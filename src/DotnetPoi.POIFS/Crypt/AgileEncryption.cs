using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using OpenMcdf;

namespace DotnetPoi.POIFS.Crypt;

public enum EncryptionMode
{
    agile
}

public sealed class EncryptionInfo
{
    internal const ushort AgileMajor = 4;
    internal const ushort AgileMinor = 4;
    internal const uint AgileFlags = 0x40;

    private readonly AgileEncryptionParameters _parameters;

    public EncryptionInfo(EncryptionMode mode)
    {
        if (mode != EncryptionMode.agile)
        {
            throw new NotSupportedException("Only agile encryption is currently supported.");
        }

        _parameters = AgileEncryptionParameters.CreateDefault();
        Encryptor = new Encryptor(_parameters);
        Decryptor = new Decryptor(_parameters);
    }

    public EncryptionInfo(Stream poifsStream)
    {
        ArgumentNullException.ThrowIfNull(poifsStream);
        var streams = OpenMcdfStreams.ReadStreams(poifsStream);
        if (!streams.TryGetValue("EncryptionInfo", out var encryptionInfo))
        {
            throw new InvalidDataException("The OLE2 file does not contain an EncryptionInfo stream.");
        }

        _parameters = AgileEncryptionParameters.Parse(encryptionInfo);
        Encryptor = new Encryptor(_parameters);
        Decryptor = new Decryptor(_parameters, streams);
    }

    public Encryptor Encryptor { get; }

    public Decryptor Decryptor { get; }
}

public sealed class Encryptor
{
    private readonly AgileEncryptionParameters _parameters;

    internal Encryptor(AgileEncryptionParameters parameters)
    {
        _parameters = parameters;
    }

    public void confirmPassword(string password)
    {
        Span<byte> keySpec = stackalloc byte[16];
        Span<byte> keySalt = stackalloc byte[16];
        Span<byte> verifierSalt = stackalloc byte[16];
        Span<byte> verifier = stackalloc byte[16];
        Span<byte> integritySalt = stackalloc byte[20];
        RandomNumberGenerator.Fill(keySpec);
        RandomNumberGenerator.Fill(keySalt);
        RandomNumberGenerator.Fill(verifierSalt);
        RandomNumberGenerator.Fill(verifier);
        RandomNumberGenerator.Fill(integritySalt);
        confirmPassword(password, keySpec.ToArray(), keySalt.ToArray(), verifier.ToArray(), verifierSalt.ToArray(), integritySalt.ToArray());
    }

    public void confirmPassword(string password, byte[] keySpec, byte[] keySalt, byte[] verifier, byte[] verifierSalt, byte[] integritySalt)
    {
        ArgumentNullException.ThrowIfNull(keySpec);
        ArgumentNullException.ThrowIfNull(keySalt);
        ArgumentNullException.ThrowIfNull(verifier);
        ArgumentNullException.ThrowIfNull(verifierSalt);
        ArgumentNullException.ThrowIfNull(integritySalt);

        if (keySpec.Length != _parameters.KeySizeBytes || keySalt.Length != _parameters.BlockSize
            || verifier.Length != _parameters.BlockSize || verifierSalt.Length != _parameters.BlockSize
            || integritySalt.Length != _parameters.HashSize)
        {
            throw new ArgumentException("Agile encryption salts and keys must match POI default sizes.");
        }

        _parameters.KeySalt = keySalt.ToArray();
        _parameters.VerifierSalt = verifierSalt.ToArray();
        _parameters.SecretKey = keySpec.ToArray();

        var pwHash = AgileCrypto.HashPassword(password, verifierSalt, _parameters.SpinCount);
        _parameters.EncryptedVerifierHashInput = AgileCrypto.HashInput(_parameters, pwHash, AgileCrypto.VerifierInputBlock, verifier, encrypt: true);
        _parameters.EncryptedVerifierHashValue = AgileCrypto.HashInput(_parameters, pwHash, AgileCrypto.HashedVerifierBlock, SHA1.HashData(verifier), encrypt: true);
        _parameters.EncryptedKeyValue = AgileCrypto.HashInput(_parameters, pwHash, AgileCrypto.CryptoKeyBlock, keySpec, encrypt: true);

        var hmacKeyIv = AgileCrypto.GenerateIv(_parameters.KeySalt, AgileCrypto.IntegrityKeyBlock, _parameters.BlockSize);
        _parameters.EncryptedHmacKey = AgileCrypto.AesCbcNoPadding(keySpec, hmacKeyIv,
            AgileCrypto.GetBlock0(integritySalt, AgileCrypto.GetNextBlockSize(integritySalt.Length, _parameters.BlockSize)),
            encrypt: true);
        _parameters.IntegritySalt = integritySalt.ToArray();
    }

    public void encryptPackage(byte[] packageBytes, Stream output)
    {
        ArgumentNullException.ThrowIfNull(packageBytes);
        ArgumentNullException.ThrowIfNull(output);

        if (_parameters.SecretKey is null || _parameters.IntegritySalt is null)
        {
            confirmPassword(Decryptor.DEFAULT_PASSWORD);
        }

        var encryptedPackageBody = AgileCrypto.EncryptPackage(packageBytes, _parameters);
        _parameters.UpdateIntegrity(encryptedPackageBody, packageBytes.Length);
        var encryptedPackage = new byte[8 + encryptedPackageBody.Length];
        BinaryPrimitives.WriteInt64LittleEndian(encryptedPackage, packageBytes.Length);
        Buffer.BlockCopy(encryptedPackageBody, 0, encryptedPackage, 8, encryptedPackageBody.Length);

        OpenMcdfStreams.Write(output, _parameters.WriteEncryptionInfo(), encryptedPackage);
    }
}

public sealed class Decryptor
{
    public const string DEFAULT_PASSWORD = "VelvetSweatshop";
    public const string DEFAULT_POIFS_ENTRY = "EncryptedPackage";

    private readonly AgileEncryptionParameters _parameters;
    private readonly IReadOnlyDictionary<string, byte[]> _streams;

    internal Decryptor(AgileEncryptionParameters parameters, IReadOnlyDictionary<string, byte[]>? streams = null)
    {
        _parameters = parameters;
        _streams = streams ?? new Dictionary<string, byte[]>(StringComparer.Ordinal);
    }

    public bool verifyPassword(string password)
    {
        var pwHash = AgileCrypto.HashPassword(password, _parameters.VerifierSalt, _parameters.SpinCount);
        var verifierInput = AgileCrypto.HashInput(_parameters, pwHash, AgileCrypto.VerifierInputBlock, _parameters.EncryptedVerifierHashInput, encrypt: false);
        var verifierHash = SHA1.HashData(verifierInput);
        var verifierHashDec = AgileCrypto.GetBlock0(AgileCrypto.HashInput(_parameters, pwHash, AgileCrypto.HashedVerifierBlock, _parameters.EncryptedVerifierHashValue, encrypt: false), _parameters.HashSize);
        if (!CryptographicOperations.FixedTimeEquals(verifierHash, verifierHashDec))
        {
            return false;
        }

        _parameters.SecretKey = AgileCrypto.GetBlock0(AgileCrypto.HashInput(_parameters, pwHash, AgileCrypto.CryptoKeyBlock, _parameters.EncryptedKeyValue, encrypt: false), _parameters.KeySizeBytes);
        return true;
    }

    public byte[] getData()
    {
        if (_parameters.SecretKey is null)
        {
            throw new InvalidOperationException("Call verifyPassword before decrypting the package.");
        }

        if (!_streams.TryGetValue(DEFAULT_POIFS_ENTRY, out var encryptedPackage) || encryptedPackage.Length < 8)
        {
            throw new InvalidDataException("The OLE2 file does not contain an EncryptedPackage stream.");
        }

        var length = BinaryPrimitives.ReadInt64LittleEndian(encryptedPackage);
        var encryptedLength = AgileCrypto.GetEncryptedPackageLength(checked((int)length));
        var body = encryptedPackage.AsSpan(8, encryptedLength).ToArray();
        return AgileCrypto.DecryptPackage(body, checked((int)length), _parameters);
    }
}

internal sealed class AgileEncryptionParameters
{
    private const string EncNs = "http://schemas.microsoft.com/office/2006/encryption";
    private const string PassNs = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

    public int BlockSize { get; private init; } = 16;
    public int KeyBits { get; private init; } = 128;
    public int KeySizeBytes => KeyBits / 8;
    public int HashSize { get; private init; } = 20;
    public int SpinCount { get; private init; } = 100000;

    public byte[] KeySalt { get; set; } = Array.Empty<byte>();
    public byte[] VerifierSalt { get; set; } = Array.Empty<byte>();
    public byte[] EncryptedVerifierHashInput { get; set; } = Array.Empty<byte>();
    public byte[] EncryptedVerifierHashValue { get; set; } = Array.Empty<byte>();
    public byte[] EncryptedKeyValue { get; set; } = Array.Empty<byte>();
    public byte[] EncryptedHmacKey { get; set; } = Array.Empty<byte>();
    public byte[] EncryptedHmacValue { get; set; } = Array.Empty<byte>();
    public byte[]? SecretKey { get; set; }
    public byte[]? IntegritySalt { get; set; }

    public static AgileEncryptionParameters CreateDefault()
    {
        return new AgileEncryptionParameters();
    }

    public static AgileEncryptionParameters Parse(byte[] encryptionInfo)
    {
        if (encryptionInfo.Length < 8)
        {
            throw new InvalidDataException("EncryptionInfo stream is too short.");
        }

        var major = BinaryPrimitives.ReadUInt16LittleEndian(encryptionInfo);
        var minor = BinaryPrimitives.ReadUInt16LittleEndian(encryptionInfo.AsSpan(2));
        var flags = BinaryPrimitives.ReadUInt32LittleEndian(encryptionInfo.AsSpan(4));
        if (major != EncryptionInfo.AgileMajor || minor != EncryptionInfo.AgileMinor || flags != EncryptionInfo.AgileFlags)
        {
            throw new NotSupportedException("Only Agile EncryptionInfo version 4.4 is supported.");
        }

        using var xml = new MemoryStream(encryptionInfo, 8, encryptionInfo.Length - 8);
        var doc = new XmlDocument { PreserveWhitespace = true };
        doc.Load(xml);
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("e", EncNs);
        ns.AddNamespace("p", PassNs);

        var keyData = (XmlElement?)doc.SelectSingleNode("/e:encryption/e:keyData", ns)
            ?? throw new InvalidDataException("Missing keyData.");
        var encryptedKey = (XmlElement?)doc.SelectSingleNode("/e:encryption/e:keyEncryptors/e:keyEncryptor/p:encryptedKey", ns)
            ?? throw new InvalidDataException("Missing password encryptedKey.");
        var dataIntegrity = (XmlElement?)doc.SelectSingleNode("/e:encryption/e:dataIntegrity", ns)
            ?? throw new InvalidDataException("Missing dataIntegrity.");

        return new AgileEncryptionParameters
        {
            BlockSize = ReadInt(keyData, "blockSize"),
            KeyBits = ReadInt(keyData, "keyBits"),
            HashSize = ReadInt(keyData, "hashSize"),
            KeySalt = ReadBin(keyData, "saltValue"),
            VerifierSalt = ReadBin(encryptedKey, "saltValue"),
            SpinCount = ReadInt(encryptedKey, "spinCount"),
            EncryptedVerifierHashInput = ReadBin(encryptedKey, "encryptedVerifierHashInput"),
            EncryptedVerifierHashValue = ReadBin(encryptedKey, "encryptedVerifierHashValue"),
            EncryptedKeyValue = ReadBin(encryptedKey, "encryptedKeyValue"),
            EncryptedHmacKey = ReadBin(dataIntegrity, "encryptedHmacKey"),
            EncryptedHmacValue = ReadBin(dataIntegrity, "encryptedHmacValue")
        };
    }

    public void UpdateIntegrity(byte[] encryptedPackageBody, int plainLength)
    {
        if (SecretKey is null || IntegritySalt is null)
        {
            throw new InvalidOperationException("Password must be confirmed before writing integrity data.");
        }

        var hmacKey = AgileCrypto.GetBlock0(IntegritySalt, AgileCrypto.GetNextBlockSize(IntegritySalt.Length, BlockSize));
        using var hmac = new HMACSHA1(hmacKey);
        Span<byte> size = stackalloc byte[8];
        BinaryPrimitives.WriteInt64LittleEndian(size, plainLength);
        hmac.TransformBlock(size.ToArray(), 0, size.Length, null, 0);
        hmac.TransformFinalBlock(encryptedPackageBody, 0, encryptedPackageBody.Length);

        var hmacValue = AgileCrypto.GetBlock0(hmac.Hash!, AgileCrypto.GetNextBlockSize(hmac.Hash!.Length, BlockSize));
        var hmacValueIv = AgileCrypto.GenerateIv(KeySalt, AgileCrypto.IntegrityValueBlock, BlockSize);
        EncryptedHmacValue = AgileCrypto.AesCbcNoPadding(SecretKey, hmacValueIv, hmacValue, encrypt: true);
    }

    public byte[] WriteEncryptionInfo()
    {
        using var memory = new MemoryStream();
        Span<byte> header = stackalloc byte[8];
        BinaryPrimitives.WriteUInt16LittleEndian(header, EncryptionInfo.AgileMajor);
        BinaryPrimitives.WriteUInt16LittleEndian(header[2..], EncryptionInfo.AgileMinor);
        BinaryPrimitives.WriteUInt32LittleEndian(header[4..], EncryptionInfo.AgileFlags);
        memory.Write(header);

        var xml = BuildEncryptionInfoXml();
        var xmlBytes = Encoding.UTF8.GetBytes(xml);
        memory.Write(xmlBytes);

        return memory.ToArray();
    }

    private string BuildEncryptionInfoXml()
    {
        XNamespace ns = EncNs;
        XNamespace p = PassNs;

        var keyData = new XElement(ns + "keyData",
            new XAttribute("blockSize", BlockSize),
            new XAttribute("cipherAlgorithm", "AES"),
            new XAttribute("cipherChaining", "ChainingModeCBC"),
            new XAttribute("hashAlgorithm", "SHA1"),
            new XAttribute("hashSize", HashSize),
            new XAttribute("keyBits", KeyBits),
            new XAttribute("saltSize", BlockSize),
            new XAttribute("saltValue", Convert.ToBase64String(KeySalt)));

        var dataIntegrity = new XElement(ns + "dataIntegrity",
            new XAttribute("encryptedHmacKey", Convert.ToBase64String(EncryptedHmacKey)),
            new XAttribute("encryptedHmacValue", Convert.ToBase64String(EncryptedHmacValue)));

        var encryptedKey = new XElement(p + "encryptedKey",
            new XAttribute("blockSize", BlockSize),
            new XAttribute("cipherAlgorithm", "AES"),
            new XAttribute("cipherChaining", "ChainingModeCBC"),
            new XAttribute("encryptedKeyValue", Convert.ToBase64String(EncryptedKeyValue)),
            new XAttribute("encryptedVerifierHashInput", Convert.ToBase64String(EncryptedVerifierHashInput)),
            new XAttribute("encryptedVerifierHashValue", Convert.ToBase64String(EncryptedVerifierHashValue)),
            new XAttribute("hashAlgorithm", "SHA1"),
            new XAttribute("hashSize", HashSize),
            new XAttribute("keyBits", KeyBits),
            new XAttribute("saltSize", BlockSize),
            new XAttribute("saltValue", Convert.ToBase64String(VerifierSalt)),
            new XAttribute("spinCount", SpinCount));

        var doc = new XDocument(
            new XElement(ns + "encryption",
                new XAttribute(XNamespace.Xmlns + "p", PassNs),
                keyData,
                dataIntegrity,
                new XElement(ns + "keyEncryptors",
                    new XElement(ns + "keyEncryptor",
                        new XAttribute("uri", PassNs),
                        encryptedKey))));

        return doc.ToString(SaveOptions.DisableFormatting);
    }

    private static int ReadInt(XmlElement element, string name)
    {
        return int.Parse(element.GetAttribute(name), System.Globalization.CultureInfo.InvariantCulture);
    }

    private static byte[] ReadBin(XmlElement element, string name)
    {
        return Convert.FromBase64String(element.GetAttribute(name));
    }
}

internal static class AgileCrypto
{
    public static readonly byte[] VerifierInputBlock = LongToBytes(0xfea7d2763b4b9e79UL);
    public static readonly byte[] HashedVerifierBlock = LongToBytes(0xd7aa0f6d3061344eUL);
    public static readonly byte[] CryptoKeyBlock = LongToBytes(0x146e0be7abacd0d6UL);
    public static readonly byte[] IntegrityKeyBlock = LongToBytes(0x5fb2ad010cb9e1f6UL);
    public static readonly byte[] IntegrityValueBlock = LongToBytes(0xa0677f02b22c8433UL);

    public static byte[] HashPassword(string? password, byte[] salt, int spinCount)
    {
        password ??= Decryptor.DEFAULT_PASSWORD;
        using var sha1 = SHA1.Create();
        var passwordBytes = Encoding.Unicode.GetBytes(password);
        sha1.TransformBlock(salt, 0, salt.Length, null, 0);
        sha1.TransformFinalBlock(passwordBytes, 0, passwordBytes.Length);
        var hash = sha1.Hash!;
        Span<byte> iterator = stackalloc byte[4];
        for (var i = 0; i < spinCount; i++)
        {
            BinaryPrimitives.WriteInt32LittleEndian(iterator, i);
            using var iterSha = SHA1.Create();
            iterSha.TransformBlock(iterator.ToArray(), 0, 4, null, 0);
            iterSha.TransformFinalBlock(hash, 0, hash.Length);
            hash = iterSha.Hash!;
        }

        return hash;
    }

    public static byte[] HashInput(AgileEncryptionParameters p, byte[] pwHash, byte[] blockKey, byte[] inputKey, bool encrypt)
    {
        var key = GenerateKey(pwHash, blockKey, p.KeySizeBytes);
        var iv = GenerateIv(p.VerifierSalt, null, p.BlockSize);
        var data = GetBlock0(inputKey, GetNextBlockSize(inputKey.Length, p.BlockSize));
        return AesCbcNoPadding(key, iv, data, encrypt);
    }

    public static byte[] EncryptPackage(byte[] packageBytes, AgileEncryptionParameters p)
    {
        return TransformPackage(packageBytes, packageBytes.Length, p, encrypt: true);
    }

    public static byte[] DecryptPackage(byte[] encryptedPackage, int length, AgileEncryptionParameters p)
    {
        var decrypted = TransformPackage(encryptedPackage, length, p, encrypt: false);
        Array.Resize(ref decrypted, length);
        return decrypted;
    }

    public static int GetEncryptedPackageLength(int plainLength)
    {
        var fullChunks = plainLength / 4096;
        var rem = plainLength % 4096;
        if (rem == 0)
        {
            return plainLength;
        }

        return fullChunks * 4096 + ((rem / 16) + 1) * 16;
    }

    public static byte[] GenerateKey(byte[] passwordHash, byte[] blockKey, int keySize)
    {
        using var sha1 = SHA1.Create();
        sha1.TransformBlock(passwordHash, 0, passwordHash.Length, null, 0);
        sha1.TransformFinalBlock(blockKey, 0, blockKey.Length);
        return GetBlock36(sha1.Hash!, keySize);
    }

    public static byte[] GenerateIv(byte[] salt, byte[]? blockKey, int blockSize)
    {
        byte[] iv;
        if (blockKey is null)
        {
            iv = salt;
        }
        else
        {
            using var sha1 = SHA1.Create();
            sha1.TransformBlock(salt, 0, salt.Length, null, 0);
            sha1.TransformFinalBlock(blockKey, 0, blockKey.Length);
            iv = sha1.Hash!;
        }

        return GetBlock36(iv, blockSize);
    }

    public static byte[] GetBlock0(byte[] hash, int size) => GetBlockX(hash, size, 0x00);

    public static byte[] AesCbcNoPadding(byte[] key, byte[] iv, byte[] data, bool encrypt)
    {
        using var aes = Aes.Create();
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None;
        aes.Key = key;
        aes.IV = iv;
        using var transform = encrypt ? aes.CreateEncryptor() : aes.CreateDecryptor();
        return transform.TransformFinalBlock(data, 0, data.Length);
    }

    private static byte[] TransformPackage(byte[] input, int plainLength, AgileEncryptionParameters p, bool encrypt)
    {
        if (p.SecretKey is null)
        {
            throw new InvalidOperationException("Secret key has not been derived.");
        }

        using var output = new MemoryStream();
        var offset = 0;
        var block = 0;
        while (offset < input.Length)
        {
            var remaining = input.Length - offset;
            var chunkLength = Math.Min(4096, remaining);
            var last = encrypt
                ? offset + chunkLength == plainLength && chunkLength < 4096
                : offset + chunkLength == input.Length && plainLength % 4096 != 0;
            var chunk = new byte[chunkLength];
            Buffer.BlockCopy(input, offset, chunk, 0, chunkLength);

            var blockKey = new byte[4];
            BinaryPrimitives.WriteInt32LittleEndian(blockKey, block);
            var iv = GenerateIv(p.KeySalt, blockKey, p.BlockSize);
            var transformed = TransformPackageChunk(p.SecretKey, iv, chunk, encrypt, last);
            output.Write(transformed);

            offset += chunkLength;
            block++;
        }

        return output.ToArray();
    }

    private static byte[] TransformPackageChunk(byte[] key, byte[] iv, byte[] chunk, bool encrypt, bool last)
    {
        using var aes = Aes.Create();
        aes.Mode = CipherMode.CBC;
        aes.Padding = last ? PaddingMode.PKCS7 : PaddingMode.None;
        aes.Key = key;
        aes.IV = iv;
        using var transform = encrypt ? aes.CreateEncryptor() : aes.CreateDecryptor();
        return transform.TransformFinalBlock(chunk, 0, chunk.Length);
    }

    public static int GetNextBlockSize(int inputLen, int blockSize)
    {
        return (int)Math.Ceiling(inputLen / (double)blockSize) * blockSize;
    }

    private static byte[] GetBlock36(byte[] hash, int size) => GetBlockX(hash, size, 0x36);

    private static byte[] GetBlockX(byte[] hash, int size, byte fill)
    {
        var result = Enumerable.Repeat(fill, size).ToArray();
        Buffer.BlockCopy(hash, 0, result, 0, Math.Min(hash.Length, result.Length));
        return result;
    }

    private static byte[] LongToBytes(ulong value)
    {
        var bytes = new byte[8];
        BinaryPrimitives.WriteUInt64BigEndian(bytes, value);
        return bytes;
    }
}

internal static class OpenMcdfStreams
{
    public static void Write(Stream output, byte[] encryptionInfo, byte[] encryptedPackage)
    {
        using var root = RootStorage.Create(output, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen);
        using (var stream = root.CreateStream("EncryptionInfo"))
        {
            stream.Write(encryptionInfo);
        }

        using (var stream = root.CreateStream(Decryptor.DEFAULT_POIFS_ENTRY))
        {
            stream.Write(encryptedPackage);
        }

        root.Flush(consolidate: true);
    }

    public static Dictionary<string, byte[]> ReadStreams(Stream input)
    {
        using var root = RootStorage.Open(input, StorageModeFlags.LeaveOpen);
        return new Dictionary<string, byte[]>(StringComparer.Ordinal)
        {
            ["EncryptionInfo"] = ReadStream(root, "EncryptionInfo"),
            [Decryptor.DEFAULT_POIFS_ENTRY] = ReadStream(root, Decryptor.DEFAULT_POIFS_ENTRY)
        };
    }

    private static byte[] ReadStream(RootStorage root, string name)
    {
        using var stream = root.OpenStream(name);
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }
}
