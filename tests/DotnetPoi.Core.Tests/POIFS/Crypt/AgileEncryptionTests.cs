using DotnetPoi.POIFS.Crypt;
using Xunit;

namespace DotnetPoi.POIFS.Tests.Crypt;

public class AgileEncryptionTests
{
    [Theory]
    [InlineData(15)]
    [InlineData(16)]
    [InlineData(17)]
    [InlineData(3292)]
    [InlineData(3293)]
    [InlineData(4000)]
    [InlineData(4096)]
    [InlineData(4097)]
    public void EncryptDecrypt_AgilePayload_RoundTrips(int length)
    {
        var plain = Enumerable.Range(0, length).Select(i => (byte)(i * 31)).ToArray();
        var info = new EncryptionInfo(EncryptionMode.agile);
        info.Encryptor.confirmPassword("f",
            Hex("00112233445566778899AABBCCDDEEFF"),
            Hex("102132435465768798A9BACBDCEDFE0F"),
            Hex("2031425364758697A8B9CADBECFD0E1F"),
            Hex("30415263748596A7B8C9DAEBFC0D1E2F"),
            Hex("405162738495A6B7C8D9EAFF0011223344556677"));

        using var encrypted = new MemoryStream();
        info.Encryptor.encryptPackage(plain, encrypted);
        encrypted.Position = 0;

        var read = new EncryptionInfo(encrypted);
        Assert.True(read.Decryptor.verifyPassword("f"));
        Assert.Equal(plain, read.Decryptor.getData());
        Assert.False(read.Decryptor.verifyPassword("wrong"));
    }

    private static byte[] Hex(string hex)
    {
        var result = new byte[hex.Length / 2];
        for (var i = 0; i < result.Length; i++)
        {
            result[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
        }

        return result;
    }
}
