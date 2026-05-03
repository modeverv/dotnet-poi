# Phase 3.4 Agile Encryption Notes

This note captures the working state of the OOXML Agile encryption slice and the traps found while making Excel accept the generated file. Read this before changing `src/DotnetPoi.POIFS/Crypt/AgileEncryption.cs`, especially before removing the temporary `OpenMcdf` dependency.

## Current Working State

- `XSSFWorkbook.writeEncrypted(Stream, string)` writes a normal xlsx package to memory, then wraps it as an Agile-encrypted OLE2 compound file.
- `EncryptionInfo` / `EncryptedPackage` CFB streams are currently written with `OpenMcdf`.
- Apache POI can decrypt and read the generated fixture.
- Microsoft Excel opens the generated encrypted xlsx without the previous corruption warning after the integrity HMAC fixes below.

The generated interop fixture is:

```text
tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase3_4-agile-encrypted.xlsx
```

Password used by tests:

```text
f
```

## Why OpenMcdf Is Currently Used

An earlier hand-written minimal CFB writer produced files Apache POI could read, but Excel still treated them as corrupt or suspicious. POI is permissive enough that "POI can decrypt" is not sufficient for Excel compatibility.

The known fragile CFB areas are:

- directory sibling tree ordering and red/black node metadata
- mini FAT vs regular FAT stream placement
- root storage mini-stream metadata
- sector allocation and stream cutoff behavior

For Phase 3.4, `OpenMcdf` is intentionally used as a compatibility crutch so the cryptographic work can stabilize independently from a full POIFS writer. Removing it should be treated as POIFS work, not as a simple dependency cleanup.

## Excel-Critical Agile Details

These details were necessary for Excel to open the file cleanly:

- `EncryptionInfo` version header must be:
  - major `0x0004`
  - minor `0x0004`
  - flags `0x00000040`
- `EncryptedPackage` stream begins with an 8-byte little-endian original package length.
- Encrypted package data is AES-CBC in 4096-byte chunks.
- Chunk IV is `SHA1(keySalt + UInt32LE(chunkIndex))`, truncated/padded to block size.
- Full 4096-byte chunks use `PaddingMode.None`.
- The final partial chunk uses `PaddingMode.PKCS7`.
- Password hash loop is POI-compatible:
  - `H0 = SHA1(verifierSalt + UTF-16LE(password))`
  - each spin hashes `UInt32LE(i) + previousHash`
- Block keys are the POI/MS-OFFCRYPTO byte sequences:
  - verifier input: `FE A7 D2 76 3B 4B 9E 79`
  - verifier hash: `D7 AA 0F 6D 30 61 34 4E`
  - encrypted key: `14 6E 0B E7 AB AC D0 D6`
  - integrity HMAC key: `5F B2 AD 01 0C B9 E1 F6`
  - integrity HMAC value: `A0 67 7F 02 B2 2C 84 33`

## Integrity HMAC Trap

This was the final Excel warning culprit.

For SHA1, the integrity salt is 20 bytes. It must be zero-padded to the next AES block boundary, i.e. 32 bytes, in two separate places:

- before encrypting `encryptedHmacKey`
- before using it as the HMAC key for `encryptedHmacValue`

Using 16 bytes in either place can still let Apache POI decrypt the package, but Excel may show:

```text
This file may be corrupted or tampered with, and its contents cannot be trusted.
```

or the localized equivalent. If the user clicks Yes, Excel may still open the workbook, which makes this easy to miss in automated POI-only tests.

## EncryptionInfo XML Shape

The XML is currently written without an XML declaration and with attribute order matching the known working implementation in:

```text
/Users/seijiro/Sync/sync_work/me/SetPassToExceldotNet/src/ExcelEncryptor/Encrypt.cs
```

The current implementation builds the XML with `XDocument.ToString(SaveOptions.DisableFormatting)` to match that working file shape.

Do not "clean up" this XML writer casually. Byte-level XML shape has caused previous Office encryption integrity mismatches.

## Tests To Run

Run these after any Agile encryption or CFB change:

```bash
dotnet test tests/DotnetPoi.POIFS.Tests/DotnetPoi.POIFS.Tests.csproj
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --filter "FullyQualifiedName~Write_AgileEncryptedWorkbook_CreatesFixtureForPoi"
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest#readPhase34AgileEncryptedWorkbook
```

The xUnit runner may need elevated sandbox permissions because it opens a local test communication socket.

Manual Excel verification is still required for this phase. POI interop alone does not validate the same integrity path Excel validates.

## OpenMcdf Removal Plan

When replacing `OpenMcdf` with an in-repo POIFS writer, preserve behavior first and refactor second.

Recommended sequence:

1. Keep `AgileEncryption.cs` cryptographic output unchanged.
2. Add a POIFS writer test that writes exactly two streams:
   - `EncryptionInfo`
   - `EncryptedPackage`
3. Compare the resulting CFB with an `OpenMcdf`-generated file structurally:
   - stream names and lengths
   - FAT chain lengths
   - mini FAT usage
   - root storage directory tree
   - sector size and mini stream cutoff
4. Validate with Apache POI.
5. Validate manually with Excel and confirm no corruption/tamper warning appears.
6. Only then remove `OpenMcdf` from `src/DotnetPoi.POIFS/DotnetPoi.POIFS.csproj`.

Avoid padding streams just to dodge mini FAT behavior. Excel notices more CFB detail than POI. A proper in-repo implementation should either support mini streams correctly or intentionally store streams in regular FAT while keeping header/directory metadata consistent.

## Known Backlog

- Implement full in-repo POIFS/CFB writer and reader.
- Add a test helper that recomputes Agile integrity HMAC from `EncryptedPackage` and `EncryptionInfo`.
- Add an Excel/manual verification checklist artifact for release candidates.
- Consider supporting AES-192/AES-256 and SHA-256+ after the AES-128/SHA1 path is fully stabilized.
