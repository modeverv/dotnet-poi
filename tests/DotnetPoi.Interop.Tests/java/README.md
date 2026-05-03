# XML Parity Fixture Generator (Java)

This Maven test generates OOXML XML fixtures using Apache POI and writes them to:

```
../fixtures/xml-parity/
```

## Run

```sh
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=XmlParityFixtureGeneratorTest
```

## Notes

- The POI version is pinned to 5.5.1 in `pom.xml` for stable fixtures.
- Output XML files are flattened as `<case>__<entry-path>.xml`.
