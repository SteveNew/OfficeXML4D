unit Office4D.Tests.Metadata;

interface

uses
  System.SysUtils,
  System.DateUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Metadata,
  Office4D.Package;

type
  [TestFixture]
  TMetadataTests = class(TOffice4DTests)
  private
    FParser: TMetadataParser;

    const SampleCoreXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">' +
      '<dc:title>Test Document</dc:title>' +
      '<dc:subject>Testing</dc:subject>' +
      '<dc:creator>John Doe</dc:creator>' +
      '<dc:description>A test document</dc:description>' +
      '<cp:keywords>test, sample</cp:keywords>' +
      '<cp:lastModifiedBy>Jane Doe</cp:lastModifiedBy>' +
      '<cp:revision>5</cp:revision>' +
      '<cp:category>Testing</cp:category>' +
      '<dcterms:created>2026-01-15T10:30:00Z</dcterms:created>' +
      '<dcterms:modified>2026-01-20T14:45:00Z</dcterms:modified>' +
      '</cp:coreProperties>';

  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure Parse_Title_ExtractsTitle;

    [Test]
    procedure Parse_Creator_ExtractsAuthor;

    [Test]
    procedure Parse_Subject_ExtractsSubject;

    [Test]
    procedure Parse_Keywords_ExtractsKeywords;

    [Test]
    procedure Parse_LastModifiedBy_ExtractsLastModifiedBy;

    [Test]
    procedure Parse_Revision_ExtractsRevision;

    [Test]
    procedure Parse_Created_ExtractsCreatedDate;

    [Test]
    procedure Parse_Modified_ExtractsModifiedDate;

    [Test]
    procedure Parse_MissingFields_ReturnsEmptyStrings;

    [Test]
    procedure Parse_RealDocx_ExtractsMetadata;
  end;

implementation

{ TMetadataTests }

procedure TMetadataTests.Setup;
begin
  FParser := TMetadataParser.Create;
end;

procedure TMetadataTests.TearDown;
begin
  FParser.Free;
end;

procedure TMetadataTests.Parse_Title_ExtractsTitle;
begin
  var Meta := FParser.Parse(SampleCoreXml);

  Assert.AreEqual('Test Document', Meta.Title);
end;

procedure TMetadataTests.Parse_Creator_ExtractsAuthor;
begin
  var Meta := FParser.Parse(SampleCoreXml);

  Assert.AreEqual('John Doe', Meta.Creator);
end;

procedure TMetadataTests.Parse_Subject_ExtractsSubject;
begin
  var Meta := FParser.Parse(SampleCoreXml);

  Assert.AreEqual('Testing', Meta.Subject);
end;

procedure TMetadataTests.Parse_Keywords_ExtractsKeywords;
begin
  var Meta := FParser.Parse(SampleCoreXml);

  Assert.AreEqual('test, sample', Meta.Keywords);
end;

procedure TMetadataTests.Parse_LastModifiedBy_ExtractsLastModifiedBy;
begin
  var Meta := FParser.Parse(SampleCoreXml);

  Assert.AreEqual('Jane Doe', Meta.LastModifiedBy);
end;

procedure TMetadataTests.Parse_Revision_ExtractsRevision;
begin
  var Meta := FParser.Parse(SampleCoreXml);

  Assert.AreEqual('5', Meta.Revision);
end;

procedure TMetadataTests.Parse_Created_ExtractsCreatedDate;
begin
  var Meta := FParser.Parse(SampleCoreXml);

  Assert.AreEqual(2026, YearOf(Meta.Created));
  Assert.AreEqual(1, MonthOf(Meta.Created));
  Assert.AreEqual(15, DayOf(Meta.Created));
end;

procedure TMetadataTests.Parse_Modified_ExtractsModifiedDate;
begin
  var Meta := FParser.Parse(SampleCoreXml);

  Assert.AreEqual(2026, YearOf(Meta.Modified));
  Assert.AreEqual(1, MonthOf(Meta.Modified));
  Assert.AreEqual(20, DayOf(Meta.Modified));
end;

procedure TMetadataTests.Parse_MissingFields_ReturnsEmptyStrings;
begin
  var MinimalXml := '<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"></cp:coreProperties>';

  var Meta := FParser.Parse(MinimalXml);

  Assert.IsEmpty(Meta.Title);
  Assert.IsEmpty(Meta.Creator);
  Assert.AreEqual(Double(0), Double(Meta.Created));
end;

procedure TMetadataTests.Parse_RealDocx_ExtractsMetadata;
begin
  var Package := TOXMLPackage.Create;
  try
    Package.Open(GetWordSamplePath);
    var CoreXml := Package.GetPartContent('docProps/core.xml');
    var Meta := FParser.Parse(CoreXml);

    Assert.AreEqual('Marco Geuze', Meta.LastModifiedBy);
    Assert.IsTrue(Meta.Created > 0, 'Created date should be set');
    Assert.IsTrue(Meta.Modified > 0, 'Modified date should be set');
  finally
    Package.Free;
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TMetadataTests);

end.
