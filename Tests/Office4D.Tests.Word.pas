unit Office4D.Tests.Word;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Word;

type
  [TestFixture]
  TWordReadTests = class
  private
    FDoc: IWordDocument;

    function GetSamplesPath: string;
    function GetWordSamplePath: string;
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure LoadFromFile_ValidDocx_LoadsDocument;

    [Test]
    procedure LoadFromFile_NonExistentFile_RaisesException;

    [Test]
    procedure GetText_SimpleParagraph_ReturnsText;

    [Test]
    procedure GetText_ReturnsHelloWorld;

    [Test]
    procedure ParagraphCount_AfterLoad_ReturnsCount;

    [Test]
    procedure Paragraph_GetText_ReturnsText;

    [Test]
    procedure Paragraph_RunCount_ReturnsRunCount;

    [Test]
    procedure Metadata_LastModifiedBy_ReturnsAuthor;

    [Test]
    procedure Metadata_Created_ReturnsDate;
  end;

  [TestFixture]
  TWordAdvancedTests = class
  private
    FDoc: IWordDocument;
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure GetText_WithLineBreak_IncludesLineBreak;

    [Test]
    procedure GetText_WithTab_IncludesTab;

    [Test]
    procedure Tables_SimpleTable_ReturnsCells;

    [Test]
    procedure Tables_TableCount_ReturnsCount;

    [Test]
    procedure Table_RowCount_ReturnsRows;

    [Test]
    procedure Table_CellText_ReturnsText;
  end;

  [TestFixture]
  TWordDOMTests = class
  private
    FDoc: IWordDocument;
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure AddParagraph_Empty_AddsToParagraphs;

    [Test]
    procedure AddRun_ToParagraph_AddsText;

    [Test]
    procedure Paragraph_MultipleRuns_ConcatenatesText;
  end;

implementation

uses
  Office4D.Errors;

{ TWordReadTests }

function TWordReadTests.GetSamplesPath: string;
begin
  Result := TPath.GetFullPath(TPath.Combine(ExtractFilePath(ParamStr(0)), '..\..\..\Samples'));
end;

function TWordReadTests.GetWordSamplePath: string;
begin
  Result := TPath.Combine(GetSamplesPath, 'Word\simple_word.docx');
end;

procedure TWordReadTests.Setup;
begin
  FDoc := TWordDocumentFactory.CreateDocument;
end;

procedure TWordReadTests.TearDown;
begin
  FDoc := nil;
end;

procedure TWordReadTests.LoadFromFile_ValidDocx_LoadsDocument;
begin
  FDoc.LoadFromFile(GetWordSamplePath);

  Assert.IsTrue(FDoc.ParagraphCount > 0);
end;

procedure TWordReadTests.LoadFromFile_NonExistentFile_RaisesException;
begin
  Assert.WillRaise(
    procedure
    begin
      FDoc.LoadFromFile('C:\nonexistent\file.docx');
    end,
    EPackageNotFound
  );
end;

procedure TWordReadTests.GetText_SimpleParagraph_ReturnsText;
begin
  FDoc.LoadFromFile(GetWordSamplePath);

  Assert.IsNotEmpty(FDoc.Text);
end;

procedure TWordReadTests.GetText_ReturnsHelloWorld;
begin
  FDoc.LoadFromFile(GetWordSamplePath);

  var Text := Trim(FDoc.Text);
  Assert.Contains(Text, 'Hello');
  Assert.Contains(Text, 'World');
end;

procedure TWordReadTests.ParagraphCount_AfterLoad_ReturnsCount;
begin
  FDoc.LoadFromFile(GetWordSamplePath);

  Assert.AreEqual(1, FDoc.ParagraphCount);
end;

procedure TWordReadTests.Paragraph_GetText_ReturnsText;
begin
  FDoc.LoadFromFile(GetWordSamplePath);

  Assert.Contains(FDoc.Paragraphs[0].Text, 'Hello');
end;

procedure TWordReadTests.Paragraph_RunCount_ReturnsRunCount;
begin
  FDoc.LoadFromFile(GetWordSamplePath);

  Assert.IsTrue(FDoc.Paragraphs[0].RunCount >= 1);
end;

procedure TWordReadTests.Metadata_LastModifiedBy_ReturnsAuthor;
begin
  FDoc.LoadFromFile(GetWordSamplePath);

  Assert.AreEqual('Marco Geuze', FDoc.Metadata.LastModifiedBy);
end;

procedure TWordReadTests.Metadata_Created_ReturnsDate;
begin
  FDoc.LoadFromFile(GetWordSamplePath);

  Assert.IsTrue(FDoc.Metadata.Created > 0, 'Created date should be set');
end;

{ TWordDOMTests }

procedure TWordDOMTests.Setup;
begin
  FDoc := TWordDocumentFactory.CreateDocument;
end;

procedure TWordDOMTests.TearDown;
begin
  FDoc := nil;
end;

procedure TWordDOMTests.AddParagraph_Empty_AddsToParagraphs;
begin
  FDoc.AddParagraph;

  Assert.AreEqual(1, FDoc.ParagraphCount);
end;

procedure TWordDOMTests.AddRun_ToParagraph_AddsText;
begin
  var Para := FDoc.AddParagraph;
  Para.AddRun('Test text');

  Assert.AreEqual('Test text', Para.Text);
end;

procedure TWordDOMTests.Paragraph_MultipleRuns_ConcatenatesText;
begin
  var Para := FDoc.AddParagraph;
  Para.AddRun('Hello ');
  Para.AddRun('World');

  Assert.AreEqual('Hello World', Para.Text);
end;

{ TWordAdvancedTests }

procedure TWordAdvancedTests.Setup;
begin
  FDoc := TWordDocumentFactory.CreateDocument;
end;

procedure TWordAdvancedTests.TearDown;
begin
  FDoc := nil;
end;

procedure TWordAdvancedTests.GetText_WithLineBreak_IncludesLineBreak;
begin
  var Para := FDoc.AddParagraph;
  Para.AddRun('Line1');
  Para.AddLineBreak;
  Para.AddRun('Line2');

  Assert.Contains(Para.Text, 'Line1');
  Assert.Contains(Para.Text, 'Line2');
  Assert.Contains(Para.Text, sLineBreak);
end;

procedure TWordAdvancedTests.GetText_WithTab_IncludesTab;
begin
  var Para := FDoc.AddParagraph;
  Para.AddRun('Col1');
  Para.AddTab;
  Para.AddRun('Col2');

  Assert.Contains(Para.Text, #9);
end;

procedure TWordAdvancedTests.Tables_SimpleTable_ReturnsCells;
begin
  var Table := FDoc.AddTable(2, 2);
  Table.Cells[0, 0].Text := 'A1';
  Table.Cells[0, 1].Text := 'B1';
  Table.Cells[1, 0].Text := 'A2';
  Table.Cells[1, 1].Text := 'B2';

  Assert.AreEqual('A1', Table.Cells[0, 0].Text);
  Assert.AreEqual('B2', Table.Cells[1, 1].Text);
end;

procedure TWordAdvancedTests.Tables_TableCount_ReturnsCount;
begin
  FDoc.AddTable(2, 2);
  FDoc.AddTable(3, 3);

  Assert.AreEqual(2, FDoc.TableCount);
end;

procedure TWordAdvancedTests.Table_RowCount_ReturnsRows;
begin
  var Table := FDoc.AddTable(3, 2);

  Assert.AreEqual(3, Table.RowCount);
  Assert.AreEqual(2, Table.ColCount);
end;

procedure TWordAdvancedTests.Table_CellText_ReturnsText;
begin
  var Table := FDoc.AddTable(1, 1);
  Table.Cells[0, 0].Text := 'Test Cell';

  Assert.AreEqual('Test Cell', Table.Cells[0, 0].Text);
end;

initialization
  TDUnitX.RegisterTestFixture(TWordReadTests);
  TDUnitX.RegisterTestFixture(TWordDOMTests);
  TDUnitX.RegisterTestFixture(TWordAdvancedTests);

end.
