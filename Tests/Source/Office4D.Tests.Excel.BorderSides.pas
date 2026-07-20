unit Office4D.Tests.Excel.BorderSides;

interface

uses
  System.SysUtils,
  System.IOUtils,
  System.RegularExpressions,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Excel;

type
  [TestFixture]
  TExcelBorderSidesTests = class(TOffice4DTests)
  private
    FWorkbook: IExcelWorkbook;
    FTempFile: string;

    function ExtractCellStyleIndex(const SheetXml, CellRef: string): Integer;

  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure SaveToFile_DifferentTopBorderColors_ProduceDifferentStyleIndices;

    [Test]
    procedure SaveToFile_BorderOnlyCell_WritesStyledEmptyCell;

    [Test]
    procedure RoundTrip_UniformBorderSetter_AppliesToAllFourSides;

    [Test]
    procedure SaveToFile_PlainCell_HasNoStyleAttribute;

    [Test]
    procedure RoundTrip_MixedSideBorders_PreservesEachSideIndependently;

    [Test]
    procedure RoundTrip_OutlinePattern_ProducesCorrectSideCountPerPosition;

    [Test]
    procedure GetBorderStyle_MultiSideQuery_ReturnsFirstMatchingSideInOrder;

    [Test]
    procedure SaveToFile_ColorWithoutStyle_ProducesNoBorderElement;

    [Test]
    procedure LoadFromFile_AsymmetricBottomOnlyBorder_ReadsAllFourSidesIndependently;

    [Test]
    procedure RoundTrip_BottomOnlyColor_DoesNotBleedIntoOtherSides;
  end;

implementation

uses
  System.Classes,
  System.Zip,
  Office4D.Package;

{ TExcelBorderSidesTests }

procedure TExcelBorderSidesTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
  FTempFile := TPath.Combine(TPath.GetTempPath, 'border_test_' + TGUID.NewGuid.ToString + '.xlsx');
end;

procedure TExcelBorderSidesTests.TearDown;
begin
  FWorkbook := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

function TExcelBorderSidesTests.ExtractCellStyleIndex(const SheetXml, CellRef: string): Integer;
var
  M: TMatch;
begin
  M := TRegEx.Match(SheetXml, '<c\s+r="' + CellRef + '"[^>]*\ss="(\d+)"', [roIgnoreCase]);
  if M.Success then
    Result := StrToIntDef(M.Groups[1].Value, 0)
  else
    Result := 0; // no s= attribute present means the default style, index 0
end;

procedure TExcelBorderSidesTests.SaveToFile_DifferentTopBorderColors_ProduceDifferentStyleIndices;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');

  Sheet.Cell['A1'].BorderStyle[[TExcelBorderSide.Top]] := TExcelBorderStyle.Thin;
  Sheet.Cell['A1'].BorderColor[[TExcelBorderSide.Top]] := $FF0000;

  Sheet.Cell['B1'].BorderStyle[[TExcelBorderSide.Top]] := TExcelBorderStyle.Thin;
  Sheet.Cell['B1'].BorderColor[[TExcelBorderSide.Top]] := $0000FF;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    const IdxA1 = ExtractCellStyleIndex(SheetXml, 'A1');
    const IdxB1 = ExtractCellStyleIndex(SheetXml, 'B1');

    Assert.AreNotEqual(IdxA1, IdxB1,
      'Cells with different top border colors must resolve to different style indices -- ' +
      'a failure here means the style-key field order/count has drifted');
  finally
    Package.Free;
  end;
end;

procedure TExcelBorderSidesTests.SaveToFile_BorderOnlyCell_WritesStyledEmptyCell;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].BorderStyle[[TExcelBorderSide.Bottom]] := TExcelBorderStyle.Thin;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsTrue(Pos('r="A1" s="', SheetXml) > 0,
      'A cell with only a border side set must still be written with a non-default style, ' +
      'even though it has no value and no other formatting');
  finally
    Package.Free;
  end;
end;

procedure TExcelBorderSidesTests.RoundTrip_UniformBorderSetter_AppliesToAllFourSides;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Boxed';
  Sheet.Cell['A1'].BorderStyle[AllBorderSides] := TExcelBorderStyle.Medium;
  Sheet.Cell['A1'].BorderColor[AllBorderSides] := $123456;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  const Cell = Workbook2.Sheets[0].Cell['A1'];

  Assert.AreEqual(Ord(TExcelBorderStyle.Medium), Ord(Cell.BorderStyle[[TExcelBorderSide.Top]]), 'Top style');
  Assert.AreEqual(Ord(TExcelBorderStyle.Medium), Ord(Cell.BorderStyle[[TExcelBorderSide.Right]]), 'Right style');
  Assert.AreEqual(Ord(TExcelBorderStyle.Medium), Ord(Cell.BorderStyle[[TExcelBorderSide.Bottom]]), 'Bottom style');
  Assert.AreEqual(Ord(TExcelBorderStyle.Medium), Ord(Cell.BorderStyle[[TExcelBorderSide.Left]]), 'Left style');
  Assert.AreEqual(Cardinal($123456), Cell.BorderColor[[TExcelBorderSide.Top]], 'Top color');
  Assert.AreEqual(Cardinal($123456), Cell.BorderColor[[TExcelBorderSide.Right]], 'Right color');
  Assert.AreEqual(Cardinal($123456), Cell.BorderColor[[TExcelBorderSide.Bottom]], 'Bottom color');
  Assert.AreEqual(Cardinal($123456), Cell.BorderColor[[TExcelBorderSide.Left]], 'Left color');
end;

procedure TExcelBorderSidesTests.SaveToFile_PlainCell_HasNoStyleAttribute;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Plain';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsFalse(Pos('r="A1" s="', SheetXml) > 0,
      'A cell with no formatting at all must resolve to the default style (index 0) and ' +
      'carry no s= attribute');
  finally
    Package.Free;
  end;
end;

procedure TExcelBorderSidesTests.RoundTrip_MixedSideBorders_PreservesEachSideIndependently;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'X';
  Sheet.Cell['A1'].BorderStyle[[TExcelBorderSide.Top]] := TExcelBorderStyle.Thin;
  Sheet.Cell['A1'].BorderColor[[TExcelBorderSide.Top]] := $000000;
  Sheet.Cell['A1'].BorderStyle[[TExcelBorderSide.Right]] := TExcelBorderStyle.Thick;
  Sheet.Cell['A1'].BorderColor[[TExcelBorderSide.Right]] := $FF0000;
  // Bottom deliberately left as None -- the one side that must NOT come back set.
  Sheet.Cell['A1'].BorderStyle[[TExcelBorderSide.Left]] := TExcelBorderStyle.Dashed;
  Sheet.Cell['A1'].BorderColor[[TExcelBorderSide.Left]] := $0000FF;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  const Cell = Workbook2.Sheets[0].Cell['A1'];

  Assert.AreEqual(Ord(TExcelBorderStyle.Thin), Ord(Cell.BorderStyle[[TExcelBorderSide.Top]]), 'Top style survives round-trip');
  Assert.AreEqual(Cardinal($000000), Cell.BorderColor[[TExcelBorderSide.Top]], 'Top color survives round-trip');
  Assert.AreEqual(Ord(TExcelBorderStyle.Thick), Ord(Cell.BorderStyle[[TExcelBorderSide.Right]]), 'Right style survives round-trip');
  Assert.AreEqual(Cardinal($FF0000), Cell.BorderColor[[TExcelBorderSide.Right]], 'Right color survives round-trip');
  Assert.AreEqual(Ord(TExcelBorderStyle.None), Ord(Cell.BorderStyle[[TExcelBorderSide.Bottom]]), 'Bottom correctly stayed unset');
  Assert.AreEqual(Ord(TExcelBorderStyle.Dashed), Ord(Cell.BorderStyle[[TExcelBorderSide.Left]]), 'Left style survives round-trip');
  Assert.AreEqual(Cardinal($0000FF), Cell.BorderColor[[TExcelBorderSide.Left]], 'Left color survives round-trip');
end;

procedure TExcelBorderSidesTests.RoundTrip_OutlinePattern_ProducesCorrectSideCountPerPosition;
const
  Cols: array[0..2] of string = ('A', 'B', 'C');
var
  c, r, SideCount: Integer;
  Addr: string;
  Side: TExcelBorderSide;
  Sides: TExcelBorderSides;
  Cell: IExcelCell;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');

  for c := 0 to 2 do
    for r := 0 to 2 do
    begin
      Addr := Cols[c] + IntToStr(r + 1);
      Cell := Sheet.Cell[Addr];

      Sides := [];
      if r = 0 then Include(Sides, TExcelBorderSide.Top);
      if r = 2 then Include(Sides, TExcelBorderSide.Bottom);
      if c = 0 then Include(Sides, TExcelBorderSide.Left);
      if c = 2 then Include(Sides, TExcelBorderSide.Right);

      if Sides <> [] then
      begin
        Cell.BorderStyle[Sides] := TExcelBorderStyle.Thin;
        Cell.BorderColor[Sides] := $000000;
      end;
    end;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  const Sheet2 = Workbook2.Sheets[0];

  for c := 0 to 2 do
    for r := 0 to 2 do
    begin
      Addr := Cols[c] + IntToStr(r + 1);
      Cell := Sheet2.Cell[Addr];

      SideCount := 0;
      for Side := Low(TExcelBorderSide) to High(TExcelBorderSide) do
        if Cell.BorderStyle[[Side]] <> TExcelBorderStyle.None then
          Inc(SideCount);

      if (c = 1) and (r = 1) then
        Assert.AreEqual(0, SideCount, 'Center cell must have zero border sides')
      else if ((c = 0) or (c = 2)) and ((r = 0) or (r = 2)) then
        Assert.AreEqual(2, SideCount, Format('Corner cell %s must have exactly two sides', [Addr]))
      else
        Assert.AreEqual(1, SideCount, Format('Edge-midpoint cell %s must have exactly one side', [Addr]));
    end;
end;

procedure TExcelBorderSidesTests.GetBorderStyle_MultiSideQuery_ReturnsFirstMatchingSideInOrder;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  const Cell = Sheet.Cell['A1'];
  Cell.BorderStyle[[TExcelBorderSide.Top]] := TExcelBorderStyle.Thin;
  Cell.BorderStyle[[TExcelBorderSide.Left]] := TExcelBorderStyle.Thick;

  // Documented contract: returns the first matching side in Top/Right/Bottom/Left
  // declaration order -- Top wins here even though Left was also set to a different value.
  Assert.AreEqual(Ord(TExcelBorderStyle.Thin), Ord(Cell.BorderStyle[[TExcelBorderSide.Top, TExcelBorderSide.Left]]),
    'Querying [ebsTop, ebsLeft] together must return the Top value per the documented ' +
    'first-match-in-enum-order contract');
end;

procedure TExcelBorderSidesTests.SaveToFile_ColorWithoutStyle_ProducesNoBorderElement;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'X';
  Sheet.Cell['A1'].BorderColor[[TExcelBorderSide.Top]] := $FF0000; // color set, style left at None

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    // A1 is the only formatted cell in this file, so if any <top style="..."> element
    // appears anywhere in styles.xml, it can only have come from this cell -- and it
    // shouldn't, since BorderStyle was never set (color alone must not render a visible side).
    Assert.IsFalse(Pos('<top style=', StylesXml) > 0,
      'Setting BorderColor without BorderStyle must not render a visible <top> border side');
  finally
    Package.Free;
  end;
end;

procedure TExcelBorderSidesTests.LoadFromFile_AsymmetricBottomOnlyBorder_ReadsAllFourSidesIndependently;
const
  ContentTypesXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '</Types>';
  RelsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>';
  WorkbookXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>';
  // Realistic shape of what Excel itself writes for a bottom-only border: the other
  // three sides are present as empty self-closing elements, not omitted -- only
  // <bottom> carries style + color.
  StylesXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill>' +
    '<fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="2">' +
    '<border><left/><right/><top/><bottom/></border>' +
    '<border><left/><right/><top/><bottom style="thin"><color rgb="FF000000"/></bottom></border>' +
    '</borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="2">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>' +
    '</cellXfs></styleSheet>';
  // A1 self-closing with no value -- the common "empty cell under a header, styled
  // with just a bottom rule" case, which also exercises the empty-styled-cell parse path.
  SheetXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetData><row r="1"><c r="A1" s="1"/></row></sheetData></worksheet>';
begin
  var Zip := TZipFile.Create;
  try
    Zip.Open(FTempFile, zmWrite);
    Zip.Add(TEncoding.UTF8.GetBytes(ContentTypesXml), '[Content_Types].xml');
    Zip.Add(TEncoding.UTF8.GetBytes(RelsXml), '_rels/.rels');
    Zip.Add(TEncoding.UTF8.GetBytes(WorkbookXml), 'xl/workbook.xml');
    Zip.Add(TEncoding.UTF8.GetBytes(StylesXml), 'xl/styles.xml');
    Zip.Add(TEncoding.UTF8.GetBytes(SheetXml), 'xl/worksheets/sheet1.xml');
    Zip.Close;
  finally
    Zip.Free;
  end;

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  const Cell = Workbook2.Sheets[0].Cell['A1'];

  // This is the exact case ParseStyles used to get wrong -- it only ever read <left>,
  // so a bottom-only border like this would previously have been silently dropped.
  Assert.AreEqual(Ord(TExcelBorderStyle.None), Ord(Cell.BorderStyle[[TExcelBorderSide.Top]]), 'Top must read as None');
  Assert.AreEqual(Ord(TExcelBorderStyle.None), Ord(Cell.BorderStyle[[TExcelBorderSide.Left]]), 'Left must read as None');
  Assert.AreEqual(Ord(TExcelBorderStyle.None), Ord(Cell.BorderStyle[[TExcelBorderSide.Right]]), 'Right must read as None');
  Assert.AreEqual(Ord(TExcelBorderStyle.Thin), Ord(Cell.BorderStyle[[TExcelBorderSide.Bottom]]), 'Bottom must read as Thin');
  Assert.AreEqual(Cardinal($000000), Cell.BorderColor[[TExcelBorderSide.Bottom]], 'Bottom color must read correctly');
end;

procedure TExcelBorderSidesTests.RoundTrip_BottomOnlyColor_DoesNotBleedIntoOtherSides;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'X';
  // A non-black colour on a single side. Black is $000000 = 0 and so indistinguishable
  // from "no colour", which is exactly what would hide a bleed -- the other three sides
  // are written as self-closing <top/>/<left/>/<right/>, and none of them may pick up
  // this bottom colour when parsed back.
  Sheet.Cell['A1'].BorderStyle[[TExcelBorderSide.Bottom]] := TExcelBorderStyle.Thin;
  Sheet.Cell['A1'].BorderColor[[TExcelBorderSide.Bottom]] := $FF0000;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  const Cell = Workbook2.Sheets[0].Cell['A1'];

  Assert.AreEqual(Cardinal($FF0000), Cell.BorderColor[[TExcelBorderSide.Bottom]], 'Bottom colour must survive round-trip');
  Assert.AreEqual(Cardinal(0), Cell.BorderColor[[TExcelBorderSide.Top]], 'Top colour must not bleed from bottom');
  Assert.AreEqual(Cardinal(0), Cell.BorderColor[[TExcelBorderSide.Left]], 'Left colour must not bleed from bottom');
  Assert.AreEqual(Cardinal(0), Cell.BorderColor[[TExcelBorderSide.Right]], 'Right colour must not bleed from bottom');
end;

initialization
  TDUnitX.RegisterTestFixture(TExcelBorderSidesTests);

end.
