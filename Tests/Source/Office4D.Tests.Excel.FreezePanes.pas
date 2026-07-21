unit Office4D.Tests.Excel.FreezePanes;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Excel;

type
  [TestFixture]
  TExcelFreezePanesTests = class(TOffice4DTests)
  private
    FWorkbook: IExcelWorkbook;
    FTempFile: string;

  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure SaveToFile_NoFreeze_OmitsSheetViews;

    [Test]
    procedure FreezePanes_BothRowsAndColumns_SetsFrozenCounts;

    [Test]
    procedure FreezePanes_RowsOnly_LeavesColumnsAtZero;

    [Test]
    procedure FreezePanes_ColumnsOnly_LeavesRowsAtZero;

    [Test]
    procedure FreezePanes_TopLeftCorner_IsEquivalentToNoFreeze;

    [Test]
    procedure FreezePanes_InvalidAddress_RaisesException;

    [Test]
    procedure UnfreezePanes_AfterFreezing_ClearsFrozenState;

    [Test]
    procedure RoundTrip_FreezeBothRowsAndColumns_PreservesFrozenCounts;

    [Test]
    procedure SaveToFile_WithFreezePanes_ContainsCorrectPaneAttributes;

    [Test]
    procedure SaveToFile_FreezeRowsOnly_UsesBottomLeftActivePane;

    [Test]
    procedure SaveToFile_FreezeColumnsOnly_UsesTopRightActivePane;

    [Test]
    procedure SaveToFile_WithFreezePanesAndColumnWidths_SheetViewsPrecedesCols;
  end;

implementation

uses
  Office4D.Errors,
  Office4D.Package;

{ TExcelFreezePanesTests }

procedure TExcelFreezePanesTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
  FTempFile := TPath.Combine(TPath.GetTempPath, 'freezepanes_test_' + TGUID.NewGuid.ToString + '.xlsx');
end;

procedure TExcelFreezePanesTests.TearDown;
begin
  FWorkbook := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

procedure TExcelFreezePanesTests.SaveToFile_NoFreeze_OmitsSheetViews;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsFalse(Pos('<sheetViews>', SheetXml) > 0,
      'A sheet with no frozen panes must not emit a sheetViews element at all');
  finally
    Package.Free;
  end;
end;

procedure TExcelFreezePanesTests.FreezePanes_BothRowsAndColumns_SetsFrozenCounts;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.FreezePanes('C2');

  Assert.AreEqual(1, Sheet.FrozenRows, 'Row 1 should be frozen (topLeftCell row - 1)');
  Assert.AreEqual(2, Sheet.FrozenColumns, 'Columns A and B should be frozen (topLeftCell col - 1)');
end;

procedure TExcelFreezePanesTests.FreezePanes_RowsOnly_LeavesColumnsAtZero;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.FreezePanes('A2');

  Assert.AreEqual(1, Sheet.FrozenRows, 'Row 1 should be frozen');
  Assert.AreEqual(0, Sheet.FrozenColumns, 'No columns should be frozen');
end;

procedure TExcelFreezePanesTests.FreezePanes_ColumnsOnly_LeavesRowsAtZero;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.FreezePanes('B1');

  Assert.AreEqual(0, Sheet.FrozenRows, 'No rows should be frozen');
  Assert.AreEqual(1, Sheet.FrozenColumns, 'Column A should be frozen');
end;

procedure TExcelFreezePanesTests.FreezePanes_TopLeftCorner_IsEquivalentToNoFreeze;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');

  // Freezing at A1 (the top-left corner itself) has nothing above/left of it to freeze --
  // this matches Excel's own UI, where selecting A1 and choosing "Freeze Panes" is a no-op.
  Sheet.FreezePanes('A1');

  Assert.AreEqual(0, Sheet.FrozenRows);
  Assert.AreEqual(0, Sheet.FrozenColumns);
end;

procedure TExcelFreezePanesTests.FreezePanes_InvalidAddress_RaisesException;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');

  Assert.WillRaise(
    procedure
    begin
      Sheet.FreezePanes('NotACell');
    end,
    EExcelWorkbookException,
    'An unparseable cell address must raise rather than silently freezing nothing or defaulting'
  );
end;

procedure TExcelFreezePanesTests.UnfreezePanes_AfterFreezing_ClearsFrozenState;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.FreezePanes('C2');
  Sheet.UnfreezePanes;

  Assert.AreEqual(0, Sheet.FrozenRows);
  Assert.AreEqual(0, Sheet.FrozenColumns);
end;

procedure TExcelFreezePanesTests.RoundTrip_FreezeBothRowsAndColumns_PreservesFrozenCounts;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Header';
  Sheet.FreezePanes('C2');

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  const Sheet2 = Workbook2.Sheets[0];

  Assert.AreEqual(1, Sheet2.FrozenRows, 'Frozen row count should survive round-trip');
  Assert.AreEqual(2, Sheet2.FrozenColumns, 'Frozen column count should survive round-trip');
end;

procedure TExcelFreezePanesTests.SaveToFile_WithFreezePanesAndColumnWidths_SheetViewsPrecedesCols;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';
  Sheet.FreezePanes('B2');
  Sheet.SetColumnWidth('A', 20);

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    const SheetViewsPos = Pos('<sheetViews>', SheetXml);
    const ColsPos = Pos('<cols>', SheetXml);

    Assert.IsTrue(SheetViewsPos > 0, 'sheetViews element should be present');
    Assert.IsTrue(ColsPos > 0, 'cols element should be present');
    // OOXML's schema is order-sensitive here -- sheetViews must precede cols, or the
    // file can render incorrectly (or not open) in strict readers even though it's
    // well-formed XML. This ordering bug has bitten other libraries in exactly this spot.
    Assert.IsTrue(SheetViewsPos < ColsPos, 'sheetViews must be written before cols in the worksheet XML');
  finally
    Package.Free;
  end;
end;

procedure TExcelFreezePanesTests.SaveToFile_WithFreezePanes_ContainsCorrectPaneAttributes;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';
  Sheet.FreezePanes('C2');

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsTrue(Pos('xSplit="2"', SheetXml) > 0, 'xSplit should equal the frozen column count');
    Assert.IsTrue(Pos('ySplit="1"', SheetXml) > 0, 'ySplit should equal the frozen row count');
    Assert.IsTrue(Pos('topLeftCell="C2"', SheetXml) > 0, 'topLeftCell should be the first scrollable cell');
    Assert.IsTrue(Pos('state="frozen"', SheetXml) > 0, 'state should be frozen, not split');
  finally
    Package.Free;
  end;
end;

procedure TExcelFreezePanesTests.SaveToFile_FreezeRowsOnly_UsesBottomLeftActivePane;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';
  Sheet.FreezePanes('A2'); // rows only

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    // Freezing rows only puts the active pane at bottom-left, not bottom-right -- getting
    // this wrong doesn't corrupt the file, but the scroll/selection behavior around the
    // freeze boundary misbehaves in Excel.
    Assert.IsTrue(Pos('activePane="bottomLeft"', SheetXml) > 0,
      'Row-only freeze must use activePane="bottomLeft"');
    Assert.IsFalse(Pos('activePane="bottomRight"', SheetXml) > 0,
      'Row-only freeze must not use the both-axes activePane value');
  finally
    Package.Free;
  end;
end;

procedure TExcelFreezePanesTests.SaveToFile_FreezeColumnsOnly_UsesTopRightActivePane;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';
  Sheet.FreezePanes('B1'); // columns only

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsTrue(Pos('activePane="topRight"', SheetXml) > 0,
      'Column-only freeze must use activePane="topRight"');
    Assert.IsFalse(Pos('activePane="bottomRight"', SheetXml) > 0,
      'Column-only freeze must not use the both-axes activePane value');
  finally
    Package.Free;
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TExcelFreezePanesTests);

end.
