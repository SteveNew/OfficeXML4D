unit Office4D.Word;

interface

uses
  System.Classes,
  System.SysUtils,
  System.Generics.Collections,
  Office4D.Metadata;

type
  {$SCOPEDENUMS ON}
  TListStyle = (None, Bullet, Numbered);
  TPageOrientation = (Portrait, Landscape);
  TParagraphAlignment = (Left, Center, Right, Justify);
  TLineSpacingRule = (Auto, Exact, AtLeast);
  TBorderStyle = (None, Single, Double, Dashed, Dotted, Thick);
  {$SCOPEDENUMS OFF}

  TLineSpacing = record
    Line: Integer;
    Rule: TLineSpacingRule;
    Before: Integer;
    After: Integer;
    class function Create(ALine: Integer; ARule: TLineSpacingRule; ABefore, AAfter: Integer): TLineSpacing; static;
    class function Single: TLineSpacing; static;
    class function OneAndHalf: TLineSpacing; static;
    class function Double: TLineSpacing; static;
  end;

  TParagraphIndent = record
    Left: Integer;
    Right: Integer;
    FirstLine: Integer;
    class function Create(ALeft, ARight, AFirstLine: Integer): TParagraphIndent; static;
    class function HalfInch: TParagraphIndent; static;
  end;

  TTableBorder = record
    Style: TBorderStyle;
    Width: Integer;
    Color: string;
    class function Create(AStyle: TBorderStyle; AWidth: Integer; const AColor: string): TTableBorder; static;
    class function Default: TTableBorder; static;
    class function None: TTableBorder; static;
  end;

  TTableBorders = record
    Top: TTableBorder;
    Bottom: TTableBorder;
    Left: TTableBorder;
    Right: TTableBorder;
    InsideH: TTableBorder;
    InsideV: TTableBorder;
    class function Box(const Border: TTableBorder): TTableBorders; static;
    class function All(const Border: TTableBorder): TTableBorders; static;
    class function None: TTableBorders; static;
  end;

  TPageMargins = record
    Top: Integer;
    Bottom: Integer;
    Left: Integer;
    Right: Integer;
    class function Create(const Top, Bottom, Left, Right: Integer): TPageMargins; static;
  end;

  IWordRun = interface;
  IWordParagraph = interface;
  IWordTableCell = interface;
  IWordTable = interface;
  IWordHeaderFooter = interface;
  IWordDocument = interface;

  TWordImage = record
    Data: TBytes;
    Width: Integer;
    Height: Integer;
    Extension: string;
    class function Create(const AData: TBytes; const AExtension: string; AWidth, AHeight: Integer): TWordImage; static;
  end;

  IWordRun = interface
    ['{A1B2C3D4-1111-2222-3333-444455556666}']
    function GetText: string;
    procedure SetText(const Value: string);
    function GetBold: Boolean;
    procedure SetBold(const Value: Boolean);
    function GetItalic: Boolean;
    procedure SetItalic(const Value: Boolean);
    function GetUnderline: Boolean;
    procedure SetUnderline(const Value: Boolean);
    function GetHyperlink: string;
    procedure SetHyperlink(const Value: string);
    function GetFontName: string;
    procedure SetFontName(const Value: string);
    function GetFontSize: Integer;
    procedure SetFontSize(const Value: Integer);
    function GetFontColor: string;
    procedure SetFontColor(const Value: string);
    function GetImage: TWordImage;
    function HasImage: Boolean;
    procedure AddImage(const AData: TBytes; const AExtension: string; AWidthPx, AHeightPx: Integer);

    property Text: string read GetText write SetText;
    property Bold: Boolean read GetBold write SetBold;
    property Italic: Boolean read GetItalic write SetItalic;
    property Underline: Boolean read GetUnderline write SetUnderline;
    property Hyperlink: string read GetHyperlink write SetHyperlink;
    property FontName: string read GetFontName write SetFontName;
    property FontSize: Integer read GetFontSize write SetFontSize;
    property FontColor: string read GetFontColor write SetFontColor;
    property Image: TWordImage read GetImage;
  end;

  IWordParagraph = interface
    ['{A1B2C3D4-2222-3333-4444-555566667777}']
    function GetText: string;
    function GetRunCount: Integer;
    function GetRun(Index: Integer): IWordRun;
    function AddRun(const Text: string): IWordRun;
    procedure AddLineBreak;
    procedure AddTab;
    procedure AddPageBreak;
    function GetListStyle: TListStyle;
    procedure SetListStyle(const Value: TListStyle);
    function GetAlignment: TParagraphAlignment;
    procedure SetAlignment(const Value: TParagraphAlignment);
    function GetLineSpacing: TLineSpacing;
    procedure SetLineSpacing(const Value: TLineSpacing);
    function GetIndent: TParagraphIndent;
    procedure SetIndent(const Value: TParagraphIndent);

    property Text: string read GetText;
    property RunCount: Integer read GetRunCount;
    property Runs[Index: Integer]: IWordRun read GetRun;
    property ListStyle: TListStyle read GetListStyle write SetListStyle;
    property Alignment: TParagraphAlignment read GetAlignment write SetAlignment;
    property LineSpacing: TLineSpacing read GetLineSpacing write SetLineSpacing;
    property Indent: TParagraphIndent read GetIndent write SetIndent;
  end;

  IWordTableCell = interface
    ['{A1B2C3D4-4444-5555-6666-777788889999}']
    function GetText: string;
    procedure SetText(const Value: string);
    function GetShading: string;
    procedure SetShading(const Value: string);
    function GetWidth: Integer;
    procedure SetWidth(const Value: Integer);

    property Text: string read GetText write SetText;
    property Shading: string read GetShading write SetShading;
    property Width: Integer read GetWidth write SetWidth;
  end;

  IWordTable = interface
    ['{A1B2C3D4-5555-6666-7777-888899990000}']
    function GetRowCount: Integer;
    function GetColCount: Integer;
    function GetCell(Row, Col: Integer): IWordTableCell;
    function GetBorders: TTableBorders;
    procedure SetBorders(const Value: TTableBorders);
    procedure SetColumnWidths(const Widths: array of Integer);

    property RowCount: Integer read GetRowCount;
    property ColCount: Integer read GetColCount;
    property Cells[Row, Col: Integer]: IWordTableCell read GetCell;
    property Borders: TTableBorders read GetBorders write SetBorders;
  end;

  IWordHeaderFooter = interface
    ['{A1B2C3D4-6666-7777-8888-999900001111}']
    function GetText: string;
    procedure SetText(const Value: string);

    property Text: string read GetText write SetText;
  end;

  IWordDocument = interface
    ['{A1B2C3D4-3333-4444-5555-666677778888}']
    procedure LoadFromFile(const FileName: string);
    procedure LoadFromStream(const Stream: TStream);
    procedure SaveToFile(const FileName: string);
    procedure SaveToStream(const Stream: TStream);

    function GetText: string;
    function GetMetadata: TDocumentMetadata;

    function GetParagraphCount: Integer;
    function GetParagraph(Index: Integer): IWordParagraph;
    function AddParagraph: IWordParagraph;

    function GetTableCount: Integer;
    function GetTable(Index: Integer): IWordTable;
    function AddTable(const Rows, Cols: Integer): IWordTable;

    property Text: string read GetText;
    property Metadata: TDocumentMetadata read GetMetadata;
    property ParagraphCount: Integer read GetParagraphCount;
    property Paragraphs[Index: Integer]: IWordParagraph read GetParagraph;
    property TableCount: Integer read GetTableCount;
    property Tables[Index: Integer]: IWordTable read GetTable;

    function GetPageOrientation: TPageOrientation;
    procedure SetPageOrientation(const Value: TPageOrientation);
    function GetPageMargins: TPageMargins;
    procedure SetPageMargins(const Value: TPageMargins);
    function GetHeader: IWordHeaderFooter;
    function GetFooter: IWordHeaderFooter;

    property PageOrientation: TPageOrientation read GetPageOrientation write SetPageOrientation;
    property PageMargins: TPageMargins read GetPageMargins write SetPageMargins;
    property Header: IWordHeaderFooter read GetHeader;
    property Footer: IWordHeaderFooter read GetFooter;
  end;

  TWordDocumentFactory = class
  public
    class function CreateDocument: IWordDocument;
  end;

implementation

uses
  Office4D.Word.Document;

class function TWordDocumentFactory.CreateDocument: IWordDocument;
begin
  Result := TWordDocument.Create;
end;

{ TPageMargins }

class function TPageMargins.Create(const Top, Bottom, Left, Right: Integer): TPageMargins;
begin
  Result.Top := Top;
  Result.Bottom := Bottom;
  Result.Left := Left;
  Result.Right := Right;
end;

{ TLineSpacing }

class function TLineSpacing.Create(ALine: Integer; ARule: TLineSpacingRule; ABefore, AAfter: Integer): TLineSpacing;
begin
  Result.Line := ALine;
  Result.Rule := ARule;
  Result.Before := ABefore;
  Result.After := AAfter;
end;

class function TLineSpacing.Single: TLineSpacing;
begin
  Result := TLineSpacing.Create(240, TLineSpacingRule.Auto, 0, 0);
end;

class function TLineSpacing.OneAndHalf: TLineSpacing;
begin
  Result := TLineSpacing.Create(360, TLineSpacingRule.Auto, 0, 0);
end;

class function TLineSpacing.Double: TLineSpacing;
begin
  Result := TLineSpacing.Create(480, TLineSpacingRule.Auto, 0, 0);
end;

{ TParagraphIndent }

class function TParagraphIndent.Create(ALeft, ARight, AFirstLine: Integer): TParagraphIndent;
begin
  Result.Left := ALeft;
  Result.Right := ARight;
  Result.FirstLine := AFirstLine;
end;

class function TParagraphIndent.HalfInch: TParagraphIndent;
begin
  Result := TParagraphIndent.Create(720, 0, 0);
end;

{ TTableBorder }

class function TTableBorder.Create(AStyle: TBorderStyle; AWidth: Integer; const AColor: string): TTableBorder;
begin
  Result.Style := AStyle;
  Result.Width := AWidth;
  Result.Color := AColor;
end;

class function TTableBorder.Default: TTableBorder;
begin
  Result := TTableBorder.Create(TBorderStyle.Single, 4, '000000');
end;

class function TTableBorder.None: TTableBorder;
begin
  Result := TTableBorder.Create(TBorderStyle.None, 0, '');
end;

{ TTableBorders }

class function TTableBorders.Box(const Border: TTableBorder): TTableBorders;
begin
  Result.Top := Border;
  Result.Bottom := Border;
  Result.Left := Border;
  Result.Right := Border;
  Result.InsideH := TTableBorder.None;
  Result.InsideV := TTableBorder.None;
end;

class function TTableBorders.All(const Border: TTableBorder): TTableBorders;
begin
  Result.Top := Border;
  Result.Bottom := Border;
  Result.Left := Border;
  Result.Right := Border;
  Result.InsideH := Border;
  Result.InsideV := Border;
end;

class function TTableBorders.None: TTableBorders;
begin
  Result.Top := TTableBorder.None;
  Result.Bottom := TTableBorder.None;
  Result.Left := TTableBorder.None;
  Result.Right := TTableBorder.None;
  Result.InsideH := TTableBorder.None;
  Result.InsideV := TTableBorder.None;
end;

{ TWordImage }

class function TWordImage.Create(const AData: TBytes; const AExtension: string; AWidth, AHeight: Integer): TWordImage;
begin
  Result.Data := AData;
  Result.Extension := AExtension;
  Result.Width := AWidth;
  Result.Height := AHeight;
end;

end.
