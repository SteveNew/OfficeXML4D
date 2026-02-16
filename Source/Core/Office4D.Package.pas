unit Office4D.Package;

interface

uses
  System.Classes,
  System.SysUtils,
  System.Zip;

type
  TOXMLPackage = class
  private
    FFileName: string;
    FIsOpen: Boolean;
    FZipFile: TZipFile;
    FOwnsStream: Boolean;
    FStream: TStream;

    function NormalizePath(const PartName: string): string;
  public
    constructor Create;
    destructor Destroy; override;

    procedure Open(const FileName: string); overload;
    procedure Open(const Stream: TStream); overload;
    procedure Close;
    function PartExists(const PartName: string): Boolean;
    function GetPartStream(const PartName: string): TStream;
    function GetPartContent(const PartName: string): string;

    property FileName: string read FFileName;
    property IsOpen: Boolean read FIsOpen;
  end;

implementation

uses
  Office4D.Errors;

{ TOXMLPackage }

constructor TOXMLPackage.Create;
begin
  inherited Create;
  FIsOpen := False;
  FZipFile := nil;
  FStream := nil;
  FOwnsStream := False;
end;

destructor TOXMLPackage.Destroy;
begin
  Close;
  inherited;
end;

function TOXMLPackage.NormalizePath(const PartName: string): string;
begin
  Result := StringReplace(PartName, '\', '/', [rfReplaceAll]);
end;

procedure TOXMLPackage.Open(const FileName: string);
begin
  Close;

  if not FileExists(FileName) then
    raise EPackageNotFound.CreateFmt('Package file not found: %s', [FileName]);

  FZipFile := TZipFile.Create;
  try
    FZipFile.Open(FileName, zmRead);
  except
    on E: Exception do
    begin
      FreeAndNil(FZipFile);
      raise EPackageInvalid.CreateFmt('Invalid package file: %s (%s)', [FileName, E.Message]);
    end;
  end;

  FFileName := FileName;
  FIsOpen := True;
end;

procedure TOXMLPackage.Open(const Stream: TStream);
begin
  Close;

  FZipFile := TZipFile.Create;
  try
    FZipFile.Open(Stream, zmRead);
  except
    on E: Exception do
    begin
      FreeAndNil(FZipFile);
      raise EPackageInvalid.CreateFmt('Invalid package stream: %s', [E.Message]);
    end;
  end;

  FFileName := '';
  FStream := Stream;
  FOwnsStream := False;
  FIsOpen := True;
end;

procedure TOXMLPackage.Close;
begin
  if FZipFile <> nil then
  begin
    if FZipFile.Mode <> zmClosed then
      FZipFile.Close;
    FreeAndNil(FZipFile);
  end;
  FFileName := '';
  FStream := nil;
  FOwnsStream := False;
  FIsOpen := False;
end;

function TOXMLPackage.PartExists(const PartName: string): Boolean;
begin
  Result := False;

  if not FIsOpen then
    Exit;

  var NormalizedName := NormalizePath(PartName);

  for var I := 0 to FZipFile.FileCount - 1 do
  begin
    if SameText(NormalizePath(FZipFile.FileNames[I]), NormalizedName) then
      Exit(True);
  end;
end;

function TOXMLPackage.GetPartStream(const PartName: string): TStream;
begin
  if not FIsOpen then
    raise EPackageException.Create('Package is not open');

  var NormalizedName := NormalizePath(PartName);

  for var I := 0 to FZipFile.FileCount - 1 do
  begin
    if SameText(NormalizePath(FZipFile.FileNames[I]), NormalizedName) then
    begin
      var Bytes: TBytes;
      FZipFile.Read(I, Bytes);
      Exit(TBytesStream.Create(Bytes));
    end;
  end;

  raise EPartNotFound.CreateFmt('Part not found in package: %s', [PartName]);
end;

function TOXMLPackage.GetPartContent(const PartName: string): string;
begin
  var Stream := GetPartStream(PartName);
  try
    var Reader := TStreamReader.Create(Stream, TEncoding.UTF8);
    try
      Result := Reader.ReadToEnd;
    finally
      Reader.Free;
    end;
  finally
    Stream.Free;
  end;
end;

end.
