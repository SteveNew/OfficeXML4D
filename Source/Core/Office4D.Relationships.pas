unit Office4D.Relationships;

interface

uses
  System.Classes,
  System.SysUtils,
  System.Generics.Collections,
  System.RegularExpressions;

type
  TRelationship = record
    Id: string;
    RelType: string;
    Target: string;
  end;

  TRelationshipList = TList<TRelationship>;

  TRelationships = class
  private
    FRelationships: TRelationshipList;

    function GetAttributeValue(const Element, AttrName: string): string;
  public
    constructor Create;
    destructor Destroy; override;

    procedure LoadFromXml(const XmlContent: string);
    function GetById(const Id: string): TRelationship;
    function GetByType(const RelType: string): TRelationship;
    function GetTargetByType(const RelType: string): string;
    function HasType(const RelType: string): Boolean;
    function Count: Integer;

    property Relationships: TRelationshipList read FRelationships;
  end;

implementation

uses
  Office4D.Errors,
  Office4D.Types;

{ TRelationships }

constructor TRelationships.Create;
begin
  inherited Create;
  FRelationships := TRelationshipList.Create;
end;

destructor TRelationships.Destroy;
begin
  FRelationships.Free;
  inherited;
end;

function TRelationships.GetAttributeValue(const Element, AttrName: string): string;
begin
  Result := '';
  var Pattern := AttrName + '=[''"]([^''"]*)[''"]';
  var Match := TRegEx.Match(Element, Pattern, [roIgnoreCase]);
  if Match.Success and (Match.Groups.Count > 1) then
    Result := Match.Groups[1].Value;
end;

procedure TRelationships.LoadFromXml(const XmlContent: string);
begin
  FRelationships.Clear;

  var Pattern := '<Relationship\s+[^>]*/?>';
  var Matches := TRegEx.Matches(XmlContent, Pattern, [roIgnoreCase]);

  for var Match in Matches do
  begin
    var Element := Match.Value;
    var Rel: TRelationship;
    Rel.Id := GetAttributeValue(Element, 'Id');
    Rel.RelType := GetAttributeValue(Element, 'Type');
    Rel.Target := GetAttributeValue(Element, 'Target');
    FRelationships.Add(Rel);
  end;
end;

function TRelationships.GetById(const Id: string): TRelationship;
begin
  for var Rel in FRelationships do
  begin
    if SameText(Rel.Id, Id) then
      Exit(Rel);
  end;

  raise EOfficeDocumentException.CreateFmt('Relationship not found: %s', [Id]);
end;

function TRelationships.GetByType(const RelType: string): TRelationship;
begin
  for var Rel in FRelationships do
  begin
    if SameText(Rel.RelType, RelType) then
      Exit(Rel);
  end;

  raise EOfficeDocumentException.CreateFmt('Relationship type not found: %s', [RelType]);
end;

function TRelationships.GetTargetByType(const RelType: string): string;
begin
  Result := '';
  for var Rel in FRelationships do
  begin
    if SameText(Rel.RelType, RelType) then
      Exit(Rel.Target);
  end;
end;

function TRelationships.HasType(const RelType: string): Boolean;
begin
  Result := False;
  for var Rel in FRelationships do
  begin
    if SameText(Rel.RelType, RelType) then
      Exit(True);
  end;
end;

function TRelationships.Count: Integer;
begin
  Result := FRelationships.Count;
end;

end.
