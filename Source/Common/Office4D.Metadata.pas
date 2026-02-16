unit Office4D.Metadata;

interface

uses
  System.SysUtils,
  System.DateUtils,
  System.RegularExpressions;

type
  TDocumentMetadata = record
    Title: string;
    Subject: string;
    Creator: string;
    Keywords: string;
    Description: string;
    LastModifiedBy: string;
    Revision: string;
    Created: TDateTime;
    Modified: TDateTime;
    Category: string;

    procedure Clear;
  end;

  TMetadataParser = class
  private
    function GetElementValue(const Xml, ElementName: string): string;
    function ParseW3CDateTime(const DateStr: string): TDateTime;
  public
    function Parse(const XmlContent: string): TDocumentMetadata;
  end;

implementation

{ TDocumentMetadata }

procedure TDocumentMetadata.Clear;
begin
  Title := '';
  Subject := '';
  Creator := '';
  Keywords := '';
  Description := '';
  LastModifiedBy := '';
  Revision := '';
  Created := 0;
  Modified := 0;
  Category := '';
end;

{ TMetadataParser }

function TMetadataParser.GetElementValue(const Xml, ElementName: string): string;
begin
  Result := '';
  var Pattern := '<[^>]*?' + ElementName + '[^>]*>([^<]*)</[^>]*?' + ElementName + '>';
  var Match := TRegEx.Match(Xml, Pattern, [roIgnoreCase, roSingleLine]);
  if Match.Success and (Match.Groups.Count > 1) then
    Result := Trim(Match.Groups[1].Value);
end;

function TMetadataParser.ParseW3CDateTime(const DateStr: string): TDateTime;
begin
  Result := 0;
  if DateStr = '' then
    Exit;

  Result := ISO8601ToDate(DateStr, False);
end;

function TMetadataParser.Parse(const XmlContent: string): TDocumentMetadata;
begin
  Result.Clear;

  Result.Title := GetElementValue(XmlContent, 'title');
  Result.Subject := GetElementValue(XmlContent, 'subject');
  Result.Creator := GetElementValue(XmlContent, 'creator');
  Result.Description := GetElementValue(XmlContent, 'description');

  Result.Keywords := GetElementValue(XmlContent, 'keywords');
  Result.LastModifiedBy := GetElementValue(XmlContent, 'lastModifiedBy');
  Result.Revision := GetElementValue(XmlContent, 'revision');
  Result.Category := GetElementValue(XmlContent, 'category');

  Result.Created := ParseW3CDateTime(GetElementValue(XmlContent, 'created'));
  Result.Modified := ParseW3CDateTime(GetElementValue(XmlContent, 'modified'));
end;

end.
