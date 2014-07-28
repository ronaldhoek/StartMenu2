unit AppFuncU;

interface

uses
  Windows;

  function IsLink(const aFileName: string): boolean;

  function GetSpecialFolder( const aHandle: HWND; const aFolder: Integer;
    var Location: String ): LongWord;

  procedure CheckMSCFile(var aFile: string; out aIconIndex: Word);

implementation

uses
  SysUtils, ShlObj, XMLIntf, XMLDoc, Variants;

function IsLink(const aFileName: string): boolean;
var
  aExt: string;
begin
  aExt := LowerCase( ExtractFileExt(aFileName) );
  Result := (aExt = '.lnk') or (aExt = '.url') or (aExt = '.pif');
end;

function GetSpecialFolder( const aHandle: HWND; const aFolder: Integer;
  var Location: String ): LongWord;
var
   pidl:      PItemIDList;
   hRes:      HRESULT;
   RealPath:  Array[0..MAX_PATH] of Char;
   Success:   Boolean;
begin
   Result := 0;
   hRes   := SHGetSpecialFolderLocation( aHandle, aFolder, pidl );
   if hRes = NO_ERROR then
   begin
      Success := SHGetPathFromIDList( pidl, RealPath );
      if Success then
         Location := String( RealPath ) + '\'
      else
         Result := LongWord( E_UNEXPECTED );
      GlobalFreePtr( pidl );
   end else
      Result := hRes;
end;

procedure CheckMSCFile(var aFile: string; out aIconIndex: Word);
var
  aXMLNode: IXMLNode;
  s: string;
begin
  aIconIndex := 0;

  if SameText(ExtractFileExt(aFile), '.msc') then
  begin
    // MSC-files bevatten een sectie waarin beschreven staat welk icoon
    // gebruikt moet wordern. Als deze sectie leeg is, dan moet het standaard
    // icoon van "%winsys%\mmc.exe" genomen worden.
    try
      aXMLNode := LoadXMLDocument(aFile).ChildNodes['MMC_ConsoleFile'].
        ChildNodes['VisualAttributes'].ChildNodes['Icon'];
{  <VisualAttributes>
    <Icon Index="1" File="C:\WINDOWS\system32\mmc.exe">
      <Image Name="Large" BinaryRefIndex="0"/>
      <Image Name="Small" BinaryRefIndex="1"/>
    </Icon>
  </VisualAttributes> }
      s := VarToStr(aXMLNode.Attributes['File']);
      if FileExists(s) then
      begin
        aFile := s;
        aXMLNode := aXMLNode.ChildNodes.First;
        while Assigned(aXMLNode) do
        begin
          if SameText(aXMLNode.NodeName, 'Image') then
          begin
            aIconIndex := aXMLNode.Attributes['BinaryRefIndex'];
            s := VarToStr(aXMLNode.Attributes['Name']);
            if SameText(s, 'Small') then Break;
          end;
          aXMLNode := aXMLNode.NextSibling;
        end;
      end;
    except
      // Geen geldig MSC-bestand
    end;
  end;
end;

end.
