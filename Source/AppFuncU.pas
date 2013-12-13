unit AppFuncU;

interface

uses
  Windows;

  function GetSpecialFolder( const aHandle: HWND; const aFolder: Integer;
    var Location: String ): LongWord;

implementation

uses
  ShlObj;

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

end.
