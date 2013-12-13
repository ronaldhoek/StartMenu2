{****************************************************}
{        TWallpaper Component for Delphi 16/32       }
{      Copyright (c) 1998, UtilMind Solutions        }
{             Home Page: www.utilmind.com            }
{              E-Mail: info@utilmind.com             }
{****************************************************}
{                 IMPORTANT NOTE:                    }
{  This code may be used and modified by anyone so   }
{ long as this header and copyright information      }
{ remains intact. By using this code you agree to    }
{ indemnify UtilMind Solutions from any liability    }
{ that might arise from its use. You must obtain     }
{ written consent before selling or redistributing   }
{ this code                                          }
{****************************************************}

unit Walpaper;

interface

uses
  {$IFDEF WIN32} Windows, Registry, {$ELSE} WinTypes, WinProcs, IniFiles, {$ENDIF}
  Classes, Controls, SysUtils;

type
  TWallPaper = class(TComponent)
  private
    PC: Array[0..$FF] of Char;
{$IFDEF WIN32}
    Reg: TRegistry;
{$ELSE}
    Reg: TIniFile;
    WinIniPath: String;
{$ENDIF}

    function GetWallpaper: String;
    procedure SetWallpaper(Value: String);
    function GetTile: Boolean;
    procedure SetTile(Value: Boolean);
    function GetStretch: Boolean;
    procedure SetStretch(Value: Boolean);
  public
{$IFNDEF WIN32}
    constructor Create(aOwner: TComponent); override;
{$ENDIF}
  published
    property Wallpaper: String read GetWallpaper write SetWallpaper;
    property Tile: Boolean read GetTile write SetTile;
    property Stretch: Boolean read GetStretch write SetStretch;
  end;

implementation

{$IFNDEF WIN32}
constructor TWallpaper.Create(aOwner: TComponent);
begin
  inherited Create(aOwner);
  GetWindowsDirectory(PC, $FF);
  WinIniPath := StrPas(PC) + '\WIN.INI';
end;
{$ENDIF}

function TWallpaper.GetWallpaper: String;
begin
{$IFDEF WIN32}
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    Reg.OpenKey('\Control Panel\desktop\', False);
    Result := Reg.ReadString('Wallpaper');
{$ELSE}
  Reg := TIniFile.Create(WinIniPath);
  try
    Result := Reg.ReadString('Desktop', 'Wallpaper', '');
{$ENDIF}
  finally
    Reg.Free;
  end;
end;

procedure TWallpaper.SetWallpaper(Value: String);
begin
  if not (csDesigning in ComponentState) and
     not (csLoading in ComponentState) and
     not (csReading in ComponentState) then
  begin
    StrPCopy(PC, Value);
    SystemParametersInfo(spi_SetDeskWallpaper, 0, @PC, spif_UpdateIniFile);
  end;
end;

function TWallpaper.GetTile: Boolean;
begin
{$IFDEF WIN32}
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    Reg.OpenKey('\Control Panel\desktop\', False);
    Result := Boolean(StrToInt(Reg.ReadString('TileWallpaper')));
{$ELSE}
  Reg := TIniFile.Create(WinIniPath);
  try
    Result := Reg.ReadBool('Desktop', 'TileWallpaper', False);
{$ENDIF}
  finally
    Reg.Free;
  end;
end;

procedure TWallpaper.SetTile(Value: Boolean);
begin
  if not (csDesigning in ComponentState) and
     not (csLoading in ComponentState) and
     not (csReading in ComponentState) then
  begin
{$IFDEF WIN32}
    Reg := TRegistry.Create;
    try
      Reg.RootKey := HKEY_CURRENT_USER;
      Reg.OpenKey('\Control Panel\desktop\', False);
      Reg.WriteString('TileWallpaper', IntToStr(Integer(Value)));
{$ELSE}
    Reg := TIniFile.Create(WinIniPath);
    try
      Reg.WriteBool('Desktop', 'TileWallpaper', Value);
{$ENDIF}
    finally
      Reg.Free;
    end;
    SetWallpaper(Wallpaper);
  end;
end;

function TWallpaper.GetStretch: Boolean;
var
  i: Integer;
begin
{$IFDEF WIN32}
  Reg := TRegistry.Create;
  try
    try
      Reg.RootKey := HKEY_CURRENT_USER;
      Reg.OpenKey('\Control Panel\desktop\', False);
      i := StrToInt(Reg.ReadString('WallpaperStyle'));
    except
      i := 1;
    end;
{$ELSE}
  Reg := TIniFile.Create(WinIniPath);
  try
    i := Reg.ReadInteger('Desktop', 'WallpaperStyle', 0);
{$ENDIF}
  finally
    Reg.Free;
  end;
  Result := i = 2;
end;

procedure TWallpaper.SetStretch(Value: Boolean);
var
  v: Integer;
begin
  if not (csDesigning in ComponentState) and
     not (csLoading in ComponentState) and
     not (csReading in ComponentState) then
  begin
    if Value then v := 2 else v := 0;
{$IFDEF WIN32}
    Reg := TRegistry.Create;
    try
      Reg.RootKey := HKEY_CURRENT_USER;
      Reg.OpenKey('\Control Panel\desktop\', False);
      Reg.WriteString('WallpaperStyle', IntToStr(v));
{$ELSE}
    Reg := TIniFile.Create(WinIniPath);
    try
      Reg.WriteInteger('Desktop', 'WallpaperStyle', v);
{$ENDIF}
    finally
      Reg.Free;
    end;
    SetWallpaper(Wallpaper);
  end;
end;

end.
