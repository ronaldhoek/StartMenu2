{$I DebugSettings.inc}
unit MainU;

interface

uses
  Forms, Walpaper, Dialogs, ExtDlgs, Menus, Classes, ExtCtrls, ActnList,
  ImgList, Controls, StdActns, StdCtrls, ComCtrls, Vcl.Mask, JvExMask, JvSpin;

type
  TfrmMain = class(TForm)
    actnAutoChange: TAction;
    actnBeheerMenu: TAction;
    actnChangeWallPaper: TAction;
    actnClose: TAction;
    actnRefresh: TAction;
    actnStartUpAll: TAction;
    actnStartupCommonOnly: TAction;
    actnStartupUserOnly: TAction;
    actnWallPaperStrech: TAction;
    actnWallPaperTile: TAction;
    Alleengedeeldeitems1: TMenuItem;
    Alleenhuidigegebruiker1: TMenuItem;
    alMain: TActionList;
    Automatischwijzigen1: TMenuItem;
    Beheermenu1: TMenuItem;
    Beide1: TMenuItem;
    Cancel1: TMenuItem;
    ChangeWP1: TMenuItem;
    Close1: TMenuItem;
    dlgWallPaper: TOpenPictureDialog;
    ilMenu_Default: TImageList;
    ilMenu_LMB: TImageList;
    ilMenu_StartUp: TImageList;
    Image1: TImage;
    mnuExecStartup: TMenuItem;
    mnuSepMenuList: TMenuItem;
    mnuStartMenuItems: TMenuItem;
    N1: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    pmnuLMB: TPopupMenu;
    pmnuRMB: TPopupMenu;
    edtIconIndex: TJvSpinEdit;
    Stretch1: TMenuItem;
    Tile1: TMenuItem;
    tiMain: TTrayIcon;
    tmrWPChange: TTimer;
    tvDebug: TTreeView;
    Ververslijst1: TMenuItem;
    procedure actnAutoChangeExecute(Sender: TObject);
    procedure actnAutoChangeUpdate(Sender: TObject);
    procedure actnBeheerMenuExecute(Sender: TObject);
    procedure actnChangeWallPaperExecute(Sender: TObject);
    procedure actnCloseExecute(Sender: TObject);
    procedure actnRefreshExecute(Sender: TObject);
    procedure actnStartUpAllExecute(Sender: TObject);
    procedure actnStartUpActionUpdate(Sender: TObject);
    procedure actnStartupCommonOnlyExecute(Sender: TObject);
    procedure actnStartupUserOnlyExecute(Sender: TObject);
    procedure actnWallPaperStrechExecute(Sender: TObject);
    procedure actnWallPaperStrechUpdate(Sender: TObject);
    procedure actnWallPaperTileExecute(Sender: TObject);
    procedure actnWallPaperTileUpdate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure tiMainClick(Sender: TObject);
    procedure tmrWPChangeTimer(Sender: TObject);
    procedure tvDebugClick(Sender: TObject);
  private
    FBaseMenuItems: TList;
    FFullLinkNames: TStringList;
    FLastWPChange: TDateTime;
    FLoadingCount: Integer;
    FPopulateThread: TThread;
    FStartUpCommon: string;
    FStartUpItems: TStrings;
    FStartUpUser: string;
    wpEditer: TWallPaper;
    function AppIniFile: string;
    function AppPath: string;
    procedure ChangeWallPaperIfWanted;
    procedure ExeStartUpItem(const Index: integer; const aFolderID: Integer = 0);
    function GetLoading: Boolean;
    procedure OnMenuItemClick(Sender: TObject);
    procedure OnMenuItemStartUpClick(Sender: TObject);
    procedure OnPopulateDone(Sender: TObject);
    procedure PopulateDebugTree;
    procedure PopulateItems;
  protected
    procedure AddLoading;
    procedure LoadingChanged; virtual;
    procedure RemoveLoading;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    property Loading: Boolean read GetLoading;
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.DFM}
uses
  Windows, SysUtils, ShlObj, JPeg, Graphics, ShellApi, DateUtils, ActiveX,
  MY_Regutil, MY_RestUtil, MY_WinUtil, MY_DirList, AppFuncU;

const
  sSectionWallPaper = 'WallPaper';
  sIdentAutoChange = 'AutoChange';
  sIdentAutoChangeInterval = 'AutoChangeInterval';

type
  TMenuPopulateThread = class(TThread)
  private
    FForm: TfrmMain;
    function GetImageIndex(aLinkFile: string; ImageList: TImageList): integer;
    procedure PopulateStartMenu;
    procedure PopulateStartUpList;
    function StartUpFileName(const Index: integer): string;
  protected
    procedure Execute; override;
  public
    constructor Create(aForm: TfrmMain);
  end;

{ TMenuPopulateThread }

constructor TMenuPopulateThread.Create(aForm: TfrmMain);
begin
  inherited Create;
  FForm := aForm;
  Self.OnTerminate := FForm.OnPopulateDone;
  Self.FreeOnTerminate := True;

  // DIT IS NIET SYNC MET DE MAIN (VCL) THREAD!!!

  // Verwijderen toegevoegd menuitems
  while aForm.FBaseMenuItems.Count > 0 do
  begin
    TObject(aForm.FBaseMenuItems[0]).Free;
    aForm.FBaseMenuItems.Delete(0);
  end;
  // Neem standaard iconen over
  aForm.ilMenu_LMB.Assign(aForm.ilMenu_Default);

  // Verwijderen startmenu items
  aForm.mnuStartMenuItems.Clear;
  // Neem standaard iconen over
  aForm.ilMenu_StartUp.Assign(aForm.ilMenu_Default);
end;

procedure TMenuPopulateThread.Execute;
begin
  CoInitialize(nil);
  try
    PopulateStartMenu;
    PopulateStartUpList;
  finally
    CoUninitialize;
  end;
end;

function TMenuPopulateThread.GetImageIndex(aLinkFile: string; ImageList:
    TImageList): integer;
var
  iIcon: Word;
  aIcon: HIcon;
  Icon: TIcon;
  idxIcon: integer;
begin
  CheckMSCFile(aLinkFile, iIcon);

  aIcon := ExtractAssociatedIcon(MainInstance,PChar(aLinkFile),iIcon);
  if aIcon <> 0 then
  begin
    Icon := TIcon.Create;
    try
      Icon.Handle := aIcon;
      // Dit MOET even synchroon met VCL thread
      Synchronize(
        procedure begin
          idxIcon := ImageList.AddIcon(Icon);
        end);
      Result := idxIcon;
      // Result := ImageList.Count-1;
    finally
      FreeAndNil(Icon);
    end;
  end else
    if IsLink(aLinkFile) then
      Result:= 2 // Link icoon
    else Result:= 3; // Onbekend icoon
end;

procedure TMenuPopulateThread.PopulateStartMenu;

  procedure InternalPopulate(aDirList: IDirectoryList; mnuItem: TMenuItem);

    function CreateMenuItem(const aFileName: string; const Directory: boolean): TMenuItem;
    var
      aNewMenuItem: TMenuItem;
    begin
      aNewMenuItem := TMenuItem.Create(FForm);
      if Directory then
      begin
        aNewMenuItem.Caption := aFileName;
        aNewMenuItem.ImageIndex := 1; // Directory icoon
      end else
      begin
        aNewMenuItem.Caption := ChangeFileExt(aFileName, ''); // extentie eraf laten
        aNewMenuItem.Tag := FForm.FFullLinkNames.Add(IncludeTrailingPathDelimiter(aDirList.Directory) + aFileName);
        aNewMenuItem.ImageIndex := GetImageIndex(FForm.FFullLinkNames[aNewMenuItem.Tag], FForm.ilMenu_LMB);
        aNewMenuItem.OnClick := FForm.OnMenuItemClick;
      end;

      // Toevoegen aan menu via Sync doen!
      Synchronize(procedure begin
          if Assigned(mnuItem) then
            mnuItem.Add(aNewMenuItem)
          else begin
            FForm.pmnuLMB.Items.Insert(FForm.pmnuLMB.Items.IndexOf(FForm.mnuSepMenuList), aNewMenuItem);
            FForm.FBaseMenuItems.Add(aNewMenuItem);
          end
        end);

      Result := aNewMenuItem;
    end;

  var
    n: integer;
  begin
    if not Assigned(aDirList) then
      Exit;

    for n := 0 to aDirList.Count-1 do
      with aDirList.Items[n] do
    case aType of
      itDirectory: InternalPopulate(SubItems, CreateMenuItem(Name, True) );
      itFile: CreateMenuItem(Name, False);
    end;
  end;

var
  aDL: IDirectoryList;
begin
  aDL := NewDirectoryList(FForm.AppPath + 'Items', '*');
  InternalPopulate(aDL, nil);
end;

procedure TMenuPopulateThread.PopulateStartUpList;

  procedure AddFoldersToList(const aFolderID: integer; var aFolder: string);
  var
    n: integer;
  begin
    if GetSpecialFolder(Self.Handle, aFolderID, aFolder) = NO_ERROR then
      with NewDirectoryList(aFolder, '*.lnk') do
    begin
      aFolder := IncludeTrailingPathDelimiter(aFolder);
      for n:=0 to Count-1 do
        with Items[n] do
          if aType = itFile then
            FForm.FStartUpItems.AddObject(Name, Pointer(aFolderID));
    end else
      aFolder := '';
  end;

var
  n: integer;
  aMnuItem: TMenuItem;
begin
  AddFoldersToList(CSIDL_COMMON_STARTUP, FForm.FStartUpCommon);
  AddFoldersToList(CSIDL_STARTUP, FForm.FStartUpUser);

  for n:=0 to FForm.FStartUpItems.Count-1 do
  begin
    aMnuItem := TMenuItem.Create(FForm);
    with aMnuItem do
    begin
      OnClick := FForm.OnMenuItemStartUpClick;
      Caption := FForm.FStartUpItems[n];
      if IsLink(Caption) then
        Caption := Copy(Caption, 1, Length(Caption)-4);
      Tag := n;
      ImageIndex := GetImageIndex(StartUpFileName(n), FForm.ilMenu_StartUp);
    end;

    // Toevoegen aan menu via Sync doen!
    Synchronize(procedure begin
        FForm.mnuStartMenuItems.Add(aMnuItem);
      end);
  end;
end;

function TMenuPopulateThread.StartUpFileName(const Index: integer): string;
begin
  case Integer(FForm.FStartUpItems.Objects[Index]) of
    CSIDL_COMMON_STARTUP:
      Result := FForm.FStartUpCommon + FForm.FStartUpItems[Index];
    CSIDL_STARTUP:
      Result := FForm.FStartUpUser + FForm.FStartUpItems[Index];
    else
      Raise Exception.CreateFmt('Ongeldig opstartmenu item "%s"', [FForm.FStartUpItems[Index]]);
  end;
end;

{ TfrmMain }

procedure TfrmMain.actnAutoChangeExecute(Sender: TObject);
begin
  WriteIni( AppIniFile, sSectionWallPaper, sIdentAutoChange, (Sender as TAction).Checked );
end;

procedure TfrmMain.actnAutoChangeUpdate(Sender: TObject);
begin
  (Sender as TAction).Checked :=
    ReadIni( AppIniFile, sSectionWallPaper, sIdentAutoChange, RegBool );
end;

procedure TfrmMain.actnBeheerMenuExecute(Sender: TObject);
begin
  if ForceDirectories(AppPath + 'Items') then
    ShellExecute(Self.Handle, nil, 'explorer.exe',
      PChar('/e,/root,' + AppPath + 'Items'), nil, SW_SHOW);
end;

procedure TfrmMain.actnChangeWallPaperExecute(Sender: TObject);
const
  sWPFileName = 'WP_IconImage.bmp';
var
  aJpeg:TJPEGImage;
  aBmp:TBitmap;
  aExt:string;
  aPath:array [0..MAX_PATH] of Char;
begin
  with dlgWallPaper do
    If Execute then
  begin
    aExt := LowerCase( ExtractFileExt(FileName) );
    if (aExt='.jpg') or (aExt='.jpeg') then
    begin
      aJpeg:=TJPEGImage.Create;
      try
        aJpeg.LoadFromFile(FileName);
        aBmp:=TBitmap.Create;
        try
          aBmp.Assign(aJpeg);
          GetWindowsDirectory(aPath, MAX_PATH);
          aBmp.SaveToFile( IncludeTrailingPathDelimiter(aPath) + sWPFileName );
        finally
          FreeAndNil(aBmp);
        end;
      finally
        FreeAndNil(aJpeg);
      end;
      wpEditer.Wallpaper := IncludeTrailingPathDelimiter(aPath) + sWPFileName;
    end else
      wpEditer.Wallpaper := FileName;
  end;
end;

procedure TfrmMain.actnCloseExecute(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TfrmMain.actnRefreshExecute(Sender: TObject);
begin
  PopulateItems;
end;

procedure TfrmMain.actnStartUpAllExecute(Sender: TObject);
var
  n: integer;
begin
  if Loading then Exit;
  for n:=0 to FStartUpItems.Count-1 do
    ExeStartUpItem(n);
end;

procedure TfrmMain.actnStartupCommonOnlyExecute(Sender: TObject);
var
  n: integer;
begin
  if Loading then Exit;
  for n:=0 to FStartUpItems.Count-1 do
    ExeStartUpItem(n, CSIDL_COMMON_STARTUP);
end;

procedure TfrmMain.actnStartupUserOnlyExecute(Sender: TObject);
var
  n: integer;
begin
  if Loading then Exit;
  for n:=0 to FStartUpItems.Count-1 do
    ExeStartUpItem(n, CSIDL_STARTUP);
end;

procedure TfrmMain.actnWallPaperStrechExecute(Sender: TObject);
begin
  wpEditer.Stretch := (Sender as TAction).Checked;
end;

procedure TfrmMain.actnWallPaperStrechUpdate(Sender: TObject);
begin
  (Sender as TAction).Checked := wpEditer.Stretch;
end;

procedure TfrmMain.actnWallPaperTileExecute(Sender: TObject);
begin
  wpEditer.Tile := (Sender as TAction).Checked;
end;

procedure TfrmMain.actnWallPaperTileUpdate(Sender: TObject);
begin
  (Sender as TAction).Checked := wpEditer.Tile;
end;

function TfrmMain.AppIniFile: string;
begin
  Result := ChangeFileExt(Application.ExeName, '.ini');
end;

function TfrmMain.AppPath: string;
begin
  Result := ExtractFilePath(Application.ExeName);
end;

procedure TfrmMain.ChangeWallPaperIfWanted;
begin
  if ReadIni( AppIniFile, sSectionWallPaper, sIdentAutoChange, RegBool ) and
     ( Now > IncMinute(FLastWPChange, ReadIni( AppIniFile, sSectionWallPaper, sIdentAutoChangeInterval, RegInt )) ) then
  begin
    // Selecteer volgende wallpaper
  end;
end;

constructor TfrmMain.Create(AOwner: TComponent);
begin
  FLastWPChange := 0;
  FFullLinkNames:= TStringList.Create;
  FBaseMenuItems := TList.Create;
  FStartUpItems := TStringList.Create;
  TStringList(FStartUpItems).Sorted := True;
  inherited;
  wpEditer := TWallPaper.Create(self);
end;

destructor TfrmMain.Destroy;
begin
  FreeAndNil(FFullLinkNames);
  FreeAndNil(FBaseMenuItems);
  FreeAndNil(FStartUpItems);
  inherited;
end;

procedure TfrmMain.actnStartUpActionUpdate(Sender: TObject);
begin
  (Sender as TCustomAction).Enabled := not Loading;
end;

procedure TfrmMain.AddLoading;
begin
  Inc(FLoadingCount);
  if FLoadingCount = 1 then
    LoadingChanged;
end;

procedure TfrmMain.ExeStartUpItem(const Index: integer; const aFolderID:
    Integer = 0);
var
  aFile: string;
  iItemFolderID: Integer;
begin
  if Loading then Exit;

  iItemFolderID := Integer(FStartUpItems.Objects[Index]);
  if (aFolderID <> 0) and (aFolderID <> iItemFolderID) then
    Exit; // Ongeldig item

  case iItemFolderID of
    CSIDL_COMMON_STARTUP:
      aFile := FStartUpCommon + FStartUpItems[Index];
    CSIDL_STARTUP:
      aFile := FStartUpUser + FStartUpItems[Index];
  else
    Exit; // Niets doen
  end;

  ShellExecute(Self.Handle, nil, PChar(aFile), nil, nil, SW_SHOW);
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  tiMain.Icon := Application.Icon;
  PopulateItems;
  ChangeWallPaperIfWanted;
end;

function TfrmMain.GetLoading: Boolean;
begin
  Result := FLoadingCount > 0;
end;

procedure TfrmMain.LoadingChanged;
begin
  // Overerf om een actie te ondernemen zodra Loading property wijzigd
  if Loading then
    tiMain.Hint := 'Bezig met opbouwen...'
  else
    tiMain.Hint := 'Je tweede Startmenu';
end;

procedure TfrmMain.OnMenuItemClick(Sender: TObject);
begin
  ShellExecute(Self.Handle, nil, PChar(FFullLinkNames[(Sender as TMenuItem).Tag]),
    nil, nil, SW_SHOW);
end;

procedure TfrmMain.OnMenuItemStartUpClick(Sender: TObject);
begin
  ExeStartUpItem( (Sender as TMenuItem).Tag );
end;

procedure TfrmMain.OnPopulateDone(Sender: TObject);
begin
  if FPopulateThread = Sender then
  begin
    FPopulateThread := nil;

    // Debug treeview
  {$IFDEF DEBUG}
    PopulateDebugTree;
  {$ENDIF}

    RemoveLoading;
  end;
end;

procedure TfrmMain.PopulateDebugTree;

  procedure AddMenuItem(const aItem: TMenuItem; aParent: TTreeNode);
  var
    n: Integer;
  begin
    aParent := tvDebug.Items.AddChild(aParent, aItem.Caption);
    aParent.ImageIndex := aItem.ImageIndex;
    aParent.SelectedIndex := aItem.ImageIndex;

    if aParent.ImageIndex = 1 then // directory
    begin
      for n:=0 to aItem.Count-1 do
        AddMenuItem(aItem.Items[n], aParent);
    end else
      aParent.Data := Pointer(aItem.Tag);
  end;

var
  n: Integer;
begin
  tvDebug.Items.BeginUpdate;
  try
    tvDebug.Items.Clear;
    for n:=0 to pmnuLMB.Items.Count-1 do
    begin
      if pmnuLMB.Items[n] = mnuSepMenuList then Break;
      AddMenuItem(pmnuLMB.Items[n], nil);
    end;
  finally
    tvDebug.Items.EndUpdate;
  end;
end;

procedure TfrmMain.PopulateItems;
begin
  if Loading then Exit;
  AddLoading;
  FPopulateThread := TMenuPopulateThread.Create(Self);
end;

procedure TfrmMain.RemoveLoading;
begin
  if FLoadingCount > 0 then
  begin
    Dec(FLoadingCount);
    if FLoadingCount = 0 then
      LoadingChanged;
  end;
end;

procedure TfrmMain.tiMainClick(Sender: TObject);
var
  aPoint: TPoint;
begin
  if GetCursorPos(aPoint) then
  begin
    SetForeGroundWindow(Handle); //Zet window als actief
    pmnuLMB.Popup(aPoint.X, aPoint.Y);
//    PostMessage(Handle, $0000, 0, 0);
//    Application.ProcessMessages; //To let the OnClick event run first
  end;
end;

procedure TfrmMain.tmrWPChangeTimer(Sender: TObject);
begin
  If ReadIni(AppIniFile, 'Algemeen', 'Quit', RegBool) = True then
    Application.Terminate;

  if IsIconTrayReady then
    (Sender as TTimer).Interval := 1000 * 60
  else
    ChangeWallPaperIfWanted;
end;

procedure TfrmMain.tvDebugClick(Sender: TObject);
var
  iIcon: Word;
  aIcon: HIcon;
  aNode: TTreeNode;
  aFileName: string;
begin
  aNode := (Sender as TTreeView).Selected;
  if (aNode = nil) or (aNode.ImageIndex = 1) then
  begin
    Image1.Picture.Assign(nil);
    Exit;
  end;

  aFileName := FFullLinkNames[Integer(aNode.Data)];
  iIcon := edtIconIndex.AsInteger;

  CheckMSCFile(aFileName, iIcon);

  edtIconIndex.AsInteger := iIcon;

  aIcon := ExtractAssociatedIcon(MainInstance,PChar(aFileName), iIcon);
  if aIcon <> 0 then
    Image1.Picture.Icon.Handle := aIcon
  else
    Image1.Picture.Assign(nil);
end;

end.



