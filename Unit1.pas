unit Unit1;

interface

uses
  SysUtils, Classes, Controls, Forms, Dialogs, ComCtrls, StdCtrls, ShellCtrls,
  INIfiles, FileCtrl;

type
  TForm1 = class(TForm)
    lblFolder: TLabel;
    lblWhat: TLabel;
    txtWhat: TEdit;
    lblWhere: TLabel;
    txtWhere: TEdit;
    cmdGo: TButton;
    txtFilter: TEdit;
    lblFilter: TLabel;
    cmdHelp: TButton;
    grpSize: TGroupBox;
    rbAll: TRadioButton;
    rbSmaller: TRadioButton;
    rbLarger: TRadioButton;
    txtSize: TEdit;
    lblMB: TLabel;
    grpFilesFolders: TGroupBox;
    rbFiles: TRadioButton;
    rbFolders: TRadioButton;
    rbBoth: TRadioButton;
    txtDepth: TEdit;
    lblDepth: TLabel;
    lblDepth2: TLabel;
    procedure cmdGoClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure cmdHelpClick(Sender: TObject);
    procedure GetAllSubFolders(sPath: string; FolderDepth: Integer);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  C: TShellTreeView;
  FilesNames: TStringList;

implementation

{$R *.dfm}

function GetSizeOfDir(dir: string; subdir: Boolean): Int64;
var
  rec: TSearchRec;
  found: Integer;
begin
  Result := 0;
  if dir[Length(dir)] <> '\' then
    dir := dir + '\';
  found := FindFirst(dir + '*.*', faAnyFile, rec);
  while found = 0 do
  begin
    Inc(Result, rec.Size);
    if (rec.Attr and faDirectory > 0) and (rec.Name[1] <> '.') and (subdir = True) then
      Inc(Result, GetSizeOfDir(dir + rec.Name, True));
    found := FindNext(rec);
  end;
  FindClose(rec);
end;

function GetSizeOfFile(const FileName: string): Int64;
var
  Rec: TSearchRec;
begin
  Result := 0;
  //Find what we assume is a file and get its size
  if (FindFirst(FileName, faAnyFile, Rec) = 0) then
    Result := Rec.Size;
  //If the file happens to have a folder attribute, get its size recursively
  if ((Rec.Attr and faDirectory) = faDirectory) then
    Result := GetSizeOfDir(FileName, True); //True ios for sub-folders
  //Close record
  FindClose(Rec);
end;

function ExtractFileNameWoExt(const FileName: string): string;
var
  i: integer;
begin
  i := LastDelimiter('.' + PathDelim + DriveDelim, FileName);
  if (i = 0) or (FileName[i] <> '.') then
    i := MaxInt;
  Result := ExtractFileName(Copy(FileName, 1, i - 1));
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  myINI: TINIFile;
  ThisSetting: string;
begin
  //Initialise options from INI file
  myINI := TINIFile.Create(ExtractFilePath(Application.EXEName) + 'DirList.ini');
  txtFilter.text := myINI.ReadString('Settings', 'Filter', '*.*');
  ThisSetting := myINI.ReadString('Settings', 'SizeChoice', 'All');
  if ThisSetting = 'All' then
    rbAll.Checked := true
  else if ThisSetting = 'Smaller' then
    rbSmaller.Checked := true
  else
    rbLarger.Checked := true;
  txtSize.text := IntToStr(myINI.ReadInteger('Settings', 'SizeValue', 0));
  ThisSetting := myINI.ReadString('Settings', 'FilesFolders', 'Files');
  if ThisSetting = 'Files' then
    rbFiles.Checked := true
  else if ThisSetting = 'Folders' then
    rbFolders.Checked := true
  else
    rbBoth.Checked := true;
  txtDepth.Text := IntToStr(myINI.ReadInteger('Settings', 'Depth', 0));
  txtWhat.text := myINI.ReadString('Settings', 'What', 'N,123,SM');
  txtWhere.text := myINI.ReadString('Settings', 'Where', 'output.csv');
  myINI.Free;
  C := TShellTreeView.Create(Self);
  with C do
  begin
    Left := 8;
    Top := lblFolder.Top + lblfolder.Height + 8;
    Width := cmdGo.Left - left - 8;
    Height := cmdGo.Top + cmdGo.Height - Top;
    Visible := True;
    Parent := Self;
    ObjectTypes := [otFolders];
    Root := 'rfMyComputer';
    UseShellImages := True;
    AutoRefresh := False;
    Indent := 23;
    ParentColor := False;
    RightClickSelect := True;
    ShowRoot := False;
    TabOrder := 0;
  end;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
var
  myINI: TIniFile;
  ThisSetting: string;
begin
  //Save settings to INI file
  myINI := TINIFile.Create(ExtractFilePath(Application.EXEName) + 'DirList.ini');
  myINI.WriteString('Settings', 'Filter', txtFilter.text);
  if rbAll.Checked then
    ThisSetting := 'All'
  else if rbSmaller.Checked then
    ThisSetting := 'Smaller'
  else
    ThisSetting := 'Larger';
  myINI.WriteString('Settings', 'SizeChoice', ThisSetting);
  myINI.WriteInteger('Settings', 'SizeValue', StrToInt(txtSize.text));
  if rbFiles.Checked then
    ThisSetting := 'Files'
  else if rbFolders.Checked then
    ThisSetting := 'Folders'
  else
    ThisSetting := 'Both';
  myINI.WriteString('Settings', 'FilesFolders', ThisSetting);
  myINI.WriteInteger('Settings', 'Depth', StrToInt(txtDepth.text));
  myINI.WriteString('Settings', 'What', txtWhat.text);
  myINI.WriteString('Settings', 'Where', txtWhere.text);
  myINI.Free;
end;

procedure TForm1.cmdHelpClick(Sender: TObject);
begin
  ShowMessage('D: drive' + sLineBreak + 'P: path' + sLineBreak + 'F: folder' + sLineBreak + 'N: name w/o extension' + sLineBreak + 'NE: name w/ extension' + sLineBreak + 'E: extension' + sLineBreak + 'SB: size in bytes' + sLineBreak + 'SK: size in KB' + sLineBreak + 'SM: sizein MB' + sLineBreak + 'A: attributes' + sLineBreak + 'Other: text as typed');
end;

procedure TForm1.cmdGoClick(Sender: TObject);
var
  Path, FullName: string;
  OutD, OutP, OutF, OutN, OutNE, OutE, outSB, outSK, outSM, outA, OutputString: string;
  List: TStrings;
  ThisFileSize, SizeLimit: Int64;
  i, j, FileAttributes, FilesListed: Integer;
  OutputFile: TextFile;
  AppendToFile: Boolean;
begin
  //Get Path for ShellTreeView
  Path := IncludeTrailingPathDelimiter(C.Path);
  //Create list of files names from selected path
  FilesNames := TStringList.Create;
  GetAllSubFolders(Path, 0);
  //Get list of properties to output
  List := TStringList.Create;
  ExtractStrings([','], [], PChar(txtWhat.Text), List);
  SizeLimit := StrToInt(txtSize.Text);
  //Try to open output file
  AssignFile(OutputFile, PChar(txtWhere.Text));
  //Check if file exists
  AppendToFile := False;
  if fileexists(PChar(txtWhere.Text)) then
    if messagedlg('The output file already exists.' + #13#10 + 'Do you want to append to it (no to overwrite)? ', mtCustom, [mbYes, mbNo], 0) = mrYes then
      AppendToFile := True;
  if AppendToFile then
    append(OutputFile)
  else
    ReWrite(OutputFile);
  FilesListed := 0;
  for i := 0 to FilesNames.Count - 1 do
  begin
    //Prepare properties values
    FullName := FilesNames[i];
    OutD := ExtractFileDrive(FullName);
    OutP := ExtractFilePath(FullName);
    OutF := ExtractFileDir(FullName);
    OutN := ExtractFileNameWoExt(FullName);
    OutNE := ExtractFileName(FullName);
    OutE := ExtractFileExt(FullName);
    ThisFileSize := GetSizeOfFile(FullName);
    outSB := IntToStr(ThisFileSize);
    outSK := IntToStr(ThisFileSize div 1024);
    outSM := IntToStr(ThisFileSize div 1048576);
    //If comma (,) in names/path strings, encompass in double quotes (")
    //This ensures the CSV file doesn't think these commas define extra columns
    if Pos(',', OutP) > 0 then
      OutP := '"' + OutP + '"';
    if Pos(',', OutF) > 0 then
      OutF := '"' + OutF + '"';
    if Pos(',', OutN) > 0 then
    begin
      OutN := '"' + OutN + '"';
      OutNE := '"' + OutNE + '"';
    end;
    //File attributes
    FileAttributes := filegetattr(FullName);
    outA := '';
    if FileAttributes and faReadOnly > 0 then
      outA := outA + 'R';
    if FileAttributes and faHidden > 0 then
      outA := outA + 'H';
    if FileAttributes and faSysFile > 0 then
      outA := outA + 'S';
    if FileAttributes and faArchive > 0 then
      outA := outA + 'A';
    OutputString := '';
    //Export file if size limit is met
    ThisFileSize := ThisFileSize div 1048576;
    if rbAll.Checked or (rbSmaller.Checked and (ThisFileSize < SizeLimit)) or (rbLarger.Checked and (ThisFileSize > SizeLimit)) then
    begin
      for j := 0 to List.Count - 1 do
      begin
        if List[j] = 'D' then
          OutputString := OutputString + OutD
        else if List[j] = 'P' then
          OutputString := OutputString + OutP
        else if List[j] = 'F' then
          OutputString := OutputString + OutF
        else if List[j] = 'N' then
          OutputString := OutputString + OutN
        else if List[j] = 'NE' then
          OutputString := OutputString + OutNE
        else if List[j] = 'E' then
          OutputString := OutputString + OutE
        else if List[j] = 'SB' then
          OutputString := OutputString + outSB
        else if List[j] = 'SK' then
          OutputString := OutputString + outSK
        else if List[j] = 'SM' then
          OutputString := OutputString + outSM
        else if List[j] = 'A' then
          OutputString := OutputString + outA
        else
          OutputString := OutputString + List[j];
        if j < (List.count - 1) then
          OutputString := OutputString + ',';
      end;
      WriteLn(OutputFile, OutputString);
      FilesListed := FilesListed + 1;
    end;
  end;
  //Close the file
  CloseFile(OutputFile);
  List.Free;
  if AppendToFile then
    showmessage('Added information for ' + IntToStr(FilesListed) + ' files to the file ' + PChar(txtWhere.Text))
  else
    showmessage('Wrote information for ' + IntToStr(FilesListed) + ' files to the file ' + PChar(txtWhere.Text));
end;

procedure TForm1.GetAllSubFolders(sPath: string; FolderDepth: Integer);
var
  Path: string;
  Rec: TSearchRec;
begin
  try
    Path := IncludeTrailingBackslash(sPath);
    if FindFirst(Path + PChar(txtFilter.Text), faDirectory, Rec) = 0 then
    try
      repeat
        if (Rec.Name <> '.') and (Rec.Name <> '..') then
        begin
          //Only add to list if matches requirements, i.e. files only, folders only or both
          if rbBoth.Checked or (rbFiles.Checked and not ((Rec.Attr and faDirectory) = faDirectory)) or (rbFolders.Checked and ((Rec.Attr and faDirectory) = faDirectory)) then
            FilesNames.add(Path + Rec.Name);
          if (FolderDepth < StrToInt(txtDepth.Text)) or (StrToInt(txtDepth.Text) = -1) then
            GetAllSubFolders(Path + Rec.Name, FolderDepth + 1);
        end;
      until FindNext(Rec) <> 0;
    finally
      FindClose(Rec);
    end;
  except
    on e: Exception do
      Showmessage('Error: ' + e.Message);
  end;
end;

end.

