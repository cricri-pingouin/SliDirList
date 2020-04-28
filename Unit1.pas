Unit Unit1;

Interface

Uses
  SysUtils, Classes, Controls, Forms, Dialogs, ComCtrls, StdCtrls, ShellCtrls,
  INIfiles;

Type
  TForm1 = Class(TForm)
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
    Procedure cmdGoClick(Sender: TObject);
    Procedure FormCreate(Sender: TObject);
    Procedure FormClose(Sender: TObject; Var Action: TCloseAction);
    Procedure cmdHelpClick(Sender: TObject);
    Procedure GetAllSubFolders(sPath: String; FolderDepth: Integer);
  Private
    { Private declarations }
  Public
    { Public declarations }
  End;

Var
  Form1: TForm1;
  C: TShellTreeView;
  FilesNames: TStringList;

Implementation

{$R *.dfm}

Function GetSizeOfDir(dir: String; subdir: Boolean): Int64;
Var
  rec: TSearchRec;
  found: Integer;
Begin
  Result := 0;
  If dir[Length(dir)] <> '\' Then
    dir := dir + '\';
  found := FindFirst(dir + '*.*', faAnyFile, rec);
  While found = 0 Do
  Begin
    Inc(Result, rec.Size);
    If (rec.Attr And faDirectory > 0) And (rec.Name[1] <> '.') And (subdir = True) Then
      Inc(Result, GetSizeOfDir(dir + rec.Name, True));
    found := FindNext(rec);
  End;
  FindClose(rec);
End;

Function GetSizeOfFile(Const FileName: String): Int64;
Var
  Rec: TSearchRec;
Begin
  Result := 0;
  //Find what we assume is a file and get its size
  If (FindFirst(FileName, faAnyFile, Rec) = 0) Then
    Result := Rec.Size;
  //If the file happens to have a folder attribute, get its size recursively
  If ((Rec.Attr And faDirectory) = faDirectory) Then
    Result := GetSizeOfDir(FileName, True); //True ios for sub-folders
  //Close record
  FindClose(Rec);
End;

Function ExtractFileNameWoExt(Const FileName: String): String;
Var
  i: integer;
Begin
  i := LastDelimiter('.' + PathDelim + DriveDelim, FileName);
  If (i = 0) Or (FileName[i] <> '.') Then
    i := MaxInt;
  Result := ExtractFileName(Copy(FileName, 1, i - 1));
End;

Procedure TForm1.FormCreate(Sender: TObject);
Var
  myINI: TINIFile;
  ThisSetting: String;
Begin
  //Initialise options from INI file
  myINI := TINIFile.Create(ExtractFilePath(Application.EXEName) + 'DirList.ini');
  txtFilter.text := myINI.ReadString('Settings', 'Filter', '*.*');
  ThisSetting := myINI.ReadString('Settings', 'SizeChoice', 'All');
  If ThisSetting = 'All' Then
    rbAll.Checked := true
  Else If ThisSetting = 'Smaller' Then
    rbSmaller.Checked := true
  Else
    rbLarger.Checked := true;
  txtSize.text := IntToStr(myINI.ReadInteger('Settings', 'SizeValue', 0));
  ThisSetting := myINI.ReadString('Settings', 'FilesFolders', 'Files');
  If ThisSetting = 'Files' Then
    rbFiles.Checked := true
  Else If ThisSetting = 'Folders' Then
    rbFolders.Checked := true
  Else
    rbBoth.Checked := true;
  txtDepth.Text := IntToStr(myINI.ReadInteger('Settings', 'Depth', 0));
  txtWhat.text := myINI.ReadString('Settings', 'What', 'N,123,SM');
  txtWhere.text := myINI.ReadString('Settings', 'Where', 'output.csv');
  myINI.Free;
  C := TShellTreeView.Create(Self);
  With C Do
  Begin
    Left := 8;
    Top := 32;
    Width := 329;
    Height := 369;
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
  End;
End;

Procedure TForm1.FormClose(Sender: TObject; Var Action: TCloseAction);
Var
  myINI: TIniFile;
  ThisSetting: String;
Begin
  //Save settings to INI file
  myINI := TINIFile.Create(ExtractFilePath(Application.EXEName) + 'DirList.ini');
  myINI.WriteString('Settings', 'Filter', txtFilter.text);
  If rbAll.Checked Then
    ThisSetting := 'All'
  Else If rbSmaller.Checked Then
    ThisSetting := 'Smaller'
  Else
    ThisSetting := 'Larger';
  myINI.WriteString('Settings', 'SizeChoice', ThisSetting);
  myINI.WriteInteger('Settings', 'SizeValue', StrToInt(txtSize.text));
  If rbFiles.Checked Then
    ThisSetting := 'Files'
  Else If rbFolders.Checked Then
    ThisSetting := 'Folders'
  Else
    ThisSetting := 'Both';
  myINI.WriteString('Settings', 'FilesFolders', ThisSetting);
  myINI.WriteInteger('Settings', 'Depth', StrToInt(txtDepth.text));
  myINI.WriteString('Settings', 'What', txtWhat.text);
  myINI.WriteString('Settings', 'Where', txtWhere.text);
  myINI.Free;
End;

Procedure TForm1.cmdHelpClick(Sender: TObject);
Begin
  ShowMessage('D: drive' + sLineBreak + 'P: path' + sLineBreak + 'F: folder' + sLineBreak + 'N: name w/o extension' + sLineBreak + 'NE: name w/ extension' + sLineBreak + 'E: extension' + sLineBreak + 'SB: size in bytes' + sLineBreak + 'SK: size in KB' + sLineBreak + 'SM: sizein MB' + sLineBreak + 'A: attributes' + sLineBreak + 'Other: text as typed');
End;

Procedure TForm1.cmdGoClick(Sender: TObject);
Var
  Path, FullName: String;
  OutD, OutP, OutF, OutN, OutNE, OutE, outSB, outSK, outSM, outA, OutputString: String;
  List: TStrings;
  ThisFileSize, SizeLimit: Int64;
  i, j, FileAttributes, FilesListed: Integer;
  OutputFile: TextFile;
  AppendToFile: Boolean;
Begin
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
  If fileexists(PChar(txtWhere.Text)) Then
    If messagedlg('The output file already exists.' + #13#10 + 'Do you want to append to it (no to overwrite)? ', mtCustom, [mbYes, mbNo], 0) = mrYes Then
      AppendToFile := True;
  If AppendToFile Then
    append(OutputFile)
  Else
    ReWrite(OutputFile);
  FilesListed := 0;
  For i := 0 To FilesNames.Count - 1 Do
  Begin
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
    outSK := IntToStr(ThisFileSize Div 1024);
    outSM := IntToStr(ThisFileSize Div 1048576);
    //If comma (,) in names/path strings, encompass in double quotes (")
    //This ensures the CSV file doesn't think these commas define extra columns
    If Pos(',', OutP) > 0 Then
      OutP := '"' + OutP + '"';
    If Pos(',', OutF) > 0 Then
      OutF := '"' + OutF + '"';
    If Pos(',', OutN) > 0 Then
    Begin
      OutN := '"' + OutN + '"';
      OutNE := '"' + OutNE + '"';
    End;
    //File attributes
    FileAttributes := filegetattr(FullName);
    outA := '';
    If FileAttributes And faReadOnly > 0 Then
      outA := outA + 'R';
    If FileAttributes And faHidden > 0 Then
      outA := outA + 'H';
    If FileAttributes And faSysFile > 0 Then
      outA := outA + 'S';
    If FileAttributes And faArchive > 0 Then
      outA := outA + 'A';
    OutputString := '';
    //Export file if size limit is met
    ThisFileSize := ThisFileSize Div 1048576;
    If rbAll.Checked Or (rbSmaller.Checked And (ThisFileSize < SizeLimit)) Or (rbLarger.Checked And (ThisFileSize > SizeLimit)) Then
    Begin
      For j := 0 To List.Count - 1 Do
      Begin
        If List[j] = 'D' Then
          OutputString := OutputString + OutD
        Else If List[j] = 'P' Then
          OutputString := OutputString + OutP
        Else If List[j] = 'F' Then
          OutputString := OutputString + OutF
        Else If List[j] = 'N' Then
          OutputString := OutputString + OutN
        Else If List[j] = 'NE' Then
          OutputString := OutputString + OutNE
        Else If List[j] = 'E' Then
          OutputString := OutputString + OutE
        Else If List[j] = 'SB' Then
          OutputString := OutputString + outSB
        Else If List[j] = 'SK' Then
          OutputString := OutputString + outSK
        Else If List[j] = 'SM' Then
          OutputString := OutputString + outSM
        Else If List[j] = 'A' Then
          OutputString := OutputString + outA
        Else
          OutputString := OutputString + List[j];
        If j < (List.count - 1) Then
          OutputString := OutputString + ',';
      End;
      WriteLn(OutputFile, OutputString);
      FilesListed := FilesListed + 1;
    End;
  End;
  //Close the file
  CloseFile(OutputFile);
  List.Free;
  If AppendToFile Then
    showmessage('Added information for ' + IntToStr(FilesListed) + ' files to the file ' + PChar(txtWhere.Text))
  Else
    showmessage('Wrote information for ' + IntToStr(FilesListed) + ' files to the file ' + PChar(txtWhere.Text));
End;

Procedure TForm1.GetAllSubFolders(sPath: String; FolderDepth: Integer);
Var
  Path: String;
  Rec: TSearchRec;
Begin
  Try
    Path := IncludeTrailingBackslash(sPath);
    If FindFirst(Path + PChar(txtFilter.Text), faDirectory, Rec) = 0 Then
    Try
      Repeat
        If (Rec.Name <> '.') And (Rec.Name <> '..') Then
        Begin
          //Only add to list if matches requirements, i.e. files only, folders only or both
          If rbBoth.Checked Or (rbFiles.Checked And Not ((Rec.Attr And faDirectory) = faDirectory)) Or (rbFolders.Checked And ((Rec.Attr And faDirectory) = faDirectory)) Then
            FilesNames.add(Path + Rec.Name);
          If (FolderDepth < StrToInt(txtDepth.Text)) Or (StrToInt(txtDepth.Text) = -1) Then
            GetAllSubFolders(Path + Rec.Name, FolderDepth + 1);
        End;
      Until FindNext(Rec) <> 0;
    Finally
      FindClose(Rec);
    End;
  Except
    On e: Exception Do
      Showmessage('Error: ' + e.Message);
  End;
End;

End.

