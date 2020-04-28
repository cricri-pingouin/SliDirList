program DirList;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  ShellConsts in 'ShellConsts.pas',
  ShellCtrls in 'ShellCtrls.pas';

{$R *.res}
{$SetPEFlags 1}

begin
  Application.Initialize;
  Application.Title := 'DirList';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
