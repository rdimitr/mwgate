program mwgate;

uses
  Vcl.Forms,
  Windows,
  mwgatemain in 'mwgatemain.pas' {Form1};

{$R *.res}
var
 MutexHandle : THandle;
const
 MutexName = '68ebf0a0-78ea-4890-901d-dc942a911b36';

begin
  MutexHandle := OpenMutex(MUTEX_ALL_ACCESS, false, MutexName);
  if MutexHandle <> 0 then begin
     CloseHandle(MutexHandle);
     halt;
  end;
  MutexHandle := CreateMutex(nil, false, MutexName);
  
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'ЕГИСЗ-документы';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
  
  CloseHandle(MutexHandle);
end.