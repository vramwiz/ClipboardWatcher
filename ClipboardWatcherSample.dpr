program ClipboardWatcherSample;

uses
  Vcl.Forms,
  MainForm in 'MainForm.pas' {FormMain},
  ClipboardWatcher in 'ClipboardWatcher.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFormMain, FormMain);
  Application.Run;
end.
