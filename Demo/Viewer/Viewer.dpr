program Viewer;

uses
  Vcl.Forms,
  Main in 'Main.pas' {MainForm},
  Preview in '..\..\Preview.pas';

{$R *.RES}

begin
  Application.Title := 'Print Preview Viewer';
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
