program RichPrint;

uses
  Vcl.Forms,
  Main in 'Main.pas' {MainForm},
  Preview in '..\..\Preview.pas';

{$R *.RES}

begin
  Application.Title := 'TPrintPreview Demo';
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
