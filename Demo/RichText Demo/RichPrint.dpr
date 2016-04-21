program RichPrint;

uses
  Vcl.Forms,
  MainRichPrint in 'MainRichPrint.pas' {MainForm},
  Preview in '..\..\Preview.pas';

{$R *.RES}

begin
  Application.Title := 'TPrintPreview Demo';
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
