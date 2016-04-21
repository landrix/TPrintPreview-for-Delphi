unit MainRichPrint;

interface

uses
  System.SysUtils, System.Classes,
  Winapi.Windows, Winapi.Messages,
  Vcl.Graphics, Vcl.Controls,Vcl.Forms, Vcl.ImgList,Vcl.Dialogs, Vcl.StdCtrls,
  Vcl.ExtCtrls, Vcl.Tabs, Vcl.ComCtrls, Vcl.Menus, Vcl.ToolWin,Vcl.ExtDlgs,
  Preview;

type
  TMainForm = class(TForm)
    Toolbar: TPanel;
    ZoomComboBox: TComboBox;
    PrintButton: TButton;
    Label1: TLabel;
    Image1: TImage;
    PrinterSetupDialog: TPrinterSetupDialog;
    OpenButton: TButton;
    OpenDialog: TOpenDialog;
    RichEdit: TRichEdit;
    Splitter1: TSplitter;
    Panel1: TPanel;
    PageNavigator: TTabSet;
    DirectPrint: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure ZoomComboBoxChange(Sender: TObject);
    procedure PrintButtonClick(Sender: TObject);
    procedure OpenButtonClick(Sender: TObject);
    procedure PageNavigatorChange(Sender: TObject; NewTab: Integer;
      var AllowChange: Boolean);
    procedure PrintPreviewChange(Sender: TObject);
    procedure PrintPreviewBeforePrint(Sender: TObject);
    procedure PrintPreviewProgress(Sender: TObject; Done, Total: Integer);
    procedure PrintPreviewAfterPrint(Sender: TObject);
    procedure PrintPreviewNewPage(Sender: TObject);
    procedure PrintPreviewBeginDoc(Sender: TObject);
    procedure PrintPreviewEndDoc(Sender: TObject);
    procedure PrintPreviewZoomChange(Sender: TObject);
  private
    PrintPreview: TPrintPreview;
    ThumbnailPreview: TThumbnailPreview;
  private
    procedure RenderRichEdit;
  end;

var
  MainForm: TMainForm;

implementation

{$R *.DFM}

procedure TMainForm.FormCreate(Sender: TObject);
begin
  PrintPreview:= TPrintPreview.Create(self);
  PrintPreview.Parent := self;
  with PrintPreview do
  begin
    Left := 0;
    Top := 0;
    Width := 557;
    Height := 529;
    HorzScrollBar.Margin := 10;
    HorzScrollBar.Tracking := True;
    VertScrollBar.Margin := 10;
    VertScrollBar.Tracking := True;
    ParentFont := True;
    TabOrder := 0;
    PaperView.BorderColor := clNavy;
    PaperView.ShadowWidth := 4;
    PrintJobTitle := 'TPrintPreview Sample Print';
    UsePrinterOptions := True;
    OnBeginDoc := PrintPreviewBeginDoc;
    OnEndDoc := PrintPreviewEndDoc;
    OnNewPage := PrintPreviewNewPage;
    OnChange := PrintPreviewChange;
    OnZoomChange := PrintPreviewZoomChange;
    OnProgress := PrintPreviewProgress;
    OnBeforePrint := PrintPreviewBeforePrint;
    OnAfterPrint := PrintPreviewAfterPrint;
  end;

  ThumbnailPreview := TThumbnailPreview.Create(self);
  ThumbnailPreview.Parent := self;
  with ThumbnailPreview do
  begin
    Left := 0;
    Top := 29;
    Width := 115;
    Height := 558;
    TabOrder := 2;
    PrintPreview := PrintPreview;
    PaperView.ShadowWidth := 1;
  end;

  Randomize;
  PrintPreview.ZoomState := zsZoomToFit;
  ZoomComboBox.ItemIndex := 6; // Zoom to Fit (Whole Page)
  if ParamCount = 1 then
  begin
    RichEdit.Lines.LoadFromFile(ParamStr(1));
    RenderRichEdit;
  end;
end;

procedure TMainForm.ZoomComboBoxChange(Sender: TObject);
begin
  case ZoomComboBox.ItemIndex of
    0: PrintPreview.Zoom := 50;
    1: PrintPreview.Zoom := 100;
    2: PrintPreview.Zoom := 150;
    3: PrintPreview.Zoom := 200;
    4: PrintPreview.ZoomState := zsZoomToWidth;
    5: PrintPreview.ZoomState := zsZoomToHeight;
    6: PrintPreview.ZoomState := zsZoomToFit;
  end;
end;

procedure TMainForm.PrintButtonClick(Sender: TObject);
begin
  if (PrintPreview.State = psReady) and PrinterSetupDialog.Execute then
  begin
    if not DirectPrint.Checked then
      PrintPreview.Print
    else
    begin
      PrintPreview.DirectPrint := True;
      try
        RenderRichEdit;
      finally
        PrintPreview.DirectPrint := False;
      end;
    end;
  end;
end;

procedure TMainForm.OpenButtonClick(Sender: TObject);
begin
  if OpenDialog.Execute then
  begin
    RichEdit.Lines.LoadFromFile(OpenDialog.FileName);
    RenderRichEdit;
  end;
end;

procedure TMainForm.PageNavigatorChange(Sender: TObject; NewTab: Integer;
  var AllowChange: Boolean);
begin
  PrintPreview.CurrentPage := NewTab + 1;
end;

procedure TMainForm.PrintPreviewChange(Sender: TObject);
begin
  while PageNavigator.Tabs.Count < PrintPreview.TotalPages do
    PageNavigator.Tabs.Add(IntToStr(PageNavigator.Tabs.Count + 1));
  while PageNavigator.Tabs.Count > PrintPreview.TotalPages do
    PageNavigator.Tabs.Delete(PageNavigator.Tabs.Count - 1);
  PageNavigator.TabIndex := PrintPreview.CurrentPage - 1;

  if PrintPreview.State = psReady then
    PrintButton.Enabled := PrintPreview.PrinterInstalled and (PrintPreview.TotalPages > 0)
  else
    Application.ProcessMessages;
end;

procedure TMainForm.PrintPreviewZoomChange(Sender: TObject);
begin
  case PrintPreview.ZoomState of
    zsZoomToFit: ZoomComboBox.ItemIndex := 6;
    zsZoomToHeight: ZoomComboBox.ItemIndex := 5;
    zsZoomToWidth: ZoomComboBox.ItemIndex := 4;
  else
    case PrintPreview.Zoom of
      200: ZoomComboBox.ItemIndex := 3;
      150: ZoomComboBox.ItemIndex := 2;
      100: ZoomComboBox.ItemIndex := 1;
      50: ZoomComboBox.ItemIndex := 0;
    else
      ZoomComboBox.ItemIndex := -1;
    end;
  end;
end;

procedure TMainForm.PrintPreviewBeginDoc(Sender: TObject);
begin
  Caption := Application.Title + ' - Creating pages...';

  PrintButton.Enabled := False;
  OpenButton.Enabled := False;
end;

procedure TMainForm.PrintPreviewEndDoc(Sender: TObject);
begin
  Caption := Application.Title;

  PrintButton.Enabled := PrintPreview.PrinterInstalled and (PrintPreview.TotalPages > 0);
  OpenButton.Enabled := True;
end;

procedure TMainForm.PrintPreviewBeforePrint(Sender: TObject);
begin
  Screen.Cursor := crHourglass;
  Caption := Application.Title + ' - Preparing to print...';

  PrintButton.Enabled := False;
  OpenButton.Enabled := False;
end;

procedure TMainForm.PrintPreviewAfterPrint(Sender: TObject);
begin
  Caption := Application.Title;
  Screen.Cursor := crDefault;

  PrintButton.Enabled := PrintPreview.PrinterInstalled and (PrintPreview.TotalPages > 0);
  OpenButton.Enabled := True;
end;

procedure TMainForm.PrintPreviewProgress(Sender: TObject;
  Done, Total: Integer);
begin
  Caption := Format('%s - Printing... (%.1f%% done)',
    [Application.Title, Done / Total * 100]);
  Update;
end;

procedure TMainForm.PrintPreviewNewPage(Sender: TObject);
var
  R: TRect;
begin
  with PrintPreview do
  begin
    // The following line ensures one pixel pen width in any mapping mode.
    Canvas.Pen.Width := 0;
    Canvas.Brush.Style := bsCLear;
    // Draws a frame with 1cm margin
    SetRect(R, 1000, 1000, PaperWidth - 1000, PaperHeight - 1000);
    Canvas.Rectangle(R.Left, R.Top, R.Right, R.Bottom);
    // Sets font's size to 8
    Canvas.Font.Size := 8;
    // Draws the page number under the frame
    Canvas.TextOut(R.Left, R.Bottom, Format('Page %d', [TotalPages+1]));
  end;
end;

procedure TMainForm.RenderRichEdit;
var
  ImageRect: array[Boolean] of TRect;
  TextRect: array[Boolean] of TRect;
  Toggled: Boolean;
  Offset: Integer;
  R: TRect;
begin
  with PrintPreview do
  begin
    Units := mmHiMetric; // All units are in 1/100th of millimeter
    BeginDoc;
    try
      SetRect(TextRect[False], 2000, 2000, PaperWidth div 2 - 500, PaperHeight div 2 + 3000);
      SetRect(ImageRect[False], PaperWidth div 2 + 500, 2000, PaperWidth - 2000, PaperHeight div 2 - 4000);
      SetRect(TextRect[True], PaperWidth div 2 + 500, PaperHeight div 2 - 3000, PaperWidth - 2000, PaperHeight - 2000);
      SetRect(ImageRect[True], 2000, PaperHeight div 2 + 4000, PaperWidth div 2 - 500, PaperHeight - 2000);
      Offset := 0;
      Toggled := False;
      while (Offset >= 0) and
            (PaintRichText(TextRect[Toggled], RichEdit, 1, @Offset) <> 0) do
      begin
        Application.ProcessMessages;
        R := TextRect[Toggled];
        InflateRect(R, 300, 300);
        Canvas.Rectangle(R.Left, R.Top, R.Right, R.Bottom);
        PaintGraphicEx(ImageRect[Toggled], Image1.Picture.Graphic, True, False, True);
        if Toggled and (Offset >= 0) then
          NewPage;
        Toggled := not Toggled;
      end;
    finally
      EndDoc;
    end;
  end;
end;

end.

