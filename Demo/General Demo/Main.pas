unit Main;

{$I DELPHIAREA.INC}

interface

uses
  System.SysUtils, System.Classes,
  Winapi.Windows, Winapi.Messages,
  Vcl.Graphics, Vcl.Controls,Vcl.Forms, Vcl.ImgList,Vcl.Dialogs, Vcl.StdCtrls,
  Vcl.ExtCtrls, Vcl.Tabs, Vcl.ComCtrls, Vcl.Menus, Vcl.ToolWin,Vcl.ExtDlgs,
  Preview;

type
  TMainForm = class(TForm)
    Image1: TImage;
    Image2: TImage;
    PrinterSetupDialog: TPrinterSetupDialog;
    OpenDialog: TOpenDialog;
    SaveDialog: TSaveDialog;
    RichEdit1: TRichEdit;
    Splitter: TSplitter;
    SavePDFDialog: TSaveDialog;
    GrayscaleOptionsPanel: TPanel;
    Bevel3: TBevel;
    Label4: TLabel;
    Contrast: TTrackBar;
    Brightness: TTrackBar;
    Label5: TLabel;
    btnGrayReset: TButton;
    ThumbViewerPopupMenu: TPopupMenu;
    PagePopupMenu: TPopupMenu;
    AddPageBeforeCurrentMenuItem: TMenuItem;
    AddPageAfterCurrentMenuItem: TMenuItem;
    N1: TMenuItem;
    ReplaceCurrentMenuItem: TMenuItem;
    EditCurrentMenuItem: TMenuItem;
    N2: TMenuItem;
    DeleteCurrent: TMenuItem;
    ReduceThumbnailSizeMenuItem: TMenuItem;
    EnlargeThumbnailSizeMenuItem: TMenuItem;
    N3: TMenuItem;
    MoveCurrentUpMenuItem: TMenuItem;
    MoveCurrentDownMenuItem: TMenuItem;
    N4: TMenuItem;
    AddPageMenuItem: TMenuItem;
    MoveCurrentFirstMenuItem: TMenuItem;
    MoveCurrentLastMenuItem: TMenuItem;
    DeleteAllMenuItem: TMenuItem;
    N5: TMenuItem;
    HotTrackMenuItem: TMenuItem;
    MultiSelectMenuItem: TMenuItem;
    N6: TMenuItem;
    ArrangeLeftMenuItem: TMenuItem;
    ArrangeTopMenuItem: TMenuItem;
    ToolBar: TToolBar;
    btnOpen: TToolButton;
    btnSave: TToolButton;
    btnSavePDF: TToolButton;
    ToolButton4: TToolButton;
    btnPrint: TToolButton;
    btnDirectPrint: TToolButton;
    ToolButton7: TToolButton;
    btnZoomOut: TToolButton;
    btnZoom: TToolButton;
    btnZoomIn: TToolButton;
    ToolButton11: TToolButton;
    btnPageSetup: TToolButton;
    btnUnits: TToolButton;
    btnPrintableArea: TToolButton;
    ImageList: TImageList;
    ZoomPopupMenu: TPopupMenu;
    UnitsPopupMenu: TPopupMenu;
    ToolButton1: TToolButton;
    btnGrayscale: TToolButton;
    ToolButton2: TToolButton;
    btnFirstPage: TToolButton;
    btnPriorPage: TToolButton;
    btnNextPage: TToolButton;
    btnLastPage: TToolButton;
    Pixels1: TMenuItem;
    N01mm1: TMenuItem;
    N001mm1: TMenuItem;
    N001inch1: TMenuItem;
    N0001inch1: TMenuItem;
    Twips1: TMenuItem;
    Points1: TMenuItem;
    ZoomActualSize: TMenuItem;
    ZoomPageWidth: TMenuItem;
    ZoomPageHeight: TMenuItem;
    ZoomWholePage: TMenuItem;
    PageSetupDialog: TPageSetupDialog;
    StatusBar: TStatusBar;
    Bevel1: TBevel;
    Grayscale1: TMenuItem;
    ThumbGrayPreview: TMenuItem;
    ThumbGrayNever: TMenuItem;
    ThumbGrayAlways: TMenuItem;
    btnRandomPages: TToolButton;
    ToolButton3: TToolButton;
    PreviewPopupMenu: TPopupMenu;
    AddRandomPages1: TMenuItem;
    Clear1: TMenuItem;
    btnPrinterSetup: TToolButton;
    btnSaveTIF: TToolButton;
    SaveTIFDialog: TSaveDialog;
    ExportCurrent: TMenuItem;
    N7: TMenuItem;
    CopyCurrent: TMenuItem;
    SavePictureDialog: TSavePictureDialog;
    procedure FormCreate(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnOpenClick(Sender: TObject);
    procedure PrintPreviewChange(Sender: TObject);
    procedure PrintPreviewProgress(Sender: TObject; Done, Total: Integer);
    procedure PrintPreviewNewPage(Sender: TObject);
    procedure PrintPreviewZoomChange(Sender: TObject);
    procedure btnGrayscaleClick(Sender: TObject);
    procedure btnSavePDFClick(Sender: TObject);
    procedure ContrastChange(Sender: TObject);
    procedure BrightnessChange(Sender: TObject);
    procedure btnGrayResetClick(Sender: TObject);
    procedure AddPageBeforeCurrentMenuItemClick(Sender: TObject);
    procedure AddPageAfterCurrentMenuItemClick(Sender: TObject);
    procedure ReplaceCurrentMenuItemClick(Sender: TObject);
    procedure EditCurrentMenuItemClick(Sender: TObject);
    procedure DeleteCurrentClick(Sender: TObject);
    procedure ReduceThumbnailSizeMenuItemClick(Sender: TObject);
    procedure EnlargeThumbnailSizeMenuItemClick(Sender: TObject);
    procedure MoveCurrentUpMenuItemClick(Sender: TObject);
    procedure MoveCurrentDownMenuItemClick(Sender: TObject);
    procedure AddPageMenuItemClick(Sender: TObject);
    procedure MoveCurrentFirstMenuItemClick(Sender: TObject);
    procedure MoveCurrentLastMenuItemClick(Sender: TObject);
    procedure PagePopupMenuPopup(Sender: TObject);
    procedure PrintPreviewStateChange(Sender: TObject);
    procedure DeleteAllMenuItemClick(Sender: TObject);
    procedure ThumbViewerPopupMenuPopup(Sender: TObject);
    procedure HotTrackMenuItemClick(Sender: TObject);
    procedure MultiSelectMenuItemClick(Sender: TObject);
    procedure ArrangeLeftMenuItemClick(Sender: TObject);
    procedure ArrangeTopMenuItemClick(Sender: TObject);
    procedure btnPrintableAreaClick(Sender: TObject);
    procedure btnDirectPrintClick(Sender: TObject);
    procedure btnPageSetupClick(Sender: TObject);
    procedure btnZoomOutClick(Sender: TObject);
    procedure btnZoomInClick(Sender: TObject);
    procedure btnZoomClick(Sender: TObject);
    procedure btnUnitsClick(Sender: TObject);
    procedure btnFirstPageClick(Sender: TObject);
    procedure btnPriorPageClick(Sender: TObject);
    procedure btnNextPageClick(Sender: TObject);
    procedure btnLastPageClick(Sender: TObject);
    procedure ZoomActualSizeClick(Sender: TObject);
    procedure ZoomPageWidthClick(Sender: TObject);
    procedure ZoomPageHeightClick(Sender: TObject);
    procedure ZoomWholePageClick(Sender: TObject);
    procedure ZoomPopupMenuPopup(Sender: TObject);
    procedure UnitsClick(Sender: TObject);
    procedure UnitsPopupMenuPopup(Sender: TObject);
    procedure ThumbGrayNeverClick(Sender: TObject);
    procedure ThumbGrayAlwaysClick(Sender: TObject);
    procedure ThumbGrayPreviewClick(Sender: TObject);
    procedure btnRandomPagesClick(Sender: TObject);
    procedure btnPrinterSetupClick(Sender: TObject);
    procedure btnSaveTIFClick(Sender: TObject);
    procedure ExportCurrentClick(Sender: TObject);
    procedure CopyCurrentClick(Sender: TObject);
  private
    PrintPreview: TPrintPreview;
    ThumbnailPreview: TThumbnailPreview;
  private
    PageBoundsAfterMargin: TRect;
    procedure DrawImageTextPage;
    procedure DrawImageOnlyPage;
    procedure DrawRichTextPage;
    procedure DrawRandomCircles;
    procedure GeneratePages;
  end;

var
  MainForm: TMainForm;

implementation

{$R *.DFM}

uses
  System.Math, Vcl.Printers, Vcl.Clipbrd;

procedure TMainForm.FormCreate(Sender: TObject);
var
  SampleRTF: String;
begin
  PrintPreview:= TPrintPreview.Create(self);
  PrintPreview.Parent := self;
  with PrintPreview do
  begin
    Left := 164;
    Top := 26;
    Width := 620;
    Height := 470;
    HorzScrollBar.Margin := 10;
    HorzScrollBar.Smooth := True;
    HorzScrollBar.Tracking := True;
    VertScrollBar.Margin := 10;
    VertScrollBar.Smooth := True;
    VertScrollBar.Tracking := True;
    BorderStyle := bsNone           ;
    Font.Charset := DEFAULT_CHARSET;
    Font.Color := clWindowText;
    Font.Height := -11;
    Font.Name := 'Tahoma';
    Font.Style := [];
    PopupMenu := PreviewPopupMenu;
    TabOrder := 4;
    PaperView.PopupMenu := PagePopupMenu;
    PDFDocumentInfo.Producer := 'TPrintPreview Component';
    PDFDocumentInfo.Creator := 'TPrintPreview General Demo';
    PDFDocumentInfo.Author := 'delphiarea.com';
    PDFDocumentInfo.Subject := 'Auto generated pages';
    PDFDocumentInfo.Title := 'General Demo';
    PrintableAreaColor := clRed;
    PrintJobTitle := 'TPrintPreview Sample Print';
    OnNewPage := PrintPreviewNewPage;
    OnChange := PrintPreviewChange;
    OnStateChange := PrintPreviewStateChange;
    OnZoomChange := PrintPreviewZoomChange;
    OnProgress := PrintPreviewProgress;
  end;

  ThumbnailPreview := TThumbnailPreview.Create(self);
  ThumbnailPreview.Parent := self;
  with ThumbnailPreview do
  begin
    Left := 0;
    Top := 26;
    Width := 161;
    Height := 470;
    BorderStyle := bsNone;
    PopupMenu := ThumbViewerPopupMenu;
    TabOrder := 1;
    AllowReorder := True;
    MarkerColor := clRed;
    PrintPreview := PrintPreview;
    PaperView.PopupMenu := PagePopupMenu;
    PaperView.ShadowWidth := 1;
    Zoom := 15                 ;
  end;


  PrintPreview.Zoom := 100;
  PrintPreview.Grayscale := [];
  {$IFDEF COMPILER7_UP}
  PrintPreview.SetPageSetupParameters(PageSetupDialog);
  {$ENDIF}
  btnPrintableArea.Down := PrintPreview.ShowPrintableArea;
  btnUnits.Caption := UnitsPopupMenu.Items[Ord(PrintPreview.Units)].Caption;
  Brightness.Position := PrintPreview.GrayBrightness;
  Contrast.Position := PrintPreview.GrayContrast;
  if (ParamCount > 0) and FileExists(ParamStr(1)) then
    RichEdit1.Lines.LoadFromFile(ParamStr(1))
  else
  begin
    SampleRTF := ExtractFilePath(Application.ExeName) + 'TEAMWORK.rtf';
    if FileExists(SampleRTF) then
      RichEdit1.Lines.LoadFromFile(SampleRTF);
  end;
end;

procedure TMainForm.FormActivate(Sender: TObject);
begin
  Update;
  GeneratePages;
end;

procedure TMainForm.btnRandomPagesClick(Sender: TObject);
var
  S: String;
  I: Integer;
begin
  S := '100';
  if InputQuery('Number of Pages', 'Enter number of pages to add:', S) then
    for I := 1 to StrToInt(S) do
      if PrintPreview.BeginInsert(MaxInt) then
        try
          DrawRandomCircles;
        finally
          PrintPreview.EndInsert(False);
        end;
end;

procedure TMainForm.btnPrintableAreaClick(Sender: TObject);
begin
  PrintPreview.ShowPrintableArea := not PrintPreview.ShowPrintableArea;
end;

procedure TMainForm.btnPageSetupClick(Sender: TObject);
begin
  {$IFDEF COMPILER7_UP}
  if PageSetupDialog.Execute then
    GeneratePages;
  {$ENDIF}
end;

procedure TMainForm.btnPrinterSetupClick(Sender: TObject);
begin
  if PrinterSetupDialog.Execute then
  begin
    PrintPreview.GetPrinterOptions;
    {$IFDEF COMPILER7_UP}
    PrintPreview.SetPageSetupParameters(PageSetupDialog);
    {$ENDIF}
    GeneratePages;
  end;
end;

procedure TMainForm.btnPrintClick(Sender: TObject);
begin
  PrintPreview.SetPrinterOptions;
  if PrinterSetupDialog.Execute then
    PrintPreview.Print;
end;

procedure TMainForm.btnDirectPrintClick(Sender: TObject);
begin
  if PrinterSetupDialog.Execute then
  begin
    PrintPreview.DirectPrint := True;
    try
      GeneratePages;
    finally
      PrintPreview.DirectPrint := False;
    end;
  end;
end;

procedure TMainForm.btnFirstPageClick(Sender: TObject);
begin
  PrintPreview.CurrentPage := 1;
end;

procedure TMainForm.btnPriorPageClick(Sender: TObject);
begin
  with PrintPreview do CurrentPage := CurrentPage - 1;
end;

procedure TMainForm.btnNextPageClick(Sender: TObject);
begin
  with PrintPreview do CurrentPage := CurrentPage + 1;
end;

procedure TMainForm.btnLastPageClick(Sender: TObject);
begin
  with PrintPreview do CurrentPage := TotalPages;
end;

procedure TMainForm.btnSaveClick(Sender: TObject);
begin
  if SaveDialog.Execute then
    PrintPreview.SaveToFile(SaveDialog.FileName);
end;

procedure TMainForm.btnSavePDFClick(Sender: TObject);
begin
  if SavePDFDialog.Execute then
    PrintPreview.SaveAsPDF(SavePDFDialog.FileName);
end;

procedure TMainForm.btnSaveTIFClick(Sender: TObject);
begin
  if SaveTIFDialog.Execute then
    PrintPreview.SaveAsTIF(SaveTIFDialog.FileName);
end;

procedure TMainForm.btnUnitsClick(Sender: TObject);
begin
  PrintPreview.Units := TUnits((Ord(PrintPreview.Units) + 1) mod (Ord(High(TUnits)) + 1));
  GeneratePages;
end;

procedure TMainForm.btnZoomClick(Sender: TObject);
begin
  case PrintPreview.ZoomState of
    zsZoomOther:
      if PrintPreview.Zoom = 100 then
        PrintPreview.ZoomState := zsZoomToWidth
      else
        PrintPreview.Zoom := 100;
    zsZoomToWidth:
      PrintPreview.ZoomState := zsZoomToHeight;
    zsZoomToHeight:
      PrintPreview.ZoomState := zsZoomToFit;
    zsZoomToFit:
      PrintPreview.Zoom := 100;
  end;
end;

procedure TMainForm.btnZoomInClick(Sender: TObject);
begin
  with PrintPreview do Zoom := Zoom + ZoomStep;
end;

procedure TMainForm.btnZoomOutClick(Sender: TObject);
begin
  with PrintPreview do Zoom := Zoom - ZoomStep;
end;

procedure TMainForm.btnOpenClick(Sender: TObject);
begin
  if OpenDialog.Execute then
    PrintPreview.LoadFromFile(OpenDialog.FileName);
end;

procedure TMainForm.btnGrayscaleClick(Sender: TObject);
begin
  if PrintPreview.Grayscale = [] then
  begin
    btnGrayscale.ImageIndex := 22;
    GrayscaleOptionsPanel.Visible := True;
    PrintPreview.Grayscale := [gsPreview, gsPrint]
  end
  else
  begin
    btnGrayscale.ImageIndex := 23;
    GrayscaleOptionsPanel.Visible := False;
    PrintPreview.Grayscale := [];
  end;
end;

procedure TMainForm.ContrastChange(Sender: TObject);
begin
  PrintPreview.GrayContrast := Contrast.Position;
end;

procedure TMainForm.BrightnessChange(Sender: TObject);
begin
  PrintPreview.GrayBrightness := Brightness.Position;
end;

procedure TMainForm.btnGrayResetClick(Sender: TObject);
begin
  Brightness.Position := 0;
  Contrast.Position := 0;
end;

// PrintPreview event handlers start here

procedure TMainForm.PrintPreviewStateChange(Sender: TObject);
begin
  if PrintPreview.State = psReady then
  begin
    Screen.Cursor := crDefault;
    btnOpen.Enabled := True;
    btnSave.Enabled := (PrintPreview.TotalPages > 0);
    btnSavePDF.Enabled := PrintPreview.CanSaveAsPDF and (PrintPreview.TotalPages > 0);
    btnSaveTIF.Enabled := PrintPreview.CanSaveAsTIF and (PrintPreview.TotalPages > 0);
    btnPrint.Enabled := PrintPreview.PrinterInstalled and (PrintPreview.TotalPages > 0);
    btnDirectPrint.Enabled := PrintPreview.PrinterInstalled and (PrintPreview.TotalPages > 0);
    btnPageSetup.Enabled := {$IFDEF COMPILER7_UP} True {$ELSE} False {$ENDIF};
    btnUnits.Enabled := True;
    StatusBar.Panels[0].Text := Format('Page %d of %d',
      [PrintPreview.CurrentPage, PrintPreview.TotalPages]);
    StatusBar.Panels[5].Text := '';
    StatusBar.Panels[6].Text := '';
  end
  else if not (PrintPreview.State in [psInserting, psEditing, psReplacing]) then
  begin
    if PrintPreview.State = psCreating then
      Screen.Cursor := crAppStart
    else
      Screen.Cursor := crHourGlass;
    btnOpen.Enabled := False;
    btnSave.Enabled := False;
    btnSavePDF.Enabled := False;
    btnSaveTIF.Enabled := False;
    btnPrint.Enabled := False;
    btnDirectPrint.Enabled := False;
    btnPageSetup.Enabled := False;
    btnUnits.Enabled := False;
    case PrintPreview.State of
      psCreating:
      begin
        StatusBar.Panels[2].Text := PrintPreview.FormName;
        if PrintPreview.IsPaperRotated then
          StatusBar.Panels[3].Text := 'Landscape'
        else
          StatusBar.Panels[3].Text := 'Portrait';
        case PrintPreview.Units of
          mmPixel:
            StatusBar.Panels[4].Text := 'Pixels';
          mmLoMetric:
            StatusBar.Panels[4].Text := '1/10 mm';
          mmHiMetric:
            StatusBar.Panels[4].Text := '1/100 mm';
          mmLoEnglish:
            StatusBar.Panels[4].Text := '1/100"';
          mmHiEnglish:
            StatusBar.Panels[4].Text := '1/1000"';
          mmTWIPS:
            StatusBar.Panels[4].Text := 'TWIPS (1/1440")';
          mmPoints:
            StatusBar.Panels[4].Text := 'Points (1/72")';
        end;
        StatusBar.Panels[6].Text := 'Creating pages...';
      end;
      psLoading:
        StatusBar.Panels[6].Text := 'Loading pages from file...';
      psSaving:
        StatusBar.Panels[6].Text := 'Saving pages to file...';
      psSavingPDF:
        StatusBar.Panels[6].Text := 'Saving pages as PDF...';
      psSavingTIF:
        StatusBar.Panels[6].Text := 'Saving pages as multi-frame TIFF...';
      psPrinting:
        StatusBar.Panels[6].Text := 'Printing pages...';
    end;
  end;
  Update;
end;

procedure TMainForm.PrintPreviewChange(Sender: TObject);
begin
  StatusBar.Panels[0].Text := Format('Page %d of %d',
    [PrintPreview.CurrentPage, PrintPreview.TotalPages]);
  btnFirstPage.Enabled := (PrintPreview.CurrentPage > 1);
  btnPriorPage.Enabled := (PrintPreview.CurrentPage > 1);
  btnNextPage.Enabled := (PrintPreview.CurrentPage < PrintPreview.TotalPages);
  btnLastPage.Enabled := (PrintPreview.CurrentPage < PrintPreview.TotalPages);
  if PrintPreview.State in [psCreating, psInserting] then
  begin
    // allow user to navigate generated pages while other pages are stil in progress to generate
    Application.ProcessMessages;
  end;
end;

procedure TMainForm.PrintPreviewZoomChange(Sender: TObject);
begin
  StatusBar.Panels[1].Text := Format('%%%d', [PrintPreview.Zoom]);
  btnZoomOut.Enabled := (PrintPreview.Zoom > PrintPreview.ZoomMin);
  btnZoomIn.Enabled := (PrintPreview.Zoom < PrintPreview.ZoomMax);
  case PrintPreview.ZoomState of
    zsZoomToWidth:
    begin
      ZoomPageWidth.Checked := True;
      btnZoom.ImageIndex := ZoomPageHeight.ImageIndex;
      btnZoom.Hint := 'Zoom to page height';
    end;
    zsZoomToHeight:
    begin
      ZoomPageHeight.Checked := True;
      btnZoom.ImageIndex := ZoomWholePage.ImageIndex;
      btnZoom.Hint := 'Zoom to whole page';
    end;
    zsZoomToFit:
    begin
      ZoomWholePage.Checked := True;
      btnZoom.ImageIndex := ZoomActualSize.ImageIndex;
      btnZoom.Hint := 'Zoom to actual size';
    end;
  else
    if PrintPreview.Zoom = 100 then
    begin
      ZoomActualSize.Checked := True;
      btnZoom.ImageIndex := ZoomPageWidth.ImageIndex;
      btnZoom.Hint := 'Zoom to page width';
    end
    else
    begin
      ZoomActualSize.Checked := False;
      ZoomPageWidth.Checked := False;
      ZoomPageHeight.Checked := False;
      ZoomWholePage.Checked := False;
      btnZoom.ImageIndex := ZoomActualSize.ImageIndex;
      btnZoom.Hint := 'Zoom to actual size';
    end;
  end;
end;

procedure TMainForm.PrintPreviewProgress(Sender: TObject;
  Done, Total: Integer);
begin
  StatusBar.Panels[5].Text := FormatFloat('#,##0.0%', Done / Total * 100);
  StatusBar.Update;
end;

procedure TMainForm.PrintPreviewNewPage(Sender: TObject);
begin
  with PrintPreview do
  begin
    Canvas.Pen.Width := XFrom(mmLoMetric, 5); { 0.05 mm }
    Canvas.Pen.Color := clBlack;
    Canvas.Brush.Style := bsClear;
    // We are going to draw a rectangle on the page to distinguish the page's
    // margin. The margin is already calculated in the GeneratePages function.
    with PageBoundsAfterMargin do
      Canvas.Rectangle(Left, Top, Right, Bottom);
  end;
end;

// functions creating pages start here

procedure TMainForm.GeneratePages;
begin
  with PrintPreview do
  begin
    {$IFDEF COMPILER7_UP}
    PageBoundsAfterMargin := GetPageSetupParameters(PageSetupDialog);
    {$ELSE}
    PageBoundsAfterMargin := PageBounds;
    with PointFrom(mmLoMetric, 100, 100) do
      InflateRect(PageBoundsAfterMargin, -X, -Y);
    {$ENDIF}
    BeginDoc;
    try
      DrawImageTextPage;
      NewPage;
      DrawImageOnlyPage;
      NewPage;
      DrawRichTextPage;
    finally
      EndDoc;
    end;
  end;
end;

// In this example, the code is independent of the Units property of
// PrintPreview. If you use only one measuremnt unit for PrintPreview, you can
// easily use constant values instead of passing them to conversion methods.
// I also tried to write the code independent of the paper size.

procedure TMainForm.DrawImageTextPage;
var
  R: TRect;
  OneCM: TPoint;
  SavedBottom: Integer;
  Text: String;
begin
  with PrintPreview do
  begin
    R := PageBoundsAfterMargin;
    // Let's know how many units reperesents 1cm
    OneCM := PointFrom(mmLoMetric, 100, 100);
    // 1cm margin to look better
    InflateRect(R, -OneCM.X, -OneCM.Y);
    // We want to place an image horizontally in the top center of the paper.
    // In addition, we want the image height does not exceed 3 cm limit.
    SavedBottom := R.Bottom;
    R.Bottom := R.Top + 3 * OneCM.Y;
    PaintGraphicEx(R, Image1.Picture.Graphic, True, True, True);
    // We are going to draw a frame and write some text inside it. The new
    // frame is 1cm under the image boundary.
    R.Top := R.Bottom + OneCM.Y;
    R.Bottom := SavedBottom;
    // draw the frame
    Canvas.Pen.Width := XFrom(mmLoMetric, 5); { 0.05 mm }
    Canvas.Rectangle(R.Left, R.Top, R.Right, R.Bottom);
    // write the frame's dimensions under ir
    Text := Format('%d x %d (%s)', [R.Right - R.Left, R.Bottom - R.Top, btnUnits.Caption]);
    Canvas.Font.Size := 8;
    Canvas.TextOut(R.Left, R.Bottom + Canvas.Pen.Width, Text);
    // start with 12pt font size
    Canvas.Font.Size := 12;
    // while we have not reached to the frame's bottom (2 mm space), do...
    InflateRect(R, -OneCM.X div 5, -OneCM.Y div 5);
    while R.Top - Canvas.Font.Height <= R.Bottom do
    begin
      // randomly we select the font color
      Canvas.Font.Color := RGB(Random(256), Random(256), Random(256));
      // draw the text
      Canvas.TextRect(R, R.Left, R.Top, 'Powered by Borland Delphi.');
      // move the frame's top to the next line,
      Inc(R.Top, -Canvas.Font.Height);
      // and increase the font size by 1pt
      Canvas.Font.Size := Canvas.Font.Size + 1;
    end;
  end;
end;

procedure TMainForm.DrawImageOnlyPage;
var
  PR: TRect;
begin
  with PrintPreview do
  begin
    PR := PageBoundsAfterMargin;
    with PointFrom(mmLoMetric, 50, 50) do { 0.5 cm additional margin }
      InflateRect(PR, -X, -Y);
    PaintGraphicEx(PR, Image2.Picture.Graphic, True, False, True);
  end;
end;

procedure TMainForm.DrawRichTextPage;
var
  PR: TRect;
begin
  with PrintPreview do
  begin
    PR := PageBoundsAfterMargin;
    with PointFrom(mmLoMetric, 100, 100) do { 1 cm additional margin }
      InflateRect(PR, -X, -Y);
    PaintRichText(PR, RichEdit1, 0, nil);
  end;
end;

procedure TMainForm.DrawRandomCircles;
var
  Center: TPoint;
  Radius: Integer;
  MaxRadius: Integer;
  MaxSize: TSize;
  PR, R: TRect;
  I: Integer;
begin
  Randomize;
  with PrintPreview do
  begin
    PR := PageBoundsAfterMargin;
    with PointFrom(mmLoMetric, 50, 50) do { 0.5 cm additional margin }
      InflateRect(PR, -X, -Y);
    MaxSize.cx := PR.Right - PR.Left;
    MaxSize.cy := PR.Bottom - PR.Top;
    if MaxSize.cx < MaxSize.cy then
      MaxRadius := MaxSize.cx div 10
    else
      MaxRadius := MaxSize.cy div 10;
    Canvas.Pen.Mode := pmMask;
    Canvas.Pen.Width := XFrom(mmLoMetric, 1); { 0.1 mm }
    for I := 1 to 20 do
    begin
      Canvas.Pen.Color := RGB(Random(256), Random(256), Random(256));
      Canvas.Brush.Color := RGB(Random(256), Random(256), Random(256));
      Radius := Random(MaxRadius);
      Center.X := PR.Left + Radius + Random(MaxSize.cx - 2 * Radius);
      Center.Y := PR.Top + Radius + Random(MaxSize.cy - 2 * Radius);
      R.TopLeft := Center; R.BottomRight := Center;
      InflateRect(R, Radius, Radius);
      Canvas.Ellipse(R.Left, R.Top, R.Right, R.Bottom);
    end;
  end;
end;

// ThumbViewerPopupMenu starts here

procedure TMainForm.ReduceThumbnailSizeMenuItemClick(Sender: TObject);
begin
  ThumbnailPreview.Zoom := ThumbnailPreview.Zoom - 1;
end;

procedure TMainForm.EnlargeThumbnailSizeMenuItemClick(Sender: TObject);
begin
  ThumbnailPreview.Zoom := ThumbnailPreview.Zoom + 1;
end;

procedure TMainForm.HotTrackMenuItemClick(Sender: TObject);
begin
  ThumbnailPreview.HotTrack := not ThumbnailPreview.HotTrack;
end;

procedure TMainForm.MultiSelectMenuItemClick(Sender: TObject);
begin
  ThumbnailPreview.MultiSelect := not ThumbnailPreview.MultiSelect;
end;

procedure TMainForm.ArrangeLeftMenuItemClick(Sender: TObject);
begin
  ThumbnailPreview.IconOptions.Arrangement := iaLeft;
end;

procedure TMainForm.ArrangeTopMenuItemClick(Sender: TObject);
begin
  ThumbnailPreview.IconOptions.Arrangement := iaTop;
end;

procedure TMainForm.AddPageMenuItemClick(Sender: TObject);
begin
  if PrintPreview.BeginInsert(MaxInt) then
    try
      DrawRandomCircles;
    finally
      PrintPreview.EndInsert(False);
    end;
  PrintPreview.CurrentPage := PrintPreview.TotalPages;
end;

procedure TMainForm.DeleteAllMenuItemClick(Sender: TObject);
begin
  PrintPreview.Clear;
end;

procedure TMainForm.ThumbGrayAlwaysClick(Sender: TObject);
begin
  ThumbnailPreview.Grayscale := tgsAlways;
end;

procedure TMainForm.ThumbGrayNeverClick(Sender: TObject);
begin
  ThumbnailPreview.Grayscale := tgsNever;
end;

procedure TMainForm.ThumbGrayPreviewClick(Sender: TObject);
begin
  ThumbnailPreview.Grayscale := tgsPreview;
end;

procedure TMainForm.ThumbViewerPopupMenuPopup(Sender: TObject);
begin
  ReduceThumbnailSizeMenuItem.Enabled := (ThumbnailPreview.Zoom > 1);
  HotTrackMenuItem.Checked := ThumbnailPreview.HotTrack;
  MultiSelectMenuItem.Checked := ThumbnailPreview.MultiSelect;
  if ThumbnailPreview.IconOptions.Arrangement = iaTop then
    ArrangeTopMenuItem.Checked := True
  else
    ArrangeLeftMenuItem.Checked := True;
  DeleteAllMenuItem.Enabled := (PrintPreview.TotalPages > 0);
  ThumbGrayPreview.Checked := (ThumbnailPreview.Grayscale = tgsPreview);
  ThumbGrayAlways.Checked := (ThumbnailPreview.Grayscale = tgsAlways);
  ThumbGrayNever.Checked := (ThumbnailPreview.Grayscale = tgsNever);
end;

// PagePopupMenu starts here

procedure TMainForm.AddPageBeforeCurrentMenuItemClick(Sender: TObject);
begin
  if PrintPreview.BeginInsert(PrintPreview.CurrentPage) then
  begin
    try
      DrawRandomCircles;
    finally
      PrintPreview.EndInsert(False);
    end;
    PrintPreview.CurrentPage := PrintPreview.CurrentPage - 1;
  end;
end;

procedure TMainForm.AddPageAfterCurrentMenuItemClick(Sender: TObject);
begin
  if PrintPreview.BeginInsert(PrintPreview.CurrentPage + 1) then
  begin
    try
      DrawRandomCircles;
    finally
      PrintPreview.EndInsert(False);
    end;
    PrintPreview.CurrentPage := PrintPreview.CurrentPage + 1;
  end;
end;

procedure TMainForm.ReplaceCurrentMenuItemClick(Sender: TObject);
begin
  if PrintPreview.BeginReplace(PrintPreview.CurrentPage) then
    try
      DrawRandomCircles;
    finally
      PrintPreview.EndReplace(False);
    end;
end;

procedure TMainForm.EditCurrentMenuItemClick(Sender: TObject);
begin
  if PrintPreview.BeginEdit(PrintPreview.CurrentPage) then
    try
      DrawRandomCircles;
    finally
      PrintPreview.EndEdit(False);
    end;
end;

procedure TMainForm.DeleteCurrentClick(Sender: TObject);
begin
  PrintPreview.Delete(PrintPreview.CurrentPage);
end;

procedure TMainForm.MoveCurrentUpMenuItemClick(Sender: TObject);
begin
  PrintPreview.Exchange(PrintPreview.CurrentPage, PrintPreview.CurrentPage - 1);
end;

procedure TMainForm.MoveCurrentDownMenuItemClick(Sender: TObject);
begin
  PrintPreview.Exchange(PrintPreview.CurrentPage, PrintPreview.CurrentPage + 1);
end;

procedure TMainForm.MoveCurrentFirstMenuItemClick(Sender: TObject);
begin
  PrintPreview.Move(PrintPreview.CurrentPage, 1);
end;

procedure TMainForm.MoveCurrentLastMenuItemClick(Sender: TObject);
begin
  PrintPreview.Move(PrintPreview.CurrentPage, PrintPreview.TotalPages);
end;

procedure TMainForm.ExportCurrentClick(Sender: TObject);
begin
  if SavePictureDialog.Execute then
    PrintPreview.Pages[PrintPreview.CurrentPage].SaveToFile(SavePictureDialog.FileName);
end;

procedure TMainForm.CopyCurrentClick(Sender: TObject);
begin
  Clipboard.Assign(PrintPreview.Pages[PrintPreview.CurrentPage]);
end;

procedure TMainForm.PagePopupMenuPopup(Sender: TObject);
begin
  with PrintPreview do
  begin
    MoveCurrentFirstMenuItem.Enabled := (CurrentPage > 1);
    MoveCurrentUpMenuItem.Enabled := (CurrentPage > 1);
    MoveCurrentDownMenuItem.Enabled := (CurrentPage < TotalPages);
    MoveCurrentLastMenuItem.Enabled := (CurrentPage < TotalPages);
  end;
end;

// ZoomPopupMenu starts here

procedure TMainForm.ZoomActualSizeClick(Sender: TObject);
begin
  PrintPreview.Zoom := 100;
end;

procedure TMainForm.ZoomPageWidthClick(Sender: TObject);
begin
  PrintPreview.ZoomState := zsZoomToWidth;
end;

procedure TMainForm.ZoomPageHeightClick(Sender: TObject);
begin
  PrintPreview.ZoomState := zsZoomToHeight;
end;

procedure TMainForm.ZoomWholePageClick(Sender: TObject);
begin
  PrintPreview.ZoomState := zsZoomToFit;
end;

procedure TMainForm.ZoomPopupMenuPopup(Sender: TObject);
begin
end;

// UnitsPopupMenu starts here

procedure TMainForm.UnitsClick(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to UnitsPopupMenu.Items.Count - 1 do
    if Sender = UnitsPopupMenu.Items[I] then
    begin
      PrintPreview.Units := TUnits(I);
      GeneratePages;
      Break;
    end;
end;

procedure TMainForm.UnitsPopupMenuPopup(Sender: TObject);
begin
  UnitsPopupMenu.Items[Ord(PrintPreview.Units)].Checked := True;
end;

end.

