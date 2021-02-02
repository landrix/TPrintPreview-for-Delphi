{------------------------------------------------------------------------------}
{                                                                              }
{  Print Preview Components                                                    }
{  by Kambiz R. Khojasteh                                                      }
{                                                                              }
{  kambiz@delphiarea.com                                                       }
{  http://www.delphiarea.com                                                   }
{                                                                              }
{  TPrintPreview v5.94                                                         }
{  TPaperPreview v2.20                                                         }
{  TThumbnailPreview v2.12                                                     }
{                                                                              }
{------------------------------------------------------------------------------}

{------------------------------------------------------------------------------}
{  Use Synopse library to output preview as PDF document                       }
{  Get the library from http://www.synopse.info                                }
{------------------------------------------------------------------------------}
{.$DEFINE PDF_SYNOPSE}

{------------------------------------------------------------------------------}
{  Use dsPDF library to output preview as PDF document                         }
{  Get the newest library from http://delphistep.cis.si/dspdf.htm              }
{------------------------------------------------------------------------------}
{.$DEFINE PDF_DSPDF}

{------------------------------------------------------------------------------}
{  Use wPDF library to output preview as PDF document                          }
{  Get the newest library from http://www.wpcubed.com/products/wpdf/index.htm  }
{------------------------------------------------------------------------------}
{.$DEFINE PDF_WPDF}

{------------------------------------------------------------------------------}
{  Register Components in IDE                                                  }
{------------------------------------------------------------------------------}
{.$DEFINE REGISTER}

unit Preview;

interface

uses
  System.SysUtils, System.Classes,
  Winapi.Windows, Winapi.Messages, Winapi.WinSpool,
  Vcl.Graphics, Vcl.Controls,Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
  Vcl.ComCtrls, Vcl.Menus, Vcl.Printers;

{------------------------------------------------------------------------------}
{  If you need transparent image printing, set AllowTransparentDIB to True.    }
{                                                                              }
{  Note: Transparency on printers is not guaranteed. Instead, combine images   }
{  as needed, and then draw the final image to the printer.                    }
{------------------------------------------------------------------------------}
var
  AllowTransparentDIB: Boolean = False;

const
  crHand = 10;
  crGrab = 11; //pvg

type
  EPrintPreviewError = class(Exception);
  EPreviewLoadError = class(EPrintPreviewError);
  EPDFLibraryError = class(EPrintPreviewError);
  EPDFError = class(EPrintPreviewError);

  { TTemporaryFileStream }

  TTemporaryFileStream = class(THandleStream)
  public
    constructor Create;
    destructor Destroy; override;
  end;

  { TIntegerList }

  TIntegerList = class(TList)
  private
    function GetItems(Index: Integer): Integer;
    procedure SetItems(Index: Integer; Value: Integer);
  public
    function Add(Value: Integer): Integer;
    procedure Insert(Index: Integer; Value: Integer);
    function Remove(Value: Integer): Integer;
    function Extract(Value: Integer): Integer;
    function First: Integer;
    function Last: Integer;
    function IndexOf(Value: Integer): Integer;
    procedure Sort;
    procedure SaveToStream(Stream: TStream);
    procedure LoadFromStream(Stream: TStream);
    property Items[Index: Integer]: Integer read GetItems write SetItems; default;
  end;

  { TMetafileList }

  TMetafileList = class;

  TMetafileEntryState = (msInMemory, msInStorage, msDirty);
  TMetafileEntryStates = set of TMetafileEntryState;

  TMetafileEntry = class(TObject)
  private
    FOwner: TMetafileList;
    FMetafile: TMetafile;
    FStates: TMetafileEntryStates;
    FOffset: Int64;
    FSize: Int64;
    TouchCount: Integer;
    procedure MetafileChanged(Sender: TObject);
  protected
    constructor CreateInMemory(AOwner: TMetafileList; AMetafile: TMetafile);
    constructor CreateInStorage(AOwner: TMetafileList;
      const AOffset, ASize: Int64);
    procedure CopyToMemory;
    procedure CopyToStorage;
    function IsMoreRequiredThan(Another: TMetafileEntry): Boolean;
    procedure Touch;
    property Owner: TMetafileList read FOwner;
    property States: TMetafileEntryStates read FStates;
    property Offset: Int64 read FOffset;
    property Size: Int64 read FSize;
  public
    constructor Create(AOwner: TMetafileList);
    destructor Destroy; override;
    property Metafile: TMetafile read FMetafile;
  end;

  TSingleChangeEvent = procedure(Sender: TObject; Index: Integer) of object;
  TMultipleChangeEvent = procedure(Sender: TObject; StartIndex, EndIndex: Integer) of object;

  TMetafileList = class(TObject)
  private
    FEntries: TList;
    FCachedEntries: TList;
    FStorage: TStream;
    FCacheSize: Integer;
    FOnSingleChange: TSingleChangeEvent;
    FOnMultipleChange: TMultipleChangeEvent;
    function GetCount: Integer;
    function GetItems(Index: Integer): TMetafileEntry;
    function GetMetafiles(Index: Integer): TMetafile;
    procedure SetCacheSize(Value: Integer);
  protected
    procedure Reset;
    procedure ReduceCacheEntries(NumOfEntries: Integer);
    function GetCachedEntry(Index: Integer): TMetafileEntry;
    procedure EntryChanged(Entry: TMetafileEntry);
    procedure DoSingleChange(Index: Integer);
    procedure DoMultipleChange(StartIndex, EndIndex: Integer);
    property Storage: TStream read FStorage;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Clear;
    function Add(AMetafile: TMetafile): Integer;
    procedure Insert(Index: Integer; AMetafile: TMetafile);
    procedure Delete(Index: Integer);
    procedure Exchange(Index1, Index2: Integer);
    procedure Move(Index, NewIndex: Integer);
    function LoadFromStream(Stream: TStream): Boolean;
    procedure SaveToStream(Stream: TStream);
    procedure LoadFromFile(const FileName: String);
    procedure SaveToFile(const FileName: String);
    property Count: Integer read GetCount;
    property Items[Index: Integer]: TMetafileEntry read GetItems;
    property Metafiles[Index: Integer]: TMetafile read GetMetafiles; default;
    property CacheSize: Integer read FCacheSize write SetCacheSize;
    property OnSingleChange: TSingleChangeEvent read FOnSingleChange write FOnSingleChange;
    property OnMultipleChange: TMultipleChangeEvent read FOnMultipleChange write FOnMultipleChange;
  end;

  { TPaperPreviewOptions }

  TUpdateSeverity = (usNone, usRedraw, usRecreate);

  TPaperPreviewChangeEvent = procedure(Sender: TObject;
    Severity: TUpdateSeverity) of object;

  TPaperPreviewOptions = class(TPersistent)
  private
    FPaperColor: TColor;
    FBorderColor: TColor;
    FBorderWidth: TBorderWidth;
    FShadowColor: TColor;
    FShadowWidth: TBorderWidth;
    FCursor: TCursor;
    FDragCursor: TCursor;
    FGrabCursor: TCursor; //pvg
    FPopupMenu: TPopupMenu;
    FHint: String;
    FOnChange: TPaperPreviewChangeEvent;
    procedure SetPaperColor(Value: TColor);
    procedure SetBorderColor(Value: TColor);
    procedure SetBorderWidth(Value: TBorderWidth);
    procedure SetShadowColor(Value: TColor);
    procedure SetShadowWidth(Value: TBorderWidth);
    procedure SetCursor(Value: TCursor);
    procedure SetDragCursor(Value: TCursor);
    procedure SetGrabCursor(Value: TCursor); //pvg
    procedure SetPopupMenu(Value: TPopupMenu);
    procedure SetHint(const Value: String);
  protected
    procedure DoChange(Severity: TUpdateSeverity);
  public
    constructor Create;
    procedure Assign(Source: TPersistent); override;
    procedure AssignTo(Dest: TPersistent); override;
    procedure CalcDimensions(PaperWidth, PaperHeight: Integer;
      out PaperRect, BoxRect: TRect);
    procedure Draw(Canvas: TCanvas; const BoxRect: TRect);
    property OnChange: TPaperPreviewChangeEvent read FOnChange write FOnChange;
  published
    property BorderColor: TColor read FBorderColor write SetBorderColor default clBlack;
    property BorderWidth: TBorderWidth read FBorderWidth write SetBorderWidth default 1;
    property Cursor: TCursor read FCursor write SetCursor default crDefault;
    property DragCursor: TCursor read FDragCursor write SetDragCursor default crHand;
    property GrabCursor: TCursor read FGrabCursor write SetGrabCursor default crGrab; //pvg
    property Hint: String read FHint write SetHint;
    property PaperColor: TColor read FPaperColor write SetPaperColor default clWhite;
    property PopupMenu: TPopupMenu read FPopupMenu write SetPopupMenu;
    property ShadowColor: TColor read FShadowColor write SetShadowColor default clBtnShadow;
    property ShadowWidth: TBorderWidth read FShadowWidth write SetShadowWidth default 3;
  end;

  { TPaperPreview }

  TPaperPaintEvent = procedure(Sender: TObject; Canvas: TCanvas;
    const Rect: TRect) of object;

  TPaperPreview = class(TCustomControl)
  private
    FPreservePaperSize: Boolean;
    FPaperColor: TColor;
    FBorderColor: TColor;
    FBorderWidth: TBorderWidth;
    FShadowColor: TColor;
    FShadowWidth: TBorderWidth;
    FShowCaption: Boolean;
    FAlignment: TAlignment;
    FWordWrap: Boolean;
    FCaptionHeight: Integer;
    FOnResize: TNotifyEvent;
    FOnPaint: TPaperPaintEvent;
    FOnMouseEnter: TNotifyEvent;
    FOnMouseLeave: TNotifyEvent;
    FPageRect: TRect;
    OffScreen: TBitmap;
    IsOffScreenPrepared: Boolean;
    IsOffScreenReady: Boolean;
    LastVisibleRect: TRect;
    LastVisiblePageRect: TRect;
    PageCanvas: TCanvas;
    procedure SetPaperWidth(Value: Integer);
    function GetPaperWidth: Integer;
    procedure SetPaperHeight(Value: Integer);
    function GetPaperHeight: Integer;
    function GetPaperSize: TPoint;
    procedure SetPaperSize(const Value: TPoint);
    procedure SetPaperColor(Value: TColor);
    procedure SetBorderColor(Value: TColor);
    procedure SetBorderWidth(Value: TBorderWidth);
    procedure SetShadowColor(Value: TColor);
    procedure SetShadowWidth(Value: TBorderWidth);
    procedure SetShowCaption(Value: Boolean);
    procedure SetAlignment(Value: TAlignment);
    procedure SetWordWrap(Value: Boolean);
    procedure UpdateCaptionHeight;
    procedure WMSize(var Message: TWMSize); message WM_SIZE;
    procedure WMEraseBkgnd(var Message: TWMEraseBkgnd); message WM_ERASEBKGND;
    procedure CMMouseEnter(var Message: TMessage); message CM_MOUSEENTER;
    procedure CMMouseLeave(var Message: TMessage); message CM_MOUSELEAVE;
    procedure CMColorChanged(var Message: TMessage); message CM_COLORCHANGED;
    procedure CMFontChanged(var Message: TMessage); message CM_FONTCHANGED;
    procedure CMTextChanged(var Message: TMessage); message CM_TEXTCHANGED;
    procedure BiDiModeChanged(var Message: TMessage); message CM_BIDIMODECHANGED;
  protected
    procedure Paint; override;
    procedure DrawPage(Canvas: TCanvas); virtual;
    function ActualWidth(Value: Integer): Integer; virtual;
    function ActualHeight(Value: Integer): Integer; virtual;
    function LogicalWidth(Value: Integer): Integer; virtual;
    function LogicalHeight(Value: Integer): Integer; virtual;
    procedure InvalidateAll; virtual;
    property CaptionHeight: Integer read FCaptionHeight;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure Invalidate; override;
    function ClientToPaper(const Pt: TPoint): TPoint;
    function PaperToClient(const Pt: TPoint): TPoint;
    procedure SetBoundsEx(ALeft, ATop, APaperWidth, APaperHeight: Integer);
    property PaperSize: TPoint read GetPaperSize write SetPaperSize;
    property PageRect: TRect read FPageRect;
  published
    property Align;
    property Alignment: TAlignment read FAlignment write SetAlignment default taCenter;
    property BiDiMode;
    property BorderColor: TColor read FBorderColor write SetBorderColor default clBlack;
    property BorderWidth: TBorderWidth read FBorderWidth write SetBorderWidth default 1;
    property Caption;
    property Color;
    property Cursor;
    property DragCursor;
    property DragMode;
    property Font;
    property ParentBiDiMode;
    property ParentColor;
    property ParentFont;
    property ParentShowHint;
    property PopupMenu;
    property PaperColor: TColor read FPaperColor write SetPaperColor default clWhite;
    property PaperWidth: Integer read GetPaperWidth write SetPaperWidth;
    property PaperHeight: Integer read GetPaperHeight write SetPaperHeight;
    property PreservePaperSize: Boolean read FPreservePaperSize write FPreservePaperSize default True;
    property ShadowColor: TColor read FShadowColor write SetShadowColor default clBtnShadow;
    property ShadowWidth: TBorderWidth read FShadowWidth write SetShadowWidth default 3;
    property ShowCaption: Boolean read FShowCaption write SetShowCaption default False;
    property ShowHint;
    property TabOrder;
    property TabStop;
    property Visible;
    property WordWrap: Boolean read FWordWrap write SetWordWrap default True;
    property OnClick;
    property OnDblClick;
    property OnDragDrop;
    property OnDragOver;
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
    property OnMouseDown;
    property OnMouseMove;
    property OnMouseUp;
    property OnMouseEnter: TNotifyEvent read FOnMouseEnter write FOnMouseEnter;
    property OnMouseLeave: TNotifyEvent read FOnMouseLeave write FOnMouseLeave;
    property OnResize: TNotifyEvent read FOnResize write FOnResize;
    property OnPaint: TPaperPaintEvent read FOnPaint write FOnPaint;
  end;

  { TPrintPreview}

  TPDFDocumentInfo = class;
  TThumbnailPreview = class;

  TVertAlign = (vaTop, vaCenter, vaBottom);  //rmk
  THorzAlign = (haLeft, haCenter, haRight);  //rmk

  TGrayscaleOption = (gsPreview, gsPrint);
  TGrayscaleOptions = set of TGrayscaleOption;

  TPreviewState = (psReady, psCreating, psPrinting, psEditing, psReplacing,
    psInserting, psLoading, psSaving, psSavingPDF, psSavingTIF);

  TZoomState = (zsZoomOther, zsZoomToWidth, zsZoomToHeight, zsZoomToFit);

  TUnits = (mmPixel, mmLoMetric, mmHiMetric, mmLoEnglish, mmHiEnglish, mmTWIPS, mmPoints);

  TPaperType = (pLetter, pLetterSmall, pTabloid, pLedger, pLegal, pStatement,
    pExecutive, pA3, pA4, pA4Small, pA5, pB4, pB5, pFolio, pQuatro, p10x14,
    p11x17, pNote, pEnv9, pEnv10, pEnv11, pEnv12, pEnv14, pCSheet, pDSheet,
    pESheet, pEnvDL, pEnvC5, pEnvC3, pEnvC4, pEnvC6, pEnvC65, pEnvB4, pEnvB5,
    pEnvB6, pEnvItaly, pEnvMonarch, pEnvPersonal, pFanfoldUSStd, pFanfoldGermanStd,
    pFanfoldGermanLegal, pB4ISO, pJapanesePostcard, p9x11, p10x11, p15x11,
    pEnvInvite, pLetterExtra, pLegalExtra, pTabloidExtra, pA4Extra, pLetterTransverse,
    pA4Transverse, pLetterExtraTransverse, pAPlus, pBPlus, pLetterPlus, pA4Plus,
    pA5Transverse, pB5Transverse, pA3Extra, pA5Extra, pB5Extra, pA2, pA3Transverse,
    pA3ExtraTransverse, pCustom);

  TPageProcessingChoice = (pcAccept, pcIgnore, pcCancellAll);

  TPreviewPageProcessingEvent = procedure(Sender: TObject; PageNo: Integer;
    var Choice: TPageProcessingChoice) of object;

  TPreviewPageDrawEvent = procedure(Sender: TObject; PageNo: Integer;
    Canvas: TCanvas) of object;

  TPreviewProgressEvent = procedure(Sender: TObject; Done, Total: Integer) of object;

  TPrintPreview = class(TScrollBox)
  private
    FThumbnailViews: TList;
    FPaperView: TPaperPreview;
    FPaperViewOptions: TPaperPreviewOptions;
    FPrintJobTitle: String;
    FPageList: TMetafileList;
    FPageCanvas: TCanvas;
    FUnits: TUnits;
    FDeviceExt: TPoint;
    FLogicalExt: TPoint;
    FPageExt: TPoint;
    FOrientation: TPrinterOrientation;
    FCurrentPage: Integer;
    FPaperType: TPaperType;
    FState: TPreviewState;
    FZoom: Integer;
    FZoomState: TZoomState;
    FZoomSavePos: Boolean;
    FZoomMin: Integer;
    FZoomMax: Integer;
    FZoomStep: Integer;
    FLastZoom: Integer;
    FUsePrinterOptions: Boolean;
    FDirectPrint: Boolean;
    FDirectPrinting: Boolean;
    FDirectPrintPageCount: Integer;
    FOldMousePos: TPoint;
    FCanScrollHorz: Boolean;
    FCanScrollVert: Boolean;
    FIsDragging: Boolean;
    FCanvasPageNo: Integer;
    FFormName: String;
    FVirtualFormName: String;
    FAnnotation: Boolean;
    FBackground: Boolean;
    FGrayscale: TGrayscaleOptions;
    FGrayBrightness: Integer;
    FGrayContrast: Integer;
    FShowPrintableArea: Boolean;
    FPrintableAreaColor: TColor;
    FPDFDocumentInfo: TPDFDocumentInfo;
    FOnBeginDoc: TNotifyEvent;
    FOnEndDoc: TNotifyEvent;
    FOnNewPage: TNotifyEvent;
    FOnEndPage: TNotifyEvent;
    FOnChange: TNotifyEvent;
    FOnStateChange: TNotifyEvent;
    FOnPaperChange: TNotifyEvent;
    FOnProgress: TPreviewProgressEvent;
    FOnPageProcessing: TPreviewPageProcessingEvent;
    FOnBeforePrint: TNotifyEvent;
    FOnAfterPrint: TNotifyEvent;
    FOnZoomChange: TNotifyEvent;
    FOnAnnotation: TPreviewPageDrawEvent;
    FOnBackground: TPreviewPageDrawEvent;
    FOnPrintAnnotation: TPreviewPageDrawEvent;
    FOnPrintBackground: TPreviewPageDrawEvent;
    PageMetafile: TMetafile;
    AnnotationMetafile: TMetafile;
    BackgroundMetafile: TMetafile;
    WheelAccumulator: Integer;
    ReferenceDC: HDC;
    procedure SetPaperViewOptions(Value: TPaperPreviewOptions);
    procedure SetUnits(Value: TUnits);
    procedure SetPaperType(Value: TPaperType);
    function GetPaperWidth: Integer;
    procedure SetPaperWidth(Value: Integer);
    function GetPaperHeight: Integer;
    procedure SetPaperHeight(Value: Integer);
    procedure SetAnnotation(Value: Boolean);
    procedure SetBackground(Value: Boolean);
    procedure SetGrayscale(Value: TGrayscaleOptions);
    procedure SetGrayBrightness(Value: Integer);
    procedure SetGrayContrast(Value: Integer);
    function GetCacheSize: Integer;
    procedure SetCacheSize(Value: Integer);
    function GetFormName: String;
    procedure SetFormName(const Value: String);
    procedure SetPDFDocumentInfo(Value: TPDFDocumentInfo);
    function GetPageBounds: TRect;
    function GetPrinterPageBounds: TRect;
    function GetPrinterPhysicalPageBounds: TRect;
    procedure SetOrientation(Value: TPrinterOrientation);
    procedure SetZoomState(Value: TZoomState);
    procedure SetZoom(Value: Integer);
    procedure SetZoomMin(Value: Integer);
    procedure SetZoomMax(Value: Integer);
    procedure SetCurrentPage(Value: Integer);
    function GetTotalPages: Integer;
    function GetPages(PageNo: Integer): TMetafile;
    function GetCanvas: TCanvas;
    function GetPrinterInstalled: Boolean;
    function GetPrinter: TPrinter;
    procedure SetShowPrintableArea(Value: Boolean);
    procedure SetPrintableAreaColor(Value: TColor);
    procedure SetDirectPrint(Value: Boolean);
    function GetIsDummyFormName: Boolean;
    function GetSystemDefaultUnits: TUnits;
    function GetUserDefaultUnits: TUnits;
    function IsZoomStored: Boolean;
    procedure PaperClick(Sender: TObject);
    procedure PaperDblClick(Sender: TObject);
    procedure PaperMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure PaperMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure PaperMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure PaperViewOptionsChanged(Sender: TObject; Severity: TUpdateSeverity);
    procedure PagesChanged(Sender: TObject; PageStartIndex, PageEndIndex: Integer);
    procedure PageChanged(Sender: TObject; PageIndex: Integer);
    procedure PaintPage(Sender: TObject; Canvas: TCanvas; const Rect: TRect);
    procedure CNKeyDown(var Message: TWMKey); message CN_KEYDOWN;
    procedure WMMouseWheel(var Message: TMessage); message WM_MOUSEWHEEL;
    procedure WMEraseBkgnd(var Message: TWMEraseBkgnd); message WM_ERASEBKGND;
    procedure WMPaint(var Message: TWMPaint); message WM_PAINT;
    procedure WMHScroll(var Message: TWMScroll); message WM_HSCROLL;
    procedure WMVScroll(var Message: TWMScroll); message WM_VSCROLL;
  protected
    procedure Loaded; override;
    procedure Resize; override;
    procedure DoPaperChange; virtual;
    procedure DoAnnotation(PageNo: Integer); virtual;
    procedure DoBackground(PageNo: Integer); virtual;
    procedure DoProgress(Done, Total: Integer); virtual;
    function DoPageProcessing(PageNo: Integer): TPageProcessingChoice; virtual;
    procedure ChangeState(NewState: TPreviewState);
    procedure PreviewPage(PageNo: Integer; Canvas: TCanvas; const Rect: TRect); virtual;
    procedure PrintPage(PageNo: Integer; Canvas: TCanvas; const Rect: TRect); virtual;
    function FindPaperTypeBySize(APaperWidth, APaperHeight: Integer): TPaperType;
    function FindPaperTypeByID(ID: Integer): TPaperType;
    function GetPaperTypeSize(APaperType: TPaperType;
      out APaperWidth, APaperHeight: Integer; OutUnits: TUnits): Boolean;
    procedure SetPaperSize(AWidth, AHeight: Integer);
    procedure SetPaperSizeOrientation(AWidth, AHeight: Integer;
      AOrientation: TPrinterOrientation);
    procedure ResetPrinterDC;
    procedure InitializePrinting; virtual;
    procedure FinalizePrinting(Succeeded: Boolean); virtual;
    function GetVisiblePageRect: TRect;
    procedure SetVisiblePageRect(const Value: TRect);
    procedure UpdateZoomEx(X, Y: Integer); virtual;
    function CalculateViewSize(const Space: TPoint): TPoint; virtual;
    procedure UpdateExtends; virtual;
    procedure CreateMetafileCanvas(out AMetafile: TMetafile; out ACanvas: TCanvas); virtual;
    procedure CloseMetafileCanvas(var AMetafile: TMetafile; var ACanvas: TCanvas); virtual;
    procedure CreatePrinterCanvas(out ACanvas: TCanvas); virtual;
    procedure ClosePrinterCanvas(var ACanvas: TCanvas); virtual;
    procedure ScaleCanvas(ACanvas: TCanvas); virtual;
  public
    function HorzPixelsPerInch: Integer; virtual;
    function VertPixelsPerInch: Integer; virtual;
  protected
    procedure RegisterThumbnailViewer(ThumbnailView: TThumbnailPreview); virtual;
    procedure UnregisterThumbnailViewer(ThumbnailView: TThumbnailPreview); virtual;
    procedure RebuildThumbnails; virtual;
    procedure UpdateThumbnails(StartIndex, EndIndex: Integer); virtual;
    procedure RepaintThumbnails(StartIndex, EndIndex: Integer); virtual;
    procedure RecolorThumbnails(OnlyGrays: Boolean); virtual;
    procedure SyncThumbnail; virtual;
    function LoadPageInfo(Stream: TStream): Boolean; virtual;
    procedure SavePageInfo(Stream: TStream); virtual;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure ConvertPoints(var Points; NumPoints: Integer; InUnits, OutUnits: TUnits);
    function ConvertXY(X, Y: Integer; InUnits, OutUnits: TUnits): TPoint;
    function ConvertX(X: Integer; InUnits, OutUnits: TUnits): Integer;
    function ConvertY(Y: Integer; InUnits, OutUnits: TUnits): Integer;
    function ConvertRect(Rec: TRect; InUnits: TUnits; OutUnits: TUnits): TRect;
    function BoundsFrom(AUnits: TUnits; ALeft, ATop, AWidth, AHeight: Integer): TRect;
    function RectFrom(AUnits: TUnits; ALeft, ATop, ARight, ABottom: Integer): TRect;
    function PointFrom(AUnits: TUnits; X, Y: Integer): TPoint;
    function XFrom(AUnits: TUnits; X: Integer): Integer;
    function YFrom(AUnits: TUnits; Y: Integer): Integer;
    function ScreenToPreview(X, Y: Integer): TPoint;
    function PreviewToScreen(X, Y: Integer): TPoint;
    function ScreenToPaper(const Pt: TPoint): TPoint;
    function PaperToScreen(const Pt: TPoint): TPoint;
    function ClientToPaper(const Pt: TPoint): TPoint;
    function PaperToClient(const Pt: TPoint): TPoint;
    function PaintGraphic(X, Y: Integer; Graphic: TGraphic): TPoint;
    function PaintGraphicEx(const Rect: TRect; Graphic: TGraphic;
      Proportinal, ShrinkOnly, Center: Boolean): TRect;
    function PaintGraphicEx2(const Rect: TRect; Graphic: TGraphic;   //rmk
      VertAlign: TVertAlign; HorzAlign: THorzAlign): TRect;          //rmk
    function PaintWinControl(X, Y: Integer; WinControl: TWinControl): TPoint;
    function PaintWinControlEx(const Rect: TRect; WinControl: TWinControl;
      Proportinal, ShrinkOnly, Center: Boolean): TRect;
    function PaintWinControlEx2(const Rect: TRect; WinControl: TWinControl;
      VertAlign: TVertAlign; HorzAlign: THorzAlign): TRect;
    function PaintRichText(const Rect: TRect; RichEdit: TCustomRichEdit;
      MaxPages: Integer; pOffset: PInteger = nil): Integer;
    function GetRichTextRect(var Rect: TRect; RichEdit: TCustomRichEdit;
      pOffset: PInteger = nil): Integer;
    procedure Clear;
    function Delete(PageNo: Integer): Boolean;
    function Move(PageNo, NewPageNo: Integer): Boolean;
    function Exchange(PageNo1, PageNo2: Integer): Boolean;
    function BeginReplace(PageNo: Integer): Boolean;
    procedure EndReplace(Cancel: Boolean = False);
    function BeginEdit(PageNo: Integer): Boolean;
    procedure EndEdit(Cancel: Boolean = False);
    function BeginInsert(PageNo: Integer): Boolean;
    procedure EndInsert(Cancel: Boolean = False);
    function BeginAppend: Boolean;
    procedure EndAppend(Cancel: Boolean = False);
    procedure BeginDoc;
    procedure EndDoc;
    procedure NewPage;
    procedure Print;
    procedure PrintPages(FromPage, ToPage: Integer);
    procedure PrintPagesEx(Pages: TIntegerList);
    procedure UpdateZoom;
    procedure UpdateAnnotation;
    procedure UpdateBackground;
    procedure SetPrinterOptions;
    procedure GetPrinterOptions;
    procedure SetPageSetupParameters(PageSetupDialog: TPageSetupDialog);
    function GetPageSetupParameters(PageSetupDialog: TPageSetupDialog): TRect;
    function FetchFormNames(FormNames: TStrings): Boolean;
    function GetFormSize(const AFormName: String; out FormWidth, FormHeight: Integer): Boolean;
    function AddNewForm(const AFormName: String; FormWidth, FormHeight: DWORD): Boolean;
    function RemoveForm(const AFormName: String): Boolean;
    procedure DrawPage(PageNo: Integer; Canvas: TCanvas; const Rect: TRect; Gray: Boolean); virtual;
    procedure LoadFromStream(Stream: TStream);
    procedure SaveToStream(Stream: TStream);
    procedure LoadFromFile(const FileName: String);
    procedure SaveToFile(const FileName: String);
    procedure SaveAsTIF(const FileName: String);
    function CanSaveAsTIF: Boolean;
    procedure SaveAsPDF(const FileName: String);
    function CanSaveAsPDF: Boolean;
    function IsPaperCustom: Boolean;
    function IsPaperRotated: Boolean;
    property Canvas: TCanvas read GetCanvas;
    property CanvasPageNo: Integer read FCanvasPageNo;
    property TotalPages: Integer read GetTotalPages;
    property State: TPreviewState read FState;
    property PageSize: TPoint read FPageExt;
    property PageDevicePixels: TPoint read FDeviceExt;
    property PageLogicalPixels: TPoint read FLogicalExt;
    property PageBounds: TRect read GetPageBounds;
    property PrinterPageBounds: TRect read GetPrinterPageBounds;
    property PrinterPhysicalPageBounds: TRect read GetPrinterPhysicalPageBounds;
    property PrinterInstalled: Boolean read GetPrinterInstalled;
    property Printer: TPrinter read GetPrinter;
    property PaperViewControl: TPaperPreview read FPaperView;
    property CurrentPage: Integer read FCurrentPage write SetCurrentPage;
    property FormName: String read GetFormName write SetFormName;
    property IsDummyFormName: Boolean read GetIsDummyFormName;
    property SystemDefaultUnits: TUnits read GetSystemDefaultUnits;
    property UserDefaultUnits: TUnits read GetUserDefaultUnits;
    property CanScrollHorz: Boolean read FCanScrollHorz;
    property CanScrollVert: Boolean read FCanScrollVert;
    property Pages[PageNo: Integer]: TMetafile read GetPages;
  published
    property Align default alClient;
    property Annotation: Boolean read FAnnotation write SetAnnotation default False;
    property Background: Boolean read FBackground write SetBackground default False;
    property CacheSize: Integer read GetCacheSize write SetCacheSize default 10;
    property DirectPrint: Boolean read FDirectPrint write SetDirectPrint default False;
    property Grayscale: TGrayscaleOptions read FGrayscale write SetGrayscale default [];
    property GrayBrightness: Integer read FGrayBrightness write SetGrayBrightness default 0;
    property GrayContrast: Integer read FGrayContrast write SetGrayContrast default 0;
    property Units: TUnits read FUnits write SetUnits default mmHiMetric;
    property Orientation: TPrinterOrientation read FOrientation write SetOrientation default poPortrait;
    property PaperType: TPaperType read FPaperType write SetPaperType default pA4;
    property PaperView: TPaperPreviewOptions read FPaperViewOptions write SetPaperViewOptions;
    property PaperWidth: Integer read GetPaperWidth write SetPaperWidth stored IsPaperCustom;
    property PaperHeight: Integer read GetPaperHeight write SetPaperHeight stored IsPaperCustom;
    property ParentFont default False;
    property PDFDocumentInfo: TPDFDocumentInfo read FPDFDocumentInfo write SetPDFDocumentInfo;
    property PrintableAreaColor: TColor read FPrintableAreaColor write SetPrintableAreaColor default clSilver;
    property PrintJobTitle: String read FPrintJobTitle write FPrintJobTitle;
    property ShowPrintableArea: Boolean read fShowPrintableArea write SetShowPrintableArea default False;
    property TabStop default True;
    property UsePrinterOptions: Boolean read FUsePrinterOptions write FUsePrinterOptions default False;
    property ZoomState: TZoomState read FZoomState write SetZoomState default zsZoomToFit;
    property Zoom: Integer read FZoom write SetZoom stored IsZoomStored;
    property ZoomMin: Integer read FZoomMin write SetZoomMin default 10;
    property ZoomMax: Integer read FZoomMax write SetZoomMax default 1000;
    property ZoomSavePos: Boolean read FZoomSavePos write FZoomSavePos default True;
    property ZoomStep: Integer read FZoomStep write FZoomStep default 10;
    property OnBeginDoc: TNotifyEvent read FOnBeginDoc write FOnBeginDoc;
    property OnEndDoc: TNotifyEvent read FOnEndDoc write FOnEndDoc;
    property OnNewPage: TNotifyEvent read FOnNewPage write FOnNewPage;
    property OnEndPage: TNotifyEvent read FOnEndPage write FOnEndPage;
    property OnChange: TNotifyEvent read FOnChange write FOnChange;
    property OnStateChange: TNotifyEvent read FOnStateChange write FOnStateChange;
    property OnZoomChange: TNotifyEvent read FOnZoomChange write FOnZoomChange;
    property OnPaperChange: TNotifyEvent read FOnPaperChange write FOnPaperChange;
    property OnProgress: TPreviewProgressEvent read FOnProgress write FOnProgress;
    property OnPageProcessing: TPreviewPageProcessingEvent read FOnPageProcessing write FOnPageProcessing;
    property OnBeforePrint: TNotifyEvent read FOnBeforePrint write FOnBeforePrint;
    property OnAfterPrint: TNotifyEvent read FOnAfterPrint write FOnAfterPrint;
    property OnAnnotation: TPreviewPageDrawEvent read FOnAnnotation write FOnAnnotation;
    property OnBackground: TPreviewPageDrawEvent read FOnBackground write FOnBackground;
    property OnPrintAnnotation: TPreviewPageDrawEvent read FOnPrintAnnotation write FOnPrintAnnotation;
    property OnPrintBackground: TPreviewPageDrawEvent read FOnPrintBackground write FOnPrintBackground;
  end;

  { TThumbnailDragObject }

  TThumbnailDragObject = class(TDragControlObject)
  private
    FDragImages: TDragImageList;
    FPageNo: Integer;
    FDropAfter: Boolean;
  protected
    function GetDragImages: TDragImageList; override;
    function GetDragCursor(Accepted: Boolean; X, Y: Integer): TCursor; override;
  public
    constructor Create(AControl: TThumbnailPreview; APageNo: Integer); reintroduce;
    destructor Destroy; override;
    procedure HideDragImage; override;
    procedure ShowDragImage; override;
    property PageNo: Integer read FPageNo;
    property DropAfter: Boolean read FDropAfter write FDropAfter;
  end;

  { TThumbnailPreview }

  TThumbnailGrayscale = (tgsPreview, tgsNever, tgsAlways);

  TThumbnailMarkerOption = (moMove, moSizeTopLeft, moSizeTopRight,
    moSizeBottomLeft, moSizeBottomRight);

  TThumbnailMarkerAction = (maNone, maMove, maResize);

  TPageNotifyEvent = procedure(Sender: TObject; PageNo: Integer) of object;

  TPageInfoTipEvent = procedure(Sender: TObject; PageNo: Integer;
    var InfoTip: String) of object;

  TPageThumbnailDrawEvent = procedure(Sender: TObject; PageNo: Integer;
    Canvas: TCanvas; const Rect: TRect; var DefaultDraw: Boolean) of object;

  TThumbnailPreview = class(TCustomListView)
  private
    FZoom: Integer;
    FMarkerColor: TColor;
    FSpacingHorizontal: Integer;
    FSpacingVertical: Integer;
    FGrayscale: TThumbnailGrayscale;
    FIsGrayscaled: Boolean;
    FPrintPreview: TPrintPreview;
    FPaperViewOptions: TPaperPreviewOptions;
    FCurrentIndex: Integer;
    FAllowReorder: Boolean;
    FDropTarget: Integer;
    FDisableTheme: Boolean;
    FOnPageBeforeDraw: TPageThumbnailDrawEvent;
    FOnPageAfterDraw: TPageThumbnailDrawEvent;
    FOnPageInfoTip: TPageInfoTipEvent;
    FOnPageClick: TPageNotifyEvent;
    FOnPageDblClick: TPageNotifyEvent;
    FOnPageSelect: TPageNotifyEvent;
    FOnPageUnselect: TPageNotifyEvent;
    PageRect, BoxRect: TRect;
    Page: TBitmap;
    CursorPageNo: Integer;
    DefaultDragObject: TThumbnailDragObject;
    MarkerRect, UpdatingMarkerRect: TRect;
    MarkerOfs, MarkerPivotPt: TPoint;
    MarkerAction: TThumbnailMarkerAction;
    MarkerDragging: Boolean;
    procedure SetZoom(Value: Integer);
    procedure SetMarkerColor(Value: TColor);
    procedure SetSpacingHorizontal(Value: Integer);
    procedure SetSpacingVertical(Value: Integer);
    procedure SetGrayscale(Value: TThumbnailGrayscale);
    procedure SetPrintPreview(Value: TPrintPreview);
    procedure SetPaperViewOptions(Value: TPaperPreviewOptions);
    procedure PaperViewOptionsChanged(Sender: TObject; Severity: TUpdateSeverity);
    procedure SetCurrentIndex(Index: Integer);
    function GetSelected: Integer;
    procedure SetSelected(Value: Integer);
    procedure SetDisableTheme(Value: Boolean);
    procedure CMHintShow(var Message: TCMHintShow); message CM_HINTSHOW;
    procedure CMFontChanged(var Message: TMessage); message CM_FONTCHANGED;
    procedure WMSetCursor(var Message: TWMSetCursor); message WM_SETCURSOR;
    procedure WMEraseBkgnd(var Message: TWMEraseBkgnd); message WM_ERASEBKGND;
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure CreateWnd; override;
    procedure DestroyWnd; override;
    procedure ApplySpacing; virtual;
    procedure InsertMark(Index: Integer; After: Boolean);
    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure Click; override;
    procedure DblClick; override;
    function GetPopupMenu: TPopupMenu; override;
    function OwnerDataFetch(Item: TListItem; Request: TItemRequest): Boolean; override;
    function OwnerDataHint(StartIndex, EndIndex: Integer): Boolean; override;
    function IsCustomDrawn(Target: TCustomDrawTarget; Stage: TCustomDrawStage): Boolean; override;
    function CustomDrawItem(Item: TListItem; State: TCustomDrawState;
      Stage: TCustomDrawStage): Boolean; override;
    procedure Change(Item: TListItem; Change: Integer); override;
    procedure DragOver(Source: TObject; X, Y: Integer; State: TDragState;
      var Accept: Boolean); override;
    procedure DoStartDrag(var DragObject: TDragObject); override;
    procedure DoEndDrag(Target: TObject; X, Y: Integer); override;
    procedure RebuildThumbnails;
    procedure UpdateThumbnails(StartIndex, EndIndex: Integer);
    procedure RepaintThumbnails(StartIndex, EndIndex: Integer);
    procedure RecolorThumbnails;
    procedure InvalidateMarker(Rect: TRect);
    function GetMarkerArea: TRect;
    procedure SetMarkerArea(const Value: TRect);
    property CurrentIndex: Integer read FCurrentIndex write SetCurrentIndex;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure DragDrop(Source: TObject; X, Y: Integer); override;
    function PageAtCursor: Integer;
    function PageAt(X, Y: Integer): Integer;
    procedure GetSelectedPages(Pages: TIntegerList);
    procedure SetSelectedPages(Pages: TIntegerList);
    procedure DeleteSelected; override;
    procedure PrintSelected; virtual;
    property IsGrayscaled: Boolean read FIsGrayscaled;
    property Selected: Integer read GetSelected write SetSelected;
    property DropTarget: Integer read FDropTarget;
  published
    property Align default alLeft;
    property AllowReorder: Boolean read FAllowReorder write FAllowReorder default False;
    property Anchors;
    property BevelEdges;
    property BevelInner;
    property BevelOuter;
    property BevelKind default bkNone;
    property BevelWidth;
    property BiDiMode;
    property BorderStyle;
    property BorderWidth;
    property Color;
    property Constraints;
    property Ctl3D;
    property DisableTheme: Boolean read FDisableTheme write SetDisableTheme default False;
    property DragCursor;
    property DragKind;
    property DragMode;
    property Enabled;
    property Font;
    property FlatScrollBars;
    property Grayscale: TThumbnailGrayscale read FGrayscale write SetGrayscale default tgsPreview;
    property HideSelection;
    property HotTrack;
    property HotTrackStyles;
    property HoverTime;
    property IconOptions;
    property MarkerColor: TColor read FMarkerColor write SetMarkerColor default clBlue;
    property MultiSelect;
    property PaperView: TPaperPreviewOptions read FPaperViewOptions write SetPaperViewOptions;
    property ParentColor default True;
    property ParentBiDiMode;
    property ParentFont;
    property ParentShowHint;
    property PopupMenu;
    property PrintPreview: TPrintPreview read FPrintPreview write SetPrintPreview;
    property ShowHint;
    property SpacingHorizontal: Integer read FSpacingHorizontal write SetSpacingHorizontal default 8;
    property SpacingVertical: Integer read FSpacingVertical write SetSpacingVertical default 8;
    property TabOrder;
    property TabStop default True;
    property Visible;
    property Zoom: Integer read FZoom write SetZoom default 10;
    property OnClick;
    property OnEndDock;
    property OnEndDrag;
    property OnEnter;
    property OnExit;
    property OnDragDrop;
    property OnDragOver;
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
    property OnMouseDown;
    property OnMouseMove;
    property OnMouseUp;
    property OnResize;
    property OnStartDock;
    property OnStartDrag;
    property OnPageBeforeDraw: TPageThumbnailDrawEvent read FOnPageBeforeDraw write FOnPageBeforeDraw;
    property OnPageAfterDraw: TPageThumbnailDrawEvent read FOnPageAfterDraw write FOnPageAfterDraw;
    property OnPageInfoTip: TPageInfoTipEvent read FOnPageInfoTip write FOnPageInfoTip;
    property OnPageClick: TPageNotifyEvent read FOnPageClick write FOnPageClick;
    property OnPageDblClick: TPageNotifyEvent read FOnPageDblClick write FOnPageDblClick;
    property OnPageSelect: TPageNotifyEvent read FOnPageSelect write FOnPageSelect;
    property OnPageUnselect: TPageNotifyEvent read FOnPageUnselect write FOnPageUnselect;
  end;

  { TPDFDocumentInfo }

  TPDFDocumentInfo = class(TPersistent)
  private
    FProducer: AnsiString;
    FCreator: AnsiString;
    FAuthor: AnsiString;
    FSubject: AnsiString;
    FTitle: AnsiString;
    FKeywords: AnsiString;
  public
    procedure Assign(Source: TPersistent); override;
  published
    property Producer: AnsiString read FProducer write FProducer;
    property Creator: AnsiString read FCreator write FCreator;
    property Author: AnsiString read FAuthor write FAuthor;
    property Subject: AnsiString read FSubject write FSubject;
    property Title: AnsiString read FTitle write FTitle;
    property Keywords: AnsiString read FKeywords write FKeywords;
  end;

  {$IFDEF PDF_DSPDF}
  { TdsPDF }

  TdsPDF = class(TObject)
  private
    Handle: HMODULE;
    pBeginDoc: function(FileName: PAnsiChar): Integer; stdcall;
    pEndDoc: function: Integer; stdcall;
    pNewPage: function: Integer; stdcall;
    pPrintPageMemory: function(Buffer: Pointer; BufferSize: Integer): Integer; stdcall;
    pPrintPageFile: function(FileName: PAnsiChar): Integer; stdcall;
    pSetParameters: function(OffsetX, OffsetY: Integer; ConverterX, ConverterY: Double): Integer; stdcall;
    pSetPage: function(PageSize, Orientation, Width, Height: Integer): Integer; stdcall;
    pSetDocumentInfo: function(What: Integer; Value: PAnsiChar): Integer; stdcall;
    function PDFPageSizeOf(PaperType: TPaperType): Integer;
  public
    constructor Create; virtual;
    destructor Destroy; override;
    function Exists: Boolean;
    procedure SetDocumentInfoEx(Info: TPDFDocumentInfo);
    function SetDocumentInfo(What: Integer; const Value: AnsiString): Integer;
    function SetPage(PaperType: TPaperType; Orientation: TPrinterOrientation; mmWidth, mmHeight: Integer): Integer;
    function SetParameters(OffsetX, OffsetY: Integer; const ConverterX, ConverterY: Double): Integer;
    function BeginDoc(const FileName: AnsiString): Integer;
    function EndDoc: Integer;
    function NewPage: Integer;
    function RenderMemory(Buffer: Pointer; BufferSize: Integer): Integer;
    function RenderFile(const FileName: AnsiString): Integer;
    function RenderMetaFile(Metafile: TMetafile): Integer;
  end;
  {$ENDIF}

  { GDIPlusSubset }

  TGDIPlusSubset = class(TObject)
  private
    Handle: HMODULE;
    Token: ULONG;
    ThreadToken: ULONG;
    pUnhook: Pointer;
  protected
    GdiplusStartup: function(out Token: ULONG; Input, Output: Pointer): HRESULT; stdcall;
    GdiplusShutdown: procedure(Token: ULONG); stdcall;
    GdipGetDpiX: function(Graphics: Pointer; out Resolution: Single): HRESULT; stdcall;
    GdipGetDpiY: function(Graphics: Pointer; out Resolution: Single): HRESULT; stdcall;
    GdipDrawImageRectRect: function(Graphics, Image: Pointer;
      dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight: Single;
      SrcUnit: Integer; ImageAttributes: Pointer; Callback: Pointer;
      CallbackData: Pointer): HRESULT; stdcall;
    GdipCreateFromHDC: function(hDC: HDC; out Graphics: Pointer): HRESULT; stdcall;
    GdipGetImageGraphicsContext: function(Image: Pointer; out Graphics: Pointer): HRESULT; stdcall;
    GdipDeleteGraphics: function(Graphics: Pointer): HRESULT; stdcall;
    GdipCreateMetafileFromEmf: function(hEMF: HENHMETAFILE; DeleteEMF: BOOL; out Metafile: Pointer): HRESULT; stdcall;
    GdipCreateBitmapFromScan0: function(Width, Height: Integer; Stride: Integer;
      Format: Integer; scan0: PBYTE; out Bitmap: Pointer): HRESULT; stdcall;
    GdipDisposeImage: function(Image: Pointer): HRESULT; stdcall;
    GdipBitmapSetResolution: function(Bitmap: Pointer; dpiX, dpiY: Single): HRESULT; stdcall;
    GdipGetImageHorizontalResolution: function(Image: Pointer; out Resolution: Single): HRESULT; stdcall;
    GdipGetImageVerticalResolution: function(Image: Pointer; out Resolution: Single): HRESULT; stdcall;
    GdipGetImageWidth: function(Image: Pointer; out Width: UINT): HRESULT; stdcall;
    GdipGetImageHeight: function(Image: Pointer; out Height: UINT): HRESULT; stdcall;
    GdipGraphicsClear: function(Graphics: Pointer; Color: UINT): HRESULT; stdcall;
    GdipGetImageEncodersSize: function(out NumEncoders, Size: UINT): HRESULT; stdcall;
    GdipGetImageEncoders: function(NumEncoders, Size: UINT; Encoders: Pointer): HRESULT; stdcall;
    GdipSaveImageToFile: function(Image: Pointer; Filename: PWideChar;
      const clsidEncoder: TGUID; EncoderParams: Pointer): HRESULT; stdcall;
    GdipSaveAddImage: function(Image, NewImage: Pointer; EncoderParams: Pointer): HRESULT; stdcall;
  protected
    function CteateBitmap(Metafile: TMetafile; BackColor: TColor): Pointer;
    function GetEncoderClsid(const MimeType: WideString; out Clsid: TGUID): Boolean;
  public
    constructor Create; virtual;
    destructor Destroy; override;
    function Exists: Boolean;
    procedure Draw(Canvas: TCanvas; const Rect: TRect; Metafile: TMetafile);
    function MultiFrameBegin(const FileName: WideString;
      FirstPage: TMetafile; BackColor: TColor): Pointer;
    procedure MultiFrameNext(MF: Pointer;
      NextPage: TMetafile; BackColor: TColor);
    procedure MultiFrameEnd(MF: Pointer);
  end;

  TPaperSizeInfo = record
    ID: SmallInt;
    Width, Height: Integer;
    Units: TUnits;
  end;

const
  // Paper Sizes
  PaperSizes: array[TPaperType] of TPaperSizeInfo = (
    (ID: DMPAPER_LETTER;                  Width: 08500;     Height: 11000;     Units: mmHiEnglish),
    (ID: DMPAPER_LETTER;                  Width: 08500;     Height: 11000;     Units: mmHiEnglish),
    (ID: DMPAPER_TABLOID;                 Width: 11000;     Height: 17000;     Units: mmHiEnglish),
    (ID: DMPAPER_LEDGER;                  Width: 17000;     Height: 11000;     Units: mmHiEnglish),
    (ID: DMPAPER_LEGAL;                   Width: 08500;     Height: 14000;     Units: mmHiEnglish),
    (ID: DMPAPER_STATEMENT;               Width: 05500;     Height: 08500;     Units: mmHiEnglish),
    (ID: DMPAPER_EXECUTIVE;               Width: 07250;     Height: 10500;     Units: mmHiEnglish),
    (ID: DMPAPER_A3;                      Width: 02970;     Height: 04200;     Units: mmLoMetric),
    (ID: DMPAPER_A4;                      Width: 02100;     Height: 02970;     Units: mmLoMetric),
    (ID: DMPAPER_A4SMALL;                 Width: 02100;     Height: 02970;     Units: mmLoMetric),
    (ID: DMPAPER_A5;                      Width: 01480;     Height: 02100;     Units: mmLoMetric),
    (ID: DMPAPER_B4;                      Width: 02500;     Height: 03540;     Units: mmLoMetric),
    (ID: DMPAPER_B5;                      Width: 01820;     Height: 02570;     Units: mmLoMetric),
    (ID: DMPAPER_FOLIO;                   Width: 08500;     Height: 13000;     Units: mmHiEnglish),
    (ID: DMPAPER_QUARTO;                  Width: 02150;     Height: 02750;     Units: mmLoMetric),
    (ID: DMPAPER_10X14;                   Width: 10000;     Height: 14000;     Units: mmHiEnglish),
    (ID: DMPAPER_11X17;                   Width: 11000;     Height: 17000;     Units: mmHiEnglish),
    (ID: DMPAPER_NOTE;                    Width: 08500;     Height: 11000;     Units: mmHiEnglish),
    (ID: DMPAPER_ENV_9;                   Width: 03875;     Height: 08875;     Units: mmHiEnglish),
    (ID: DMPAPER_ENV_10;                  Width: 04125;     Height: 09500;     Units: mmHiEnglish),
    (ID: DMPAPER_ENV_11;                  Width: 04500;     Height: 10375;     Units: mmHiEnglish),
    (ID: DMPAPER_ENV_12;                  Width: 04750;     Height: 11000;     Units: mmHiEnglish),
    (ID: DMPAPER_ENV_14;                  Width: 05000;     Height: 11500;     Units: mmHiEnglish),
    (ID: DMPAPER_CSHEET;                  Width: 17000;     Height: 22000;     Units: mmHiEnglish),
    (ID: DMPAPER_DSHEET;                  Width: 22000;     Height: 34000;     Units: mmHiEnglish),
    (ID: DMPAPER_ESHEET;                  Width: 34000;     Height: 44000;     Units: mmHiEnglish),
    (ID: DMPAPER_ENV_DL;                  Width: 01100;     Height: 02200;     Units: mmLoMetric),
    (ID: DMPAPER_ENV_C5;                  Width: 01620;     Height: 02290;     Units: mmLoMetric),
    (ID: DMPAPER_ENV_C3;                  Width: 03240;     Height: 04580;     Units: mmLoMetric),
    (ID: DMPAPER_ENV_C4;                  Width: 02290;     Height: 03240;     Units: mmLoMetric),
    (ID: DMPAPER_ENV_C6;                  Width: 01140;     Height: 01620;     Units: mmLoMetric),
    (ID: DMPAPER_ENV_C65;                 Width: 01140;     Height: 02290;     Units: mmLoMetric),
    (ID: DMPAPER_ENV_B4;                  Width: 02500;     Height: 03530;     Units: mmLoMetric),
    (ID: DMPAPER_ENV_B5;                  Width: 01760;     Height: 02500;     Units: mmLoMetric),
    (ID: DMPAPER_ENV_B6;                  Width: 01760;     Height: 01250;     Units: mmLoMetric),
    (ID: DMPAPER_ENV_ITALY;               Width: 01100;     Height: 02300;     Units: mmLoMetric),
    (ID: DMPAPER_ENV_MONARCH;             Width: 03875;     Height: 07500;     Units: mmHiEnglish),
    (ID: DMPAPER_ENV_PERSONAL;            Width: 03625;     Height: 06500;     Units: mmHiEnglish),
    (ID: DMPAPER_FANFOLD_US;              Width: 14875;     Height: 11000;     Units: mmHiEnglish),
    (ID: DMPAPER_FANFOLD_STD_GERMAN;      Width: 08500;     Height: 12000;     Units: mmHiEnglish),
    (ID: DMPAPER_FANFOLD_LGL_GERMAN;      Width: 08500;     Height: 13000;     Units: mmHiEnglish),
    (ID: DMPAPER_ISO_B4;                  Width: 02500;     Height: 03530;     Units: mmLoMetric),
    (ID: DMPAPER_JAPANESE_POSTCARD;       Width: 01000;     Height: 01480;     Units: mmLoMetric),
    (ID: DMPAPER_9X11;                    Width: 09000;     Height: 11000;     Units: mmHiEnglish),
    (ID: DMPAPER_10X11;                   Width: 10000;     Height: 11000;     Units: mmHiEnglish),
    (ID: DMPAPER_15X11;                   Width: 15000;     Height: 11000;     Units: mmHiEnglish),
    (ID: DMPAPER_ENV_INVITE;              Width: 02200;     Height: 02200;     Units: mmLoMetric),
    (ID: DMPAPER_LETTER_EXTRA;            Width: 09500;     Height: 12000;     Units: mmHiEnglish),
    (ID: DMPAPER_LEGAL_EXTRA;             Width: 09500;     Height: 15000;     Units: mmHiEnglish),
    (ID: DMPAPER_TABLOID_EXTRA;           Width: 11690;     Height: 18000;     Units: mmHiEnglish),
    (ID: DMPAPER_A4_EXTRA;                Width: 09270;     Height: 12690;     Units: mmHiEnglish),
    (ID: DMPAPER_LETTER_TRANSVERSE;       Width: 08500;     Height: 11000;     Units: mmHiEnglish),
    (ID: DMPAPER_A4_TRANSVERSE;           Width: 02100;     Height: 02970;     Units: mmLoMetric),
    (ID: DMPAPER_LETTER_EXTRA_TRANSVERSE; Width: 09500;     Height: 12000;     Units: mmHiEnglish),
    (ID: DMPAPER_A_PLUS;                  Width: 02270;     Height: 03560;     Units: mmLoMetric),
    (ID: DMPAPER_B_PLUS;                  Width: 03050;     Height: 04870;     Units: mmLoMetric),
    (ID: DMPAPER_LETTER_PLUS;             Width: 08500;     Height: 12690;     Units: mmHiEnglish),
    (ID: DMPAPER_A4_PLUS;                 Width: 02100;     Height: 03300;     Units: mmLoMetric),
    (ID: DMPAPER_A5_TRANSVERSE;           Width: 01480;     Height: 02100;     Units: mmLoMetric),
    (ID: DMPAPER_B5_TRANSVERSE;           Width: 01820;     Height: 02570;     Units: mmLoMetric),
    (ID: DMPAPER_A3_EXTRA;                Width: 03220;     Height: 04450;     Units: mmLoMetric),
    (ID: DMPAPER_A5_EXTRA;                Width: 01740;     Height: 02350;     Units: mmLoMetric),
    (ID: DMPAPER_B5_EXTRA;                Width: 02010;     Height: 02760;     Units: mmLoMetric),
    (ID: DMPAPER_A2;                      Width: 04200;     Height: 05940;     Units: mmLoMetric),
    (ID: DMPAPER_A3_TRANSVERSE;           Width: 02970;     Height: 04200;     Units: mmLoMetric),
    (ID: DMPAPER_A3_EXTRA_TRANSVERSE;     Width: 03220;     Height: 04450;     Units: mmLoMetric),
    (ID: DMPAPER_USER;                    Width: 0;         Height: 0;         Units: mmPixel));

type
  TPrintPreviewHelper = class(TObject)
  private
    class procedure RaiseOutOfMemory;
    class procedure DrawBitmapAsDIB(DC: HDC; Bitmap: TBitmap; const Rect: TRect);
  public
    class function ConvertUnits(Value, DPI: Integer; InUnits, OutUnits: TUnits): Integer;
    class procedure DrawGraphic(Canvas: TCanvas; X, Y: Integer; Graphic: TGraphic);
    class procedure StretchDrawGraphic(Canvas: TCanvas; const Rect: TRect; Graphic: TGraphic);
    class procedure DrawGrayscale(Canvas: TCanvas; X, Y: Integer; Graphic: TGraphic;
      Brightness: Integer = 0; Contrast: Integer = 0);
    class procedure StretchDrawGrayscale(Canvas: TCanvas; const Rect: TRect; Graphic: TGraphic;
      Brightness: Integer = 0; Contrast: Integer = 0);
    class function CreateWinControlImage(WinControl: TWinControl): TGraphic;
    class procedure ConvertBitmapToGrayscale(Bitmap: TBitmap;
      Brightness: Integer = 0; Contrast: Integer = 0);
    class procedure SmoothDraw(Canvas: TCanvas; const Rect: TRect; Metafile: TMetafile);
    class procedure SwapValues(var A, B: Integer);
    class function ScaleToDeviceContext(DC: HDC; const Pt: TPoint): TPoint;
  end;

{$IFDEF PDF_DSPDF}
function dsPDF: TdsPDF;
{$ENDIF}

{$IFDEF REGISTER}
procedure Register;
{$ENDIF}

implementation

{$R Preview.RES}

uses
  {$IFDEF PDF_SYNOPSE} SynPdf, {$ENDIF}
  {$IFDEF PDF_WPDF} WPPDFR1,WPPDFR2, {$ENDIF}
  System.Types,
  Vcl.ImgList,
  WinApi.RichEdit, WinApi.CommCtrl, System.Math;

resourcestring
  SOutOfMemoryError = 'There is not enough memory to create a new page';
  SLoadError        = 'The content cannot be loaded';
{$IFDEF PDF_DSPDF}
  SdsPDFError       = 'The dsPDF library is not available';
{$ENDIF}
  PDFError          = 'No PDF support';
  SRotated          = 'Rotated';

const
  TextAlignFlags: array[TAlignment] of DWORD = (DT_LEFT, DT_RIGHT, DT_CENTER);
  TextWordWrapFlags: array[Boolean] of DWORD = (DT_END_ELLIPSIS, DT_WORDBREAK);

type
  TStreamHeader = packed record
    Signature: array[0..3] of AnsiChar;
    Version: Word;
  end;

const
  PageInfoHeader: TStreamHeader = (Signature: 'DAPI'; Version: $0550);
  PageListHeader: TStreamHeader = (Signature: 'DAPL'; Version: $0550);

{ Helper Functions }

var _gdiPlus: TGDIPlusSubset = nil;

function gdiPlus: TGDIPlusSubset;
begin
  if not Assigned(_gdiPlus) then
    _gdiPlus := TGDIPlusSubset.Create;
  Result := _gdiPlus;
end;

{$IFDEF PDF_DSPDF}
var _dsPDF: TdsPDF = nil;

function dsPDF: TdsPDF;
begin
  if not Assigned(_dsPDF) then
    _dsPDF := TdsPDF.Create;
  Result := _dsPDF;
end;
{$ENDIF}

procedure TransparentStretchDIBits(dstDC: HDC;
  dstX, dstY: Integer; dstW, dstH: Integer;
  srcX, srcY: Integer; srcW, srcH: Integer;
  bmpBits: Pointer; var bmpInfo: TBitmapInfo;
  mskBits: Pointer; var mskInfo: TBitmapInfo;
  Usage: DWORD);
var
  MemDC: HDC;
  MemBmp: HBITMAP;
  Save: THandle;
  crText, crBack: TColorRef;
  memInfo: pBitmapInfo;
  memBits: Pointer;
  HeaderSize: DWORD;
  ImageSize: DWORD;
begin
  MemDC := CreateCompatibleDC(0);
  try
    MemBmp := CreateCompatibleBitmap(dstDC, srcW, srcH);
    try
      Save := SelectObject(MemDC, MemBmp);
      SetStretchBltMode(MemDC, ColorOnColor);
      StretchDIBits(MemDC, 0, 0, srcW, srcH, 0, 0, srcW, srcH, mskBits, mskInfo, Usage, SrcCopy);
      StretchDIBits(MemDC, 0, 0, srcW, srcH, 0, 0, srcW, srcH, bmpBits, bmpInfo, Usage, SrcErase);
      if Save <> 0 then SelectObject(MemDC, Save);
      GetDIBSizes(MemBmp, HeaderSize, ImageSize);
      GetMem(memInfo, HeaderSize);
      try
        GetMem(memBits, ImageSize);
        try
          GetDIB(MemBmp, 0, memInfo^, memBits^);
          crText := SetTextColor(dstDC, RGB(0, 0, 0));
          crBack := SetBkColor(dstDC, RGB(255, 255, 255));
          SetStretchBltMode(dstDC, ColorOnColor);
          StretchDIBits(dstDC, dstX, dstY, dstW, dstH, srcX, srcY, srcW, srcH, mskBits, mskInfo, Usage, SrcAnd);
          StretchDIBits(dstDC, dstX, dstY, dstW, dstH, srcX, srcY, srcW, srcH, memBits, memInfo^, Usage, SrcInvert);
          SetTextColor(dstDC, crText);
          SetBkColor(dstDC, crBack);
        finally
          FreeMem(memBits, ImageSize);
        end;
      finally
        FreeMem(memInfo, HeaderSize);
      end;
    finally
      DeleteObject(MemBmp);
    end;
  finally
    DeleteDC(MemDC);
  end;
end;

class procedure TPrintPreviewHelper.DrawBitmapAsDIB(DC: HDC; Bitmap: TBitmap; const Rect: TRect);
var
  BitmapHeader: pBitmapInfo;
  BitmapImage: Pointer;
  HeaderSize: DWORD;
  ImageSize: DWORD;
  MaskBitmapHeader: pBitmapInfo;
  MaskBitmapImage: Pointer;
  maskHeaderSize: DWORD;
  MaskImageSize: DWORD;
begin
  GetDIBSizes(Bitmap.Handle, HeaderSize, ImageSize);
  GetMem(BitmapHeader, HeaderSize);
  try
    GetMem(BitmapImage, ImageSize);
    try
      GetDIB(Bitmap.Handle, Bitmap.Palette, BitmapHeader^, BitmapImage^);
      if AllowTransparentDIB and Bitmap.Transparent then
      begin
        GetDIBSizes(Bitmap.MaskHandle, MaskHeaderSize, MaskImageSize);
        GetMem(MaskBitmapHeader, MaskHeaderSize);
        try
          GetMem(MaskBitmapImage, MaskImageSize);
          try
            GetDIB(Bitmap.MaskHandle, 0, MaskBitmapHeader^, MaskBitmapImage^);
            TransparentStretchDIBits(
              DC,                              // handle of destination device context
              Rect.Left, Rect.Top,             // upper-left corner of destination rectagle
              Rect.Right - Rect.Left,          // width of destination rectagle
              Rect.Bottom - Rect.Top,          // height of destination rectagle
              0, 0,                            // upper-left corner of source rectangle
              Bitmap.Width, Bitmap.Height,     // width and height of source rectangle
              BitmapImage,                     // address of bitmap bits
              BitmapHeader^,                   // bitmap data
              MaskBitmapImage,                 // address of mask bitmap bits
              MaskBitmapHeader^,               // mask bitmap data
              DIB_RGB_COLORS                   // usage: the color table contains literal RGB values
            );
          finally
            FreeMem(MaskBitmapImage, MaskImageSize)
          end;
        finally
          FreeMem(MaskBitmapHeader, maskHeaderSize);
        end;
      end
      else
      begin
        SetStretchBltMode(DC, ColorOnColor);
        StretchDIBits(
          DC,                                  // handle of destination device context
          Rect.Left, Rect.Top,                 // upper-left corner of destination rectagle
          Rect.Right - Rect.Left,              // width of destination rectagle
          Rect.Bottom - Rect.Top,              // height of destination rectagle
          0, 0,                                // upper-left corner of source rectangle
          Bitmap.Width, Bitmap.Height,         // width and height of source rectangle
          BitmapImage,                         // address of bitmap bits
          BitmapHeader^,                       // bitmap data
          DIB_RGB_COLORS,                      // usage: the color table contains literal RGB values
          SrcCopy                              // raster operation code: copy source pixels
        );
      end;
    finally
      FreeMem(BitmapImage, ImageSize)
    end;
  finally
    FreeMem(BitmapHeader, HeaderSize);
  end;
end;

class procedure TPrintPreviewHelper.DrawGraphic(Canvas: TCanvas; X, Y: Integer; Graphic: TGraphic);
var
  Rect: TRect;
begin
  Rect.Left := X;
  Rect.Top := Y;
  Rect.Right := X + Graphic.Width;
  Rect.Bottom := Y + Graphic.Height;
  StretchDrawGraphic(Canvas, Rect, Graphic);
end;

class procedure TPrintPreviewHelper.StretchDrawGraphic(Canvas: TCanvas; const Rect: TRect; Graphic: TGraphic);
var
  Bitmap: TBitmap;
begin
  if Graphic is TBitmap then
    DrawBitmapAsDIB(Canvas.Handle, TBitmap(Graphic), Rect)
  else if Graphic is TMetafile then
    TPrintPreviewHelper.SmoothDraw(Canvas, Rect, TMetafile(Graphic))
  else
  begin
    Bitmap := TBitmap.Create;
    try
      Bitmap.Canvas.Brush.Color := clWhite;
      Bitmap.Width := Graphic.Width;
      Bitmap.Height := Graphic.Height;
      Bitmap.Canvas.Draw(0, 0, Graphic);
      Bitmap.Transparent := Graphic.Transparent;
      DrawBitmapAsDIB(Canvas.Handle, Bitmap, Rect)
    finally
      Bitmap.Free;
    end;
  end;
end;

class procedure TPrintPreviewHelper.DrawGrayscale(Canvas: TCanvas; X, Y: Integer; Graphic: TGraphic;
  Brightness, Contrast: Integer);
var
  Rect: TRect;
begin
  Rect.Left := X;
  Rect.Top := Y;
  Rect.Right := X + Graphic.Width;
  Rect.Bottom := Y + Graphic.Height;
  TPrintPreviewHelper.StretchDrawGrayscale(Canvas, Rect, Graphic, Brightness, Contrast);
end;

class procedure TPrintPreviewHelper.StretchDrawGrayscale(Canvas: TCanvas; const Rect: TRect;
  Graphic: TGraphic; Brightness, Contrast: Integer);
var
  Bitmap: TBitmap;
begin
  Bitmap := TBitmap.Create;
  try
    Bitmap.Canvas.Brush.Color := clWhite;
    Bitmap.Width := Graphic.Width;
    Bitmap.Height := Graphic.Height;
    Bitmap.Canvas.Draw(0, 0, Graphic);
    Bitmap.Transparent := Graphic.Transparent;
    TPrintPreviewHelper.ConvertBitmapToGrayscale(Bitmap, Brightness, Contrast);
    DrawBitmapAsDIB(Canvas.Handle, Bitmap, Rect);
  finally
    Bitmap.Free;
  end;
end;

class function TPrintPreviewHelper.CreateWinControlImage(WinControl: TWinControl): TGraphic;
var
  Metafile: TMetafile;
  MetaCanvas: TCanvas;
begin
  Metafile := TMetafile.Create;
  try
    Metafile.Width := WinControl.Width;
    Metafile.Height := WinControl.Height;
    MetaCanvas := TMetafileCanvas.Create(Metafile, 0);
    try
      MetaCanvas.Lock;
      try
        WinControl.PaintTo(MetaCanvas.Handle, 0, 0);
      finally
        MetaCanvas.Unlock;
      end;
    finally
      MetaCanvas.Free;
    end;
  except
    Metafile.Free;
    raise;
  end;
  Result := Metafile;
end;

class procedure TPrintPreviewHelper.ConvertBitmapToGrayscale(Bitmap: TBitmap; Brightness, Contrast: Integer);
// If we consider RGB values in range [0,1] and contrast and brightness in
// rannge [-1,+1], the formula of this function became:
// Gray = Red * 0.30 + Green * 0.59 + Blue * 0.11
// GrayBC = 0.5 + (Gray - 0.5) * (1 + Contrast) + Brighness
// FinalGray = Confine GrayBC in range [0,1]
var
  Pixel: PRGBQuad;
  TransPixel: TRGBQuad;
  X, Y: Integer;
  Gray: Integer;
  Offset: Integer;
  Scale: Integer;
begin
  Bitmap.PixelFormat := pf32bit;
  TransPixel.rgbRed := GetRValue(Bitmap.TransparentColor);
  TransPixel.rgbGreen := GetGValue(Bitmap.TransparentColor);
  TransPixel.rgbBlue := GetBValue(Bitmap.TransparentColor);
  if Bitmap.Transparent then
    TransPixel.rgbReserved := 0
  else
    TransPixel.rgbReserved := 255;
  Scale := 100 + Contrast;
  Offset := 128 + (255 * Brightness - 128 * Scale) div 100;
  Pixel := Bitmap.ScanLine[Bitmap.Height - 1];
  for Y := 0 to Bitmap.Height - 1 do
  begin
    for X := 0 to Bitmap.Width - 1 do
    begin
      if PDWORD(Pixel)^ <> PDWORD(@TransPixel)^ then
        with Pixel^ do
        begin
          Gray := Offset + (rgbRed * 30 + rgbGreen * 59 + rgbBlue * 11) * Scale div 10000;
          if Gray > 255 then
            Gray := 255
          else if Gray < 0 then
            Gray := 0;
          rgbRed := Gray;
          rgbGreen := Gray;
          rgbBlue := Gray;
        end;
      Inc(Pixel);
    end;
  end;
end;

class function TPrintPreviewHelper.ScaleToDeviceContext(DC: HDC; const Pt: TPoint): TPoint;
var
  Handle: HDC;
begin
  Handle := DC;
  if DC = 0 then
    Handle := GetDC(0);
  try
    Result.X := Round(Pt.X * GetDeviceCaps(Handle, HORZRES) / GetDeviceCaps(Handle, DESKTOPHORZRES));
    Result.Y := Round(Pt.Y * GetDeviceCaps(Handle, VERTRES) / GetDeviceCaps(Handle, DESKTOPVERTRES));
  finally
    if DC = 0 then
      ReleaseDC(0, Handle);
  end;
end;

class procedure TPrintPreviewHelper.SmoothDraw(Canvas: TCanvas; const Rect: TRect; Metafile: TMetafile);
begin
  gdiPlus.Draw(Canvas, Rect, Metafile);
end;

{ TTemporaryFileStream }

constructor TTemporaryFileStream.Create;
// Delphi 2009 bug: do not use Unicode string here!
var
  TempPath: array[0..MAX_PATH] of AnsiChar;
  TempFile: array[0..MAX_PATH] of AnsiChar;
begin
  GetTempPathA(SizeOf(TempPath), TempPath);
  GetTempFileNameA(TempPath, 'DA', 0, TempFile);
  inherited Create(CreateFileA(TempFile, GENERIC_READ or GENERIC_WRITE, 0, nil,
    CREATE_ALWAYS, FILE_ATTRIBUTE_TEMPORARY or FILE_FLAG_RANDOM_ACCESS or
    FILE_FLAG_DELETE_ON_CLOSE, 0));
end;

destructor TTemporaryFileStream.Destroy;
begin
  FileClose(Handle);
  inherited Destroy;
end;

{ TIntegerList }

function TIntegerList.GetItems(Index: Integer): Integer;
begin
  Result := Integer(Get(Index));
end;

procedure TIntegerList.SetItems(Index: Integer; Value: Integer);
begin
  Put(Index, Pointer(Value));
end;

function TIntegerList.Add(Value: Integer): Integer;
begin
  Result := inherited Add(Pointer(Value));
end;

procedure TIntegerList.Insert(Index, Value: Integer);
begin
  inherited Insert(Index, Pointer(Value));
end;

function TIntegerList.Remove(Value: Integer): Integer;
begin
  Result := inherited Remove(Pointer(Value));
end;

function TIntegerList.Extract(Value: Integer): Integer;
begin
  Result := Integer(inherited Extract(Pointer(Value)));
end;

function TIntegerList.IndexOf(Value: Integer): Integer;
begin
  Result := inherited IndexOf(Pointer(Value));
end;

function TIntegerList.First: Integer;
begin
  Result := Integer(inherited First);
end;

function TIntegerList.Last: Integer;
begin
  Result := Integer(inherited Last);
end;

function IntegerCompare(Item1, Item2: Pointer): Integer;
begin
  Result := Integer(Item1) - Integer(Item2);
end;

procedure TIntegerList.Sort;
begin
  inherited Sort(IntegerCompare);
end;

procedure TIntegerList.LoadFromStream(Stream: TStream);
var
  V, I: Integer;
begin
  Clear;
  Stream.ReadBuffer(V, SizeOf(V));
  Count := V;
  for I := 0 to Count - 1 do
  begin
    Stream.ReadBuffer(V, SizeOf(V));
    Items[I] := V;
  end;
end;

procedure TIntegerList.SaveToStream(Stream: TStream);
var
  V, I: Integer;
begin
  V := Count;
  Stream.WriteBuffer(V, SizeOf(V));
  for I := 0 to Count - 1 do
  begin
    V := Items[I];
    Stream.WriteBuffer(V, SizeOf(V));
  end;
end;

{ TMetafileEntry }

constructor TMetafileEntry.Create(AOwner: TMetafileList);
begin
  FOwner := AOwner;
  FMetafile := TMetafile.Create;
  FMetafile.OnChange := MetafileChanged;
  FStates := [msInMemory];
end;

constructor TMetafileEntry.CreateInMemory(AOwner: TMetafileList;
  AMetafile: TMetafile);
begin
  FOwner := AOwner;
  FMetafile := TMetafile.Create;
  FMetafile.Assign(AMetafile);
  FMetafile.OnChange := MetafileChanged;
  FStates := [msInMemory, msDirty];
end;

constructor TMetafileEntry.CreateInStorage(AOwner: TMetafileList;
  const AOffset, ASize: Int64);
begin
  FOwner := AOwner;
  FOffset := AOffset;
  FSize := ASize;
  FStates := [msInStorage];
end;

destructor TMetafileEntry.Destroy;
begin
  if FMetafile <> nil then
    FMetafile.Free;
  inherited Destroy;
end;

procedure TMetafileEntry.MetafileChanged(Sender: TObject);
var
  CanNotifyIt: Boolean;
begin
  CanNotifyIt := (FSize <> 0) or (msDirty in FStates);
  Include(FStates, msDirty);
  if CanNotifyIt then
    FOwner.EntryChanged(Self);
end;

procedure TMetafileEntry.CopyToMemory;
begin
  if (msInStorage in FStates) and not (msInMemory in FStates) then
  begin
    FOwner.Storage.Seek(FOffset, soBeginning);
    FMetafile := TMetafile.Create;
    FMetafile.LoadFromStream(FOwner.Storage);
    FMetafile.OnChange := MetafileChanged;
    Include(FStates, msInMemory);
    TouchCount := 0;
  end;
end;

procedure TMetafileEntry.CopyToStorage;
begin
  if msDirty in FStates then
  begin
    if (msInStorage in FStates) and (FOffset + FSize = FOwner.Storage.Size) then
    begin
      FOwner.Storage.Seek(FOffset, soBeginning);
      FMetafile.SaveToStream(FOwner.Storage);
      FSize := FOwner.Storage.Position - FOffset;
      if msInStorage in FStates then
        FOwner.Storage.Size := FOwner.Storage.Position;
    end
    else
    begin
      FOffset := FOwner.Storage.Seek(0, soEnd);
      FMetafile.SaveToStream(FOwner.Storage);
      FSize := FOwner.Storage.Position - FOffset;
    end;
    Include(FStates, msInStorage);
    Exclude(FStates, msInMemory);
    Exclude(FStates, msDirty);
    FMetafile.Free;
    FMetafile := nil;
  end;
end;

function TMetafileEntry.IsMoreRequiredThan(Another: TMetafileEntry): Boolean;
begin
  Result := Self.TouchCount > Another.TouchCount;
end;

procedure TMetafileEntry.Touch;
begin
  Inc(TouchCount);
end;

{ TMetafileList }

constructor TMetafileList.Create;
begin
  inherited Create;
  FEntries := TList.Create;
  FCachedEntries := TList.Create;
  FCacheSize := 10;
end;

destructor TMetafileList.Destroy;
begin
  Reset;
  FCachedEntries.Free;
  FEntries.Free;
  inherited Destroy;
end;

function TMetafileList.GetCount: Integer;
begin
  Result := FEntries.Count;
end;

function TMetafileList.GetItems(Index: Integer): TMetafileEntry;
begin
  Result := GetCachedEntry(Index);
end;

function TMetafileList.GetMetafiles(Index: Integer): TMetafile;
begin
  Result := Items[Index].Metafile;
end;

procedure TMetafileList.SetCacheSize(Value: Integer);
begin
  if Value < 1 then
    Value := 1;
  if FCacheSize <> Value then
  begin
    FCacheSize := Value;
    if FCachedEntries.Count > FCacheSize then
      ReduceCacheEntries(FCacheSize);
  end;
end;

procedure TMetafileList.ReduceCacheEntries(NumOfEntries: Integer);
var
  I: Integer;
  LessRequiredIndex: Integer;
  LessRequired: TMetafileEntry;
  Entry: TMetafileEntry;
begin
  while FCachedEntries.Count > NumOfEntries do
  begin
    LessRequiredIndex := FCachedEntries.Count - 1;
    LessRequired := TMetafileEntry(FCachedEntries[LessRequiredIndex]);
    for I := LessRequiredIndex - 1 downto 0 do
    begin
      Entry := TMetafileEntry(FCachedEntries[I]);
      if LessRequired.IsMoreRequiredThan(Entry) then
      begin
        LessRequired := Entry;
        LessRequiredIndex := I;
      end;
    end;
    if msDirty in LessRequired.States then
    begin
      if FStorage = nil then
        FStorage := TTemporaryFileStream.Create;
      LessRequired.CopyToStorage;
    end;
    FCachedEntries.Delete(LessRequiredIndex);
  end;
end;

function TMetafileList.GetCachedEntry(Index: Integer): TMetafileEntry;
begin
  Result := TMetafileEntry(FEntries[Index]);
  if not (msInMemory in Result.States) then
  begin
    if FCachedEntries.Count >= FCacheSize then
      ReduceCacheEntries(FCacheSize - 1);
    Result.CopyToMemory;
    FCachedEntries.Add(Result);
  end;
  Result.Touch;
end;

procedure TMetafileList.Reset;
var
  I: Integer;
begin
  FCachedEntries.Clear;
  for I := FEntries.Count - 1 downto 0 do
    TMetafileEntry(FEntries[I]).Free;
  FEntries.Clear;
  if Assigned(FStorage) then
  begin
    FStorage.Free;
    FStorage := nil;
  end;
end;

procedure TMetafileList.EntryChanged(Entry: TMetafileEntry);
var
  Index: Integer;
begin
  Index := FEntries.IndexOf(Entry);
  if Index >= 0 then
    DoSingleChange(Index);
end;

procedure TMetafileList.DoSingleChange(Index: Integer);
begin
  if Assigned(FOnSingleChange) then
    FOnSingleChange(Self, Index);
end;

procedure TMetafileList.DoMultipleChange(StartIndex, EndIndex: Integer);
begin
  if Assigned(FOnMultipleChange) then
    FOnMultipleChange(Self, StartIndex, EndIndex);
end;

procedure TMetafileList.Clear;
begin
  if FEntries.Count > 0 then
  begin
    Reset;
    DoMultipleChange(0, -1);
  end;
end;

function TMetafileList.Add(AMetafile: TMetafile): Integer;
begin
  Result := FEntries.Count;
  Insert(Result, AMetafile);
end;

procedure TMetafileList.Insert(Index: Integer; AMetafile: TMetafile);
var
  Entry: TMetafileEntry;
begin
  if Index < 0 then
    Index := 0
  else if Index > FEntries.Count then
    Index := FEntries.Count;
  ReduceCacheEntries(FCacheSize - 1);
  Entry := TMetafileEntry.CreateInMemory(Self, AMetafile);
  FEntries.Insert(Index, Entry);
  FCachedEntries.Add(Entry);
  DoMultipleChange(Index, Count - 1);
end;

procedure TMetafileList.Delete(Index: Integer);
var
  Entry: TMetafileEntry;
begin
  if (FEntries.Count = 1) and (Index = 0) then
    Clear
  else
  begin
    Entry := TMetafileEntry(FEntries[Index]);
    if msInMemory in Entry.States then
      FCachedEntries.Remove(Entry);
    if (msInStorage in Entry.States) and(Entry.Offset + Entry.Size = FStorage.Size) then
      FStorage.Size := Entry.Offset;
    FEntries.Delete(Index);
    Entry.Free;
    DoMultipleChange(Index, Count - 1);
  end;
end;

procedure TMetafileList.Exchange(Index1, Index2: Integer);
begin
  if Index1 <> Index2 then
  begin
    FEntries.Exchange(Index1, Index2);
    if Index1 < Index2 then
      DoMultipleChange(Index1, Index2)
    else
      DoMultipleChange(Index2, Index1);
  end;
end;

procedure TMetafileList.Move(Index, NewIndex: Integer);
begin
  if Index <> NewIndex then
  begin
    FEntries.Move(Index, NewIndex);
    if Assigned(FOnMultipleChange) then
    begin
      if Index < NewIndex then
        DoMultipleChange(Index, NewIndex)
      else
        DoMultipleChange(NewIndex, Index);
    end;
  end;
end;

function TMetafileList.LoadFromStream(Stream: TStream): Boolean;
var
  Header: TStreamHeader;
  Offsets: TIntegerList;
  Entry: TMetafileEntry;
  Size, Offset: Int64;
  DataSize: DWORD;
  I: Integer;
begin
  Result := False;
  Stream.ReadBuffer(Header, SizeOf(Header));
  if CompareMem(@Header.Signature, @PageListHeader.Signature, SizeOf(Header.Signature)) then
  begin
    Clear;
    Offsets := TIntegerList.Create;
    try
      Stream.ReadBuffer(DataSize, SizeOf(DataSize));
      Offsets.LoadFromStream(Stream);
      if Offsets.Count <= CacheSize then
      begin
        for I := 0 to Offsets.Count - 1 do
        begin
          Entry := TMetafileEntry.Create(Self);
          Entry.Metafile.LoadFromStream(Stream);
          FEntries.Add(Entry);
          FCachedEntries.Add(Entry);
        end;
      end
      else
      begin
        FStorage := TTemporaryFileStream.Create;
        FStorage.CopyFrom(Stream, DataSize);
        Offset := 0;
        for I := 0 to Offsets.Count - 1 do
        begin
          if I < Offsets.Count - 1 then
            Size := DWORD(Offsets[I + 1]) - Offset
          else
            Size := FStorage.Size - Offset;
          Entry := TMetafileEntry.CreateInStorage(Self, Offset, Size);
          FEntries.Add(Entry);
          Inc(Offset, Size);
        end;
      end;
    finally
      Offsets.Free;
    end;
    if FEntries.Count > 0 then
      DoMultipleChange(0, Count - 1);
    Result := True;
  end;
end;

procedure TMetafileList.SaveToStream(Stream: TStream);
var
  Offsets: TIntegerList;
  Entry: TMetafileEntry;
  HeaderOffset: Int64;
  BaseOffset: Int64;
  DataSize: DWORD;
  I: Integer;
begin
  Stream.WriteBuffer(PageListHeader, SizeOf(PageListHeader));
  HeaderOffset := Stream.Position;
  Stream.WriteBuffer(DataSize, SizeOf(DataSize));
  Offsets := TIntegerList.Create;
  try
    Offsets.Count := FEntries.Count;
    Offsets.SaveToStream(Stream);
    BaseOffset := Stream.Position;
    for I := 0 to FEntries.Count - 1 do
    begin
      Offsets[I] := DWORD(Stream.Position - BaseOffset);
      Entry := TMetafileEntry(FEntries[I]);
      if (msInStorage in Entry.States) and not (msDirty in Entry.States) then
      begin
        FStorage.Seek(Entry.Offset, soBeginning);
        Stream.CopyFrom(FStorage, Entry.Size);
      end
      else
        Entry.Metafile.SaveToStream(Stream);
    end;
    DataSize := DWORD(Stream.Position - BaseOffset);
    Stream.Seek(HeaderOffset, soBeginning);
    Stream.WriteBuffer(DataSize, SizeOf(DataSize));
    Offsets.SaveToStream(Stream);
    Stream.Seek(DataSize, soCurrent);
  finally
    Offsets.Free;
  end;
end;

procedure TMetafileList.LoadFromFile(const FileName: String);
var
  FileStream: TFileStream;
begin
  FileStream := TFileStream.Create(FileName, fmOpenRead or fmShareDenyWrite);
  try
    LoadFromStream(FileStream);
  finally
    FileStream.Free;
  end;
end;

procedure TMetafileList.SaveToFile(const FileName: String);
var
  FileStream: TFileStream;
begin
  FileStream := TFileStream.Create(FileName, fmCreate or fmShareExclusive);
  try
    SaveToStream(FileStream);
  finally
    FileStream.Free;
  end;
end;

{ TPaperPreviewOptions }

constructor TPaperPreviewOptions.Create;
begin
  inherited Create;
  FBorderColor := clBlack;
  FBorderWidth := 1;
  FCursor := crDefault;
  FDragCursor := crHand;
  FGrabCursor := crGrab;  //pvg
  FPaperColor := clWhite;
  FShadowColor := clBtnShadow;
  FShadowWidth := 3;
end;

procedure TPaperPreviewOptions.Assign(Source: TPersistent);
begin
  if Source is TPaperPreviewOptions then
  begin
    BorderColor := TPaperPreviewOptions(Source).BorderColor;
    BorderWidth :=  TPaperPreviewOptions(Source).BorderWidth;
    ShadowColor := TPaperPreviewOptions(Source).ShadowColor;
    ShadowWidth := TPaperPreviewOptions(Source).ShadowWidth;
    Cursor := TPaperPreviewOptions(Source).Cursor;
    DragCursor := TPaperPreviewOptions(Source).DragCursor;
    GrabCursor := TPaperPreviewOptions(Source).GrabCursor; //pvg
    Hint := TPaperPreviewOptions(Source).Hint;
    PaperColor := TPaperPreviewOptions(Source).PaperColor;
    PopupMenu := TPaperPreviewOptions(Source).PopupMenu;
  end
  else
    inherited Assign(Source);
end;

procedure TPaperPreviewOptions.AssignTo(Dest: TPersistent);
begin
  if Dest is TPaperPreviewOptions then
    Dest.Assign(Self)
  else if Dest is TPaperPreview then
  begin
    TPaperPreview(Dest).PaperColor := PaperColor;
    TPaperPreview(Dest).BorderColor := BorderColor;
    TPaperPreview(Dest).BorderWidth := BorderWidth;
    TPaperPreview(Dest).ShadowColor := ShadowColor;
    TPaperPreview(Dest).ShadowWidth := ShadowWidth;
    TPaperPreview(Dest).Cursor := Cursor;
    TPaperPreview(Dest).PopupMenu := PopupMenu;
    TPaperPreview(Dest).Hint := Hint;
  end
  else
    inherited AssignTo(Dest);
end;

procedure TPaperPreviewOptions.DoChange(Severity: TUpdateSeverity);
begin
  if Assigned(FOnChange) then
    FOnChange(self, Severity);
end;

procedure TPaperPreviewOptions.SetPaperColor(Value: TColor);
begin
  if PaperColor <> Value then
  begin
    FPaperColor := Value;
    DoChange(usRedraw);
  end;
end;

procedure TPaperPreviewOptions.SetBorderColor(Value: TColor);
begin
  if BorderColor <> Value then
  begin
    FBorderColor := Value;
    DoChange(usRedraw);
  end;
end;

procedure TPaperPreviewOptions.SetBorderWidth(Value: TBorderWidth);
begin
  if BorderWidth <> Value then
  begin
    FBorderWidth := Value;
    DoChange(usRecreate);
  end;
end;

procedure TPaperPreviewOptions.SetShadowColor(Value: TColor);
begin
  if ShadowColor <> Value then
  begin
    FShadowColor := Value;
    DoChange(usRedraw);
  end;
end;

procedure TPaperPreviewOptions.SetShadowWidth(Value: TBorderWidth);
begin
  if ShadowWidth <> Value then
  begin
    FShadowWidth := Value;
    DoChange(usRecreate);
  end;
end;

procedure TPaperPreviewOptions.SetCursor(Value: TCursor);
begin
  if Cursor <> Value then
  begin
    FCursor := Value;
    DoChange(usNone);
  end;
end;

procedure TPaperPreviewOptions.SetDragCursor(Value: TCursor);
begin
  if DragCursor <> Value then
  begin
    FDragCursor := Value;
    DoChange(usNone);
  end;
end;

procedure TPaperPreviewOptions.SetGrabCursor(Value: TCursor); //pvg
begin
  if GrabCursor <> Value then
  begin
    FGrabCursor := Value;
    DoChange(usNone);
  end;
end;

procedure TPaperPreviewOptions.SetHint(const Value: String);
begin
  if Hint <> Value then
  begin
    FHint := Value;
    DoChange(usNone);
  end;
end;

procedure TPaperPreviewOptions.SetPopupMenu(Value: TPopupMenu);
begin
  if PopupMenu <> Value then
  begin
    FPopupMenu := Value;
    DoChange(usNone);
  end;
end;

procedure TPaperPreviewOptions.CalcDimensions(PaperWidth, PaperHeight: Integer;
  out PaperRect, BoxRect: TRect);
begin
  PaperRect.Left := BorderWidth;
  PaperRect.Right := PaperRect.Left + PaperWidth;
  PaperRect.Top := BorderWidth;
  PaperRect.Bottom := PaperRect.Top + PaperHeight;
  BoxRect.Left := 0;
  BoxRect.Top := 0;
  BoxRect.Right := BorderWidth + PaperWidth + BorderWidth + ShadowWidth;
  BoxRect.Bottom := BorderWidth + PaperHeight + BorderWidth + ShadowWidth;
end;

procedure TPaperPreviewOptions.Draw(Canvas: TCanvas; const BoxRect: TRect);
var
  R: TRect;
begin
  if ShadowWidth > 0 then
  begin
    R.Left := BoxRect.Right - ShadowWidth;
    R.Right := BoxRect.Right;
    R.Top := 0;
    R.Bottom := ShadowWidth;
    Canvas.FillRect(R);
    R.Left := 0;
    R.Right := ShadowWidth;
    R.Top := BoxRect.Bottom - ShadowWidth;
    R.Bottom := BoxRect.Bottom;
    Canvas.FillRect(R);
    Canvas.Brush.Color := ShadowColor;
    Canvas.Brush.Style := bsSolid;
    R.Left := BoxRect.Right - ShadowWidth;
    R.Right := BoxRect.Right;
    R.Top := BoxRect.Top + ShadowWidth;
    R.Bottom := BoxRect.Bottom;
    Canvas.FillRect(R);
    R.Left := ShadowWidth;
    R.Top := BoxRect.Bottom - ShadowWidth;
    Canvas.FillRect(R);
  end;
  if BorderWidth > 0 then
  begin
    Canvas.Pen.Width := BorderWidth;
    Canvas.Pen.Style := psInsideFrame;
    Canvas.Pen.Color := BorderColor;
    Canvas.Brush.Style := bsClear;
    Canvas.Rectangle(BoxRect.Left, BoxRect.Top,
      BoxRect.Right - ShadowWidth, BoxRect.Bottom - ShadowWidth);
  end;
end;

{ TPaperPreview }

constructor TPaperPreview.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  ControlStyle := ControlStyle + [csOpaque, csDisplayDragImage];
  OffScreen := TBitmap.Create;
  PageCanvas := TCanvas.Create;
  FPreservePaperSize := True;
  FBorderColor := clBlack;
  FBorderWidth := 1;
  FPaperColor := clWhite;
  FShadowColor := clBtnShadow;
  FShadowWidth := 3;
  FShowCaption := False;
  FAlignment := taCenter;
  FWordWrap := True;
  Width := 100;
  Height := 150;
end;

destructor TPaperPreview.Destroy;
begin
  OffScreen.Free;
  PageCanvas.Free;
  inherited Destroy;
end;

procedure TPaperPreview.Invalidate;
begin
  IsOffScreenReady := False;
  if WindowHandle <> 0 then
    InvalidateRect(WindowHandle, @PageRect, False);
end;

procedure TPaperPreview.InvalidateAll;
begin
  IsOffScreenPrepared := False;
  if WindowHandle <> 0 then
    InvalidateRect(WindowHandle, nil, False);
end;

procedure TPaperPreview.Paint;
var
  OffDC: HDC;
  VisibleRect: TRect;
  VisiblePageRect: TRect;
  SavedDC: Integer;
begin
  if IntersectRect(VisibleRect, Canvas.ClipRect, ClientRect) then
  begin
    if not IsOffScreenPrepared or
      (VisibleRect.Left < LastVisibleRect.Left) or
      (VisibleRect.Top < LastVisibleRect.Top) or
      (VisibleRect.Right > LastVisibleRect.Right) or
      (VisibleRect.Bottom > LastVisibleRect.Bottom) then
    begin
      OffScreen.Width := VisibleRect.Right - VisibleRect.Left;
      OffScreen.Height := VisibleRect.Bottom - VisibleRect.Top;
      OffDC := OffScreen.Canvas.Handle;
      SetWindowOrgEx(OffDC, VisibleRect.Left, VisibleRect.Top, nil);
      DrawPage(OffScreen.Canvas);
      SetWindowOrgEx(OffDC, 0, 0, nil);
      LastVisibleRect := VisibleRect;
      IsOffScreenPrepared := True;
      IsOffScreenReady := False;
    end;
    if IntersectRect(VisiblePageRect, VisibleRect, PageRect) then
    begin
      if not IsOffScreenReady or
        (VisiblePageRect.Left < LastVisiblePageRect.Left) or
        (VisiblePageRect.Top < LastVisiblePageRect.Top) or
        (VisiblePageRect.Right > LastVisiblePageRect.Right) or
        (VisiblePageRect.Bottom > LastVisiblePageRect.Bottom) then
      begin
        OffDC := OffScreen.Canvas.Handle;
        SelectClipRgn(OffDC, 0);
        SetWindowOrgEx(OffDC, LastVisibleRect.Left, LastVisibleRect.Top, nil);
        with VisiblePageRect do
          IntersectClipRect(OffDC, Left, Top, Right, Bottom);
        with OffScreen.Canvas do
        begin
          Brush.Color := PaperColor;
          Brush.Style := bsSolid;
          FillRect(VisiblePageRect);
        end;
        if Assigned(FOnPaint) then
        begin
          SavedDC := SaveDC(OffDC);
          PageCanvas.Handle := OffDC;
          try
            FOnPaint(Self, PageCanvas, PageRect);
          finally
            PageCanvas.Handle := 0;
            RestoreDC(OffDC, SavedDC);
          end;
        end;
        SetWindowOrgEx(OffDC, 0, 0, nil);
        LastVisiblePageRect := VisiblePageRect;
        IsOffScreenReady := True;
      end;
    end;
    Canvas.Draw(LastVisibleRect.Left, LastVisibleRect.Top, OffScreen);
  end;
end;

procedure TPaperPreview.DrawPage(Canvas: TCanvas);
var
  Rect: TRect;
  Flags: DWORD;
begin
  Canvas.Pen.Mode := pmCopy;
  if ShowCaption and (Caption <> '') then
  begin
    Rect.Left := 0;
    Rect.Top := Height - CaptionHeight;
    Rect.Right := Width - ShadowWidth + 1;
    Rect.Bottom := Height;
    if RectVisible(Canvas.Handle, Rect) then
    begin
      Canvas.Brush.Color := Color;
      Canvas.Brush.Style := bsSolid;
      Canvas.Font.Assign(Font);
      Canvas.FillRect(Rect);
      InflateRect(Rect, 0, -1);
      Flags := TextAlignFlags[Alignment] or TextWordWrapFlags[WordWrap]
            or DT_NOPREFIX;
      Flags := DrawTextBiDiModeFlags(Flags);
      DrawText(Canvas.Handle, PChar(Caption), Length(Caption), Rect, Flags);
    end;
  end;
  if ShadowWidth > 0 then
  begin
    Canvas.Brush.Color := Color;
    Canvas.Brush.Style := bsSolid;
    Rect.Left := Width - ShadowWidth;
    Rect.Right := Width;
    Rect.Top := 0;
    Rect.Bottom := ShadowWidth;
    Canvas.FillRect(Rect);
    Rect.Left := 0;
    Rect.Right := ShadowWidth;
    Rect.Top := Height - CaptionHeight - ShadowWidth;
    Rect.Bottom := Height - CaptionHeight;
    Canvas.FillRect(Rect);
    Canvas.Brush.Color := ShadowColor;
    Canvas.Brush.Style := bsSolid;
    Rect.Left := Width - ShadowWidth;
    Rect.Top := ShadowWidth;
    Rect.Right := Width;
    Rect.Bottom := Height - CaptionHeight;
    Canvas.FillRect(Rect);
    Rect.Left := ShadowWidth;
    Rect.Top := Height - CaptionHeight - ShadowWidth;
    Canvas.FillRect(Rect);
  end;
  if BorderWidth > 0 then
  begin
    Canvas.Pen.Width := BorderWidth;
    Canvas.Pen.Style := psInsideFrame;
    Canvas.Pen.Color := BorderColor;
    Canvas.Brush.Style := bsClear;
    Canvas.Rectangle(0, 0, Width - ShadowWidth, Height - CaptionHeight - ShadowWidth);
  end;
end;

procedure TPaperPreview.UpdateCaptionHeight;
var
  Rect: TRect;
  Flags: DWORD;
  NewCaptionHeight: Integer;
  SavedSize: TPoint;
  DC: HDC;
begin
  if ShowCaption then
  begin
    Flags := TextAlignFlags[Alignment] or TextWordWrapFlags[WordWrap]
          or DT_NOPREFIX or DT_CALCRECT;
    Flags := DrawTextBiDiModeFlags(Flags);
    Rect.Left := 0;
    Rect.Right := Width - ShadowWidth;
    Rect.Top := 0;
    Rect.Bottom := 0;
    Dec(Rect.Right, ShadowWidth);
    if HandleAllocated then
      DC := Canvas.Handle
    else
    begin
      DC := CreateCompatibleDC(0);
      SelectObject(DC, Font.Handle);
    end;
    DrawText(DC, PChar(Caption), Length(Caption), Rect, Flags);
    if HandleAllocated then
      DeleteDC(DC);
    NewCaptionHeight := Rect.Bottom - Rect.Top + 2;
  end
  else
    NewCaptionHeight := 0;
  if CaptionHeight <> NewCaptionHeight then
  begin
    SavedSize := PaperSize;
    FCaptionHeight := NewCaptionHeight;
    if PreservePaperSize then
      PaperSize := SavedSize
    else
      InvalidateAll;
  end;
end;

function TPaperPreview.ClientToPaper(const Pt: TPoint): TPoint;
begin
  Result.X := Pt.X - BorderWidth;
  Result.Y := Pt.Y - BorderWidth;
end;

function TPaperPreview.PaperToClient(const Pt: TPoint): TPoint;
begin
  Result.X := Pt.X + BorderWidth;
  Result.Y := Pt.Y + BorderWidth;
end;

procedure TPaperPreview.SetBoundsEx(ALeft, ATop, APaperWidth, APaperHeight: Integer);
begin
  FPageRect.Left := BorderWidth;
  FPageRect.Top := BorderWidth;
  FPageRect.Right := FPageRect.Left + APaperWidth;
  FPageRect.Bottom := FPageRect.Top + APaperHeight;
  SetBounds(ALeft, ATop, ActualWidth(APaperWidth), ActualHeight(APaperHeight));
end;

function TPaperPreview.ActualWidth(Value: Integer): Integer;
begin
  Result := Value + 2 * FBorderWidth + FShadowWidth;
end;

function TPaperPreview.ActualHeight(Value: Integer): Integer;
begin
  Result := Value + 2 * FBorderWidth + FShadowWidth + CaptionHeight;
end;

function TPaperPreview.LogicalWidth(Value: Integer): Integer;
begin
  Result := Value - 2 * FBorderWidth - FShadowWidth;
end;

function TPaperPreview.LogicalHeight(Value: Integer): Integer;
begin
  Result := Value - 2 * FBorderWidth - FShadowWidth - CaptionHeight;
end;

procedure TPaperPreview.SetPaperWidth(Value: Integer);
begin
  ClientWidth := ActualWidth(Value);
end;

function TPaperPreview.GetPaperWidth: Integer;
begin
  Result := LogicalWidth(Width);
end;

procedure TPaperPreview.SetPaperHeight(Value: Integer);
begin
  ClientHeight := ActualHeight(Value);
end;

function TPaperPreview.GetPaperHeight: Integer;
begin
  Result := LogicalHeight(ClientHeight);
end;

procedure TPaperPreview.SetPaperSize(const Value: TPoint);
begin
  SetBoundsEx(Left, Top, Value.X, Value.Y);
end;

function TPaperPreview.GetPaperSize: TPoint;
begin
  Result.X := LogicalWidth(Width);
  Result.Y := LogicalHeight(Height);
end;

procedure TPaperPreview.SetPaperColor(Value: TColor);
begin
  if PaperColor <> Value then
  begin
    FPaperColor := Value;
    InvalidateAll;
  end;
end;

procedure TPaperPreview.SetBorderColor(Value: TColor);
begin
  if BorderColor <> Value then
  begin
    FBorderColor := Value;
    InvalidateAll;
  end;
end;

procedure TPaperPreview.SetBorderWidth(Value: TBorderWidth);
var
  SavedSize: TPoint;
begin
  if BorderWidth <> Value then
  begin
    SavedSize := PaperSize;
    FBorderWidth := Value;
    if PreservePaperSize then
      PaperSize := SavedSize
    else
      InvalidateAll;
  end;
end;

procedure TPaperPreview.SetShadowColor(Value: TColor);
begin
  if ShadowColor <> Value then
  begin
    FShadowColor := Value;
    InvalidateAll;
  end;
end;

procedure TPaperPreview.SetShadowWidth(Value: TBorderWidth);
var
  SavedSize: TPoint;
begin
  if ShadowWidth <> Value then
  begin
    SavedSize := PaperSize;
    FShadowWidth := Value;
    if PreservePaperSize then
      PaperSize := SavedSize
    else
      InvalidateAll;
  end;
end;

procedure TPaperPreview.SetShowCaption(Value: Boolean);
begin
  if ShowCaption <> Value then
  begin
    FShowCaption := Value;
    UpdateCaptionHeight;
  end;
end;

procedure TPaperPreview.SetAlignment(Value: TAlignment);
begin
  if Alignment <> Value then
  begin
    FAlignment := Value;
    if ShowCaption then
      InvalidateAll;
  end;
end;

procedure TPaperPreview.SetWordWrap(Value: Boolean);
begin
  if WordWrap <> Value then
  begin
    FWordWrap := Value;
    if ShowCaption then
      UpdateCaptionHeight;
  end;
end;

procedure TPaperPreview.WMSize(var Message: TWMSize);
begin
  inherited;
  FPageRect.Left := BorderWidth;
  FPageRect.Top := BorderWidth;
  FPageRect.Right := FPageRect.Left + LogicalWidth(Width);
  FPageRect.Bottom := FPageRect.Top + LogicalHeight(Height);
  InvalidateAll;
  if Assigned(OnResize) then
    OnResize(Self);
end;

procedure TPaperPreview.WMEraseBkgnd(var Message: TWMEraseBkgnd);
begin
  Message.Result := 1;
end;

procedure TPaperPreview.CMMouseEnter(var Message: TMessage);
begin
  inherited;
  if Assigned(FOnMouseEnter) then
    FOnMouseEnter(Self);
end;

procedure TPaperPreview.CMMouseLeave(var Message: TMessage);
begin
  inherited;
  if Assigned(FOnMouseLeave) then
    FOnMouseLeave(Self);
end;

procedure TPaperPreview.CMColorChanged(var Message: TMessage);
begin
  inherited;
  InvalidateAll;
end;

procedure TPaperPreview.CMFontChanged(var Message: TMessage);
begin
  inherited;
  if ShowCaption then
  begin
    UpdateCaptionHeight;
    InvalidateAll;
  end;
end;

procedure TPaperPreview.CMTextChanged(var Message: TMessage);
begin
  inherited;
  if ShowCaption then
  begin
    UpdateCaptionHeight;
    InvalidateAll;
  end;
end;

procedure TPaperPreview.BiDiModeChanged(var Message: TMessage);
begin
  inherited;
  if ShowCaption then
    InvalidateAll
  else
    Invalidate;
end;

{ TPDFDocumentInfo }

procedure TPDFDocumentInfo.Assign(Source: TPersistent);
begin
  if Source is TPDFDocumentInfo then
  begin
    Producer := TPDFDocumentInfo(Source).Producer;
    Creator := TPDFDocumentInfo(Source).Creator;
    Author := TPDFDocumentInfo(Source).Author;
    Subject := TPDFDocumentInfo(Source).Subject;
    Title := TPDFDocumentInfo(Source).Title;
    Keywords := TPDFDocumentInfo(Source).Keywords;
  end
  else
    inherited Assign(Source);
end;

{ TPrintPreviewHelper }

class procedure TPrintPreviewHelper.RaiseOutOfMemory;
begin
  raise EOutOfMemory.Create(SOutOfMemoryError);
end;

class procedure TPrintPreviewHelper.SwapValues(var A, B: Integer);
var
  T: Integer;
begin
  T := A;
  A := B;
  B := T;
end;

class function TPrintPreviewHelper.ConvertUnits(Value, DPI: Integer; InUnits, OutUnits: TUnits): Integer;
begin
  Result := Value;
  case InUnits of
    mmLoMetric:
      case OutUnits of
        mmLoMetric: Result := Value;
        mmHiMetric: Result := Value * 10;
        mmLoEnglish: Result := Round(Value * 100 / 254);
        mmHiEnglish: Result := Round(Value * 1000 / 254);
        mmPoints: Result := Round(Value * 72 / 254);
        mmTWIPS: Result := Round(Value * 1440 / 254);
        mmPixel: Result := Round(Value * DPI / 254);
      end;
    mmHiMetric:
      case OutUnits of
        mmLoMetric: Result := Value div 10;
        mmHiMetric: Result := Value;
        mmLoEnglish: Result := Round(Value * 100 / 2540);
        mmHiEnglish: Result := Round(Value * 1000 / 2540);
        mmPoints: Result := Round(Value * 72 / 2540);
        mmTWIPS: Result := Round(Value * 1440 / 2540);
        mmPixel: Result := Round(Value * DPI / 2540);
      end;
    mmLoEnglish:
      case OutUnits of
        mmLoMetric: Result := Round(Value * 254 / 100);
        mmHiMetric: Result := Round(Value * 2540 / 100);
        mmLoEnglish: Result := Value;
        mmHiEnglish: Result := Value * 10;
        mmPoints: Result := Round(Value * 72 / 100);
        mmTWIPS: Result := Round(Value * 1440 / 100);
        mmPixel: Result := Round(Value * DPI / 100);
      end;
    mmHiEnglish:
      case OutUnits of
        mmLoMetric: Result := Round(Value * 254 / 1000);
        mmHiMetric: Result := Round(Value * 2540 / 1000);
        mmLoEnglish: Result := Value div 10;
        mmHiEnglish: Result := Value;
        mmPoints: Result := Round(Value * 72 / 1000);
        mmTWIPS: Result := Round(Value * 1440 / 1000);
        mmPixel: Result := Round(Value * DPI / 1000);
      end;
    mmPoints:
      case OutUnits of
        mmLoMetric: Result := Round(Value * 254 / 72);
        mmHiMetric: Result := Round(Value * 2540 / 72);
        mmLoEnglish: Result := Round(Value * 100 / 72);
        mmHiEnglish: Result := Round(Value * 1000 / 72);
        mmPoints: Result := Value;
        mmTWIPS: Result := Value * 20;
        mmPixel: Result := Round(Value * DPI / 72);
      end;
    mmTWIPS:
      case OutUnits of
        mmLoMetric: Result := Round(Value * 254 / 1440);
        mmHiMetric: Result := Round(Value * 2540 / 1440);
        mmLoEnglish: Result := Round(Value * 100 / 1440);
        mmHiEnglish: Result := Round(Value * 1000 / 1440);
        mmPoints: Result := Value div 20;
        mmTWIPS: Result := Value;
        mmPixel: Result := Round(Value * DPI / 1440);
      end;
    mmPixel:
      case OutUnits of
        mmLoMetric: Result := Round(Value * 254 / DPI);
        mmHiMetric: Result := Round(Value * 2540 / DPI);
        mmLoEnglish: Result := Round(Value * 100 / DPI);
        mmHiEnglish: Result := Round(Value * 1000 / DPI);
        mmPoints: Result := Round(Value * 72 / DPI);
        mmTWIPS: Result := Round(Value * 1440 / DPI);
        mmPixel: Result := Value;
      end;
  end;
end;

{ TPrintPreview }

constructor TPrintPreview.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  ControlStyle := ControlStyle - [csAcceptsControls] + [csDisplayDragImage];
  VertScrollBar.Visible := true;
  HorzScrollBar.Visible := true;
  VertScrollBar.Tracking := true;
  HorzScrollBar.Tracking := true;
  Align := alClient;
  TabStop := True;
  ParentFont := False;
  Font.Name := 'Arial';
  Font.Size := 8;
  FPrintableAreaColor := clSilver;
  FPDFDocumentInfo := TPDFDocumentInfo.Create;
  FPageList := TMetafileList.Create;
  FPageList.OnMultipleChange := PagesChanged;
  FPageList.OnSingleChange := PageChanged;
  FPaperViewOptions := TPaperPreviewOptions.Create;
  FPaperViewOptions.OnChange := PaperViewOptionsChanged;
  FPaperView := TPaperPreview.Create(Self);
  with FPaperView do
  begin
    Parent := Self;
    TabStop := False;
    Visible := False;
    OnPaint := PaintPage;
    OnClick := PaperClick;
    OnDblClick := PaperDblClick;
    OnMouseDown := PaperMouseDown;
    OnMouseMove := PaperMouseMove;
    OnMouseUp := PaperMouseUp;
  end;
  FPaperViewOptions.AssignTo(FPaperView);
  FState := psReady;
  FZoom := 100;
  FZoomMin := 10;
  FZoomMax := 1000;
  FZoomStep := 10;
  FZoomSavePos := True;
  FZoomState := zsZoomToFit;
  FUnits := mmHiMetric;
  FOrientation := poPortrait;
  SetPaperType(pA4);
  UpdateExtends;
end;

destructor TPrintPreview.Destroy;
begin
  FPageList.Free;
  FPaperView.Free;
  FPaperViewOptions.Free;
  FPDFDocumentInfo.Free;
  if Assigned(AnnotationMetafile) then
    AnnotationMetafile.Free;
  if Assigned(BackgroundMetafile) then
    BackgroundMetafile.Free;
  if Assigned(FThumbnailViews) then
  begin
    FThumbnailViews.Free;
    FThumbnailViews := nil;
  end;
  inherited Destroy;
end;

procedure TPrintPreview.Loaded;
begin
  inherited Loaded;
  UpdateExtends;
  UpdateZoom;
end;

function TPrintPreview.ConvertX(X: Integer; InUnits, OutUnits: TUnits): Integer;
begin
  Result := TPrintPreviewHelper.ConvertUnits(X, HorzPixelsPerInch, InUnits, OutUnits);
end;

function TPrintPreview.ConvertY(Y: Integer; InUnits, OutUnits: TUnits): Integer;
begin
  Result := TPrintPreviewHelper.ConvertUnits(Y, VertPixelsPerInch, InUnits, OutUnits);
end;

function TPrintPreview.ConvertXY(X, Y: Integer; InUnits, OutUnits: TUnits): TPoint;
begin
  Result.X := TPrintPreviewHelper.ConvertUnits(X, HorzPixelsPerInch, InUnits, OutUnits);
  Result.Y := TPrintPreviewHelper.ConvertUnits(Y, VertPixelsPerInch, InUnits, OutUnits);
end;

procedure TPrintPreview.ConvertPoints(var Points; NumPoints: Integer;
  InUnits, OutUnits: TUnits);
var
  pPoints: PPoint;
begin
  pPoints := @Points;
  while NumPoints > 0 do
  begin
    with pPoints^ do
    begin
      X := TPrintPreviewHelper.ConvertUnits(X, HorzPixelsPerInch, InUnits, OutUnits);
      Y := TPrintPreviewHelper.ConvertUnits(Y, VertPixelsPerInch, InUnits, OutUnits);
    end;
    Inc(pPoints);
    Dec(NumPoints);
  end;
end;

function TPrintPreview.ConvertRect(Rec: TRect; InUnits,
  OutUnits: TUnits): TRect;
begin
  Result.Left := TPrintPreviewHelper.ConvertUnits(Rec.Left, HorzPixelsPerInch, InUnits, OutUnits);
  Result.Top := TPrintPreviewHelper.ConvertUnits(Rec.Top, VertPixelsPerInch, InUnits, OutUnits);
  Result.Right := TPrintPreviewHelper.ConvertUnits(Rec.Right, HorzPixelsPerInch, InUnits, OutUnits);
  Result.Bottom := TPrintPreviewHelper.ConvertUnits(Rec.Bottom, VertPixelsPerInch, InUnits, OutUnits);
  Result.TopLeft.X := Result.Left;
  Result.TopLeft.Y := Result.Top;
  Result.BottomRight.X := Result.Right;
  Result.BottomRight.Y := Result.Bottom;
end;

function TPrintPreview.BoundsFrom(AUnits: TUnits;
  ALeft, ATop, AWidth, AHeight: Integer): TRect;
begin
  Result := RectFrom(AUnits, ALeft, ATop, ALeft + AWidth, ATop + AHeight);
end;

function TPrintPreview.RectFrom(AUnits: TUnits;
  ALeft, ATop, ARight, ABottom: Integer): TRect;
begin
  Result.TopLeft := PointFrom(AUnits, ALeft, ATop);
  Result.BottomRight := PointFrom(AUnits, ARight, ABottom);
end;

function TPrintPreview.PointFrom(AUnits: TUnits; X, Y: Integer): TPoint;
begin
  Result := ConvertXY(X, Y, AUnits, FUnits);
end;

function TPrintPreview.XFrom(AUnits: TUnits; X: Integer): Integer;
begin
  Result := ConvertX(X, AUnits, FUnits);
end;

function TPrintPreview.YFrom(AUnits: TUnits; Y: Integer): Integer;
begin
  Result := ConvertY(Y, AUnits, FUnits);
end;

function TPrintPreview.ScreenToPreview(X, Y: Integer): TPoint;
begin
  Result.X := ConvertX(MulDiv(X, HorzPixelsPerInch, Screen.PixelsPerInch), mmPixel, FUnits);
  Result.Y := ConvertY(MulDiv(Y, VertPixelsPerInch, Screen.PixelsPerInch), mmPixel, FUnits);
end;

function TPrintPreview.PreviewToScreen(X, Y: Integer): TPoint;
begin
  Result.X := MulDiv(ConvertX(X, FUnits, mmPixel), Screen.PixelsPerInch, HorzPixelsPerInch);
  Result.Y := MulDiv(ConvertY(Y, FUnits, mmPixel), Screen.PixelsPerInch, VertPixelsPerInch);
end;

function TPrintPreview.ScreenToPaper(const Pt: TPoint): TPoint;
begin
  Result := FPaperView.ScreenToClient(Pt);
  Result := FPaperView.ClientToPaper(Result);
  Result.X := MulDiv(Result.X, 100, FZoom);
  Result.Y := MulDiv(Result.Y, 100, FZoom);
  Result := ScreenToPreview(Result.X, Result.Y);
end;

function TPrintPreview.PaperToScreen(const Pt: TPoint): TPoint;
begin
  Result := PreviewToScreen(Pt.X, Pt.Y);
  Result.X := MulDiv(Result.X, FZoom, 100);
  Result.Y := MulDiv(Result.Y, FZoom, 100);
  Result := FPaperView.PaperToClient(Result);
  Result := FPaperView.ClientToScreen(Result);
end;

function TPrintPreview.ClientToPaper(const Pt: TPoint): TPoint;
begin
  Result := ScreenToPaper(ClientToScreen(Pt));
end;

function TPrintPreview.PaperToClient(const Pt: TPoint): TPoint;
begin
  Result := ScreenToClient(PaperToScreen(Pt));
end;

function TPrintPreview.PaintGraphic(X, Y: Integer; Graphic: TGraphic): TPoint;
var
  Rect: TRect;
begin
  Result := ScreenToPreview(Graphic.Width, Graphic.Height);
  Rect.Left := X;
  Rect.Right := X + Result.X;
  Rect.Top := Y;
  Rect.Bottom := Y + Result.Y;
  TPrintPreviewHelper.StretchDrawGraphic(Canvas, Rect, Graphic);
end;

function TPrintPreview.PaintGraphicEx(const Rect: TRect; Graphic: TGraphic;
  Proportinal, ShrinkOnly, Center: Boolean): TRect;
var
  gW, gH: Integer;
  rW, rH: Integer;
  W, H: Integer;
begin
  with ScreenToPreview(Graphic.Width, Graphic.Height) do
  begin
    gW := X;
    gH := Y;
  end;
  rW := Rect.Right - Rect.Left;
  rH := Rect.Bottom - Rect.Top;
  if not ShrinkOnly or (gW > rW) or (gH > rH) then
  begin
    if Proportinal then
    begin
      if (rW / gW) < (rH / gH) then
      begin
        H := MulDiv(gH, rW, gW);
        W := rW;
      end
      else
      begin
        W := MulDiv(gW, rH, gH);
        H := rH;
      end;
    end
    else
    begin
      W := rW;
      H := rH;
    end;
  end
  else
  begin
    W := gW;
    H := gH;
  end;
  if Center then
  begin
    Result.Left := Rect.Left + (rW - W) div 2;
    Result.Top := Rect.Top + (rH - H) div 2;
  end
  else
    Result.TopLeft := Rect.TopLeft;
  Result.Right := Result.Left + W;
  Result.Bottom := Result.Top + H;
  TPrintPreviewHelper.StretchDrawGraphic(Canvas, Result, Graphic);
end;

//rmk
function TPrintPreview.PaintGraphicEx2(const Rect: TRect; Graphic: TGraphic;
  VertAlign: TVertAlign; HorzAlign: THorzAlign): TRect;
var
  gW, gH: Integer;
  rW, rH: Integer;
  W, H: Integer;
begin
  with ScreenToPreview(Graphic.Width, Graphic.Height) do
  begin
    gW := X;
    gH := Y;
  end;
  rW := Rect.Right - Rect.Left;
  rH := Rect.Bottom - Rect.Top;

  if (gW > rW) or (gH > rH) then
  begin
    if (rW / gW) < (rH / gH) then
    begin
      H := MulDiv(gH, rW, gW);
      W := rW;
    end
    else
    begin
      W := MulDiv(gW, rH, gH);
      H := rH;
    end;
  end
  else
  begin
    W := gW;
    H := gH;
  end;

  Case VertAlign of
    vaTop   : Result.Top := Rect.Top;
    vaCenter: Result.Top := Rect.Top + (rH - H) div 2;
    vaBottom: Result.Top := Rect.Bottom - H;
  else
    Result.Top := Rect.Top + (rH - H) div 2;
  end;

  Case HorzAlign of
    haLeft  : Result.Left := Rect.Left;
    haCenter: Result.Left := Rect.Left + (rW - W) div 2;
    haRight : Result.Left := Rect.Right - W;
  else
    Result.Left := Rect.Left + (rW - W) div 2;
  end;

  Result.Right := Result.Left + W;
  Result.Bottom := Result.Top + H;

  TPrintPreviewHelper.StretchDrawGraphic(Canvas, Result, Graphic);
end;

function TPrintPreview.PaintWinControl(X, Y: Integer;
  WinControl: TWinControl): TPoint;
var
  Graphic: TGraphic;
begin
  Graphic := TPrintPreviewHelper.CreateWinControlImage(WinControl);
  try
    PaintGraphic(X, Y, Graphic);
  finally
    Graphic.Free;
  end;
end;

function TPrintPreview.PaintWinControlEx(const Rect: TRect;
  WinControl: TWinControl; Proportinal, ShrinkOnly, Center: Boolean): TRect;
var
  Graphic: TGraphic;
begin
  Graphic := TPrintPreviewHelper.CreateWinControlImage(WinControl);
  try
    PaintGraphicEx(Rect, Graphic, Proportinal, ShrinkOnly, Center);
  finally
    Graphic.Free;
  end;
end;

function TPrintPreview.PaintWinControlEx2(const Rect: TRect;
  WinControl: TWinControl; VertAlign: TVertAlign; HorzAlign: THorzAlign): TRect;
var
  Graphic: TGraphic;
begin
  Graphic := TPrintPreviewHelper.CreateWinControlImage(WinControl);
  try
    PaintGraphicEx2(Rect, Graphic, VertAlign, HorzAlign);
  finally
    Graphic.Free;
  end;
end;

function TPrintPreview.PaintRichText(const Rect: TRect;
  RichEdit: TCustomRichEdit; MaxPages: Integer; pOffset: PInteger): Integer;
var
  Range: TFormatRange;
  RectTWIPS: TRect;
  SaveIndex: Integer;
  MaxLen: Integer;
  TextLenEx: TGetTextLengthEx;
begin
  Result := 0;
  RectTWIPS := Rect;
  ConvertPoints(RectTWIPS, 2, FUnits, mmTWIPS);
  FillChar(Range, SizeOf(TFormatRange), 0);
  if pOffset = nil then
    Range.chrg.cpMin := 0
  else
    Range.chrg.cpMin := pOffset^;
  TextLenEx.flags := GTL_DEFAULT;
  TextLenEx.codepage := CP_UTF8;
  MaxLen := SendMessage(RichEdit.Handle, EM_GETTEXTLENGTHEX, WPARAM(@TextLenEx), 0);
  SaveIndex := SaveDC(FPageCanvas.Handle);
  try
    SendMessage(RichEdit.Handle, EM_FORMATRANGE, 0, 0);
    repeat
      if Result > 0  then
      begin
        RestoreDC(FPageCanvas.Handle, SaveIndex);
        NewPage;
        SaveIndex := SaveDC(FPageCanvas.Handle);
      end;
      Range.chrg.cpMax := -1;
      Range.rc := RectTWIPS;
      Range.rcPage := RectTWIPS;
      Range.hdc := FPageCanvas.Handle;
      SetMapMode(FPageCanvas.Handle, MM_TEXT);
      Range.chrg.cpMin := SendMessage(RichEdit.Handle, EM_FORMATRANGE, 0, LPARAM(@Range));
      SendMessage(RichEdit.Handle, EM_DISPLAYBAND, 0, LPARAM(@Range.rc));
      if Range.chrg.cpMin <> -1 then
        Inc(Result);
    until (Range.chrg.cpMin >= MaxLen) or (Range.chrg.cpMin = -1) or
          ((MaxPages > 0) and (Result >= MaxPages));
  finally
    SendMessage(RichEdit.Handle, EM_FORMATRANGE, 0, 0);
    RestoreDC(FPageCanvas.Handle, SaveIndex);
  end;
  if pOffset <> nil then
    if Range.chrg.cpMin < MaxLen then
      pOffset^ := Range.chrg.cpMin
    else
      pOffset^ := -1;
end;

function TPrintPreview.GetRichTextRect(var Rect: TRect;
  RichEdit: TCustomRichEdit; pOffset: PInteger): Integer;
var
  Range: TFormatRange;
  RectTWIPS: TRect;
  SaveIndex: Integer;
  MaxLen: Integer;
  TextLenEx: TGetTextLengthEx;
begin
  RectTWIPS := Rect;
  ConvertPoints(RectTWIPS, 2, FUnits, mmTWIPS);
  FillChar(Range, SizeOf(TFormatRange), 0);
  Range.rc := RectTWIPS;
  Range.rcPage := RectTWIPS;
  Range.hdc := FPageCanvas.Handle;
  Range.chrg.cpMax := -1;
  if pOffset = nil then
    Range.chrg.cpMin := 0
  else
    Range.chrg.cpMin := pOffset^;
  SaveIndex := SaveDC(FPageCanvas.Handle);
  try
    SetMapMode(FPageCanvas.Handle, MM_TEXT);
    SendMessage(RichEdit.Handle, EM_FORMATRANGE, 0, 0);
    Range.chrg.cpMin := SendMessage(RichEdit.Handle, EM_FORMATRANGE, 0, LPARAM(@Range));
    if Range.chrg.cpMin = -1 then
      Rect.Bottom := Rect.Top
    else
      Rect.Bottom := ConvertY(Range.rc.Bottom, mmTWIPS, FUnits);
  finally
    SendMessage(RichEdit.Handle, EM_FORMATRANGE, 0, 0);
    RestoreDC(FPageCanvas.Handle, SaveIndex);
  end;
  if pOffset <> nil then
  begin
    TextLenEx.flags := GTL_DEFAULT;
    TextLenEx.codepage := CP_UTF8;
    MaxLen := SendMessage(RichEdit.Handle, EM_GETTEXTLENGTHEX, WPARAM(@TextLenEx), 0);
    if Range.chrg.cpMin < MaxLen then
      pOffset^ := Range.chrg.cpMin
    else
      pOffset^ := -1;
  end;
  Result := Rect.Bottom;
end;

procedure TPrintPreview.WMEraseBkgnd(var Message: TWMEraseBkgnd);
begin
  Message.Result := 1
end;

procedure TPrintPreview.WMPaint(var Message: TWMPaint);
var
  DC: HDC;
  PaintStruct: TPaintStruct;
begin
  DC := Message.DC;
  if Message.DC = 0 then
    DC := BeginPaint(WindowHandle, PaintStruct);
  try
    if FPaperView.Visible then
      with FPaperView.BoundsRect do
        ExcludeClipRect(DC, Left, Top, Right, Bottom);
    FillRect(DC, PaintStruct.rcPaint, Brush.Handle);
  finally
    if Message.DC = 0 then
      EndPaint(WindowHandle, PaintStruct);
  end;
end;

procedure TPrintPreview.CNKeyDown(var Message: TWMKey);
var
  Key: Word;
  Shift: TShiftState;
begin
  with Message do
  begin
    Key := CharCode;
    Shift := KeyDataToShiftState(KeyData);
  end;
  if (Key = VK_HOME) and (Shift = []) then
    Perform(WM_HSCROLL, SB_LEFT, 0)
  else if (Key = VK_HOME) and (Shift = [ssCtrl]) then
    Perform(WM_VSCROLL, SB_TOP, 0)
  else if (Key = VK_END) and (Shift = []) then
    Perform(WM_HSCROLL, SB_RIGHT, 0)
  else if (Key = VK_END) and (Shift = [ssCtrl]) then
    Perform(WM_VSCROLL, SB_BOTTOM, 0)
  else if (Key = VK_LEFT) and (Shift = [ssShift]) then
    Perform(WM_HSCROLL, MakeLong(SB_THUMBPOSITION, HorzScrollBar.Position - 1), 0)
  else if (Key = VK_LEFT) and (Shift = []) then
    Perform(WM_HSCROLL, SB_LINELEFT, 0)
  else if (Key = VK_LEFT) and (Shift = [ssCtrl]) then
    Perform(WM_HSCROLL, SB_PAGELEFT, 0)
  else if (Key = VK_RIGHT) and (Shift = [ssShift]) then
    Perform(WM_HSCROLL, MakeLong(SB_THUMBPOSITION, HorzScrollBar.Position + 1), 0)
  else if (Key = VK_RIGHT) and (Shift = []) then
    Perform(WM_HSCROLL, SB_LINERIGHT, 0)
  else if (Key = VK_RIGHT) and (Shift = [ssCtrl]) then
    Perform(WM_HSCROLL, SB_PAGERIGHT, 0)
  else if (Key = VK_UP) and (Shift = [ssShift]) then
    Perform(WM_VSCROLL, MakeLong(SB_THUMBPOSITION, VertScrollBar.Position - 1), 0)
  else if (Key = VK_UP) and (Shift = []) then
    Perform(WM_VSCROLL, SB_LINEUP, 0)
  else if (Key = VK_UP) and (Shift = [ssCtrl]) then
    Perform(WM_VSCROLL, SB_PAGEUP, 0)
  else if (Key = VK_DOWN) and (Shift = [ssShift]) then
    Perform(WM_VSCROLL, MakeLong(SB_THUMBPOSITION, VertScrollBar.Position + 1), 0)
  else if (Key = VK_DOWN) and (Shift = []) then
    Perform(WM_VSCROLL, SB_LINEDOWN, 0)
  else if (Key = VK_DOWN) and (Shift = [ssCtrl]) then
    Perform(WM_VSCROLL, SB_PAGEDOWN, 0)
  else if (Key = VK_NEXT) and (Shift = []) then
    CurrentPage := CurrentPage + 1
  else if (Key = VK_NEXT) and (Shift = [ssCtrl]) then
    CurrentPage := TotalPages
  else if (Key = VK_PRIOR) and (Shift = []) then
    CurrentPage := CurrentPage - 1
  else if (Key = VK_PRIOR) and (Shift = [ssCtrl]) then
    CurrentPage := 1
  else if (Key = VK_ADD) and (Shift = []) then
    Zoom := Zoom + ZoomStep
  else if (Key = VK_SUBTRACT) and (Shift = []) then
    Zoom := Zoom - ZoomStep
  else
    inherited;
end;

procedure TPrintPreview.WMMouseWheel(var Message: TMessage);
var
  Amount: Integer;
  ScrollDir: Integer;
  Shift: TShiftState;
  I: Integer;
begin
  if PtInRect(ClientRect, ScreenToClient(Mouse.CursorPos)) then
  begin
    Message.Result := 0;
    Inc(WheelAccumulator, SmallInt(Message.WParamHi));
    Amount := WheelAccumulator div WHEEL_DELTA;
    if Amount <> 0 then
    begin
      WheelAccumulator := WheelAccumulator mod WHEEL_DELTA;
      Shift := KeyboardStateToShiftState;
      if Shift = [] then
      begin
        ScrollDir := SB_LINEUP;
        if Amount < 0 then
        begin
          ScrollDir := SB_LINEDOWN;
          Amount := -Amount;
        end;
        for I := 1 to Amount do
          Perform(WM_VSCROLL, ScrollDir, 0);
      end
      else if Shift = [ssCtrl] then
        Zoom := Zoom + ZoomStep * Amount
      else if (Shift = [ssShift]) or (Shift = [ssMiddle]) then
        CurrentPage := CurrentPage + Amount;
    end;
  end;
end;

procedure TPrintPreview.WMHScroll(var Message: TWMScroll);
begin
  inherited;
  Update;
  SyncThumbnail;
end;

procedure TPrintPreview.WMVScroll(var Message: TWMScroll);
begin
  inherited;
  Update;
  SyncThumbnail;
end;

procedure TPrintPreview.PaperClick(Sender: TObject);
begin
  Click;
end;

procedure TPrintPreview.PaperDblClick(Sender: TObject);
begin
  DblClick;
end;

procedure TPrintPreview.PaperMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  Pt: TPoint;
begin
  if not Focused and Enabled then SetFocus;
  //pvg begin
  if (Sender = FPaperView) and (FCanScrollHorz or FCanScrollVert) then
  begin
    FIsDragging := True;
    FPaperView.Cursor := FPaperViewOptions.GrabCursor;
    FPaperView.Perform(WM_SETCURSOR, FPaperView.Handle, HTCLIENT);
  end;
  //pvg end
  Pt.X := X;
  Pt.Y := Y;
  FOldMousePos := Pt;
  MapWindowPoints(FPaperView.Handle, Handle, Pt, 1);
  MouseDown(Button, Shift, Pt.X, Pt.Y);
end;

procedure TPrintPreview.PaperMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
var
  Delta: TPoint;
  Pt: TPoint;
begin
  Pt.X := X;
  Pt.Y := Y;
  MapWindowPoints(FPaperView.Handle, Handle, Pt, 1);
  MouseMove(Shift, Pt.X, Pt.Y);
  if ssLeft in Shift then
  begin
    if FCanScrollHorz then
    begin
      Delta.X := X - FOldMousePos.X;
      if not (AutoScroll and HorzScrollBar.Visible) then
      begin
        if FPaperView.Left + Delta.X < ClientWidth - HorzScrollBar.Margin - FPaperView.Width then
          Delta.X := ClientWidth - HorzScrollBar.Margin - FPaperView.Width - FPaperView.Left
        else if FPaperView.Left + Delta.X > HorzScrollBar.Margin then
          Delta.X := HorzScrollBar.Margin - FPaperView.Left;
        FPaperView.Left := FPaperView.Left + Delta.X;
      end
      else
        HorzScrollBar.Position := HorzScrollBar.Position - Delta.X;
    end;
    if FCanScrollVert then
    begin
      Delta.Y := Y - FOldMousePos.Y;
      if not (AutoScroll and VertScrollBar.Visible) then
      begin
        if FPaperView.Top + Delta.Y < ClientHeight - VertScrollBar.Margin - FPaperView.Height then
          Delta.Y := ClientHeight - VertScrollBar.Margin - FPaperView.Height - FPaperView.Top
        else if FPaperView.Top + Delta.Y > VertScrollBar.Margin then
          Delta.Y := VertScrollBar.Margin - FPaperView.Top;
        FPaperView.Top := FPaperView.Top + Delta.Y;
      end
      else
        VertScrollBar.Position := VertScrollBar.Position - Delta.Y;
    end;
    if (FCanScrollHorz and (Delta.X <> 0)) or (FCanScrollVert and (Delta.Y <> 0)) then
    begin
      Update;
      SyncThumbnail;
    end;
  end;
end;

procedure TPrintPreview.PaperMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  Pt: TPoint;
begin
  Pt.X := X;
  Pt.Y := Y;
  MapWindowPoints(FPaperView.Handle, Handle, Pt, 1);
  MouseUp(Button, Shift, Pt.X, Pt.Y);
  //pvg begin
  if FIsDragging then
  begin
    FIsDragging := False;
    FPaperView.Cursor := FPaperViewOptions.DragCursor;
  end;
  //pvg end
end;

function TPrintPreview.GetSystemDefaultUnits: TUnits;
var
  Data: array[0..1] of Char;
begin
  GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_IMEASURE, Data, 2);
  if Data[0] = '0' then
    Result := mmHiMetric
  else
    Result := mmHiEnglish;
end;

function TPrintPreview.GetUserDefaultUnits: TUnits;
var
  Data: array[0..1] of Char;
begin
  GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IMEASURE, Data, 2);
  if Data[0] = '0' then
    Result := mmHiMetric
  else
    Result := mmHiEnglish;
end;

procedure TPrintPreview.SetPageSetupParameters(PageSetupDialog: TPageSetupDialog);
var
  OutUnit: TUnits;
begin
  case PageSetupDialog.Units of
    pmMillimeters: OutUnit := mmHiMetric;
    pmInches: OutUnit := mmHiEnglish;
  else
    OutUnit := UserDefaultUnits;
  end;
  if Printer.Orientation = Orientation then
  begin
    PageSetupDialog.PageWidth := ConvertX(PaperWidth, FUnits, OutUnit);
    PageSetupDialog.PageHeight := ConvertY(PaperHeight, FUnits, OutUnit);
  end
  else
  begin
    Printer.Orientation := Orientation;
    PageSetupDialog.PageWidth := ConvertX(PaperHeight, FUnits, OutUnit);
    PageSetupDialog.PageHeight := ConvertY(PaperWidth, FUnits, OutUnit);
  end
end;

function TPrintPreview.GetPageSetupParameters(PageSetupDialog: TPageSetupDialog): TRect;
var
  InUnit: TUnits;
  NewWidth, NewHeight: Integer;
begin
  case PageSetupDialog.Units of
    pmMillimeters: InUnit := mmHiMetric;
    pmInches: InUnit := mmHiEnglish;
  else
    InUnit := UserDefaultUnits;
  end;
  NewWidth := ConvertX(PageSetupDialog.PageWidth, InUnit, FUnits);
  NewHeight := ConvertY(PageSetupDialog.PageHeight, InUnit, FUnits);
  SetPaperSizeOrientation(NewWidth, NewHeight, Printer.Orientation);
  Result := PageBounds;
  Inc(Result.Left, ConvertX(PageSetupDialog.MarginLeft, InUnit, FUnits));
  Inc(Result.Top, ConvertY(PageSetupDialog.MarginTop, InUnit, FUnits));
  Dec(Result.Right, ConvertX(PageSetupDialog.MarginRight, InUnit, FUnits));
  Dec(Result.Bottom, ConvertX(PageSetupDialog.MarginBottom, InUnit, FUnits));
end;

procedure TPrintPreview.SetPrinterOptions;
var
  DeviceMode: THandle;
  DevMode: PDeviceMode;
  Device, Driver, Port: array[0..MAX_PATH] of Char;
  DriverInfo2: PDriverInfo2;
  DriverInfo2Size: DWORD;
  hPrinter: THandle;
  PaperSize: TPoint;
begin
  if PrinterInstalled then
  begin
    Printer.GetPrinter(Device, Driver, Port, DeviceMode);
    DevMode := PDevMode(GlobalLock(DeviceMode));
    try
      with DevMode^ do
      begin
        dmFields := dmFields and not
          (DM_FORMNAME or DM_PAPERSIZE or DM_PAPERWIDTH or DM_PAPERLENGTH);
        if not IsDummyFormName then
        begin
          dmFields := dmFields or DM_FORMNAME;
          StrPLCopy(dmFormName, FormName, CCHFORMNAME);
        end;
        if PaperType = pCustom then
        begin
          PaperSize := ConvertXY(PaperWidth, PaperHeight, Units, mmLoMetric);
          if FOrientation = poLandscape then
            TPrintPreviewHelper.SwapValues(PaperSize.X, PaperSize.Y);
          dmFields := dmFields or DM_PAPERSIZE;
          dmPaperSize := DMPAPER_USER;
          dmFields := dmFields or DM_PAPERWIDTH;
          dmPaperWidth := PaperSize.X;
          dmFields := dmFields or DM_PAPERLENGTH;
          dmPaperLength := PaperSize.Y;
        end
        else
        begin
          dmFields := dmFields or DM_PAPERSIZE;
          dmPaperSize := PaperSizes[PaperType].ID;
        end;
        dmFields := dmFields or DM_ORIENTATION;
        case FOrientation of
          poPortrait: dmOrientation := DMORIENT_PORTRAIT;
          poLandscape: dmOrientation := DMORIENT_LANDSCAPE;
        end;
      end;
    finally
      GlobalUnlock(DeviceMode);
    end;
    ResetDC(Printer.Handle, DevMode^);
    OpenPrinter(Device, hPrinter, nil);
    try
      GetPrinterDriver(hPrinter, nil, 2, nil, 0, DriverInfo2Size);
      GetMem(DriverInfo2, DriverInfo2Size);
      try
        GetPrinterDriver(hPrinter, nil, 2, DriverInfo2, DriverInfo2Size, DriverInfo2Size);
        StrPCopy(Driver, ExtractFileName(StrPas(DriverInfo2^.PDriverPath)));
      finally
        FreeMem(DriverInfo2, DriverInfo2Size);
      end;
    finally
      ClosePrinter(hPrinter);
    end;
    Printer.SetPrinter(Device, Driver, Port, DeviceMode);
  end;
end;

procedure TPrintPreview.GetPrinterOptions;
var
  DeviceMode: THandle;
  Device, Driver, Port: array[0..MAX_PATH] of Char;
  NewWidth, NewHeight: Integer;
  NewOrientation: TPrinterOrientation;
  NewPaperType: TPaperType;
begin
  if PrinterInstalled then
  begin
    Printer.GetPrinter(Device, Driver, Port, DeviceMode);
    with PDevMode(GlobalLock(DeviceMode))^ do
      try
        NewOrientation := Orientation;
        if (dmFields and DM_ORIENTATION) = DM_ORIENTATION then
          case dmOrientation of
            DMORIENT_PORTRAIT: NewOrientation := poPortrait;
            DMORIENT_LANDSCAPE: NewOrientation := poLandscape;
          end;
        NewPaperType := pCustom;
        if (dmFields and DM_PAPERSIZE) = DM_PAPERSIZE then
          NewPaperType := FindPaperTypeByID(dmPaperSize);
        if NewPaperType = pCustom then
        begin
          NewWidth := TPrintPreviewHelper.ConvertUnits(GetDeviceCaps(Printer.Handle, PHYSICALWIDTH),
            GetDeviceCaps(Printer.Handle, LOGPIXELSX), mmPixel, Units);
          NewHeight := TPrintPreviewHelper.ConvertUnits(GetDeviceCaps(Printer.Handle, PHYSICALHEIGHT),
            GetDeviceCaps(Printer.Handle, LOGPIXELSY), mmPixel, Units);
        end
        else
        begin
          GetPaperTypeSize(NewPaperType, NewWidth, NewHeight, Units);
          if NewOrientation = poLandscape then
            TPrintPreviewHelper.SwapValues(NewWidth, NewHeight);
        end;
        SetPaperSizeOrientation(NewWidth, NewHeight, NewOrientation);
        if (dmFields and DM_FORMNAME) = DM_FORMNAME then
        begin
          FFormName := StrPas(dmFormName);
          FVirtualFormName := '';
        end;
      finally
        GlobalUnlock(DeviceMode);
      end;
  end;
end;

procedure TPrintPreview.ResetPrinterDC;
var
  DeviceMode: THandle;
  DevMode: PDeviceMode;
  Device, Driver, Port: array[0..MAX_PATH] of Char;
begin
  if PrinterInstalled then
  begin
    Printer.GetPrinter(Device, Driver, Port, DeviceMode);
    DevMode := PDevMode(GlobalLock(DeviceMode));
    try
      ResetDC(Printer.Canvas.Handle, DevMode^);
    finally
      GlobalUnlock(DeviceMode);
    end;
  end;
end;

procedure TPrintPreview.InitializePrinting;
begin
  if Assigned(FOnBeforePrint) then
    FOnBeforePrint(Self);
  if not UsePrinterOptions then
    SetPrinterOptions;
  Printer.Title := PrintJobTitle;
  Printer.BeginDoc;
  if not UsePrinterOptions then
    ResetPrinterDC;
end;

procedure TPrintPreview.FinalizePrinting(Succeeded: Boolean);
begin
  if not Succeeded and Printer.Printing then
    Printer.Abort;
  if Printer.Printing then
    Printer.EndDoc;
  Printer.Title := '';
  if Assigned(FOnAfterPrint) then
    FOnAfterPrint(Self);
end;

function TPrintPreview.FetchFormNames(FormNames: TStrings): Boolean;
var
  DeviceMode: THandle;
  Device, Driver, Port: array[0..MAX_PATH] of Char;
  hPrinter: THandle;
  pFormsInfo, pfi: PFormInfo1;
  BytesNeeded: DWORD;
  FormCount: DWORD;
  I: Integer;
begin
  Result := False;
  FormNames.BeginUpdate;
  try
    FormNames.Clear;
    if PrinterInstalled then
    begin
      Printer.GetPrinter(Device, Driver, Port, DeviceMode);
      OpenPrinter(Device, hPrinter, nil);
      try
        BytesNeeded := 0;
        EnumForms(hPrinter, 1, nil, 0, BytesNeeded, FormCount);
        if BytesNeeded > 0 then
        begin
          FormCount := BytesNeeded div SizeOf(TFormInfo1);
          GetMem(pFormsInfo, BytesNeeded);
          try
            if EnumForms(hPrinter, 1, pFormsInfo, BytesNeeded, BytesNeeded, FormCount) then
            begin
              Result := True;
              pfi := pFormsInfo;
              for I := 0 to FormCount - 1 do
              begin
                if (pfi^.Size.cx > 10) and (pfi^.Size.cy > 10) then
                  FormNames.Add(pfi^.pName);
                Inc(pfi);
              end;
            end;
          finally
            FreeMem(pFormsInfo);
          end;
        end;
      finally
        ClosePrinter(hPrinter);
      end;
    end;
  finally
    FormNames.EndUpdate;
  end;
end;

function TPrintPreview.GetFormSize(const AFormName: String;
  out FormWidth, FormHeight: Integer): Boolean;
var
  DeviceMode: THandle;
  Device, Driver, Port: array[0..MAX_PATH] of Char;
  hPrinter: THandle;
  pFormInfo: PFormInfo1;
  BytesNeeded: DWORD;
begin
  Result := False;
  if PrinterInstalled then
  begin
    Printer.GetPrinter(Device, Driver, Port, DeviceMode);
    OpenPrinter(Device, hPrinter, nil);
    try
      BytesNeeded := 0;
      GetForm(hPrinter, PChar(AFormName), 1, nil, 0, BytesNeeded);
      if BytesNeeded > 0 then
      begin
        GetMem(pFormInfo, BytesNeeded);
        try
          if GetForm(hPrinter, PChar(AFormName), 1, pFormInfo, BytesNeeded, BytesNeeded) then
          begin
            with ConvertXY(pFormInfo.Size.cx div 10, pFormInfo.Size.cy div 10, mmHiMetric, Units) do
            begin
              FormWidth := X;
              FormHeight := Y;
            end;
            Result := True;
          end;
        finally
          FreeMem(pFormInfo);
        end;
      end;
    finally
      ClosePrinter(hPrinter);
    end;
  end;
end;

function TPrintPreview.AddNewForm(const AFormName: String;
  FormWidth, FormHeight: DWORD): Boolean;
var
  DeviceMode: THandle;
  Device, Driver, Port: array[0..MAX_PATH] of Char;
  hPrinter: THandle;
  FormInfo: TFormInfo1;
begin
  Result := False;
  if PrinterInstalled then
  begin
    Printer.GetPrinter(Device, Driver, Port, DeviceMode);
    OpenPrinter(Device, hPrinter, nil);
    try
      with FormInfo do
      begin
        Flags := 0;
        pName := PChar(AFormName);
        with ConvertXY(FormWidth, FormHeight, Units, mmHiMetric) do
        begin
          Size.cx := X * 10;
          Size.cy := Y * 10;
        end;
        SetRect(ImageableArea, 0, 0, Size.cx, Size.cy);
      end;
      if AddForm(hPrinter, 1, @FormInfo) then
      begin
        if CompareText(AFormName, FVirtualFormName) = 0 then
          FVirtualFormName := '';
        Result := True;
      end;
    finally
      ClosePrinter(hPrinter);
    end;
  end;
end;

function TPrintPreview.RemoveForm(const AFormName: String): Boolean;
var
  DeviceMode: THandle;
  Device, Driver, Port: array[0..MAX_PATH] of Char;
  hPrinter: THandle;
begin
  Result := False;
  if PrinterInstalled then
  begin
    Printer.GetPrinter(Device, Driver, Port, DeviceMode);
    OpenPrinter(Device, hPrinter, nil);
    try
      if DeleteForm(hPrinter, PChar(AFormName)) then
      begin
        if CompareText(AFormName, FFormName) = 0 then
        begin
          FVirtualFormName := FFormName;
          FFormName := '';
        end;
        Result := True;
      end;
    finally
      ClosePrinter(hPrinter);
    end;
  end;
end;

function TPrintPreview.GetFormName: String;
var
  DeviceMode: THandle;
  Device, Driver, Port: array[0..MAX_PATH] of Char;
  hPrinter: THandle;
  PaperSize: TPoint;
  mmPaperSize: TPoint;
  mmFormSize: TPoint;
  pForms, pf: PFormInfo1;
  BytesNeeded: DWORD;
  FormCount: DWORD;
  IsRotated: Boolean;
  Metric: Boolean;
  I: Integer;
begin
  Result := FFormName;
  if FFormName = '' then
  begin
    if (FVirtualFormName = '') and PrinterInstalled then
    begin
      IsRotated := (Orientation = poLandscape);
      if PaperType <> pCustom then
         GetPaperTypeSize(PaperType, PaperSize.X, PaperSize.Y, mmHiMetric)
      else if IsRotated then
         PaperSize := ConvertXY(PaperHeight, PaperWidth, Units, mmHiMetric)
      else
         PaperSize := ConvertXY(PaperWidth, PaperHeight, Units, mmHiMetric);
      mmPaperSize.X := Round(PaperSize.X / 100);
      mmPaperSize.Y := Round(PaperSize.Y / 100);
      Printer.GetPrinter(Device, Driver, Port, DeviceMode);
      OpenPrinter(Device, hPrinter, nil);
      try
        BytesNeeded := 0;
        EnumForms(hPrinter, 1, nil, 0, BytesNeeded, FormCount);
        if BytesNeeded > 0 then
        begin
          FormCount := BytesNeeded div SizeOf(TFormInfo1);
          GetMem(pForms, BytesNeeded);
          try
            if EnumForms(hPrinter, 1, pForms, BytesNeeded, BytesNeeded, FormCount) then
            begin
              pf := pForms;
              for I := 0 to FormCount - 1 do
              begin
                mmFormSize.X := Round(pf^.Size.cx / 1000);
                mmFormSize.Y := Round(pf^.Size.cy / 1000);
                if (mmFormSize.X = mmPaperSize.X) and (mmFormSize.Y = mmPaperSize.Y) then
                begin
                  FFormName := pf^.pName;
                  FVirtualFormName := '';
                  Result := FFormName;
                  Exit;
                end
                else if (mmFormSize.X = mmPaperSize.Y) and (mmFormSize.Y = mmPaperSize.X) then
                  FVirtualFormName := pf^.pName;
                Inc(pf);
              end;
            end;
          finally
            FreeMem(pForms);
          end;
        end;
      finally
        ClosePrinter(hPrinter);
      end;
      if FVirtualFormName <> '' then
        IsRotated := not IsRotated
      else
      begin
        Metric := True;
        case Units of
          mmLoEnglish, mmHiEnglish:
            Metric := False;
          mmLoMetric, mmHiMetric:
            Metric := True;
        else
          case UserDefaultUnits of
            mmLoEnglish, mmHiEnglish:
              Metric := False;
            mmLoMetric, mmHiMetric:
              Metric := True;
          end;
        end;
        if IsRotated then
          TPrintPreviewHelper.SwapValues(mmPaperSize.X, mmPaperSize.Y);
        if Metric then
          FVirtualFormName := Format('%umm x %umm', [mmPaperSize.X, mmPaperSize.Y])
        else
          with ConvertXY(PaperSize.X, PaperSize.Y, mmHiMetric, mmHiEnglish) do
            FVirtualFormName := Format('%g" x %g"', [Round(X / 100) / 10, Round(Y / 100) / 10]);
      end;
      if IsRotated then
        FVirtualFormName := FVirtualFormName + ' ' + SRotated;
    end;
    Result := FVirtualFormName;
  end;
end;

procedure TPrintPreview.SetFormName(const Value: String);
var
  FormWidth, FormHeight: Integer;
begin
  if (CompareText(FFormName, Value) <> 0) and (FState = psReady) and
      GetFormSize(Value, FormWidth, FormHeight) and
     (FormWidth <> 0) and (FormHeight <> 0) then
  begin
    if Orientation = poPortrait then
      SetPaperSize(FormWidth, FormHeight)
    else
      SetPaperSize(FormHeight, FormWidth);
    FFormName := Value;
    FVirtualFormName := '';
  end;
end;

function TPrintPreview.GetIsDummyFormName: Boolean;
begin
  Result := (CompareText(FormName, FVirtualFormName) = 0);
end;

function TPrintPreview.FindPaperTypeBySize(APaperWidth, APaperHeight: Integer): TPaperType;
var
  Paper: TPaperType;
  InputSize: TPoint;
  PaperSize: TPoint;
begin
  Result := pCustom;
  InputSize := ConvertXY(APaperWidth, APaperHeight, Units, mmHiMetric);
  InputSize.X := Round(InputSize.X / 100);
  InputSize.Y := Round(InputSize.Y / 100);
  for Paper := Low(TPaperType) to High(TPaperType) do
  begin
    PaperSize := ConvertXY(PaperSizes[Paper].Width, PaperSizes[Paper].Height,
      PaperSizes[Paper].Units, mmHiMetric);
    PaperSize.X := Round(PaperSize.X / 100);
    PaperSize.Y := Round(PaperSize.Y / 100);
    if (PaperSize.X = InputSize.X) and (PaperSize.Y = InputSize.Y) then
    begin
      Result := Paper;
      Exit;
    end;
  end;
end;

function TPrintPreview.FindPaperTypeByID(ID: Integer): TPaperType;
var
  Paper: TPaperType;
begin
  Result := pCustom;
  for Paper := Low(TPaperType) to High(TPaperType) do
    if PaperSizes[Paper].ID = ID then
    begin
      Result := Paper;
      Exit;
    end;
end;

function TPrintPreview.GetPaperTypeSize(APaperType: TPaperType;
  out APaperWidth, APaperHeight: Integer; OutUnits: TUnits): Boolean;
begin
  Result := False;
  if APaperType <> pCustom then
  begin
    APaperWidth := ConvertX(PaperSizes[APaperType].Width, PaperSizes[APaperType].Units, OutUnits);
    APaperHeight := ConvertY(PaperSizes[APaperType].Height, PaperSizes[APaperType].Units, OutUnits);
    Result := True;
  end;
end;

procedure TPrintPreview.Resize;
begin
  inherited Resize;
  UpdateZoom;
end;

function TPrintPreview.GetVisiblePageRect: TRect;
begin
  Result := FPaperView.PageRect;
  MapWindowPoints(FPaperView.Handle, Handle, Result, 2);
  IntersectRect(Result, Result, ClientRect);
  MapWindowPoints(Handle, FPaperView.Handle, Result, 2);
  OffsetRect(Result, -FPaperView.BorderWidth, -FPaperView.BorderWidth);
  Result.Left := MulDiv(Result.Left, 100, Zoom);
  Result.Top := MulDiv(Result.Top, 100, Zoom);
  Result.Right := MulDiv(Result.Right, 100, Zoom);
  Result.Bottom := MulDiv(Result.Bottom, 100, Zoom);
end;

procedure TPrintPreview.SetVisiblePageRect(const Value: TRect);
var
  OldZoom: Integer;
  Space: TPoint;
  W, H: Integer;
begin
  OldZoom := FLastZoom;
  Space.X := ClientWidth - 2 * HorzScrollBar.Margin;
  Space.Y := ClientHeight - 2 * VertScrollBar.Margin;
  W := FPaperView.ActualWidth(Value.Right - Value.Left);
  H := FPaperView.ActualHeight(Value.Bottom - Value.Top);
  if Space.X / W < Space.Y / H then
    FZoom := MulDiv(100, Space.X, W)
  else
    FZoom := MulDiv(100, Space.Y, H);
  UpdateZoomEx(Value.Left, Value.Top);
  if OldZoom = FZoom then
  begin
    SyncThumbnail;
    if FZoomState <> zsZoomOther then
    begin
      FZoomState := zsZoomOther;
      if Assigned(FOnZoomChange) then
        FOnZoomChange(Self);
    end;
  end;
end;

function TPrintPreview.CalculateViewSize(const Space: TPoint): TPoint;
begin
  with FPaperView do
  begin
    case FZoomState of
      zsZoomOther:
      begin
        Result.X := ActualWidth(MulDiv(FLogicalExt.X, FZoom, 100));
        Result.Y := ActualHeight(MulDiv(FLogicalExt.Y, FZoom, 100));
      end;
      zsZoomToWidth:
      begin
        Result.X := Space.X;
        Result.Y := ActualHeight(MulDiv(LogicalWidth(Result.X), FLogicalExt.Y, FLogicalExt.X));
      end;
      zsZoomToHeight:
      begin
        Result.Y := Space.Y;
        Result.X := ActualWidth(MulDiv(LogicalHeight(Result.Y), FLogicalExt.X, FLogicalExt.Y));
      end;
      zsZoomToFit:
      begin
        if (FLogicalExt.Y / FLogicalExt.X) < (Space.Y / Space.X) then
        begin
          Result.X := Space.X;
          Result.Y := ActualHeight(MulDiv(LogicalWidth(Result.X), FLogicalExt.Y, FLogicalExt.X));
        end
        else
        begin
          Result.Y := Space.Y;
          Result.X := ActualWidth(MulDiv(LogicalHeight(Result.Y), FLogicalExt.X, FLogicalExt.Y));
        end;
      end;
    end;
    if FZoomState <> zsZoomOther then
      FZoom := Round((100 * LogicalHeight(Result.Y)) / FLogicalExt.Y);
  end;
end;

{$WARNINGS OFF}
procedure TPrintPreview.UpdateZoomEx(X, Y: Integer);
var
  Space: TPoint;
  Position: TPoint;
  ViewPos: TPoint;
  ViewSize: TPoint;
  Percent: TPoint;
begin
  if not HandleAllocated or (csLoading in ComponentState) or
    (not (csDesigning in ComponentState) and (FPageList.Count = 0))
  then
    Exit;

  Space.X := ClientWidth - 2 * HorzScrollBar.Margin;
  Space.Y := ClientHeight - 2 * VertScrollBar.Margin;

  if (Space.X <= 0) or (Space.Y <= 0) then
    Exit;

  if FZoomSavePos and (FCurrentPage <> 0) then
  begin
    Position.X := MulDiv(HorzScrollbar.Position, 100, HorzScrollBar.Range - Space.X);
    if Position.X < 0 then Position.X := 0;
    Position.Y := MulDiv(VertScrollbar.Position, 100, VertScrollbar.Range - Space.Y);
    if Position.Y < 0 then Position.Y := 0;
  end;

  if AutoScroll then
  begin
    if HorzScrollBar.IsScrollBarVisible then
      Inc(Space.Y, GetSystemMetrics(SM_CYHSCROLL));
    if VertScrollBar.IsScrollBarVisible then
      Inc(Space.X, GetSystemMetrics(SM_CXVSCROLL));
  end;

  SendMessage(WindowHandle, WM_SETREDRAW, 0, 0);

  try

    DisableAutoRange;

    try

      HorzScrollbar.Position := 0;
      VertScrollbar.Position := 0;

      ViewSize := CalculateViewSize(Space);

      FCanScrollHorz := (ViewSize.X > Space.X);
      FCanScrollVert := (ViewSize.Y > Space.Y);

      if AutoScroll then
      begin
        if FCanScrollHorz then
        begin
           Dec(Space.Y, GetSystemMetrics(SM_CYHSCROLL));
           FCanScrollVert := (FPaperView.Height > Space.Y);
           if FCanScrollVert then
             Dec(Space.X, GetSystemMetrics(SM_CXVSCROLL));
           ViewSize := CalculateViewSize(Space);
        end
        else if FCanScrollVert then
        begin
           Dec(Space.X, GetSystemMetrics(SM_CXVSCROLL));
           FCanScrollHorz := (FPaperView.Width > Space.X);
           if FCanScrollHorz then
             Dec(Space.Y, GetSystemMetrics(SM_CYHSCROLL));
           ViewSize := CalculateViewSize(Space);
        end;
      end;

      ViewPos.X := HorzScrollBar.Margin;
      if not FCanScrollHorz then
        Inc(ViewPos.X, (Space.X - ViewSize.X) div 2);

      ViewPos.Y := VertScrollBar.Margin;
      if not FCanScrollVert then
        Inc(ViewPos.Y, (Space.Y - ViewSize.Y) div 2);

      FPaperView.SetBounds(ViewPos.X, ViewPos.Y, ViewSize.X, ViewSize.Y);

    finally
      EnableAutoRange;
    end;

    if FCurrentPage <> 0 then
    begin
      if FCanScrollHorz then
      begin
        if X >= 0 then
          HorzScrollbar.Position := MulDiv(X, HorzScrollBar.Range, FLogicalExt.X)
        else if FZoomSavePos then
          HorzScrollbar.Position := MulDiv(Position.X, HorzScrollBar.Range - Space.X, 100);
      end;
      if FCanScrollVert then
      begin
        if Y >= 0 then
          VertScrollBar.Position := MulDiv(Y, VertScrollBar.Range, FLogicalExt.Y)
        else if FZoomSavePos then
          VertScrollbar.Position := MulDiv(Position.Y, VertScrollbar.Range - Space.Y, 100);
      end;
    end;

  finally
    SendMessage(WindowHandle, WM_SETREDRAW, 1, 0);
    Invalidate;
  end;

  FIsDragging := False;
  if FCanScrollHorz or FCanScrollVert then
    FPaperView.Cursor := FPaperViewOptions.DragCursor
  else
    FPaperView.Cursor := FPaperViewOptions.Cursor;

  if (ViewSize.X <> FPaperView.Width) or (ViewSize.Y <> FPaperView.Height) then
  begin
    Percent.X := (MulDiv(100, FPaperView.Width, FLogicalExt.X) div FZoomStep) * FZoomStep;
    Percent.Y := (MulDiv(100, FPaperView.Height, FLogicalExt.Y) div FZoomStep) * FZoomStep;
    if Percent.X < Percent.Y then
      FZoom := Percent.X
    else
      FZoom := Percent.Y;
    UpdateZoomEx(X, Y);
  end
  else
  begin
    if FLastZoom <> FZoom then
    begin
      FLastZoom := FZoom;
      Update;
      if Assigned(FOnZoomChange) then
        FOnZoomChange(Self);
    end;
    SyncThumbnail;
  end;
end;
{$WARNINGS ON}

procedure TPrintPreview.UpdateZoom;
begin
  UpdateZoomEx(-1, -1);
end;

procedure TPrintPreview.ChangeState(NewState: TPreviewState);
begin
  if FState <> NewState then
  begin
    FState := NewState;
    if Assigned(FOnStateChange) then
      FOnStateChange(Self);
  end;
end;

procedure TPrintPreview.PaintPage(Sender: TObject; Canvas: TCanvas;
  const Rect: TRect);
var
  sx, sy: Double;
begin
  if (FCurrentPage >= 1) and (FCurrentPage <= TotalPages) then
  begin
    PreviewPage(FCurrentPage, Canvas, Rect);
    if FShowPrintableArea then
    begin
      sx := (Rect.Right - Rect.Left) / FPageExt.X;
      sy := (Rect.Bottom - Rect.Top) / FPageExt.Y;
      with Canvas, PrinterPageBounds do
      begin
        Pen.Mode := pmMask;
        Pen.Width := 0;
        Pen.Style := psDot;
        Pen.Color := FPrintableAreaColor;
        MoveTo(Round(sx * Left), Rect.Top);
        LineTo(Round(sx * Left), Rect.Bottom);
        MoveTo(Round(sx * Right), Rect.Top);
        LineTo(Round(sx * Right), Rect.Bottom);
        MoveTo(Rect.Left, Round(sy * Top));
        LineTo(Rect.Right, Round(sy * Top));
        MoveTo(Rect.Left, Round(sy * Bottom));
        LineTo(Rect.Right, Round(sy * Bottom));
      end;
    end;
  end;
end;

procedure TPrintPreview.PreviewPage(PageNo: Integer; Canvas: TCanvas;
  const Rect: TRect);
begin
  if Assigned(BackgroundMetafile) then
    Canvas.StretchDraw(Rect, BackgroundMetafile);
  DrawPage(PageNo, Canvas, Rect, gsPreview in FGrayscale);
  if Assigned(AnnotationMetafile) then
    Canvas.StretchDraw(Rect, AnnotationMetafile);
end;

procedure TPrintPreview.PrintPage(PageNo: Integer; Canvas: TCanvas;
  const Rect: TRect);
begin
  if Assigned(FOnPrintBackground) then
    FOnPrintBackground(Self, PageNo, Canvas);
  if gsPrint in Grayscale then
    TPrintPreviewHelper.StretchDrawGrayscale(Canvas, Rect, FPageList[PageNo-1], FGrayBrightness, FGrayContrast)
  else
    Canvas.StretchDraw(Rect, FPageList[PageNo-1]);
  if Assigned(FOnPrintAnnotation) then
    FOnPrintAnnotation(Self, PageNo, Canvas);
end;

procedure TPrintPreview.DrawPage(PageNo: Integer; Canvas: TCanvas;
  const Rect: TRect; Gray: Boolean);
var
  Bitmap: TBitmap;
  VisibleRect: TRect;
  BitmapRect: TRect;
begin
  if not Gray then
    gdiPlus.Draw(Canvas, Rect, FPageList[PageNo-1])
  else if IntersectRect(VisibleRect, Canvas.ClipRect, Rect) then
  begin
    InflateRect(VisibleRect, 1, 1);
    BitmapRect := Rect;
    OffsetRect(BitmapRect, -VisibleRect.Left, -VisibleRect.Top);
    Bitmap := TBitmap.Create;
    try
      Bitmap.Canvas.Brush.Color := FPaperView.PaperColor;
      Bitmap.Width := VisibleRect.Right - VisibleRect.Left;
      Bitmap.Height := VisibleRect.Bottom - VisibleRect.Top;
      Bitmap.TransparentColor := FPaperView.PaperColor;
      Bitmap.Transparent := True;
      gdiPlus.Draw(Bitmap.Canvas, BitmapRect, FPageList[PageNo-1]);
      TPrintPreviewHelper.ConvertBitmapToGrayscale(Bitmap, FGrayBrightness, FGrayContrast);
      Canvas.Draw(VisibleRect.Left, VisibleRect.Top, Bitmap);
    finally
      Bitmap.Free;
    end;
  end;
end;

procedure TPrintPreview.PaperViewOptionsChanged(Sender: TObject;
  Severity: TUpdateSeverity);
begin
  FPaperViewOptions.AssignTo(FPaperView);
  if Severity = usRecreate then
    UpdateZoom;
end;

procedure TPrintPreview.PagesChanged(Sender: TObject;
  PageStartIndex, PageEndIndex: Integer);
var
  Rebuild: Boolean;
begin
  Rebuild := False;
  if PageEndIndex < 0 then
  begin
    FCurrentPage := 0;
    FPaperView.Visible := False;
    Repaint;
  end
  else
  begin
    if FCurrentPage = 0 then
    begin
      FCurrentPage := 1;
      UpdateZoom;
      FPaperView.Visible := True;
      Rebuild := True;
    end;
    if (FCurrentPage >= PageStartIndex + 1) and (FCurrentPage <= PageEndIndex + 1) then
    begin
      DoBackground(FCurrentPage);
      DoAnnotation(FCurrentPage);
      FPaperView.Repaint;
    end;
    Update;
  end;
  if Rebuild then
    RebuildThumbnails
  else
    UpdateThumbnails(PageStartIndex, PageEndIndex);
  if Assigned(FOnChange) then
    FOnChange(Self);
end;

procedure TPrintPreview.PageChanged(Sender: TObject; PageIndex: Integer);
begin
  if PageIndex + 1 = FCurrentPage then
  begin
    DoBackground(FCurrentPage);
    DoAnnotation(FCurrentPage);
    FPaperView.Repaint;
  end;
  RepaintThumbnails(PageIndex, PageIndex);
  if Assigned(FOnChange) then
    FOnChange(Self);
end;

function TPrintPreview.HorzPixelsPerInch: Integer;
begin
  if ReferenceDC <> 0 then
    Result := GetDeviceCaps(ReferenceDC, LOGPIXELSX)
  else
    Result := Screen.PixelsPerInch;
end;

function TPrintPreview.VertPixelsPerInch: Integer;
begin
  if ReferenceDC <> 0 then
    Result := GetDeviceCaps(ReferenceDC, LOGPIXELSY)
  else
    Result := Screen.PixelsPerInch;
end;

procedure TPrintPreview.SetPaperViewOptions(Value: TPaperPreviewOptions);
begin
  FPaperViewOptions.Assign(Value);
end;

procedure TPrintPreview.SetUnits(Value: TUnits);
begin
  if FUnits <> Value then
  begin
    if FPaperType <> pCustom then
    begin
      GetPaperTypeSize(FPaperType, FPageExt.X, FPageExt.Y, Value);
      if FOrientation = poLandscape then
        TPrintPreviewHelper.SwapValues(FPageExt.X, FPageExt.Y);
    end
    else
      ConvertPoints(FPageExt, 1, FUnits, Value);
    if Assigned(FPageCanvas) then
    begin
      FPageCanvas.Pen.Width := ConvertX(FPageCanvas.Pen.Width, FUnits, Value);
      ScaleCanvas(FPageCanvas);
    end;
    FUnits := Value;
  end;
end;

procedure TPrintPreview.DoPaperChange;
begin
  FFormName := '';
  FVirtualFormName := '';
  UpdateExtends;
  UpdateZoom;
  if Assigned(FOnPaperChange) then
    FOnPaperChange(Self);
end;

procedure TPrintPreview.SetPaperType(Value: TPaperType);
begin
  if (FPaperType <> Value) and (FState = psReady) then
  begin
    FPaperType := Value;
    if FPaperType <> pCustom then
    begin
      with PaperSizes[FPaperType] do
        FPageExt := ConvertXY(Width, Height, Units, FUnits);
      if FOrientation = poLandscape then
        TPrintPreviewHelper.SwapValues(FPageExt.X, FPageExt.Y);
      DoPaperChange;
    end;
  end;
end;

procedure TPrintPreview.SetPaperSize(AWidth, AHeight: Integer);
begin
  if AWidth < 1 then AWidth := 1;
  if AHeight < 1 then AHeight := 1;
  if ((FPageExt.X <> AWidth) or (FPageExt.Y <> AHeight)) and (FState = psReady) then
  begin
    FPageExt.X := AWidth;
    FPageExt.Y := AHeight;
    if FOrientation = poLandscape then
      FPaperType := FindPaperTypeBySize(FPageExt.Y, FPageExt.X)
    else
      FPaperType := FindPaperTypeBySize(FPageExt.X, FPageExt.Y);
    DoPaperChange;
  end;
end;

procedure TPrintPreview.SetOrientation(Value: TPrinterOrientation);
begin
  if (FOrientation <> Value) and (FState = psReady) then
  begin
    FOrientation := Value;
    TPrintPreviewHelper.SwapValues(FPageExt.X, FPageExt.Y);
    DoPaperChange;
  end;
end;

procedure TPrintPreview.SetPaperSizeOrientation(AWidth, AHeight: Integer;
  AOrientation: TPrinterOrientation);
begin
  if AWidth < 1 then AWidth := 1;
  if AHeight < 1 then AHeight := 1;
  if (FOrientation <> AOrientation) or
     ((AOrientation = FOrientation) and ((FPageExt.X <> AWidth) or (FPageExt.Y <> AHeight))) or
     ((AOrientation <> FOrientation) and ((FPageExt.X <> AHeight) or (FPageExt.Y <> AWidth))) then
  begin
    FPageExt.X := AWidth;
    FPageExt.Y := AHeight;
    FOrientation := AOrientation;
    if FOrientation = poPortrait then
      FPaperType := FindPaperTypeBySize(FPageExt.X, FPageExt.Y)
    else
      FPaperType := FindPaperTypeBySize(FPageExt.Y, FPageExt.X);
    DoPaperChange;
  end;
end;

function TPrintPreview.GetPaperWidth: Integer;
begin
  Result := FPageExt.X;
end;

procedure TPrintPreview.SetPaperWidth(Value: Integer);
begin
  SetPaperSize(Value, FPageExt.Y);
end;

function TPrintPreview.GetPaperHeight: Integer;
begin
  Result := FPageExt.Y;
end;

procedure TPrintPreview.SetPaperHeight(Value: Integer);
begin
  SetPaperSize(FPageExt.X, Value);
end;

function TPrintPreview.GetPageBounds: TRect;
begin
  Result.Left := 0;
  Result.Top := 0;
  Result.BottomRight := FPageExt;
end;

function TPrintPreview.GetPrinterPageBounds: TRect;
var
  Offset: TPoint;
  Size: TPoint;
  DPI: TPoint;
begin
  if PrinterInstalled then
  begin
    DPI.X := GetDeviceCaps(Printer.Handle, LOGPIXELSX);
    DPI.Y := GetDeviceCaps(Printer.Handle, LOGPIXELSY);
    Offset.X := GetDeviceCaps(Printer.Handle, PHYSICALOFFSETX);
    Offset.Y := GetDeviceCaps(Printer.Handle, PHYSICALOFFSETY);
    Offset.X := TPrintPreviewHelper.ConvertUnits(Offset.X, DPI.X, mmPixel, Units);
    Offset.Y := TPrintPreviewHelper.ConvertUnits(Offset.Y, DPI.Y, mmPixel, Units);
    Size.X := GetDeviceCaps(Printer.Handle, HORZRES);                           //Mixy
    Size.Y := GetDeviceCaps(Printer.Handle, VERTRES);                           //Mixy
    Size.X := TPrintPreviewHelper.ConvertUnits(Size.X, DPI.X, mmPixel, Units);                      //Mixy
    Size.Y := TPrintPreviewHelper.ConvertUnits(Size.Y, DPI.Y, mmPixel, Units);                      //Mixy
    SetRect(Result, Offset.X, Offset.Y, Offset.X + Size.X, Offset.Y + Size.Y);  //Mixy
  end
  else
    Result := PageBounds;
end;

function TPrintPreview.GetPrinterPhysicalPageBounds: TRect;
begin
  Result.Left := 0;
  Result.Top := 0;
  Result.Right := 0;
  Result.Bottom := 0;
  if PrinterInstalled then
  begin
    if UsePrinterOptions then
    begin
      Result.Right := GetDeviceCaps(Printer.Handle, PHYSICALWIDTH);
      Result.Bottom := GetDeviceCaps(Printer.Handle, PHYSICALHEIGHT);
    end
    else
    begin
      Result.Right := TPrintPreviewHelper.ConvertUnits(FPageExt.X,
        GetDeviceCaps(Printer.Handle, LOGPIXELSX), FUnits, mmPixel);
      Result.Bottom := TPrintPreviewHelper.ConvertUnits(FPageExt.Y,
        GetDeviceCaps(Printer.Handle, LOGPIXELSY), FUnits, mmPixel);
    end;
    OffsetRect(Result,
       -GetDeviceCaps(Printer.Handle, PHYSICALOFFSETX),
       -GetDeviceCaps(Printer.Handle, PHYSICALOFFSETY));
  end;
end;

function TPrintPreview.IsPaperCustom: Boolean;
begin
  Result := (FPaperType = pCustom);
end;

function TPrintPreview.IsPaperRotated: Boolean;
begin
  Result := (FOrientation = poLandscape);
end;

procedure TPrintPreview.SetZoom(Value: Integer);
var
  OldZoom: Integer;
begin
  if Value < FZoomMin then Value := FZoomMin
  else if Value > FZoomMax then Value := FZoomMax;
  if (FZoom <> Value) or (FZoomState <> zsZoomOther) then
  begin
    OldZoom := FZoom;
    FZoom := Value;
    FZoomState := zsZoomOther;
    UpdateZoom;
    if (OldZoom = FZoom) and Assigned(FOnZoomChange) then
      FOnZoomChange(Self);
  end;
end;

function TPrintPreview.IsZoomStored: Boolean;
begin
  Result := (FZoomState = zsZoomOther) and (FZoom <> 100);
end;

procedure TPrintPreview.SetZoomMin(Value: Integer);
begin
  if (FZoomMin <> Value) and (Value >= 1) and (Value <= FZoomMax) then
  begin
    FZoomMin := Value;
    if (FZoomState = zsZoomOther) and (FZoom < FZoomMin) then
      Zoom := FZoomMin;
  end;
end;

procedure TPrintPreview.SetZoomMax(Value: Integer);
begin
  if (FZoomMax <> Value) and (Value >= FZoomMin) then
  begin
    FZoomMax := Value;
    if (FZoomState = zsZoomOther) and (FZoom > FZoomMax) then
      Zoom := FZoomMax;
  end;
end;

procedure TPrintPreview.SetZoomState(Value: TZoomState);
var
  OldZoom: Integer;
begin
  if FZoomState <> Value then
  begin
    OldZoom := FZoom;
    FZoomState := Value;
    UpdateZoom;
    if (OldZoom = FZoom) and Assigned(FOnZoomChange) then
      FOnZoomChange(Self);
  end;
end;

procedure TPrintPreview.SetCurrentPage(Value: Integer);
begin
  if TotalPages <> 0 then
  begin
    if Value < 1 then Value := 1;
    if Value > TotalPages then Value := TotalPages;
    if FCurrentPage <> Value then
    begin
      FCurrentPage := Value;
      DoBackground(FCurrentPage);
      DoAnnotation(FCurrentPage);
      FPaperView.Repaint;
      SyncThumbnail;
      if Assigned(FOnChange) then
        FOnChange(Self);
    end;
  end;
end;

procedure TPrintPreview.SetGrayscale(Value: TGrayscaleOptions);
begin
  if Grayscale <> Value then
  begin
    FGrayscale := Value;
    FPaperView.Repaint;
    RecolorThumbnails(False);
  end;
end;

procedure TPrintPreview.SetGrayBrightness(Value: Integer);
begin
  if Value < -100 then
    Value := -100
  else if Value > 100 then
    Value := 100;
  if GrayBrightness <> Value then
  begin
    FGrayBrightness := Value;
    if gsPreview in Grayscale then
    begin
      FPaperView.Repaint;
      RecolorThumbnails(True);
    end;
  end;
end;

procedure TPrintPreview.SetGrayContrast(Value: Integer);
begin
  if Value < -100 then
    Value := -100
  else if Value > 100 then
    Value := 100;
  if GrayContrast <> Value then
  begin
    FGrayContrast := Value;
    if gsPreview in Grayscale then
    begin
      FPaperView.Repaint;
      RecolorThumbnails(True);
    end;
  end;
end;

function TPrintPreview.GetCacheSize: Integer;
begin
  Result := FPageList.CacheSize;
end;

procedure TPrintPreview.SetCacheSize(Value: Integer);
begin
  FPageList.CacheSize := Value;
end;

procedure TPrintPreview.SetShowPrintableArea(Value: Boolean);
begin
  if FShowPrintableArea <> Value then
  begin
    FShowPrintableArea := Value;
    if CurrentPage <> 0 then
      FPaperView.Refresh;
  end;
end;

procedure TPrintPreview.SetPrintableAreaColor(Value: TColor);
begin
  if FPrintableAreaColor <> Value then
  begin
    FPrintableAreaColor := Value;
    if FShowPrintableArea and (CurrentPage <> 0) then
      FPaperView.Refresh;
  end;
end;

procedure TPrintPreview.SetDirectPrint(Value: Boolean);
begin
  if FDirectPrint <> Value then
  begin
    FDirectPrint := Value;
    if FDirectPrint and PrinterInstalled then
      ReferenceDC := Printer.Handle
    else
      ReferenceDC := 0;
    UpdateExtends;
  end;
end;

procedure TPrintPreview.SetPDFDocumentInfo(Value: TPDFDocumentInfo);
begin
  FPDFDocumentInfo.Assign(Value);
end;

function TPrintPreview.GetTotalPages: Integer;
begin
  if FDirectPrinting then
    Result := FDirectPrintPageCount
  else
    Result := FPageList.Count;
end;

function TPrintPreview.GetPages(PageNo: Integer): TMetafile;
begin
  if (PageNo >= 1) and (PageNo <= TotalPages) then
    Result := FPageList[PageNo-1]
  else
    Result := nil;
end;

function TPrintPreview.GetCanvas: TCanvas;
begin
  if Assigned(FPageCanvas) then
    Result := FPageCanvas
  else
    Result := Printer.Canvas;
end;

function TPrintPreview.GetPrinterInstalled: Boolean;
begin
  Result := (Printer.Printers.Count > 0);
end;

function TPrintPreview.GetPrinter: TPrinter;
begin
  Result := Vcl.Printers.Printer;
end;

procedure TPrintPreview.ScaleCanvas(ACanvas: TCanvas);
var
  FontSize: Integer;
  LogExt, DevExt: TPoint;
begin
  LogExt := FPageExt;
  DevExt.X := TPrintPreviewHelper.ConvertUnits(LogExt.X,
    GetDeviceCaps(ACanvas.Handle, LOGPIXELSX), FUnits, mmPixel);
  DevExt.Y := TPrintPreviewHelper.ConvertUnits(LogExt.Y,
    GetDeviceCaps(ACanvas.Handle, LOGPIXELSY), FUnits, mmPixel);
  SetMapMode(ACanvas.Handle, MM_ANISOTROPIC);
  SetWindowExtEx(ACanvas.Handle, LogExt.X, LogExt.Y, nil);
  SetViewPortExtEx(ACanvas.Handle, DevExt.X, DevExt.Y, nil);
  SetViewportOrgEx(ACanvas.Handle,
    -GetDeviceCaps(ACanvas.Handle, PHYSICALOFFSETX),
    -GetDeviceCaps(ACanvas.Handle, PHYSICALOFFSETY), nil);
  FontSize := ACanvas.Font.Size;
  ACanvas.Font.PixelsPerInch :=
    MulDiv(GetDeviceCaps(ACanvas.Handle, LOGPIXELSY), LogExt.Y, DevExt.Y);
  ACanvas.Font.Size := FontSize;
end;

procedure TPrintPreview.UpdateExtends;
begin
  FDeviceExt.X := ConvertX(FPageExt.X, FUnits, mmPixel);
  FDeviceExt.Y := ConvertX(FPageExt.Y, FUnits, mmPixel);
  FLogicalExt.X := MulDiv(FDeviceExt.X, Screen.PixelsPerInch, HorzPixelsPerInch);
  FLogicalExt.Y := MulDiv(FDeviceExt.Y, Screen.PixelsPerInch, VertPixelsPerInch);
end;

procedure TPrintPreview.CreateMetafileCanvas(out AMetafile: TMetafile;
  out ACanvas: TCanvas);
var
  aDC:HDC;
begin
  AMetafile := TMetafile.Create;
  try
    aDC:=ReferenceDC;
    if aDC=0 then
      aDC:=FPaperView.Canvas.Handle;
    with TPrintPreviewHelper.ScaleToDeviceContext(aDC, FDeviceExt) do
    begin
      AMetafile.Width := X;
      AMetafile.Height := Y;
    end;
    ACanvas := TMetafileCanvas.CreateWithComment(AMetafile, ReferenceDC,
      ClassName, PrintJobTitle);
    if ACanvas.Handle = 0 then
    begin
      ACanvas.Free;
      ACanvas := nil;
      TPrintPreviewHelper.RaiseOutOfMemory;
    end;
  except
    AMetafile.Free;
    AMetafile := nil;
    raise;
  end;
  ACanvas.Font.Assign(Font);
  ScaleCanvas(ACanvas);
  SetBkColor(ACanvas.Handle, RGB(255, 255, 255));
  SetBkMode(ACanvas.Handle, TRANSPARENT);
end;

procedure TPrintPreview.CloseMetafileCanvas(var AMetafile: TMetafile;
  var ACanvas: TCanvas);
begin
  ACanvas.Free;
  ACanvas := nil;
  if AMetafile.Handle = 0 then
  begin
    AMetafile.Free;
    AMetafile := nil;
    TPrintPreviewHelper.RaiseOutOfMemory;
  end;
end;

procedure TPrintPreview.CreatePrinterCanvas(out ACanvas: TCanvas);
begin
  ACanvas := TCanvas.Create;
  try
    ACanvas.Handle := Printer.Handle;
    ScaleCanvas(ACanvas);
  except
    ACanvas.Free;
    ACanvas := nil;
    raise;
  end;
end;

procedure TPrintPreview.ClosePrinterCanvas(var ACanvas: TCanvas);
begin
  ACanvas.Handle := 0;
  ACanvas.Free;
  ACanvas := nil;
end;

procedure TPrintPreview.Clear;
begin
  FPageList.Clear;
end;

procedure TPrintPreview.BeginDoc;
begin
  if FState = psReady then
  begin
    FPageCanvas := nil;
    if not FDirectPrint then
    begin
      Clear;
      ChangeState(psCreating);
      if UsePrinterOptions then
        GetPrinterOptions;
      FDirectPrinting := False;
      ReferenceDC := 0;
    end
    else
    begin
      ChangeState(psPrinting);
      FDirectPrinting := True;
      FDirectPrintPageCount := 0;
      if UsePrinterOptions then
        GetPrinterOptions
      else
        SetPrinterOptions;
      Printer.Title := PrintJobTitle;
      Printer.BeginDoc;
      ReferenceDC := Printer.Handle;
    end;
    UpdateExtends;
    if Assigned(FOnBeginDoc) then
      FOnBeginDoc(Self);
    NewPage;
  end
end;

procedure TPrintPreview.EndDoc;
begin
  if ((FState = psCreating) and not FDirectPrinting) or
     ((FState = psPrinting) and FDirectPrinting) then
  begin
    if Assigned(FOnEndPage) then
      FOnEndPage(Self);
    FCanvasPageNo := 0;
    if not FDirectPrinting then
    begin
      try
        CloseMetafileCanvas(PageMetafile, FPageCanvas);
        FPageList.Add(PageMetafile);
      finally
        PageMetafile.Free;
        PageMetafile := nil;
      end;
    end
    else
    begin
      Inc(FDirectPrintPageCount);
      ClosePrinterCanvas(FPageCanvas);
      Printer.EndDoc;
      FDirectPrinting := False;
    end;
    if Assigned(FOnEndDoc) then
      FOnEndDoc(Self);
    ChangeState(psReady);
  end;
end;

procedure TPrintPreview.NewPage;
begin
  if ((FState = psCreating) and not FDirectPrinting) or
     ((FState = psPrinting) and FDirectPrinting) then
  begin
    if Assigned(FPageCanvas) and Assigned(FOnEndPage) then
      FOnEndPage(Self);
    if not FDirectPrinting then
    begin
      if Assigned(FPageCanvas) then
      begin
        CloseMetafileCanvas(PageMetafile, FPageCanvas);
        try
          FPageList.Add(PageMetafile);
        finally
          PageMetafile.Free;
          PageMetafile := nil;
        end;
      end;
      CreateMetafileCanvas(PageMetafile, FPageCanvas);
    end
    else
    begin
      if Assigned(FPageCanvas) then
      begin
        Inc(FDirectPrintPageCount);
        Printer.NewPage;
      end
      else
        CreatePrinterCanvas(FPageCanvas);
      FPageCanvas.Font.Assign(Font);
    end;
    Inc(FCanvasPageNo);
    if Assigned(FOnNewPage) then
      FOnNewPage(Self);
  end;
end;

function TPrintPreview.BeginEdit(PageNo: Integer): Boolean;
begin
  Result := False;
  if (FState = psReady) and (PageNo > 0) and (PageNo <= TotalPages) then
  begin
    ChangeState(psEditing);
    CreateMetafileCanvas(PageMetafile, FPageCanvas);
    FCanvasPageNo := PageNo;
    FPageCanvas.StretchDraw(PageBounds, FPageList[FCanvasPageNo - 1]);
    Result := True;
  end;
end;

procedure TPrintPreview.EndEdit(Cancel: Boolean);
begin
  if FState = psEditing then
  begin
    try
      CloseMetafileCanvas(PageMetafile, FPageCanvas);
      if not Cancel then
        FPageList[FCanvasPageNo - 1].Assign(PageMetafile);
    finally
      PageMetafile.Free;
      PageMetafile := nil;
      FCanvasPageNo := 0;
      ChangeState(psReady);
    end;
  end;
end;

function TPrintPreview.BeginReplace(PageNo: Integer): Boolean;
begin
  Result := False;
  if (FState = psReady) and (PageNo > 0) and (PageNo <= TotalPages) then
  begin
    ChangeState(psReplacing);
    CreateMetafileCanvas(PageMetafile, FPageCanvas);
    FCanvasPageNo := PageNo;
    if Assigned(FOnNewPage) then
      FOnNewPage(Self);
    Result := True;
  end;
end;

procedure TPrintPreview.EndReplace(Cancel: Boolean);
begin
  if FState = psReplacing then
  begin
    try
      CloseMetafileCanvas(PageMetafile, FPageCanvas);
      if not Cancel then
        FPageList[FCanvasPageNo - 1].Assign(PageMetafile);
    finally
      PageMetafile.Free;
      PageMetafile := nil;
      FCanvasPageNo := 0;
      ChangeState(psReady);
    end;
  end;
end;

function TPrintPreview.BeginInsert(PageNo: Integer): Boolean;
begin
  Result := False;
  if FState = psReady then
  begin
    ChangeState(psInserting);
    CreateMetafileCanvas(PageMetafile, FPageCanvas);
    if PageNo <= 0 then
      FCanvasPageNo := 1
    else if PageNo > TotalPages then
      FCanvasPageNo := TotalPages + 1
    else
      FCanvasPageNo := PageNo;
    if Assigned(FOnNewPage) then
      FOnNewPage(Self);
    Result := True;
  end;
end;

procedure TPrintPreview.EndInsert(Cancel: Boolean);
begin
  if FState = psInserting then
  begin
    try
      CloseMetafileCanvas(PageMetafile, FPageCanvas);
      if not Cancel then
      begin
        if FCurrentPage >= FCanvasPageNo then
          Inc(FCurrentPage);
        FPageList.Insert(FCanvasPageNo - 1, PageMetafile);
      end;
    finally
      PageMetafile.Free;
      PageMetafile := nil;
      FCanvasPageNo := 0;
      ChangeState(psReady);
    end;
  end;
end;

function TPrintPreview.BeginAppend: Boolean;
begin
  Result := BeginInsert(MaxInt);
end;

procedure TPrintPreview.EndAppend(Cancel: Boolean);
begin
  EndInsert(Cancel);
end;

function TPrintPreview.Delete(PageNo: Integer): Boolean;
begin
  Result := False;
  if (FState = psReady) and (PageNo > 0) and (PageNo <= TotalPages) then
  begin
    if (PageNo < FCurrentPage) or ((PageNo = FCurrentPage) and (PageNo = TotalPages)) then
      Dec(FCurrentPage);
    FPageList.Delete(PageNo - 1);
    Result := True;
  end;
end;

function TPrintPreview.Move(PageNo, NewPageNo: Integer): Boolean;
begin
  Result := False;
  if (FState = psReady) and (PageNo <> NewPageNo) and
     (PageNo > 0) and (PageNo <= TotalPages) and
     (NewPageNo > 0) and (NewPageNo <= TotalPages) then
  begin
    if PageNo = FCurrentPage then
      FCurrentPage := NewPageNo
    else if NewPageNo = FCurrentPage then
      FCurrentPage := NewPageNo;
    FPageList.Move(PageNo - 1, NewPageNo - 1);
    Result := True;
  end;
end;

function TPrintPreview.Exchange(PageNo1, PageNo2: Integer): Boolean;
begin
  Result := False;
  if (FState = psReady) and (PageNo1 <> PageNo2) and
     (PageNo1 > 0) and (PageNo1 <= TotalPages) and
     (PageNo2 > 0) and (PageNo2 <= TotalPages) then
  begin
    if PageNo1 = FCurrentPage then
      FCurrentPage := PageNo2
    else if PageNo2 = FCurrentPage then
      FCurrentPage := PageNo1;
    FPageList.Exchange(PageNo1 - 1, PageNo2 - 1);
    Result := True;
  end;
end;

function TPrintPreview.LoadPageInfo(Stream: TStream): Boolean;
var
  Header: TStreamHeader;
  Data: Integer;
begin
  Result := False;
  Stream.ReadBuffer(Header, SizeOf(Header));
  if CompareMem(@Header.Signature, @PageInfoHeader.Signature, SizeOf(Header.Signature))  then
  begin
    Stream.ReadBuffer(Data, SizeOf(Data));
    FOrientation := TPrinterOrientation(Data);
    Stream.ReadBuffer(Data, SizeOf(Data));
    FPaperType := TPaperType(Data);
    Stream.ReadBuffer(Data, SizeOf(Data));
    FPageExt.X := ConvertX(Data, mmHiMetric, FUnits);
    Stream.ReadBuffer(Data, SizeOf(Data));
    FPageExt.Y := ConvertY(Data, mmHiMetric, FUnits);
    UpdateExtends;
    Result := True;
  end;
end;

procedure TPrintPreview.SavePageInfo(Stream: TStream);
var
  Data: Integer;
begin
  Stream.WriteBuffer(PageInfoHeader, SizeOf(PageInfoHeader));
  Data := Ord(FOrientation);
  Stream.WriteBuffer(Data, SizeOf(Data));
  Data := Ord(FPaperType);
  Stream.WriteBuffer(Data, SizeOf(Data));
  Data := ConvertX(FPageExt.X, FUnits, mmHiMetric);
  Stream.WriteBuffer(Data, SizeOf(Data));
  Data := ConvertY(FPageExt.Y, FUnits, mmHiMetric);
  Stream.WriteBuffer(Data, SizeOf(Data));
end;

procedure TPrintPreview.LoadFromStream(Stream: TStream);
begin
  ChangeState(psLoading);
  try
    if not LoadPageInfo(Stream) or not FPageList.LoadFromStream(Stream) then
      raise EPreviewLoadError.Create(SLoadError);
  finally
    ChangeState(psReady);
  end;
end;

procedure TPrintPreview.SaveToStream(Stream: TStream);
begin
  ChangeState(psSaving);
  try
    SavePageInfo(Stream);
    FPageList.SaveToStream(Stream);
  finally
    ChangeState(psReady);
  end;
end;

procedure TPrintPreview.LoadFromFile(const FileName: String);
var
  FileStream: TFileStream;
begin
  FileStream := TFileStream.Create(FileName, fmOpenRead or fmShareDenyWrite);
  try
    LoadFromStream(FileStream);
  finally
    FileStream.Free;
  end;
end;

procedure TPrintPreview.SaveToFile(const FileName: String);
var
  FileStream: TFileStream;
begin
  FileStream := TFileStream.Create(FileName, fmCreate or fmShareExclusive);
  try
    SaveToStream(FileStream);
  finally
    FileStream.Free;
  end;
end;

procedure TPrintPreview.SaveAsTIF(const FileName: String);
var
  MF: Pointer;
  PageNo: Integer;
begin
  if (TotalPages > 0) and gdiPlus.Exists then
  begin
    ChangeState(psSavingTIF);
    try
      MF := nil;
      try
        DoProgress(0, TotalPages);
        for PageNo := 1 to TotalPages do
        begin
          case DoPageProcessing(PageNo) of
            pcAccept:
             if Assigned(MF) then
                gdiPlus.MultiFrameNext(MF, Pages[PageNo], PaperView.PaperColor)
              else
                MF := gdiPlus.MultiFrameBegin(WideString(FileName), Pages[PageNo], PaperView.PaperColor);
            pcCancellAll:
              Exit;
          end;
          DoProgress(PageNo, TotalPages);
        end;
      finally
        if Assigned(MF) then
          gdiPlus.MultiFrameEnd(MF);
      end;
    finally
      ChangeState(psReady);
    end;
  end;
end;

function TPrintPreview.CanSaveAsTIF: Boolean;
begin
  Result := gdiPlus.Exists;
end;

procedure TPrintPreview.SaveAsPDF(const FileName: String);
{$IFDEF PDF_SYNOPSE}
var
  PageNo: Integer;
  pdf: TPdfDocument;
{$ELSEIF PDF_DSPDF}
var
  PageNo: Integer;
  AnyPageRendered: Boolean;
{$ELSEIF PDF_WPDF}
{$IFEND}
begin
{$IFDEF PDF_SYNOPSE}
  pdf := TPdfDocument.Create;
  try
    ChangeState(psSavingPDF);
    try
      pdf.Info.CreationDate := Now;
      pdf.Info.Creator := PDFDocumentInfo.Creator;
      pdf.Info.Author := PDFDocumentInfo.Author;
      pdf.Info.Subject := PDFDocumentInfo.Subject;
      pdf.Info.Title := PDFDocumentInfo.Title;
      pdf.DefaultPageWidth := ConvertX(PaperWidth, Units, mmPoints);
      pdf.DefaultPageHeight := ConvertY(PaperHeight, Units, mmPoints);
      pdf.NewDoc;
      DoProgress(0, TotalPages);
      for PageNo := 1 to TotalPages do
      begin
        case DoPageProcessing(PageNo) of
          pcAccept:
          begin
            pdf.AddPage;
            pdf.Canvas.RenderMetaFile(Pages[PageNo]);
          end;
          pcCancellAll:
            Exit;
        end;
        DoProgress(PageNo, TotalPages);
      end;
      pdf.SaveToFile(FileName);
    finally
      ChangeState(psReady);
    end;
  finally
    pdf.Free;
  end;
{$ELSEIF PDF_DSPDF}
  if dsPDF.Exists then
  begin
    ChangeState(psSavingPDF);
    try
      dsPDF.BeginDoc(AnsiString(FileName));
      try
        dsPDF.SetDocumentInfoEx(PDFDocumentInfo);
        AnyPageRendered := False;
        DoProgress(0, TotalPages);
        for PageNo := 1 to TotalPages do
        begin
          case DoPageProcessing(PageNo) of
            pcAccept:
            begin
              if AnyPageRendered then
                dsPDF.NewPage;
              dsPDF.SetPage(PaperType, Orientation,
                ConvertX(PaperWidth, Units, mmHiMetric),
                ConvertY(PaperHeight, Units, mmHiMetric));
              dsPDF.RenderMetaFile(Pages[PageNo]);
              AnyPageRendered := True;
            end;
            pcCancellAll:
              Exit;
          end;
          DoProgress(PageNo, TotalPages);
        end;
      finally
        dsPDF.EndDoc;
      end;
    finally
      ChangeState(psReady);
    end;
  end
  else
    raise EPDFLibraryError.Create(SdsPDFError);
{$ELSEIF PDF_WPDF}

{$ELSE}
  raise EPDFError.Create(PDFError);
{$IFEND}
end;

function TPrintPreview.CanSaveAsPDF: Boolean;
begin
  {$IFDEF PDF_SYNOPSE}
  Result := True;
  {$ELSEIF PDF_DSPDF}}
  Result := dsPDF.Exists;
  {$ELSEIF PDF_WPDF}}
  Result := true;
  {$ELSE}
  Result := false;
  {$IFEND}
end;

procedure TPrintPreview.Print;
begin
  PrintPages(1, TotalPages);
end;

procedure TPrintPreview.PrintPages(FromPage, ToPage: Integer);
var
  I: Integer;
  Pages: TIntegerList;
begin
  if FromPage < 1 then
    FromPage := 1;
  if ToPage > TotalPages then
    ToPage := TotalPages;
  if FromPage <= TotalPages then
  begin
    Pages := TIntegerList.Create;
    try
      Pages.Capacity := ToPage - FromPage + 1;
      for I := FromPage to ToPage do
        Pages.Add(I);
      PrintPagesEx(Pages);
    finally
      Pages.Free;
    end;
  end;
end;

procedure TPrintPreview.PrintPagesEx(Pages: TIntegerList);
var
  I: Integer;
  PageRect: TRect;
  Succeeded: Boolean;
  AnyPagePrinted: Boolean;
begin
  if (FState = psReady) and PrinterInstalled and (Pages.Count > 0) then
  begin
    ChangeState(psPrinting);
    try
      Succeeded := False;
      InitializePrinting;
      try
        PageRect := PrinterPhysicalPageBounds;
        AnyPagePrinted := False;
        for I := 0 to Pages.Count - 1 do
        begin
          DoProgress(0, Pages.Count);
          case DoPageProcessing(Pages[I]) of
            pcAccept:
            begin
              if AnyPagePrinted then
                Printer.NewPage;
              PrintPage(Pages[I], Printer.Canvas, PageRect);
              AnyPagePrinted := True;
            end;
            pcCancellAll:
              Exit;
          end;
        end;
        DoProgress(Pages.Count, Pages.Count);
        Succeeded := True;
      finally
        FinalizePrinting(Succeeded);
      end;
    finally
      ChangeState(psReady);
    end;
  end;
end;

procedure TPrintPreview.DoProgress(Done, Total: Integer);
begin
  if Assigned(FOnProgress) then
    FOnProgress(Self, Done, Total);
end;

function TPrintPreview.DoPageProcessing(PageNo: Integer): TPageProcessingChoice;
begin
  Result := pcAccept;
  if Assigned(FOnPageProcessing) then
    FOnPageProcessing(Self, PageNo, Result);
end;

procedure TPrintPreview.RegisterThumbnailViewer(ThumbnailView: TThumbnailPreview);
begin
  if ThumbnailView <> nil then
  begin
    if FThumbnailViews = nil then
      FThumbnailViews := TList.Create;
    if FThumbnailViews.IndexOf(ThumbnailView) < 0 then
    begin
      FThumbnailViews.Add(ThumbnailView);
      FreeNotification(ThumbnailView);
    end;
  end;
end;

procedure TPrintPreview.UnregisterThumbnailViewer(ThumbnailView: TThumbnailPreview);
begin
  if FThumbnailViews <> nil then
  begin
    if FThumbnailViews.Remove(ThumbnailView) >= 0 then
    begin
      RemoveFreeNotification(ThumbnailView);
      if FThumbnailViews.Count = 0 then
      begin
        FThumbnailViews.Free;
        FThumbnailViews := nil;
      end;
    end;
  end;
end;

procedure TPrintPreview.RebuildThumbnails;
var
  I: Integer;
begin
  if FThumbnailViews <> nil then
    for I := 0 to FThumbnailViews.Count - 1 do
      TThumbnailPreview(FThumbnailViews[I]).RebuildThumbnails;
end;

procedure TPrintPreview.UpdateThumbnails(StartIndex, EndIndex: Integer);
var
  I: Integer;
begin
  if FThumbnailViews <> nil then
    for I := 0 to FThumbnailViews.Count - 1 do
      TThumbnailPreview(FThumbnailViews[I]).UpdateThumbnails(StartIndex, EndIndex);
end;

procedure TPrintPreview.RepaintThumbnails(StartIndex, EndIndex: Integer);
var
  I: Integer;
begin
  if FThumbnailViews <> nil then
    for I := 0 to FThumbnailViews.Count - 1 do
      TThumbnailPreview(FThumbnailViews[I]).RepaintThumbnails(StartIndex, EndIndex);
end;

procedure TPrintPreview.RecolorThumbnails(OnlyGrays: Boolean);
var
  I: Integer;
  Viewer: TThumbnailPreview;
begin
  if FThumbnailViews <> nil then
    for I := 0 to FThumbnailViews.Count - 1 do
    begin
      Viewer := TThumbnailPreview(FThumbnailViews[I]);
      if not OnlyGrays then
        Viewer.RecolorThumbnails
      else if Viewer.IsGrayscaled then
        Viewer.RepaintThumbnails(0, TotalPages - 1);
    end;
end;

procedure TPrintPreview.SyncThumbnail;
var
  I: Integer;
begin
  if FThumbnailViews <> nil then
    for I := 0 to FThumbnailViews.Count - 1 do
      with TThumbnailPreview(FThumbnailViews[I]) do
      begin
        if CurrentIndex <> CurrentPage - 1 then
          CurrentIndex := CurrentPage - 1
        else
          RepaintThumbnails(CurrentIndex, CurrentIndex);
        Update;
      end;
end;

procedure TPrintPreview.SetAnnotation(Value: Boolean);
begin
  if FAnnotation <> Value then
  begin
    FAnnotation := Value;
    DoAnnotation(FCurrentPage);
    FPaperView.Repaint;
  end;
end;

procedure TPrintPreview.UpdateAnnotation;
begin
  if FAnnotation then
  begin
    DoAnnotation(FCurrentPage);
    FPaperView.Repaint;
  end;
end;

procedure TPrintPreview.DoAnnotation(PageNo: Integer);
var
  AnnotationCanvas: TCanvas;
begin
  if Assigned(AnnotationMetafile) then
  begin
    AnnotationMetafile.Free;
    AnnotationMetafile := nil;
  end;
  if FAnnotation and (PageNo > 0) and Assigned(FOnAnnotation) then
  begin
    CreateMetafileCanvas(AnnotationMetafile, AnnotationCanvas);
    try
      FOnAnnotation(Self, PageNo, AnnotationCanvas);
    finally
      CloseMetafileCanvas(AnnotationMetafile, AnnotationCanvas);
    end;
  end
end;

procedure TPrintPreview.SetBackground(Value: Boolean);
begin
  if FBackground <> Value then
  begin
    FBackground := Value;
    DoBackground(FCurrentPage);
    FPaperView.Repaint;
  end;
end;

procedure TPrintPreview.UpdateBackground;
begin
  if FBackground then
  begin
    DoBackground(FCurrentPage);
    FPaperView.Repaint;
  end;
end;

procedure TPrintPreview.DoBackground(PageNo: Integer);
var
  BackgroundCanvas: TCanvas;
begin
  if Assigned(BackgroundMetafile) then
  begin
    BackgroundMetafile.Free;
    BackgroundMetafile := nil;
  end;
  if FBackground and (PageNo > 0) and Assigned(FOnBackground) then
  begin
    CreateMetafileCanvas(BackgroundMetafile, BackgroundCanvas);
    try
      FOnBackground(Self, PageNo, BackgroundCanvas);
    finally
      CloseMetafileCanvas(BackgroundMetafile, BackgroundCanvas);
    end;
  end
end;

{ TThumbnailDragObject }

constructor TThumbnailDragObject.Create(AControl: TThumbnailPreview;
  APageNo: Integer);
var
  HotSpot: TPoint;
begin
  inherited Create(AControl);
  FPageNo := APageNo;
  if (APageNo <> 0) and (APageNo <= AControl.Items.Count) and
     (AControl.SelCount = 1) and Assigned(AControl.PrintPreview) then
  begin
    // prepare image
    with AControl do
    begin
      Page.Canvas.Pen.Mode := pmCopy;
      Page.Canvas.Brush.Color := PaperView.PaperColor;
      Page.Canvas.Brush.Style := bsSolid;
      Page.Canvas.FillRect(AControl.PageRect);
      PrintPreview.DrawPage(APageNo, Page.Canvas, PageRect, IsGrayscaled);
      // calculate hot spot
      HotSpot := ScreenToClient(Mouse.CursorPos);
      with Items[APageNo-1].Position do
      begin
        Dec(HotSpot.X, X);
        Dec(HotSpot.Y, Y);
      end;
      // set drag image
      FDragImages := TDragImageList.CreateSize(Page.Width, Page.Height);
      FDragImages.AddMasked(Page, Color);
      FDragImages.SetDragImage(0, HotSpot.X, HotSpot.Y);
    end;
  end;
end;

destructor TThumbnailDragObject.Destroy;
begin
  if Assigned(FDragImages) then
  begin
    if FDragImages.Dragging then
      FDragImages.EndDrag;
    FDragImages.Free;
  end;
  inherited Destroy;
end;

function TThumbnailDragObject.GetDragCursor(Accepted: Boolean;
  X, Y: Integer): TCursor;
begin
  if Accepted then
    Result := TThumbnailPreview(Control).DragCursor
  else
    Result := crNoDrop;
end;

function TThumbnailDragObject.GetDragImages: TDragImageList;
begin
  Result := FDragImages;
end;

procedure TThumbnailDragObject.ShowDragImage;
begin
  if Assigned(FDragImages) then
    FDragImages.ShowDragImage;
end;

procedure TThumbnailDragObject.HideDragImage;
begin
  if Assigned(FDragImages) then
    FDragImages.HideDragImage;
end;

{ TThumbnailPreview }

procedure FixControlStyles(Parent: TControl);
var
  I: Integer;
begin
  Parent.ControlStyle := Parent.ControlStyle + [csDisplayDragImage];
  if Parent is TWinControl then
    with TWinControl(Parent) do
      for I := 0 to ControlCount - 1 do
        FixControlStyles(Controls[I]);
end;

constructor TThumbnailPreview.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  if Assigned(AOwner) and (AOwner is TControl) then
    FixControlStyles(TControl(AOwner));
  FZoom := 10;
  FSpacingHorizontal := 8;
  FSpacingVertical := 8;
  FMarkerColor := clBlue;
  FPaperViewOptions := TPaperPreviewOptions.Create;
  FPaperViewOptions.OnChange := PaperViewOptionsChanged;
  Page := TBitmap.Create;
  ParentColor := True;
  ReadOnly := True;
  ViewStyle := vsIcon;
  LargeImages := TImageList.Create(nil);
  Align := alLeft;
end;

destructor TThumbnailPreview.Destroy;
var
  Images: TCustomImageList;
begin
  Images := LargeImages;
  LargeImages := nil;
  Images.Free;
  FPaperViewOptions.Free;
  Page.Free;
  inherited;
end;

procedure TThumbnailPreview.CMFontChanged(var Message: TMessage);
begin
  inherited;
  ApplySpacing;
end;

procedure TThumbnailPreview.CMHintShow(var Message: TCMHintShow);
begin
  inherited;
  if CursorPageNo <> 0 then
  begin
    if PaperView.Hint <> '' then
      Message.HintInfo^.HintStr := PaperView.Hint;
    if Assigned(OnPageInfoTip) then
      FOnPageInfoTip(Self, CursorPageNo, Message.HintInfo^.HintStr);
  end;
end;

procedure TThumbnailPreview.WMSetCursor(var Message: TWMSetCursor);
var
  ActiveCursor: TCursor;
begin
  case MarkerAction of
    maMove:
      if MarkerDragging then
        ActiveCursor := FPaperViewOptions.FGrabCursor
      else
        ActiveCursor := FPaperViewOptions.FDragCursor;
    maResize:
      ActiveCursor := crSizeNWSE;
  else
    if CursorPageNo <> 0 then
      ActiveCursor := FPaperViewOptions.Cursor
    else
      ActiveCursor := Cursor;
  end;
  if ActiveCursor <> crDefault then
  begin
    SetCursor(Screen.Cursors[ActiveCursor]);
    Message.Result := 1;
  end
  else
    inherited;
end;

procedure TThumbnailPreview.WMEraseBkgnd(var Message: TWMEraseBkgnd);
var
  Item: TListItem;
  Org: TPoint;
  CR, IR: TRect;
  SavedDC: Integer;
  I: Integer;
begin
  SavedDC := SaveDC(Message.DC);
  try
    CR := ClientRect;
    Item := GetNearestItem(Point(0, 0), sdAll);
    if Assigned(Item) then
    begin
      Org := ViewOrigin;
      for I := Item.Index to Items.Count - 1 do
      begin
        Item := Items[I];
        IR := BoxRect;
        with Item.DisplayRect(drIcon) do
          OffsetRect(IR,
            (Left + Right - BoxRect.Right) div 2,
            (Top + Bottom - BoxRect.Bottom) div 2);
        ExcludeClipRect(Message.DC, IR.Left, IR.Top, IR.Right, IR.Bottom);
        IR := Item.DisplayRect(drLabel);
        if not IntersectRect(IR, IR, CR) then
          Break;
        ExcludeClipRect(Message.DC, IR.Left, IR.Top, IR.Right, IR.Bottom);
      end;
    end;
    FillRect(Message.DC, CR, Brush.Handle);
  finally
    RestoreDC(Message.DC, SavedDC);
  end;
  Message.Result := 1;
end;

procedure TThumbnailPreview.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (AComponent = PrintPreview) and (Operation = opRemove) then
    PrintPreview := nil;
end;

var SetWindowTheme: function(hwnd: HWND; pszSubAppName: PChar; pszSubIdList: PChar): HRESULT; stdcall;

procedure TThumbnailPreview.CreateWnd;
begin
  inherited CreateWnd;
  if DisableTheme then
  begin
    if not Assigned(SetWindowTheme) then
      @SetWindowTheme := GetProcAddress(GetModuleHandle('UxTheme.dll'), 'SetWindowTheme');
    if Assigned(SetWindowTheme) then
      SetWindowTheme(WindowHandle, nil, '');
  end;
  RebuildThumbnails;
end;

procedure TThumbnailPreview.DestroyWnd;
begin
  FCurrentIndex := -1;
  Items.Clear;
  inherited DestroyWnd;
end;

function TThumbnailPreview.GetPopupMenu: TPopupMenu;
begin
  Result := inherited GetPopupMenu;
  if Assigned(PaperView.PopupMenu) and (PageAtCursor <> 0) then
    Result := PaperView.PopupMenu;
end;

procedure TThumbnailPreview.MouseMove(Shift: TShiftState; X, Y: Integer);
var
  NewCursorPageNo: Integer;
  NewPos: Integer;
  Pt: TPoint;
  R: TRect;
begin
  if MarkerDragging then
  begin
    case MarkerAction of
      maMove:
      begin
        if PrintPreview.CanScrollHorz then
        begin
          NewPos := PrintPreview.HorzScrollBar.Position
                  + MulDiv(X - MarkerPivotPt.X, PrintPreview.Zoom, Zoom);
          if NewPos < 0 then
            NewPos := 0
          else if NewPos > PrintPreview.HorzScrollBar.Range then
            NewPos := PrintPreview.HorzScrollBar.Range;
          PrintPreview.Perform(WM_HSCROLL, MakeLong(SB_THUMBPOSITION, NewPos), 0);
        end;
        if PrintPreview.CanScrollVert then
        begin
          NewPos := PrintPreview.VertScrollBar.Position
                  + MulDiv(Y - MarkerPivotPt.Y, PrintPreview.Zoom, Zoom);
          if NewPos < 0 then
            NewPos := 0
          else if NewPos > PrintPreview.VertScrollBar.Range then
            NewPos := PrintPreview.VertScrollBar.Range;
          PrintPreview.Perform(WM_VSCROLL, MakeLong(SB_THUMBPOSITION, NewPos), 0);
        end;
        MarkerPivotPt := Point(X, Y);
      end;
      maResize:
      begin
        InvalidateMarker(UpdatingMarkerRect);
        UpdatingMarkerRect := MarkerRect;
        Inc(UpdatingMarkerRect.Right, X - MarkerPivotPt.X);
        Inc(UpdatingMarkerRect.Bottom, Y - MarkerPivotPt.Y);
        if UpdatingMarkerRect.Right < UpdatingMarkerRect.Left + 8 then
          UpdatingMarkerRect.Right := UpdatingMarkerRect.Left + 8;
        if UpdatingMarkerRect.Bottom < UpdatingMarkerRect.Top + 8 then
          UpdatingMarkerRect.Bottom := UpdatingMarkerRect.Top + 8;
        IntersectRect(UpdatingMarkerRect, UpdatingMarkerRect, PageRect);
        InvalidateMarker(UpdatingMarkerRect);
        Update;
      end;
    end;
  end
  else
  begin
    NewCursorPageNo := PageAt(X, Y);
    if NewCursorPageNo <> CursorPageNo then
    begin
      CursorPageNo := NewCursorPageNo;
      if ShowHint then
        Application.CancelHint;
    end;
    MarkerAction := maNone;
    if (CursorPageNo <> 0) and (CursorPageNo = PrintPreview.CurrentPage) then
    begin
      Pt := Point(X - MarkerOfs.X, Y - MarkerOfs.Y);
      R.TopLeft := MarkerRect.BottomRight;
      R.BottomRight := MarkerRect.BottomRight;
      InflateRect(R, 4, 4);
      if PtInRect(R, Pt) then
        MarkerAction := maResize
      else if PtInRect(MarkerRect, Pt) and
        (PrintPreview.CanScrollHorz or PrintPreview.CanScrollVert)
      then
        MarkerAction := maMove;
    end;
  end;
  inherited MouseMove(Shift, X, Y);
end;

procedure TThumbnailPreview.MouseDown(Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  inherited MouseDown(Button, Shift, X, Y);
  if not Dragging and not (ssDouble in Shift) and (Button = mbLeft) then
  begin
    if MarkerAction <> maNone then
    begin
      UpdatingMarkerRect := MarkerRect;
      MarkerPivotPt := Point(X, Y);
      MarkerDragging := True;
      SetCapture(Handle);
      Perform(WM_SETCURSOR, Handle, HTCLIENT);
    end
    else if AllowReorder and (SelCount = 1) then
      BeginDrag(False);
  end;
end;

procedure TThumbnailPreview.MouseUp(Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if MarkerDragging then
  begin
    MarkerDragging := False;
    ReleaseCapture;
    Perform(WM_SETCURSOR, Handle, HTCLIENT);
    if not (MarkerAction in [maNone, maMove]) then
    begin
      InvalidateMarker(UpdatingMarkerRect);
      SetMarkerArea(UpdatingMarkerRect);
    end;
  end;
  inherited MouseUp(Button, Shift, X, Y);
end;

procedure TThumbnailPreview.Click;
begin
  inherited Click;
  if (CursorPageNo <> 0) and Assigned(FOnPageClick) then
    FOnPageClick(Self, CursorPageNo);
end;

procedure TThumbnailPreview.DblClick;
begin
  inherited DblClick;
  if (CursorPageNo <> 0) and Assigned(FOnPageDblClick) then
    FOnPageDblClick(Self, CursorPageNo);
end;

function TThumbnailPreview.OwnerDataFetch(Item: TListItem;
  Request: TItemRequest): Boolean;
begin
  if irText in Request then
    Item.Caption := IntToStr(Item.Index + 1);
  Result := True;
end;

function TThumbnailPreview.OwnerDataHint(StartIndex, EndIndex: Integer): Boolean;
var
  I: Integer;
begin
  for I := StartIndex to EndIndex do
    Items[I].Caption := IntToStr(I + 1);
  Result := True;
end;

function TThumbnailPreview.IsCustomDrawn(Target: TCustomDrawTarget;
  Stage: TCustomDrawStage): Boolean;
begin
  Result := (Target = dtItem);
end;

function TThumbnailPreview.CustomDrawItem(Item: TListItem;
  State: TCustomDrawState; Stage: TCustomDrawStage): Boolean;
var
  PageCanvas: TCanvas;
  PageNo: Integer;
  DefaultDraw: Boolean;
  X, Y, W, H: Integer;
  Rect: TRect;
  DC: HDC;
begin
  Result := True;
  if (Stage = cdPrePaint) and (Item <> nil) and
     (Item.Index >= 0) and (Item.Index < Items.Count) then
  begin
    PageNo := Item.Index + 1;
    DefaultDraw := True;
    // prepare thumbnail
    PageCanvas := Page.Canvas;
    PageCanvas.Pen.Mode := pmCopy;
    PageCanvas.Brush.Color := PaperView.PaperColor;
    PageCanvas.Brush.Style := bsSolid;
    PageCanvas.FillRect(PageRect);
    if Assigned(FOnPageBeforeDraw) then
      FOnPageBeforeDraw(Self, PageNo, PageCanvas, PageRect, DefaultDraw);
    if DefaultDraw then
      PrintPreview.DrawPage(PageNo, PageCanvas, PageRect, IsGrayscaled);
    if Assigned(FOnPageAfterDraw) then
      FOnPageAfterDraw(Self, PageNo, PageCanvas, PageRect, DefaultDraw);
    // draw marker on the thumbnail
    if PageNo = PrintPreview.CurrentPage then
    begin
      if not MarkerDragging or (MarkerAction in [maNone, maMove]) then
      begin
        IntersectRect(Rect, PageRect, GetMarkerArea);
        MarkerRect := Rect;
      end
      else
        Rect := UpdatingMarkerRect;
      with PageCanvas, Rect do
      begin
        Pen.Mode := pmCopy;
        Pen.Style := psInsideFrame;
        Pen.Width := 2;
        Pen.Color := MarkerColor;
        Brush.Style := bsClear;
        Rectangle(Left, Top, Right, Bottom);
        Brush.Color := MarkerColor;
        Rect.Left := Rect.Right - 5;
        Rect.Top := Rect.Bottom - 5;
        FillRect(Rect);
      end;
    end;
    // draw thumbnial
    Rect := Item.DisplayRect(drIcon);
    X := (Rect.Left + Rect.Right - BoxRect.Right) div 2;
    Y := (Rect.Top + Rect.Bottom - BoxRect.Bottom) div 2;
    W := Rect.Right - Rect.Left;
    H := Rect.Bottom - Rect.Top;
    DC := GetDC(WindowHandle);
    try
      BitBlt(DC, X, Y, W, H, PageCanvas.Handle, 0, 0, SRCCOPY);
    finally
      ReleaseDC(WindowHandle, DC);
    end;
    MarkerOfs := Rect.TopLeft;
  end;
end;

procedure TThumbnailPreview.Change(Item: TListItem; Change: Integer);
begin
  if Assigned(PrintPreview) and (Change = LVIF_STATE) and Assigned(Item) then
  begin
    if Item.Selected and (ItemIndex >= 0) then
      CurrentIndex := ItemIndex;
    if Item.Selected and Assigned(FOnPageSelect) then
      FOnPageSelect(Self, Item.Index + 1)
    else if not Item.Selected and Assigned(FOnPageUnselect) then
      FOnPageUnselect(Self, Item.Index + 1);
  end;
end;

function TThumbnailPreview.GetSelected: Integer;
begin
  Result := ItemIndex + 1;
end;

procedure TThumbnailPreview.SetSelected(Value: Integer);
begin
  ItemIndex := Value - 1;
end;

procedure TThumbnailPreview.SetDisableTheme(Value: Boolean);
begin
  if FDisableTheme <> Value then
  begin
    FDisableTheme := Value;
    RecreateWnd;
  end;
end;

procedure TThumbnailPreview.DoStartDrag(var DragObject: TDragObject);
begin
  FDropTarget := 0;
  DefaultDragObject := nil;
  if (SelCount = 1) and (DragObject = nil) and Assigned(PrintPreview) then
  begin
    DefaultDragObject := TThumbnailDragObject.Create(Self, Selected);
    DragObject := DefaultDragObject;
  end;
  inherited DoStartDrag(DragObject);
  if Assigned(DefaultDragObject) and (DragObject <> DefaultDragObject) then
  begin
    DefaultDragObject.Free;
    DefaultDragObject := nil;
  end;
end;

procedure TThumbnailPreview.DragOver(Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
begin
  FDropTarget := PageAt(X, Y);
  inherited DragOver(Source, X, Y, State, Accept);
  if Assigned(DefaultDragObject) then
  begin
    InsertMark(DropTarget - 1, DefaultDragObject.DropAfter);
    if AllowReorder and (DropTarget <> 0) and (Source = DefaultDragObject) and (SelCount = 1) then
      Accept := AllowReorder;
  end;
end;

procedure TThumbnailPreview.DragDrop(Source: TObject; X, Y: Integer);
begin
  FDropTarget := PageAt(X, Y);
  inherited DragDrop(Source, X, Y);
  if AllowReorder and Assigned(PrintPreview) and
    (Source = DefaultDragObject) and (SelCount = 1) and (DropTarget <> 0) then
  begin
    if DefaultDragObject.DropAfter and (DropTarget < PrintPreview.TotalPages) then
      PrintPreview.Move(Selected, DropTarget + 1)
    else
      PrintPreview.Move(Selected, DropTarget);
  end;
end;

procedure TThumbnailPreview.DoEndDrag(Target: TObject; X, Y: Integer);
begin
  inherited DoEndDrag(Target, X, Y);
  FDropTarget := 0;
  if Assigned(DefaultDragObject) then
  begin
    InsertMark(-1, False);
    DefaultDragObject.Free;
    DefaultDragObject := nil;
  end;
end;

procedure TThumbnailPreview.InvalidateMarker(Rect: TRect);
begin
  OffsetRect(Rect, MarkerOfs.X, MarkerOfs.Y);
  InvalidateRect(Handle, @Rect, False);
end;

function TThumbnailPreview.GetMarkerArea: TRect;
begin
  Result := PrintPreview.GetVisiblePageRect;
  Result.Left := MulDiv(Result.Left, Zoom, 100);
  Result.Top := MulDiv(Result.Top, Zoom, 100);
  Result.Right := MulDiv(Result.Right, Zoom, 100);
  Result.Bottom := MulDiv(Result.Bottom, Zoom, 100);
  OffsetRect(Result, FPaperViewOptions.BorderWidth, FPaperViewOptions.BorderWidth);
end;

procedure TThumbnailPreview.SetMarkerArea(const Value: TRect);
var
  R: TRect;
begin
  R := Value;
  OffsetRect(R, -FPaperViewOptions.BorderWidth, -FPaperViewOptions.BorderWidth);
  R.Left := MulDiv(R.Left, 100, Zoom);
  R.Top := MulDiv(R.Top, 100, Zoom);
  R.Right := MulDiv(R.Right, 100, Zoom);
  R.Bottom := MulDiv(R.Bottom, 100, Zoom);
  PrintPreview.SetVisiblePageRect(R);
end;

procedure TThumbnailPreview.RebuildThumbnails;
var
  PageWidth, PageHeight: Integer;
begin
  if Assigned(PrintPreview) then
  begin
    SendMessage(WindowHandle, WM_SETREDRAW, 0, 0);
    PageWidth := MulDiv(PrintPreview.PageLogicalPixels.X, Zoom, 100);
    PageHeight := MulDiv(PrintPreview.PageLogicalPixels.Y, Zoom, 100);
    PaperView.CalcDimensions(PageWidth, PageHeight, PageRect, BoxRect);
    Page.Canvas.Pen.Mode := pmCopy;
    Page.Canvas.Brush.Color := Color;
    Page.Canvas.Brush.Style := bsSolid;
    Page.Width := BoxRect.Right;
    Page.Height := BoxRect.Bottom;
    PaperView.Draw(Page.Canvas, BoxRect);
    LargeImages.Width := Page.Width;
    LargeImages.Height := Page.Height;
    ApplySpacing;
    Items.Count := PrintPreview.TotalPages;
    CurrentIndex := PrintPreview.CurrentPage - 1;
    SendMessage(WindowHandle, WM_SETREDRAW, 1, 0);
    Repaint;
  end
end;

procedure TThumbnailPreview.UpdateThumbnails(StartIndex, EndIndex: Integer);
begin
  if Assigned(PrintPreview) then
  begin
    Items.Count := PrintPreview.TotalPages;
    RepaintThumbnails(StartIndex, EndIndex);
    CurrentIndex := PrintPreview.CurrentPage - 1;
  end;
end;

procedure TThumbnailPreview.RepaintThumbnails(StartIndex, EndIndex: Integer);
var
  Item: TListItem;
  CR, IR: TRect;
  I: Integer;
begin
  Item := GetNearestItem(Point(0, 0), sdAll);
  if Assigned(Item) then
  begin
    if StartIndex < Item.Index then
      StartIndex := Item.Index;
    if EndIndex >= Items.Count then
      EndIndex := Items.Count - 1;
    if StartIndex <= EndIndex then
    begin
      CR := ClientRect;
      for I := StartIndex to EndIndex do
      begin
        IR := Items[I].DisplayRect(drIcon);
        if not IntersectRect(IR, IR, CR) then
          Exit;
        InvalidateRect(WindowHandle, @IR, False);
      end;
    end;
  end;
end;

procedure TThumbnailPreview.RecolorThumbnails;
var
  WasGrayscaled: Boolean;
begin
  if Assigned(PrintPreview) then
  begin
    WasGrayscaled := IsGrayscaled;
    FIsGrayscaled := (Grayscale = tgsAlways) or
      ((Grayscale = tgsPreview) and (gsPreview in PrintPreview.Grayscale));
    if WasGrayscaled <> IsGrayscaled then
      RepaintThumbnails(0, Items.Count - 1);
  end;
end;

procedure TThumbnailPreview.ApplySpacing;
const
  LVM_SETICONSPACING = LVM_FIRST + 53;
var
  tm: TTextMetric;
  X, Y: Integer;
begin
  if WindowHandle <> 0 then
  begin
    GetTextMetrics(Canvas.Handle, tm);
    X := SpacingHorizontal + LargeImages.Width;
    Y := SpacingVertical + LargeImages.Height
       + tm.tmHeight + tm.tmAscent - tm.tmDescent - tm.tmInternalLeading;
    SendMessage(WindowHandle, LVM_SETICONSPACING, 0, MakeLong(X, Y));
  end;
end;

procedure TThumbnailPreview.InsertMark(Index: Integer; After: Boolean);
const
  LVM_SETINSERTMARK = LVM_FIRST + 166;
  LVIM_AFTER = $00000001;
type
  LVINSERTMARK = packed record
    cbSize: UINT;
    dwFlags: DWORD;
    iItem: Integer;
    dwReserved: DWORD;
  end;
var
  im: LVINSERTMARK;
begin
  if WindowHandle <> 0 then
  begin
    FillChar(im, SizeOf(im), 0);
    im.cbSize := SizeOf(im);
    if After then im.dwFlags := LVIM_AFTER;
    im.iItem := Index;
    SendMessage(WindowHandle, LVM_SETINSERTMARK, 0, LPARAM(@im));
  end;
end;

procedure TThumbnailPreview.PaperViewOptionsChanged(Sender: TObject;
  Severity: TUpdateSeverity);
begin
  if Assigned(PrintPreview) then
  begin
    if Severity = usRecreate then
      RebuildThumbnails
    else if Severity = usRedraw then
    begin
      PaperView.Draw(Page.Canvas, BoxRect);
      RepaintThumbnails(0, Items.Count - 1);
    end;
  end;
end;

procedure TThumbnailPreview.SetPaperViewOptions(Value: TPaperPreviewOptions);
begin
  FPaperViewOptions.Assign(Value);
end;

procedure TThumbnailPreview.SetMarkerColor(Value: TColor);
begin
  if FMarkerColor <> Value then
  begin
    FMarkerColor := Value;
    if CurrentIndex >= 0 then
      InvalidateMarker(MarkerRect);
  end;
end;

procedure TThumbnailPreview.SetSpacingHorizontal(Value: Integer);
begin
  if FSpacingHorizontal <> Value then
  begin
    FSpacingHorizontal := Value;
    RebuildThumbnails;
  end;
end;

procedure TThumbnailPreview.SetSpacingVertical(Value: Integer);
begin
  if FSpacingVertical <> Value then
  begin
    FSpacingVertical := Value;
    RebuildThumbnails;
  end;
end;

procedure TThumbnailPreview.SetGrayscale(Value: TThumbnailGrayscale);
begin
  if FGrayscale <> Value then
  begin
    FGrayscale := Value;
    RecolorThumbnails;
  end;
end;

procedure TThumbnailPreview.SetZoom(Value: Integer);
begin
  if (FZoom <> Value) and (Value >= 1) then
  begin
    FZoom := Value;
    RebuildThumbnails;
  end;
end;

procedure TThumbnailPreview.SetCurrentIndex(Index: Integer);
var
  OldIndex: Integer;
begin
  if not Assigned(PrintPreview) then
    FCurrentIndex := -1
  else
  begin
    if Index >= Items.Count then
      Index := Items.Count - 1;
    if (CurrentIndex <> Index) and (Index >= 0) then
    begin
      OldIndex := CurrentIndex;
      ItemIndex := Index;
      FCurrentIndex := Index;
      Items[Index].MakeVisible(False);
      if OldIndex < 0 then
        RepaintThumbnails(CurrentIndex, CurrentIndex)
      else if OldIndex - CurrentIndex = 1 then
        RepaintThumbnails(CurrentIndex, OldIndex)
      else if CurrentIndex - OldIndex = 1 then
        RepaintThumbnails(OldIndex, CurrentIndex)
      else
      begin
        RepaintThumbnails(OldIndex, OldIndex);
        RepaintThumbnails(CurrentIndex, CurrentIndex);
      end;
      PrintPreview.CurrentPage := CurrentIndex + 1;
    end
    else if ItemIndex <> CurrentIndex then
      ItemIndex := CurrentIndex;
  end;
end;

procedure TThumbnailPreview.SetPrintPreview(Value: TPrintPreview);
begin
  if FPrintPreview <> Value then
  begin
    if Assigned(FPrintPreview) then
      FPrintPreview.UnregisterThumbnailViewer(Self);
    FPrintPreview := Value;
    if Assigned(FPrintPreview) then
    begin
      FPrintPreview.RegisterThumbnailViewer(Self);
      if Grayscale = tgsPreview then
        FIsGrayscaled := (gsPreview in PrintPreview.Grayscale);
      OwnerData := True;
      RebuildThumbnails;
    end
    else
    begin
      OwnerData := False;
      CurrentIndex := -1;
      if Grayscale = tgsPreview then
        FIsGrayscaled := False;
    end;
  end;
end;

function TThumbnailPreview.PageAt(X, Y: Integer): Integer;
var
  Item: TListItem;
begin
  Item := GetItemAt(X, Y);
  if Assigned(Item) then
    Result := Item.Index + 1
  else
    Result := 0;
end;

function TThumbnailPreview.PageAtCursor: Integer;
begin
  with ScreenToClient(Mouse.CursorPos) do
    Result := PageAt(X, Y);
end;

procedure TThumbnailPreview.GetSelectedPages(Pages: TIntegerList);
var
  I: Integer;
begin
  Pages.Clear;
  if SelCount > 0 then
    for I := ItemIndex to Items.Count - 1 do
      if Items[I].Selected then
        Pages.Add(I + 1);
end;

procedure TThumbnailPreview.SetSelectedPages(Pages: TIntegerList);
var
  I: Integer;
begin
  ClearSelection;
  for I := 0 to Pages.Count - 1 do
    Items[Pages[I]].Selected := True;
end;

procedure TThumbnailPreview.DeleteSelected;
var
  Pages: TIntegerList;
  I: Integer;
begin
  if (SelCount > 0) and Assigned(PrintPreview) then
  begin
    Pages := TIntegerList.Create;
    try
      GetSelectedPages(Pages);
      for I := Pages.Count - 1 downto 0 do
        PrintPreview.Delete(Pages[I]);
    finally
      Pages.Free;
    end;
  end
  else
  begin
    inherited DeleteSelected;
  end;
end;

procedure TThumbnailPreview.PrintSelected;
var
  Pages: TIntegerList;
begin
  if (SelCount > 0) and Assigned(PrintPreview) then
  begin
    Pages := TIntegerList.Create;
    try
      GetSelectedPages(Pages);
      PrintPreview.PrintPagesEx(Pages);
    finally
      Pages.Free;
    end;
  end;
end;

{$IFDEF PDF_DSPDF}

{ TdsPDF }

constructor TdsPDF.Create;
begin
  Handle := LoadLibrary('dspdf.dll');
  if Handle > 0 then
  begin
    @pBeginDoc := GetProcAddress(Handle, 'BeginDoc');
    @pEndDoc := GetProcAddress(Handle, 'EndDoc');
    @pNewPage := GetProcAddress(Handle, 'NewPage');
    @pPrintPageMemory := GetProcAddress(Handle, 'PrintPageM');
    @pPrintPageFile := GetProcAddress(Handle, 'PrintPageF');
    @pSetParameters := GetProcAddress(Handle, 'SetParameters');
    @pSetPage := GetProcAddress(Handle, 'SetPage');
    @pSetDocumentInfo := GetProcAddress(Handle, 'SetDocumentInfo');
  end;
end;

destructor TdsPDF.Destroy;
begin
  if Handle > 0 then
    FreeLibrary(Handle);
  inherited;
end;

function TdsPDF.Exists: Boolean;
begin
  Result := (Handle > 0);
end;

function TdsPDF.PDFPageSizeOf(PaperType: TPaperType): Integer;
begin
  case PaperType of
    pCustom     : Result := 00;
    pLetter     : Result := 01;
    pLegal      : Result := 04;
    pExecutive  : Result := 11;
    pA3         : Result := 03;
    pA4         : Result := 02;
    pA5         : Result := 09;
    pB4         : Result := 08;
    pB5         : Result := 05;
    pFolio      : Result := 10;
    pEnvDL      : Result := 15;
    pEnvB4      : Result := 12;
    pEnvB5      : Result := 13;
    pEnvMonarch : Result := 16;
  else
    Result := 0; // Default to custom
  end;
end;

procedure TdsPDF.SetDocumentInfoEx(Info: TPDFDocumentInfo);
begin
  if Assigned(pSetDocumentInfo) then
    with Info do
    begin
      if Producer <> '' Then SetDocumentInfo(0, Producer);
      if Author <> '' Then SetDocumentInfo(1, Author);
      if Creator <> '' Then SetDocumentInfo(2, Creator);
      if Subject <> '' Then SetDocumentInfo(3, Subject);
      if Title <> '' Then SetDocumentInfo(4, Title);
    end;
end;

function TdsPDF.SetDocumentInfo(What: Integer; const Value: AnsiString): Integer;
begin
  if Assigned(pSetDocumentInfo) then
    Result := pSetDocumentInfo(What, PAnsiChar(Value))
  else
    Result := -1;
end;

function TdsPDF.SetPage(PaperType: TPaperType; Orientation: TPrinterOrientation;
  mmWidth, mmHeight: Integer): Integer;
begin
  if not Assigned(pSetPage) then
    raise EPDFLibraryError.CreateFmt(SdsPDFError, ['SetPage']);
  Result := pSetPage(PDFPageSizeOf(PaperType), Ord(Orientation), mmWidth, mmHeight);
end;

function TdsPDF.SetParameters(OffsetX, OffsetY: Integer;
  const ConverterX, ConverterY: Double): Integer;
begin
  if not Assigned(pSetParameters) then
    raise EPDFLibraryError.CreateFmt(SdsPDFError, ['SetParameters']);
  Result := pSetParameters(OffsetX, OffsetY, ConverterX, ConverterY);
end;

function TdsPDF.BeginDoc(const FileName: AnsiString): Integer;
begin
  if not Assigned(pBeginDoc) then
    raise EPDFLibraryError.CreateFmt(SdsPDFError, ['BeginDoc']);
  Result := pBeginDoc(PAnsiChar(FileName));
end;

function TdsPDF.EndDoc: Integer;
begin
  if not Assigned(pEndDoc) then
    raise EPDFLibraryError.CreateFmt(SdsPDFError, ['EndDoc']);
  Result := pEndDoc();
end;

function TdsPDF.NewPage: Integer;
begin
  if not Assigned(pNewPage) then
    raise EPDFLibraryError.CreateFmt(SdsPDFError, ['NewPage']);
  Result := pNewPage();
end;

function TdsPDF.RenderMemory(Buffer: Pointer; BufferSize: Integer): Integer;
begin
  if not Assigned(pPrintPageMemory) then
    raise EPDFLibraryError.CreateFmt(SdsPDFError, ['PrintPageMemory']);
  Result := pPrintPageMemory(Buffer, BufferSize);
end;

function TdsPDF.RenderFile(const FileName: AnsiString): Integer;
begin
  if not Assigned(pPrintPageFile) then
    raise EPDFLibraryError.CreateFmt(SdsPDFError, ['PrintPageFile']);
  Result := pPrintPageFile(PAnsiChar(FileName));
end;

function TdsPDF.RenderMetaFile(Metafile: TMetafile): Integer;
var
  Stream: TMemoryStream;
begin
  Stream := TMemoryStream.Create;
  try
    Metafile.SaveToStream(Stream);
    Result := RenderMemory(Stream.Memory, Stream.Size);
  finally
    Stream.Free;
  end;
end;
{$ENDIF}

{ TGDIPlusSubset }

type
  TNotificationHookProc = function(out token: ULONG): HRESULT; stdcall;
  TNotificationUnhookProc = procedure(token: ULONG); stdcall;

  PEncoderParameter = ^TEncoderParameter;
  TEncoderParameter = record
    Guid: TGUID;
    NumberOfValues: ULONG;
    Type_: ULONG;
    Value: Pointer;
  end;

  PEncoderParameters = ^TEncoderParameters;
  TEncoderParameters = record
    Count: DWORD;
    Parameter: array[0..0] of TEncoderParameter;
  end;

  PMultiFrameRec = ^TMultiFrameRec;
  TMultiFrameRec = record
    EncoderParameters: TEncoderParameters;
    EncoderValue: ULONG;
    Image: Pointer;
  end;

  PGdiplusStartupInput = ^TGdiplusStartupInput;
  TGdiplusStartupInput = record
    GdiplusVersion: DWORD;
    DebugEventCallback: Pointer;
    SuppressBackgroundThread: BOOL;
    SuppressExternalCodecs: BOOL;
  end;

  PGdiplusStartupOutput = ^TGdiplusStartupOutput;
  TGdiplusStartupOutput = record
    NotificationHook: TNotificationHookProc;
    NotificationUnhook: TNotificationUnhookProc;
  end;


constructor TGDIPlusSubset.Create;
var
  Input: TGDIPlusStartupInput;
  Output: TGdiplusStartupOutput;
begin
  Handle := LoadLibrary('gdiplus.dll');
  if Handle > 0 then
  begin
    @GdiplusStartup := GetProcAddress(Handle, 'GdiplusStartup');
    @GdiplusShutdown := GetProcAddress(Handle, 'GdiplusShutdown');
    @GdipGetDpiX := GetProcAddress(Handle, 'GdipGetDpiX');
    @GdipGetDpiY := GetProcAddress(Handle, 'GdipGetDpiY');
    @GdipDrawImageRectRect := GetProcAddress(Handle, 'GdipDrawImageRectRect');
    @GdipCreateFromHDC := GetProcAddress(Handle, 'GdipCreateFromHDC');
    @GdipGetImageGraphicsContext := GetProcAddress(Handle, 'GdipGetImageGraphicsContext');
    @GdipDeleteGraphics := GetProcAddress(Handle, 'GdipDeleteGraphics');
    @GdipCreateMetafileFromEmf := GetProcAddress(Handle, 'GdipCreateMetafileFromEmf');
    @GdipCreateBitmapFromScan0 := GetProcAddress(Handle, 'GdipCreateBitmapFromScan0');
    @GdipDisposeImage := GetProcAddress(Handle, 'GdipDisposeImage');
    @GdipBitmapSetResolution := GetProcAddress(Handle, 'GdipBitmapSetResolution');
    @GdipGetImageHorizontalResolution := GetProcAddress(Handle, 'GdipGetImageHorizontalResolution');
    @GdipGetImageVerticalResolution := GetProcAddress(Handle, 'GdipGetImageVerticalResolution');
    @GdipGetImageWidth := GetProcAddress(Handle, 'GdipGetImageWidth');
    @GdipGetImageHeight := GetProcAddress(Handle, 'GdipGetImageHeight');
    @GdipGraphicsClear := GetProcAddress(Handle, 'GdipGraphicsClear');
    @GdipGetImageEncodersSize := GetProcAddress(Handle, 'GdipGetImageEncodersSize');
    @GdipGetImageEncoders := GetProcAddress(Handle, 'GdipGetImageEncoders');
    @GdipSaveImageToFile := GetProcAddress(Handle, 'GdipSaveImageToFile');
    @GdipSaveAddImage := GetProcAddress(Handle, 'GdipSaveAddImage');
    // init GDI+
    with Input do
    begin
      GdiplusVersion := 1;
      DebugEventCallback := nil;
      SuppressBackgroundThread := True;
      SuppressExternalCodecs := False;
    end;
    if GdiplusStartup(Token, @Input, @Output) <> S_OK then
      Token := 0
    else if Assigned(Output.NotificationHook) then
    begin
      Output.NotificationHook(ThreadToken);
      TNotificationUnhookProc(pUnhook) := Output.NotificationUnhook;
    end;
  end;
end;

destructor TGDIPlusSubset.Destroy;
begin
  if Handle > 0 then
  begin
    if (ThreadToken <> 0) and Assigned(pUnhook) then
      TNotificationUnhookProc(pUnhook)(ThreadToken);
    if Token <> 0 then
      GdiplusShutdown(Token);
    FreeLibrary(Handle);
  end;
  inherited Destroy;
end;

function TGDIPlusSubset.Exists;
begin
  Result := (Handle > 0) and (Token <> 0);
end;

function TGDIPlusSubset.CteateBitmap(Metafile: TMetafile; BackColor: TColor): Pointer;
const
  OpaqueColor = $FF000000;
  UntiPixels = 2;
  PixelFormatRGB32 = (32 shl 8) or $00020000 or 9;
var
  Graphics, Image: Pointer;
  dpiX, dpiY: Single;
  Width, Height: UINT;
begin
  Result := nil;
  GdipCreateMetafileFromEmf(Metafile.Handle, False, Image);
  try
    GdipGetImageHorizontalResolution(Image, dpiX);
    GdipGetImageVerticalResolution(Image, dpiY);
    GdipGetImageWidth(Image, Width);
    GdipGetImageHeight(Image, Height);
    GdipCreateBitmapFromScan0(Width, Height, 0, PixelFormatRGB32, nil, Result);
    GdipBitmapSetResolution(Result, dpiX, dpiY);
    GdipGetImageGraphicsContext(Result, Graphics);
    try
      GdipGraphicsClear(Graphics, DWORD(ColorToRGB(BackColor)) or OpaqueColor);
      GdipDrawImageRectRect(Graphics, Image, 0, 0, Width, Height,
        0, 0, Width, Height, UntiPixels, nil, nil, nil);
    finally
      GdipDeleteGraphics(Graphics);
    end;
  finally
    gdipDisposeImage(Image);
  end;
end;

procedure TGDIPlusSubset.Draw(Canvas: TCanvas; const Rect: TRect;
  Metafile: TMetafile);
const
  UnitPixels = 2;
var
  DC: HDC;
  gResX, gResY: Single;
  xScale, yScale: Single;
  Graphics, Image: Pointer;
  ImageWidth, ImageHeight: UINT;
begin
  if Exists then
  begin
    DC := Canvas.Handle;
    GdipCreateFromHDC(DC, Graphics);
    try
      GdipGetDpiX(Graphics, gResX);
      GdipGetDpiY(Graphics, gResY);
      xScale := Screen.PixelsPerInch / gResX;
      yScale := Screen.PixelsPerInch / gResY;
      GdipCreateMetafileFromEmf(Metafile.Handle, False, Image);
      try
        GdipGetImageWidth(Image, ImageWidth);
        GdipGetImageHeight(Image, ImageHeight);
        GdipDrawImageRectRect(Graphics, Image, Rect.Left * xScale, Rect.Top * yScale,
          (Rect.Right - Rect.Left) * xScale, (Rect.Bottom - Rect.Top) * yScale, 0, 0,
          ImageWidth, ImageHeight, UnitPixels, nil, nil, nil);
      finally
        GdipDisposeImage(Image);
      end;
    finally
      GdipDeleteGraphics(Graphics);
    end;
  end
  else
    Canvas.StretchDraw(Rect, Metafile);
end;

function TGDIPlusSubset.GetEncoderClsid(const MimeType: WideString;
  out Clsid: TGUID): Boolean;
type
  PImageCodecInfo = ^TImageCodecInfo;
  TImageCodecInfo = packed record
    Clsid             : TGUID;
    FormatID          : TGUID;
    CodecName         : PWideChar;
    DllName           : PWideChar;
    FormatDescription : PWideChar;
    FilenameExtension : PWideChar;
    MimeType          : PWideChar;
    Flags             : DWORD;
    Version           : DWORD;
    SigCount          : DWORD;
    SigSize           : DWORD;
    SigPattern        : PBYTE;
    SigMask           : PBYTE;
  end;
var
  I: Integer;
  NumEncoders, Size: UINT;
  ImageCodecInfoList: PImageCodecInfo;
  ImageCodecInfo: PImageCodecInfo;
begin
  Result := False;
  if Succeeded(GdipGetImageEncodersSize(NumEncoders, Size)) then
  begin
    GetMem(ImageCodecInfoList, Size);
    try
      GdipGetImageEncoders(NumEncoders, Size, ImageCodecInfoList);
      ImageCodecInfo := ImageCodecInfoList;
      for I := 0 to NumEncoders - 1 do
      begin
        if lstrcmpiW(ImageCodecInfo^.MimeType, PWideChar(MimeType)) = 0 then
        begin
          Clsid := ImageCodecInfo^.Clsid;
          Result := True;
          Exit;
        end;
        Inc(ImageCodecInfo);
      end;
    finally
      FreeMem(ImageCodecInfoList, Size);
    end;
  end;
end;

function TGDIPlusSubset.MultiFrameBegin(const FileName: WideString;
  FirstPage: TMetafile; BackColor: TColor): Pointer;
const
  EncoderSaveFlag: TGUID = '{292266FC-AC40-47BF-8CFC-A85B89A655DE}';
  EncoderParameterValueTypeLong = 4;
  EncoderValueMultiFrame = 18;
  EncoderValueFrameDimensionPage = 23;
var
  EncoderClsid: TGUID;
  MF: PMultiFrameRec;
begin
  Result := nil;
  if GetEncoderClsid('image/tiff', EncoderClsid) then
  begin
    GetMem(MF, SizeOf(TMultiFrameRec));
    try
      MF^.Image := CteateBitmap(FirstPage, BackColor);
      MF^.EncoderParameters.Count := 1;
      MF^.EncoderParameters.Parameter[0].Guid := EncoderSaveFlag;
      MF^.EncoderParameters.Parameter[0].Type_ := EncoderParameterValueTypeLong;
      MF^.EncoderParameters.Parameter[0].NumberOfValues := 1;
      MF^.EncoderParameters.Parameter[0].Value := @(MF^.EncoderValue);
      MF^.EncoderValue := EncoderValueMultiFrame;
      if Failed(GdipSaveImageToFile(MF^.Image, PWideChar(FileName),
         EncoderClsid, @(MF^.EncoderParameters))) then
      begin
        RaiseLastOSError;
      end;
      MF^.EncoderValue := EncoderValueFrameDimensionPage;
      Result := MF;
    except
      if MF^.Image <> nil then
        GdipDisposeImage(MF^.Image);
       FreeMem(MF, SizeOf(TMultiFrameRec));
    end;
  end;
end;

procedure TGDIPlusSubset.MultiFrameNext(MF: Pointer;
  NextPage: TMetafile; BackColor: TColor);
var
  Image: Pointer;
begin
  Image := CteateBitmap(NextPage, BackColor);
  try
    GdipSaveAddImage(PMultiFrameRec(MF)^.Image, Image,
      @(PMultiFrameRec(MF)^.EncoderParameters));
  finally
    GdipDisposeImage(Image);
  end;
end;

procedure TGDIPlusSubset.MultiFrameEnd(MF: Pointer);
begin
  if PMultiFrameRec(MF)^.Image <> nil then
    GdipDisposeImage(PMultiFrameRec(MF)^.Image);
  FreeMem(MF, SizeOf(TMultiFrameRec));
end;

{ Components' Registration }

{$IFDEF REGISTER}
procedure Register;
begin
  RegisterComponents('Delphi Area', [TPrintPreview, TThumbnailPreview, TPaperPreview]);
end;
{$ENDIF}

initialization
  Screen.Cursors[crHand] := LoadCursor(hInstance, 'CURSOR_HAND');
  Screen.Cursors[crGrab] := LoadCursor(hInstance, 'CURSOR_GRAB');
finalization
  {$IFDEF PDF_DSPDF}if Assigned(_dsPDF) then _dsPDF.Free;{$ENDIF}
  if Assigned(_gdiPlus) then _gdiPlus.Free;
end.
