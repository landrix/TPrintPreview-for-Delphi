unit Cairo.ExportIntf;
 {$IFDEF CAIRODLLEXPORT}
  {$UNDEF CAIRODLLUSE}
 {$ELSE}
 {$DEFINE CAIRODLLUSE}
 {$ENDIF}
interface
uses
  Winapi.Windows;

 type
  ICairoPDF = interface
  ['{153A5662-EF4C-4B1C-9B6D-4278B7AE0607}']
     procedure CreatePDF(const aFilename : WideString; const aWidth, aHeight: Integer);  stdcall;
     procedure AddPage; stdcall;
     procedure RenderMetaFile(const aMeta : HENHMETAFILE); stdcall;

     procedure SetCreator(const Value: WideString); stdcall;
     procedure SetAuthor(const Value: WideString);  stdcall;
     procedure SetSubject(const Value: WideString); stdcall;
     procedure SetTitle(const Value: WideString);   stdcall;
     procedure SetCreationDate(const Value: TDateTime); stdcall;

     property CreationDate : TDateTime write SetCreationDate;
     property Creator : WideString write SetCreator;
     property Author  : WideString write SetAuthor;
     property Subject : WideString write SetSubject;
     property Title   : WideString write SetTitle;


  end;

  ICairoSVG = interface
     ['{35729AF9-2A37-4C3D-B1C3-DD0DB98794A4}']
     procedure  CreateSVG(const aFilename : WideString; const aWidth, aHeight: Integer); stdcall;
     procedure RenderMetaFile(const aMeta : HENHMETAFILE); stdcall;
  end;

   ICairoExporter = interface
     ['{2663886C-72AA-4F46-97C9-7AE9FA9981B5}']
     function GetCairoPDF : ICairoPDF; stdcall;
     function GetCairoSVG : ICairoSVG; stdcall;
   end;


  {$IFDEF CAIRODLLUSE}
   function CairoExporter : ICairoExporter;  stdcall;
  {$ENDIF}


implementation
{$IFDEF CAIRODLLUSE}
  uses
   System.SysUtils;

Var GCairoExporter : ICairoExporter = nil;
    GDllHandle : THandle = 0;
    GChecked : Boolean = false;
    GCall : function : ICairoExporter;  stdcall;


function CairoExporter : ICairoExporter;
begin
  if not gChecked then
   begin
    gChecked := true;
    GDllHandle :=  SafeLoadLibrary('CairoExport.dll');
    if GDllHandle <> 0 then
    begin
     @GCall := GetProcAddress(GDllHandle, 'CairoExporter') ;
     if @GCall <> nil then
     GCairoExporter := GCall();
    end;
   end;
  result := GCairoExporter;
end;



initialization

finalization
  GCairoExporter := nil;
 {$ENDIF}
end.
