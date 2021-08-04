unit service_unit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.SvcMgr, Vcl.Dialogs, IdBaseComponent,
  IdComponent, IdCustomTCPServer, IdCustomHTTPServer, IdHTTPServer, IdContext,
  Forms, WinSvc, CommCtrl, ActiveX, PsAPI,shellapi, registry,
  unit_WriteDataExchFlow, unit_DataExchFlow,unit_SaveAsFlow, unit_HttpCreate,
   unit_DataRequest, ExcelXP, Data.DB, Data.Win.ADODB ;




type
  Tmr_service = class(TService)
    HTTPServer: TIdHTTPServer;
    DataSource1: TDataSource;
    ADOQuery1: TADOQuery;
    procedure ServiceCreate(Sender: TObject);
    procedure ServiceStart(Sender: TService; var Started: Boolean);
    procedure ServiceStop(Sender: TService; var Stopped: Boolean);
    procedure HTTPServerCommandGet(AContext: TIdContext;
      ARequestInfo: TIdHTTPRequestInfo; AResponseInfo: TIdHTTPResponseInfo);
  private
    { Private declarations }
  public


    function GetServiceController: TServiceController; override;

    { Public declarations }
  end;

var
  mr_service: Tmr_service;
  SDataExchFlow:DataExchFlow;
  SWriteDataExchFlow:WriteDataExchFlow;
  DHttpStreamRoot:THttpStreamRoot;
  ErrorStringList:TStringList;
  ExcelConnect: TExcelWorkbook;



implementation


uses unit_form_install;

{$R *.DFM}

{$R List.res}




function initioal_system_privilege():boolean;
var registry:Tregistry;
begin
  registry := TRegistry.Create;
  registry.RootKey := HKEY_LOCAL_MACHINE;
  if registry.OpenKey('SAM\SAM\Domains',false)=true
    then
      initioal_system_privilege:=true
      else
       initioal_system_privilege:=false;
  registry.CloseKey;
//  initioal_system_privilege:=true; //                   :))
end;

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  mr_service.Controller(CtrlCode);
end;

function Tmr_service.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

procedure Tmr_service.HTTPServerCommandGet(AContext: TIdContext;
  ARequestInfo: TIdHTTPRequestInfo; AResponseInfo: TIdHTTPResponseInfo);
  var
    forCycle:integer;
    i:string;
    FileSave:TStringList;
      B: TBytes;
begin
 {if (ARequestInfo.AuthUsername <> ' ') or (ARequestInfo.AuthPassword <> ' ') then
  begin
    AResponseInfo.ContentText :='<i>ERROR!</i>';
    AResponseInfo.AuthRealm :='autontification';
    exit;
  end; }
  i:= ARequestInfo.Document;
    AResponseInfo.ContentType := 'text/html; charset=windows-1251';
  if i='/' then
   if   form_install.stopFlow=true then begin
       AResponseInfo.ContentText:= DHttpStreamRoot.uplodStreamRoot(form_install.instantValueData)
   end else
     AResponseInfo.ContentText:=ErrorStringList.Text
   else
  if i = '/statistic' then begin

   if ARequestInfo.Params.Count <> 0 then
    if form_install.saveAs = false then  form_install.saveAs:= true else
          form_install.saveAs:= false;

         AResponseInfo.ContentText:=DHttpStreamRoot.uplodStreamStatistic();
          //  DHttpStreamRoot.uplodStreamStatistic();
  end else  //statistic

 if i='/error' then
        AResponseInfo.ContentText:= form_install.errorStream.Text
 else
 begin
   //file
     AResponseInfo.ServeFile(AContext,form_install.dirFile+i);
 end;
 //StaticticValueData



end;

procedure Tmr_service.ServiceCreate(Sender: TObject);
  var forCycle:byte;
begin
form_install.errorStream:=TStringList.Create;
StatisticData:= TStatisticData.Create;
StatisticOnXls:=TStatisticData.Create;
form_install.dirFile:=extractfilepath(application.ExeName);
form_install.errorStream.Add('Œ¯Ë·ÍË');


DHttpStreamRoot:=THttpStreamRoot.Create;
HTTPServer.DefaultPort:=100;


ErrorStringList:=TStringList.Create;
ErrorStringList.Add('Error connecting to the Com Port');
ErrorStringList.Add('Check the connection to the Device');
ErrorStringList.Add('restart the service "mr-234"');

if (ParamStr(1) = '') then
  //install_function
    if initioal_system_privilege = false then begin
      application.Run;
      form_install.systemStart:= false;
      form_install.Show;
    end else
      form_install.systemStart:=true;
end;

procedure Tmr_service.ServiceStart(Sender: TService; var Started: Boolean);
var
  registry:TRegistry;
begin

  HTTPServer.Active:=true;
  registry:= Tregistry.Create;
  registry.RootKey:=HKEY_LOCAL_MACHINE;
  registry.OpenKey('SYSTEM\CurrentControlSet\services\mr_service',true);
    form_install.addressDev:= registry.ReadInteger('addressDev');
    form_install.voltageCoef:= registry.ReadInteger('voltageCoef');
    form_install.currentCoef:= registry.ReadInteger('currentCoef');
    form_install.timerSleepRequests:= registry.ReadInteger('timerSleepRequests');
    form_install.timerSleepRequestsNext:= registry.ReadInteger('timerSleepRequestsNext');
    form_install.saveAs:= registry.ReadBool('saveAs');
    SDataExchFlow:=DataExchFlow.Create;
    SDataExchFlow.comAdres:= registry.ReadString('comAdres');

 registry.CloseKey;
 registry.Free;

 //service_start_Flow
   SDataExchFlow.FreeOnTerminate:=true;
   SDataExchFlow.Priority:=tpLower;
   SDataExchFlow.Resume;

    SWriteDataExchFlow:= WriteDataExchFlow.Create;
    SWriteDataExchFlow.FreeOnTerminate:=true;
    SWriteDataExchFlow.Priority:=tpLower;
    SWriteDataExchFlow.Resume;

end;

procedure Tmr_service.ServiceStop(Sender: TService; var Stopped: Boolean);
begin
  HTTPServer.Active:=false;
  form_install.stopFlow:=false;
  //SDataExchFlow.comConecting:=false;
  SDataExchFlow.StopConnectCom;
  SDataExchFlow.Terminate;
  SWriteDataExchFlow.Terminate;
  HTTPServer.Active:=false;
  HTTPServer.Destroy;
  HTTPServer.Free;
//   pDataExchFlow.Free;
//   pWriteDataExchFlow.Free;
end;

end.
