unit unit_SaveAsFlow;

interface

uses
    Winapi.Windows, Winapi.Messages, System.SysUtils, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.SvcMgr, Vcl.Dialogs, IdBaseComponent,
  IdComponent, IdCustomTCPServer, IdCustomHTTPServer, IdHTTPServer, IdContext,
  Forms, WinSvc, CommCtrl, ActiveX, PsAPI,shellapi, registry, unit_DataRequest,
  comobj,ExcelXP;

type
  SaveAsFlow = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

  var
    nameMdb:String;


implementation
uses  service_unit, unit_form_install,unit_WriteDataExchFlow,unit_DataExchFlow;

{ SaveAsFlow }
 procedure ExtractRes(ResType, ResName, ResNewName : String);
var
Res : TResourceStream;
Begin
// ResType - тип ресурса в нашем случае это программа с типом EXEFILE
// ResName - имя ресурса
// ResNewName - соответственно имя нового файла
Res:=TResourceStream.Create(Hinstance, Resname, Pchar(ResType));
Res.SavetoFile(ResNewName);
Res.Free;
end;


function nameFileConvect(timeDat:TDateTime):string;
  var
    tempString:string;
    forCycle:integer;
begin
  tempString:= DateToStr(timeDat)+'_'+TimeToStr(timeDat);
  for forCycle:= 1 to length(tempString) do
    if tempString[forCycle]=':' then  tempString[forCycle]:='-';
    nameFileConvect:=tempString;
end;

procedure newFileMdb();
var   tempTimeData: TDateTime;
begin
  tempTimeData:=now;
  nameMdb:= nameFileConvect(tempTimeData)+'.mdb';
  ExtractRes('EXEFILE', 'LIST',
          form_install.dirFile+   '\'+ nameMdb);
end;

procedure SaveAsFlow.Execute;
var
  indexXlsFile:integer;
  f:System.Text; //класс текстового файла
begin
CoInitialize(nil);
 newFileMdb;
  //подключение по пути form_install.dirFile+   '\'+ nameXls
  mr_service.ADOQuery1.ConnectionString:=
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source="'+form_install.dirFile+   '\'+ nameMdb+'";Persist Security Info=False';
  mr_service.ADOQuery1.Active:=false;
  indexXlsFile:=1;
  form_install.startSaveAs:=true;
   while (form_install.saveAs and form_install.stopFlow) do begin
     //запись -----------------------------
    //open table minimal
    try
    with mr_service.ADOQuery1 do begin
     Active:=false;
     SQL.Clear;
     SQL.Add('SELECT * FROM data');
     Active:=true;
     Last;
     Insert;
     Edit;
     FieldByName('Date').AsDateTime:= Date;    ///?
     FieldByName('Time').AsDateTime:= Time;
     FieldByName('Uab_min').AsFloat:=  StatisticOnXls.minLineVoltege[0];
     FieldByName('Ubc_min').AsFloat:=  StatisticOnXls.minLineVoltege[1];
     FieldByName('Uca_min').AsFloat:=  StatisticOnXls.minLineVoltege[2];
     FieldByName('Ua_min').AsFloat:=  StatisticOnXls.minPhaseVoltege[0];
     FieldByName('Ub_min').AsFloat:=  StatisticOnXls.minPhaseVoltege[1];
     FieldByName('Uc_min').AsFloat:=  StatisticOnXls.minPhaseVoltege[2];
     FieldByName('Ia_min').AsFloat:=  StatisticOnXls.minCurrent[0];
     FieldByName('Ib_min').AsFloat:=  StatisticOnXls.minCurrent[1];
     FieldByName('Ic_min').AsFloat:=  StatisticOnXls.minCurrent[2];
     FieldByName('f_min').AsFloat:=   StatisticOnXls.frequency[0];
     FieldByName('Uab_ave').AsFloat:=  StatisticOnXls.aveLineVoltege[0];
     FieldByName('Ubc_ave').AsFloat:=  StatisticOnXls.aveLineVoltege[1];
     FieldByName('Uca_ave').AsFloat:=  StatisticOnXls.aveLineVoltege[2];
     FieldByName('Ua_ave').AsFloat:=  StatisticOnXls.avePhaseVoltege[0];
     FieldByName('Ub_ave').AsFloat:=  StatisticOnXls.avePhaseVoltege[1];
     FieldByName('Uc_ave').AsFloat:=  StatisticOnXls.avePhaseVoltege[2];
     FieldByName('Ia_ave').AsFloat:=  StatisticOnXls.aveCurrent[0];
     FieldByName('Ib_ave').AsFloat:=  StatisticOnXls.aveCurrent[1];
     FieldByName('Ic_ave').AsFloat:=  StatisticOnXls.aveCurrent[2];
     FieldByName('f_ave').AsFloat:=   StatisticOnXls.frequency[1];
     FieldByName('Uab_max').AsFloat:=  StatisticOnXls.maxLineVoltege[0];
     FieldByName('Ubc_max').AsFloat:=  StatisticOnXls.maxLineVoltege[1];
     FieldByName('Uca_max').AsFloat:=  StatisticOnXls.maxLineVoltege[2];
     FieldByName('Ua_max').AsFloat:=  StatisticOnXls.maxPhaseVoltege[0];
     FieldByName('Ub_max').AsFloat:=  StatisticOnXls.maxPhaseVoltege[1];
     FieldByName('Uc_max').AsFloat:=  StatisticOnXls.maxPhaseVoltege[2];
     FieldByName('Ia_max').AsFloat:=  StatisticOnXls.maxCurrent[0];
     FieldByName('Ib_max').AsFloat:=  StatisticOnXls.maxCurrent[1];
     FieldByName('Ic_max').AsFloat:=  StatisticOnXls.maxCurrent[2];
     FieldByName('f_max').AsFloat:=   StatisticOnXls.frequency[2];
     Post;
     Active:=false;
    end;
     except on E:exception do begin
      form_install.errorStream.Add(E.Message+'___'+E.StackTrace);
      form_install.errorStream.SaveToFile(form_install.dirFile+'\e.txt');
    end;
    end;
    //---------------------------------------
    indexXlsFile:=indexXlsFile+1;
    sleep(30000);
    if indexXlsFile >=3000 then begin
      //new file
     //  закрыть старй файл
      indexXlsFile:=1;
      newFileMdb;
      //подключаемя к новый
  mr_service.ADOQuery1.ConnectionString:=
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source="'+form_install.dirFile+
                             '\'+ nameMdb+'";Persist Security Info=False';
    end;
  end; //Save Cycle
  form_install.startSaveAs:=false;
  { Place thread code here }
end;

end.
