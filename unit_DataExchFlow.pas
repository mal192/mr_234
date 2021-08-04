unit unit_DataExchFlow;

interface

uses
   Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
   System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
   Vcl.StdCtrls, Registry, ShellApi, unit_DataRequest, unit_SaveAsFlow;


type
  DataExchFlow = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  public

   comAdres:string;
   ByteReaded: cardinal; //количество считанных байт

   procedure StopConnectCom;

  end;

    var
       comHandle: THandle;
       acceptValue: byte;
        comConecting: boolean;
        StatisticData:TStatisticData;
        StatisticOnXls:TStatisticData;

implementation
uses  service_unit, unit_form_install, unit_WriteDataExchFlow, unit_HttpCreate;

 var
  ct: TCommTimeouts;
  dcb: TDCB;
  PSaveAsFlow:SaveAsFlow;
  oneStart:boolean;
  indexXlsMath:integer;
  countXlsMath:integer;

{kall Code ;)) }
procedure StatisticCreate ();
  var forCycle:integer;
  begin
   for forCycle := 0 to 2 do StatisticData.minStatisticDateTime[forCycle]:=now;
   for forCycle := 0 to 2 do StatisticData.maxStatisticDateTime[forCycle]:=now;
   for forCycle := 0 to 2 do StatisticData.minPhaseVoltege[forCycle]:=form_install.instantValueData[forCycle];
   for forCycle := 0 to 2 do StatisticData.avePhaseVoltege[forCycle]:=form_install.instantValueData[forCycle];
   for forCycle := 0 to 2 do StatisticData.maxPhaseVoltege[forCycle]:=form_install.instantValueData[forCycle];
    for forCycle := 0 to 2 do StatisticData.minCurrent[forCycle]:=form_install.instantValueData[forCycle+3];
    for forCycle := 0 to 2 do StatisticData.aveCurrent[forCycle]:=form_install.instantValueData[forCycle+3];
    for forCycle := 0 to 2 do StatisticData.maxCurrent[forCycle]:=form_install.instantValueData[forCycle+3];
   for forCycle := 0 to 2 do StatisticData.minLineVoltege[forCycle]:= form_install.instantValueData[forCycle+6];
   for forCycle := 0 to 2 do StatisticData.maxLineVoltege[forCycle]:= form_install.instantValueData[forCycle+6];
   for forCycle := 0 to 2 do StatisticData.aveLineVoltege[forCycle]:= form_install.instantValueData[forCycle+6];
   for forCycle := 0 to 2 do StatisticData.frequency[forCycle]:= form_install.instantValueData[12];
end;

procedure StatisticXlsCreate ();
  var forCycle:integer;
  begin
   for forCycle := 0 to 2 do StatisticOnXls.minPhaseVoltege[forCycle]:=form_install.instantValueData[forCycle];
   for forCycle := 0 to 2 do StatisticOnXls.avePhaseVoltege[forCycle]:=form_install.instantValueData[forCycle];
   for forCycle := 0 to 2 do StatisticOnXls.maxPhaseVoltege[forCycle]:=form_install.instantValueData[forCycle];
    for forCycle := 0 to 2 do StatisticOnXls.minCurrent[forCycle]:=form_install.instantValueData[forCycle+3];
    for forCycle := 0 to 2 do StatisticOnXls.aveCurrent[forCycle]:=form_install.instantValueData[forCycle+3];
    for forCycle := 0 to 2 do StatisticOnXls.maxCurrent[forCycle]:=form_install.instantValueData[forCycle+3];
   for forCycle := 0 to 2 do StatisticOnXls.minLineVoltege[forCycle]:= form_install.instantValueData[forCycle+6];
   for forCycle := 0 to 2 do StatisticOnXls.maxLineVoltege[forCycle]:= form_install.instantValueData[forCycle+6];
   for forCycle := 0 to 2 do StatisticOnXls.aveLineVoltege[forCycle]:= form_install.instantValueData[forCycle+6];
   for forCycle := 0 to 2 do StatisticOnXls.frequency[forCycle]:= form_install.instantValueData[12];
end;

procedure StaticUpdate ();
  var forCycle:integer;
begin
// частота средняя
StatisticData.frequency[1]:= (StatisticData.frequency[1] + form_install.instantValueData[12])/2;
//частот минимальная
if (StatisticData.frequency[0]>form_install.instantValueData[12]) then
    StatisticData.frequency[0]:= form_install.instantValueData[12];
//частота максимальная
if (StatisticData.frequency[2]<form_install.instantValueData[12]) then
    StatisticData.frequency[2]:= form_install.instantValueData[12];

//мин.фаза а проверяем по фазным
if StatisticData.minPhaseVoltege[0]>form_install.instantValueData[0] then begin
  StatisticData.minStatisticDateTime[0]:= now;
  StatisticData.minLineVoltege[0]:= form_install.instantValueData[6];
  StatisticData.minPhaseVoltege[0]:= form_install.instantValueData[0];
  StatisticData.minCurrent[0]:=form_install.instantValueData[3];
end;
//мин.фаза b проверяем по фазным
if StatisticData.minPhaseVoltege[1]>form_install.instantValueData[1] then begin
  StatisticData.minStatisticDateTime[1]:= now;
  StatisticData.minLineVoltege[1]:= form_install.instantValueData[7];
  StatisticData.minPhaseVoltege[1]:= form_install.instantValueData[1];
  StatisticData.minCurrent[1]:=form_install.instantValueData[4];
end;
//мин.фаза C проверяем по фазным
if StatisticData.minPhaseVoltege[2]>form_install.instantValueData[2] then begin
  StatisticData.minStatisticDateTime[2]:= now;
  StatisticData.minLineVoltege[2]:= form_install.instantValueData[8];
  StatisticData.minPhaseVoltege[2]:= form_install.instantValueData[2];
  StatisticData.minCurrent[2]:=form_install.instantValueData[5];
end;
//макс.фаза а проверяем по фазным
if StatisticData.maxPhaseVoltege[0]<form_install.instantValueData[0] then begin
  StatisticData.maxStatisticDateTime[0]:= now;
  StatisticData.maxLineVoltege[0]:= form_install.instantValueData[6];
  StatisticData.maxPhaseVoltege[0]:= form_install.instantValueData[0];
  StatisticData.maxCurrent[0]:=form_install.instantValueData[3];
end;
//макс.фаза b проверяем по фазным
if StatisticData.maxPhaseVoltege[1]<form_install.instantValueData[1] then begin
  StatisticData.maxStatisticDateTime[1]:= now;
  StatisticData.maxLineVoltege[1]:= form_install.instantValueData[7];
  StatisticData.maxPhaseVoltege[1]:= form_install.instantValueData[1];
  StatisticData.maxCurrent[1]:=form_install.instantValueData[4];
end;
//макс.фаза C проверяем по фазным
if StatisticData.maxPhaseVoltege[2]<form_install.instantValueData[2] then begin
  StatisticData.maxStatisticDateTime[2]:= now;
  StatisticData.maxLineVoltege[2]:= form_install.instantValueData[8];
  StatisticData.maxPhaseVoltege[2]:= form_install.instantValueData[2];
  StatisticData.maxCurrent[2]:=form_install.instantValueData[5];
end;

//срезнач а
  StatisticData.aveLineVoltege[0]:= (form_install.instantValueData[6]+
                                     StatisticData.aveLineVoltege[0])/2;
  StatisticData.avePhaseVoltege[0]:= (form_install.instantValueData[0]+
                                     StatisticData.avePhaseVoltege[0])/2;
  StatisticData.aveCurrent[0]:= (form_install.instantValueData[3]+
                                     StatisticData.aveCurrent[0])/2;
//срезнач b
  StatisticData.aveLineVoltege[1]:= (form_install.instantValueData[7]+
                                     StatisticData.aveLineVoltege[1])/2;
  StatisticData.avePhaseVoltege[1]:= (form_install.instantValueData[1]+
                                     StatisticData.avePhaseVoltege[1])/2;
  StatisticData.aveCurrent[1]:= (form_install.instantValueData[4]+
                                     StatisticData.aveCurrent[1])/2;
//срезнач c
  StatisticData.aveLineVoltege[2]:= (form_install.instantValueData[8]+
                                     StatisticData.aveLineVoltege[2])/2;
  StatisticData.avePhaseVoltege[2]:= (form_install.instantValueData[2]+
                                     StatisticData.avePhaseVoltege[2])/2;
  StatisticData.aveCurrent[2]:= (form_install.instantValueData[5]+
                                     StatisticData.aveCurrent[2])/2;

end;

procedure StaticXlsUpdate ();
  var forCycle:integer;
begin
// частота средняя
StatisticOnXls.frequency[1]:= (StatisticOnXls.frequency[1] + form_install.instantValueData[12])/2;
//частот минимальная
if (StatisticOnXls.frequency[0]>form_install.instantValueData[12]) then
    StatisticOnXls.frequency[0]:= form_install.instantValueData[12];
//частота максимальная
if (StatisticOnXls.frequency[2]<form_install.instantValueData[12]) then
    StatisticOnXls.frequency[2]:= form_install.instantValueData[12];

//мин.фаза а проверяем по фазным
if StatisticOnXls.minPhaseVoltege[0]>form_install.instantValueData[0] then begin
  StatisticOnXls.minLineVoltege[0]:= form_install.instantValueData[6];
  StatisticOnXls.minPhaseVoltege[0]:= form_install.instantValueData[0];
  StatisticOnXls.minCurrent[0]:=form_install.instantValueData[3];
end;
//мин.фаза b проверяем по фазным
if StatisticOnXls.minPhaseVoltege[1]>form_install.instantValueData[1] then begin
  StatisticOnXls.minLineVoltege[1]:= form_install.instantValueData[7];
  StatisticOnXls.minPhaseVoltege[1]:= form_install.instantValueData[1];
  StatisticOnXls.minCurrent[1]:=form_install.instantValueData[4];
end;
//мин.фаза C проверяем по фазным
if StatisticOnXls.minPhaseVoltege[2]>form_install.instantValueData[2] then begin
  StatisticOnXls.minLineVoltege[2]:= form_install.instantValueData[8];
  StatisticOnXls.minPhaseVoltege[2]:= form_install.instantValueData[2];
  StatisticOnXls.minCurrent[2]:=form_install.instantValueData[5];
end;
//макс.фаза а проверяем по фазным
if StatisticOnXls.maxPhaseVoltege[0]<form_install.instantValueData[0] then begin
  StatisticOnXls.maxLineVoltege[0]:= form_install.instantValueData[6];
  StatisticOnXls.maxPhaseVoltege[0]:= form_install.instantValueData[0];
  StatisticOnXls.maxCurrent[0]:=form_install.instantValueData[3];
end;
//макс.фаза b проверяем по фазным
if StatisticOnXls.maxPhaseVoltege[1]<form_install.instantValueData[1] then begin
  StatisticOnXls.maxLineVoltege[1]:= form_install.instantValueData[7];
  StatisticOnXls.maxPhaseVoltege[1]:= form_install.instantValueData[1];
  StatisticOnXls.maxCurrent[1]:=form_install.instantValueData[4];
end;
//макс.фаза C проверяем по фазным
if StatisticOnXls.maxPhaseVoltege[2]<form_install.instantValueData[2] then begin
  StatisticOnXls.maxLineVoltege[2]:= form_install.instantValueData[8];
  StatisticOnXls.maxPhaseVoltege[2]:= form_install.instantValueData[2];
  StatisticOnXls.maxCurrent[2]:=form_install.instantValueData[5];
end;

//срезнач а
  StatisticOnXls.aveLineVoltege[0]:= (form_install.instantValueData[6]+
                                     StatisticOnXls.aveLineVoltege[0])/2;
  StatisticOnXls.avePhaseVoltege[0]:= (form_install.instantValueData[0]+
                                     StatisticOnXls.avePhaseVoltege[0])/2;
  StatisticOnXls.aveCurrent[0]:= (form_install.instantValueData[3]+
                                     StatisticOnXls.aveCurrent[0])/2;
//срезнач b
  StatisticOnXls.aveLineVoltege[1]:= (form_install.instantValueData[7]+
                                     StatisticOnXls.aveLineVoltege[1])/2;
  StatisticOnXls.avePhaseVoltege[1]:= (form_install.instantValueData[1]+
                                     StatisticOnXls.avePhaseVoltege[1])/2;
  StatisticOnXls.aveCurrent[1]:= (form_install.instantValueData[4]+
                                     StatisticOnXls.aveCurrent[1])/2;
//срезнач c
  StatisticOnXls.aveLineVoltege[2]:= (form_install.instantValueData[8]+
                                     StatisticOnXls.aveLineVoltege[2])/2;
  StatisticOnXls.avePhaseVoltege[2]:= (form_install.instantValueData[2]+
                                     StatisticOnXls.avePhaseVoltege[2])/2;
  StatisticOnXls.aveCurrent[2]:= (form_install.instantValueData[5]+
                                     StatisticOnXls.aveCurrent[2])/2;

end;

procedure DataExchFlow.StopConnectCom;
begin
  CloseHandle(comHandle);
end;

function dataTransform(x:byte; y:byte; z:byte):Real;
var
  tempMultiple:integer;
  tempResult:Real;
begin
  tempMultiple:= x div 64;
  tempResult:= (x-(64*tempMultiple))*655.36+y*0.01+z*2.56;
  if x >= 128 then tempResult:=tempResult*-1;
  dataTransform:= tempResult;
end;

procedure DataExchFlow.Execute;
var
  numberComPort:string;
  MyBuff: array[0..255] of byte; //буфер для чтения данных
  zapr_podk:array[0..255] of byte; // буфер для ответа на проверку подключения

  Str: string; //вспомогательная строка
  Status: DWord; //статус устройства
  bufferStr:String; 
  forCycle:integer;
  floatTempData:Real;

begin
countXlsMath:= 30000 div
      (form_install.timerSleepRequests+form_install.timerSleepRequestsNext);
oneStart:=true;

  ByteReaded:=0;
// connect_comport
//messagedlg(comAdres, mterror, [mbOk], 0);
  comHandle := CreateFile(PChar(comAdres), GENERIC_READ or
                              GENERIC_WRITE,
                              FILE_SHARE_READ or FILE_SHARE_WRITE,
                              nil, OPEN_EXISTING,
                              FILE_ATTRIBUTE_NORMAL, 0);
                              
  if (comHandle < 0) or not SetupComm(comHandle, 2048, 2048)or not
    GetCommState(comHandle, dcb) then exit; //init error
   
  dcb.BaudRate:= 9600;
  dcb.ByteSize:= 8;
  if not SetCommState(comHandle, dcb)or not GetCommTimeouts(comHandle, ct)
    then exit; //error
  ct.ReadTotalTimeoutConstant:= 50;
  ct.ReadIntervalTimeout:= 50;
  ct.ReadTotalTimeoutMultiplier:= 1;
  ct.WriteTotalTimeoutMultiplier:= 0;
  ct.WriteTotalTimeoutConstant:= 10;
  if not SetCommTimeouts(comHandle, ct)
    or not SetCommMask(comHandle, EV_RING + EV_RXCHAR + EV_RXFLAG + EV_TXEMPTY)
    then  exit;  //error

  //run FLOW write
    sleep(500);
      comConecting:= true;
      form_install.stopFlow:=true;

  while form_install.stopFlow do begin //reading data
    //получим статус COM-порта устройства
    if not GetCommModemStatus(comHandle, Status) then begin
      {ошибка при получении статуса устройства}
      SysErrorMessage(GetLastError);
      form_install.stopFlow:=false;
      comConecting:=false;
      Exit;
    end;
    //читаем буфер из Com-порта
    FillChar(MyBuff, SizeOf(MyBuff), #0);
    if not ReadFile(comHandle, MyBuff, SizeOf(MyBuff), ByteReaded, nil) then
    begin
      {ошибка при чтении данных}
      SysErrorMessage(GetLastError);
      Exit;
    end; {ошибка при чтении данных}
    //данные пришли
    if ByteReaded > 0 then
    begin {отображение данных}
    if acceptValue <> NULL_ACCEPT then begin
        if acceptValue = TEST_CONECTION then begin
          acceptValue:= NULL_ACCEPT;
          if (MyBuff[1]=0)and(MyBuff[2]=1) then else
              comConecting:=false;
        end;

        if acceptValue = ADDRESS_DEVICE then begin
          acceptValue:= NULL_ACCEPT;
            form_install.addressDev:= MyBuff[2];
            form_install.MemoInfo.Lines.Add('address:');
            form_install.MemoInfo.Lines.Add(IntToStr(form_install.addressDev));
        end;

        if acceptValue = VOLT_CURRENT_FACTOR then begin
          acceptValue:= NULL_ACCEPT;
            form_install.MemoInfo.Lines.Add('коэффициент U:');
            form_install.voltageCoef:=  MyBuff[2];
            form_install.MemoInfo.Lines.Add(IntToStr(form_install.voltageCoef));
            form_install.MemoInfo.Lines.Add('коэффициент I:');
            form_install.currentCoef:=  MyBuff[4];
            form_install.MemoInfo.Lines.Add(IntToStr(form_install.currentCoef));
            form_install.button_next.Enabled:=true;
        end;
        //voltave
        if (acceptValue >=VOLTAGE_ONE_PHASE)and
           (acceptValue <=VOLTAGE_THEE_PHASE) then
        begin
          form_install.instantValueData[acceptValue-10]:=
          dataTransform(MyBuff[1],MyBuff[2],MyBuff[3])*form_install.voltageCoef;
          acceptValue:= NULL_ACCEPT;
        end;
         //current
        if (acceptValue >=CURRENT_ONE_PHASE)and
           (acceptValue <=CURRENT_THEE_PHASE) then
        begin
          form_install.instantValueData[acceptValue-10]:=
          dataTransform(MyBuff[1],MyBuff[2],MyBuff[3])*form_install.currentCoef/10;
          acceptValue:= NULL_ACCEPT;
        end;
         //angle Phase          for forCycle := 9 to 11 do
         if (acceptValue >=PHASE_ANGLE_ONE)and
           (acceptValue <=PHASE_ANGLE_THEE) then
        begin
          floatTempData:= dataTransform(MyBuff[1],MyBuff[2],MyBuff[3]);
          if (floatTempData <-360)or(floatTempData>360) then floatTempData:=0;
          form_install.instantValueData[acceptValue-10]:=floatTempData;
          acceptValue:= NULL_ACCEPT;
        end;
          if (acceptValue =FREQUENCY_ACCEPT) then
        begin
          form_install.instantValueData[acceptValue-10]:=
              dataTransform(MyBuff[1],MyBuff[2],MyBuff[3]);
          acceptValue:= NULL_ACCEPT;

            form_install.instantValueData[6]:=
 Exp(0.5 * ln(Exp(2 * Ln(form_install.instantValueData[0]))+
              Exp(2 * Ln(form_install.instantValueData[1]))-
 form_install.instantValueData[0]*form_install.instantValueData[1]*
              2*(cos(form_install.instantValueData[9]*pi/180))));
            form_install.instantValueData[7]:=
 Exp(0.5 * ln(Exp(2 * Ln(form_install.instantValueData[0]))+
              Exp(2 * Ln(form_install.instantValueData[2]))-
 form_install.instantValueData[0]*form_install.instantValueData[2]*
              2*(cos(form_install.instantValueData[10]*pi/180))));
            form_install.instantValueData[8]:=
 Exp(0.5 * ln(Exp(2 * Ln(form_install.instantValueData[1]))+
              Exp(2 * Ln(form_install.instantValueData[2]))-
 form_install.instantValueData[1]*form_install.instantValueData[2]*
              2*(cos(form_install.instantValueData[11]*pi/180))));

 //создание массива статистика ( а когда и где еще)
   if oneStart then begin
      oneStart:= false; //создание статистики после первого измерения
      StatisticCreate;
      StatisticXlsCreate;
   end else begin
      StaticUpdate;
      StaticXlsUpdate;
      indexXlsMath:=indexXlsMath+1;
   end;

   if indexXlsMath >= countXlsMath then begin
     StatisticXlsCreate;
     indexXlsMath:=0;
   end;

 {}

  if (form_install.startSaveAs = false)and(form_install.saveAs) then begin
          //старт потока записи
            PSaveAsFlow:=SaveAsFlow.Create;
          PSaveAsFlow.FreeOnTerminate:=true;
          PSaveAsFlow.Priority:=tpLower;
          PSaveAsFlow.Resume;
  end;

        end;
      end; //not accept
    end;{отображение данных}
    sleep(10);
  end; //reading data

  { Place thread code here }
end;

end.
