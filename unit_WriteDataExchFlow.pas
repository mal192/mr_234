unit unit_WriteDataExchFlow;

interface

uses
   Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
   System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
   Vcl.StdCtrls, Registry, ShellApi, unit_DataRequest,unit_DataExchFlow,
   unit_SaveAsFlow;

type
  WriteDataExchFlow = class(TThread)
  private
    { Private declarations }
  protected

  public
        procedure Execute; override;
  end;

  var
    bitWriteInstant: array[0..12] of TDataRequest_;


implementation
uses service_unit, unit_form_install;


procedure initialBitWriteInstant();
begin
  bitWriteInstant[0]:=TDataRequest_.Create(request0, form_install.addressDev);
  bitWriteInstant[1]:=TDataRequest_.Create(request1, form_install.addressDev);
  bitWriteInstant[2]:=TDataRequest_.Create(request2, form_install.addressDev);
  bitWriteInstant[3]:=TDataRequest_.Create(request3, form_install.addressDev);
  bitWriteInstant[4]:=TDataRequest_.Create(request4, form_install.addressDev);
  bitWriteInstant[5]:=TDataRequest_.Create(request5, form_install.addressDev);
  bitWriteInstant[9]:=TDataRequest_.Create(request6, form_install.addressDev);
  bitWriteInstant[10]:=TDataRequest_.Create(request7, form_install.addressDev);
  bitWriteInstant[11]:=TDataRequest_.Create(request8, form_install.addressDev);
  bitWriteInstant[12]:=TDataRequest_.Create(request9, form_install.addressDev);
end;


procedure WriteDataExchFlow.Execute;
  var
    DtestBitWrite: TDataRequest_;
    DopenSessionBitWrite: TDataRequest_;
    bitWriteInstall: array[0..1] of TDataRequest_;
    forCycle: Byte;
    i: cardinal;
  label restart;
begin

//opensession

 //messagedlg('',mterror,[],0);
 restart:
if comConecting then begin

acceptValue:=NULL_ACCEPT;
DtestBitWrite:= TDataRequest_.Create(testBitWrite,0);
  WriteFile(comHandle,  DtestBitWrite.messag, DtestBitWrite.count, i, nil);
  sleep(form_install.timerSleepRequests);

acceptValue:= TEST_CONECTION;
DopenSessionBitWrite:= TDataRequest_.Create(openSessionBitWrite,0);
  WriteFile(comHandle, DopenSessionBitWrite.messag, DopenSessionBitWrite.count, i, nil);
  sleep(form_install.timerSleepRequests);

if form_install.systemStart = false then begin //service_install
    form_install.MemoInfo.Lines.Add('секундочку...');

    bitWriteInstall[0]:= TDataRequest_.Create(bitWriteInstall_0,0);
    acceptValue:=ADDRESS_DEVICE;
    bitWriteInstall[0].messag[3]:=182;
    bitWriteInstall[0].messag[4]:=3;
    WriteFile(comHandle,  bitWriteInstall[0].messag, bitWriteInstall[0].count, i, nil);
       sleep(form_install.timerSleepRequests);


    bitWriteInstall[1]:= TDataRequest_.Create(bitWriteInstall_1,0);
    acceptValue:=VOLT_CURRENT_FACTOR;
    WriteFile(comHandle,  bitWriteInstall[1].messag, bitWriteInstall[1].count, i, nil);
  sleep(form_install.timerSleepRequests);
end; // end install

initialBitWriteInstant;

{циклимся на стоп}
if form_install.systemStart then  
   while (form_install.stopFlow and comConecting) do begin
    acceptValue:= TEST_CONECTION;
    DopenSessionBitWrite:= TDataRequest_.Create(openSessionBitWrite,0);
    WriteFile(comHandle, DopenSessionBitWrite.messag, DopenSessionBitWrite.count, i, nil);
    sleep(form_install.timerSleepRequests);

      for forCycle := 0 to 12 do
       // begin  //write instant Value Data
        if (forCycle<6)or(forCycle>8) then begin
          acceptValue:= forCycle+10;
          WriteFile(comHandle,  bitWriteInstant[forCycle].messag,
                                bitWriteInstant[forCycle].count, i, nil);
          sleep(form_install.timerSleepRequests);
        end;


     sleep(form_install.timerSleepRequestsNext);
   end; // FLOW_service_Write
 //service run

 end else  //comConecting
begin
    sleep(form_install.timerSleepRequestsNext);
  GoTo restart;
end;

  { Place thread code here }
end;

end.
