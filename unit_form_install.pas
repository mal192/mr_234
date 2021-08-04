unit unit_form_install;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
   Vcl.StdCtrls, Registry, ShellApi, comobj,
  unit_DataExchFlow, unit_WriteDataExchFlow,unit_SaveAsFlow, Vcl.Menus,
  Data.Win.ADODB, Data.DB, Vcl.Grids, Vcl.DBGrids;

type
  Tform_install = class(TForm)
    button_exit: TButton;
    button_next: TButton;
    MemoInfo: TMemo;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    OpenDialog1: TOpenDialog;
    ADOQuery1: TADOQuery;
    ADODataSet1: TADODataSet;
    procedure FormShow(Sender: TObject);
    procedure button_exitClick(Sender: TObject);
    procedure click_serviceUninstall(Sender: TObject);
    procedure click_serviceInstall(Sender: TObject);
    procedure click_serviceInstallEnd(Sender: TObject);
    procedure click_comboBox(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
  private


    { Private declarations }
  public
    systemStart: boolean;
    voltageCoef: Byte;
    currentCoef: Byte;
    addressDev: Byte;
    timerSleepRequests:integer;
    timerSleepRequestsNext:integer;
    stopFlow:boolean;
    instantValueData: array [0..12] of real; // rezult data
    saveAs:boolean;
    startSaveAs:boolean;
    dirFile:string;
    errorStream:TStringList;
    { Public declarations }
  end;
const
nam:string = 'Ма';
nam1:string = 'в А. '+#28+'В.';
tel:integer = 791914;
var
  form_install: Tform_install;
  service_installed:boolean;
  listComPort:TComboBox;
  pDataExchFlow:DataExchFlow;
  pWriteDataExchFlow:WriteDataExchFlow;

implementation
uses  service_unit,  unit_DataRequest,  unit_HttpCreate;


{$R *.dfm}

 procedure ExtractRes(ResType, ResName, ResNewName : String);
var
Res : TResourceStream;
Begin
  Res:=TResourceStream.Create(Hinstance, Resname, Pchar(ResType));
  Res.SavetoFile(ResNewName);
  Res.Free;
end;


function list_ComPort():TStringList;
var
  registry: TRegistry;
  tempList: TStringList;
  listCom:  TStringList;
  forCycle: Byte;
begin
  list_ComPort:= TStringList.Create;
  tempList:= TStringList.Create;
  listCom:= TStringList.Create;
    registry:=Tregistry.Create;
    registry.RootKey:=HKEY_LOCAL_MACHINE;
    registry.OpenKey('HARDWARE\DEVICEMAP\SERIALCOMM',false);
    registry.GetValueNames(tempList);
    registry.CloseKey;
    for forCycle:=0 to tempList.Count-1 do begin
      registry:=Tregistry.Create;
      registry.RootKey:=HKEY_LOCAL_MACHINE;
      registry.OpenKey('HARDWARE\DEVICEMAP\SERIALCOMM',true);
      if registry.ReadString(tempList[forCycle])<>'' then
        listCom.Add(registry.ReadString(tempList[forCycle]));
      registry.CloseKey;
      registry.Free;
    end;
  list_ComPort.AddStrings(listCom);
  tempList.Clear;
  tempList.Free;
  listCom.Clear;
  listCom.Free;
end;

procedure Tform_install.click_comboBox(Sender: TObject);
begin
  pDataExchFlow:= DataExchFlow.Create;
  pDataExchFlow.FreeOnTerminate:=true;
  pDataExchFlow.Priority:=tpLower;
   pDataExchFlow.comAdres:=listComPort.Text;
//  pDataExchFlow.Synchronize(pWriteDataExchFlow, pWriteDataExchFlow.Execute);
 pDataExchFlow.Resume;

    pWriteDataExchFlow:= WriteDataExchFlow.Create;
    pWriteDataExchFlow.FreeOnTerminate:=true;
    pWriteDataExchFlow.Priority:=tpLower;
//    pWriteDataExchFlow.Synchronize(pDataExchFlow,pDataExchFlow.Execute);
   pWriteDataExchFlow.Resume;
end;

procedure Tform_install.click_serviceInstallEnd(Sender: TObject);
var
  registry: TRegistry;
  h:hwnd;
begin
registry := TRegistry.Create;
 registry.RootKey:=HKEY_LOCAL_MACHINE;
 registry.OpenKey('SYSTEM\CurrentControlSet\services\mr_service',true);
    registry.WriteInteger('addressDev', 0);
    registry.WriteInteger('voltageCoef', voltageCoef);
    registry.WriteInteger('currentCoef', currentCoef);
    registry.WriteInteger('timerSleepRequests', 700);
    registry.WriteInteger('timerSleepRequestsNext', 700);
    registry.WriteString('comAdres',pDataExchFlow.comAdres);
    registry.WriteBool('SaveAs',false);
    registry.WriteString('Description','Служба мониторинга параметров '+
                            'электрической сети, электросчетчика Меркурий 234');
 registry.CloseKey;
 registry.Free;
  sleep(150);
  shellexecute(h,nil,pchar(application.ExeName),pchar('/install'),'',sw_show);
  application.Terminate;
end;

procedure Tform_install.click_serviceUninstall(Sender: TObject);
var h:hwnd;
begin
  sleep(500);
  shellexecute(h,nil,pchar(application.ExeName),pchar('/uninstall'),'',sw_show);
  application.Terminate;
end;

procedure Tform_install.click_serviceInstall(Sender: TObject);
begin
  MemoInfo.Lines.Clear;
  listComPort:=Tcombobox.Create(form_install);
  listComPort.Parent:= Self;
  listComPort.Left:= MemoInfo.Left;
  listComPort.Top:=  MemoInfo.Top;
  listComPort.Font:= MemoInfo.Font;
  listComPort.Width:= 153;
  listComPort.Height:=24;
  listComPort.Items.Clear;
  listComPort.Items.AddStrings(list_ComPort);
  listComPort.OnClick:= click_comboBox;
  listComPort.Show;
    MemoInfo.Top:= MemoInfo.Top + listComPort.Height + 3;
    MemoInfo.Height:= MemoInfo.Height - listComPort.Height;
      button_next.Caption:= 'Установить';
      button_next.Enabled:= false;
      button_next.OnClick:= click_serviceInstallEnd;


end;

procedure Tform_install.button_exitClick(Sender: TObject);
begin
application.Terminate;
end;

procedure Tform_install.FormShow(Sender: TObject);
var registry: Tregistry;
begin


//initioal_service
  timerSleepRequests:=550;
  service_installed:=false;
  registry := TRegistry.Create;
  registry.RootKey := HKEY_LOCAL_MACHINE;
  if registry.OpenKey('SYSTEM\CurrentControlSet\services\mr_service',false)=true
    then begin
      if registry.ReadString('DisplayName') = 'mr_234' then
        service_installed:=true
      else
        service_installed:=false;
  end;
  registry.CloseKey;
  registry.Free;

  if service_installed=true then begin
  //delete_service
    form_install.Caption:='uninstall';
    form_install.MemoInfo.Text:='Вы действительно хотите удалить службу? ';
   button_next.OnClick:= click_serviceUninstall;

  end    else begin
  //install_service
    form_install.Caption:='install';
    form_install.MemoInfo.Text:=
    'Проверьте, подключен ли счетчик к ПК'+#13+#10+
    ' ВНИМАНИЕ!!!'+#13+#10+'Данный файл будет прописан как служба Windows.'+
    #13+#10+'Убедитесь, что его расположение окончательное. ';
    button_next.OnClick:= click_serviceInstall;
  end;

end;

procedure Tform_install.N1Click(Sender: TObject);
begin
messagedlg('Приложение для мониторинга и сохранения показателей '+#13+#10+
  'напряжений и токов в сети 6кВ.'+#13+#10+'Показатели считываются из счетчика электрической сети Меркурий-234.'+
  #13+#10+ 'По всем вопросам +'+inttostr(tel)+'8'+#28+'9'+#28+'44'+#28+'7'+'.'+#13+#10+'                            '+
  nam+#28+#1083+#1100+#1094+#28+#1077+#28+nam1
  ,mtinformation,[mbOk],0);

end;

procedure Tform_install.N2Click(Sender: TObject);
var
  filePath:string;
  fileName:string;
  forCycle:integer;
    xlsFile:olevariant;
    s:string;
    xlsConect:variant;
begin
if opendialog1.Execute then begin
  fileName:= opendialog1.FileName;
  filePath:= extractfilepath(fileName);
  fileName:= ExtractFileName(fileName);
  xlsFile:= filePath+    fileName+'.xlsx';
  ExtractRes('EXEFILE', 'XLS', xlsFile);
  ADOQuery1.ConnectionString:=
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source="'
      +filePath+ fileName+'";Persist Security Info=False';
  ADOQuery1.Active:=false;
  ADOQuery1.SQL.Clear;
  ADOQuery1.SQL.Add('SELECT * FROM data ORDER BY Код ASC;');
  ADOQuery1.Active:=true;
  ADOQuery1.First;

  xlsConect:= createOleObject('Excel.Application');
  xlsConect.displayalerts:= false;
  xlsConect.visible:=true;
  xlsConect.Workbooks.Open(xlsFile);
  for forCycle := 1 to  ADOQuery1.RecordCount do
  begin
   xlsConect.Cells[forCycle+2,1]:=  ADOQuery1.FieldByName('Код').AsInteger;
   xlsConect.Cells[forCycle+2,2]:=  ADOQuery1.FieldByName('Date').AsDateTime;
   xlsConect.Cells[forCycle+2,3]:=  TimeToStr(ADOQuery1.FieldByName('Time').AsDateTime);
   xlsConect.Cells[forCycle+2,4]:=  ADOQuery1.FieldByName('Uab_min').AsFloat;
   xlsConect.Cells[forCycle+2,5]:=  ADOQuery1.FieldByName('Ubc_min').AsFloat;
   xlsConect.Cells[forCycle+2,6]:=  ADOQuery1.FieldByName('Uca_min').AsFloat;
   xlsConect.Cells[forCycle+2,7]:=  ADOQuery1.FieldByName('Ua_min').AsFloat;
   xlsConect.Cells[forCycle+2,8]:=  ADOQuery1.FieldByName('Ub_min').AsFloat;
   xlsConect.Cells[forCycle+2,9]:=  ADOQuery1.FieldByName('Uc_min').AsFloat;
   xlsConect.Cells[forCycle+2,10]:=  ADOQuery1.FieldByName('Ia_min').AsFloat;
   xlsConect.Cells[forCycle+2,11]:= ADOQuery1.FieldByName('Ib_min').AsFloat;
   xlsConect.Cells[forCycle+2,12]:= ADOQuery1.FieldByName('Ic_min').AsFloat;
   xlsConect.Cells[forCycle+2,13]:= ADOQuery1.FieldByName('f_min').AsFloat;
   xlsConect.Cells[forCycle+2,14]:= ADOQuery1.FieldByName('Uab_ave').AsFloat;
   xlsConect.Cells[forCycle+2,15]:= ADOQuery1.FieldByName('Ubc_ave').AsFloat;
   xlsConect.Cells[forCycle+2,16]:= ADOQuery1.FieldByName('Uca_ave').AsFloat;
   xlsConect.Cells[forCycle+2,17]:= ADOQuery1.FieldByName('Ua_ave').AsFloat;
   xlsConect.Cells[forCycle+2,18]:= ADOQuery1.FieldByName('Ub_ave').AsFloat;
   xlsConect.Cells[forCycle+2,19]:= ADOQuery1.FieldByName('Uc_ave').AsFloat;
   xlsConect.Cells[forCycle+2,20]:= ADOQuery1.FieldByName('Ia_ave').AsFloat;
   xlsConect.Cells[forCycle+2,21]:= ADOQuery1.FieldByName('Ib_ave').AsFloat;
   xlsConect.Cells[forCycle+2,22]:= ADOQuery1.FieldByName('Ic_ave').AsFloat;
   xlsConect.Cells[forCycle+2,23]:= ADOQuery1.FieldByName('f_ave').AsFloat;
   xlsConect.Cells[forCycle+2,24]:= ADOQuery1.FieldByName('Uab_max').AsFloat;
   xlsConect.Cells[forCycle+2,25]:= ADOQuery1.FieldByName('Ubc_max').AsFloat;
   xlsConect.Cells[forCycle+2,26]:= ADOQuery1.FieldByName('Uca_max').AsFloat;
   xlsConect.Cells[forCycle+2,27]:= ADOQuery1.FieldByName('Ua_max').AsFloat;
   xlsConect.Cells[forCycle+2,28]:= ADOQuery1.FieldByName('Ub_max').AsFloat;
   xlsConect.Cells[forCycle+2,29]:= ADOQuery1.FieldByName('Uc_max').AsFloat;
   xlsConect.Cells[forCycle+2,30]:= ADOQuery1.FieldByName('Ia_max').AsFloat;
   xlsConect.Cells[forCycle+2,31]:= ADOQuery1.FieldByName('Ib_max').AsFloat;
   xlsConect.Cells[forCycle+2,32]:= ADOQuery1.FieldByName('Ic_max').AsFloat;
   xlsConect.Cells[forCycle+2,33]:= ADOQuery1.FieldByName('f_max').AsFloat;
   ADOQuery1.Next;
  end;
     xlsConect.Save;
//  xlsConect.Quit;
 // ADOQuery1.Active:=false;
  MessageDlg('Отчет готов',mtinformation, [mbOk], 0);


  //ShowMessage();
//  ExtractRes

end;

end;

end.
