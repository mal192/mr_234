program mr_234;

uses
  Vcl.SvcMgr,
  service_unit in 'service_unit.pas' {mr_service: TService},
  unit_form_install in 'unit_form_install.pas' {form_install},
  unit_DataRequest in 'unit_DataRequest.pas',
  unit_WriteDataExchFlow in 'unit_WriteDataExchFlow.pas',
  unit_HttpCreate in 'unit_HttpCreate.pas',
  unit_SaveAsFlow in 'unit_SaveAsFlow.pas';

{$R *.RES}

begin
  // Windows 2003 Server requires StartServiceCtrlDispatcher to be
  // called before CoRegisterClassObject, which can be called indirectly
  // by Application.Initialize. TServiceApplication.DelayInitialize allows
  // Application.Initialize to be called from TService.Main (after
  // StartServiceCtrlDispatcher has been called).
  //
  // Delayed initialization of the Application object may affect
  // events which then occur prior to initialization, such as
  // TService.OnCreate. It is only recommended if the ServiceApplication
  // registers a class object with OLE and is intended for use with
  // Windows 2003 Server.
  //
   Application.DelayInitialize := True;
  //
  if not Application.DelayInitialize or Application.Installing then
    Application.Initialize;
      Application.CreateForm(Tform_install, form_install);
  Application.CreateForm(Tmr_service, mr_service);
  Application.Run;
end.
