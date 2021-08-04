object mr_service: Tmr_service
  OldCreateOrder = False
  OnCreate = ServiceCreate
  AllowPause = False
  DisplayName = 'mr_234'
  Interactive = True
  StartType = stManual
  OnStart = ServiceStart
  OnStop = ServiceStop
  Height = 150
  Width = 215
  object HTTPServer: TIdHTTPServer
    Bindings = <>
    OnCommandGet = HTTPServerCommandGet
    Left = 120
    Top = 80
  end
  object DataSource1: TDataSource
    DataSet = ADOQuery1
    Left = 72
    Top = 24
  end
  object ADOQuery1: TADOQuery
    Parameters = <>
    Left = 40
    Top = 80
  end
end
