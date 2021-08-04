unit unit_DataRequest;

interface
uses
  System.Classes,   Winapi.Windows, Winapi.Messages, System.SysUtils,
  System.Variants,  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls, Vcl.ExtCtrls;

  type  TDataRequest_ = class (TObject)
    public
      count: integer;
      messag: array [0..20] of byte;
      constructor Create(dat: array of byte; addres:byte);
  end;

    type  TStatisticData = class (TObject)
    public
    minStatisticDateTime:array[0..2] of TDateTime;
    maxStatisticDateTime:array[0..2] of TDateTime;
    minLineVoltege: array [0..2] of Real;
    aveLineVoltege: array [0..2] of Real;
    maxLineVoltege: array [0..2] of Real;
    minPhaseVoltege: array [0..2] of Real;
    avePhaseVoltege: array [0..2] of Real;
    maxPhaseVoltege: array [0..2] of Real;
    minCurrent: array [0..2] of Real;
    aveCurrent: array [0..2] of Real;
    maxCurrent: array [0..2] of Real;
    frequency:  array [0..2] of Real;
     constructor Create();
    end;

  const
    NULL_ACCEPT: Byte = 0;
    ADDRESS_DEVICE: Byte = 1;
    VOLT_CURRENT_FACTOR: Byte = 2;
    VOLTAGE_ONE_PHASE: Byte  = 10;  //U1
    VOLTAGE_TWO_PHASE: Byte  = 11;  //U2
    VOLTAGE_THEE_PHASE: Byte = 12;  //U3
    CURRENT_ONE_PHASE: Byte  = 13;  //I1
    CURRENT_TWO_PHASE: Byte  = 14;  //I2
    CURRENT_THEE_PHASE: Byte = 15;  //I3

    PHASE_ANGLE_ONE: Byte    = 19;  //y1
    PHASE_ANGLE_TWO: Byte    = 20;  //y2
    PHASE_ANGLE_THEE: Byte   = 21;  //y3
    FREQUENCY_ACCEPT: Byte   = 22;  //F
    TEST_CONECTION: Byte     = 99;  //8-1

    testBitWrite:array [0..0] of Byte = (0); //test
    openSessionBitWrite:array [0..7] of Byte = (1, 1, 1, 1, 1, 1, 1, 1); //openSession
    bitWriteInstall_0:array [0..1] of Byte = (8, 5); //adress
    bitWriteInstall_1:array [0..1] of Byte = (8, 2);//transformation coefficient

    request0: array [0..2] of byte = (8, 17, 17);
    request1: array [0..2] of byte = (8, 17, 18);
    request2: array [0..2] of byte = (8, 17, 19);
    request3: array [0..2] of byte = (8, 17, 33);
    request4: array [0..2] of byte = (8, 17, 34);
    request5: array [0..2] of byte = (8, 17, 35);
    request6: array [0..2] of byte = (8, 17, 81);
    request7: array [0..2] of byte = (8, 17, 82);
    request8: array [0..2] of byte = (8, 17, 83);
    request9: array [0..2] of byte = (8, 17, 64);
implementation

 constructor TStatisticData.Create;
 begin

 end;

constructor TDataRequest_.Create;
  var t:integer;
      i, j: Integer;   res:Cardinal;  buf:string;
begin
  count:= Length(dat)+3;
  messag[0]:=addres;
  for t := 0 to Length(dat)-1 do messag[t+1]:= dat[t];
    //CRC-16 MODBus
     res := 65535;
     for i := 0 to count-3 do
      begin
        Res := Res xor (messag[i]);
        for j := 0 to 7 do
          if (Res and $0001) <> 0 then
            Res := (Res shr 1) xor 40961
          else
            Res := Res shr 1;
      end;
     buf:=inttohex(res,0);
     messag[count-2]:= StrToInt('$'+buf[3]+buf[4]);
     messag[count-1]:= StrToInt('$'+buf[1]+buf[2]);
end;

end.
