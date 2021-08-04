unit unit_HttpCreate;

interface
uses
  System.Classes,   Winapi.Windows, Winapi.Messages, System.SysUtils,
  System.Variants,  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls, Vcl.ExtCtrls, unit_DataExchFlow, unit_DataRequest,
  unit_form_install;

    type  THttpStreamRoot = class (TObject)
    public
      HttpStreamRootData: array [0..13] of TStringList;
      HttpStatisticData: array [0..36] of TStringList;
      function uplodStreamRoot(instantValue: array of Real):string;
      function uplodStreamStatistic():string;
      constructor Create();
    end;

implementation


function THttpStreamRoot.uplodStreamRoot(instantValue: array of Real):string;
var
  forCycle:integer;
  uplodStreamRootTemp:TStringList;
begin
  uplodStreamRootTemp:=TStringList.Create;

  uplodStreamRootTemp.Clear;
  for forCycle := 0 to 12 do begin
    uplodStreamRootTemp.AddStrings(HttpStreamRootData[forCycle]);
  uplodStreamRootTemp.Add(FloatToStrf(instantValue[forCycle],ffFixed,8,2));
  end;
    uplodStreamRootTemp.AddStrings(HttpStreamRootData[13]);

    uplodStreamRoot:=(uplodStreamRootTemp.Text);
    uplodStreamRootTemp.Clear;
    uplodStreamRootTemp.Free;
end;

function THttpStreamRoot.uplodStreamStatistic():String;
var
  forCycle:integer;
  uplodTemp:TStringList;
  data:array [0..35]of string;
  direkt:TSearchRec;

begin
//Data_min 0, 4, 8
for forCycle := 0 to 2 do
  data[forCycle*4]:= DateToStr(StatisticData.minStatisticDateTime[forCycle])+' '+
                     TimeToStr(StatisticData.minStatisticDateTime[forCycle]);
for forCycle := 0 to 2 do
  data[forCycle*4+1]:= FloatToStrf(StatisticData.minLineVoltege[forCycle],ffFixed,8,2);
for forCycle := 0 to 2 do
  data[forCycle*4+2]:= FloatToStrf(StatisticData.minPhaseVoltege[forCycle],ffFixed,8,2);
for forCycle := 0 to 2 do
  data[forCycle*4+3]:= FloatToStrf(StatisticData.minCurrent[forCycle],ffFixed,8,2);
 data[12]:= FloatToStrf(StatisticData.frequency[0],ffFixed,8,2);

 // average
for forCycle := 0 to 2 do
  data[forCycle*3+13]:= FloatToStrf(StatisticData.aveLineVoltege[forCycle],ffFixed,8,2);
for forCycle := 0 to 2 do
  data[forCycle*3+14]:= FloatToStrf(StatisticData.avePhaseVoltege[forCycle],ffFixed,8,2);
for forCycle := 0 to 2 do
  data[forCycle*3+15]:= FloatToStrf(StatisticData.aveCurrent[forCycle],ffFixed,8,2);
  data[22]:= FloatToStrf(StatisticData.frequency[1],ffFixed,8,2);
//max
for forCycle := 0 to 2 do
  data[forCycle*4+23]:= DateToStr(StatisticData.maxStatisticDateTime[forCycle])+' '+
                     TimeToStr(StatisticData.maxStatisticDateTime[forCycle]);
for forCycle := 0 to 2 do
  data[forCycle*4+24]:= FloatToStrf(StatisticData.maxLineVoltege[forCycle],ffFixed,8,2);
for forCycle := 0 to 2 do
  data[forCycle*4+25]:= FloatToStrf(StatisticData.maxPhaseVoltege[forCycle],ffFixed,8,2);
for forCycle := 0 to 2 do
  data[forCycle*4+26]:= FloatToStrf(StatisticData.maxCurrent[forCycle],ffFixed,8,2);
data[35]:= FloatToStrf(StatisticData.frequency[2],ffFixed,8,2);


  uplodTemp:=TStringList.Create;
  uplodTemp.Clear;
 uplodTemp.Add('<html>');
 uplodTemp.Add('<meta charset="UTF-8">');
  if form_install.saveAs = false then begin
      uplodTemp.Add('<font size="14">');
      uplodTemp.Add('<form name="input" action="/statistic" method="post">');
      uplodTemp.Add('<input name="save" maxlength=255  value="on" type=hidden>');
      uplodTemp.Add('<input type="submit" value="Включить запись"></form>');
  end else begin
      uplodTemp.Add('<font size="14">');
      uplodTemp.Add('<form name="input" action="/statistic" method="post">');
      uplodTemp.Add('<input name="save" maxlength=255 value="off" type=hidden>');
      uplodTemp.Add('<input type="submit" value="Отключить запись"></form>');
  end;


  for forCycle := 0 to 35 do begin
    uplodTemp.AddStrings(HttpStatisticData[forCycle]);
    uplodTemp.Add((data[forCycle]));
  end;
    uplodTemp.AddStrings(HttpStatisticData[36]);
//add string save XLS  *.xlsx

  uplodTemp.Add('<table>');
  If FindFirst(form_install.dirFile+'*.mdb',faAnyFile,direkt)=0 then
    begin
      if FileExists(form_install.dirFile+direkt.Name)
              and (direkt.Name<>'.') and (direkt.Name<>'..')then
        begin
          uplodTemp.Add('<form name="input" action="/'+direkt.Name+'" method="post">');
          uplodTemp.Add('<input name="d" size=100 value="'+direkt.Name+direkt.Name+'" type=hidden>');
          uplodTemp.Add('<tr><td>'+direkt.Name+'</td><td>');
          uplodTemp.Add('<input type="submit" value="Скачать"></td></tr></form>');
        end;  //if directory
      while FindNext(direkt)=0 do
      if FileExists(form_install.dirFile+direkt.Name)and (direkt.Name<>'.') and (direkt.Name<>'..')then
        begin
          uplodTemp.Add('<form name="input" action="/'+direkt.Name+'" method="post">');
          uplodTemp.Add('<input name="d" size=100 value="'+direkt.Name+direkt.Name+'" type=hidden>');
          uplodTemp.Add('<tr><td>'+direkt.Name+'</td><td>');
          uplodTemp.Add('<input type="submit" value="Скачать"></td></tr></form>');
        end;
      end;    //  If FindFirst
  uplodTemp.Add('</table>');

//
    uplodTemp.Add('</html>');
    uplodStreamStatistic:= (uplodTemp.Text);
    uplodTemp.Clear;
    uplodTemp.Free;
end;


constructor THttpStreamRoot.Create;
var
  forCycle:Integer;
begin
  for forCycle := 0 to 13 do HttpStreamRootData[forCycle]:=TStringList.Create;
  for forCycle := 0 to 36 do HttpStatisticData[forCycle]:=TStringList.Create;
HttpStreamRootData[0].Clear;
with HttpStreamRootData[0]   do begin
  add('<html> <p align=right>');
  add('<span style='+#39+'font-size:16.0pt; font-family:"Cambria","serif"; mso-style-textfill-fill-colortransforms:lumm=75000'+#39+'>');
  add('<a href="/statistic">');
  add('<span style='+#39+'color:#77933C; mso-style-textfill-fill-colortransforms:lumm=75000'+#39+'>');
  add('Статистика</span></a></span></p>');
  add('<table  style='+#39+'border-collapse:collapse;'+#39'>');
  add('<tr style='+#39+'mso-yfti-irow:0;'+#39'>');
  add('<td  style='+#39+'border:solid windowtext 1.0pt;  background:#548DD4;  padding:0cm 5.4pt 0cm 5.4pt;'+#39+'>');
  add('<b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('<span style='+#39+'font-size:18.0pt; font-family:"Courier New"; color:white; mso-themecolor:background1'+#39+'>');
  add('Наименование</span></b></td>');
  add('<td style='+#39+'border:solid windowtext 1.0pt; background:#548DD4;padding:0cm 5.4pt 0cm 5.4pt'+#39'>');
  add('<b style='+#39+'mso-bidi-font-weight:normal'+#39'>');
  add('<span style='+#39+'font-size:18.0pt;font-family:"Courier New";color:white; mso-themecolor:background1'+#39'>');
  add('Фаза А</span></b></td>');
  add('<td style='+#39+'border:solid windowtext 1.0pt; background:#548DD4;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('<b style='+#39+'mso-bidi-font-weight:normal'+#39'>');
  add('<span style='+#39+'font-size:18.0pt;font-family:"Courier New";color:white; mso-themecolor:background1'+#39+'>');
  add('Фаза В </span></b> </td>');
  add('<td style='+#39+'border:solid windowtext 1.0pt; background:#548DD4;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('<b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('<span style='+#39+'font-size:18.0pt;font-family:"Courier New";color:white; mso-themecolor:background1'+#39+'>');
  add('Фаза С</span></b> </td>');
  add('<td style='+#39'border:solid windowtext 1.0pt;background:#548DD4;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add('</tr>    <tr style='+#39+'mso-yfti-irow:1'+#39+'>');
  add('<td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('<b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('<span style='+#39+'font-size:18.0pt;font-family:"Courier New";color:white; mso-themecolor:background1'+#39'>');
  add('Напряжение фазы  </span></b>  </td>');
  add('<td style='+#39'border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39'>');
  add('<span style='+#39+'font-size:18.0pt;font-family:"Courier New";'+#39'>');
end;
 with HttpStreamRootData[1]   do begin
   add('</span>  </td>');
   add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
   add('<span style='+#39+'font-size:18.0pt;font-family:"Courier New";'+#39+'>');
end;
HttpStreamRootData[2].AddStrings(HttpStreamRootData[1]);
 with HttpStreamRootData[3]   do begin
   add('</span>  </td>');
   add('<td style='+#39+'border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
   add('</tr> <tr style='+#39+'mso-yfti-irow:2'+#39+'>');
   add('<td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
   add('<b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
   add('<span style='+#39+'font-size:18.0pt;font-family:"Courier New";color:white; mso-themecolor:background1'+#39+'>');
   add('Ток</span></b>  </td>');
   add('<td style='+#39+'border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
   add('<span style='+#39+'font-size:18.0pt;font-family:"Courier New";'+#39'>');
 end;
HttpStreamRootData[4].AddStrings(HttpStreamRootData[2]);
HttpStreamRootData[5].AddStrings(HttpStreamRootData[4]);
with HttpStreamRootData[6]   do begin
  add('</span>  </td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('</tr>             <tr style='+#39+'mso-yfti-irow:3'+#39+'>');
  add('<td  style='+#39+'border:solid windowtext 1.0pt;  background:#548DD4;  padding:0cm 5.4pt 0cm 5.4pt;'+#39+'>');
  add('<b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('<span style='+#39+'font-size:18.0pt; font-family:"Courier New"; color:white; mso-themecolor:background1'+#39+'>');
  add('Наименование</span></b>  </td>');
  add(' <td  style='+#39+'border:solid windowtext 1.0pt;  background:#548DD4;  padding:0cm 5.4pt 0cm 5.4pt;'+#39+'>');
  add('  <b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('<span style='+#39+'font-size:18.0pt; font-family:"Courier New"; color:white; mso-themecolor:background1'+#39+'>');
  add(' Фаза АВ</span></b>  </td>');
  add(' <td  style='+#39+'border:solid windowtext 1.0pt;  background:#548DD4;  padding:0cm 5.4pt 0cm 5.4pt;'+#39+'>');
  add(' <b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('<span style='+#39+'font-size:18.0pt; font-family:"Courier New"; color:white; mso-themecolor:background1'+#39+'>');
  add('Фаза AC</span></b>  </td>');
  add('  <td  style='+#39+'border:solid windowtext 1.0pt;  background:#548DD4;  padding:0cm 5.4pt 0cm 5.4pt;'+#39+'>');
  add('  <b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('<span style='+#39+'font-size:18.0pt; font-family:"Courier New"; color:white; mso-themecolor:background1'+#39+'>');
  add('Фаза ВС</span></b>  </td>');
  add('<td  style='+#39+'border:solid windowtext 1.0pt;  background:#548DD4;  padding:0cm 5.4pt 0cm 5.4pt;'+#39+'></td>');
  add(' </tr>');
  add('');
  add(' <tr style='+#39+'mso-yfti-irow:4'+#39+'>');
  add('  <td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('  <b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('  <span style='+#39+'font-size:18.0pt;font-family:"Courier New";color:white; mso-themecolor:background1'+#39+'>');
  add('  Напряжение линии</span></b>  </td>');
  add('  <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:18.0pt;font-family:"Courier New";'+#39+'>');
end;
//-@ul1@-
with HttpStreamRootData[7]   do begin
  add(' </span>  </td>');
  add('  <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:18.0pt;font-family:"Courier New";'+#39+'>');
end;
//-@ul2@-
HttpStreamRootData[8].AddStrings(HttpStreamRootData[7]);
//-@ul3@-
with HttpStreamRootData[9]   do begin
  add(' </span>  </td>');
  add('  <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'> </td>');
  add(' </tr>');
  add('');
  add(' <tr style='+#39+'mso-yfti-irow:5'+#39+'>');
  add('  <td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('  <b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('  <span style='+#39+'font-size:18.0pt;font-family:"Courier New";color:white; mso-themecolor:background1'+#39+'>');
  add('  Угол сдвига фаз</span>  </td>');
  add('  <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:18.0pt;font-family:"Courier New";'+#39+'>');
end;
// -@y1@-
HttpStreamRootData[10].AddStrings(HttpStreamRootData[8]);
//-@y2@-
HttpStreamRootData[11].AddStrings(HttpStreamRootData[10]);
//-@y3@-
with HttpStreamRootData[12]   do begin
  add(' </span>  </td>');
  add(' <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>   </td>');
  add(' </tr>');
  add(' ');
  add(' <tr style='+#39+'mso-yfti-irow:6;'+#39+'>');
  add('  <td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('  <span style='+#39+'font-size:18.0pt;font-family:"Courier New";color:white; mso-themecolor:background1'+#39+'>');
  add('  Частота</span></b>  </td>');
  add('  <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('  <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('  <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('  <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:18.0pt;font-family:"Courier New";'+#39+'>');
end;
//-@f@-
with HttpStreamRootData[13]   do begin
  add(' </span>  </td>');
  add(' </tr>');
  add('</table>');
  add('<META HTTP-EQUIV="REFRESH" CONTENT=0; URL="/">');
  add('</html>');
end;
//--------------------------
with HttpStatisticData[0] do begin
  add('<table  style='+#39+'border-collapse:collapse;'+#39+'>');
  add('<tr style='+#39+'mso-yfti-irow:0;'+#39+'>');
  add('<td style='+#39+' border:solid windowtext 1.0pt;  background:#548DD4;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('<p class=MsoNormal >');
  add('<span style='+#39+'font-size:16.0pt;mso-style-textfill-fill-colortransforms:lumm=75000'+#39+'>');
  add('<a href="/">');
  add('<span style='+#39+'color:white;'+#39+'><- Назад</span></a></span></p></td> ');
  add('<td style='+#39+' border:solid windowtext 1.0pt;  background:#548DD4;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('<p class=MsoNormal >');
  add('<span style='+#39+'font-size:16.0pt;mso-style-textfill-fill-colortransforms:lumm=75000'+#39+'>');
  add('<a href="/statistic"><span style='+#39+'color:white;'+#39+'>0бновить</span></a></span></p></td>');
  add('</tr>');
  add('<tr style='+#39+'mso-yfti-irow:1;'+#39+'>');
  add('<td  style='+#39+'border:solid windowtext 1.0pt;  background:#548DD4; padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('<b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('<span style='+#39+'font-size:14.0pt; font-family:"Courier New"; color:white; mso-themecolor:background1'+#39+'>');
  add('Дата и время</span></b>  </td>');
  add('<td  style='+#39+'border:solid windowtext 1.0pt;  background:#548DD4; padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('<b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('<span style='+#39+'font-size:14.0pt; font-family:"Courier New"; color:white; mso-themecolor:background1'+#39+'>');
  add('Напряжение линии</span></b>  </td>');
  add('<td  style='+#39+'border:solid windowtext 1.0pt;  background:#548DD4; padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('<b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('<span style='+#39+'font-size:14.0pt; font-family:"Courier New"; color:white; mso-themecolor:background1'+#39+'>');
  add('Напряжение фазы</span></b>  </td>');
  add('<td  style='+#39+'border:solid windowtext 1.0pt;  background:#548DD4; padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('<b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('<span style='+#39+'font-size:14.0pt; font-family:"Courier New"; color:white; mso-themecolor:background1'+#39+'>');
  add('Ток</span></b>  </td>');
  add('<td  style='+#39+'border:solid windowtext 1.0pt;  background:#548DD4; padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('<b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('<span style='+#39+'font-size:14.0pt; font-family:"Courier New"; color:white; mso-themecolor:background1'+#39+'>');
  add('Частота</span></b></p>  </td>');
  add(' </tr>');
  add(' <tr style='+#39+'mso-yfti-irow:2'+#39+'>');
  add('  <td  colspan=5 style='+#39+'border:solid windowtext 1.0pt; background:#548DD4;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('  <p style='+#39+'text-align:center;'+#39+'>  <b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('  <span style='+#39+'font-size:14.0pt;font-family:"Cambria","serif";color:white;mso-themecolor:background1'+#39+'>');
  add('Минимальные значения');
  add('  </span></b></p>  </td>');
  add(' </tr>');
  add(' <tr style='+#39+'mso-yfti-irow:3'+#39+'>');
  add(' <td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:14.0pt;font-family:"Courier New";color:white; mso-themecolor:background1'+#39+'>');
end;
// -@time_a_min@-'
with HttpStatisticData[1] do begin
  add('  </span>  </td>');
  add('  <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:14.0pt;font-family:"Courier New";'+#39+'>');
 end;
// -@uab_min@-
HttpStatisticData[2].AddStrings(HttpStatisticData[1]);
//'-@ua_min@-
HttpStatisticData[3].AddStrings(HttpStatisticData[2]);
//-@ia_min@-
with HttpStatisticData[4] do begin
  add('</span> </td>');
  add('    <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' </tr>');
  add('  <tr style='+#39+'mso-yfti-irow:4'+#39+'>');
  add(' <td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:14.0pt;font-family:"Courier New";color:white; mso-themecolor:background1'+#39+'>');
 end;
 //-@time_b_min@-
HttpStatisticData[5].AddStrings(HttpStatisticData[3]);
 //-@ubc_min@-
HttpStatisticData[6].AddStrings(HttpStatisticData[5]);
 //-@ub_min@-
HttpStatisticData[7].AddStrings(HttpStatisticData[6]);
//-@ib_min@-
with HttpStatisticData[8] do begin
  add('</span>  </td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' </tr>');
  add(' <tr style='+#39+'mso-yfti-irow:5'+#39+'>');
  add('  <td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:14.0pt;font-family:"Courier New";color:white; mso-themecolor:background1'+#39+'>');
end;
   //-@time_c_min@-
HttpStatisticData[9].AddStrings(HttpStatisticData[7]);
 //-@uac_min@-
HttpStatisticData[10].AddStrings(HttpStatisticData[9]);
 //-@uc_min@-
HttpStatisticData[11].AddStrings(HttpStatisticData[10]);
//-@ic_min@-
with HttpStatisticData[12] do begin
  add('</span>  </td>');
  add('  <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('</tr>  ');
  add('    <tr style='+#39+'mso-yfti-irow:6'+#39+'>');
  add('<td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:14.0pt;font-family:"Courier New";'+#39+'>');
end;
 // -@f_min@-
with HttpStatisticData[13] do begin
  add('</span>  </td> </tr>');
  add(' <tr style='+#39+'mso-yfti-irow:7'+#39+'>');
  add('  <td  colspan=5 style='+#39+'border:solid windowtext 1.0pt; background:#548DD4;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('  <p style='+#39+'text-align:center;'+#39+'>  <b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('  <span style='+#39+'font-size:14.0pt;font-family:"Cambria","serif";color:white;mso-themecolor:background1'+#39+'>');
  add('Средние значения');
  add('  </span></b>  </td> </tr>');
  add(' <tr style='+#39+'mso-yfti-irow:8'+#39+'>');
  add(' <td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add('  <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:14.0pt;font-family:"Courier New";'+#39+'>');
end;
 //-@uab_sr@-
HttpStatisticData[14].AddStrings(HttpStatisticData[11]);
//-@ua_sr@-
HttpStatisticData[15].AddStrings(HttpStatisticData[14]);
//-@ia_sr@-
with HttpStatisticData[16] do begin
  add('</span>  </td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' </tr>');
  add(' <tr style='+#39+'mso-yfti-irow:9'+#39+'>');
  add('<td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add('  <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:14.0pt;font-family:"Courier New";'+#39+'>');
end;
 //-@ubc_sr@-
HttpStatisticData[17].AddStrings(HttpStatisticData[15]);
 //-@ub_sr@-
HttpStatisticData[18].AddStrings(HttpStatisticData[17]);
 //-@ib_sr@-
with HttpStatisticData[19] do begin
  add('</span>  </td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add(' </tr>');
  add(' <tr style='+#39+'mso-yfti-irow:10'+#39+'>');
  add('<td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add('  <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:14.0pt;font-family:"Courier New";'+#39+'>');
end;
  // -@uac_sr@-
HttpStatisticData[20].AddStrings(HttpStatisticData[18]);
//-@uc_sr@-
HttpStatisticData[21].AddStrings(HttpStatisticData[20]);
//-@ic_sr@-
with HttpStatisticData[22] do begin
  add('</span>  </td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'> </td>');
  add(' </tr>');
  add('<tr style='+#39+'mso-yfti-irow:11'+#39+'>');
  add('<td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add('  <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:14.0pt;font-family:"Courier New";'+#39+'>');
end;
  //-@f_sr@-
with HttpStatisticData[23] do begin
  add('</span>  </td>');
  add(' </tr>');
  add(' <tr style='+#39+'mso-yfti-irow:12'+#39+'>');
  add('  <td  colspan=5 style='+#39+'border:solid windowtext 1.0pt; background:#548DD4;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add('  <p style='+#39+'text-align:center;'+#39+'>  <b style='+#39+'mso-bidi-font-weight:normal'+#39+'>');
  add('  <span style='+#39+'font-size:14.0pt;font-family:"Cambria","serif";color:white;mso-themecolor:background1'+#39+'>');
  add('Максимальные значения');
  add('  </span></b></p>  </td> </tr>');
  add(' <tr style='+#39+'mso-yfti-irow:13'+#39+'>');
  add(' <td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:14.0pt;font-family:"Courier New";color:white; mso-themecolor:background1'+#39+'>');
end;
  // -@time_a_max@-
HttpStatisticData[24].AddStrings(HttpStatisticData[21]);
//-@uab_max@-
HttpStatisticData[25].AddStrings(HttpStatisticData[24]);
//-@ua_max@-
HttpStatisticData[26].AddStrings(HttpStatisticData[25]);
//-@ia_max@-
with HttpStatisticData[27] do begin
  add('</span>  </td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add(' </tr>');
  add(' <tr style='+#39+'mso-yfti-irow:14'+#39+'>');
  add('  <td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:14.0pt;font-family:"Courier New";color:white; mso-themecolor:background1'+#39+'>');
end;
 //-@time_b_max@-
HttpStatisticData[28].AddStrings(HttpStatisticData[26]);
 //-@ubc_max@-
HttpStatisticData[29].AddStrings(HttpStatisticData[28]);
 //-@ub_max@-
HttpStatisticData[30].AddStrings(HttpStatisticData[29]);
 //-@ib_max@-
with HttpStatisticData[31] do begin
  add('</span>  </td>');
  add(' <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'> </td>');
  add('  </tr>');
  add(' <tr style='+#39+'mso-yfti-irow:15;'+#39+'>');
  add('   <td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:14.0pt;font-family:"Courier New";color:white; mso-themecolor:background1'+#39+'>');
end;
     //-@time_c_max@-
HttpStatisticData[32].AddStrings(HttpStatisticData[26]);
 //-@ubc_max@-
HttpStatisticData[33].AddStrings(HttpStatisticData[28]);
 //-@ub_max@-
HttpStatisticData[34].AddStrings(HttpStatisticData[29]);
 //'-@ic_max@-
with HttpStatisticData[35] do begin
  add('</span>  </td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'> </td>  ');
  add(' </tr>');
  add('<tr style='+#39+'mso-yfti-irow:16'+#39+'>');
  add('<td style='+#39+'border:solid windowtext 1.0pt; background:#92D050;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add('<td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'></td>');
  add('  <td style='+#39+'  border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'+#39+'>');
  add(' <span style='+#39+'font-size:14.0pt;font-family:"Courier New";'+#39+'>');
end;
HttpStatisticData[36].add('</span>  </td> </tr></table>');
end;


end.
