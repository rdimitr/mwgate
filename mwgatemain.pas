unit mwgatemain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, Vcl.Grids, IniFiles, ShellApi,
  Vcl.ExtCtrls, Vcl.ComCtrls, REST.Types, REST.Client, Data.Bind.Components,
  Data.Bind.ObjectScope, System.JSON, System.Generics.Collections, ComObj;

type
  TForm1 = class(TForm)
    StringGrid1: TStringGrid;
    Panel1: TPanel;
    BitBtn1: TBitBtn;
    StatusBar1: TStatusBar;
    Panel2: TPanel;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    RESTRequest1: TRESTRequest;
    RESTClient1: TRESTClient;
    RESTResponse1: TRESTResponse;
    Label1: TLabel;
    Label2: TLabel;
    ComboBox1: TComboBox;
    Label3: TLabel;
    RowCountLabel: TLabel;
    BitBtn4: TBitBtn;
    SaveDialog1: TSaveDialog;
    Bevel1: TBevel;
    Bevel2: TBevel;
    SaveDoctorCB: TCheckBox;
    LimitRowsCB: TCheckBox;
    BitBtn5: TBitBtn;
    BitBtn6: TBitBtn;
    BitBtn7: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure StringGrid1FixedCellClick(Sender: TObject; ACol, ARow: Integer);
    procedure StringGrid1DblClick(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    function GridIsEmpty: boolean;
    procedure BitBtn5Click(Sender: TObject);
    procedure StringGrid1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
    AppName, DBname, MWserver, MWApi, DTformat, DTdelim, CSVdelim: string;
    isStoredDoctor, isReadLimit: boolean;
    nameStoredDoctor: string;
    LimitRows: integer;
    rowsJSONArray: TJSONArray;
    rowsJsonArrEnum: TJSONArray.TEnumerator;
    sortFlagsArray: array[0..5] of integer;
    DateTimeFormatString: TFormatSettings;
    procedure RefreshDoctors;
    procedure FillGrid(fltrString: string);
    procedure RefreshIniFile;
    procedure WorkIniFile;
    procedure SortGrid(Column: integer);
    procedure ShowRecCount;
    function GetExportFileName(ext: string): String;
  public
    { Public declarations }
    isFirstRun: boolean;
  end;

var
  Form1: TForm1;
  IniFile: TIniFile;

const
  NO_ORDER   = 0;
  ASC_ORDER  = 1;
  DESC_ORDER = 2;
  EXCEL_FILE_EXT = '.xlsx';
  CSV_FILE_EXT = '.csv';

implementation

{$R *.dfm}


function CompareInteger(x1, x2: integer): integer;
begin
  if (x1>x2) then Result:=1 else if (x1<x2) then Result:=-1 else Result:=0
end;


function CompareDate(s1, s2: string; fmt:TFormatSettings): integer;
var d1, d2: TDateTime;
begin
  d1:=StrToDateTime(s1, fmt);
  d2:=StrToDateTime(s2, fmt);
  if (d1>d2) then Result:=1 else if (d1<d2) then Result:=-1 else Result:=0
end;

function TForm1.GridIsEmpty: boolean;
begin
  if (StringGrid1.Cells[0,1] = EmptyStr) and (StringGrid1.Cells[5,1] = EmptyStr) then
     Result:=True
  else
     Result:=False;
end;

procedure TForm1.ShowRecCount;
begin
  if GridIsEmpty then
      RowCountLabel.Caption:='Записей: 0'
  else
      RowCountLabel.Caption:='Записей: ' + IntToStr(StringGrid1.RowCount-1);
end;


procedure TForm1.BitBtn1Click(Sender: TObject);
var parStr: string;
begin
   if (StringGrid1.Cells[5, StringGrid1.Row] = EmptyStr) or (StringGrid1.RowCount=1) then
       ShowMessage('Просмотр документа невозможен. Отсутствуют данные в таблице!')
   else begin
       parStr := ' -server=' + MWserver + ' -cname=' + DBname + ' -runMacro=formedit(' +
                 StringGrid1.Cells[5, StringGrid1.Row] + ')';
       ShellExecute(handle, 'open', PWideChar(AppName), PWideChar(parStr), nil, SW_SHOWNORMAL);
   end;
end;


procedure TForm1.RefreshDoctors;
var j: integer;
begin
     with ComboBox1 do begin
          Clear;
          Items.Add('<Все>');
          for j := 1 to StringGrid1.RowCount-1 do
              if Items.IndexOf(StringGrid1.Cells[4, j]) = -1 then
                 Items.Add(StringGrid1.Cells[4, j]);
          ItemIndex := 0;
     end;
end;


procedure TForm1.SpeedButton2Click(Sender: TObject);
var fltrExpr: string;
begin
   if ((Combobox1.Items.Count > 0) and (StringGrid1.RowCount>1) and (not GridIsEmpty)) or
      (ComboBox1.Items[ComboBox1.ItemIndex] = '<Все>') then begin
            fltrExpr:=ComboBox1.Items[ComboBox1.ItemIndex];
            if fltrExpr = '<Все>' then begin
               //fltrExpr:=EmptyStr;
               BitBtn2Click(Sender);
            end
            else begin
               FillGrid(fltrExpr);
               ShowRecCount;
            end;
   end;
end;


procedure TForm1.SortGrid(Column: integer);
var
  i, j: integer;
  tmpRow: TStringList;
  resCompare: integer;
begin
  if StringGrid1.RowCount < 3 then exit;

  if (sortFlagsArray[Column] =  NO_ORDER) or (sortFlagsArray[Column] =  DESC_ORDER) then
     sortFlagsArray[Column]:=ASC_ORDER
  else
     sortFlagsArray[Column]:=DESC_ORDER;

  tmpRow:= TStringList.Create;
  try
    for i:=1 to StringGrid1.RowCount-1 do begin    //i:=0
      for j:=i+1 to StringGrid1.RowCount-1 do begin

        if Column in [2,4] then
           resCompare:=AnsiCompareStr(StringGrid1.Cells[Column, i], StringGrid1.Cells[Column, j])
        else if Column in [0,1,5] then
                resCompare:=CompareInteger(StrToInt(StringGrid1.Cells[Column, i]),
                                           StrToInt(StringGrid1.Cells[Column, j]))
             else
                resCompare:=CompareDate(StringGrid1.Cells[Column, i], StringGrid1.Cells[Column, j],
                                        DateTimeFormatString);

        if sortFlagsArray[Column]=DESC_ORDER then resCompare:=-resCompare;

        if resCompare>0 then
          begin
              tmpRow.Assign(StringGrid1.Rows[i]);
              StringGrid1.Rows[i]:=StringGrid1.Rows[j];
              StringGrid1.Rows[j]:=tmpRow;
          end;
      end;
    end;
  finally
    tmpRow.Free;
  end;
end;


procedure TForm1.StringGrid1DblClick(Sender: TObject);
begin
     BitBtn1Click(Sender);
end;

procedure TForm1.StringGrid1FixedCellClick(Sender: TObject; ACol,
  ARow: Integer);
var sortCol: integer;
begin
     if ARow = 0 then begin
        sortCol:=ACol;
        SortGrid(sortCol);
     end;
end;


procedure TForm1.StringGrid1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
     BitBtn1Click(Sender);
end;

procedure TForm1.FillGrid(fltrString: string);
var  detailJSONArray: TJSONArray;
     maxRowNum, currRow: integer;
     nameDoctor: string;
begin
  Screen.Cursor:=crHourGlass;
  try
     maxRowNum:=2;
     StringGrid1.RowCount:=maxRowNum;
     rowsJsonArrEnum:=rowsJSONArray.GetEnumerator;
     if (isFirstRun) and (isStoredDoctor) and (nameStoredDoctor <> EmptyStr)
         and (Combobox1.Items[Combobox1.ItemIndex] <> '<Все>') then fltrString:=nameStoredDoctor;

     while rowsJsonArrEnum.MoveNext do begin
        detailJSONArray := rowsJSONArrEnum.Current as TJSONArray;

        nameDoctor:=Copy(detailJSONArray.Items[5].ToString,2,length(detailJSONArray.Items[5].ToString)-2);
        if (fltrString = EmptyStr) or ((fltrString <> EmptyStr) and (fltrString = nameDoctor)) then begin
              with StringGrid1 do begin
                RowCount:= maxRowNum;
                currRow := maxRowNum - 1;
                Cells[0,currRow]:=detailJSONArray.Items[0].ToString;
                Cells[1,currRow]:=detailJSONArray.Items[1].ToString;
                Cells[2,currRow]:=Copy(detailJSONArray.Items[2].ToString,2,length(detailJSONArray.Items[2].ToString)-2);
                Cells[3,currRow]:=Copy(detailJSONArray.Items[3].ToString,2,length(detailJSONArray.Items[3].ToString)-6);
                Cells[4,currRow]:=nameDoctor;
                Cells[5,currRow]:=detailJSONArray.Items[6].ToString;
              end;
              inc(maxRowNum);
        end;
     end;
  finally
     TThread.Synchronize(nil,
         procedure
           begin
               Screen.Cursor := crDefault;
           end
      );
  end;
end;


procedure TForm1.BitBtn2Click(Sender: TObject);
var   jsRequest: TJSONObject;
      beginDT, endDT: string;
      x, k: integer;
      tmpJSONArray: TJSONArray;
begin
  Screen.Cursor:=crSQLWait;
  try
     beginDT := DateToStr(DateTimePicker1.DateTime) + ' 00:00:01';
     endDT := DateToStr(DateTimePicker2.DateTime) + ' 23:59:59';

     jsRequest := TJSONObject.Create();
     jsRequest.AddPair('beginInterval', beginDT);
     jsRequest.AddPair('endInterval', endDT);
     jsRequest.AddPair('listStatus', TJSONArray.Create());
     jsRequest.AddPair('listDoctors', TJSONArray.Create());
     RESTRequest1.AddBody(jsRequest);
     jsRequest.Free();
     RESTRequest1.Execute();

     jsRequest:=TJSONObject.Create;
     jsRequest.Parse(TEncoding.UTF8.GetBytes(RESTResponse1.Content),0);
     rowsJSONArray:=TJSONArray.Create();
     if isReadLimit and (LimitRows > 0) and ((jsRequest.Get('rows').JsonValue as TJSONArray).Count>LimitRows) then begin
         tmpJSONArray:=TJSONArray.Create();
         tmpJSONArray:=jsRequest.Get('rows').JsonValue as TJSONArray;
         for x := 0 to tmpJSONArray.Count-1 do begin
             if x<LimitRows then begin
                   rowsJSONArray.AddElement( tmpJSONArray.Items[x] );
             end;
         end;
     end
     else
         rowsJSONArray:= jsRequest.Get('rows').JsonValue as TJSONArray;
  finally
      TThread.Synchronize(nil,
         procedure
           begin
               Screen.Cursor := crDefault;
           end
       );
  end;
  FillGrid(EmptyStr);

  RefreshDoctors;
  if (isFirstRun) then isFirstRun:=False;
  for k:=0 to length(sortFlagsArray)-1 do sortFlagsArray[k]:=NO_ORDER;
  ShowRecCount;
end;


procedure TForm1.BitBtn3Click(Sender: TObject);
begin
     if Application.MessageBox('Вы хотите выйти из программы?', 'Выход',
        MB_ICONQUESTION + MB_YESNO) = IDYES then begin
             RefreshIniFile;
             Application.Terminate;
     end;
end;


function TForm1.GetExportFileName(ext: string): String;
begin
  SaveDialog1.FileName:='export';
  if ext = CSV_FILE_EXT then
      SaveDialog1.Filter := 'CSV file (*.csv)|*.CSV'
  else
      SaveDialog1.Filter := 'MS EXCEL file (*.xlsx)|*.XLSX';
  if SaveDialog1.Execute then begin
     Result := SaveDialog1.FileName;
     if LowerCase(ExtractFileExt(Result)) <> ext then Result := Result + ext;
  end
  else
     Result:=EmptyStr;
end;



procedure TForm1.BitBtn4Click(Sender: TObject);
var
  ExcelApp, Sheet: variant;
  Col, Row: Word;
  sXLSXfileName: string;
begin
  Screen.Cursor:=crHourglass;
  ExcelApp := CreateOleObject('Excel.Application');
  try
    ExcelApp.Visible := false;
    ExcelApp.Workbooks.Add;
    Sheet := ExcelApp.ActiveWorkbook.Worksheets[1];
    for Col := 0 to StringGrid1.ColCount - 1 do
        for Row := 0 to StringGrid1.RowCount - 1 do
            Sheet.Cells[Row + 1, Col + 1] := StringGrid1.Cells[Col, Row];
    sXLSXfileName:=GetExportFileName(EXCEL_FILE_EXT);
    if sXLSXfileName <> EmptyStr then begin
       ExcelApp.ActiveWorkbook.SaveAs(sXLSXfileName);
       ShowMessage('Данные сохранены!');
       ShellExecute(handle, 'open', PWideChar(sXLSXfileName), nil, nil, SW_SHOWNORMAL);
    end;
  finally
    ExcelApp.Application.Quit;
    ExcelApp := unassigned;
    TThread.Synchronize(nil,
         procedure
           begin
               Screen.Cursor := crDefault;
           end
     );
  end;
end;


procedure TForm1.BitBtn5Click(Sender: TObject);
var
  i : Integer;
  listCSV : TStrings;
  tmp_str: string;
  sCSVfileName: string;
begin
  Screen.Cursor:=crHourglass;
  try
     listCSV:= TStringList.Create;
     for i:= 0 to StringGrid1.RowCount-1 do begin
         StringGrid1.Rows[i].StrictDelimiter:= true;
         StringGrid1.Rows[i].Delimiter:= CSVdelim[1];
         tmp_str:= StringGrid1.Rows[i].DelimitedText;
         listCSV.Add(tmp_str);
     end;
     sCSVfileName:=GetExportFileName(CSV_FILE_EXT);
     if sCSVfileName <> EmptyStr then begin
        listCSV.SaveToFile(sCSVfileName);
        ShowMessage('Данные сохранены!');
        ShellExecute(handle, 'open', PWideChar(sCSVfileName), nil, nil, SW_SHOWNORMAL);
     end;
  finally
    TThread.Synchronize(nil,
         procedure
           begin
               Screen.Cursor := crDefault;
           end
     );
  end;
end;


procedure TForm1.RefreshIniFile;
var sIniFile: string;
begin
  try
    sIniFile:=ChangeFileExt(Application.ExeName,'.ini');
    IniFile:= TIniFile.Create(sIniFile);
    with IniFile do begin
         WriteBool('SAVED', 'SaveDoctor', SaveDoctorCB.Checked);
         if SaveDoctorCB.Checked = True then begin
            if (ComboBox1.Items[ComboBox1.ItemIndex] <> EmptyStr) and (ComboBox1.Items[ComboBox1.ItemIndex]<>'<Все>')then begin
               WriteString('SAVED', 'NameDoctor', ComboBox1.Items[ComboBox1.ItemIndex]);
            end
            else begin
               WriteBool('SAVED', 'SaveDoctor', false);
               WriteString('SAVED', 'NameDoctor', EmptyStr);
            end;
         end
         else begin
            WriteString('SAVED', 'NameDoctor', EmptyStr);
         end;
         WriteBool('SAVED', 'ReadLimit', LimitRowsCB.Checked);
    end;
  finally
    IniFile.Free;
  end;
end;


procedure TForm1.WorkIniFile;
var sIniFile: string;
begin
  try
    sIniFile:=ChangeFileExt(Application.ExeName,'.ini');
    IniFile:= TIniFile.Create(sIniFile);

    with IniFile do begin
      if not FileExists(sIniFile) then begin
         WriteString('MW', 'Application','C:\Medwork\Medwork.exe');
         WriteString('MW', 'MW-Server','192.168.0.142');
         WriteString('MW', 'Database','Medwork');
         WriteString('MW', 'MW-Api','http://192.168.0.145:5050/listdocs');
         WriteString('DATE', 'ShortFormat', 'dd.mm.yyyy');
         WriteString('DATE', 'DateDelimiter', '.');
         WriteString('DATE', 'CSVDelimiter', '#');
         WriteBool('SAVED', 'SaveDoctor', false);
         WriteString('SAVED', 'NameDoctor', EmptyStr);
         WriteInteger('SAVED','LimitRows', 300);
         WriteBool('SAVED', 'ReadLimit', true);
      end;

      AppName:=ReadString('MW', 'Application','C:\Medwork\Medwork.exe');
      MWserver:= ReadString('MW', 'MW-Server','192.168.0.142');
      DBname:= ReadString('MW', 'Database','Medwork');
      MWApi:= ReadString('MW', 'MW-Api','http://127.0.0.1:5050/listdocs');

      DTformat:=ReadString('DATE', 'ShortFormat', 'dd.mm.yyyy');
      DTdelim:=ReadString('DATE', 'DateDelimiter', '.');
      CSVdelim:=ReadString('DATE', 'CSVDelimiter', '#');

      isStoredDoctor:=ReadBool('SAVED', 'SaveDoctor',false);
      nameStoredDoctor:=ReadString('SAVED', 'NameDoctor', EmptyStr);
      LimitRows:=ReadInteger('SAVED','LimitRows', 300);
      isReadLimit:=ReadBool('SAVED', 'ReadLimit', true);

    end;

  finally
    IniFile.Free;
  end;
end;


procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
     RefreshIniFile;
end;

procedure TForm1.FormCreate(Sender: TObject);
var k: integer;
begin
   isFirstRun:=True;
   WorkIniFile();
   DateTimePicker1.DateTime:=Now-5;
   DateTimePicker2.DateTime:=Now;

   StatusBar1.Panels[0].Text := 'IP Medwork: ' + MWserver;
   StatusBar1.Panels[1].Text:= 'Database: ' + DBname;
   StatusBar1.Panels[2].Text:= 'MW Api: ' + MWApi;

   RestClient1.BaseURL:=MWApi;

   with StringGrid1 do begin
      Cells[0,0] := 'А/карта';
      Cells[1,0] := 'Посещение';
      Cells[2,0] := 'ФИО пациента';
      Cells[3,0] := 'Дата создания';
      Cells[4,0] := 'ФИО врача';
      Cells[5,0] := '№ документа';
   end;

   for k:=0 to length(sortFlagsArray)-1 do sortFlagsArray[k]:=NO_ORDER;
   SaveDoctorCB.Checked:=isStoredDoctor;
   LimitRowsCB.Checked:=isReadLimit;
   if isStoredDoctor then begin
      Combobox1.Items.Clear;
      Combobox1.Items.Add('<Все>');
      Combobox1.Items.Add(nameStoredDoctor);
      Combobox1.ItemIndex:=1;
   end;

   {$WARN SYMBOL_PLATFORM OFF}
       DateTimeFormatString:=TFormatSettings.Create(LOCALE_USER_DEFAULT);
       DateTimeFormatString.ShortDateFormat := DTformat;
       DateTimeFormatString.DateSeparator := DTdelim[1];
   {$WARN SYMBOL_PLATFORM ON}

end;

end.
