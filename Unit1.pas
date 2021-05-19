unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ShellApi, ExtCtrls;

type
  TForm1 = class(TForm)
    Barcode: TEdit;
    ButtonClear: TButton;
    LabelSimpleStatus: TLabel;
    Panel1: TPanel;
    ListBox1: TListBox;
    LabelStatus: TLabel;
    EditRoot: TEdit;
    ButtonGenerate: TButton;
    ClearScripts: TButton;
    ButtonShowHide: TButton;
    Button3: TButton;
    EditPath: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    CheckBoxOpenMorbidelliEditor: TCheckBox;
    PanelDetailCount: TPanel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    Label3: TLabel;
    procedure BarcodeChange(Sender: TObject);
    procedure ButtonGenerateClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ClearScriptsClick(Sender: TObject);
    procedure ButtonClearClick(Sender: TObject);
    procedure ButtonShowHideClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormDeactivate(Sender: TObject);
    procedure EditPathMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure BarcodeMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ListBox1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure CheckBoxOpenMorbidelliEditorClick(Sender: TObject);
    procedure RadioButton1Click(Sender: TObject);
    procedure RadioButton2Click(Sender: TObject);
    protected Procedure LastFocus(var Mess : TMessage);
    message  WM_ACTIVATE;
  private
    { Private declarations }
  public
    { Public declarations }
  end;               

var
  Form1: TForm1;
  DirCount, FileCount: integer;
  dirs: TStrings;
  PREFIX, ROOTPATH: String;
  FULLSTATE, FEW_SCRIPTS: boolean;
implementation

{$R *.dfm}

Procedure TForm1.LastFocus(var Mess : TMessage);
Begin

  if  Mess.wParam = WA_INACTIVE Then
    Barcode.Color := clWindow
  else
    Barcode.Color := clYellow;

  if Mess.Msg =WM_SYSCOMMAND then
    if Mess.WParam=SC_MINIMIZE then
    begin
      Application.Restore;
    end;

  Barcode.SetFocus;

  Inherited;

End;

procedure ManageFormDesign;
begin
  //Full application interface or compact
  if FULLSTATE then
  begin
    Form1.Width := 730;
    Form1.Height := 210;
    Form1.Panel1.Visible := True;
    Form1.AlphaBlend := False;
    Form1.ButtonShowHide.Caption := 'Hide';
  end
  else
  begin
    Form1.Width := 230;
    Form1.Height := 210;
    if FEW_SCRIPTS then Form1.Height := 154;
    Form1.Panel1.Visible := False;
    Form1.AlphaBlend := True;
    Form1.ButtonShowHide.Caption := 'Show';
  end;

end;

procedure TForm1.BarcodeChange(Sender: TObject);
var f, f_script: TextFile;
  Exist, LeadingZeros: Boolean;
  i: byte;
  detail, DetailName, FileResultName, path, edit, scriptPath, st_script, st: String;

begin

  //Reading control commands 11841, 321321 or 654654
  if StrPos('*', PChar(Barcode.Text)) <> nil then
    Barcode.Text := ''
  else if Barcode.Text = '11841' then
  begin
    FULLSTATE := not FULLSTATE;
    Barcode.Text := '';
  end
  else if Barcode.Text = '-' then
  begin
    Form1.ButtonGenerate.Click();
    Barcode.Text := '';
    LabelSimpleStatus.Caption := 'Generate at ' + TimeToStr(Time);
    Exit;
  end
  else if Barcode.Text = '654654' then
  begin
    Form1.Close;
  end;

  //Set full or comapct application interface
  ManageFormDesign;

  if CheckBoxOpenMorbidelliEditor.Checked then
    FileResultName := 'Scripts\result-editor.txt'
  else
    FileResultName := 'Scripts\result.txt';

  if FileExists(FileResultName) then
  BEGIN
    edit := Barcode.Text;

    if (length(Barcode.Text) = 13) and (Barcode.Text[1] = PREFIX) then
    //Entering detail barcode
      begin
        try
        //Transforming barcode to simple code adding symbol '+'
          LeadingZeros := True;
          st := '';
          for i := length(PREFIX) + 1 to Length(edit)-1 do
          begin
            if edit[i] = '0' then
              begin
                if not LeadingZeros then
                  st := st + edit[i]
              end
            else
            begin
              LeadingZeros := False;
              st := st + edit[i];
            end;
          end;
          edit := st + '+';
        except
        end;
      end;

    if edit <> '' then
    if (edit[length(edit)] = '+') then
    begin
      if Radiobutton1.Checked then
      begin
        edit[length(edit)] := '@';
        edit := edit  + '+'
      end;
      AssignFile(f, FileResultName);
      Reset(f);
      Exist := False;

      While not eof(f) and not Exist do
      begin
         readln(f, detail);
         readln(f, path);
         if RadioButton1.Checked then detail := detail;
         if detail = copy(edit, 1, length(edit) - 1) then
         begin
            Exist := True;
            if form1.CheckBoxOpenMorbidelliEditor.Checked then
            begin
              DetailName := detail + '-editor.ahk';
            end
            else
            begin
                DetailName := detail + '.ahk';
            end;
            scriptPath := GetCurrentDir + '\Scripts\' +  DetailName;
            ShellExecute(0, 'Open', PChar(scriptPath), nil, nil, 1);
            LabelStatus.Caption := 'Τΰιλ ' + scriptPath;
            LabelSimpleStatus.Caption := 'Script ' +  DetailName + ' found';

            AssignFile(f_script, scriptPath);

            Reset(f_script);
            ReadLN(f_script, st_script);
            ReadLN(f_script, st_script);
            st_script := StringReplace(st_script, '{SPACE}', ' ', [rfReplaceAll, rfIgnoreCase]);
            st_script := StringReplace(st_script, '{+}', '+', [rfReplaceAll, rfIgnoreCase]);
            st_script := StringReplace(st_script, ROOTPATH, '', [rfReplaceAll, rfIgnoreCase]);
            st_script := StringReplace(st_script, 'FileName = ', '', [rfReplaceAll, rfIgnoreCase]);
            EditPath.Text := st_script;
            CloseFile(f_script);

            Barcode.Text := '';
         end;
      end;

      CloseFile(f);
    end;

    if (not Exist) and (Barcode.Text <> '') then
    begin
      LabelStatus.Caption := 'File ' + DetailName + ' NOT FOUND';
      LabelSimpleStatus.Caption := 'Code ' + DetailName + ' not found';
      EditPath.Text := '';
    end;

    if length(Barcode.Text) >= 13 then
      Barcode.Text := '';
  END
  else
    LabelSimpleStatus.Caption := 'results.txt not found';
end;

//***************************************************************************
//Generating AHK script by detail number and file path 
procedure GenerateScript(detail, path: String);
var f1, f2, f_template1, f_template2: TextFile;
  st: String;

begin
  AssignFile(f1, 'Scripts\' + detail + '-editor.ahk');
  AssignFile(f2, 'Scripts\' + detail + '.ahk');

  AssignFile(f_template1, 'config-editor.ahk');//Template AHK script
  AssignFile(f_template2, 'config.ahk');//Template AHK script

  if FEW_SCRIPTS then
  begin
    Rewrite(f1);
    Reset(f_template1);
    while not eof(f_template1) do
    begin
      Readln(f_template1, st);
      //Skip insignificant lines
      if length(st) > 0 then
        if st[1] <> ';' then
          Writeln(f1, st)
        //If line starts at ;* then replace it
        else if (st[1] = ';')and(st[2] = '*') then
        begin
          st := StringReplace(path, ' ', '{SPACE}', [rfReplaceAll, rfIgnoreCase]);
          st := StringReplace(st, '+', '{+}', [rfReplaceAll, rfIgnoreCase]);
          Writeln(f1, 'FileName = ' + st);
        end;
    end;
    CloseFile(f1);
    CloseFile(f_template1);
  end;

  Rewrite(f2);
  Reset(f_template2);
  while not eof(f_template2) do
  begin
    Readln(f_template2, st);
    //Skip insignificant lines
    if length(st) > 0 then
      if st[1] <> ';' then
        Writeln(f2, st)
      //If line starts at ;* then replace it
      else if (st[1] = ';')and(st[2] = '*') then
      begin
        st := StringReplace(path, ' ', '{SPACE}', [rfReplaceAll, rfIgnoreCase]);
        st := StringReplace(st, '+', '{+}', [rfReplaceAll, rfIgnoreCase]);
        Writeln(f2, 'FileName = ' + st);
      end;
  end;
  CloseFile(f2);
  CloseFile(f_template2);
end;
//***************************************************************************

procedure GetFiles(const dir: string; list: TStrings);
var rec: TSearchRec;

begin
  if FindFirst(dir + '\*.*', faAnyFile, rec) = 0 then
  repeat
    if (rec.Name = '.') or (rec.Name = '..') then Continue;
    if (rec.Attr and faDirectory) <> 0 then
    begin
      inc(dirCount);
      dirs.Add(dir + '\' + rec.Name);
      GetFiles(dir + '\' + rec.Name, list);
    end
    else
    begin
      inc(fileCount);
      list.Add(dir + '\' + rec.Name);
    end;
  until FindNext(rec) <> 0;
  FindClose(rec);
end;

procedure TForm1.ButtonGenerateClick(Sender: TObject);
var files: TStringList;
  i, k, j, st_start,st_end, details_counter: integer;
  f: TextFile;
  st, filename, detail, swap: String;

begin
  details_counter := 0;

  if Form1.CheckBoxOpenMorbidelliEditor.Checked then
    AssignFile(f, 'Scripts\result-editor.txt')
  else
    AssignFile(f, 'Scripts\result.txt');
  Rewrite(f);

  dirs := TStringList.Create;
  dirCount := 1;
  files := TStringList.Create;
  fileCount := 0;

  GetFiles(ROOTPATH, files);
  ListBox1.Clear;
  for i := 0 to files.Count - 1 do
  begin
    st := files[i];
    k := length(st);

    //----> Get filename without file path and extension
    while (st[k] <> '.')and(k > 0) do
    begin
      dec(k);
    end;
    if k = 0 then k := length(st) + 1;
    st_end := k;
    while st[k] <> '\' do
    begin
      dec(k);
    end;
    st_start := k + 1;
    filename := copy(st, st_start, st_end - st_start);
    //<----

    //This code simply works as follows:
    //IN
    //filename: gdsfgsjgh shfgsjgf +789 +1123  +121.xlsx
    //OUT
    //GenerateScript(123, gdsfgsjgh shfgsjgf +789 +1123  +121.xlsx)
    //GenerateScript(1123, gdsfgsjgh shfgsjgf +789 +1123  +121.xlsx)
    //GenerateScript(789, gdsfgsjgh shfgsjgf +789 +1123  +121.xlsx)
    k:= length(filename);
    detail := '';
    while k > 0 do
    begin
      if filename[k] = '+' then
      begin
        swap := '';
        for j:=length(detail) downto 1 do swap := swap + detail[j];
        if Pos('@',  files[i]) > 0 then
            swap := swap + '@';
        writeln(f, swap);
        writeln(f, files[i]);
        GenerateScript(swap, files[i]);
        ListBox1.AddItem(StringReplace(files[i], ROOTPATH, '', [rfReplaceAll, rfIgnoreCase]) + ' (' + swap + ')', Sender);
        inc(details_counter);
        detail := '';
      end

      else
      begin
        if (filename[k] >= '0') and (filename[k] <= '9') then
          detail := detail + filename[k];
      end;
      dec(k);
    end;
  end;

  files.Free;
  dirs.Free;
  CloseFile(f);
  Form1.LabelStatus.Caption := IntToStr(fileCount) + ' files in ' + intToStr(dirCount) + ' folders were processed (' + IntToStr(details_counter) + ' scripts created)';

  Barcode.Color := clYellow;
  Barcode.SetFocus;
end;

procedure TForm1.FormCreate(Sender: TObject);
var f: TextFile;
  st: String;

begin
  AssignFile(f, 'config.ahk');
  Reset(f);
  Readln(f, st);
  ROOTPATH := StringReplace(st, ';', '', [rfReplaceAll, rfIgnoreCase]);
  Readln(f, st);
  ROOTPATH := StringReplace(st, ';', '', [rfReplaceAll, rfIgnoreCase]);
  EditRoot.Text := st;
  PREFIX := '1';
  FEW_SCRIPTS := False;

  Readln(f, st);
  ListBox1.AddItem(st, Sender);
  Readln(f, st);
    ListBox1.AddItem(st, Sender);
  if st = ';YES' then
  begin
    CheckBoxOpenMorbidelliEditor.Visible := True;
    FEW_SCRIPTS := True;
  end;

  FULLSTATE := False;
  ManageFormDesign;
  CloseFile(f);
end;

procedure TForm1.ClearScriptsClick(Sender: TObject);
var files: TStringList;
  i:integer;

begin
  files := TStringList.Create;
  GetFiles(GetCurrentDir + '\Scripts', files);
  ListBox1.Clear;
  for i := 0 to files.Count - 1 do
  begin
    DeleteFile(PChar(files[i]));
    ListBox1.AddItem(StringReplace(files[i], GetCurrentDir + '\Scripts\', '', [rfReplaceAll, rfIgnoreCase]) + ' deleted', Sender);
  end;

  Form1.LabelStatus.Caption := IntToStr(files.Count) + ' files deleted';

  files.Free;

  Barcode.Color := clYellow;
  Barcode.SetFocus;
end;

procedure TForm1.ButtonClearClick(Sender: TObject);
begin
  LabelStatus.Caption := '';
  ListBox1.Clear;
  Barcode.Text := '';
  Barcode.Color := clYellow;
  Barcode.SetFocus;
end;

procedure TForm1.ButtonShowHideClick(Sender: TObject);
begin
  FULLSTATE := not FULLSTATE;
  ManageFormDesign;
  Barcode.Text := '';
  Barcode.Color := clYellow;
  Barcode.SetFocus;
end;

procedure TForm1.FormActivate(Sender: TObject);
begin
  Barcode.Color := clYellow;
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  Form1.Close;
end;

procedure TForm1.FormDeactivate(Sender: TObject);
begin
  Barcode.Color := clWindow;
end;

procedure TForm1.EditPathMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  Barcode.Color := clWindow;
end;

procedure TForm1.BarcodeMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  Barcode.Color := clYellow;
end;

procedure TForm1.ListBox1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  Barcode.Color := clYellow;
  Barcode.SetFocus;
end;

procedure TForm1.CheckBoxOpenMorbidelliEditorClick(Sender: TObject);
begin
  Barcode.Color := clYellow;
  Barcode.SetFocus;
end;

procedure TForm1.RadioButton1Click(Sender: TObject);
begin
  Barcode.Color := clYellow;
  Barcode.SetFocus;
end;

procedure TForm1.RadioButton2Click(Sender: TObject);
begin
  Barcode.Color := clYellow;
  Barcode.SetFocus;
end;

end.
