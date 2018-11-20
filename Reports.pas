unit Reports;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  orBrowser, DBTables, RxQuery, Db, DBClient, Menus, ActnList,
  StdCtrls, ComCtrls, jpeg, ExtCtrls, ToolWin, Grids, DBGrids, RXDBCtrl,
  CheckLst, Mask, Buttons, RXSpin, cxGridLevel, cxClasses,
  cxControls, cxGridCustomView, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxGrid, cxStyles, cxCustomData, cxGraphics, cxFilter,
  cxData, cxDataStorage, cxEdit, cxDBData, RxGIF, GIFImage, rxToolEdit,
  rxPlacemnt, cxLookAndFeels, cxLookAndFeelPainters;

type
  TfrmReports = class(TfrmorBrowser)
    PageControl1: TPageControl;
    tsNodes: TTabSheet;
    fcDateStop: TTabSheet;
    tsNar2: TTabSheet;
    TabSheet4: TTabSheet;
    rpTypeC: TCheckListBox;
    Label1: TLabel;
    Label2: TLabel;
    rpNode: TCheckListBox;
    rpStatus: TCheckListBox;
    Label3: TLabel;
    rpDep: TCheckListBox;
    Label4: TLabel;
    rpNasos: TCheckListBox;
    rpPED: TCheckListBox;
    Label10: TLabel;
    Label11: TLabel;
    fcDate: TGroupBox;
    Label7: TLabel;
    Label8: TLabel;
    btnDateForward: TSpeedButton;
    btnDateBackward: TSpeedButton;
    FromDate: TDateEdit;
    ToDate: TDateEdit;
    rpDep_s: TCheckListBox;
    Label9: TLabel;
    rpNGDU_PDK: TCheckListBox;
    Label13: TLabel;
    RxSpinEdit1: TRxSpinEdit;
    Label14: TLabel;
    ComboBox1: TComboBox;
    Label15: TLabel;
    Label17: TLabel;
    rpGroupNodeR: TCheckListBox;
    RadioGroup1: TRadioGroup;
    fcNew: TRadioGroup;
    tsNar3: TTabSheet;
    Label5: TLabel;
    Label6: TLabel;
    Label16: TLabel;
    rpNGDU_S: TCheckListBox;
    rpZavod: TCheckListBox;
    rpDEPOSIT_S: TCheckListBox;
    rpEPU: TCheckListBox;
    Label12: TLabel;
    rpGroupNodeR1: TCheckListBox;
    Label18: TLabel;
    rpPrichina: TCheckListBox;
    Label19: TLabel;
    tsKompl: TTabSheet;
    rpNGDU_Kompl: TCheckListBox;
    Label20: TLabel;
    GroupBox1: TGroupBox;
    Label21: TLabel;
    Label22: TLabel;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    DateEdit1: TDateEdit;
    DateEdit2: TDateEdit;
    cbDateField: TComboBox;
    procedure FormShow(Sender: TObject);
    procedure ResetAllFilter; override;
    procedure AssignFilter; override;
    procedure ColumnFilterClick(Sender: TObject);
    procedure RadioGroup1Click(Sender: TObject);
    procedure ResetFilterExecute(Sender: TObject);
    procedure cbDateFieldClick(Sender: TObject);
    procedure FillComboDateField;
  private
    { Private declarations }
  public
    { Public declarations }
   GroupString, CurView: String;
  end;

var
  frmReports: TfrmReports;
  


implementation
uses MMSG, MSQL, Globals, rxVclUtils, uConnData;

{$R *.DFM}

procedure TfrmReports.FormShow(Sender: TObject);
begin
  inherited;
  IF Global_Caption <> '' Then Caption := Global_Caption;
  //ResetAllFilter
  //ToolButton14.Click;
end;


procedure TfrmReports.ResetAllFilter;

begin
  inherited;
  reff(tsNodes);
  reff(fcDateStop);
  reff(tsNar2);
  reff(tsNar3);
  reff(TabSheet4);
  reff(tsKompl);   
end;

procedure TfrmReports.AssignFilter;

begin
  inherited;
  IF IS_Diagramm <> ''
  Then
   Begin
    srxQuery.MacroByName('Viewer').AsString := CurView;
    srxQuery.MacroByName('GROUP').AsString := GroupString
   End
  Else
   Begin
   // srxQuery.MacroByName('GROUP').AsString := OrderString;
   End;
end;

procedure TfrmReports.ColumnFilterClick(Sender: TObject);
begin
  inherited;
 grid.TitleButtons := ColumnFilter.Down;
end;

procedure TfrmReports.RadioGroup1Click(Sender: TObject);
begin
  inherited;
  Case RadioGroup1.ItemIndex of
   0: Begin
       frmReports.CurView := 'SELECT "Õ¿—Œ—", COUNT(ID) AS " ŒÀ»◊≈—“¬Œ_Œ“ ¿«Œ¬", SUM("ŒÚ‡·ÓÚ‡Î") AS "Œ“–¿¡Œ“¿À" FROM RP_WORKDAYS A';
       frmReports.GroupString := 'GROUP BY "Õ¿—Œ—"';
       ToolButton8.Click;
      End;
   1: Begin
       frmReports.CurView := 'SELECT "ƒ¬»√¿“≈À‹", COUNT(ID) AS " ŒÀ»◊≈—“¬Œ_Œ“ ¿«Œ¬", SUM("ŒÚ‡·ÓÚ‡Î") AS "Œ“–¿¡Œ“¿À" FROM RP_WORKDAYS A';
       frmReports.GroupString := 'GROUP BY "ƒ¬»√¿“≈À‹"';
       ToolButton8.Click;
      End;
   2: Begin
       frmReports.CurView := 'SELECT "”—“¿ÕŒ¬ ¿", COUNT(ID) AS " ŒÀ»◊≈—“¬Œ_Œ“ ¿«Œ¬", SUM("ŒÚ‡·ÓÚ‡Î") AS "Œ“–¿¡Œ“¿À" FROM RP_WORKDAYS A';
       frmReports.GroupString := 'GROUP BY "”—“¿ÕŒ¬ ¿"';
       ToolButton8.Click;
      End;
  End;
end;

procedure TfrmReports.ResetFilterExecute(Sender: TObject);
begin
 // srxQuery.MacroByName('Filter').AsString := '';
  inherited;
// rpDEPOSIT_S.
  frmReports.FillComboDateField;
  frmReports.cbDateField.ItemIndex := 0;

end;

procedure TfrmReports.cbDateFieldClick(Sender: TObject);
begin
  inherited;
  Global_Caption_Date := TField(cbDateField.Items.Objects[cbDateField.ItemIndex]).FieldName;
  DateFilterChange(Sender);
end;

procedure TfrmReports.FillComboDateField;
var k: integer;
begin
  cbDateField.Items.Clear;
  for k := 0 to pred(cdsBrowser.FieldCount) do
    if (cdsBrowser.Fields[k] is TDateField) or
      (cdsBrowser.Fields[k] is TDateTimeField) then
      cbDateField.Items.AddObject(cdsBrowser.Fields[k].DisplayLabel, cdsBrowser.Fields[k]);
end;

end.
