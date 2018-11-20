unit RemBrowser;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  orBrowser, DBTables, RxQuery, Db, DBClient, Menus, ActnList,
  ComCtrls, StdCtrls, ExtCtrls, ToolWin, Grids, DBGrids, RXDBCtrl,
  Mask, Buttons, jpeg, CheckLst, cxStyles, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxGridLevel, cxClasses, cxControls,
  cxGridCustomView, cxGrid, cxPropertiesStore, Variants, rxStrUtils,
  cxCustomData, cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit,
  cxDBData, cxCheckBox, RxGIF, GIFImage, rxToolEdit, rxPlacemnt, cxLookAndFeels,
  cxLookAndFeelPainters;

type
  TfrmRemBrowser = class(TfrmorBrowser)
    tbNodesType: TToolBar;
    tcNodesType: TTabControl;
    Label1: TLabel;
    Label2: TLabel;
    fcTypeC: TComboBox;
    fcModel: TComboBox;
    fcSect: TRadioGroup;
    fcNew: TRadioGroup;
    fcCtrl: TCheckBox;
    Label3: TLabel;
    fcDepartment: TComboBox;
    cdsBrowserID: TFloatField;
    cdsBrowserNODESID: TFloatField;
    cdsBrowserREMOUNTTYPE: TFloatField;
    cdsBrowserREMOUNTTYPE_S: TStringField;
    cdsBrowserDEPARTMENT: TFloatField;
    cdsBrowserDEPARTMENT_S: TStringField;
    cdsBrowserSTEP: TFloatField;
    cdsBrowserSTEP_S: TStringField;
    cdsBrowserDATEIN: TDateTimeField;
    cdsBrowserDATEEND: TDateTimeField;
    cdsBrowserNODE: TFloatField;
    cdsBrowserSECT: TStringField;
    cdsBrowserNUM: TStringField;
    cdsBrowserINVNUM: TStringField;
    cdsBrowserACCOUNTER: TFloatField;
    cdsBrowserACCOUNTER_S: TStringField;
    cdsBrowserNEW: TStringField;
    cdsBrowserCTRL: TStringField;
    cdsBrowserFROMDEPARTMENT: TFloatField;
    cdsBrowserFROMDEPARTMENT_S: TStringField;
    cdsBrowserWORKDAYS: TFloatField;
    cdsBrowserTYPEC: TFloatField;
    cdsBrowserTYPEC_S: TStringField;
    cdsBrowserMODEL: TFloatField;
    Label4: TLabel;
    fcRemountType: TComboBox;
    Label5: TLabel;
    fcStep: TComboBox;
    faRemountType: TCheckBox;
    faStep: TCheckBox;
    Label6: TLabel;
    fcFromDepartment: TComboBox;
    fcDate: TGroupBox;
    Label7: TLabel;
    Label8: TLabel;
    btnDateForward: TSpeedButton;
    btnDateBackward: TSpeedButton;
    FromDate: TDateEdit;
    ToDate: TDateEdit;
    cbDateField: TComboBox;
    DoStep: TAction;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    N15: TMenuItem;
    ReturnToRemount: TAction;
    NextStep: TAction;
    ShowNode: TAction;
    ToolButton6: TToolButton;
    ToolButton20: TToolButton;
    Label9: TLabel;
    fcACCOUNTER: TComboBox;
    N16: TMenuItem;
    N17: TMenuItem;
    N18: TMenuItem;
    N19: TMenuItem;
    N20: TMenuItem;
    plNoSect: TPanel;
    cdsBrowserNODE_S: TStringField;
    fcNODE: TEdit;
    cdsBrowserMODEL_S: TStringField;
    cdsBrowserREAL_LENGTHL: TFloatField;
    cdsBrowserMECH_POVREGD: TFloatField;
    cdsBrowserSROSTKOV: TFloatField;
    cdsBrowserUDLINIT: TFloatField;
    cdsBrowserCENA: TFloatField;
    cdsBrowserSUM_DEFECT: TFloatField;
    cdsBrowserSUM_NEW: TFloatField;
    rpGroupNode: TCheckListBox;
    Label10: TLabel;
    cxGrid1DEDBTableView1ID: TcxGridDBColumn;
    cxGrid1DEDBTableView1REMOUNTTYPE_S: TcxGridDBColumn;
    cxGrid1DEDBTableView1DEPARTMENT_S: TcxGridDBColumn;
    cxGrid1DEDBTableView1STEP_S: TcxGridDBColumn;
    cxGrid1DEDBTableView1DATEIN: TcxGridDBColumn;
    cxGrid1DEDBTableView1DATEEND: TcxGridDBColumn;
    cxGrid1DEDBTableView1SECT: TcxGridDBColumn;
    cxGrid1DEDBTableView1NUM: TcxGridDBColumn;
    cxGrid1DEDBTableView1INVNUM: TcxGridDBColumn;
    cxGrid1DEDBTableView1ACCOUNTER_S: TcxGridDBColumn;
    cxGrid1DEDBTableView1NEW: TcxGridDBColumn;
    cxGrid1DEDBTableView1CTRL: TcxGridDBColumn;
    cxGrid1DEDBTableView1FROMDEPARTMENT_S: TcxGridDBColumn;
    cxGrid1DEDBTableView1WORKDAYS: TcxGridDBColumn;
    cxGrid1DEDBTableView1TYPEC_S: TcxGridDBColumn;
    cxGrid1DEDBTableView1REAL_LENGTHL: TcxGridDBColumn;
    cxGrid1DEDBTableView1MECH_POVREGD: TcxGridDBColumn;
    cxGrid1DEDBTableView1SROSTKOV: TcxGridDBColumn;
    cxGrid1DEDBTableView1UDLINIT: TcxGridDBColumn;
    cxGrid1DEDBTableView1CENA: TcxGridDBColumn;
    cxGrid1DEDBTableView1MODEL_S: TcxGridDBColumn;
    cxGrid1DEDBTableView1SUM_DEFECT: TcxGridDBColumn;
    cxGrid1DEDBTableView1SUM_NEW: TcxGridDBColumn;
    cxStyleRepository1: TcxStyleRepository;
    CTRLStyle: TcxStyle;
    cdsBrowserGRUPPA: TStringField;
    cxGrid1DEDBTableView1Gruppa: TcxGridDBColumn;
    cxPropertiesStore1: TcxPropertiesStore;
    cdsBrowserWELL: TStringField;
    cdsBrowserKUST: TStringField;
    cdsBrowserDEPOSIT_S: TStringField;
    cdsBrowserNGDU_S: TStringField;
    cxGrid1DEDBTableView1WELL: TcxGridDBColumn;
    cxGrid1DEDBTableView1KUST: TcxGridDBColumn;
    cxGrid1DEDBTableView1DEPOSIT_S: TcxGridDBColumn;
    cxGrid1DEDBTableView1NGDU_S: TcxGridDBColumn;
    cdsBrowserREAL_LENGTHL_A: TFloatField;
    cxGrid1DEDBTableView1REAL_LENGTHL_A: TcxGridDBColumn;
    procedure FormCreate(Sender: TObject);
    procedure tbNodesTypeResize(Sender: TObject);
    procedure SetDefaultOrder; override;
    procedure ResetAllFilter; override;
    procedure CustomiseFields; override;
    procedure ActionUpDate; override;
    procedure DoDelete(var ID: integer; var DelRes: boolean); override;
    procedure DoBeforeFiscal; override;
    procedure tcNodesTypeChange(Sender: TObject);
    procedure GridGetBtnParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; var SortMarker: TSortMarker;
      IsDown: Boolean);
    procedure FormShow(Sender: TObject);
    procedure cbMainResize(Sender: TObject);
    procedure cbDateFieldClick(Sender: TObject);
    procedure AddRecordExecute(Sender: TObject);
    procedure AddHistRecordExecute(Sender: TObject);
    procedure EditRecordExecute(Sender: TObject);
    procedure DoStepExecute(Sender: TObject);
    procedure ReturnToRemountExecute(Sender: TObject);
    procedure NextStepExecute(Sender: TObject);
    procedure ShowNodeExecute(Sender: TObject);
    procedure cxGrid1DEDBTableView1CTRLStylesGetContentStyle(
      Sender: TcxCustomGridTableView; ARecord: TcxCustomGridRecord;
      AItem: TcxCustomGridTableItem; out AStyle: TcxStyle);
    procedure ViewRecordExecute(Sender: TObject);
    procedure DeleteRecordExecute(Sender: TObject);
  private
    { Private declarations }
    FCurNodeType: integer;
    FRmBrowserMode: boolean;
    procedure CustomiseFilter;
    procedure SetCurNode(const Value: integer);
    procedure SetRmBrowserMode(const Value: boolean);
    procedure FillComboDateField;
    procedure SENDTOREM(Historycaly: Boolean);
  public
    { Public declarations }
    procedure ShowRemBrowser(Remount, Modal: Boolean);
    property CurNodeType: integer read FCurNodeType write SetCurNode;
    property RmBrowserMode: boolean read FRmBrowserMode write SetRmBrowserMode;
  end;

var
  frmRemBrowser: TfrmRemBrowser;

implementation

uses MMSG, MSql, uConnData, globals, rxVclUtils, NodeBrowser, RemountEditor,
     NodeEditor, SelStatNode;

{$R *.DFM}

{ TfrmRemBrowser }

procedure TfrmRemBrowser.ShowRemBrowser(Remount, Modal: Boolean);
begin
  RmBrowserMode := Remount;
  ShowBrowser(Modal);
end;

procedure TfrmRemBrowser.SetCurNode(const Value: integer);
begin
  FCurNodeType := Value;
end;

procedure TfrmRemBrowser.ActionUpDate;
begin
  inherited;
  DoStep.Visible := AddRecord.Visible and (TAccessLevel(ArmCurrentInfo.UserLevel) in ulAllWorker);
  NextStep.Visible := DoStep.Visible;
  ReturnToRemount.Visible := DoStep.Visible and (TAccessLevel(ArmCurrentInfo.UserLevel) in ulAdmLevel);
  DoStep.Enabled := ViewRecord.Enabled and (cdsBrowserSTEP.AsInteger < idStepEnd);
  NextStep.Enabled := ViewRecord.Enabled and (cdsBrowserSTEP.AsInteger < idStepTest);
  ReturnToRemount.Enabled := ViewRecord.Enabled and (cdsBrowserSTEP.AsInteger = idStepEnd);
  ShowNode.Enabled := ViewRecord.Enabled;
end;

procedure TfrmRemBrowser.FormCreate(Sender: TObject);
begin
  inherited;
  if Run_app_name = apManager then
   fcDepartment.Enabled := true
  else
   fcDepartment.Enabled := false;
  HistorySupport := TAccessLevel(ArmCurrentInfo.UserLevel) in ulAdmLevel;
  FillComboDateField;
  PopulateCombosGSD(tcNodesType.Tabs, gsdNodeM, '', ArmCurrentInfo.NodeList, '', true, false, -ofsNodeM, -ofsNodeM);
  FViewers.Add('VW_RM');
  PopulateCombosGSD(fcStep.Items, gsdRemStep, AllStr, '', '', true, false, 0, 0);
  PopulateCombosGSD(fcACCOUNTER.Items, gsdAccounter, AllStr, '', '', true, false, 0, 0);
end;

procedure TfrmRemBrowser.tbNodesTypeResize(Sender: TObject);
begin
  inherited;
  tcNodesType.Width := tcNodesType.Parent.ClientWidth;
end;

procedure TfrmRemBrowser.CustomiseFilter;
begin
  PopulateCombosGSD(fcTypeC.Items, CurNodeType + ofsType, AllStr, '', '', true, false, 0, 0);
  PopulateCombosGSD(fcModel.Items, CurNodeType + ofsmodel, AllStr, '', '', true, false, 0, 0);
  PopulateCombosDEPT(fcDepartment.Items, ALLStrDepTypes, 'Все ЦЕХА', GetListDeptWorkWithNT(CurNodeType), True);
  PopulateCombosGSD(rpGroupNode.Items, 1401, '', '', '', true, false, 0, 0);
  //fcFromDepartment.Items.Assign(fcDepartment.Items);
  PopulateCombosDEPT(fcFromDepartment.Items, ALLStrDepTypes, Allstr, '1', True);
  fcSect.Visible := CurNodeType in [1, 3];
  Label10.Visible := CurNodeType in [1];
  rpGroupNode.Visible := Label10.Visible;
end;

procedure TfrmRemBrowser.CustomiseFields;
begin
  inherited;
  case CurNodeType of
    1, 3: cdsBrowserSECT.Visible := true;
  end;
  cdsBrowserWORKDAYS.Visible := FRmBrowserMode;
  cdsBrowserINVNUM.Visible :=true;
end;

procedure TfrmRemBrowser.SetDefaultOrder;
begin
  inherited;
//  ADDOrder(cdsBrowserNum, false, false);
end;

procedure TfrmRemBrowser.tcNodesTypeChange(Sender: TObject);
begin
//  CurrentViewer:=FViewers.Strings[tcNodesType.tabindex];
  CurNodeType := GetComboID(tcNodesType.Tabs, tcNodesType.TabIndex);
  fcNODE.Text := IntToStr(CurNodeType);
  FilterChangeE(fcNODE);
  CustomiseFilter;
  ResetAllFilter;
  CustomiseGrid;
  if Sender <> nil then
   DataRefresh.Execute;
end;

procedure TfrmRemBrowser.GridGetBtnParams(Sender: TObject; Field: TField;
  AFont: TFont; var Background: TColor; var SortMarker: TSortMarker;
  IsDown: Boolean);
begin
  inherited;
  if field = cdsBrowserCTRL then
    if field.Value <> null then
      if field.Value = 'К' then AFont.Color := clred;
end;

procedure TfrmRemBrowser.FormShow(Sender: TObject);
begin
   IF Global_Caption = '' Then
   Begin
   if RmBrowserMode then Capt := 'Журнал ремонта' else
    Capt := 'Журнал входного контроля';
   end
  Else
   Begin
    Capt := Global_Caption;
   End;

  DefaultSearchField := cdsBrowserNum;
  tcNodesTypeChange(nil);
  inherited;
end;

procedure TfrmRemBrowser.ResetAllFilter;
begin
  inherited;
  faRemountType.Checked := FRmBrowserMode;
  FilterChangeE(faRemountType);
  faStep.Checked := true;
  FilterChangeE(faStep);
  FromDate.Date := 0;
  ToDate.Date := 0;
  cbDateField.ItemIndex := 0;
  cbDateFieldClick(cbDateField);
  if Run_app_name <> apManager then
   begin
    {если арм то устанавливать по умолчанию свой цех}
    fcDepartment.ItemIndex := fcDepartment.Items.IndexOfObject(TObject(ArmCurrentInfo.DepartmentID));
    FilterChangeE(fcDepartment);
   end;
end;

procedure TfrmRemBrowser.cbMainResize(Sender: TObject);
begin
  inherited;
  tbNodesType.width := tbNodesType.Parent.ClientWidth;
end;

procedure TfrmRemBrowser.SetRmBrowserMode(const Value: boolean);
begin
  FRmBrowserMode := Value;
  with fcRemountType do
    if FRmBrowserMode then
    {исключить вх. контроль}
      PopulateCombosGSD(Items, gsdRemountType, AllStr, '', '', true, false, 0, 0)
    else
    {включить только вх. контроль}
      items.AddObject('Входной контроль', TObject(idInCtrl));
end;

procedure TfrmRemBrowser.FillComboDateField;
var k: integer;
begin
  cbDateField.Items.Clear;
  for k := 0 to pred(cdsBrowser.FieldCount) do
    if (cdsBrowser.Fields[k] is TDateField) or
      (cdsBrowser.Fields[k] is TDateTimeField) then
      cbDateField.Items.AddObject(cdsBrowser.Fields[k].DisplayLabel, cdsBrowser.Fields[k]);
end;

procedure TfrmRemBrowser.cbDateFieldClick(Sender: TObject);
begin
  cbDateField.parent.Name := 'fc' + TField(cbDateField.Items.Objects[cbDateField.ItemIndex]).FieldName;
  DateFilterChange(Sender);
end;

procedure TfrmRemBrowser.AddRecordExecute(Sender: TObject);
begin
  inherited;
  SENDTOREM(false);
end;

procedure TfrmRemBrowser.AddHistRecordExecute(Sender: TObject);
begin
  inherited;
  SENDTOREM(true);
end;

procedure TfrmRemBrowser.SENDTOREM(Historycaly: Boolean);
var SendDir: string;
  NList: TStringList;
  FNodeSelMode: TNodeSelModeSet;
  function GetNodeListStr: string;
  var K: Integer;
  begin
    result := '';
    for k := 0 to Pred(NList.Count) do
      result := result + #10#13 + NList[k];
  end;
  procedure Sending;
  var RemID: Variant;
    k: integer;
    ACount: integer;
  begin
  User_ID(1);
    ACount := NList.Count;
    for k := Pred(ACount) downto 0 do begin
      RemID:=ExecuteStoreFUNC('SEND_NODE_REMOUNT',[Integer(NList.Objects[k]),
                                  WORD(FRmBrowserMode),
                                  WORD(Historycaly),
                                  ArmCurrentInfo.DepartmentID]);
      if RemID > 0 then
      begin
        DobeforeFiscal;
        FiscalInfo.RecordID := RemID;
        if Historycaly then FiscalInfo.Action := idFActCreateHist else
          FiscalInfo.Action := idFActCreate;
        FiscalInfo.Descript := NList[k];
        Fiscaler(FiscalInfo);
        NList.Delete(k);
      end;
    end;
  User_ID(0);
  end;
begin
  if FRmBrowserMode then begin
    SendDir := 'в ремонт';
    FNodeSelMode := [nsmInRem];
  end else begin
    SendDir := 'на вх. контроль';
    FNodeSelMode := [nsmInCtrl];
  end;

  if Historycaly then FNodeSelMode := FNodeSelMode + [nsmHist];

  NList := TStringList.Create;
  try
    Application.CreateForm(TfrmNodeBrowser, frmNodeBrowser);
    frmNodeBrowser.SelectNodes(FNodeSelMode, NList, CurNodeType);
    if NList.Count > 0 then
    begin
      if MicarMsg('', Format(OnSendNodeREMQry, [SendDir, GetNodeListStr]), MB_YESNO + MB_ICONQUESTION) = mrNo then Exit;
      Sending;
      while (NList.Count > 0) and
        (MicarMsg('', Format(SendNodeREMError, [SendDir, GetNodeListStr]), MB_YESNO + MB_ICONERROR) = mrYes) do
        Sending;
      DoAfterEditing(NoChange);
    end;
  finally
    NList.Free;
  end;
end;

procedure TfrmRemBrowser.EditRecordExecute(Sender: TObject);
begin
  inherited;
  Application.CreateForm(TfrmRemountEditor, frmRemountEditor);
  DoAfterEditing(frmRemountEditor.ShowRemount(FCurNodeType, cdsBrowserID.asInteger, FRmBrowserMode));
end;

procedure TfrmRemBrowser.DoStepExecute(Sender: TObject);
begin
  Application.CreateForm(TfrmRemountEditor, frmRemountEditor);
  DoAfterEditing(frmRemountEditor.DoStep(FCurNodeType, cdsBrowserID.asInteger, cdsBrowserSTEP.AsInteger, FRmBrowserMode));   //Переход на вкладку
end;

procedure TfrmRemBrowser.ReturnToRemountExecute(Sender: TObject);
begin
  if (MicarMsg('', Format(rmeTryNodeReturn, [cdsBrowserDepartment_S.AsString]),
    MB_YESNO + MB_ICONQUESTION) = mrYes) then
  begin
    if ExecuteStoreProc('ERETURN_NODE_REMOUNT',[cdsBrowserID.AsInteger,0])[1] = 0 then
      MicarMsg('', 'Не удалось вернуть узел.', MB_ICONINFORMATION)
    else
    begin
      FiscalAct(faRetRem,true);
      DoAfterEditing(cdsBrowserID.AsInteger);
    end;
  end;
end;

procedure TfrmRemBrowser.NextStepExecute(Sender: TObject);
begin
  inherited;
  if ExecuteStoreProc('ENEXTSTEP_NODE_REMOUNT', [cdsBrowserID.AsInteger,0])[1] = 0 then
    MicarMsg('', 'Не удалось.', MB_ICONINFORMATION)
  else
  begin
    FiscalAct(faNextStep,true);
    DoAfterEditing(cdsBrowserID.AsInteger);
  end;
end;

procedure TfrmRemBrowser.ShowNodeExecute(Sender: TObject);
begin
  Application.CreateForm(TfrmNodeEditor, frmNodeEditor);
  frmNodeEditor.ShowNode(FCurNodeType, cdsBrowserNODESID.asInteger);
end;

procedure TfrmRemBrowser.DoDelete(var ID: integer; var DelRes: boolean);
begin
  ID := cdsBrowserID.AsInteger;
  DelRes := ExecuteSQL(DeleteRecordByID2('RMJOURNAL', ID)) > 0;
end;

procedure TfrmRemBrowser.DeleteRecordExecute(Sender: TObject);
 var AStatus: Integer;
     DelRes: Boolean;
begin
  Application.CreateForm(TfrmSelStatNode, frmSelStatNode);
  AStatus := frmSelStatNode.ConfirmDelRecord(RmBrowserMode, cdsBrowserNODE.AsInteger);
  if AStatus > 0 then
   begin
    DelRes := ExecuteSQL(UpdateNodeStatus(AStatus, cdsBrowserNODESID.AsInteger)) > 0;
    DelRes := ExecuteSQL(DeleteRecordByID2('RMJOURNAL', cdsBrowserID.AsInteger)) > 0;
    if DelRes then
     begin
      FFiscAct := faDelete;
      FiscalIt;
      cdsBrowser.Delete;
     end
    else
     MicarMsg('', DeleteError, MB_ICONWARNING);
   end;
end;

procedure TfrmRemBrowser.DoBeforeFiscal;
begin
  inherited;
  with FiscalInfo do begin
    if RmBrowserMode then Block := idBlockRem else
      Block := idBlockCtrl;
  end;
end;

procedure TfrmRemBrowser.cxGrid1DEDBTableView1CTRLStylesGetContentStyle(
  Sender: TcxCustomGridTableView; ARecord: TcxCustomGridRecord;
  AItem: TcxCustomGridTableItem; out AStyle: TcxStyle);
begin
  inherited;
if (ARecord is TcxGridDataRow) and not ARecord.Selected and
   (ARecord.Values[cxGrid1DEDBTableView1CTRL.Index] = 'К') then
    AStyle := CTRLStyle;
end;

procedure TfrmRemBrowser.ViewRecordExecute(Sender: TObject);
begin
  inherited;
  Application.CreateForm(TfrmRemountEditor, frmRemountEditor);
  DoAfterEditing(frmRemountEditor.ShowRemount(FCurNodeType, cdsBrowserID.asInteger, FRmBrowserMode));
end;

end.

