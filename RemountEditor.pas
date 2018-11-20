unit RemountEditor;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  orDBMEditor, Db, DBClient, StdCtrls, Buttons, ExtCtrls, ComCtrls, Grids,
  DBGrids, RXDBCtrl, Mask, DBCtrls, NodeProfile, ParamList, cxLabel,
  ActnList, miSequence, ComObj, Variants, rxStrUtils, cxControls,
  cxContainer, cxEdit, cxTextEdit, cxMaskEdit, cxDropDownEdit, cxCalendar,
  cxDBEdit, rxToolEdit, cxGraphics, cxLookAndFeels, cxLookAndFeelPainters,
  ToolWin, cxStyles, cxCustomData, cxFilter, cxData, cxDataStorage, cxDBData,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGridLevel,
  cxClasses, cxGridCustomView, cxGrid, cxGroupBox, cxRadioGroup, jpeg, cxMRUEdit, ExcelXP;

type
  TfrmRemountEditor = class(TfrmorDBMEditor)
    pTitle: TPanel;
    pcRemount: TPageControl;
    tsCommon: TTabSheet;
    tsDF: TTabSheet;
    tsAsm: TTabSheet;
    tsTest: TTabSheet;
    pOW: TPanel;
    cdsMasterID: TFloatField;
    cdsMasterNODESID: TFloatField;
    cdsMasterNODEINFOBEFORE: TFloatField;
    cdsMasterREMOUNTTYPE: TFloatField;
    cdsMasterDEPARTMENT: TFloatField;
    cdsMasterSTEP: TFloatField;
    cdsMasterDATEIN: TDateTimeField;
    cdsMasterDATEASSEMBLE: TDateTimeField;
    cdsMasterDATETEST: TDateTimeField;
    cdsMasterDATEEND: TDateTimeField;
    cdsMasterRMJOURNALID: TFloatField;
    cdsMasterFROMDEPARTMENT: TFloatField;
    cdsMasterOILWELLSID: TFloatField;
    cdsMasterWORKDAYS: TFloatField;
    cdsMasterDATEDEFECT: TDateTimeField;
    cdsMasterMASTER: TFloatField;
    cdsMasterVISITOR: TStringField;
    cdsMasterNODEINFOAFTER: TFloatField;
    cdsMasterNGDU_S: TStringField;
    cdsMasterDEPOSIT_S: TStringField;
    cdsMasterKUST: TStringField;
    cdsMasterWELL: TStringField;
    cdsMasterNODE: TFloatField;
    cdsMasterSECT: TStringField;
    cdsMasterTYPEC: TFloatField;
    cdsMasterMODEL: TFloatField;
    cdsMasterLENGTHL: TFloatField;
    cdsMasterPRODUCER: TFloatField;
    cdsMasterCTRL: TStringField;
    cdsMasterNUM: TStringField;
    cdsMasterINVNUM: TStringField;
    cdsMasterNEW: TStringField;
    cdsMasterACCOUNTER: TFloatField;
    cdsMasterNODE_S: TStringField;
    cdsMasterTYPEC_S: TStringField;
    cdsMasterMODEL_S: TStringField;
    cdsMasterPRODUCER_S: TStringField;
    cdsMasterACCOUNTER_S: TStringField;
    pcCommon: TPageControl;
    tsNPBefore: TTabSheet;
    tsNPAfter: TTabSheet;
    cdsNPAfter: TClientDataSet;
    cdsNPAfterNODE: TFloatField;
    cdsNPAfterNODE_S: TStringField;
    cdsNPAfterSECT: TStringField;
    cdsNPAfterTYPEC: TFloatField;
    cdsNPAfterTYPEC_S: TStringField;
    cdsNPAfterMODEL: TFloatField;
    cdsNPAfterMODEL_S: TStringField;
    cdsNPAfterLENGTHL: TFloatField;
    cdsNPAfterPRODUCER: TFloatField;
    cdsNPAfterPRODUCER_S: TStringField;
    cdsNPAfterCTRL: TStringField;
    cdsNPAfterNUM: TStringField;
    cdsNPAfterINVNUM: TStringField;
    cdsNPAfterNEW: TStringField;
    cdsNPAfterACCOUNTER: TFloatField;
    cdsNPAfterACCOUNTER_S: TStringField;
    Label5: TLabel;
    cdsMasterDEPARTMENT_S: TStringField;
    cdsMasterREMOUNTTYPE_S: TStringField;
    cdsMasterSTEP_S: TStringField;
    cdsMasterFROMDEPARTMENT_S: TStringField;
    cdsMasterASMWORKER: TFloatField;
    cdsMasterTESTWORKER: TFloatField;
    cdsMasterREM: TStringField;
    cdsMasterDFWORKER: TFloatField;
    cdsMasterDFREM: TStringField;
    Label8: TLabel;
    DBText1: TDBText;
    DBText2: TDBText;
    DBText3: TDBText;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    DBEdit1: TDBEdit;
    DBText4: TDBText;
    DBText5: TDBText;
    DBText6: TDBText;
    DBText7: TDBText;
    lRemount: TLabel;
    Label4: TLabel;
    DBLookupComboBox2: TDBLookupComboBox;
    plNodeProfile: TPanel;
    Label16: TLabel;
    plPrmList: TPanel;
    plNP: TPanel;
    Bevel1: TBevel;
    cdsDetales: TClientDataSet;
    cdsDetalesKEYGRP: TFloatField;
    cdsDetalesSORT1: TFloatField;
    cdsDetalesSORT2: TFloatField;
    cdsDetalesSORT3: TFloatField;
    cdsDetalesREM: TStringField;
    srcDetales: TDataSource;
    cdsDetalesDETALE_S: TStringField;
    cdsDetalesSum: TIntegerField;
    cdsDetalesBrack: TFloatField;
    cdsMasterDFWORKER_S: TStringField;
    cdsMasterASMWORKER_S: TStringField;
    cdsMasterTESTWORKER_S: TStringField;
    cdsMasterMASTER_S: TStringField;
    plComInfo: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    lFromDept: TLabel;
    tsCabel: TTabSheet;
    ActionList1: TActionList;
    SelectOW: TAction;
    DeleteRaport: TAction;
    NewRaport: TAction;
    EditRaport: TAction;
    DeleteDetales: TAction;
    CreateDetales: TAction;
    Label3: TLabel;
    DBEdit2: TDBEdit;
    plInfoRem: TPanel;
    dbmDfRem: TDBMemo;
    plFoot: TPanel;
    Bevel2: TBevel;
    plAsmInfo: TPanel;
    Label18: TLabel;
    dbceAsm: TRxDBComboEdit;
    plDfInfo: TPanel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    dbceMaster: TRxDBComboEdit;
    DBEdit6: TDBEdit;
    dbceDefector: TRxDBComboEdit;
    plTestnfo: TPanel;
    Label19: TLabel;
    dbceTester: TRxDBComboEdit;
    Panel3: TPanel;
    Label17: TLabel;
    DBStepDateEdit: TDBDateEdit;
    splDFRem: TSplitter;
    plDetales: TPanel;
    Bevel4: TBevel;
    Panel6: TPanel;
    plRemParams: TPanel;
    plRemPrmList: TPanel;
    Panel2: TPanel;
    Splitter: TSplitter;
    PrintRDump: TAction;
    PreviewRDump: TAction;
    ShowNPassp: TAction;
    ShowEpuPassp: TAction;
    EditDetales: TAction;
    cdsMasterREAL_LENGTHL: TFloatField;
    cdsMasterMECH_POVREGD: TFloatField;
    cdsMasterSROSTKOV: TFloatField;
    cdsMasterUDLINIT: TFloatField;
    cdsMasterCENA: TFloatField;
    cdsNPAfterREAL_LENGTHL: TFloatField;
    cdsNPAfterMECH_POVREGD: TFloatField;
    cdsNPAfterSROSTKOV: TFloatField;
    cdsNPAfterUDLINIT: TFloatField;
    cdsNPAfterCENA: TFloatField;
    cdsMasterGRUPPA: TFloatField;
    cdsNPAfterGRUPPA: TFloatField;
    Label7: TLabel;
    DebTrigger: TAction;
    cxDBDateEdit1: TcxDBDateEdit;
    cdsMasterGRUPPA_S: TStringField;
    dblkcbbGRUPPA: TDBLookupComboBox;
    pnlVals: TPanel;
    GroupBox1: TGroupBox;
    ToolBar1: TToolBar;
    btnAddVal: TToolButton;
    Panel7: TPanel;
    plBtn: TPanel;
    btnCreateDetales: TSpeedButton;
    btnDeleteDetales: TSpeedButton;
    btnEditDetales: TSpeedButton;
    gDetales: TRxDBGrid;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1: TcxGrid;
    cxGrid1DBTableView1ID: TcxGridDBColumn;
    cxGrid1DBTableView1nomer: TcxGridDBColumn;
    cxGrid1DBTableView1max_tekuchest: TcxGridDBColumn;
    cxGrid1DBTableView1marka_stali: TcxGridDBColumn;
    cxGrid1DBTableView1diametr: TcxGridDBColumn;
    cxGrid1DBTableView1length: TcxGridDBColumn;
    cxGrid1DBTableView1producer_s: TcxGridDBColumn;
    cxGrid1DBTableView1new: TcxGridDBColumn;
    cxGrid1DBTableView1date_prihod: TcxGridDBColumn;
    btnRemoveVal: TToolButton;
    cxGrid1DBTableView1date_remove: TcxGridDBColumn;
    cxGrid1DBTableView1is_current_val: TcxGridDBColumn;
    cxGrid1DBTableView1date_add: TcxGridDBColumn;
    btnPassport: TToolButton;
    cxDateEnd: TcxDBDateEdit;
    CDSCabCabel: TClientDataSet;
    CDSCabCabelID: TFloatField;
    CDSCabCabelTYPEC_S: TStringField;
    CDSCabCabelGRUPPA: TStringField;
    CDSCabCabelOWNER_S: TStringField;
    CDSCabCabelMODEL_S: TStringField;
    CDSCabCabelNUM: TStringField;
    CDSCabCabelINVNUM: TStringField;
    CDSCabCabelLENGTHL: TFloatField;
    CDSCabCabelREAL_LENGTHL: TFloatField;
    CDSCabCabelPRODUCER_S: TStringField;
    CDSCabCabelNEW: TStringField;
    CDSCabCabelCENA: TFloatField;
    CDSCabCabelAGREGATID: TFloatField;
    CDSCabCabelNODE: TFloatField;
    CDSCabCabelNODE_S: TStringField;
    CDSCabCabelTYPEC: TFloatField;
    CDSCabCabelMODEL: TFloatField;
    CDSCabCabelACCOUNTER: TFloatField;
    CDSCabCabelOWNER: TFloatField;
    CDSCabCabelPRODUCER: TFloatField;
    CDSCabCabelCTRL: TStringField;
    CDSCabCabelNODEINFO: TFloatField;
    CDSCabCabelDATEPROD: TDateTimeField;
    CDSCabCabelDATEIN: TDateTimeField;
    CDSCabCabelDATETRASH: TDateTimeField;
    CDSCabCabelCURSTATUS: TFloatField;
    CDSCabCabelCURSTATUS_S: TStringField;
    CDSCabCabelDEPARTMENT: TFloatField;
    CDSCabCabelDEPARTMENT_S: TStringField;
    CDSCabCabelTODEPARTMENT: TFloatField;
    CDSCabCabelTODEPARTMENT_S: TStringField;
    CDSCabCabelTEMPID: TFloatField;
    CDSCabCabelMECH_POVREGD: TFloatField;
    CDSCabCabelSROSTKOV: TFloatField;
    CDSCabCabelUDLINIT: TFloatField;
    cdsCabSpis: TClientDataSet;
    FloatField1: TFloatField;
    FloatField2: TFloatField;
    FloatField3: TFloatField;
    cdsMasterRAPORTTYPE: TFloatField;
    cdsMasterRAPORTTYPE_S: TStringField;
    FloatField4: TFloatField;
    StringField1: TStringField;
    cdsMasterDATERAPORT: TDateTimeField;
    FloatField5: TFloatField;
    StringField2: TStringField;
    FloatField6: TFloatField;
    StringField3: TStringField;
    FloatField7: TFloatField;
    StringField4: TStringField;
    cdsMasterOWNER: TFloatField;
    cdsMasterOWNER_S: TStringField;
    FloatField8: TFloatField;
    cdsMasterFROMCAT: TFloatField;
    cdsMasterFROMCAT_S: TStringField;
    cdsMasterTOCAT: TFloatField;
    cdsMasterTOCAT_S: TStringField;
    cdsMasterREASON: TFloatField;
    cdsMasterREASON_S: TStringField;
    FloatField9: TFloatField;
    StringField5: TStringField;
    StringField6: TStringField;
    StringField7: TStringField;
    DSCabSpis: TDataSource;
    DSCabCabel: TDataSource;
    PanelCabCabel: TPanel;
    BtDiv: TButton;
    BtSpis: TButton;
    BtDelete: TButton;
    BtAdd: TButton;
    Panel12: TPanel;
    Label56: TLabel;
    LabCabLength: TLabel;
    Panel13: TPanel;
    Label57: TLabel;
    LabCabSrost: TLabel;
    Panel14: TPanel;
    LabCabKusk: TLabel;
    Label58: TLabel;
    plCableDIV: TPanel;
    Image1: TImage;
    Panel8: TPanel;
    CabLabel: TLabel;
    PlNoActive: TPanel;
    Label28: TLabel;
    Label24: TLabel;
    Label23: TLabel;
    Label35: TLabel;
    Label36: TLabel;
    cxRGNewRem: TcxRadioGroup;
    DBEProducerT: TDBEdit;
    DBETypecT: TDBEdit;
    DBEModelT: TDBEdit;
    DBEDateend: TDBEdit;
    DBEDatein: TDBEdit;
    PlInvnum: TPanel;
    Label25: TLabel;
    Label26: TLabel;
    ENum: TEdit;
    EInvNum: TEdit;
    PlLength: TPanel;
    Label29: TLabel;
    ERealLength: TEdit;
    cbSostCabLin: TCheckBox;
    PlSostCab: TPanel;
    Label37: TLabel;
    Label31: TLabel;
    Label32: TLabel;
    lblCurStat: TLabel;
    CBCurstatus: TComboBox;
    PlSpis: TPanel;
    Label52: TLabel;
    Label53: TLabel;
    Label54: TLabel;
    Label51: TLabel;
    Label55: TLabel;
    DBLookupComboBox8: TDBLookupComboBox;
    RxDBComboEdit1: TRxDBComboEdit;
    DBLookupComboBox7: TDBLookupComboBox;
    DBLookupComboBox6: TDBLookupComboBox;
    DBLookupComboBox9: TDBLookupComboBox;
    Panel9: TPanel;
    Label40: TLabel;
    Label41: TLabel;
    Label42: TLabel;
    Label43: TLabel;
    Label44: TLabel;
    Label45: TLabel;
    Label46: TLabel;
    Label49: TLabel;
    Label50: TLabel;
    Label33: TLabel;
    eRemID: TDBEdit;
    DBENew: TDBEdit;
    Nodeinfoid: TDBEdit;
    Nodeid: TDBEdit;
    DBERealLength: TDBEdit;
    DBEdit3: TDBEdit;
    DBEdit4: TDBEdit;
    DBEdit5: TDBEdit;
    DBEdit7: TDBEdit;
    DBEdit8: TDBEdit;
    DBEdit9: TDBEdit;
    DBEdit10: TDBEdit;
    DBEGroup: TDBEdit;
    DBEdit13: TDBEdit;
    CBDepartment: TDBEdit;
    DBETypec: TDBEdit;
    DBEModel: TDBEdit;
    DBEProducer: TDBEdit;
    Panel10: TPanel;
    Panel11: TPanel;
    MSrcAfter: TDataSource;
    Label27: TLabel;
    cdsCabSpisTIER_NUM: TStringField;
    DBEdit11: TDBEdit;
    Label30: TLabel;
    NPAfterOwner: TEdit;
    cxDBRadioGroup1: TcxDBRadioGroup;
    dblcbRemount: TDBLookupComboBox;
    sbPrintOTK: TSpeedButton;
    btnPreviewRDump: TSpeedButton;
    sbEpuPassport: TSpeedButton;
    sbNodePassport: TSpeedButton;
    cdsVals: TClientDataSet;
    cdsValsID: TFloatField;
    cdsValsnomer: TStringField;
    cdsValsdate_in: TDateTimeField;
    cdsValsdate_out: TDateTimeField;
    cdsValsmax_tekuchest: TFloatField;
    cdsValsmarka_stali: TStringField;
    cdsValsdiametr: TFloatField;
    cdsValslength: TFloatField;
    cdsValsproducer_s: TStringField;
    cdsValsnew: TStringField;
    cdsValsdate_prihod: TDateTimeField;
    cdsValsis_current_val: TFloatField;
    cdsValsrmjournal_id: TFloatField;
    dsVals: TDataSource;
    InCtrlStat: TcxDBRadioGroup;
    cdsMasterOUTSTATUS: TFloatField;
    dbFromDept: TcxDBTextEdit;
    cxRadioGroup1: TcxRadioGroup;
    NPAfterOwner2: TEdit;
    DBEDep: TDBEdit;
    DBEdit12: TDBEdit;
    CDSCabCabelUDLINIT_SET: TStringField;
    cxGrid2: TcxGrid;
    cxGrid2DBTableView1: TcxGridDBTableView;
    cxGrid2DBTableView1ID: TcxGridDBColumn;
    cxGrid2DBTableView1Priznak: TcxGridDBColumn;
    cxGrid2DBTableView1TYPEC_S: TcxGridDBColumn;
    cxGrid2DBTableView1MODEL_S: TcxGridDBColumn;
    cxGrid2DBTableView1NUM: TcxGridDBColumn;
    cxGrid2DBTableView1INVNUM: TcxGridDBColumn;
    cxGrid2DBTableView1REAL_LENGTHL: TcxGridDBColumn;
    cxGrid2DBTableView1GRUPPA: TcxGridDBColumn;
    cxGrid2DBTableView1OWNER_S: TcxGridDBColumn;
    cxGrid2DBTableView1PRODUCER_S: TcxGridDBColumn;
    cxGrid2DBTableView1NEW: TcxGridDBColumn;
    cxGrid2DBTableView1CENA: TcxGridDBColumn;
    cxGrid2Level1: TcxGridLevel;
    cdsNPAfterGRUPPA_S: TStringField;
    procedure CustomiseEditor; override;
    function DoPost: boolean; override;
    procedure SetSpecifyLookUpTag(DataSet: TDataSet); override;
    function GetInputDataRequest(DataSet: TDataSet): OleVariant; override;
    procedure DoBeforeFiscal; override;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ActivePageChange(Sender: TObject);
    procedure cdsDetalesCalcFields(DataSet: TDataSet);
    procedure FormShow(Sender: TObject);
    procedure cdsMasterREMOUNTTYPEValidate(Sender: TField);
    procedure cdsMasterSTEPValidate(Sender: TField);
    procedure cdsMasterCalcFields(DataSet: TDataSet);
    procedure DeleteDetalesExecute(Sender: TObject);
    procedure SelectOWExecute(Sender: TObject);
    procedure CreateDetalesExecute(Sender: TObject);
    procedure cdsNPAfterNUMValidate(Sender: TField);
    procedure PrintRDumpExecute(Sender: TObject);
    procedure ShowNPasspExecute(Sender: TObject);
    procedure ShowEpuPasspExecute(Sender: TObject);
    procedure MSrcDataChange(Sender: TObject; Field: TField);
    procedure EditDetalesExecute(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure OkBtnClick(Sender: TObject);
    procedure btDefectDblClick(Sender: TObject);
    procedure sbPrintOTKClick(Sender: TObject);
    procedure DebTriggerExecute(Sender: TObject);
    procedure btnAddValClick(Sender: TObject);
    procedure btnRemoveValClick(Sender: TObject);
    procedure cxGrid1DBTableView1CustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
    procedure btnPassportClick(Sender: TObject);
    procedure cxGrid1DBTableView1CellDblClick(Sender: TcxCustomGridTableView;
      ACellViewInfo: TcxGridTableDataCellViewInfo; AButton: TMouseButton;
      AShift: TShiftState; var AHandled: Boolean);
    procedure cdsValsAfterOpen(DataSet: TDataSet);
    procedure cdsValsdate_inChange(Sender: TField);
    procedure BtDivClick(Sender: TObject);
    procedure BtSpisClick(Sender: TObject);
    procedure BtDeleteClick(Sender: TObject);
    procedure BtAddClick(Sender: TObject);
    procedure CDSCabCabelAfterPost(DataSet: TDataSet);
    procedure CDSCabCabelBeforeDelete(DataSet: TDataSet);
    procedure CDSCabCabelBeforeOpen(DataSet: TDataSet);
    procedure FKeyPress(Sender: TObject; var Key: Char);
    procedure FKeyPressNumber(Sender: TObject; var Key: Char);
    procedure cbSostCabLinClick(Sender: TObject);
    procedure ERealLengthChange(Sender: TObject);
    procedure tsCabelShow(Sender: TObject);
    procedure CabRemoveClick(Sender: TObject);
    procedure cdsMasterReconcileError(DataSet: TCustomClientDataSet;
                        E: EReconcileError; UpdateKind: TUpdateKind;
                        var Action: TReconcileAction);
    procedure cdsMasterREMOUNTTYPEChange(Sender: TField);
    procedure InCtrlStatPropertiesChange(Sender: TObject);
    procedure cxDBDateEdit1PropertiesChange(Sender: TObject);
    procedure cxDateEndPropertiesChange(Sender: TObject);
    procedure cxGrid2DBTableView1PriznakGetPropertiesForEdit(
      Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
      var AProperties: TcxCustomEditProperties);
    procedure cxGrid2DBTableView1PriznakPropertiesChange(Sender: TObject);
    procedure CDSCabCabelUDLINIT_SETGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure CDSCabCabelUDLINIT_SETSetText(Sender: TField; const Text: string);
    procedure cdsMasterDATEINChange(Sender: TField);
  private
    { Private declarations }
    debug: TStringList;
    idb: boolean;
    FRMEditorMode: boolean;
    FNodeProfile: TfrmNodeProfile;
    FNT: Integer;
    FDoStep: integer;
    FHistoricaly: boolean;
    FTempOWInfo: Variant;
    FOWID: Integer;
    FParamList: TfrmParamList;
    FLengthl: TEdit;
    FNPAfterNumModified: boolean;
    FNPBeforeNumModified: boolean;
    procedure SetRMEditorMode(const Value: boolean);
    function PageActive(ts: TTabSheet): boolean;
    procedure BeforeDoStep;
    function CheckValidRem: Boolean;
    function IsDoubleNum: Boolean;
    procedure DebugPoint1;
    procedure date_update;
    procedure cdsCabRef;
    function GetNPInfoRemEnd(RMid: Integer): Integer;
    function GetLengthlCablineBeforeRemount(CablineID, rmID: Integer): double;
  public
    { Public declarations }
    function ShowRemount(NT, ID: Integer; RM: Boolean): Integer;
    function DoStep(NT, ID, Step: Integer; RM: Boolean): Integer;
    property RMEditorMode: boolean read FRMEditorMode write SetRMEditorMode;
  end;

const del_rem = 'BEGIN' +
                ' DELETE FROM RMREMOUNT WHERE RMJOURNALID = %s;' +
                ' DELETE FROM RMJOURNAL WHERE ID = %s; ' +
                'END;';
      upd_kablenodes_info = 'UPDATE kablenodes SET nodeinfo = %d WHERE agregatid = %d AND nodesid = %d';
      MaxCutCabLength = 1;
      MsgCabOperation = 'Операция невозможна для пустой кабельной линии.';

var
  frmRemountEditor: TfrmRemountEditor;
  CabCut: Boolean = False;
  ProcUpdate: Boolean;
  RemoveCabFromCablin: Boolean = False;
  br: Boolean;
  SostCabLin: Boolean = False;
  RemIDDel, IDCabLinForDIV: integer;
  FLengthl_After, FLengthl_Before: Double;
  RealMaxLen, MaxLen, MinLen, CurrentLen, CurrentRealLen: Double;
  TempCabLineID, TempNPAfterID: integer; // Для временного хранения ИД кабельной линии

implementation

uses MSql, MMSG, MFDumpS, RxVclUtils, globals, uConnData, NodeEditor, OWBrowser,
     CabelRaportEditor, EPUEditor, RemBrowser, uPrintDF, NodeBrowser, ExcelNZ,
     uSettings, uListValy, uValy, SelStatNode, CabLin1;

{$R *.DFM}

{ TfrmRemountEditor }

function TfrmRemountEditor.GetNPInfoRemEnd(RMid: Integer): Integer;
 const z_rm_out_info = 'SELECT nodeinfoafter FROM rmremount WHERE rmjournalid = %d';
begin
  Result := CustomSQL(Format(z_rm_out_info, [RMid]))[0];
end;

function TfrmRemountEditor.ShowRemount(NT, ID: Integer; RM: Boolean): Integer;
begin
  FDoStep := NoChange;
  pcRemount.ActivePageIndex := 0;
  FNT := NT;
  RMEditorMode := RM;
  tsCabel.TabVisible := False;
  Result := ShowEditor(ID, true);
end;

procedure TfrmRemountEditor.tsCabelShow(Sender: TObject);
 function RazborNomera(NID: string): string;
  begin
   Result := VarToStr(CustomSQL('SELECT get_cable_num('+NID+') FROM dual')[0]);
  end;
  var TempNodeInfo: Variant;
      TmpAgnCabl: Variant;
begin
  with cxRGNewRem do
   if DBENew.Text = 'н' then
    ItemIndex := 0
   else
    if DBENew.Text = 'р' then
     ItemIndex := 1;
  TempNodeInfo := CustomSQL('SELECT owner, owner_s FROM vw_nodes WHERE id = '+cdsMasterNODESID.AsString);
  if not VarIsNull(TempNodeInfo[1]) then
   begin
    NPAfterOwner.Text := TempNodeInfo[1];
    NPAfterOwner2.Text := TempNodeInfo[1];
    NPAfterOwner2.Tag := TempNodeInfo[0];
   end;
  MaxLen := cdsNPAfterLENGTHL.AsFloat;
  MinLen := 1;
  CurrentRealLen := cdsNPAfterREAL_LENGTHL.AsFloat;
  CurrentLen := cdsNPAfterLENGTHL.AsFloat;
  //--  Условие для выбора статусов кабельной линии
  PopulateCombosGSD(CBCurstatus.Items, 12, '', '1027, 1011, 1028', '', true, false, 0, 0);
  CBCurstatus.ItemIndex := CBCurstatus.Items.IndexOfObject(TObject(idcsReadyForCabl));
  if CabCut then
   if Copy(CabLabel.Caption, 0, 1) = 'Н' then
    begin
     ENum.Text := RazborNomera(cdsMasterNODESID.AsString);
     if cxRGNewRem.ItemIndex = 0 then
      begin
       TmpAgnCabl := CustomSQL(Format(SqlAgnCabl,[cdsMasterNODESID.AsInteger]));
       if TmpAgnCabl[0] > 0 then
        cxRGNewRem.ItemIndex := 1;
      end;
    end
   else
    begin
     if not (cdsCabSpis.State in [dsEdit, dsInsert]) then
      cdsCabSpis.Edit;
     cdsMasterOWNER.Value := TempNodeInfo[0];
    end;
end;

function TfrmRemountEditor.DoStep(NT, ID, Step: Integer; RM: Boolean): Integer;
begin
  FDoStep := Step;
  pcRemount.ActivePageIndex := Step - 400;
  FNT := NT;
  RMEditorMode := RM;
  tsCabel.TabVisible := False;
  Result := ShowEditor(ID, true);
end;

procedure TfrmRemountEditor.BeforeDoStep;
var fDateField: TField;
begin
  if FDoStep > 0 then begin
    fDateField := nil;
    cdsMaster.Edit;
    case FDoStep of
      idStepDF: fDateField := cdsMasterDATEDEFECT;
      idStepAsm: fDateField := cdsMasterDATEASSEMBLE;
      idStepTest: fDateField := cdsMasterDATETEST;
    end;
    if fDateField.IsNull then fDateField.Value := Date;
  end;
end;

procedure TfrmRemountEditor.BtAddClick(Sender: TObject);
  var
   NList: TStringList;
   FNodeSelMode: TNodeSelModeSet;
   k: integer;
 procedure PrepareNList;
  begin
   with CDSCabCabel do
    begin
     DisableControls;
     try
      First;
      while not eof do
       begin
        NList.AddObject('IS', TObject(FieldByName('ID').AsInteger));
        Next;
       end;
     finally
      EnableControls;
     end;
    end;
  end;
begin
 StartWait;
 Tab_ID := '6';
 Application.CreateForm(TfrmNodeBrowser, frmNodeBrowser);
 FNodeSelMode := [nsmCablin];
 NList := TStringList.Create;
 try
  CDSCabCabel.Open;
  PrepareNList;
  frmNodeBrowser.SelectNodes(FNodeSelMode, NList, idCabelLine);
  Update;
  for k := 0 to Pred(NList.Count) do
   if NList[k] = '' then
    with CDSCabCabel do
     begin
      Insert;
      CDSCabCabelID.AsInteger := Integer(NList.Objects[k]);
      CDSCabCabelAGREGATID.AsInteger:= cdsMasterNODESID.AsInteger;
      CopyCablin(Integer(NList.Objects[k]), CDSCabCabel);
      if CDSCabCabelID.AsInteger=0 then CDSCabCabel.Cancel;
      Post;
      CDSCabCabel.ApplyUpdates(0);
      CDSCabCabel.Close;
      CDSCabCabel.Open;
     end;
 finally
  NList.Free;
 end;
 StopWait;
end;

procedure TfrmRemountEditor.cdsCabRef;
 var k: integer;
     i: Double;
begin
 if not CDSCabCabel.Active then
  begin
   FLengthl_After := 0;
   LabCabLength.Caption := FloatToStr(0);
   LabCabKusk.Caption := IntToStr(0);
   LabCabSrost.Caption := IntToStr(0);
   Exit;
  end;
 i:= 0; k:= 0;
 cdsCabCabel.First;
 while not cdsCabCabel.Eof do
  begin
   k:= k+1;
   i:= i + CDSCabCabelREAL_LENGTHL.AsFloat;
   cdsCabCabel.Next;
  end;
 cdsCabCabel.First;
 FLengthl_After := i;
 LabCabLength.Caption := FloatToStr(i);
 LabCabKusk.Caption := IntToStr(k);
 if k = 0 then
  LabCabSrost.Caption := IntToStr(k)
 else
  LabCabSrost.Caption := IntToStr(k-1);
end;

function TfrmRemountEditor.CheckValidRem: Boolean;
begin
  result := true;
  if (cdsMaster.ReadOnly) or (cdsMaster.State = dsBrowse) then Exit;
  {Переход на след. этап автоматом}
  if (FDoStep > 0) and (FDoStep = cdsMasterSTEP.AsInteger) then
    cdsMasterSTEP.AsInteger := FDoStep + 1;
  {idStepDF и idrmUnKnown}
  if (cdsMasterSTEP.AsInteger <> idStepDF) and
    (cdsMasterREMOUNTTYPE.AsInteger = idrmUnKnown) then
  begin
    MicarMsg('', Format(rmeRMTypeUnKnown, [cdsMasterSTEP_S.AsString]), MB_ICONWARNING);
    ActiveControl := dblcbRemount;
    result := false;
    exit;
  end;
  {Списание idrmTrash}
  if not FHistoricaly and (cdsMasterREMOUNTTYPE.AsInteger = idrmTrash) then
    cdsMasterSTEP.AsInteger := idStepEnd;
  {Разкомплектация}
  if not FHistoricaly and (cdsMasterREMOUNTTYPE.AsInteger = idrmDismantle) then
    cdsMasterSTEP.AsInteger := idStepEnd;
  {Завершение idStepEnd}
  if (cdsMasterDATEEND.IsNull) and
    not FHistoricaly and
    (cdsMasterStep.AsInteger = idStepEnd) then
  begin
    MicarMsg('', rmeCheckDateEnd, MB_ICONWARNING);
    result := false;
    pcRemount.ActivePageIndex := 0;
//    if FRMEditorMode then
      pcCommon.ActivePageIndex := 1;
    ActivePageChange(pcRemount);
    ActiveControl := cxDateEnd;
    exit;
  end;

  {Завершение вх. контроля}
  if (cdsMasterREMOUNTTYPE.AsInteger = idInCtrl) and (cdsMasterStep.AsInteger = idStepEnd) then
   if cdsMasterOUTSTATUS.AsInteger <> idcsNoCtrl then
    if cdsMasterOUTSTATUS.AsInteger <> idcsReady then
     if cdsMasterOUTSTATUS.AsInteger <> idcsReadyForCabl then
      if cdsMasterOUTSTATUS.AsInteger <> idcsReadyForNO then
       if cdsMasterOUTSTATUS.AsInteger <> idcsPretenzia then
        begin
         MicarMsg('', 'Укажите статус входного контроля', MB_ICONWARNING);
         result := false;
         pcRemount.ActivePageIndex := 0;
         pcCommon.ActivePageIndex := 1;
         ActivePageChange(pcRemount);
         ActiveControl := InCtrlStat;
         InCtrlStat.Style.Color := $00D9D9F4;
         exit;
        end;
  {Завершение ремонта для каб. линии}
  if (cdsMasterREMOUNTTYPE.AsInteger = idrmCur) and (cdsMasterStep.AsInteger = idStepEnd) and (cdsMasterNODE.AsInteger = idCabLine) then
   if cdsMasterOUTSTATUS.AsInteger <> idcsReady then
    if cdsMasterOUTSTATUS.AsInteger <> idcsReadyForNO then
      begin
       MicarMsg('', 'Укажите статус результата ремонта', MB_ICONWARNING);
       result := false;
       pcRemount.ActivePageIndex := 0;
       pcCommon.ActivePageIndex := 1;
       ActivePageChange(pcRemount);
       ActiveControl := InCtrlStat;
       InCtrlStat.Style.Color := $00D9D9F4;
       exit;
      end;

{  if not FHistoricaly and
     (cdsMasterStep.AsInteger=idStepEnd) then
  begin
    if ((cdsMasterREMOUNTTYPE.AsInteger=idrmTrash) and
       (MicarMsg('',rmeNodeTrash,MB_YESNO+MB_ICONQUESTION)=mrNo)) or
       ((cdsMasterREMOUNTTYPE.AsInteger<>idrmTrash) and
       (MicarMsg('',rmeNodeReady,MB_YESNO+MB_ICONQUESTION)=mrNo))then
       begin
         result:=false;
         exit;
       end;
   end;}

  {Попытка возврата узла в ремонт}
  if FHistoricaly and (cdsMasterSTEP.AsInteger <> idStepEnd) and
    (MicarMsg('', Format(rmeTryNodeReturn, [cdsMasterDepartment_S.AsString]),
    MB_YESNO + MB_ICONQUESTION) = mrNo) then
  begin
    result := false;
    exit;
  end;
end;

procedure TfrmRemountEditor.CustomiseEditor;
 var
  TmpAgnCabl: Variant;
begin
  inherited;
  {изменение порядка DS в списке после открытия всех DS
  cdsNpAfter должен быть первым в списке и поститься первыв. т.к.
  его информация иногда используется во время поста cdsMaster}
  MoveDSOrderListIndex(cdsNpAfter, 0);
  btnDeleteDetales.Caption := '';
  btnCreateDetales.Caption := '';
  btnEditDetales.Caption := '';
  btnEditDetales.Enabled := (ArmCurrentInfo.PagesList.Count <> 0);
  {если этап "Завершен" значит запись историческая}
  FHistoricaly := cdsMasterSTEP.AsInteger = idStepEnd;
  cdsMasterDEPARTMENT.ReadOnly := not (TaccessLevel(ArmCurrentInfo.UserLevel) in ulAdmLevel);
  SelectOW.Visible := not cdsMaster.ReadOnly;
  FNodeProfile.dbrgSect.Visible := FNT in [1, 3];
  if Run_app_name = apCR then
   begin
    DeleteRaport.Visible := TaccessLevel(ArmCurrentInfo.UserLevel) in ulAllWorker;
    NewRaport.Visible := DeleteRaport.Visible;
   end
  else
   begin
    DeleteRaport.Visible := false;
    NewRaport.Visible := DeleteRaport.Visible;
   end;
  DeleteDetales.Visible := TaccessLevel(ArmCurrentInfo.UserLevel) in ulAllWorker;
  CreateDetales.Visible := DeleteDetales.Visible;
  EditDetales.Visible := DeleteDetales.Visible;
  pOW.Visible := RMEditorMode;
  lRemount.Visible := RMEditorMode;
  dblcbRemount.Visible := RMEditorMode;
  tsDF.TabVisible := RMEditorMode;
  tsAsm.TabVisible := RMEditorMode or (not RMEditorMode and (FNT = 6));
  lFromDept.Visible := RMEditorMode;
  dbFromDept.Visible := RMEditorMode;
  if FNT in [idCabelLine, idCabLine] then // Для кабеля и кабельной линии заполняем другой справочник групп исполнения (СРМ)
   begin
    TClientDataSet(cdsMasterGRUPPA_S.LookupDataSet).Data := SelectSQL(grp_for_kab);
    TClientDataSet(cdsNPAfterGRUPPA_S.LookupDataSet).Data := SelectSQL(grp_for_kab);
   end;
  if CabCut then
   begin
    tsCabel.TabVisible := FNT = 6;
    tsAsm.TabVisible := False;
    tsTest.TabVisible := False;
    pcRemount.ActivePage := tsCabel;
    dblcbRemount.Visible := False;
    lRemount.Visible := False;
    Label15.Visible := False;
    DBEdit1.Visible := False;
    DBText6.Visible := False;
    DBText7.Visible := False;
    Label12.Visible := False;
    Label11.Visible := False;
    DBText5.Visible := False;
    DBText4.Visible := False;
    Label13.Visible := False;
    Label14.Visible := False;
   end;
  if RemoveCabFromCablin then
   begin
    tsCabel.TabVisible := False;
    tsAsm.TabVisible := False;
    tsTest.TabVisible := False;
    pcRemount.ActivePage := tsDF;
    dblcbRemount.Visible := False;
    lRemount.Visible := False;
    Label15.Visible := False;
    DBEdit1.Visible := False;
    DBText6.Visible := False;
    DBText7.Visible := False;
    Label12.Visible := False;
    Label11.Visible := False;
    DBText5.Visible := False;
    DBText4.Visible := False;
    Label13.Visible := False;
    Label14.Visible := False;
    if cdsMasterNEW.AsString = 'н' then
     begin
      TmpAgnCabl := CustomSQL(Format(SqlAgnCabl,[cdsMasterNODESID.AsInteger]));
      if TmpAgnCabl[0] > 0 then
       begin
        cdsNPAfter.Edit;
        cdsNPAfterNEW.AsString := 'р'
       end;
     end;
   end;
end;

procedure TfrmRemountEditor.cxDBDateEdit1PropertiesChange(Sender: TObject);
begin
  // "Дата начала ремонта" не может быть больше текущей даты.
  if (cxDBDateEdit1.SelLength > 0) and RMEditorMode then
     begin
     if (StrToDateTime ( cxDBDateEdit1.Text) > Date) or (StrToDateTime ( cxDBDateEdit1.Text) < StrToDate('01.01.2005'))then
        Begin
         ShowMessage('Дата начала ремонта должна находиться в пределах от 1.01.2005 до ' +DateTimeToStr(Now));
         cxDBDateEdit1.Date:= Now;
        end;
      end;
   if cxDBDateEdit1.SelLength > 0 then
     begin
     if (StrToDateTime ( cxDBDateEdit1.Text) > Date) or (StrToDateTime ( cxDBDateEdit1.Text) < StrToDate('01.01.2005'))then
        Begin
         ShowMessage('Дата начала входного контроля должна находиться в пределах от 1.01.2005 до ' +DateTimeToStr(Now));
         cxDBDateEdit1.Date:= Now;
        end;
      end;
end;

procedure TfrmRemountEditor.cxDateEndPropertiesChange(Sender: TObject);
begin
   // "Дата завершения ремонта" узла не может быть больше текущей даты и не может быть меньше 2005 года.
  if (cxDateEnd.SelLength > 0) and RMEditorMode then
     begin
     if (StrToDateTime ( cxDateEnd.Text) > Date) or (StrToDateTime ( cxDateEnd.Text) < StrToDate('01.01.2005')) or (StrToDateTime ( cxDateEnd.Text) < StrToDateTime ( cxDBDateEdit1.Text)) then
        Begin
         ShowMessage('Дата завершения ремонта должна находиться в пределах от ' +cxDBDateEdit1.Text + ' до ' +DateTimeToStr(Now));
         cxDateEnd.Date:= Now;
        end;
      end;
   if cxDateEnd.SelLength > 0 then
     begin
     if (StrToDateTime ( cxDateEnd.Text) > Date) or (StrToDateTime ( cxDateEnd.Text) < StrToDate('01.01.2005')) or (StrToDateTime ( cxDateEnd.Text) < StrToDateTime ( cxDBDateEdit1.Text)) then
        Begin
         ShowMessage('Дата завершения входного контроля должна находиться в пределах от ' +cxDBDateEdit1.Text + ' до ' +DateTimeToStr(Now));
         cxDateEnd.Date:= Now;
        end;
      end;
end;


procedure TfrmRemountEditor.cxGrid1DBTableView1CellDblClick(
  Sender: TcxCustomGridTableView; ACellViewInfo: TcxGridTableDataCellViewInfo;
  AButton: TMouseButton; AShift: TShiftState; var AHandled: Boolean);
begin
  btnPassportClick(nil);
end;

procedure TfrmRemountEditor.cxGrid1DBTableView1CustomDrawCell(
  Sender: TcxCustomGridTableView; ACanvas: TcxCanvas;
  AViewInfo: TcxGridTableDataCellViewInfo; var ADone: Boolean);
begin
  if AViewInfo.GridRecord.Values[1] = 1 then
   begin
    ACanvas.Brush.Color := clGreen;
    ACanvas.Font.Color := clYellow;
   end;
end;

procedure TfrmRemountEditor.cxGrid2DBTableView1PriznakGetPropertiesForEdit(
  Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
  var AProperties: TcxCustomEditProperties);
begin
  if Sender.FocusedCellViewInfo.Value = '1' then
   AProperties := TcxLabelProperties.Create(Sender)
  else
   begin
    AProperties := TcxMRUEditProperties.Create(Sender);
    TcxMRUEditProperties(AProperties).LookupItems.Add(AC_Kabel);
    TcxMRUEditProperties(AProperties).LookupItems.Add(AC_Term);
    TcxMRUEditProperties(AProperties).DropDownListStyle := lsFixedList;
    TcxMRUEditProperties(AProperties).ShowEllipsis := False;
    TcxMRUEditProperties(AProperties).ImmediatePost := True;
    TcxMRUEditProperties(AProperties).OnChange := cxGrid2DBTableView1PriznakPropertiesChange;
    AProperties.ReadOnly := False;
   end;
end;

procedure TfrmRemountEditor.cxGrid2DBTableView1PriznakPropertiesChange(
  Sender: TObject);
 var ASavePosit: Integer;
begin
  if TcxMRUEdit(Sender).Text = AC_Term then
   with CDSCabCabel do
    begin
     ASavePosit := RecNo;
     DisableControls;
     First;
     while not Eof do
      begin
       if RecNo <> ASavePosit then
        if FieldByName('UDLINIT').AsInteger = 2 then
         begin
          Edit;
          FieldByName('UDLINIT').AsInteger := 0;
          FieldByName('UDLINIT_SET').Value := '0';
          Post;
         end;
       Next;
      end;
     RecNo := ASavePosit;
     EnableControls;
    end;
end;

function TfrmRemountEditor.DoPost: boolean;
begin
  if cdsMasterNODE.AsInteger = idCabLine then
   begin
    cdsMaster.DataRequest(TempCabLineID);
    cdsNPAfter.DataRequest(TempNPAfterID);
   end;
  result := IsDoubleNum and CheckValidRem and inherited DoPost;
end;

procedure TfrmRemountEditor.SetSpecifyLookUpTag(DataSet: TDataSet);
begin
  SetNodeSpecifyLookUpTag(DataSet, FNT);
end;

procedure TfrmRemountEditor.FormCreate(Sender: TObject);
begin
  inherited;

  if Screen.Height = 600 then
  begin
    Position := poDefault;
    Top := 0;
  end;

  plDfInfo.Align := alClient;
  plAsmInfo.Align := alClient;
  plTestnfo.Align := alClient;

  frmDataMod.imlGlobal.GetBitmap(33, dbceMaster.Glyph);
  frmDataMod.imlGlobal.GetBitmap(33, dbceDefector.Glyph);
  frmDataMod.imlGlobal.GetBitmap(33, dbceTester.Glyph);
  frmDataMod.imlGlobal.GetBitmap(33, dbceAsm.Glyph);

  Application.CreateForm(TfrmNodeProfile, FNodeProfile);
  FNodeProfile.NPSrc.DataSet := cdsMaster;
  MergeForm(plNP, FNodeProfile, alClient, true);

  pcCommon.ActivePageIndex := 0;
  Application.CreateForm(TfrmParamList, FParamList);
  FParamList.DoAfterFillDelPrm := AfterFillDelPrm;
  FParamList.DoAfterEditPrm := AfterEditPrm;
  idb := false;
end;

procedure TfrmRemountEditor.SetRMEditorMode(const Value: boolean);
begin
  FRMEditorMode := Value;
end;

function TfrmRemountEditor.GetInputDataRequest(
  DataSet: TDataSet): OleVariant;
begin
  result := inherited GetInputDataRequest(DataSet);
  if TClientDataSet(DataSet) = cdsNPAfter then
    Result := cdsMasterNODEINFOAFTER.AsInteger;
end;

procedure TfrmRemountEditor.FormDestroy(Sender: TObject);
begin
  inherited;
  FNodeProfile.free;
  FParamList.Free;
  if debug<>nil then debug.Free;
end;

procedure RefreshValy;
 var
  CollParams: TParamCollection;
  parametr: TParamItem;
begin
  with frmRemountEditor do
   begin
    CollParams := TParamCollection.create;
    parametr := CollParams.add;
    parametr.ParamName := 'P_REM_ID';
    parametr.ParamType := 'int';
    parametr.OraTypeName := '';
    parametr.AddValue(cdsMasterRMJOURNALID.AsString);
    cdsVals.Data := frmDataMod.sConnect.AppServer.SetPrmStoreProcGetResult
     (CollParams.Count,'PCKG_VALY.GET_VALY_FOR_REMOUNT',CollParams.GetVarArray);
    if cdsVals.RecordCount > 0 then
     with cdsVals do
      while not Eof do
       begin
        if cdsValsdate_out.AsString = '' then
         begin
          btnAddVal.Enabled := False;
          btnRemoveVal.Enabled := True;
          Break;
         end
        else
         begin
          btnAddVal.Enabled := True;
          btnRemoveVal.Enabled := False;
         end;
        Next;
       end
    else
     begin
      btnAddVal.Enabled := True;
      btnRemoveVal.Enabled := False;
     end;
   end;
end;

procedure TfrmRemountEditor.ActivePageChange(Sender: TObject);
 function GetModId(table: string): Integer;
 var cdsTmp: TClientDataSet;
  begin
    cdsTmp := TClientDataSet.Create(nil);
    cdsTmp.Data := SelectSQL('SELECT a.modifikatsia_id FROM '+table+' a WHERE a.type_id = '+cdsMasterTYPEC.AsString);
    if cdsTmp.RecordCount > 0 then
     Result := cdsTmp.Fields.Fields[0].AsInteger
    else
     result := 0;
    cdsTmp.Free;
  end;

var mPrmList: TWinControl;
  mPrmgrp: integer;
  mMasterID: integer;
  mDop: integer;
  stp: boolean;
  i: integer;
begin
  cdsDetales.ApplyUpdates(-1);
  mPrmList := nil;
  dbmDfRem.Visible := false;
  splDFRem.Visible := false;
  pnlVals.Visible := False;
  PrintRDump.Visible := false;
  btSpis.Enabled := false;
  if RMEditorMode then
   InCtrlStat.Visible := False;
  if (cdsMasterSTEP.AsInteger = 403) or (cdsMasterSTEP.AsInteger = 404) then
   stp := False
  else
   stp := True;
  if cdsMasterNODE.AsInteger = idCabLine then
   PanelCabCabel.Visible := true
  else
   PanelCabCabel.Visible := false;
  with TPageControl(Sender) do
    if ActivePage = tsCommon then begin
      ActivePageChange(pcCommon);
    end else
      if ActivePage = tsNPBefore then
      begin
        plNodeProfile.Parent := ActivePage;
        FNodeProfile.NPSrc.DataSet := cdsMaster;
        dblkcbbGRUPPA.DataSource := MSrc;
        mPrmList := plPrmList;
        mPrmgrp := FNT;
        mMasterID := cdsMasterNODEINFOBEFORE.AsInteger;
        for i := 0 to FNodeProfile.ControlCount - 1 do
         FNodeProfile.Controls[i].Enabled := False;
        dblkcbbGRUPPA.Enabled := False;
        if not RMEditorMode then
         InCtrlStat.Visible := False;
        FLengthl.Text := FloatToStr(FLengthl_Before);
      end else
        if ActivePage = tsNPAfter then
        begin
          plNodeProfile.Parent := ActivePage;
          FNodeProfile.NPSrc.DataSet := cdsNPAfter;
          dblkcbbGRUPPA.DataSource := MSrcAfter;
          mPrmList := plPrmList;
          mPrmgrp := FNT;
          mMasterID := cdsMasterNODEINFOAFTER.AsInteger;
          if not RMEditorMode then
           begin
            InCtrlStat.Visible := True;
            if cdsMasterSTEP.OldValue = idStepEnd then
             InCtrlStat.Enabled := False;
           end
          else
           if FNT = idCabLine then
            begin
             InCtrlStat.Visible := True;
             if cdsMasterSTEP.OldValue = idStepEnd then
              InCtrlStat.Enabled := False;
            end;
          case cdsMasterNODE.AsInteger of
           idCabelLine:
            begin
             for i := 0 to FNodeProfile.ControlCount - 1 do
              FNodeProfile.Controls[i].Enabled := False;
             dblkcbbGRUPPA.Enabled := False;
             if (cdsMasterREMOUNTTYPE.AsInteger = idInCtrl) and (cdsMasterSTEP.OldValue <> idStepEnd) then
              FNodeProfile.DBEdit3.Enabled := True;
            end;
           idCabLine:
           begin
            for i := 0 to FNodeProfile.ControlCount - 1 do
             FNodeProfile.Controls[i].Enabled := False;
            if cdsMasterSTEP.OldValue <> idStepEnd then
             dblkcbbGRUPPA.Enabled := True
            else
             dblkcbbGRUPPA.Enabled := False;
           end;
          else
           begin
            for i := 0 to FNodeProfile.ControlCount - 1 do
             FNodeProfile.Controls[i].Enabled := True;
            dblkcbbGRUPPA.Enabled := True;
           end;
          end;
          FLengthl.Text := FloatToStr(FLengthl_After);
       end else
          if ActivePage = tsDF then
          begin
            plRemParams.Parent := ActivePage;
            DBStepDateEdit.DataField := 'DATEDEFECT';
            cdsDetales.Filter := 'Keygrp=1';
            EditDetales.Tag := 1;
            DeleteDetales.Tag := 1;
            CreateDetales.Tag := 1;
            Splitter.Parent := ActivePage;
            Splitter.Left := 500;
            PanelCabCabel.Parent := ActivePage;
            plDetales.Parent := ActivePage;
            cdsDetalesSORT1.DisplayLabel := 'Годн.';
            cdsDetalesSORT2.DisplayLabel := 'Рест.';
            cdsDetalesSORT3.DisplayLabel := 'Спис.';
            mPrmList := plRemPrmList;
            mMasterID := cdsMasterID.AsInteger;
            case cdsMasterNODE.AsInteger of
             7: mDop := GetModId('su_types_modifikatsii_tb');
             50: mDop := GetModId('other_modif_ecn_tb');
             60: mDop := GetModId('other_modif_ped_tb');
            else
             mDop := 0;
            end;
            if mDop = 0 then
             mPrmgrp := FNT + ofsRDF
            else
             mPrmgrp := (FNT + ofsRDF) * 100 + mDop;
            dbmDfRem.Visible := True;
            splDFRem.Visible := true;
            PrintRDump.Visible := true;
            btAdd.Enabled := True and stp;
            btDiv.Enabled := True and stp;
            btSpis.Enabled := True and stp;
           if (cdsMasterREMOUNTTYPE.AsInteger = idrmDismantle) and
              (cdsMasterSTEP.AsInteger = idStepEnd) and
              (CDSCabCabel.RecordCount > 0) then
            begin
             btDelete.Enabled := True;
            end
           else
            btDelete.Enabled := True and stp;
          end else
            if ActivePage = tsAsm then
            begin
              plRemParams.Parent := ActivePage;
              DBStepDateEdit.DataField := 'DATEASSEMBLE';
              cdsDetales.Filter := 'Keygrp=2';
              EditDetales.Tag := 2;
              DeleteDetales.Tag := 2;
              CreateDetales.Tag := 2;
              Splitter.Parent := ActivePage;
              Splitter.Left := 500;
              PanelCabCabel.Parent := ActivePage;
              plDetales.Parent := ActivePage;
              cdsDetalesSORT1.DisplayLabel := 'Нов.';
              cdsDetalesSORT2.DisplayLabel := 'Б/У.';
              cdsDetalesSORT3.DisplayLabel := 'Рест.';
              mPrmList := plRemPrmList;
              if RMEditorMode then
               mPrmgrp := FNT + ofsRAsm
              else
               mPrmgrp := FNT + ofsIAsm;
{
              if cdsMasterNODE.AsInteger in [idECN, idGS] then
               begin
                RefreshValy;
                pnlVals.Visible := True;
                if Run_app_name = apCR then
                 if cdsMasterSTEP.AsInteger = idStepAsm then
                  begin
                   btnAddVal.Visible := True;
                   btnRemoveVal.Visible := True;
                  end;
               end;
}
              mMasterID := cdsMasterID.AsInteger;
              PrintRDump.Visible := true;
              btAdd.Enabled := True and stp;
              btDiv.Enabled := False and stp;
              btSpis.Enabled := False and stp;
              btDelete.Enabled := False and stp;
            end else
              if ActivePage = tsTest then
              begin
                plRemParams.Parent := ActivePage;
                DBStepDateEdit.DataField := 'DATETEST';
                plRemParams.Width := plRemParams.Constraints.MaxWidth;
                PanelCabCabel.Parent := ActivePage;
                mPrmList := plRemPrmList;
                if RMEditorMode then mPrmgrp := FNT + ofsRTest else mPrmgrp := FNT + ofsITest;
                mMasterID := cdsMasterID.AsInteger;
                PrintRDump.Visible := true;
                btAdd.Enabled := False and stp;
                btDiv.Enabled := False and stp;
                btSpis.Enabled := False and stp;
                btDelete.Enabled := False and stp;
              end;
  plDfInfo.Visible := PageActive(tsDF);
  plAsmInfo.Visible := PageActive(tsAsm);
  plTestnfo.Visible := PageActive(tsTest);
  plInfoRem.Parent := plRemParams.Parent;
  PreviewRDump.Visible:=PrintRDump.Visible;
  if not dbmDfRem.Visible then plInfoRem.Height := plFoot.Height else
    plInfoRem.Height := plFoot.Height + 60;
  splDFRem.Top := 0;
  gDetales.Columns[5].Visible := PageActive(tsDF);
  Update;
  if Assigned(mPrmList) then
   FParamList.AssignParamList(mPrmList, mPrmgrp, mMasterID, cdsMaster.ReadOnly);
end;

function TfrmRemountEditor.PageActive(ts: TTabSheet): boolean;
begin
  result := ts.PageControl.ActivePageIndex = ts.PageIndex;
end;

procedure TfrmRemountEditor.cdsDetalesCalcFields(DataSet: TDataSet);
begin
  cdsDetalesSum.AsInteger := cdsDetalesSORT1.AsInteger +
    cdsDetalesSORT2.AsInteger +
    cdsDetalesSORT3.AsInteger;
  if cdsDetalesSum.AsInteger > 0 then
    cdsDetalesBrack.AsFloat := Round(cdsDetalesSORT3.AsInteger /
      cdsDetalesSum.AsInteger * 1000) / 10;
end;

procedure TfrmRemountEditor.FormShow(Sender: TObject);
var i: integer;
begin
  if RMEditorMode then Capt := 'Ремонт узла' else
    Capt := 'Входной контроль';
  inherited;
  pcRemount.OnChange(pcRemount);
  BeforeDoStep;
  FNodeProfile.dbchbUdl.Visible := False;
   //-- Показать поле "Удлинитель"
  if FNT = idCabLine then
   begin
    cdsNPAfterGRUPPA.Required := True;
    FLengthl := TEdit.Create(FNodeProfile);
    with FLengthl do
     begin
      Parent := FNodeProfile;
      Width := FNodeProfile.DBEdit3.Width;
      Top := FNodeProfile.DBEdit3.Top;
      Left := FNodeProfile.DBEdit3.Left;
      FLengthl_Before := GetLengthlCablineBeforeRemount(cdsMasterNODESID.AsInteger, cdsMasterRMJOURNALID.AsInteger);
      Text := FloatToStr(FLengthl_Before);
      Enabled := False;
     end;
    FNodeProfile.DBEdit3.Visible := False;
    if RMEditorMode then
     begin
      InCtrlStat.Properties.Items[0].Caption := 'Готов к монтажу';
      InCtrlStat.Properties.Items[1].Destroy;
      InCtrlStat.Properties.Items[1].Destroy;
     end;
   end
  else
   InCtrlStat.Properties.Items[3].Destroy;
  if FNT <> idCabelLine then
   FNodeProfile.dbchbUdl.Visible := False
  else
   dblcbRemount.Enabled := False;
  //---- Закрываем на редактирование раздел общие сведения в "Журнале ремонта" и "Журнале входного контроля"
{
  if (StP = 'NodeEditorSel') then
   begin
    for i := 0 to FNodeProfile.ControlCount - 1 do
     FNodeProfile.Controls[i].Enabled := False;
    dblkcbbDEPARTMENT_S.Enabled := False;
    for i := 0 to plNP.ControlCount - 1 do
     if plNP.Controls[i].Tag <> 1 then
       plNP.Controls[i].Enabled := False;

    //FParamList.gParams.Enabled := False;      Параметры оставляем открытыми
    //FParamList.plBtn.Visible := False;        Параметры оставляем открытыми
   end;
}
  //----
  cdsMasterREMOUNTTYPE_S.LookupDataSet.Locate('ID', 301, [loPartialKey, loCaseInsensitive]);
  cdsMasterREMOUNTTYPE_S.LookupDataSet.Delete;

  if (cdsMasterREMOUNTTYPE.AsInteger = idInCtrl) and (cdsMasterNODE.AsInteger = idCabelLine) then
   InCtrlStat.Properties.Items.Items[0].Value := idcsReadyForCabl;
{
  else
   if cdsMasterREMOUNTTYPE.AsInteger = idInCtrl then
    InCtrlStat.Properties.Items.Items[2].Destroy;
}
  if cdsMasterNODE.AsInteger = idCabLine then
   begin
    cdsMasterREMOUNTTYPE_S.LookupDataSet.Locate('ID', 315, [loPartialKey, loCaseInsensitive]);
    cdsMasterREMOUNTTYPE_S.LookupDataSet.Delete;
    if cdsMasterSTEP.AsInteger = 404 then
     try
      CDSCabCabel.Data := SelectSQL('SELECT * FROM MICARORB.vw_kable_nodes_rm kn'+
                                    ' WHERE kn.rmjournal_id = '+cdsMasterRMJOURNALID.AsString+
                                    '   AND kn.is_before_remount = 0');
     except
      CDSCabCabel.Close;
     end;
    TempCabLineID := cdsMasterID.AsInteger;
    TempNPAfterID := cdsMasterNODEINFOAFTER.AsInteger;
    if cdsMasterSTEP.OldValue = idStepEnd then
     OkBtn.Enabled := False;
   end;
  if (cdsMasterNODE.AsInteger = idCabelLine) and CabCut then
   begin
    plCableDIV.Visible := CabCut;
    cdsMaster.Edit;
    if Copy(CabLabel.Caption, 0, 1) = 'С' then
     begin
      cdsMasterREMOUNTTYPE.AsInteger := idrmTrash;
      PlInvnum.Visible := False;
      PlSostCab.Visible := False;
      PlSpis.Visible := True;
      CBCurstatus.ItemIndex := CBCurstatus.Items.IndexOfObject(TObject(idcsInTrash));
      Activecontrol := ERealLength;
     end
    else
     begin
      cdsMasterREMOUNTTYPE.AsInteger := idrmCur;
      PlInvnum.Visible := True;
      PlSostCab.Visible := True;
      PlSpis.Visible := False;
      if SostCabLin then
       begin
        cbSostCabLin.Enabled := True;
        cbSostCabLin.Checked := True;
       end
      else
       CBCurstatus.Enabled := False;
      Activecontrol := ENum;
     end;
    cdsMasterDATEEND.AsDateTime := cdsMasterDATEIN.AsDateTime + time;
    cdsNPAfter.Edit;
    frmRemountEditor.Height := 500;
   end;
   if RemoveCabFromCablin then
    begin
     cdsMaster.Edit;
     cdsMasterREMOUNTTYPE.AsInteger := idrmCur;
     cdsMasterDATEEND.AsDateTime := cdsMasterDATEIN.AsDateTime;
     cdsNPAfter.Edit;
     frmRemountEditor.Height := 500;
    end;
  RemIDDel := cdsMasterRMJOURNALID.AsInteger;
end;

procedure TfrmRemountEditor.cdsMasterREMOUNTTYPEChange(Sender: TField);
begin
  if ProcUpdate then exit;
  if (CDSCabCabel.RecordCount <> 0) and (Sender.AsInteger = idrmDismantle) then
   begin
    ShowMessage('Состав кабельной линии должен быть пустой');
    ProcUpdate := True;
    Sender.Value := Sender.OldValue;
    ProcUpdate := False;
   end;
end;

procedure TfrmRemountEditor.cdsMasterREMOUNTTYPEValidate(Sender: TField);
begin
  {нельзя стереть}
  if Sender.IsNull then Abort;
  {нельзя выбрать вх.контроль в ремонте}
  if RMEditorMode and (Sender.AsInteger = idInCtrl) then Abort;
end;

procedure TfrmRemountEditor.cdsMasterSTEPValidate(Sender: TField);
begin
//  {нельзя изменить  Step у исторической записи если не Админ}
  if FHistoricaly and not (TaccessLevel(ArmCurrentInfo.UserLevel) in ulAdmLevel) then Abort;
  {нельзя стереть}
  if Sender.IsNull then Abort;
  {нельзя выбрать сборку не для кабельной линии на вх. контроле}
  if not RMEditorMode and (FNT <> 6) and (Sender.AsInteger = idStepAsm) then Abort;
  {нельзя выбрать дефектовку на вх. контроле}
  if not RMEditorMode and (Sender.AsInteger = idStepDF) then Abort;
end;

procedure TfrmRemountEditor.cdsMasterCalcFields(DataSet: TDataSet);
begin
  if cdsMasterOILWELLSID.AsInteger = 0 then
   Exit;
  if FOWID <> cdsMasterOILWELLSID.AsInteger then
   begin
    FOWID := cdsMasterOILWELLSID.AsInteger;
    FTempOWInfo := CustomSQL(OWLookUp_SQL(FOWID));
   end;
  if not VarIsNull(FTempOWInfo) and VarIsArray(FTempOWInfo) then
   begin
    if not VarIsNull(FTempOWInfo[0]) then
      cdsMasterDEPOSIT_S.Value := FTempOWInfo[0];
    if not VarIsNull(FTempOWInfo[1]) then
      cdsMasterNGDU_S.Value := FTempOWInfo[1];
    if not VarIsNull(FTempOWInfo[2]) then
      cdsMasterKUST.Value := FTempOWInfo[2];
    if not VarIsNull(FTempOWInfo[3]) then
      cdsMasterWell.Value := FTempOWInfo[3];
   end;
end;

procedure TfrmRemountEditor.cdsMasterDATEINChange(Sender: TField);
 const
  z_find_rem_by_date = 'select '+#39+'Начало '+#39+'||datein||'+#39+' завершение '+#39+'||dateend as info' +
                       '  from rmjournal where nodesid = %d and id != %d' +
                       '   and %s between datein and dateend';
 var
  dt: string;
begin
  dt := DateTimeToStr(Sender.AsDateTime);
  dt := 'to_date('+#39+dt+#39+','+#39+'dd.mm.yyyy HH24:MI:SS'+#39+')';
  with TClientDataSet.Create(Self) do
   begin
    Data := SelectSQL(Format(z_find_rem_by_date,[cdsMasterNODESID.AsInteger, cdsMasterID.AsInteger, dt]));
    if RecordCount > 0 then
     begin
      ShowMessage('Указанная дата пересекается с ремонтом'+#10+#13+FieldByName('INFO').AsString);
      Abort;
     end;
    Free;
   end;
end;

procedure TfrmRemountEditor.ERealLengthChange(Sender: TObject);
begin
  if ERealLength.Text = '' then
   begin
     ERealLength.Text := '0';
     ERealLength.SelectAll;
   end;
  if StrToFloat(ERealLength.Text) > CurrentRealLen then
   if CurrentRealLen <= 0 then
     ERealLength.Text := '0'
   else
     ERealLength.Text := FloatToStr(CurrentRealLen);
  cdsNPAfter.Edit;
  cdsNPAfterREAL_LENGTHL.AsFloat := CurrentRealLen - StrToFloat(ERealLength.Text);
end;

procedure TfrmRemountEditor.SelectOWExecute(Sender: TObject);
begin
  Application.CreateForm(TfrmOWBrowser, frmOWBrowser);
  DoAfterSelectSomeID(frmOWBrowser.SelectOW(cdsMasterOILWELLSID.AsInteger), cdsMasterOILWELLSID);
end;

procedure TfrmRemountEditor.CreateDetalesExecute(Sender: TObject);
begin
  if (cdsDetales.RecordCount <> 0) and (MicarMsg('', MMGCreateDetales, MB_ICONQUESTION + MB_YESNO) = mrYes) then
   begin
    DobeforeFiscal;
    FiscalInfo.Descript := 'Удаление деталей';
    Fiscaler(FiscalInfo);
    ExecuteSQL(Delete_Detales_SQL(cdsMasterID.AsInteger, TComponent(Sender).Tag));
    OpenDS(cdsDetales);

    DobeforeFiscal;
    FiscalInfo.Descript := 'Создание деталей';
    Fiscaler(FiscalInfo);
    ExecuteSQL(Create_Detales_SQL(FNT, cdsMasterID.AsInteger, TComponent(Sender).Tag));
    OpenDS(cdsDetales);
   end;
  if (cdsDetales.RecordCount = 0) then
   begin
    DobeforeFiscal;
    FiscalInfo.Descript := 'Создание деталей';
    Fiscaler(FiscalInfo);
    ExecuteSQL(Create_Detales_SQL(FNT, cdsMasterID.AsInteger, TComponent(Sender).Tag));
    OpenDS(cdsDetales);
   end;
end;

procedure TfrmRemountEditor.DeleteDetalesExecute(Sender: TObject);
begin
  if (cdsDetales.RecordCount <> 0) and (MicarMsg('', DelDetales, MB_ICONQUESTION + MB_YESNO) = mrYes) then
   begin
    DobeforeFiscal;
    FiscalInfo.Descript := 'Удаление деталей';
    Fiscaler(FiscalInfo);
    ExecuteSQL(Delete_Detales_SQL(cdsMasterID.AsInteger, TComponent(Sender).Tag));
    OpenDS(cdsDetales);
   end;
end;

procedure TfrmRemountEditor.InCtrlStatPropertiesChange(Sender: TObject);
begin
  InCtrlStat.Style.Color := clBtnFace;
end;

function TfrmRemountEditor.IsDoubleNum: Boolean;
 var AllowDouble: Boolean;
  function CheckNum(Field: TField): Boolean;
   begin
    result := true;
    if (not FNPBeforeNumModified and (Field = cdsMasterNUM)) or
      (not FNPAfterNumModified and (Field = cdsNPAfterNUM)) then
     exit;
    result := ExecuteStoreFunc('CHECK_DOUBLE_NODE',[FNT,cdsMasterNODESID.AsInteger,Field.AsString,cdsNPAfterINVNUM.AsString, cdsMasterTYPEC.AsInteger, cdsMasterPRODUCER.AsInteger]) = 0;
    if not result then
     begin
      if AllowDouble then
       begin
        if MicarMsg('', Format(DoubleNumAllow, [cdsMasternode_s.AsString, Field.AsString]), MB_ICONWARNING + MB_YESNO) = mrYes then
          Result := true;
       end
      else
       MicarMsg('', Format(DoubleNumWrn, [cdsMasternode_s.AsString, Field.AsString]), MB_ICONWARNING);
     end;
    if result and (Field = cdsMasterNUM) then
     FNPBeforeNumModified := false;
    if result and (Field = cdsNPAfterNUM) then
     FNPAfterNumModified := false;
   end;
begin
  AllowDouble := ReadMSIV('ALLOWDOUBLENUM', 'F');
  result := CheckNum(cdsMasterNUM) and CheckNum(cdsNPAfterNUM);
end;

procedure TfrmRemountEditor.cdsNPAfterNUMValidate(Sender: TField);
begin
  if Sender = cdsMasterNUM then
   FNPBeforeNumModified := true
  else
   FNPAfterNumModified := true;
end;

procedure TfrmRemountEditor.date_update;
 var
  CollParams: TParamCollection;
  parametr: TParamItem;
begin
  CollParams := TParamCollection.create;
  parametr := CollParams.add;
  parametr.ParamName := 'P_REM_ID';
  parametr.ParamType := 'int';
  parametr.OraTypeName := '';
  parametr.AddValue(cdsMasterRMJOURNALID.AsString);
  parametr := CollParams.add;
  parametr.ParamName := 'P_VAL_ID';
  parametr.ParamType := 'int';
  parametr.OraTypeName := '';
  parametr.AddValue(cdsValsID.AsString);
  parametr := CollParams.add;
  parametr.ParamName := 'P_DATE_ADD';
  parametr.ParamType := 'dat';
  parametr.OraTypeName := '';
  parametr.AddValue(cdsValsdate_in.AsString);
  parametr := CollParams.add;
  parametr.ParamName := 'P_DATE_REMOVE';
  parametr.ParamType := 'dat';
  parametr.OraTypeName := '';
  parametr.AddValue(cdsValsdate_out.AsString);
  parametr := CollParams.add;
  parametr.ParamName := 'P_DATE_ADD_OLD';
  parametr.ParamType := 'dat';
  parametr.OraTypeName := '';
  parametr.AddValue(VarToStr(cdsValsdate_in.OldValue));
  frmDataMod.sConnect.AppServer.SetPrmStoreProcGetResult
   (CollParams.Count,'PCKG_VALY.UPDATE_DATE',CollParams.GetVarArray);
  RefreshValy;
end;

procedure TfrmRemountEditor.cdsValsAfterOpen(DataSet: TDataSet);
begin
  if DataSet.RecordCount = 0 then
   btnPassport.Enabled := False
  else
   btnPassport.Enabled := True;
end;

procedure TfrmRemountEditor.cdsValsdate_inChange(Sender: TField);
begin
  if br then Exit;
  if cdsMasterDATEASSEMBLE.IsNull then
   begin
    if (Sender.AsDateTime >= cdsMasterDATEIN.AsDateTime) AND (Sender.AsDateTime <= date) then
     date_update
    else
     begin
      Application.MessageBox('Операция не может произойти до начала ремонта или больше текущей даты',
                           'Предупреждение', MB_OK + MB_ICONWARNING);
      br := True;
      Sender.Value := Sender.OldValue;
      br := False;
     end;
   end
  else
   begin
    if (Sender.AsDateTime >= cdsMasterDATEIN.AsDateTime) AND (Sender.AsDateTime <= cdsMasterDATEASSEMBLE.AsDateTime) then
     date_update
    else
     begin
      Application.MessageBox('Операция не может произойти до начала ремонта или после сборки',
                             'Предупреждение', MB_OK + MB_ICONWARNING);
      br := True;
      Sender.Value := Sender.OldValue;
      br := False;
     end;
   end;
end;

procedure TfrmRemountEditor.DoBeforeFiscal;
begin
  inherited;
  with FiscalInfo do
   begin
    if RMEditorMode then
     Block := idBlockRem
    else
     Block := idBlockCtrl;
    if FDoStep > 0 then
     Action := idFActStep;
   end;
end;

procedure TfrmRemountEditor.PrintRDumpExecute(Sender: TObject);
 var MD: TMicarRemDump;
 def_fr_xls, ASheet : Variant;
 //def_fr_xls : TExcelApplication;
 I, I1, I2, K : Integer;
 AFilePathName, AFileName: string;
 //ASheet: _WorkSheet;
begin

AFileName := 'Шаблоны\Defect_Friz.xlt';
AFilePathName := ExtractFilePath(Application.ExeName)+AFileName;
  if not FileExists(AFilePathName) then begin
  Application.MessageBox(PWideChar('Шаблон: "'+AFilePathName+'" не найден') , 'Предупреждение', MB_OK + MB_ICONWARNING);
  Exit;
  end;

def_fr_xls := CreateOleObject('Excel.Application');
//def_fr_xls := TExcelApplication.Create(Application);
def_fr_xls.Workbooks.Open(AFilePathName);

def_fr_xls.Range['J2'] := cdsMaster['DATEDEFECT'];//Дата
def_fr_xls.Range['B4'] := cdsMaster['NODE_S'];//Вид узла
def_fr_xls.Range['B5'] := DBText3.Field.AsString;//Номер узла
def_fr_xls.Range['B6'] := DBText1.Field.AsString;//Тип узла
def_fr_xls.Range['B7'] := DBText2.Field.AsString;//Модель узла
def_fr_xls.Range['B8'] := cdsMaster['PRODUCER_S'];//Завод-изготовитель
//Признак новизны узла
  if cdsMaster['NEW'] = 'н' then
    def_fr_xls.Range['B9'] := 'нов.'
      else
        def_fr_xls.Range['B9'] := 'рем.';
//Вид секции
  if cdsMaster['SECT'] = 'в' then
    def_fr_xls.Range['B10'] := 'верхняя'
      else
        if cdsMaster['SECT'] = 'с' then
          def_fr_xls.Range['B10'] := 'средняя'
            else
              if cdsMaster['SECT'] = 'н' then
                def_fr_xls.Range['B10'] := 'нижняя'
                  else
                    if cdsMaster['SECT'] = 'о' then
                      def_fr_xls.Range['B10'] := 'односекц.'
                        else
                          def_fr_xls.Range['B10'] := '';
//def_fr_xls.Range['B11'] := cdsMaster['OWNER_S'];//Владелец
def_fr_xls.Range['B12'] := 'ТОО "ККМ"';//Заказчик
def_fr_xls.Range['B13'] := DBText5.Field.AsString;//Месторождение
def_fr_xls.Range['B14'] := DBText6.Field.AsString;//Куст
def_fr_xls.Range['B15'] := DBText7.Field.AsString;//Скважина
def_fr_xls.Range['B16'] := '';//Дата остановки
def_fr_xls.Range['B17'] := DBEdit1.Text;//Отработано суток

K:=0;
//Application.MessageBox('Text', 'Caption', MB_YESNOCANCEL);
if FParamList.gParams.DataSource.DataSet.RecordCount>0 then begin
FParamList.gParams.DataSource.DataSet.First;
for I1 := 0 to FParamList.gParams.DataSource.DataSet.RecordCount - 1 do begin
  if FParamList.gParams.DataSource.DataSet.FieldByName('PVALUE').AsString <> '' then begin
  def_fr_xls.ActiveSheet.Rows[K+21+1].Insert(Shift :=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove);
  def_fr_xls.Cells[K+21,1] := FParamList.gParams.DataSource.DataSet.FieldByName('PNAME').AsString;
  def_fr_xls.Cells[K+21,2] := FParamList.gParams.DataSource.DataSet.FieldByName('PVALUE').AsString;
  K:=K+1;
  end;
FParamList.gParams.DataSource.DataSet.Next;
end;
def_fr_xls.ActiveSheet.Rows[K+21].delete;
K:=K-1;
end;

I:=0;
if gDetales.DataSource.DataSet.RecordCount>0 then begin
gDetales.DataSource.DataSet.First;
for I := 0 to gDetales.DataSource.DataSet.RecordCount - 1 do begin
def_fr_xls.ActiveSheet.Rows[I+K+26+1].Insert(Shift :=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove);
def_fr_xls.Cells[I+K+26,1] := gDetales.DataSource.DataSet.Fields[2].AsString;
def_fr_xls.Cells[I+K+26,3] := gDetales.DataSource.DataSet.Fields[1].AsString;
def_fr_xls.Cells[I+K+26,4] := gDetales.DataSource.DataSet.Fields[3].AsString;
def_fr_xls.Cells[I+K+26,5] := gDetales.DataSource.DataSet.Fields[4].AsString;
def_fr_xls.Cells[I+K+26,6] := gDetales.DataSource.DataSet.Fields[7].AsString;
def_fr_xls.Cells[I+K+26,7] := gDetales.DataSource.DataSet.Fields[6].AsString;
def_fr_xls.Cells[I+K+26,8] := gDetales.DataSource.DataSet.Fields[5].AsString;
gDetales.DataSource.DataSet.Next;
end;
def_fr_xls.ActiveSheet.Rows[I+K+26].delete;
K:=K-1;
end;

def_fr_xls.Cells[I+K+29,4] := dbceMaster.Text;//Мастер (ФИО из дефектовочной ведомости)
def_fr_xls.Cells[I+K+31,4] := dbceDefector.Text;//Дефектовщик (ФИО из дефектовочной ведомости)
def_fr_xls.Cells[I+K+33,4] := DBStepDateEdit.Field.AsString;//Дата дефектации (из дефектовочной ведомости)

def_fr_xls.Visible := True;

{  if pcRemount.ActivePage = tsDF then
   if TComponent(Sender).tag > NullItem then
    ExportDF.ExecTest(cdsMaster, FParamList.cdsParams, cdsDetales)
   else
    begin
     MD := TMicarAsmDump.Create;
     with MD do
      try
       RMMode:=FRMEditorMode;
       PrintDump(cdsMasterID.AsInteger,False);
      finally
       Free;
      end;
    end
  else
   if pcRemount.ActivePage = tsAsm then
    begin
     MD := TMicarAsmDump.Create;
     with MD do
      begin
       try
        RMMode:=FRMEditorMode;
        PrintDump(cdsMasterID.AsInteger,TComponent(Sender).tag>NullItem);
       finally
        Free;
       end;
      end;
    end
   else
    if pcRemount.ActivePage = tsTest then
     begin
      MD := TMicarTestDump.Create;
      with MD do
       begin
        try
         RMMode:=FRMEditorMode;
         PrintDump(cdsMasterID.AsInteger,TComponent(Sender).tag>NullItem);
        finally
         Free;
        end;
       end;
     end;      }
end;

procedure TfrmRemountEditor.ShowNPasspExecute(Sender: TObject);
begin
  if cdsMasterNODE.AsInteger = idCabLine then
   begin
    Application.CreateForm(TfrmCabLin1, frmCabLin1);
    frmCabLin1.ShowCabLine(cdsMasterNODESID.AsInteger);
   end
  else
   begin
    Application.CreateForm(TfrmNodeEditor, frmNodeEditor);
    frmNodeEditor.ShowNode(FNT, cdsMasterNODESID.asInteger);
   end;
end;

procedure TfrmRemountEditor.ShowEpuPasspExecute(Sender: TObject);
 var AID:integer;
begin
  AID:=ExecuteStoreFunc('FIND_EPU_BY_REM',[cdsMasterNodesID.AsInteger,
                                           cdsMasterOILWELLSID.AsInteger,
                                           cdsMasterDateIN.AsString]);
  if AID>0 then
   begin
    Application.CreateForm(TfrmEPUEDITOR, frmEPUEDITOR);
    frmEPUEDITOR.ShowEPU(AID);
   end
  else
   MicarMsg('',EPUNOTFOUND,MB_ICONWARNING);
end;

procedure TfrmRemountEditor.MSrcDataChange(Sender: TObject; Field: TField);
begin
  ShowEpuPassp.Visible:= not cdsMasterOILWELLSID.IsNull;
end;

procedure TfrmRemountEditor.EditDetalesExecute(Sender: TObject);
begin
  inherited;
  Name_Node_Select:='Детали узлов';
  Application.CreateForm(TfrmSettings, frmSettings);
  frmSettings.ShowPage(sPageClass);
  if (cdsDetales.RecordCount <> 0) then
   begin
    DobeforeFiscal;
    FiscalInfo.Descript := 'Удаление деталей';
    Fiscaler(FiscalInfo);
    ExecuteSQL(Delete_Detales_SQL(cdsMasterID.AsInteger, TComponent(Sender).Tag));
    OpenDS(cdsDetales);

   end;
  if (cdsDetales.RecordCount = 0) then
   begin
    DobeforeFiscal;
    FiscalInfo.Descript := 'Создание деталей';
    Fiscaler(FiscalInfo);
    ExecuteSQL(Create_Detales_SQL(FNT, cdsMasterID.AsInteger, TComponent(Sender).Tag));
    OpenDS(cdsDetales);
   end;
  ActivePageChange(pcRemount);
end;

procedure TfrmRemountEditor.FKeyPress(Sender: TObject; var Key: Char);
 var k : char;
begin
  k:=key;
  keytab(k);
  key:=k;
end;

procedure TfrmRemountEditor.FKeyPressNumber(Sender: TObject; var Key: Char);
begin
  if (key <> '0') and
     (key <> '1') and
     (key <> '2') and
     (key <> '3') and
     (key <> '4') and
     (key <> '5') and
     (key <> '6') and
     (key <> '7') and
     (key <> '8') and
     (key <> '9') and
     (key <> ',') and
     (ord(key) <> 127) and
     (ord(key) <> 8) then key := chr(0);
end;

procedure TfrmRemountEditor.FormClose(Sender: TObject;
  var Action: TCloseAction);
 var ExeSQL, model777, typec777, producer777: String;
     inCurStat: integer;
begin
 Action := caFree;
 if idb then
  DebugPoint1;
 if (cdsMasterNODE.AsInteger = idCabLine) and
    (ModalResult = mrOk) and
    (Action <> caNone)
 then
  CDSCabCabel.ApplyUpdates(0);
 if Run_app_name <> apManager then
  begin
   cdsMaster.Edit;
   cdsMasterGRUPPA.AsInteger := cdsNPAfterGRUPPA.AsInteger;
   if CabCut then
     if Copy(CabLabel.Caption, 0, 1) = 'С' then
      begin
        cdsCabSpis.Edit;
        cdsCabSpis.Fields.FieldByName('NODESID').AsInteger := cdsMasterNODESID.AsInteger;
        cdsCabSpis.Fields.FieldByName('RMJOURNALID').AsInteger := cdsMasterRMJOURNALID.AsInteger;
        cdsCabSpis.Fields.FieldByName('RAPORTTYPE').AsInteger := 1302;
        cdsCabSpis.Fields.FieldByName('TYPEC').AsInteger := cdsMasterTYPEC.AsInteger;
        cdsCabSpis.Fields.FieldByName('PRODUCER').AsInteger := cdsMasterPRODUCER.AsInteger;
        cdsCabSpis.Fields.FieldByName('ACCOUNTER').AsInteger := cdsMasterACCOUNTER.AsInteger;
        cdsCabSpis.Fields.FieldByName('DEPARTMENT').AsInteger := cdsMasterDEPARTMENT.AsInteger;
        cdsCabSpis.Fields.FieldByName('DATERAPORT').AsDateTime := Date;
        cdsCabSpis.Fields.FieldByName('LENGTHL').AsFloat := StrToFloat(ERealLength.Text);
      end;
  end;
  inherited;
 if CabCut and (ModalResult = mrOk) and (Action <> caNone) then
  begin
   if Copy(CabLabel.Caption, 0, 1) = 'Н' then
    begin
     if DBEModel.Text = '' then
      model777 := '0'
     else
      model777 := DBEModel.Text;
     if DBETypec.Text = '' then
      typec777 := '0'
     else
     typec777 := DBETypec.Text;
     if DBEProducer.Text = '' then
      producer777 := '0'
     else
      producer777 := DBEProducer.Text;
     if CBCurstatus.Visible then
      inCurStat := INTEGER(CBCurstatus.Items.Objects[CBCurstatus.ItemIndex])
     else
      inCurStat := idcsInCabLine;
     ExeSQL := 'begin INSERT INTO nodeprofile (ID, node, typec, model, lengthl, producer, num, invnum, NEW, real_lengthl, cena, gruppa, udlinit) VALUES (nodeprofile_id.NEXTVAL, 6, '+
            typec777+', '+model777+', '+chr(39)+'0'+chr(39)+', '+producer777+', '+chr(39)+ENum.Text+chr(39)+', '+
            chr(39)+EInvnum.Text+chr(39)+', '+chr(39)+cxRGNewRem.Properties.Items.Items[cxRGNewRem.ItemIndex].Value+chr(39)+', '+chr(39)+ERealLength.Text+chr(39)+', '+
            chr(39)+chr(39)+', '+FloatToStr(cdsMasterGRUPPA.Value)+', '+IntToStr(cxRadioGroup1.ItemIndex)+
            '); INSERT INTO nodes (ID, nodeinfo, owner, dateprod, datein, curstatus, department) VALUES (nodes_id.NEXTVAL, nodeprofile_id.CURRVAL, '+
            IntToStr(NPAfterOwner2.Tag)+', to_date ('+chr(39)+DBEDateend.Text+chr(39)+','+chr(39)+'dd.mm.yyyy hh24:mi:ss'+chr(39)+'), to_date ('+chr(39)+
            DBEDatein.Text+chr(39)+','+chr(39)+'dd.mm.yyyy hh24:mi:ss'+chr(39)+'), '+IntToStr(inCurStat)+', '+CBDepartment.Text+
            '); INSERT INTO cab_history_tb VALUES ('+
            Nodeid.Text+', '+Nodeinfoid.Text+', nodes_id.CURRVAL, nodeprofile_id.CURRVAL,'+eRemID.Text+'); ';
            if {not} cbSostCabLin.Checked then  // Проверка статуса узла
              if not CBCurstatus.Visible then
               ExeSQL := ExeSQL+' INSERT INTO kablenodes(id, agregatid, nodesid, nodeinfo)'+
               'VALUES(KABLE_ID.NEXTVAL, '+IntToStr(IDCabLinForDIV)+', nodes_id.CURRVAL, nodeprofile_id.CURRVAL); end;'
              else
               ExeSQL := ExeSQL+' end;'
            else
              ExeSQL := ExeSQL+' end;';
     if ExecuteSQL(ExeSQL) = 0 then
      begin
       Application.MessageBox('Произошла ошибка в БД'+#10+#13+'Обратитесь к администратору системы', 'Предупреждение', MB_OK + MB_ICONSTOP);
       ExeSQL := 'UPDATE NODEPROFILE SET REAL_LENGTHL = '+FloatToStr(CurrentRealLen)+' WHERE ID = (select nodeinfo from nodes where id = '+cdsMasterNODESID.AsString+')';
       ExecuteSQL(ExeSQL);
       CommitTransAction;
{
       ERealLength.Text := '0';
       cdsNPAfter.Edit;
       cdsNPAfterREAL_LENGTHL.Value := CurrentRealLen;
       cdsNPAfter.Post;
       cdsNPAfter.ApplyUpdates(0);
       Action := caNone;
}
      end;
    end
   else
    begin
     if cdsNPAfterREAL_LENGTHL.AsInteger <= 0 then
      begin
       ExeSQL := 'begin '+
                 ' delete kablenodes '+
                 '  where agregatid = '+IntToStr(IDCabLinForDIV)+
                 '    and nodesid = '+cdsMasterNODESID.AsString+'; '+
                 ' update nodes set curstatus = '+IntToStr(idcsInTrash)+
                 '  where id = '+cdsMasterNODESID.AsString+';'+
                 ' commit; '+
                 'end;';
       ExecuteSQL(ExeSQL);
      end;
    end;
   if ModalResult <> mrOk  then
    begin
     ExeSQL := 'DELETE FROM RMJOURNAL WHERE ID ='+IntToStr(RemIDDel);
     ExecuteSQL(ExeSQL);
     exit;
    end;
  end;
end;

procedure TfrmRemountEditor.OkBtnClick(Sender: TObject);

  function IsDoubleNum: Boolean;
  var AllowDouble: Boolean;
  begin
    result := true;
    AllowDouble := ReadMSIV('ALLOWDOUBLENUM', 'F');
    result:=ExecuteStoreFunc('CHECK_DOUBLE_NODE',[FNT,FID,frmRemountEditor.ENum.Text,frmRemountEditor.EInvNum.Text, cdsMasterTYPEC.AsInteger, cdsMasterPRODUCER.AsInteger])=0;
    if not result then
      if AllowDouble then
      begin
        if MicarMsg('', Format(DoubleNumAllow, ['Кабель', ENum.Text]), MB_ICONWARNING + MB_YESNO) = mrYes then
          result := true;
      end
      else MicarMsg('', Format(DoubleNumWrn, ['Кабель', ENum.Text]), MB_ICONWARNING);
  end;

 const MsgLength = 'Длина кабеля не может быть меньше %d м.';
begin
  if (cdsMasterSTEP.AsInteger = idStepEnd) and
    (cdsMasterNODE.AsInteger = idCabLine) then
   begin
    cdsNPAfter.Edit;
    cdsNPAfterSROSTKOV.AsFloat := StrToInt(LabCabSrost.Caption);
    cdsNPAfterREAL_LENGTHL.AsFloat := StrToFloat(LabCabLength.Caption);
   end;
  if CabCut then
   if (cdsMasterSTEP.AsInteger = idStepEnd) and
     (cdsMasterNODE.AsInteger = idCabelLine) then
    if Copy(CabLabel.Caption, 0, 1) = 'С' then
     begin
      if (cdsNPAfterREAL_LENGTHL.AsInteger < MaxCutCabLength) then
       if (cdsNPAfterREAL_LENGTHL.AsInteger <> 0) then
        begin
         ShowMessage(Format(MsgLength, [MaxCutCabLength]));
         ModalResult := mrNone;
         Exit;
        end
     end
    else
     begin
      if (cdsNPAfterREAL_LENGTHL.AsInteger < MaxCutCabLength) or
       (StrToFloat(ERealLength.Text) < MaxCutCabLength) then
       begin
        ShowMessage(Format(MsgLength, [MaxCutCabLength]));
        ModalResult := mrNone;
        Exit;
       end;
      if not IsDoubleNum then
       begin
        ModalResult := mrNone;
        Exit;
       end;
     end;
  cdsMaster.Edit;
  cdsMasterGRUPPA.AsInteger := cdsNPAfterGRUPPA.AsInteger;
  inherited;
end;

procedure TfrmRemountEditor.btDefectDblClick(Sender: TObject);
var
    rid: integer;
    sqlstr: string;
    ds: TClientDataSet;
    xls : variant;
    rm: TReportManager;
    prs: TStringList;
begin
//  inherited;
  ds := TClientDataSet.Create(Application);
  ds.RemoteServer := frmDataMod.sConnect;
  ds.ProviderName := 'EPUEditor';
  sqlstr := 'select Agregat.TYPEKPO from Agregat, Nodes where Nodes.ID = ' +
  IntToStr(cdsMasterNODESID.AsInteger)+ ' and Agregat.ID=Nodes.Tempid ';
  ds.DATA := SelectSql(sqlstr);
  case cdsMasterNode.AsInteger of
    // ЭЦН
    1:  if ds.Fields[0].AsInteger = 704 then
         rid := 19  // УЦПК
        else
         if ds.Fields[0].AsInteger = 703 then
          rid := 20  // УЭЦВ
         else
          rid := 25; // УЭЦН
    // Газосепаратор
    2: rid := 22;
    // Секция ПЭД
    3: rid := 21;
    // Протектор
    4: rid := 24;
    // Компенсатор
    5: rid := 23;
  end;
  try
    try
      xls := GetActiveOleObject('Excel.Application');
    except
      xls := CreateOleObject('Excel.Application');
    end;
  except
    raise Exception.Create('Не могу запустить Excel');
    Exit;
  end;
  prs := TStringList.Create;
  prs.Add(cdsMasterID.AsString);
  rm := TReportManager.Create(frmDataMod.sConnect, xls);
  rm.MakeReport(rid, prs);
  prs.Free;
  xls.Visible := true;
  rm.Free;
end;

procedure TfrmRemountEditor.BtDeleteClick(Sender: TObject);
begin
  inherited;
 if MessageDlg('Эту операцию нельзя будет отменить!', mtConfirmation, [mbOk, mbCancel], 0) = mrCancel then Exit;
 Randomize;
 try
  cdsCabCabel.Delete;
 except
   Case Random(10) of
    0: ShowMessage('Не, я так не умею!');
    1: ShowMessage('Не представляю себе это возможным!');
    2: ShowMessage('Я конечно постараюсь, но не обещаю!');
    3: ShowMessage('Удалить для меня не проблема!');
    4: ShowMessage('Вот только было бы что удалить!');
    5: ShowMessage('Раньше я знал что от меня требуют!');
    6: ShowMessage('Я, за это не отвечаю!');
    7: ShowMessage('Горе пользователи, опять просят невыполнимое!');
    8: ShowMessage('Да, да! Удалил уже все, что было возможно!');
    9: ShowMessage('Группа разработчиков АПК "МИЦАР", благодарит Вас, за тестирование кнопки "УДАЛИТЬ"!');
  end;
 end;
 cdsCabRef;
 CDSCabCabel.ApplyUpdates(-1);
end;


procedure TFrmRemountEditor.cdsMasterReconcileError(DataSet: TCustomClientDataSet;
  E: EReconcileError; UpdateKind: TUpdateKind; var Action: TReconcileAction);
begin
  if Copy(E.Message,0,3) = 'ORA' then
   Application.MessageBox(PWideChar(E.Message), 'Ошибка', MB_OK + MB_ICONSTOP)
  else
   Application.MessageBox(PWideChar(E.Message), 'Предупреждение', MB_OK + MB_ICONWARNING);
end;

procedure TfrmRemountEditor.BtDivClick(Sender: TObject);
 var RemID: Variant;
begin
  if CDSCabCabel.RecordCount = 0 then
   begin
    ShowMessage(MsgCabOperation);
    Exit;
   end;
  CabCut := True;
  SostCabLin := True;
  IDCabLinForDIV := CDSCabCabelAGREGATID.AsInteger;
  RemID:=ExecuteStoreFUNC('SEND_NODE_REMOUNT',[CDSCabCabel.Fields.FieldByName('ID').AsInteger, WORD(CabCut), WORD(CabCut), ArmCurrentInfo.DepartmentID]);
  Application.CreateForm(TfrmRemountEditor, frmRemountEditor);
  if frmRemountEditor.ShowRemount(6, RemID, True) = 0 then
   ExecuteSQL(Format(del_rem,[VarToStr(RemID), VarToStr(RemID)]))
  else
   ExecuteSQL(Format(upd_kablenodes_info, [GetNPInfoRemEnd(RemID), CDSCabCabelAGREGATID.AsInteger, CDSCabCabelID.AsInteger]));
  CabCut := False;
  StartWait;
  CDSCabCabel.Close;
  CDSCabCabel.Open;
  StopWait;
  SostCabLin := False;
end;

procedure TfrmRemountEditor.btnAddValClick(Sender: TObject);
begin
  Application.CreateForm(TfrmListValy, frmListValy);
  frmListValy.ShowValy(fvSelect, cdsMasterRMJOURNALID.AsInteger);
  RefreshValy;
end;

procedure TfrmRemountEditor.btnPassportClick(Sender: TObject);
begin
  Application.CreateForm(TfrmValy, frmValy);
  frmValy.ValyEditFormState(fveBrowse, cdsValsID.AsInteger);
end;

procedure TfrmRemountEditor.btnRemoveValClick(Sender: TObject);
 var
  CollParams: TParamCollection;
  parametr: TParamItem;
begin
  with cdsVals do
   begin
    First;
    while not Eof do
     begin
      if cdsValsis_current_val.AsInteger = 1 then Break;
      Next;
     end;
   end;
  CollParams := TParamCollection.create;
  parametr := CollParams.add;
  parametr.ParamName := 'P_REM_ID';
  parametr.ParamType := 'int';
  parametr.OraTypeName := '';
  parametr.AddValue(cdsValsrmjournal_id.AsString);
  parametr := CollParams.add;
  parametr.ParamName := 'P_VAL_ID';
  parametr.ParamType := 'int';
  parametr.OraTypeName := '';
  parametr.AddValue(cdsValsID.AsString);
  parametr := CollParams.add;
  parametr.ParamName := 'P_DATE_ADD';
  parametr.ParamType := 'dat';
  parametr.OraTypeName := '';
  parametr.AddValue(cdsValsdate_in.AsString);
  frmDataMod.sConnect.AppServer.SetPrmStoreProcGetResult
   (CollParams.Count,'PCKG_VALY.REMOVE_VAL_FROM_REMONT',CollParams.GetVarArray);
  RefreshValy;
end;

procedure TfrmRemountEditor.BtSpisClick(Sender: TObject);
 var RemID: Variant;
begin
  if CDSCabCabel.RecordCount = 0 then
   begin
    ShowMessage(MsgCabOperation);
    Exit;
   end;
  CabCut := True;
  IDCabLinForDIV := CDSCabCabelAGREGATID.AsInteger;
  RemID:=ExecuteStoreFUNC('SEND_NODE_REMOUNT',[CDSCabCabel.Fields.FieldByName('ID').AsInteger, WORD(CabCut), WORD(CabCut), ArmCurrentInfo.DepartmentID]);
  Application.CreateForm(TfrmRemountEditor, frmRemountEditor);
  frmRemountEditor.CabLabel.Caption := 'Списать кабель';
  if frmRemountEditor.ShowRemount(6, RemID, True) = 0 then
   ExecuteSQL(Format(del_rem,[VarToStr(RemID), VarToStr(RemID)]))
  else
   ExecuteSQL(Format(upd_kablenodes_info, [GetNPInfoRemEnd(RemID), CDSCabCabelAGREGATID.AsInteger, CDSCabCabelID.AsInteger, CDSCabCabelNODEINFO.AsInteger]));
  CabCut := False;
  StartWait;
  CDSCabCabel.Close;
  CDSCabCabel.Open;
  StopWait;
end;

procedure TfrmRemountEditor.CabRemoveClick(Sender: TObject);
const z_del_cab_from_cabl = 'BEGIN' +
                            '  DELETE FROM KABLENODES WHERE NODESID = %d' +
                            '     AND AGREGATID = %d;' +
                            '  UPDATE nodeprofile' +
                            '     SET new = '+#39+'р'+#39+
                            '  WHERE id = %d; %s' +
                            'END;';
 var AStatus: Integer;
     RemID: Variant;
     prz: string;
begin
  if CDSCabCabel.RecordCount = 0 then
   begin
    ShowMessage(MsgCabOperation);
    Exit;
   end;
  RemoveCabFromCablin := True;
  Application.CreateForm(TfrmSelStatNode, frmSelStatNode);
  AStatus := frmSelStatNode.SelectStatusForCab;
  if AStatus > 0 then
   begin
    RemID:=ExecuteStoreFUNC('SEND_NODE_REMOUNT',[CDSCabCabel.Fields.FieldByName('ID').AsInteger, WORD(RemoveCabFromCablin), WORD(RemoveCabFromCablin), ArmCurrentInfo.DepartmentID]);
    Application.CreateForm(TfrmRemountEditor, frmRemountEditor);
    if frmRemountEditor.ShowRemount(6, RemID, True) = 0 then
     ExecuteSQL(Format(del_rem,[VarToStr(RemID), VarToStr(RemID)]))
    else
     begin
      if CDSCabCabelUDLINIT.AsInteger = 2 then
       prz := 'UPDATE nodeprofile SET udlinit = 0 WHERE id = ' + CDSCabCabelNODEINFO.AsString +';'
      else
       prz := '';
      ExecuteSQL(Format(z_del_cab_from_cabl, [CDSCabCabelID.AsInteger, CDSCabCabelAGREGATID.AsInteger, CDSCabCabelNODEINFO.AsInteger, prz]));
      ExecuteSQL(UpdateNodeStatus(AStatus, CDSCabCabelID.AsInteger));
      StartWait;
      CDSCabCabel.Close;
      CDSCabCabel.Open;
      StopWait;
     end;
   end;
  RemoveCabFromCablin := False;
//  if MessageDlg('Эту операцию нельзя будет отменить!', mtConfirmation, [mbOk, mbCancel], 0) = mrCancel then Exit;
end;

procedure TfrmRemountEditor.cbSostCabLinClick(Sender: TObject);
begin
  if cbSostCabLin.Checked then
   begin
    CBCurstatus.ItemIndex := CBCurstatus.Items.IndexOfObject(TObject(idcsInCabLine));
    CBCurstatus.Visible := False;
    lblCurStat.Visible := False;
   end
  else
   begin
    CBCurstatus.ItemIndex := CBCurstatus.Items.IndexOfObject(TObject(idcsReadyForCabl));
    CBCurstatus.Visible := True;
    lblCurStat.Visible := True;
   end;
end;

procedure TfrmRemountEditor.CDSCabCabelAfterPost(DataSet: TDataSet);
begin
 cdsCabRef;
end;

procedure TfrmRemountEditor.CDSCabCabelBeforeDelete(DataSet: TDataSet);
begin
  CDSCabCabel.Edit;
  CDSCabCabelCURSTATUS.Value := idcsWRem;
end;

procedure TfrmRemountEditor.CDSCabCabelBeforeOpen(DataSet: TDataSet);
begin
  CDSCabCabel.DataRequest(cdsMasterNODESID.Value);
end;

procedure TfrmRemountEditor.CDSCabCabelUDLINIT_SETGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
  if (Sender.Value = '0') or VarIsNull(Sender.Value) then
   Text := AC_Kabel
  else
   if Sender.Value = '1' then
    Text := AC_Udl
   else
    if Sender.Value = '2' then
     Text := AC_Term;
end;

procedure TfrmRemountEditor.CDSCabCabelUDLINIT_SETSetText(Sender: TField;
  const Text: string);
begin
  if Text = AC_Kabel then
   begin
    Sender.Value := '0';
    CDSCabCabelUDLINIT.Value := 0;
   end
  else
   if Text = AC_Term then
    begin
     Sender.Value := '2';
     CDSCabCabelUDLINIT.Value := 2;
    end;
end;

procedure TfrmRemountEditor.sbPrintOTKClick(Sender: TObject);
begin
  ExportNZtoExcel(cdsMasterNODE.AsInteger, cdsMasterNODESID.AsInteger, cdsMasterID.AsInteger, dblcbRemount.Text);
end;

procedure TfrmRemountEditor.DebugPoint1;
var
	i: integer;
begin
  if debug=nil then debug := TStringList.Create;
  for i:=0 to cdsMaster.Fields.Count-1 do
    debug.Add(cdsMaster.Fields[i].Name+'='+cdsMaster.Fields[i].AsString);
  debug.SaveToFile(getcurrentdir+'\debug.txt');
end;

procedure TfrmRemountEditor.DebTriggerExecute(Sender: TObject);
begin
  idb := not idb;
  if idb then
    Caption := Caption+' (#Отладка)'
  else
    Caption := Copy(Caption,1,Pos('#',Caption)-3);
  Update;
end;

function TfrmRemountEditor.GetLengthlCablineBeforeRemount(CablineID, rmID: Integer): double;
 var ALen: Variant;
begin
  ALen := CustomSQL(Format(sqlGetLengthCablineBeforeRemount, [CablineID, rmID]))[0];
  if VarIsNull(ALen) then
   Result := 0
  else
   Result := ALen;
end;

end.

