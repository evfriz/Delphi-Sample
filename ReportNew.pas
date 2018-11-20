unit ReportNew;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBClient, cxGridLevel, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxClasses, cxControls, cxGridCustomView, cxGrid,
  cxContainer, cxTreeView, dxBar, dxStatusBar, cxSplitter, ExtCtrls,
  ComCtrls, dxtree,MDATA, dxdbtree, VDFFilterSL, cxGraphics, cxStyles,
  cxCustomData, cxFilter, cxData, cxDataStorage, cxEdit, cxDBData,
  cxDropDownEdit;

type
  TReportTree = class(TForm)
    TPanelClient: TPanel;
    TPanelBotton: TPanel;
    cxSplitter1: TcxSplitter;
    dxStatusBar1: TdxStatusBar;
    TPanelTree: TPanel;
    dxBarManager1: TdxBarManager;
    cxSplitter2: TcxSplitter;
    TPanelGrid: TPanel;
    cxGrid: TcxGrid;
    cxGridDBTable: TcxGridDBTableView;
    cxGrid1DBTableViewCriteria: TcxGridDBColumn;
    cxGrid1DBTableViewOperaton: TcxGridDBColumn;
    cxGrid1DBTableViewValues: TcxGridDBColumn;
    cxGridLevel: TcxGridLevel;
    CriteriaValues: TClientDataSet;
    CriteriaValuesSTATE_ID: TFloatField;
    CriteriaValuesView_Criteria: TStringField;
    CriteriaValuesView_operation: TStringField;
    CriteriaValuesView_Values: TStringField;
    DSCriteriaVal: TDataSource;
    DSReportTree: TDataSource;
    CDSReportTree: TClientDataSet;
    DBTreeView1: TdxDBTreeView;
    CDSReportTreeID: TFloatField;
    CDSReportTreeNAME: TStringField;
    CDSReportTreeTR_TR_ID: TFloatField;
    procedure CriteriaValuesView_CriteriaGetText(Sender: TField;
      var Text: String; DisplayText: Boolean);
    procedure CriteriaValuesView_CriteriaSetText(Sender: TField;
      const Text: String);
    procedure CriteriaValuesView_operationGetText(Sender: TField;
      var Text: String; DisplayText: Boolean);
    procedure CriteriaValuesView_operationSetText(Sender: TField;
      const Text: String);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ReportTree: TReportTree;
  emID               : array of integer;
  flt                : TFilterCollection;
  ldr                : TmyLoader;
  SQLstr             : string;
  Loads              : Boolean;
  Change_Filter      : Boolean = false;
  View_param_List    : TObject;
  View_criteria_List : TObject;
  tm                 : Tdatetime;
implementation

{$R *.DFM}

procedure TReportTree.CriteriaValuesView_CriteriaGetText(Sender: TField;
  var Text: String; DisplayText: Boolean);
var
i, a: integer;
begin
if Sender.AsString = '' then exit;
for i :=0 to Flt.Filters.Count-1 do
for a :=0 to TFilter(Flt.Filters.Objects[i]).CriteriaList.Count-1 do
    if StrToInt(Sender.AsString) = INTEGER (TCriteriaFilter(TFilter(Flt.Filters.Objects[i]).CriteriaList.Objects[a]).Criteria_ID) then
         Text := TFilter(Flt.Filters.Objects[i]).CriteriaList.Strings[a];
end;

procedure TReportTree.CriteriaValuesView_CriteriaSetText(Sender: TField;
  const Text: String);
var
i, a: integer;
begin
for i :=0 to Flt.Filters.Count-1 do
for a :=0 to TFilter(Flt.Filters.Objects[i]).CriteriaList.Count-1 do
    if Text = TFilter(Flt.Filters.Objects[i]).CriteriaList.Strings[a] then
       Sender.AsString := IntToStr(INTEGER (TCriteriaFilter(TFilter(Flt.Filters.Objects[i]).CriteriaList.Objects[a]).Criteria_ID));
end;

procedure TReportTree.CriteriaValuesView_operationGetText(Sender: TField;
  var Text: String; DisplayText: Boolean);
var
i, a: integer;
begin
if Sender.AsString = '' then exit;
for i :=0 to Flt.Filters.Count-1 do
for a :=0 to TFilter(Flt.Filters.Objects[i]).OperationList.Count-1 do
    if StrToInt(Sender.AsString) = INTEGER (TOperationFilter(TFilter(Flt.Filters.Objects[i]).OperationList.Objects[a]).Operation_ID) then
         Text := TFilter(Flt.Filters.Objects[i]).OperationList.Strings[a];


end;

procedure TReportTree.CriteriaValuesView_operationSetText(Sender: TField;
  const Text: String);
var
i, a: integer;
begin
for i :=0 to Flt.Filters.Count-1 do
for a :=0 to TFilter(Flt.Filters.Objects[i]).OperationList.Count-1 do
    if Text = TFilter(Flt.Filters.Objects[i]).OperationList.Strings[a] then
       Sender.AsString := IntToStr(INTEGER (TOperationFilter(TFilter(Flt.Filters.Objects[i]).OperationList.Objects[a]).Operation_ID));
end;

procedure TReportTree.FormShow(Sender: TObject);
begin
 CDSReportTree.Open;
end;

end.
