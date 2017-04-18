unit uFrmProdutos;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, System.Math, Vcl.Graphics, uClsProdutosDAO,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.StdCtrls,
  Vcl.CheckLst, Vcl.ExtCtrls, Vcl.Grids, Vcl.DBGrids, Data.DB, Datasnap.DBClient,
  Vcl.ToolWin, Vcl.ActnMan, Vcl.ActnCtrls, Vcl.Buttons, Vcl.ImgList,
  Data.DBXMSSQL, Data.FMTBcd, Data.SqlExpr, Datasnap.Provider, Vcl.DBCtrls,
  Vcl.Mask, RxToolEdit, RxDBCtrl;

type
  TFrmProdutos = class(TForm)
    pnlHeader: TPanel;
    Panel3: TPanel;
    cklstBuscarpor: TCheckListBox;
    Label1: TLabel;
    pgctrlBuscarpor: TPageControl;
    tab1: TTabSheet;
    tab2: TTabSheet;
    tab3: TTabSheet;
    tab4: TTabSheet;
    tab5: TTabSheet;
    tab6: TTabSheet;
    PageControl: TPageControl;
    tabGrid: TTabSheet;
    tabFormulario: TTabSheet;
    dbgProdutos: TDBGrid;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    cdsControle: TClientDataSet;
    dsControle: TDataSource;
    dsProdutos: TDataSource;
    cdsProdutos: TClientDataSet;
    dsGrupoProdutos: TDataSource;
    cdsGrupoProdutos: TClientDataSet;
    cdsEmpresas: TClientDataSet;
    cdsProdutosEmpresas: TClientDataSet;
    cdsNCMs: TClientDataSet;
    dsNCMs: TDataSource;
    dsProdutosEmpresas: TDataSource;
    dsEmpresas: TDataSource;
    dsTiposProdutos: TDataSource;
    cdsTipoProduto: TClientDataSet;
    cdsUnidadesMedidas: TClientDataSet;
    dsUnidadesMedidas: TDataSource;
    pnlAcoes: TPanel;
    spdInserir: TSpeedButton;
    spdAlterar: TSpeedButton;
    spdExcluir: TSpeedButton;
    spdCancelar: TSpeedButton;
    spdGravar: TSpeedButton;
    Img16Ativa: TImageList;
    SQLConnection: TSQLConnection;
    DataSetProvider1: TDataSetProvider;
    SQLDataSet1: TSQLDataSet;
    cdsEmpresasID: TIntegerField;
    cdsEmpresasEMP_RAZAOSOCIAL: TStringField;
    cdsEmpresasEMP_NOMEFANTASIA: TStringField;
    cdsEmpresasEMP_TIPOPESSOA: TStringField;
    cdsEmpresasEMP_CNPJCPF: TStringField;
    cdsEmpresasEMP_INSCRESTADUAL: TStringField;
    cdsEmpresasEMP_INSCRMUNICIPAL: TStringField;
    cdsEmpresasEMP_PISCOFINS: TSmallintField;
    cdsEmpresasEMP_ALIQPIS: TFMTBCDField;
    cdsEmpresasEMP_ALIQCOFINS: TFMTBCDField;
    cdsEmpresasEMP_CEP: TStringField;
    cdsEmpresasEMP_LOGRADOURO: TStringField;
    cdsEmpresasEMP_NUMERO: TStringField;
    cdsEmpresasEMP_BAIRRO: TStringField;
    cdsEmpresasEMP_MUNICIPIO: TStringField;
    cdsEmpresasEMP_UF: TStringField;
    cdsEmpresasEMP_PAGINA: TStringField;
    cdsEmpresasEMAIL: TStringField;
    cdsEmpresasAUD_DEL: TSmallintField;
    cdsEmpresasAUD_DATCAD: TDateField;
    cdsEmpresasAUD_USUCAD: TStringField;
    cdsEmpresasAUD_DATULTALT: TDateField;
    cdsEmpresasAUD_USUULTALT: TStringField;
    cdsEmpresasAUD_DATEXC: TDateField;
    cdsEmpresasAUD_USUEXC: TStringField;
    cdsNCMsID: TIntegerField;
    cdsNCMsCODIGOPAI_ID: TIntegerField;
    cdsNCMsNCM_CODIGO: TStringField;
    cdsNCMsNCM_CODIGOPAI: TStringField;
    cdsNCMsNCM_DESCRICAO: TStringField;
    cdsNCMsNCM_TRIBUTADO: TSmallintField;
    cdsNCMsNCM_ALIQUOTA: TFMTBCDField;
    cdsNCMsNCM_NOTA: TMemoField;
    cdsNCMsAUD_DEL: TIntegerField;
    cdsNCMsAUD_DATCAD: TDateField;
    cdsNCMsAUD_USUCAD: TStringField;
    cdsNCMsAUD_DATULTALT: TDateField;
    cdsNCMsAUD_USUULTALT: TStringField;
    cdsNCMsAUD_DATEXC: TDateField;
    cdsNCMsAUD_USUEXC: TStringField;
    cdsGrupoProdutosID: TIntegerField;
    cdsGrupoProdutosGRP_DESCRICAO: TStringField;
    cdsGrupoProdutosAUD_DEL: TSmallintField;
    cdsGrupoProdutosAUD_DATCAD: TDateField;
    cdsGrupoProdutosAUD_USUCAD: TStringField;
    cdsGrupoProdutosAUD_DATULTALT: TDateField;
    cdsGrupoProdutosAUD_USUULTALT: TStringField;
    cdsGrupoProdutosAUD_DATEXC: TDateField;
    cdsGrupoProdutosAUD_USUEXC: TStringField;
    cdsProdutosEmpresasID: TIntegerField;
    cdsProdutosEmpresasEMPRESA_ID: TIntegerField;
    cdsProdutosEmpresasPRODUTO_ID: TIntegerField;
    cdsProdutosEmpresasPRE_PRECOCUSTO: TFMTBCDField;
    cdsProdutosEmpresasPRE_CUSTOMEDIO: TFMTBCDField;
    cdsProdutosEmpresasPRE_CUSTOCONTABIL: TFMTBCDField;
    cdsProdutosEmpresasPRE_CUSTOGERENCIAL: TFMTBCDField;
    cdsProdutosEmpresasPRE_PRECOVENDA: TFMTBCDField;
    cdsProdutosEmpresasPRE_PERCCUSTOOPER: TFMTBCDField;
    cdsProdutosEmpresasPRE_MARGEMLUCRO: TFMTBCDField;
    cdsProdutosEmpresasPRE_DATAULTVENDA: TDateField;
    cdsProdutosEmpresasPRE_VLRULTVENDA: TFMTBCDField;
    cdsProdutosEmpresasPRE_DATAULTCOMPRA: TDateField;
    cdsProdutosEmpresasPRE_VLRULTCOMPRA: TFMTBCDField;
    cdsProdutosEmpresasPRE_SALDO_ANTERIOR: TFMTBCDField;
    cdsProdutosEmpresasPRE_SALDO_ATUAL: TFMTBCDField;
    cdsProdutosEmpresasPRE_SALDO_DISPONIVEL: TFMTBCDField;
    cdsProdutosEmpresasPRE_SALDO_RESERVADO: TFMTBCDField;
    cdsProdutosEmpresasPRE_STATUS: TSmallintField;
    cdsProdutosEmpresasAUD_DEL: TIntegerField;
    cdsProdutosEmpresasAUD_DATCAD: TDateField;
    cdsProdutosEmpresasAUD_USUCAD: TStringField;
    cdsProdutosEmpresasAUD_DATULTALT: TDateField;
    cdsProdutosEmpresasAUD_USUULTALT: TStringField;
    cdsProdutosEmpresasAUD_DATEXC: TDateField;
    cdsProdutosEmpresasAUD_USUEXC: TStringField;
    cdsUnidadesMedidasID: TIntegerField;
    cdsUnidadesMedidasUNM_DESCRICAO: TStringField;
    cdsUnidadesMedidasUNM_SIGLA: TStringField;
    cdsUnidadesMedidasAUD_DEL: TSmallintField;
    cdsUnidadesMedidasAUD_DATCAD: TDateField;
    cdsUnidadesMedidasAUD_USUCAD: TStringField;
    cdsUnidadesMedidasAUD_DATULTALT: TDateField;
    cdsUnidadesMedidasAUD_USUULTALT: TStringField;
    cdsUnidadesMedidasAUD_DATEXC: TDateField;
    cdsUnidadesMedidasAUD_USUEXC: TStringField;
    cdsTipoProdutoID: TIntegerField;
    cdsTipoProdutoDESCRICAO: TStringField;
    cdsControleID: TIntegerField;
    cdsControleGRUPO_PRODUTO_ID: TIntegerField;
    cdsControleUNIDADE_MEDIDA_ID: TIntegerField;
    cdsControleNCM_ID: TIntegerField;
    cdsControlePRD_CODBARRAS: TStringField;
    cdsControlePRD_CODIGO: TStringField;
    cdsControlePRD_APELIDO: TStringField;
    cdsControlePRD_DESCRICAO: TStringField;
    cdsControlePRD_DESCRICAOFISCAL: TStringField;
    cdsControlePRD_DESCRICAOTECNICA: TMemoField;
    cdsControlePRD_TIPOPRODUTO: TSmallintField;
    cdsControlePRD_PESOBRUTO: TFMTBCDField;
    cdsControlePRD_PESOLIQUIDO: TFMTBCDField;
    cdsControlePRD_FOTO: TBlobField;
    cdsControleAUD_DEL: TIntegerField;
    cdsControleAUD_DATCAD: TDateField;
    cdsControleAUD_USUCAD: TStringField;
    cdsControleAUD_DATULTALT: TDateField;
    cdsControleAUD_USUULTALT: TStringField;
    cdsControleAUD_DATEXC: TDateField;
    cdsControleAUD_USUEXC: TStringField;
    tab7: TTabSheet;
    cdsControleGRP_DESCRICAO: TStringField;
    cdsControleUNM_SIGLA: TStringField;
    cdsControleNCM_DESCRICAO: TStringField;
    spdBuscar: TSpeedButton;
    cdsProdutosID: TIntegerField;
    cdsProdutosGRUPO_PRODUTO_ID: TIntegerField;
    cdsProdutosGRP_DESCRICAO: TStringField;
    cdsProdutosUNIDADE_MEDIDA_ID: TIntegerField;
    cdsProdutosUNM_SIGLA: TStringField;
    cdsProdutosNCM_ID: TIntegerField;
    cdsProdutosNCM_DESCRICAO: TStringField;
    cdsProdutosPRD_CODBARRAS: TStringField;
    cdsProdutosPRD_CODIGO: TStringField;
    cdsProdutosPRD_APELIDO: TStringField;
    cdsProdutosPRD_DESCRICAO: TStringField;
    cdsProdutosPRD_DESCRICAOFISCAL: TStringField;
    cdsProdutosPRD_DESCRICAOTECNICA: TMemoField;
    cdsProdutosPRD_TIPOPRODUTO: TSmallintField;
    cdsProdutosPRD_PESOBRUTO: TFMTBCDField;
    cdsProdutosPRD_PESOLIQUIDO: TFMTBCDField;
    cdsProdutosPRD_FOTO: TBlobField;
    cdsProdutosAUD_DEL: TIntegerField;
    cdsProdutosAUD_DATCAD: TDateField;
    cdsProdutosAUD_USUCAD: TStringField;
    cdsProdutosAUD_DATULTALT: TDateField;
    cdsProdutosAUD_USUULTALT: TStringField;
    cdsProdutosAUD_DATEXC: TDateField;
    cdsProdutosAUD_USUEXC: TStringField;
    lblPesqCodigo: TLabel;
    edtPesqCodigo: TEdit;
    lblPesqDescricao: TLabel;
    edtPesqDescricao: TEdit;
    lkcPesqGrupoProduto: TDBLookupComboBox;
    lblPesqGrupoProduto: TLabel;
    cdsPesqGrupoProduto: TClientDataSet;
    dsPesqGrupoProduto: TDataSource;
    cdsPesqGrupoProdutoID: TIntegerField;
    cdsPesqGrupoProdutoGRP_DESCRICAO: TStringField;
    cdsPesqGrupoProdutoAUD_DEL: TSmallintField;
    cdsPesqGrupoProdutoAUD_DATCAD: TDateField;
    cdsPesqGrupoProdutoAUD_USUCAD: TStringField;
    cdsPesqGrupoProdutoAUD_DATULTALT: TDateField;
    cdsPesqGrupoProdutoAUD_USUULTALT: TStringField;
    cdsPesqGrupoProdutoAUD_DATEXC: TDateField;
    cdsPesqGrupoProdutoAUD_USUEXC: TStringField;
    lblPesqCodigoBarra: TLabel;
    edtPesqCodigoBarra: TEdit;
    cdsPesqNCMs: TClientDataSet;
    dsPesqNCMs: TDataSource;
    cdsPesqNCMsID: TIntegerField;
    cdsPesqNCMsCODIGOPAI_ID: TIntegerField;
    cdsPesqNCMsNCM_CODIGO: TStringField;
    cdsPesqNCMsNCM_CODIGOPAI: TStringField;
    cdsPesqNCMsNCM_DESCRICAO: TStringField;
    cdsPesqNCMsNCM_TRIBUTADO: TSmallintField;
    cdsPesqNCMsNCM_ALIQUOTA: TFMTBCDField;
    cdsPesqNCMsNCM_NOTA: TMemoField;
    cdsPesqNCMsAUD_DEL: TIntegerField;
    cdsPesqNCMsAUD_DATCAD: TDateField;
    cdsPesqNCMsAUD_USUCAD: TStringField;
    cdsPesqNCMsAUD_DATULTALT: TDateField;
    cdsPesqNCMsAUD_USUULTALT: TStringField;
    cdsPesqNCMsAUD_DATEXC: TDateField;
    cdsPesqNCMsAUD_USUEXC: TStringField;
    lblPesqNCMs: TLabel;
    lkcPesqNCMs: TDBLookupComboBox;
    Label2: TLabel;
    lkcPesqTipoProduto: TDBLookupComboBox;
    cdsPesqTipoProduto: TClientDataSet;
    dsPesqTipoProduto: TDataSource;
    cdsPesqTipoProdutoID: TIntegerField;
    cdsPesqTipoProdutoDESCRICAO: TStringField;
    cdsPesqUnidadesMedidas: TClientDataSet;
    dsPesqUnidadesMedidas: TDataSource;
    cdsPesqUnidadesMedidasID: TIntegerField;
    cdsPesqUnidadesMedidasUNM_DESCRICAO: TStringField;
    cdsPesqUnidadesMedidasUNM_SIGLA: TStringField;
    cdsPesqUnidadesMedidasAUD_DEL: TSmallintField;
    cdsPesqUnidadesMedidasAUD_DATCAD: TDateField;
    cdsPesqUnidadesMedidasAUD_USUCAD: TStringField;
    cdsPesqUnidadesMedidasAUD_DATULTALT: TDateField;
    cdsPesqUnidadesMedidasAUD_USUULTALT: TStringField;
    cdsPesqUnidadesMedidasAUD_DATEXC: TDateField;
    cdsPesqUnidadesMedidasAUD_USUEXC: TStringField;
    lblPesqUnidadesMedidas: TLabel;
    lkcPesqUnidadesMedidas: TDBLookupComboBox;
    edtCodigo: TDBEdit;
    lblCodigo: TLabel;
    Label3: TLabel;
    edtCodigoBarras: TDBEdit;
    Label4: TLabel;
    edtApelido: TDBEdit;
    lblDescricao: TLabel;
    edtDescricao: TDBEdit;
    lblDescricaoFiscal: TLabel;
    edtDescricaoFiscal: TDBEdit;
    dbrTipoProduto: TDBRadioGroup;
    edtDescricaoTecnica: TDBMemo;
    lblDescricaoTecnica: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    edtPesoBruto: TDBEdit;
    edtPesoLiquido: TDBEdit;
    Label7: TLabel;
    lkcNCMs: TDBLookupComboBox;
    lblNCM: TLabel;
    lblUnidadeMedida: TLabel;
    lkcUnidadeMedida: TDBLookupComboBox;
    lkcGrupoProduto: TDBLookupComboBox;
    Label8: TLabel;
    edtUsuarioCadastro: TDBEdit;
    lblUsuarioCad: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    edtUsuarioUltAlteracao: TDBEdit;
    Label11: TLabel;
    Label12: TLabel;
    edtUsuarioExclusao: TDBEdit;
    Label13: TLabel;
    spdRetornar: TSpeedButton;
    edtDataCadastro: TDBDateEdit;
    edtDataUltAlteracao: TDBDateEdit;
    edtDataExclusao: TDBDateEdit;
    procedure FormCreate(Sender: TObject);
    procedure PageControlChanging(Sender: TObject; var AllowChange: Boolean);
    procedure cklstBuscarporClickCheck(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure spdBuscarClick(Sender: TObject);
    procedure spdInserirClick(Sender: TObject);
    procedure edtCodigoExit(Sender: TObject);
    procedure spdCancelarClick(Sender: TObject);
    procedure spdRetornarClick(Sender: TObject);
    procedure spdGravarClick(Sender: TObject);
    procedure spdAlterarClick(Sender: TObject);
    procedure dbgProdutosDblClick(Sender: TObject);
    procedure spdExcluirClick(Sender: TObject);
    procedure dsControleStateChange(Sender: TObject);
    procedure dsProdutosStateChange(Sender: TObject);
  private
    { Private declarations }
    inserir, alterar, deletar, visualizar : Boolean;
    function PreencheString(Palavra, Caracter, Direcao: string;  Tamanho: integer): String;
    function RemoveCaracteresNumero(Str: String): String;
    function FormatarIntStr(Numero, Tamanho: Integer): string;
    function EAN13_DV(CodEAN13: String): String;
    function RPad(const s: String; Pad: Integer): String;
    function LPad(const s: String; Pad: Integer): String;
    function StrIsNumber(const S: String): Boolean;
    function IfThen(AValue: Boolean; const ATrue: string; AFalse: string = ''): string;overload;
    function IfThen(AValue: Boolean; const ATrue: Integer; const AFalse: Integer): Integer;overload;
    function IfThen(AValue: Boolean; const ATrue: Int64; const AFalse: Int64): Int64;overload;
    function IfThen(AValue: Boolean; const ATrue: Double; const AFalse: Double): Double;overload;
    function CharIsNum(const C: Char): Boolean;
  public
    { Public declarations }
  end;

var
  FrmProdutos: TFrmProdutos;
  produtosDAO : TProdutosDAO;

implementation

{$R *.dfm}

procedure TFrmProdutos.cklstBuscarporClickCheck(Sender: TObject);
var
  items, i, checado: Integer;
  checar: Boolean;

begin
  checado := 0;

  if cklstBuscarpor.ItemIndex = 0 then // Não Filtrar
  begin
      cklstBuscarpor.CheckAll(cbUnchecked, true, false);
      tab1.TabVisible := True;
      tab2.TabVisible := True;
      tab3.TabVisible := True;
      tab4.TabVisible := True;
      tab5.TabVisible := True;
      tab6.TabVisible := True;
      tab7.TabVisible := True;
      edtPesqCodigo.Text              := '';
      edtPesqDescricao.Text           := '';
      lkcPesqGrupoProduto.KeyValue    := 0;
      edtPesqCodigoBarra.Text         := '';
      lkcPesqNCMs.KeyValue            := 0;
      lkcPesqTipoProduto.KeyValue     := 0;
      lkcPesqUnidadesMedidas.KeyValue := 0;
      tab1.TabVisible := False;
      tab2.TabVisible := False;
      tab3.TabVisible := False;
      tab4.TabVisible := False;
      tab5.TabVisible := False;
      tab6.TabVisible := False;
      tab7.TabVisible := False;
  end
  else
  begin
    if cklstBuscarpor.Checked[0] = True then
      cklstBuscarpor.Checked[0] := False;

    items := 1;
    While items <= cklstBuscarpor.items.Count - 1 do
    begin
      if cklstBuscarpor.Checked[items] then
      begin
        checar := True;
        checado := checado + 1;
      end
      else
        checar := False;

      if checar then
      begin
        case items of
          1:begin
                tab1.TabVisible := True;
                edtPesqCodigo.SetFocus;
            end;
          2:begin
                tab2.TabVisible := True;
                edtPesqDescricao.SetFocus;
            end;
          3:begin
                tab3.TabVisible := True;
                lkcPesqGrupoProduto.SetFocus;
            end;
          4:begin
                tab4.TabVisible := True;
                edtPesqCodigoBarra.SetFocus;
            end;
          5:begin
                tab5.TabVisible := True;
                lkcPesqNCMs.SetFocus;
            end;
          6:begin
                tab6.TabVisible := True;
                lkcPesqTipoProduto.SetFocus;
            end;
          7:begin
                tab7.TabVisible := True;
                lkcPesqUnidadesMedidas.SetFocus;
            end;
        end;
        pgctrlBuscarpor.ActivePageIndex := items - 1;
      end
      else
      begin
        case items of
          1:begin
                edtPesqCodigo.Text := '';
                tab1.TabVisible := False;
            end;
          2:begin
                edtPesqDescricao.Text := '';
                tab2.TabVisible := False;
            end;
          3:begin
                lkcPesqGrupoProduto.KeyValue := 0;
                tab3.TabVisible := False;
            end;
          4:begin
                edtPesqCodigoBarra.Text := '';
                tab4.TabVisible := False;
            end;
          5:begin
                lkcPesqNCMs.KeyValue := 0;
                tab5.TabVisible := False;
            end;
          6:begin
                lkcPesqTipoProduto.KeyValue := 0;
                tab6.TabVisible := False;
            end;
          7:begin
                lkcPesqUnidadesMedidas.KeyValue := 0;
                tab7.TabVisible := False;
            end;
        end;
      end;
      items := items + 1
    end;
  end;
  if checado = 0 then
    cklstBuscarpor.Checked[0] := True;
end;

procedure TFrmProdutos.dbgProdutosDblClick(Sender: TObject);
begin
    spdAlterar.Click;
end;

procedure TFrmProdutos.dsControleStateChange(Sender: TObject);
begin
    if cdsControle.State = dsBrowse then
    begin
        spdInserir.Enabled  := True;
        if cdsControle.RecordCount > 0 then
        begin
           spdAlterar.Enabled   := True;
           spdExcluir.Enabled   := True;
        end
        else
        begin
           spdAlterar.Enabled   := False;
           spdExcluir.Enabled   := False;
        end;
        spdCancelar.Enabled     := False;
        spdGravar.Enabled       := False;
    end;

end;

procedure TFrmProdutos.dsProdutosStateChange(Sender: TObject);
begin
    if cdsProdutos.State = dsBrowse then
    begin
        spdInserir.Enabled  := True;
        if cdsProdutos.RecordCount > 0 then
        begin
           spdAlterar.Enabled   := True;
           spdExcluir.Enabled   := True;
        end
        else
        begin
           spdAlterar.Enabled   := False;
           spdExcluir.Enabled   := False;
        end;
        spdCancelar.Enabled     := False;
        spdGravar.Enabled       := False;
    end
    else if cdsProdutos.State in [dsInsert, dsEdit] then
    begin
        spdInserir.Enabled      := False;
        spdAlterar.Enabled      := False;
        spdExcluir.Enabled      := False;
        spdCancelar.Enabled     := True;
        spdGravar.Enabled       := True;
    end;

end;

procedure TFrmProdutos.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    if produtosDAO <> nil then
       produtosDAO.Free;
end;

procedure TFrmProdutos.FormCreate(Sender: TObject);
begin
  PageControl.ActivePage  := TabGrid;
  visualizar  := true;
  deletar     := false;
  inserir     := false;
  alterar     := false;

  cdsTipoProduto.CreateDataSet;
  cdsTipoProduto.Open;
  cdsTipoProduto.Append;
  cdsTipoProdutoID.Value  := 0;
  cdsTipoProdutoDESCRICAO.Value := 'Acabado';
  cdsTipoProduto.Post;
  cdsTipoProduto.Append;
  cdsTipoProdutoID.Value  := 1;
  cdsTipoProdutoDESCRICAO.Value := 'Insumo';
  cdsTipoProduto.Post;
  cdsTipoProduto.Append;
  cdsTipoProdutoID.Value  := 2;
  cdsTipoProdutoDESCRICAO.Value := 'Industrializado';
  cdsTipoProduto.Post;
  cdsTipoProduto.Append;
  cdsTipoProdutoID.Value  := 3;
  cdsTipoProdutoDESCRICAO.Value := 'Kit';
  cdsTipoProduto.Post;
  cdsTipoProduto.Append;
  cdsTipoProdutoID.Value  := 4;
  cdsTipoProdutoDESCRICAO.Value := 'Ferramentas';
  cdsTipoProduto.Post;
  cdsTipoProduto.Append;
  cdsTipoProdutoID.Value  := 5;
  cdsTipoProdutoDESCRICAO.Value := 'Componentes';
  cdsTipoProduto.Post;
  cdsTipoProduto.Append;
  cdsTipoProdutoID.Value  := 6;
  cdsTipoProdutoDESCRICAO.Value := 'Uso/Consumo';
  cdsTipoProduto.Post;

  cdsPesqTipoProduto.CreateDataSet;
  cdsPesqTipoProduto.Open;
  cdsPesqTipoProduto.Append;
  cdsPesqTipoProdutoID.Value  := 0;
  cdsPesqTipoProdutoDESCRICAO.Value := 'Acabado';
  cdsPesqTipoProduto.Post;
  cdsPesqTipoProduto.Append;
  cdsPesqTipoProdutoID.Value  := 1;
  cdsPesqTipoProdutoDESCRICAO.Value := 'Insumo';
  cdsPesqTipoProduto.Post;
  cdsPesqTipoProduto.Append;
  cdsPesqTipoProdutoID.Value  := 2;
  cdsPesqTipoProdutoDESCRICAO.Value := 'Industrializado';
  cdsPesqTipoProduto.Post;
  cdsPesqTipoProduto.Append;
  cdsPesqTipoProdutoID.Value  := 3;
  cdsPesqTipoProdutoDESCRICAO.Value := 'Kit';
  cdsPesqTipoProduto.Post;
  cdsPesqTipoProduto.Append;
  cdsPesqTipoProdutoID.Value  := 4;
  cdsPesqTipoProdutoDESCRICAO.Value := 'Ferramentas';
  cdsPesqTipoProduto.Post;
  cdsPesqTipoProduto.Append;
  cdsPesqTipoProdutoID.Value  := 5;
  cdsPesqTipoProdutoDESCRICAO.Value := 'Componentes';
  cdsPesqTipoProduto.Post;
  cdsPesqTipoProduto.Append;
  cdsPesqTipoProdutoID.Value  := 6;
  cdsPesqTipoProdutoDESCRICAO.Value := 'Uso/Consumo';
  cdsPesqTipoProduto.Post;

end;

procedure TFrmProdutos.FormShow(Sender: TObject);
begin
    produtosDAO := TProdutosDAO.Create;
    cklstBuscarporClickCheck(Sender);

    if not cdsProdutos.Active then
    begin
        cdsProdutos.CreateDataSet;
        cdsProdutos.Close;
        cdsProdutos.Data        := produtosDAO.CarregaProdutos(-1);
        cdsProdutos.Open;
    end
    else
    begin
        if cdsProdutos.RecordCount > 0 then
           cdsProdutos.EmptyDataSet;

        cdsProdutos.Close;
        cdsProdutos.Data        := produtosDAO.CarregaProdutos(-1);
        cdsProdutos.Open;
    end;

    if not cdsPesqGrupoProduto.Active then
    begin
        cdsPesqGrupoProduto.CreateDataSet;
        cdsPesqGrupoProduto.Close;
        cdsPesqGrupoProduto.Data   := produtosDAO.CarregaGruposProdutos;
        cdsPesqGrupoProduto.Open;
    end
    else
    begin
        if cdsPesqGrupoProduto.RecordCount > 0 then
           cdsPesqGrupoProduto.EmptyDataSet;

        cdsPesqGrupoProduto.Close;
        cdsPesqGrupoProduto.Data   := produtosDAO.CarregaGruposProdutos;
        cdsPesqGrupoProduto.Open;
    end;

    if not cdsEmpresas.Active then
    begin
        cdsEmpresas.CreateDataSet;
        cdsEmpresas.Close;
        cdsEmpresas.Data    := produtosDAO.CarregaEmpresas;
        cdsEmpresas.Open;
    end
    else
    begin
        if cdsEmpresas.RecordCount > 0 then
           cdsEmpresas.EmptyDataSet;

        cdsEmpresas.Close;
        cdsEmpresas.Data    := produtosDAO.CarregaEmpresas;
        cdsEmpresas.Open;
    end;

    if not cdsProdutosEmpresas.Active then
    begin
        cdsProdutosEmpresas.CreateDataSet;
        cdsProdutosEmpresas.Close;
        cdsProdutosEmpresas.Data  := produtosDAO.CarregaProdutosEmpresas(-1);
        cdsProdutosEmpresas.Open;
    end
    else
    begin
        if cdsProdutosEmpresas.RecordCount > 0 then
           cdsProdutosEmpresas.EmptyDataSet;

        cdsProdutosEmpresas.Close;
        cdsProdutosEmpresas.Data  := produtosDAO.CarregaProdutosEmpresas(-1);
        cdsProdutosEmpresas.Open;
    end;

    if not cdsPesqNCMs.Active then
    begin
        cdsPesqNCMs.CreateDataSet;
        cdsPesqNCMs.Close;
        cdsPesqNCMs.Data    := produtosDAO.CarregaNcms;
        cdsPesqNCMs.Open;
    end
    else
    begin
        if cdsPesqNCMs.RecordCount > 0 then
           cdsPesqNCMs.EmptyDataSet;

        cdsPesqNCMs.Close;
        cdsPesqNCMs.Data    := produtosDAO.CarregaNcms;
        cdsPesqNCMs.Open;
    end;

    if not cdsPesqUnidadesMedidas.Active then
    begin
        cdsPesqUnidadesMedidas.CreateDataSet;
        cdsPesqUnidadesMedidas.Close;
        cdsPesqUnidadesMedidas.Data   := produtosDAO.CarregaUnidadesMedidas;
        cdsPesqUnidadesMedidas.Open;
    end
    else
    begin
        if cdsPesqUnidadesMedidas.Recordcount > 0 then
           cdsPesqUnidadesMedidas.EmptyDataSet;

        cdsPesqUnidadesMedidas.Close;
        cdsPesqUnidadesMedidas.Data   := produtosDAO.CarregaUnidadesMedidas;
        cdsPesqUnidadesMedidas.Open;
    end;

end;

procedure TFrmProdutos.PageControlChanging(Sender: TObject;
  var AllowChange: Boolean);
begin
    if (not visualizar) and (PageControl.ActivePageIndex = 0) then
       AllowChange := true
    else if (not visualizar) and (PageControl.ActivePageIndex = 1) then
       AllowChange := false;
end;

procedure TFrmProdutos.spdAlterarClick(Sender: TObject);
var
  codigo : Integer;
begin
    visualizar  := False;
    deletar     := false;
    inserir     := false;
    alterar     := true;
    codigo      := cdsControleID.AsInteger;

    PageControl.ActivePage  := tabFormulario;

    if not cdsProdutos.Active then
    begin
        cdsProdutos.CreateDataSet;
        cdsProdutos.Close;
        cdsProdutos.Data        := produtosDAO.CarregaProdutos(codigo);
        cdsProdutos.Open;
    end
    else
    begin
        if cdsProdutos.RecordCount > 0 then
           cdsProdutos.EmptyDataSet;

        cdsProdutos.Close;
        cdsProdutos.Data        := produtosDAO.CarregaProdutos(codigo);
        cdsProdutos.Open;
    end;

    if not cdsGrupoProdutos.Active then
    begin
        cdsGrupoProdutos.CreateDataSet;
        cdsGrupoProdutos.Close;
        cdsGrupoProdutos.Data   := produtosDAO.CarregaGruposProdutos;
        cdsGrupoProdutos.Open;
    end
    else
    begin
        if cdsGrupoProdutos.RecordCount > 0 then
           cdsGrupoProdutos.EmptyDataSet;

        cdsGrupoProdutos.Close;
        cdsGrupoProdutos.Data   := produtosDAO.CarregaGruposProdutos;
        cdsGrupoProdutos.Open;
    end;

    if not cdsProdutosEmpresas.Active then
    begin
        cdsProdutosEmpresas.CreateDataSet;
        cdsProdutosEmpresas.Close;
        cdsProdutosEmpresas.Data  := produtosDAO.CarregaProdutosEmpresas(codigo);
        cdsProdutosEmpresas.Open;
    end
    else
    begin
        if cdsProdutosEmpresas.RecordCount > 0 then
           cdsProdutosEmpresas.EmptyDataSet;

        cdsProdutosEmpresas.Close;
        cdsProdutosEmpresas.Data  := produtosDAO.CarregaProdutosEmpresas(codigo);
        cdsProdutosEmpresas.Open;
    end;

    if not cdsNCMs.Active then
    begin
        cdsNCMs.CreateDataSet;
        cdsNCMs.Close;
        cdsNCMs.Data    := produtosDAO.CarregaNcms;
        cdsNCMs.Open;
    end
    else
    begin
        if cdsNCMs.RecordCount > 0 then
           cdsNCMs.EmptyDataSet;

        cdsNCMs.Close;
        cdsNCMs.Data    := produtosDAO.CarregaNcms;
        cdsNCMs.Open;
    end;

    if not cdsUnidadesMedidas.Active then
    begin
        cdsUnidadesMedidas.CreateDataSet;
        cdsUnidadesMedidas.Close;
        cdsUnidadesMedidas.Data   := produtosDAO.CarregaUnidadesMedidas;
        cdsUnidadesMedidas.Open;
    end
    else
    begin
        if cdsUnidadesMedidas.Recordcount > 0 then
           cdsUnidadesMedidas.EmptyDataSet;

        cdsUnidadesMedidas.Close;
        cdsUnidadesMedidas.Data   := produtosDAO.CarregaUnidadesMedidas;
        cdsUnidadesMedidas.Open;
    end;

    cdsProdutos.Edit;
    cdsProdutosAUD_DATULTALT.AsDateTime      := Date;
    cdsProdutosAUD_USUULTALT.AsString        := 'Administrador';
    edtCodigo.Enabled := False;
    edtCodigoBarras.SetFocus;
end;

procedure TFrmProdutos.spdBuscarClick(Sender: TObject);
var
  i, k: Integer;
  pFiltro, pOrder : string;
begin
    try
      produtosDAO.sCodigo       := '';
      produtosDAO.sDescricao    := '';
      produtosDAO.iGrupoProduto := 0;
      produtosDAO.sCodigoBarra  := '';
      produtosDAO.iNcm          := 0;
      produtosDAO.iTipoProduto  := 0;
      produtosDAO.iUnidade      := 0;

      for k := 1 to cklstBuscarpor.Items.Count - 1 do
        if cklstBuscarpor.Checked[k] then
        begin
          case k of
            1:
              produtosDAO.sCodigo := edtPesqCodigo.Text;
            2:
              produtosDAO.sDescricao := edtPesqDescricao.Text;
            3:
              produtosDAO.iGrupoProduto := lkcPesqGrupoProduto.KeyValue;
            4:
              produtosDAO.sCodigoBarra := edtPesqCodigoBarra.Text;
            5:
              produtosDAO.iNcm := lkcPesqNCMs.KeyValue;
            6:
              produtosDAO.iTipoProduto := lkcPesqTipoProduto.KeyValue;
            7:
              produtosDAO.iUnidade := lkcPesqUnidadesMedidas.KeyValue;
          end;
        end;

        if not cdsControle.Active then
        begin
          cdsControle.CreateDataSet;
          cdsControle.Close;
          cdsControle.Data := produtosDAO.Pesquisar;
          cdsControle.Open;
        end
        else
        begin
          if cdsControle.RecordCount > 0 then
            cdsControle.EmptyDataSet;

          cdsControle.Close;
          cdsControle.Data := produtosDAO.Pesquisar;
          cdsControle.Open;
        end;

    except
        on E: Exception do
        begin
          raise Exception.Create('Erro ao Localizar Produto ' + E.Message)
        end;
    end;

end;

procedure TFrmProdutos.spdCancelarClick(Sender: TObject);
begin
    cdsProdutos.Cancel;
    cdsProdutosEmpresas.Cancel;

    visualizar  := true;
    deletar     := false;
    inserir     := false;
    alterar     := false;
    PageControl.ActivePage  := TabGrid;
end;

procedure TFrmProdutos.spdExcluirClick(Sender: TObject);
var
  codigo : Integer;
begin
    codigo      := cdsControleID.AsInteger;
    Try
        if not cdsProdutos.Active then
        begin
            cdsProdutos.CreateDataSet;
            cdsProdutos.Close;
            cdsProdutos.Data        := produtosDAO.CarregaProdutos(codigo);
            cdsProdutos.Open;
        end
        else
        begin
            if cdsProdutos.RecordCount > 0 then
               cdsProdutos.EmptyDataSet;

            cdsProdutos.Close;
            cdsProdutos.Data        := produtosDAO.CarregaProdutos(codigo);
            cdsProdutos.Open;
        end;

        if not cdsProdutosEmpresas.Active then
        begin
            cdsProdutosEmpresas.CreateDataSet;
            cdsProdutosEmpresas.Close;
            cdsProdutosEmpresas.Data  := produtosDAO.CarregaProdutosEmpresas(codigo);
            cdsProdutosEmpresas.Open;
        end
        else
        begin
            if cdsProdutosEmpresas.RecordCount > 0 then
               cdsProdutosEmpresas.EmptyDataSet;

            cdsProdutosEmpresas.Close;
            cdsProdutosEmpresas.Data  := produtosDAO.CarregaProdutosEmpresas(codigo);
            cdsProdutosEmpresas.Open;
        end;


        if not produtosDAO.Excluir(cdsProdutos.Data, cdsProdutosEmpresas.Data,codigo) then
           Application.MessageBox('Erro ao Excluir Registro !',PChar(Caption), MB_ICONEXCLAMATION)
        else
           Application.MessageBox('Registro Excluido Com Sucesso !',PChar(Caption), MB_ICONEXCLAMATION);
    Finally
        spdBuscarClick(sender);
    End;

end;

procedure TFrmProdutos.spdGravarClick(Sender: TObject);
var
  iProduto : Integer;
begin
    iProduto := cdsProdutosID.AsInteger;
    try
        try
            if inserir then
            begin
                cdsProdutos.Post;
                if not produtosDAO.Inserir(cdsProdutos.Data, cdsProdutosEmpresas.Data,iProduto) then
                   Application.MessageBox('Erro ao Gravar Registro !',PChar(Caption), MB_ICONEXCLAMATION)
                else
                   Application.MessageBox('Registro Gravado Com Sucesso !',PChar(Caption), MB_ICONEXCLAMATION);
            end;

            if alterar then
            begin
                cdsProdutos.Post;
                if not produtosDAO.Alterar(cdsProdutos.Data, cdsProdutosEmpresas.Data,iProduto) then
                   Application.MessageBox('Erro ao Alterar Registro !',PChar(Caption), MB_ICONEXCLAMATION)
                else
                   Application.MessageBox('Registro Alterado Com Sucesso !',PChar(Caption), MB_ICONEXCLAMATION);
            end;

        except
            on E:Exception do
            begin
                raise E.Create('Erro ao Gravar ou Alterar Registro - '+E.Message);
            end;
        end;
    finally
        spdRetornarClick(sender);
    end;

end;

function TFrmProdutos.IfThen(AValue: Boolean; const ATrue: Integer; const AFalse: Integer): Integer;
begin
  if AValue then
    Result := ATrue
  else
    Result := AFalse;
end;

function TFrmProdutos.IfThen(AValue: Boolean; const ATrue: Int64; const AFalse: Int64): Int64;
begin
  if AValue then
    Result := ATrue
  else
    Result := AFalse;
end;

function TFrmProdutos.IfThen(AValue: Boolean; const ATrue: Double; const AFalse: Double): Double;
begin
  if AValue then
    Result := ATrue
  else
    Result := AFalse;
end;

function TFrmProdutos.IfThen(AValue: Boolean; const ATrue: string;
  AFalse: string = ''): string;
begin
  if AValue then
    Result := ATrue
  else
    Result := AFalse;
end;

function TFrmProdutos.CharIsNum(const C: Char): Boolean;
begin
  Result := CharInSet( C, ['0'..'9'] ) ;
end;

function TFrmProdutos.StrIsNumber(const S: String): Boolean;
Var
  A, LenStr : Integer ;
begin
  LenStr := Length( S ) ;
  Result := (LenStr > 0) ;
  A      := 1 ;
  while Result and ( A <= LenStr )  do
  begin
     Result := CharIsNum( S[A] ) ;
     Inc(A) ;
  end;
end ;
function TFrmProdutos.RPad(const s: String; Pad: Integer): String;
begin
if (Trim(s) = EmptyStr) or (Pad = 0) then
    Result := EmptyStr
else
    Result := Format('¬-*.*s', [Pad, Pad, s]);
end;

function TFrmProdutos.LPad(const s: String; Pad: Integer): String;
begin
    if (Trim(s) = EmptyStr) or (Pad = 0) then
        Result := EmptyStr
    else
        Result := Format('¬*.*s', [Pad, Pad, s]);
end;

function TFrmProdutos.EAN13_DV(CodEAN13: String): String;
Var A,DV : Integer ;
begin
   Result   := '' ;
   CodEAN13 := Rpad(Trim(CodEAN13),12) ;
   if not StrIsNumber( CodEAN13 ) then
      exit ;

   DV := 0;
   For A := 12 downto 1 do
      DV := DV + (StrToInt( CodEAN13[A] ) * StrToInt(IfThen(odd(A),'1','3')));

   DV := (Ceil( DV / 10 ) * 10) - DV ;

   Result := IntToStr( DV );
end;

function TFrmProdutos.PreencheString(Palavra, Caracter, Direcao: string;  Tamanho: integer): String;
var
 v : integer;
begin
    Palavra := Trim(Palavra);

    Result := '';
    for v:=1 to Tamanho-Length(Palavra) do Result := Result + Caracter;

    if Direcao = 'E' then
       Result := Result + Palavra
    else
       Result := Palavra + Result;
end;

function TFrmProdutos.RemoveCaracteresNumero(Str: String): String;
begin
     Result := StringReplace(Str,'.','',[rfReplaceAll]);
     Result := StringReplace(Str,',','',[rfReplaceAll]);
     Result := StringReplace(Str,'/','',[rfReplaceAll]);
     Result := StringReplace(Str,'-','',[rfReplaceAll]);
end;

procedure TFrmProdutos.edtCodigoExit(Sender: TObject);
var
  Codigo : String;
  CodigoAux : String;
  I : Integer;
begin
  Codigo := '';
  CodigoAux := '2';

  for I := 1 to (11 - Length(cdsProdutosPRD_CODIGO.AsString)) do
  begin
    CodigoAux := CodigoAux + '0';
  end;

  Codigo := CodigoAux + cdsProdutosPRD_CODIGO.AsString;

  cdsProdutosPRD_CODBARRAS.AsString := Codigo + EAN13_DV(Codigo);

  edtCodigoBarras.SetFocus;
end;

function TFrmProdutos.FormatarIntStr(Numero, Tamanho: Integer): string;
var
 NumeroStr : string;
begin
    if Numero > 0 then
       NumeroStr :=  IntToStr(Numero) else NumeroStr := '0';

    NumeroStr := RemoveCaracteresNumero(NumeroStr);

    Result := PreencheString(NumeroStr,'0','E',Tamanho);
end;

procedure TFrmProdutos.spdInserirClick(Sender: TObject);
var
  codigo : string;
begin
    visualizar  := False;
    deletar     := false;
    inserir     := true;
    alterar     := false;

    PageControl.ActivePage  := tabFormulario;

    if not cdsProdutos.Active then
    begin
        cdsProdutos.CreateDataSet;
        cdsProdutos.Close;
        cdsProdutos.Data        := produtosDAO.CarregaProdutos(-1);
        cdsProdutos.Open;
    end
    else
    begin
        if cdsProdutos.RecordCount > 0 then
           cdsProdutos.EmptyDataSet;

        cdsProdutos.Close;
        cdsProdutos.Data        := produtosDAO.CarregaProdutos(-1);
        cdsProdutos.Open;
    end;

    if not cdsGrupoProdutos.Active then
    begin
        cdsGrupoProdutos.CreateDataSet;
        cdsGrupoProdutos.Close;
        cdsGrupoProdutos.Data   := produtosDAO.CarregaGruposProdutos;
        cdsGrupoProdutos.Open;
    end
    else
    begin
        if cdsGrupoProdutos.RecordCount > 0 then
           cdsGrupoProdutos.EmptyDataSet;

        cdsGrupoProdutos.Close;
        cdsGrupoProdutos.Data   := produtosDAO.CarregaGruposProdutos;
        cdsGrupoProdutos.Open;
    end;

    if not cdsProdutosEmpresas.Active then
    begin
        cdsProdutosEmpresas.CreateDataSet;
        cdsProdutosEmpresas.Close;
        cdsProdutosEmpresas.Data  := produtosDAO.CarregaProdutosEmpresas(-1);
        cdsProdutosEmpresas.Open;
    end
    else
    begin
        if cdsProdutosEmpresas.RecordCount > 0 then
           cdsProdutosEmpresas.EmptyDataSet;

        cdsProdutosEmpresas.Close;
        cdsProdutosEmpresas.Data  := produtosDAO.CarregaProdutosEmpresas(-1);
        cdsProdutosEmpresas.Open;
    end;

    if not cdsNCMs.Active then
    begin
        cdsNCMs.CreateDataSet;
        cdsNCMs.Close;
        cdsNCMs.Data    := produtosDAO.CarregaNcms;
        cdsNCMs.Open;
    end
    else
    begin
        if cdsNCMs.RecordCount > 0 then
           cdsNCMs.EmptyDataSet;

        cdsNCMs.Close;
        cdsNCMs.Data    := produtosDAO.CarregaNcms;
        cdsNCMs.Open;
    end;

    if not cdsUnidadesMedidas.Active then
    begin
        cdsUnidadesMedidas.CreateDataSet;
        cdsUnidadesMedidas.Close;
        cdsUnidadesMedidas.Data   := produtosDAO.CarregaUnidadesMedidas;
        cdsUnidadesMedidas.Open;
    end
    else
    begin
        if cdsUnidadesMedidas.Recordcount > 0 then
           cdsUnidadesMedidas.EmptyDataSet;

        cdsUnidadesMedidas.Close;
        cdsUnidadesMedidas.Data   := produtosDAO.CarregaUnidadesMedidas;
        cdsUnidadesMedidas.Open;
    end;

    codigo := FormatarIntStr(produtosDAO.RetornaID('IND_PRODUTOS')+1,6);

    cdsProdutos.Insert;
    cdsProdutosID.AsInteger               := StrToInt(codigo);
    cdsProdutosPRD_CODIGO.AsString        := codigo;
    cdsProdutosPRD_TIPOPRODUTO.AsInteger  := 0;
    cdsProdutosAUD_DEL.AsInteger          := 0;
    cdsProdutosAUD_DATCAD.AsDateTime      := Date;
    cdsProdutosAUD_USUCAD.AsString        := 'Administrador';
    edtCodigo.Enabled := True;
    edtCodigo.SetFocus;

end;

procedure TFrmProdutos.spdRetornarClick(Sender: TObject);
begin
    if PageControl.ActivePageIndex = 1 then
    begin
        cdsProdutos.Cancel;
        cdsProdutosEmpresas.Cancel;

        visualizar  := true;
        deletar     := false;
        inserir     := false;
        alterar     := false;
        PageControl.ActivePage  := TabGrid;
    end
    else Close;


end;

end.
