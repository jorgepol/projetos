unit uClsProdutosDAO;

interface

Uses
   Data.DBXMSSQL, Data.FMTBcd, Data.SqlExpr, Datasnap.Provider,
   Data.DB, Datasnap.DBClient, System.SysUtils, Vcl.Forms,
   uClsEmpresasDAO;

Type
  TProdutosDAO = class(TObject)
    private
      sComando : String;
      Fconexao: TSQLConnection;
      FdsPesquisa: TDataSource;
      FqryUnidadesMedidas: TSQLQuery;
      FdspProdutosEmpresas: TDataSetProvider;
      FsdsExcUnidadesMedidas: TSQLDataSet;
      FdsUnidadesMedidas: TDataSource;
      FcdsProdutos: TClientDataSet;
      FdspNcms: TDataSetProvider;
      FqryProdutosEmpresas: TSQLQuery;
      FsdsExcProdutosEmpresas: TSQLDataSet;
      FcdsPesquisa: TClientDataSet;
      FdsProdutosEmpresas: TDataSource;
      FdspEmpresas: TDataSetProvider;
      FdspGruposProdutos: TDataSetProvider;
      FcdsUnidadesMedidas: TClientDataSet;
      FqryNcms: TSQLQuery;
      FsdsExcNcms: TSQLDataSet;
      FdsNcms: TDataSource;
      FqryEmpresas: TSQLQuery;
      FsdsExcEmpresas: TSQLDataSet;
      FqryGruposProdutos: TSQLQuery;
      FdsEmpresas: TDataSource;
      FcdsProdutosEmpresas: TClientDataSet;
      FsdsExcGruposProdutos: TSQLDataSet;
      FdsGruposProdutos: TDataSource;
      FcdsNcms: TClientDataSet;
      FdspProdutos: TDataSetProvider;
      FcdsEmpresas: TClientDataSet;
      FdspPesquisa: TDataSetProvider;
      FcdsGruposProdutos: TClientDataSet;
      FqryProdutos: TSQLQuery;
      FsdsExcProdutos: TSQLDataSet;
      FdspUnidadesMedidas: TDataSetProvider;
      FdsProdutos: TDataSource;
      FqryPesquisa: TSQLQuery;
      FsdsExcPesquisa: TSQLDataSet;
      FqryDados: TSQLQuery;
      FsdsExcDados: TSQLDataSet;
      FdsDados: TDataSource;
      FcdsDados: TClientDataSet;
      FdspDados: TDataSetProvider;
      procedure Setconexao(const Value: TSQLConnection);
      procedure SetcdsEmpresas(const Value: TClientDataSet);
      procedure SetcdsGruposProdutos(const Value: TClientDataSet);
      procedure SetcdsNcms(const Value: TClientDataSet);
      procedure SetcdsPesquisa(const Value: TClientDataSet);
      procedure SetcdsProdutos(const Value: TClientDataSet);
      procedure SetcdsProdutosEmpresas(const Value: TClientDataSet);
      procedure SetcdsUnidadesMedidas(const Value: TClientDataSet);
      procedure SetdsEmpresas(const Value: TDataSource);
      procedure SetdsGruposProdutos(const Value: TDataSource);
      procedure SetdsNcms(const Value: TDataSource);
      procedure SetdspEmpresas(const Value: TDataSetProvider);
      procedure SetdsPesquisa(const Value: TDataSource);
      procedure SetdspGruposProdutos(const Value: TDataSetProvider);
      procedure SetdspNcms(const Value: TDataSetProvider);
      procedure SetdspPesquisa(const Value: TDataSetProvider);
      procedure SetdspProdutos(const Value: TDataSetProvider);
      procedure SetdspProdutosEmpresas(const Value: TDataSetProvider);
      procedure SetdsProdutos(const Value: TDataSource);
      procedure SetdsProdutosEmpresas(const Value: TDataSource);
      procedure SetdspUnidadesMedidas(const Value: TDataSetProvider);
      procedure SetdsUnidadesMedidas(const Value: TDataSource);
      procedure SetqryEmpresas(const Value: TSQLQuery);
      procedure SetqryGruposProdutos(const Value: TSQLQuery);
      procedure SetqryNcms(const Value: TSQLQuery);
      procedure SetqryPesquisa(const Value: TSQLQuery);
      procedure SetqryProdutos(const Value: TSQLQuery);
      procedure SetqryProdutosEmpresas(const Value: TSQLQuery);
      procedure SetqryUnidadesMedidas(const Value: TSQLQuery);
      procedure SetsdsExcEmpresas(const Value: TSQLDataSet);
      procedure SetsdsExcGruposProdutos(const Value: TSQLDataSet);
      procedure SetsdsExcNcms(const Value: TSQLDataSet);
      procedure SetsdsExcPesquisa(const Value: TSQLDataSet);
      procedure SetsdsExcProdutos(const Value: TSQLDataSet);
      procedure SetsdsExcProdutosEmpresas(const Value: TSQLDataSet);
      procedure SetsdsExcUnidadesMedidas(const Value: TSQLDataSet);
      procedure SetcdsDados(const Value: TClientDataSet);
      procedure SetdsDados(const Value: TDataSource);
      procedure SetdspDados(const Value: TDataSetProvider);
      procedure SetqryDados(const Value: TSQLQuery);
      procedure SetsdsExcDados(const Value: TSQLDataSet);

    public
      iGrupoProduto, iNcm, iUnidade, iTipoProduto  : Integer;
      sCodigoBarra, sCodigo, sDescricao : String;

      constructor Create; overload;
      destructor Destroy; overload;

      function Inserir(dProduto, dProdutoEmpresa : Olevariant; pProduto : Integer):Boolean;
      function Alterar(dProduto, dProdutoEmpresa : Olevariant; pProduto : Integer):Boolean;
      function Excluir(dProduto, dProdutoEmpresa : Olevariant; pProduto : Integer):Boolean;
      function Pesquisar: Olevariant;
      function CarregaNcms:Olevariant;
      function CarregaGruposProdutos:Olevariant;
      function CarregaUnidadesMedidas:Olevariant;
      function CarregaEmpresas:Olevariant;
      function CarregaProdutos(pProduto : Integer):Olevariant;
      function CarregaProdutosEmpresas(pProduto : Integer):Olevariant;

      function SubstituirCaracter(Str: String; const Carac : String; subst: Char = #255):String;
      function DecimalSQL(const valorDecimal : Double):String;
      function RetornaID(tabela : String):Integer;

      procedure CriarDatsSets;
      procedure CriarDataSetDados;
      procedure CriarDataSetPesquisa;
      procedure CriarDataSetProdutos;
      procedure CriarDataSetProdutosEmpresas;
      procedure CriarDataSetEmpresas;
      procedure CriarDataSetNcms;
      procedure CriarDataSetUnidadesMedidas;
      procedure CriarDataSetGruposProdutos;
      procedure DestruirDataSets;
      procedure DestruirDataSetDados;
      procedure DestruirDataSetPesquisa;
      procedure DestruirDataSetProdutos;
      procedure DestruirDataSetProdutosEmpresas;
      procedure DestruirDataSetEmpresas;
      procedure DestruirDataSetNcms;
      procedure DestruirDataSetUnidadesMedidas;
      procedure DestruirDataSetGruposProdutos;


    published
      property conexao                : TSQLConnection        read Fconexao                 write Setconexao;
      property qryDados               : TSQLQuery             read FqryDados                write SetqryDados;
      property dspDados               : TDataSetProvider      read FdspDados                write SetdspDados;
      property cdsDados               : TClientDataSet        read FcdsDados                write SetcdsDados;
      property dsDados                : TDataSource           read FdsDados                 write SetdsDados;
      property sdsExcDados            : TSQLDataSet           read FsdsExcDados             write SetsdsExcDados;
      property qryPesquisa            : TSQLQuery             read FqryPesquisa             write SetqryPesquisa;
      property dspPesquisa            : TDataSetProvider      read FdspPesquisa             write SetdspPesquisa;
      property cdsPesquisa            : TClientDataSet        read FcdsPesquisa             write SetcdsPesquisa;
      property dsPesquisa             : TDataSource           read FdsPesquisa              write SetdsPesquisa;
      property sdsExcPesquisa         : TSQLDataSet           read FsdsExcPesquisa          write SetsdsExcPesquisa;
      property qryProdutos            : TSQLQuery             read FqryProdutos             write SetqryProdutos;
      property dspProdutos            : TDataSetProvider      read FdspProdutos             write SetdspProdutos;
      property cdsProdutos            : TClientDataSet        read FcdsProdutos             write SetcdsProdutos;
      property dsProdutos             : TDataSource           read FdsProdutos              write SetdsProdutos;
      property sdsExcProdutos         : TSQLDataSet           read FsdsExcProdutos          write SetsdsExcProdutos;
      property qryProdutosEmpresas    : TSQLQuery             read FqryProdutosEmpresas     write SetqryProdutosEmpresas;
      property dspProdutosEmpresas    : TDataSetProvider      read FdspProdutosEmpresas     write SetdspProdutosEmpresas;
      property cdsProdutosEmpresas    : TClientDataSet        read FcdsProdutosEmpresas     write SetcdsProdutosEmpresas;
      property dsProdutosEmpresas     : TDataSource           read FdsProdutosEmpresas      write SetdsProdutosEmpresas;
      property sdsExcProdutosEmpresas : TSQLDataSet           read FsdsExcProdutosEmpresas  write SetsdsExcProdutosEmpresas;
      property qryEmpresas            : TSQLQuery             read FqryEmpresas             write SetqryEmpresas;
      property dspEmpresas            : TDataSetProvider      read FdspEmpresas             write SetdspEmpresas;
      property cdsEmpresas            : TClientDataSet        read FcdsEmpresas             write SetcdsEmpresas;
      property dsEmpresas             : TDataSource           read FdsEmpresas              write SetdsEmpresas;
      property sdsExcEmpresas         : TSQLDataSet           read FsdsExcEmpresas          write SetsdsExcEmpresas;
      property qryNcms                : TSQLQuery             read FqryNcms                 write SetqryNcms;
      property dspNcms                : TDataSetProvider      read FdspNcms                 write SetdspNcms;
      property cdsNcms                : TClientDataSet        read FcdsNcms                 write SetcdsNcms;
      property dsNcms                 : TDataSource           read FdsNcms                  write SetdsNcms;
      property sdsExcNcms             : TSQLDataSet           read FsdsExcNcms              write SetsdsExcNcms;
      property qryUnidadesMedidas     : TSQLQuery             read FqryUnidadesMedidas      write SetqryUnidadesMedidas;
      property dspUnidadesMedidas     : TDataSetProvider      read FdspUnidadesMedidas      write SetdspUnidadesMedidas;
      property cdsUnidadesMedidas     : TClientDataSet        read FcdsUnidadesMedidas      write SetcdsUnidadesMedidas;
      property dsUnidadesMedidas      : TDataSource           read FdsUnidadesMedidas       write SetdsUnidadesMedidas;
      property sdsExcUnidadesMedidas  : TSQLDataSet           read FsdsExcUnidadesMedidas   write SetsdsExcUnidadesMedidas;
      property qryGruposProdutos      : TSQLQuery             read FqryGruposProdutos       write SetqryGruposProdutos;
      property dspGruposProdutos      : TDataSetProvider      read FdspGruposProdutos       write SetdspGruposProdutos;
      property cdsGruposProdutos      : TClientDataSet        read FcdsGruposProdutos       write SetcdsGruposProdutos;
      property dsGruposProdutos       : TDataSource           read FdsGruposProdutos        write SetdsGruposProdutos;
      property sdsExcGruposProdutos   : TSQLDataSet           read FsdsExcGruposProdutos    write SetsdsExcGruposProdutos;

  end;

implementation

{ TProdutosDAO }

function TProdutosDAO.Alterar(dProduto, dProdutoEmpresa: Olevariant; pProduto: Integer): Boolean;
var
  TD : TTransactionDesc;
  produtoID : Integer;
begin
    Try
        Try
            CriarDataSetProdutos;

            Result  := False;

            cdsProdutos.Close;
            cdsProdutos.Data      := dProduto;
            cdsProdutos.Filtered  := True;
            cdsProdutos.Filter    := 'ID = '+IntToStr(pProduto);
            cdsProdutos.Open;
            if cdsProdutos.RecordCount > 0 then
            begin
                if not conexao.InTransaction then
                begin
                    TD.TransactionID    := 99999;
                    TD.IsolationLevel   := xilREADCOMMITTED;
                    conexao.StartTransaction(TD);
                end;

                sComando  := ' '+#13;
                sComando  := sComando +'UPDATE IND_PRODUTOS '+#13;
                sComando  := sComando +'SET GRUPO_PRODUTO_ID    = '+cdsProdutos.FieldByName('GRUPO_PRODUTO_ID').AsString+', '+#13;
                sComando  := sComando +' UNIDADE_MEDIDA_ID      = '+cdsProdutos.FieldByName('UNIDADE_MEDIDA_ID').AsString+', '+#13;
                sComando  := sComando +' NCM_ID                 = '+cdsProdutos.FieldByName('NCM_ID').AsString+', '+#13;
                sComando  := sComando +' PRD_CODBARRAS          = '+QuotedStr(cdsProdutos.FieldByName('PRD_CODBARRAS').AsString)+', '+#13;
                sComando  := sComando +' PRD_CODIGO             = '+QuotedStr(cdsProdutos.FieldByName('PRD_CODIGO').AsString)+', '+#13;
                sComando  := sComando +' PRD_APELIDO            = '+QuotedStr(cdsProdutos.FieldByName('PRD_APELIDO').AsString)+', '+#13;
                sComando  := sComando +' PRD_DESCRICAO          = '+QuotedStr(cdsProdutos.FieldByName('PRD_DESCRICAO').AsString)+', '+#13;
                sComando  := sComando +' PRD_DESCRICAOFISCAL    = '+QuotedStr(cdsProdutos.FieldByName('PRD_DESCRICAOFISCAL').AsString)+', '+#13;
                sComando  := sComando +' PRD_DESCRICAOTECNICA   = '+QuotedStr(cdsProdutos.FieldByName('PRD_DESCRICAOTECNICA').AsWideString)+', '+#13;
                sComando  := sComando +' PRD_TIPOPRODUTO        = '+cdsProdutos.FieldByName('PRD_TIPOPRODUTO').AsString+', '+#13;
                sComando  := sComando +' PRD_PESOBRUTO          = '+QuotedStr(DecimalSQL(cdsProdutos.FieldByName('PRD_PESOBRUTO').AsFloat))+', '+#13;
                sComando  := sComando +' PRD_PESOLIQUIDO        = '+QuotedStr(DecimalSQL(cdsProdutos.FieldByName('PRD_PESOLIQUIDO').AsFloat))+', '+#13;
                sComando  := sComando +' PRD_FOTO               = '+QuotedStr(cdsProdutos.FieldByName('PRD_FOTO').AsWideString)+', '+#13;
                sComando  := sComando +' AUD_DEL                = 0, '+#13;
                sComando  := sComando +' AUD_DATULTALT          = '+QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',Date))+', '+#13;
                sComando  := sComando +' AUD_USUULTALT          = '+QuotedStr('Administrador')+' '+#13;
                sComando  := sComando +' WHERE ID = '+IntToStr(pProduto);
                sComando  := sComando +' '+#13;

                cdsProdutosEmpresas.Close;
                cdsProdutosEmpresas.Data      := dProdutoEmpresa;
                cdsProdutosEmpresas.Filtered  := True;
                cdsProdutosEmpresas.Filter    := 'PRODUTO_ID = '+IntToStr(pProduto);
                cdsProdutosEmpresas.Open;
                if cdsProdutosEmpresas.RecordCount > 0 then
                begin
                    while not cdsProdutosEmpresas.Eof do
                    begin
                        qryProdutosEmpresas   := TSQLQuery.Create(nil);
                        qryProdutosEmpresas.SQL.Text := 'SELECT * FROM IND_PRODUTOS_EMPRESAS WHERE ID = '+cdsProdutosEmpresas.FieldByName('ID').AsString;
                        qryProdutosEmpresas.Open;
                        if qryProdutosEmpresas.Eof then
                        begin
                            sComando  := sComando +' '+#13;
                            sComando  := sComando +'INSERT INTO IND_PRODUTOS_EMPRESAS '+#13;
                            sComando  := sComando +'( '+#13;
                            sComando  := sComando +'EMPRESA_ID,           '+#13;
                            sComando  := sComando +'PRODUTO_ID,           '+#13;
                            sComando  := sComando +'PRE_PRECOCUSTO,       '+#13;
                            sComando  := sComando +'PRE_CUSTOMEDIO,       '+#13;
                            sComando  := sComando +'PRE_CUSTOCONTABIL,    '+#13;
                            sComando  := sComando +'PRE_CUSTOGERENCIAL,   '+#13;
                            sComando  := sComando +'PRE_PRECOVENDA,       '+#13;
                            sComando  := sComando +'PRE_PERCCUSTOOPER,    '+#13;
                            sComando  := sComando +'PRE_MARGEMLUCRO,      '+#13;
                            sComando  := sComando +'PRE_DATAULTVENDA,     '+#13;
                            sComando  := sComando +'PRE_VLRULTVENDA,      '+#13;
                            sComando  := sComando +'PRE_DATAULTCOMPRA,    '+#13;
                            sComando  := sComando +'PRE_VLRULTCOMPRA,     '+#13;
                            sComando  := sComando +'PRE_SALDO_ANTERIOR,   '+#13;
                            sComando  := sComando +'PRE_SALDO_ATUAL,      '+#13;
                            sComando  := sComando +'PRE_SALDO_DISPONIVEL, '+#13;
                            sComando  := sComando +'PRE_SALDO_RESERVADO,  '+#13;
                            sComando  := sComando +'PRE_STATUS,           '+#13;
                            sComando  := sComando +'AUD_DEL,              '+#13;
                            sComando  := sComando +'AUD_DATCAD,           '+#13;
                            sComando  := sComando +'AUD_USUCAD            '+#13;
                            sComando  := sComando +' ) ';
                            sComando  := sComando +' VALUES ( '+#13;
                            sComando  := sComando +' '+cdsProdutosEmpresas.FieldByName('EMPRESA_ID').AsString+', '+#13;
                            sComando  := sComando +' '+IntToStr(produtoID)+', '+#13;
                            sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_PRECOCUSTO').AsFloat))+', '+#13;
                            sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_CUSTOMEDIO').AsFloat))+', '+#13;
                            sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_CUSTOCONTABIL').AsFloat))+', '+#13;
                            sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_CUSTOGERENCIAL').AsFloat))+', '+#13;
                            sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_PRECOVENDA').AsFloat))+', '+#13;
                            sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_PERCCUSTOOPER').AsFloat))+', '+#13;
                            sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_MARGEMLUCRO').AsFloat))+', '+#13;
                            sComando  := sComando +' '+QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',cdsProdutosEmpresas.FieldByName('PRE_DATAULTVENDA').AsDateTime))+', '+#13;
                            sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_VLRULTVENDA').AsFloat))+', '+#13;
                            sComando  := sComando +' '+QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',cdsProdutosEmpresas.FieldByName('PRE_DATAULTCOMPRA').AsDateTime))+', '+#13;
                            sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_VLRULTCOMPRA').AsFloat))+', '+#13;
                            sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_SALDO_ANTERIOR').AsFloat))+', '+#13;
                            sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_SALDO_ATUAL').AsFloat))+', '+#13;
                            sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_SALDO_DISPONIVEL').AsFloat))+', '+#13;
                            sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_SALDO_RESERVADO').AsFloat))+', '+#13;
                            sComando  := sComando +' 0, '+#13;
                            sComando  := sComando +' 0, '+#13;
                            sComando  := sComando +' '+QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',Date))+', '+#13;
                            sComando  := sComando +' '+QuotedStr('Administrador')+' '+#13;
                            sComando  := sComando +' ) '+#13;
                        end
                        else
                        begin
                            sComando  := sComando +' '+#13;
                            sComando  := sComando +'UPDATE IND_PRODUTOS_EMPRESAS '+#13;
                            sComando  := sComando +'SET EMPRESA_ID        = '+cdsProdutosEmpresas.FieldByName('EMPRESA_ID').AsString+', '+#13;
                            sComando  := sComando +'PRODUTO_ID            = '+cdsProdutosEmpresas.FieldByName('PRODUTO_ID').AsString+', '+#13;
                            sComando  := sComando +'PRE_PRECOCUSTO        = '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_PRECOCUSTO').AsFloat))+', '+#13;
                            sComando  := sComando +'PRE_CUSTOMEDIO        = '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_CUSTOMEDIO').AsFloat))+', '+#13;
                            sComando  := sComando +'PRE_CUSTOCONTABIL     = '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_CUSTOCONTABIL').AsFloat))+', '+#13;
                            sComando  := sComando +'PRE_CUSTOGERENCIAL    = '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_CUSTOGERENCIAL').AsFloat))+', '+#13;
                            sComando  := sComando +'PRE_PRECOVENDA        = '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_PRECOVENDA').AsFloat))+', '+#13;
                            sComando  := sComando +'PRE_PERCCUSTOOPER     = '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_PERCCUSTOOPER').AsFloat))+', '+#13;
                            sComando  := sComando +'PRE_MARGEMLUCRO       = '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_MARGEMLUCRO').AsFloat))+', '+#13;
                            sComando  := sComando +'PRE_DATAULTVENDA      = '+QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',cdsProdutosEmpresas.FieldByName('PRE_DATAULTVENDA').AsDateTime))+', '+#13;
                            sComando  := sComando +'PRE_VLRULTVENDA       = '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_VLRULTVENDA').AsFloat))+', '+#13;
                            sComando  := sComando +'PRE_DATAULTCOMPRA     = '+QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',cdsProdutosEmpresas.FieldByName('PRE_DATAULTCOMPRA').AsDateTime))+', '+#13;
                            sComando  := sComando +'PRE_VLRULTCOMPRA      = '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_VLRULTCOMPRA').AsFloat))+', '+#13;
                            sComando  := sComando +'PRE_SALDO_ANTERIOR    = '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_SALDO_ANTERIOR').AsFloat))+', '+#13;
                            sComando  := sComando +'PRE_SALDO_ATUAL       = '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_SALDO_ATUAL').AsFloat))+', '+#13;
                            sComando  := sComando +'PRE_SALDO_DISPONIVEL  = '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_SALDO_DISPONIVEL').AsFloat))+', '+#13;
                            sComando  := sComando +'PRE_SALDO_RESERVADO   = '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_SALDO_RESERVADO').AsFloat))+', '+#13;
                            sComando  := sComando +'PRE_STATUS            = 0, '+#13;
                            sComando  := sComando +'AUD_DEL               = 0, '+#13;
                            sComando  := sComando +'AUD_DATULTALT         = '+QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',Date))+', '+#13;
                            sComando  := sComando +'AUD_USUULTALT         = '+QuotedStr('Administrador')+' '+#13;
                            sComando  := sComando +'WHERE PRODUTO_ID      = '+IntToStr(pProduto);
                        end;

                        cdsProdutosEmpresas.Next;
                    end;

                end;

               sdsExcProdutos.Close;
               sdsExcProdutos.CommandText := sComando;
               sdsExcProdutos.ExecSQL;

               if conexao.InTransaction then
                 conexao.Commit(TD);

               result := True;

            end;

        except
            On E:Exception do
            begin
               if conexao.InTransaction then
                 conexao.Rollback(TD);

               raise E.Create('Erro ao Inserir Registro '+E.Message);
            end;
        End;
    Finally
        DestruirDataSetProdutos;
    End;

end;

function TProdutosDAO.CarregaEmpresas: Olevariant;
begin
    Try
        Try
            CriarDataSetEmpresas;

            sComando  := ' SELECT ID,  '+#13;
            sComando  := sComando +' EMP_RAZAOSOCIAL,     '+#13;
            sComando  := sComando +' EMP_NOMEFANTASIA,    '+#13;
            sComando  := sComando +' EMP_TIPOPESSOA,      '+#13;
            sComando  := sComando +' EMP_CNPJCPF,         '+#13;
            sComando  := sComando +' EMP_INSCRESTADUAL,   '+#13;
            sComando  := sComando +' EMP_INSCRMUNICIPAL,  '+#13;
            sComando  := sComando +' EMP_PISCOFINS,       '+#13;
            sComando  := sComando +' EMP_ALIQPIS,         '+#13;
            sComando  := sComando +' EMP_ALIQCOFINS,      '+#13;
            sComando  := sComando +' EMP_CEP,             '+#13;
            sComando  := sComando +' EMP_LOGRADOURO,      '+#13;
            sComando  := sComando +' EMP_NUMERO,          '+#13;
            sComando  := sComando +' EMP_BAIRRO,          '+#13;
            sComando  := sComando +' EMP_MUNICIPIO,       '+#13;
            sComando  := sComando +' EMP_UF,              '+#13;
            sComando  := sComando +' EMP_PAGINA,          '+#13;
            sComando  := sComando +' EMAIL,               '+#13;
            sComando  := sComando +' AUD_DEL,             '+#13;
            sComando  := sComando +' AUD_DATCAD,          '+#13;
            sComando  := sComando +' AUD_USUCAD,          '+#13;
            sComando  := sComando +' AUD_DATULTALT,       '+#13;
            sComando  := sComando +' AUD_USUULTALT,       '+#13;
            sComando  := sComando +' AUD_DATEXC,          '+#13;
            sComando  := sComando +' AUD_USUEXC           '+#13;
            sComando  := sComando +' FROM ADM_EMPRESAS    '+#13;
            sComando  := sComando +' WHERE AUD_DEL = 0 '+#13;

           dspEmpresas.Options  := [poAllowCommandText];

           cdsEmpresas.SetProvider(dspEmpresas);

           cdsEmpresas.Close;
           cdsEmpresas.CommandText := sComando;
           cdsEmpresas.Open;


           result := cdsEmpresas.Data;

        except
            On E:Exception do
            begin
               raise E.Create('Erro ao Carregar Empresas '+E.Message);
            end;
        End;
    Finally
        DestruirDataSetEmpresas;
    End;

end;

function TProdutosDAO.CarregaGruposProdutos: Olevariant;
begin
    Try
        Try
            CriarDataSetGruposProdutos;

            sComando  := ' SELECT ID,  '+#13;
            sComando  := sComando +' GRP_DESCRICAO,     '+#13;
            sComando  := sComando +' AUD_DEL,             '+#13;
            sComando  := sComando +' AUD_DATCAD,          '+#13;
            sComando  := sComando +' AUD_USUCAD,          '+#13;
            sComando  := sComando +' AUD_DATULTALT,       '+#13;
            sComando  := sComando +' AUD_USUULTALT,       '+#13;
            sComando  := sComando +' AUD_DATEXC,          '+#13;
            sComando  := sComando +' AUD_USUEXC           '+#13;
            sComando  := sComando +' FROM IND_GRUPOS_PRODUTOS    '+#13;
            sComando  := sComando +' WHERE AUD_DEL = 0 '+#13;

            cdsGruposProdutos           := TClientDataSet.Create(nil);
            dspGruposProdutos.Options   := [poAllowCommandText];
            cdsGruposProdutos.SetProvider(dspGruposProdutos);
            cdsGruposProdutos.Close;
            cdsGruposProdutos.CommandText := sComando;
            cdsGruposProdutos.Open;

            result := cdsGruposProdutos.Data;

        except
            On E:Exception do
            begin
               raise E.Create('Erro ao Carregar Grupo de Produtos '+E.Message);
            end;
        End;
    Finally
        DestruirDataSetGruposProdutos;
    End;

end;

function TProdutosDAO.CarregaNcms: Olevariant;
begin
    Try
        Try
            CriarDataSetNcms;

            sComando  := ' SELECT ID,  '+#13;
            sComando  := sComando +' CODIGOPAI_ID,        '+#13;
            sComando  := sComando +' NCM_CODIGO,          '+#13;
            sComando  := sComando +' NCM_CODIGOPAI,       '+#13;
            sComando  := sComando +' NCM_DESCRICAO,       '+#13;
            sComando  := sComando +' NCM_TRIBUTADO,       '+#13;
            sComando  := sComando +' NCM_ALIQUOTA,        '+#13;
            sComando  := sComando +' NCM_NOTA,            '+#13;
            sComando  := sComando +' AUD_DEL,             '+#13;
            sComando  := sComando +' AUD_DATCAD,          '+#13;
            sComando  := sComando +' AUD_USUCAD,          '+#13;
            sComando  := sComando +' AUD_DATULTALT,       '+#13;
            sComando  := sComando +' AUD_USUULTALT,       '+#13;
            sComando  := sComando +' AUD_DATEXC,          '+#13;
            sComando  := sComando +' AUD_USUEXC           '+#13;
            sComando  := sComando +' FROM COM_NCMS        '+#13;
            sComando  := sComando +' WHERE AUD_DEL = 0 '+#13;

            cdsNcms           := TClientDataSet.Create(nil);
            dspNcms.Options   := [poAllowCommandText];
            cdsNcms.SetProvider(dspNcms);
            cdsNcms.Close;
            cdsNcms.CommandText := sComando;
            cdsNcms.Open;

            result := cdsNcms.Data;

        except
            On E:Exception do
            begin
               raise E.Create('Erro ao Carregar NCMs '+E.Message);
            end;
        End;
    Finally
        DestruirDataSetNcms;
    End;

end;

function TProdutosDAO.CarregaProdutos(pProduto: Integer): Olevariant;
begin
    Try
        Try
            CriarDataSetProdutos;

            sComando  := ' SELECT PRD.ID,  '+#13;
            sComando  := sComando +' PRD.GRUPO_PRODUTO_ID,                                  '+#13;
            sComando  := sComando +' GRP.GRP_DESCRICAO,                                     '+#13;
            sComando  := sComando +' PRD.UNIDADE_MEDIDA_ID,                                 '+#13;
            sComando  := sComando +' UNM.UNM_SIGLA,                                         '+#13;
            sComando  := sComando +' PRD.NCM_ID,                                            '+#13;
            sComando  := sComando +' NCM.NCM_DESCRICAO,                                     '+#13;
            sComando  := sComando +' PRD.PRD_CODBARRAS,                                     '+#13;
            sComando  := sComando +' PRD.PRD_CODIGO,                                        '+#13;
            sComando  := sComando +' PRD.PRD_APELIDO,                                       '+#13;
            sComando  := sComando +' PRD.PRD_DESCRICAO,                                     '+#13;
            sComando  := sComando +' PRD.PRD_DESCRICAOFISCAL,                               '+#13;
            sComando  := sComando +' PRD.PRD_DESCRICAOTECNICA,                              '+#13;
            sComando  := sComando +' PRD.PRD_TIPOPRODUTO,                                   '+#13;
            sComando  := sComando +' PRD.PRD_PESOBRUTO,                                     '+#13;
            sComando  := sComando +' PRD.PRD_PESOLIQUIDO,                                   '+#13;
            sComando  := sComando +' PRD.PRD_FOTO,                                          '+#13;
            sComando  := sComando +' PRD.AUD_DEL,                                           '+#13;
            sComando  := sComando +' PRD.AUD_DATCAD,                                        '+#13;
            sComando  := sComando +' PRD.AUD_USUCAD,                                        '+#13;
            sComando  := sComando +' PRD.AUD_DATULTALT,                                     '+#13;
            sComando  := sComando +' PRD.AUD_USUULTALT,                                     '+#13;
            sComando  := sComando +' PRD.AUD_DATEXC,                                        '+#13;
            sComando  := sComando +' PRD.AUD_USUEXC                                         '+#13;
            sComando  := sComando +' FROM IND_PRODUTOS AS PRD                               '+#13;
            sComando  := sComando +' LEFT JOIN IND_GRUPOS_PRODUTOS AS GRP                   '+#13;
            sComando  := sComando +' ON GRP.ID = PRD.GRUPO_PRODUTO_ID AND GRP.AUD_DEL = 0   '+#13;
            sComando  := sComando +' LEFT JOIN IND_UNIDADE_MEDIDAS AS UNM                  '+#13;
            sComando  := sComando +' ON UNM.ID = PRD.UNIDADE_MEDIDA_ID AND UNM.AUD_DEL = 0  '+#13;
            sComando  := sComando +' LEFT JOIN COM_NCMS AS NCM                              '+#13;
            sComando  := sComando +' ON NCM.ID = PRD.NCM_ID AND NCM.AUD_DEL = 0             '+#13;
            sComando  := sComando +' WHERE PRD.AUD_DEL = 0 '+#13;

            if pProduto <> 0 then
               sComando  := sComando +' AND PRD.ID = ' +IntToStr(pProduto);

            cdsProdutos           := TClientDataSet.Create(nil);
            dspProdutos.Options   := [poAllowCommandText];
            cdsProdutos.SetProvider(dspProdutos);
            cdsProdutos.Close;
            cdsProdutos.CommandText := sComando;
            cdsProdutos.Open;

            result := cdsProdutos.Data;

        except
            On E:Exception do
            begin
               raise E.Create('Erro ao Carregar Produtos  '+E.Message);
            end;
        End;
    Finally
        DestruirDataSetProdutos;
    End;

end;

function TProdutosDAO.CarregaProdutosEmpresas(pProduto: Integer): Olevariant;
begin
    Try
        Try
            CriarDataSetProdutosEmpresas;

            sComando  := ' SELECT PRE.ID,  '+#13;
            sComando  := sComando +' PRE.EMPRESA_ID,                                        '+#13;
            sComando  := sComando +' EMP.EMP_RAZAOSOCIAL,                                   '+#13;
            sComando  := sComando +' PRE.PRODUTO_ID,                                        '+#13;
            sComando  := sComando +' PRD.PRD_DESCRICAO,                                     '+#13;
            sComando  := sComando +' PRE.PRE_PRECOCUSTO,                                    '+#13;
            sComando  := sComando +' PRE.PRE_CUSTOMEDIO,                                    '+#13;
            sComando  := sComando +' PRE.PRE_CUSTOCONTABIL,                                 '+#13;
            sComando  := sComando +' PRE.PRE_CUSTOGERENCIAL,                                '+#13;
            sComando  := sComando +' PRE.PRE_PRECOVENDA,                                    '+#13;
            sComando  := sComando +' PRE.PRE_PERCCUSTOOPER,                                 '+#13;
            sComando  := sComando +' PRE.PRE_MARGEMLUCRO,                                   '+#13;
            sComando  := sComando +' PRE.PRE_DATAULTVENDA,                                  '+#13;
            sComando  := sComando +' PRE.PRE_VLRULTVENDA,                                   '+#13;
            sComando  := sComando +' PRE.PRE_DATAULTCOMPRA,                                 '+#13;
            sComando  := sComando +' PRE.PRE_VLRULTCOMPRA,                                  '+#13;
            sComando  := sComando +' PRE.PRE_SALDO_ANTERIOR,                                '+#13;
            sComando  := sComando +' PRE.PRE_SALDO_ATUAL,                                   '+#13;
            sComando  := sComando +' PRE.PRE_SALDO_DISPONIVEL,                              '+#13;
            sComando  := sComando +' PRE.PRE_SALDO_RESERVADO,                               '+#13;
            sComando  := sComando +' PRE.PRE_STATUS,                                        '+#13;
            sComando  := sComando +' PRE.AUD_DEL,                                           '+#13;
            sComando  := sComando +' PRE.AUD_DATCAD,                                        '+#13;
            sComando  := sComando +' PRE.AUD_USUCAD,                                        '+#13;
            sComando  := sComando +' PRE.AUD_DATULTALT,                                     '+#13;
            sComando  := sComando +' PRE.AUD_USUULTALT,                                     '+#13;
            sComando  := sComando +' PRE.AUD_DATEXC,                                        '+#13;
            sComando  := sComando +' PRE.AUD_USUEXC                                         '+#13;
            sComando  := sComando +' FROM IND_PRODUTOS_EMPRESAS AS PRE                      '+#13;
            sComando  := sComando +' LEFT JOIN IND_PRODUTOS AS PRD                          '+#13;
            sComando  := sComando +' ON PRD.ID = PRE.PRODUTO_ID AND PRE.AUD_DEL = 0   '+#13;
            sComando  := sComando +' LEFT JOIN ADM_EMPRESAS AS EMP                          '+#13;
            sComando  := sComando +' ON EMP.ID = PRE.EMPRESA_ID AND EMP.AUD_DEL = 0         '+#13;
            sComando  := sComando +' WHERE PRE.AUD_DEL = 0 '+#13;

            if pProduto <> 0 then
               sComando  := sComando +' AND PRE.PRODUTO_ID = ' +IntToStr(pProduto);
            cdsProdutosEmpresas          := TClientDataSet.Create(nil);
            dspProdutosEmpresas.Options  := [poAllowCommandText];
            cdsProdutosEmpresas.SetProvider(dspProdutosEmpresas);
            cdsProdutosEmpresas.Close;
            cdsProdutosEmpresas.CommandText := sComando;
            cdsProdutosEmpresas.Open;

            result := cdsProdutosEmpresas.Data;

        except
            On E:Exception do
            begin
               raise E.Create('Erro ao Carregar Produtos Por Empresa  '+E.Message);
            end;
        End;
    Finally
        DestruirDataSetProdutosEmpresas;
    End;

end;

function TProdutosDAO.CarregaUnidadesMedidas: Olevariant;
begin
    Try
        Try
            CriarDataSetUnidadesMedidas;

            sComando  := ' SELECT ID,  '+#13;
            sComando  := sComando +' UNM_DESCRICAO,       '+#13;
            sComando  := sComando +' UNM_SIGLA,           '+#13;
            sComando  := sComando +' AUD_DEL,             '+#13;
            sComando  := sComando +' AUD_DATCAD,          '+#13;
            sComando  := sComando +' AUD_USUCAD,          '+#13;
            sComando  := sComando +' AUD_DATULTALT,       '+#13;
            sComando  := sComando +' AUD_USUULTALT,       '+#13;
            sComando  := sComando +' AUD_DATEXC,          '+#13;
            sComando  := sComando +' AUD_USUEXC           '+#13;
            sComando  := sComando +' FROM IND_UNIDADE_MEDIDAS '+#13;
            sComando  := sComando +' WHERE AUD_DEL = 0 '+#13;

            cdsUnidadesMedidas           := TClientDataSet.Create(nil);
            dspUnidadesMedidas.Options  := [poAllowCommandText];
            cdsUnidadesMedidas.SetProvider(dspUnidadesMedidas);
            cdsUnidadesMedidas.Close;
            cdsUnidadesMedidas.CommandText := sComando;
            cdsUnidadesMedidas.Open;


            result := cdsUnidadesMedidas.Data;

        except
            On E:Exception do
            begin
               raise E.Create('Erro ao Carregar Unidades de Medidas '+E.Message);
            end;
        End;
    Finally
        DestruirDataSetUnidadesMedidas;
    End;

end;

constructor TProdutosDAO.Create;
begin
    conexao                     := TSQLConnection.Create(nil);

    qryDados                    := TSQLQuery.Create(nil);
    dspDados                    := TDataSetProvider.Create(nil);
    cdsDados                    := TClientDataset.Create(nil);
    dsDados                     := TDataSource.Create(nil);
    sdsExcDados                 := TSQLDataSet.Create(nil);

    qryPesquisa                 := TSQLQuery.Create(nil);
    dspPesquisa                 := TDataSetProvider.Create(nil);
    cdsPesquisa                 := TClientDataset.Create(nil);
    dsPesquisa                  := TDataSource.Create(nil);
    sdsExcPesquisa              := TSQLDataSet.Create(nil);

    qryProdutos                 := TSQLQuery.Create(nil);
    dspProdutos                 := TDataSetProvider.Create(nil);
    cdsProdutos                 := TClientDataset.Create(nil);
    dsProdutos                  := TDataSource.Create(nil);
    sdsExcProdutos              := TSQLDataSet.Create(nil);

    qryProdutosEmpresas         := TSQLQuery.Create(nil);
    dspProdutosEmpresas         := TDataSetProvider.Create(nil);
    cdsProdutosEmpresas         := TClientDataset.Create(nil);
    dsProdutosEmpresas          := TDataSource.Create(nil);
    sdsExcProdutosEmpresas      := TSQLDataSet.Create(nil);

    qryEmpresas                 := TSQLQuery.Create(nil);
    dspEmpresas                 := TDataSetProvider.Create(nil);
    cdsEmpresas                 := TClientDataset.Create(nil);
    dsEmpresas                  := TDataSource.Create(nil);
    sdsExcEmpresas              := TSQLDataSet.Create(nil);

    qryNcms                     := TSQLQuery.Create(nil);
    dspNcms                     := TDataSetProvider.Create(nil);
    cdsNcms                     := TClientDataset.Create(nil);
    dsNcms                      := TDataSource.Create(nil);
    sdsExcNcms                  := TSQLDataSet.Create(nil);

    qryUnidadesMedidas          := TSQLQuery.Create(nil);
    dspUnidadesMedidas          := TDataSetProvider.Create(nil);
    cdsUnidadesMedidas          := TClientDataset.Create(nil);
    dsUnidadesMedidas           := TDataSource.Create(nil);
    sdsExcUnidadesMedidas       := TSQLDataSet.Create(nil);

    qryGruposProdutos           := TSQLQuery.Create(nil);
    dspGruposProdutos           := TDataSetProvider.Create(nil);
    cdsGruposProdutos           := TClientDataset.Create(nil);
    dsGruposProdutos            := TDataSource.Create(nil);
    sdsExcGruposProdutos        := TSQLDataSet.Create(nil);
end;



procedure TProdutosDAO.CriarDataSetDados;
begin
    With conexao do
    begin
        Connected                               := false;
        Params.Values['SchemaOverride']         :=  '%.dbo';
        Params.Values['DriverUnit']             :=  'Data.DBXMSSQL';
        Params.Values['DriverPackageLoader']    :=  'TDBXDynalinkDriverLoader,DBXCommonDriver190.bpl';
        Params.Values['DriverAssemblyLoader']   :=  'Borland.Data.TDBXDynalinkDriverLoader,Borland.Data.DbxCommonDriver,Version=19.0.0.0,Culture=neutral,PublicKeyToken=91d62ebb5b0d1b1b';
        Params.Values['MetaDataPackageLoader']  :=  'TDBXMsSqlMetaDataCommandFactory,DbxMSSQLDriver190.bpl';
        Params.Values['MetaDataAssemblyLoader'] :=  'Borland.Data.TDBXMsSqlMetaDataCommandFactory,Borland.Data.DbxMSSQLDriver,Version=19.0.0.0,Culture=neutral,PublicKeyToken=91d62ebb5b0d1b1b';
        Params.Values['GetDriverFunc']          :=  'getSQLDriverMSSQL';
        Params.Values['DriverName']             :=  'MSSQL';
        Params.Values['LibraryName']            :=  'dbxmss.dll';
        Params.Values['VendorLib']              :=  'sqlncli10.dll';
        Params.Values['VendorLibWin64']         :=  'sqlncli10.dll';
        Params.Values['HostName']               :=  'localhost';
        Params.Values['DataBase']               :=  'db_produtos';
        Params.Values['MaxBlobSize']            :=  '-1';
        Params.Values['LocaleCode']             :=  '0000';
        Params.Values['IsolationLevel']         :=  'ReadCommitted';
        Params.Values['OSAuthentication']       :=  'False';
        Params.Values['PrepareSQL']             :=  'True';
        Params.Values['User_Name']              :=  'sa';
        Params.Values['Password']               :=  'yeshua';
        Params.Values['BlobSize']               :=  '-1';
        Params.Values['Prepare SQL']            :=  'False';
        KeepConnection                          :=  false;
        LoginPrompt                             :=  false;
        ConnectionName                          :=  'MSSQLConnection';
        DriverName                              :=  'MSSQL';
        GetDriverFunc                           :=  'getSQLDriverMSSQL';
        VendorLib                               :=  'sqlncli10.dll';
        LibraryName                             :=  'dbxmss.dll';
        Connected                               :=  true;

//        KeepConnection                  := false;
//        LoginPrompt                     := false;
//        DriverName                      := 'MSSQL';
//        GetDriverFunc                   := 'getSQLDriverMSSQL';
//        VendorLib                       := 'sqlncli10.dll';
//        LibraryName                     := 'dbexpmss.dll';
//        Params.Values['ConnectionName'] := 'MSSQLConnection';
//        Params.Values['DriverName']     := DriverName;
//        Params.Values['LibraryName']    := LibraryName;
//        Params.Values['VendorLib']      := VendorLib;
//        Params.Values['GetDriverFunc']  := GetDriverFunc;
//        Params.Values['HostName']       := 'localhost';
//        Params.Values['Database']       := 'db_produtos';
//        Params.Values['User_Name']      := 'sa';
//        Params.Values['password']       := 'yeshua';
//        Connected                       := true;
    end;

end;

procedure TProdutosDAO.CriarDataSetEmpresas;
begin
    CriarDataSetDados;
    dspEmpresas.Name                          := 'dspEmpresas';
    dspEmpresas.DataSet                       := qryEmpresas;
    cdsEmpresas.SetProvider(dspEmpresas);
    dsEmpresas.DataSet                        := cdsEmpresas;
    qryEmpresas.SQLConnection                 := conexao;
    sdsExcEmpresas.SQLConnection              := conexao;
end;

procedure TProdutosDAO.CriarDataSetGruposProdutos;
begin
    CriarDataSetDados;
    dspGruposProdutos.Name                          := 'dspGruposProdutos';
    dspGruposProdutos.DataSet                       := qryGruposProdutos;
    cdsGruposProdutos.SetProvider(dspGruposProdutos);
    dsGruposProdutos.DataSet                        := cdsGruposProdutos;
    qryGruposProdutos.SQLConnection                 := conexao;
    sdsExcGruposProdutos.SQLConnection              := conexao;
end;

procedure TProdutosDAO.CriarDataSetNcms;
begin
    CriarDataSetDados;
    dspNcms.Name                          := 'dspNcms';
    dspNcms.DataSet                       := qryNcms;
    cdsNcms.SetProvider(dspNcms);
    dsNcms.DataSet                        := cdsNcms;
    qryNcms.SQLConnection                 := conexao;
    sdsExcNcms.SQLConnection              := conexao;
end;

procedure TProdutosDAO.CriarDataSetPesquisa;
begin
    CriarDataSetDados;
    dspPesquisa.Name                          := 'dspPesquisa';
    dspPesquisa.DataSet                       := qryPesquisa;
    cdsPesquisa.SetProvider(dspPesquisa);
    dsPesquisa.DataSet                        := cdsPesquisa;
    qryPesquisa.SQLConnection                 := conexao;
    sdsExcPesquisa.SQLConnection              := conexao;

end;

procedure TProdutosDAO.CriarDataSetProdutos;
begin
    CriarDataSetDados;
    dspProdutos.Name                          := 'dspProdutos';
    dspProdutos.DataSet                       := qryProdutos;
    cdsProdutos.SetProvider(dspProdutos);
    dsProdutos.DataSet                        := cdsProdutos;
    qryProdutos.SQLConnection                 := conexao;
    sdsExcProdutos.SQLConnection              := conexao;
    dspProdutosEmpresas.Name                  := 'dspProdutosEmpresas';
    dspProdutosEmpresas.DataSet               := qryProdutosEmpresas;
    cdsProdutosEmpresas.SetProvider(dspProdutosEmpresas);
    dsProdutosEmpresas.DataSet                := cdsProdutosEmpresas;
    qryProdutosEmpresas.SQLConnection         := conexao;
    sdsExcProdutosEmpresas.SQLConnection      := conexao;

end;

procedure TProdutosDAO.CriarDataSetProdutosEmpresas;
begin
    CriarDataSetDados;
    dspProdutosEmpresas.Name                          := 'dspProdutosEmpresas';
    dspProdutosEmpresas.DataSet                       := qryProdutosEmpresas;
    cdsProdutosEmpresas.SetProvider(dspProdutosEmpresas);
    dsProdutosEmpresas.DataSet                        := cdsProdutosEmpresas;
    qryProdutosEmpresas.SQLConnection                 := conexao;
    sdsExcProdutosEmpresas.SQLConnection              := conexao;

end;

procedure TProdutosDAO.CriarDataSetUnidadesMedidas;
begin
    CriarDataSetDados;
    dspUnidadesMedidas.Name                          := 'dspUnidadesMedidas';
    dspUnidadesMedidas.DataSet                       := qryUnidadesMedidas;
    cdsUnidadesMedidas.SetProvider(dspUnidadesMedidas);
    dsUnidadesMedidas.DataSet                        := cdsUnidadesMedidas;
    qryUnidadesMedidas.SQLConnection                 := conexao;
    sdsExcUnidadesMedidas.SQLConnection              := conexao;
end;

procedure TProdutosDAO.CriarDatsSets;
begin
    CriarDataSetDados;
    CriarDataSetPesquisa;
    CriarDataSetProdutos;
    CriarDataSetProdutosEmpresas;
    CriarDataSetEmpresas;
    CriarDataSetNcms;
    CriarDataSetUnidadesMedidas;
    CriarDataSetGruposProdutos;
end;

function TProdutosDAO.DecimalSQL(const valorDecimal: Double): String;
begin
    Result := SubstituirCaracter(FloatToStr(valorDecimal),',','.');
end;

destructor TProdutosDAO.Destroy;
begin
    conexao.Free;

    qryDados.Free;
    dspDados.Free;
    cdsDados.Free;
    dsDados.Free;
    sdsExcDados.Free;

    qryPesquisa.Free;
    dspPesquisa.Free;
    cdsPesquisa.Free;
    dsPesquisa.Free;
    sdsExcPesquisa.Free;

    qryProdutos.Free;
    dspProdutos.Free;
    cdsProdutos.Free;
    dsProdutos.Free;
    sdsExcProdutos.Free;

    qryProdutosEmpresas.Free;
    dspProdutosEmpresas.Free;
    cdsProdutosEmpresas.Free;
    dsProdutosEmpresas.Free;
    sdsExcProdutosEmpresas.Free;

    qryEmpresas.Free;
    dspEmpresas.Free;
    cdsEmpresas.Free;
    dsEmpresas.Free;
    sdsExcEmpresas.Free;

    qryNcms.Free;
    dspNcms.Free;
    cdsNcms.Free;
    dsNcms.Free;
    sdsExcNcms.Free;

    qryUnidadesMedidas.Free;
    dspUnidadesMedidas.Free;
    cdsUnidadesMedidas.Free;
    dsUnidadesMedidas.Free;
    sdsExcUnidadesMedidas.Free;

    qryGruposProdutos.Free;
    dspGruposProdutos.Free;
    cdsGruposProdutos.Free;
    dsGruposProdutos.Free;
    sdsExcGruposProdutos.Free;

end;

procedure TProdutosDAO.DestruirDataSetDados;
begin
    conexao.Connected         := False;
    qryDados.SQLConnection    := conexao;
    sdsExcDados.SQLConnection := conexao;
end;

procedure TProdutosDAO.DestruirDataSetEmpresas;
begin
    conexao.Connected         := False;
    qryEmpresas.SQLConnection    := conexao;
    sdsExcEmpresas.SQLConnection := conexao;
end;

procedure TProdutosDAO.DestruirDataSetGruposProdutos;
begin
    conexao.Connected         := False;
    qryGruposProdutos.SQLConnection    := conexao;
    sdsExcGruposProdutos.SQLConnection := conexao;

end;

procedure TProdutosDAO.DestruirDataSetNcms;
begin
    conexao.Connected         := False;
    qryNcms.SQLConnection    := conexao;
    sdsExcNcms.SQLConnection := conexao;

end;

procedure TProdutosDAO.DestruirDataSetPesquisa;
begin
    conexao.Connected         := False;
    qryPesquisa.SQLConnection    := conexao;
    sdsExcPesquisa.SQLConnection := conexao;

end;

procedure TProdutosDAO.DestruirDataSetProdutos;
begin
    conexao.Connected                     := False;
    qryProdutos.SQLConnection             := conexao;
    sdsExcProdutos.SQLConnection          := conexao;
    qryProdutosEmpresas.SQLConnection     := conexao;
    sdsExcProdutosEmpresas.SQLConnection  := conexao;

end;

procedure TProdutosDAO.DestruirDataSetProdutosEmpresas;
begin
    conexao.Connected         := False;
    qryProdutosEmpresas.SQLConnection    := conexao;
    sdsExcProdutosEmpresas.SQLConnection := conexao;

end;

procedure TProdutosDAO.DestruirDataSets;
begin
    DestruirDataSetDados;
    DestruirDataSetPesquisa;
    DestruirDataSetProdutos;
    DestruirDataSetProdutosEmpresas;
    DestruirDataSetEmpresas;
    DestruirDataSetNcms;
    DestruirDataSetUnidadesMedidas;
    DestruirDataSetGruposProdutos;
end;

procedure TProdutosDAO.DestruirDataSetUnidadesMedidas;
begin
    conexao.Connected         := False;
    qryUnidadesMedidas.SQLConnection    := conexao;
    sdsExcUnidadesMedidas.SQLConnection := conexao;

end;

function TProdutosDAO.Excluir(dProduto, dProdutoEmpresa: Olevariant; pProduto: Integer): Boolean;
var
  TD : TTransactionDesc;
begin
    Try
        Try
            CriarDataSetProdutos;

            Result  := False;

            cdsProdutos.Close;
            cdsProdutos.Data      := dProduto;
            cdsProdutos.Filtered  := True;
            cdsProdutos.Filter    := 'ID = '+IntToStr(pProduto);
            cdsProdutos.Open;
            if cdsProdutos.RecordCount > 0 then
            begin
                if not conexao.InTransaction then
                begin
                    TD.TransactionID    := 99999;
                    TD.IsolationLevel   := xilREADCOMMITTED;
                    conexao.StartTransaction(TD);
                end;

                sComando  := ' '+#13;
                sComando  := sComando + 'UPDATE IND_PRODUTOS '+#13;
                sComando  := sComando + ' SET AUD_DEL = 1, '+#13;
                sComando  := sComando + ' AUD_USUEXC  = ''Administrador'', '+#13;
                sComando  := sComando + ' AUD_DATEXC  = '+QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',Date))+' '+#13;
                sComando  := sComando + ' WHERE ID    = '+IntToStr(pProduto)+' '+#13;

                cdsProdutosEmpresas.Close;
                cdsProdutosEmpresas.Data      := dProdutoEmpresa;
                cdsProdutosEmpresas.Filtered  := True;
                cdsProdutosEmpresas.Filter    := 'PRODUTO_ID = '+IntToStr(pProduto);
                cdsProdutosEmpresas.Open;
                if cdsProdutosEmpresas.RecordCount > 0 then
                begin
                    sComando  := sComando + ' '+#13;
                    sComando  := sComando + 'UPDATE IND_PRODUTOS_EMPRESAS '+#13;
                    sComando  := sComando + ' SET AUD_DEL       = 1, '+#13;
                    sComando  := sComando + ' AUD_USUEXC        = ''Administrador'', '+#13;
                    sComando  := sComando + ' AUD_DATEXC        = '+QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',Date))+' '+#13;
                    sComando  := sComando + ' WHERE PRODUTO_ID  = '+IntToStr(pProduto)+' '+#13;
                end;

               sdsExcProdutos.Close;
               sdsExcProdutos.CommandText := sComando;
               sdsExcProdutos.ExecSQL;

               if conexao.InTransaction then
                 conexao.Commit(TD);

               result := True;

            end;

        except
            On E:Exception do
            begin
               if conexao.InTransaction then
                 conexao.Rollback(TD);

               raise E.Create('Erro ao Excluir Registro '+E.Message);
            end;
        End;
    Finally
        DestruirDataSetProdutos;
    End;
end;

function TProdutosDAO.Inserir(dProduto, dProdutoEmpresa: Olevariant; pProduto: Integer): Boolean;
var
  TD : TTransactionDesc;
  produtoID : Integer;
begin
    Try
        Try
            CriarDataSetProdutos;

            Result  := False;

            cdsProdutos.Close;
            cdsProdutos.Data      := dProduto;
            cdsProdutos.Filtered  := True;
            cdsProdutos.Filter    := 'ID = '+IntToStr(pProduto);
            cdsProdutos.Open;
            if cdsProdutos.RecordCount > 0 then
            begin
                if not conexao.InTransaction then
                begin
                    TD.TransactionID    := 99999;
                    TD.IsolationLevel   := xilREADCOMMITTED;
                    conexao.StartTransaction(TD);
                end;

                sComando  := ' '+#13;
                sComando  := sComando +'INSERT INTO IND_PRODUTOS '+#13;
                sComando  := sComando +'( '+#13;
                sComando  := sComando +' GRUPO_PRODUTO_ID,      '+#13;
                sComando  := sComando +' UNIDADE_MEDIDA_ID,     '+#13;
                sComando  := sComando +' NCM_ID,                '+#13;
                sComando  := sComando +' PRD_CODBARRAS,         '+#13;
                sComando  := sComando +' PRD_CODIGO,            '+#13;
                sComando  := sComando +' PRD_APELIDO,           '+#13;
                sComando  := sComando +' PRD_DESCRICAO,         '+#13;
                sComando  := sComando +' PRD_DESCRICAOFISCAL,   '+#13;
                sComando  := sComando +' PRD_DESCRICAOTECNICA,  '+#13;
                sComando  := sComando +' PRD_TIPOPRODUTO,       '+#13;
                sComando  := sComando +' PRD_PESOBRUTO,         '+#13;
                sComando  := sComando +' PRD_PESOLIQUIDO,       '+#13;
                sComando  := sComando +' PRD_FOTO,              '+#13;
                sComando  := sComando +' AUD_DEL,               '+#13;
                sComando  := sComando +' AUD_DATCAD,            '+#13;
                sComando  := sComando +' AUD_USUCAD             '+#13;
                sComando  := sComando +' ) ';
                sComando  := sComando +' VALUES ( '+#13;
                sComando  := sComando +' '+cdsProdutos.FieldByName('GRUPO_PRODUTO_ID').AsString+', '+#13;
                sComando  := sComando +' '+cdsProdutos.FieldByName('UNIDADE_MEDIDA_ID').AsString+', '+#13;
                sComando  := sComando +' '+cdsProdutos.FieldByName('NCM_ID').AsString+', '+#13;
                sComando  := sComando +' '+QuotedStr(cdsProdutos.FieldByName('PRD_CODBARRAS').AsString)+', '+#13;
                sComando  := sComando +' '+QuotedStr(cdsProdutos.FieldByName('PRD_CODIGO').AsString)+', '+#13;
                sComando  := sComando +' '+QuotedStr(cdsProdutos.FieldByName('PRD_APELIDO').AsString)+', '+#13;
                sComando  := sComando +' '+QuotedStr(cdsProdutos.FieldByName('PRD_DESCRICAO').AsString)+', '+#13;
                sComando  := sComando +' '+QuotedStr(cdsProdutos.FieldByName('PRD_DESCRICAOFISCAL').AsString)+', '+#13;
                sComando  := sComando +' '+QuotedStr(cdsProdutos.FieldByName('PRD_DESCRICAOTECNICA').AsWideString)+', '+#13;
                sComando  := sComando +' '+cdsProdutos.FieldByName('PRD_TIPOPRODUTO').AsString+', '+#13;
                sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutos.FieldByName('PRD_PESOBRUTO').AsFloat))+', '+#13;
                sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutos.FieldByName('PRD_PESOLIQUIDO').AsFloat))+', '+#13;
                sComando  := sComando +' '+QuotedStr(cdsProdutos.FieldByName('PRD_FOTO').AsWideString)+', '+#13;
                sComando  := sComando +' 0, '+#13;
                sComando  := sComando +' '+QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',Date))+', '+#13;
                sComando  := sComando +' '+QuotedStr('Administrador')+' '+#13;
                sComando  := sComando +' ) '+#13;

                sdsExcProdutos.Close;
                sdsExcProdutos.CommandText := sComando;
                sdsExcProdutos.ExecSQL;

                produtoID := RetornaID('IND_PRODUTOS');

                cdsProdutosEmpresas.Close;
                cdsProdutosEmpresas.Data      := dProdutoEmpresa;
                cdsProdutosEmpresas.Filtered  := True;
                cdsProdutosEmpresas.Filter    := 'PRODUTO_ID = '+IntToStr(pProduto);
                cdsProdutosEmpresas.Open;
                if cdsProdutosEmpresas.RecordCount > 0 then
                begin
                    cdsProdutosEmpresas.First;
                    while not cdsProdutosEmpresas.Eof do
                    begin
                        sComando  := sComando +' '+#13;
                        sComando  := sComando +'INSERT INTO IND_PRODUTOS_EMPRESAS '+#13;
                        sComando  := sComando +'( '+#13;
                        sComando  := sComando +'EMPRESA_ID,           '+#13;
                        sComando  := sComando +'PRODUTO_ID,           '+#13;
                        sComando  := sComando +'PRE_PRECOCUSTO,       '+#13;
                        sComando  := sComando +'PRE_CUSTOMEDIO,       '+#13;
                        sComando  := sComando +'PRE_CUSTOCONTABIL,    '+#13;
                        sComando  := sComando +'PRE_CUSTOGERENCIAL,   '+#13;
                        sComando  := sComando +'PRE_PRECOVENDA,       '+#13;
                        sComando  := sComando +'PRE_PERCCUSTOOPER,    '+#13;
                        sComando  := sComando +'PRE_MARGEMLUCRO,      '+#13;
                        sComando  := sComando +'PRE_DATAULTVENDA,     '+#13;
                        sComando  := sComando +'PRE_VLRULTVENDA,      '+#13;
                        sComando  := sComando +'PRE_DATAULTCOMPRA,    '+#13;
                        sComando  := sComando +'PRE_VLRULTCOMPRA,     '+#13;
                        sComando  := sComando +'PRE_SALDO_ANTERIOR,   '+#13;
                        sComando  := sComando +'PRE_SALDO_ATUAL,      '+#13;
                        sComando  := sComando +'PRE_SALDO_DISPONIVEL, '+#13;
                        sComando  := sComando +'PRE_SALDO_RESERVADO,  '+#13;
                        sComando  := sComando +'PRE_STATUS,           '+#13;
                        sComando  := sComando +'AUD_DEL,              '+#13;
                        sComando  := sComando +'AUD_DATCAD,           '+#13;
                        sComando  := sComando +'AUD_USUCAD            '+#13;
                        sComando  := sComando +' ) ';
                        sComando  := sComando +' VALUES ( '+#13;
                        sComando  := sComando +' '+cdsProdutosEmpresas.FieldByName('EMPRESA_ID').AsString+', '+#13;
                        sComando  := sComando +' '+IntToStr(produtoID)+', '+#13;
                        sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_PRECOCUSTO').AsFloat))+', '+#13;
                        sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_CUSTOMEDIO').AsFloat))+', '+#13;
                        sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_CUSTOCONTABIL').AsFloat))+', '+#13;
                        sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_CUSTOGERENCIAL').AsFloat))+', '+#13;
                        sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_PRECOVENDA').AsFloat))+', '+#13;
                        sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_PERCCUSTOOPER').AsFloat))+', '+#13;
                        sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_MARGEMLUCRO').AsFloat))+', '+#13;
                        sComando  := sComando +' '+QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',cdsProdutosEmpresas.FieldByName('PRE_DATAULTVENDA').AsDateTime))+', '+#13;
                        sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_VLRULTVENDA').AsFloat))+', '+#13;
                        sComando  := sComando +' '+QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',cdsProdutosEmpresas.FieldByName('PRE_DATAULTCOMPRA').AsDateTime))+', '+#13;
                        sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_VLRULTCOMPRA').AsFloat))+', '+#13;
                        sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_SALDO_ANTERIOR').AsFloat))+', '+#13;
                        sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_SALDO_ATUAL').AsFloat))+', '+#13;
                        sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_SALDO_DISPONIVEL').AsFloat))+', '+#13;
                        sComando  := sComando +' '+QuotedStr(DecimalSQL(cdsProdutosEmpresas.FieldByName('PRE_SALDO_RESERVADO').AsFloat))+', '+#13;
                        sComando  := sComando +' 0, '+#13;
                        sComando  := sComando +' 0, '+#13;
                        sComando  := sComando +' '+QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',Date))+', '+#13;
                        sComando  := sComando +' '+QuotedStr('Administrador')+' '+#13;
                        sComando  := sComando +' ) '+#13;

                        cdsProdutosEmpresas.Next;
                    end;

                end;

               sdsExcProdutosEmpresas.Close;
               sdsExcProdutosEmpresas.CommandText := sComando;
               sdsExcProdutosEmpresas.ExecSQL;

               if conexao.InTransaction then
                 conexao.Commit(TD);

               result := True;

            end;

        except
            On E:Exception do
            begin
               if conexao.InTransaction then
                 conexao.Rollback(TD);

               raise E.Create('Erro ao Inserir Registro '+E.Message);
            end;
        End;
    Finally
        DestruirDataSetProdutos;
    End;

end;

function TProdutosDAO.Pesquisar: Olevariant;
begin
    Try
        Try
            CriarDataSetPesquisa;

            sComando  := ' SELECT PRD.ID,  '+#13;
            sComando  := sComando +' PRD.GRUPO_PRODUTO_ID,                                  '+#13;
            sComando  := sComando +' GRP.GRP_DESCRICAO,                                     '+#13;
            sComando  := sComando +' PRD.UNIDADE_MEDIDA_ID,                                 '+#13;
            sComando  := sComando +' UNM.UNM_SIGLA,                                         '+#13;
            sComando  := sComando +' PRD.NCM_ID,                                            '+#13;
            sComando  := sComando +' NCM.NCM_DESCRICAO,                                     '+#13;
            sComando  := sComando +' PRD.PRD_CODBARRAS,                                     '+#13;
            sComando  := sComando +' PRD.PRD_CODIGO,                                        '+#13;
            sComando  := sComando +' PRD.PRD_APELIDO,                                       '+#13;
            sComando  := sComando +' PRD.PRD_DESCRICAO,                                     '+#13;
            sComando  := sComando +' PRD.PRD_DESCRICAOFISCAL,                               '+#13;
            sComando  := sComando +' PRD.PRD_DESCRICAOTECNICA,                              '+#13;
            sComando  := sComando +' PRD.PRD_TIPOPRODUTO,                                   '+#13;
            sComando  := sComando +' PRD.PRD_PESOBRUTO,                                     '+#13;
            sComando  := sComando +' PRD.PRD_PESOLIQUIDO,                                   '+#13;
            sComando  := sComando +' PRD.PRD_FOTO,                                          '+#13;
            sComando  := sComando +' PRD.AUD_DEL,                                           '+#13;
            sComando  := sComando +' PRD.AUD_DATCAD,                                        '+#13;
            sComando  := sComando +' PRD.AUD_USUCAD,                                        '+#13;
            sComando  := sComando +' PRD.AUD_DATULTALT,                                     '+#13;
            sComando  := sComando +' PRD.AUD_USUULTALT,                                     '+#13;
            sComando  := sComando +' PRD.AUD_DATEXC,                                        '+#13;
            sComando  := sComando +' PRD.AUD_USUEXC                                         '+#13;
            sComando  := sComando +' FROM IND_PRODUTOS AS PRD                               '+#13;
            sComando  := sComando +' LEFT JOIN IND_GRUPOS_PRODUTOS AS GRP                   '+#13;
            sComando  := sComando +' ON GRP.ID = PRD.GRUPO_PRODUTO_ID AND GRP.AUD_DEL = 0   '+#13;
            sComando  := sComando +' LEFT JOIN IND_UNIDADE_MEDIDAS AS UNM                  '+#13;
            sComando  := sComando +' ON UNM.ID = PRD.UNIDADE_MEDIDA_ID AND UNM.AUD_DEL = 0  '+#13;
            sComando  := sComando +' LEFT JOIN COM_NCMS AS NCM                              '+#13;
            sComando  := sComando +' ON NCM.ID = PRD.NCM_ID AND NCM.AUD_DEL = 0             '+#13;
            sComando  := sComando +' WHERE PRD.AUD_DEL = 0 '+#13;

            if sCodigo <> '' then
               sComando  := sComando +' AND PRD.PRD_CODIGO = ' +QuotedStr(sCodigo);

            if sCodigoBarra <> '' then
               sComando  := sComando +' AND PRD.PRD_CODBARRAS = ' +QuotedStr(sCodigoBarra);

            if sDescricao <> '' then
               sComando  := sComando +' AND PRD.PRD_DESCRICAO LIKE ' +CONCAT('%',QuotedStr(sDescricao),'%');

            if iGrupoProduto <> 0 then
               sComando  := sComando +' AND PRD.GRUPO_PRODUTO_ID = ' +IntToStr(iGrupoProduto);

            if iNcm <> 0 then
               sComando  := sComando +' AND PRD.NCM_ID = ' +IntToStr(iNcm);

            if iUnidade <> 0 then
               sComando  := sComando +' AND PRD.UNIDADE_MEDIDA_ID = ' +IntToStr(iUnidade);

            if iTipoProduto <> 0 then
               sComando  := sComando +' AND PRD.TIPOPRODUTO = ' +IntToStr(iTipoProduto);

            cdsPesquisa             := TClientDataSet.Create(nil);
            dspPesquisa.Options     := [poAllowCommandText];
            cdsPesquisa.SetProvider(dspPesquisa);
            cdsPesquisa.CommandText := sComando;
            cdsPesquisa.Open;

            result := cdsPesquisa.Data;

        except
            On E:Exception do
            begin
               raise E.Create('Erro ao Pesquisar Produtos  '+E.Message);
            end;
        End;
    Finally
        DestruirDataSetPesquisa;
    End;

end;

function TProdutosDAO.RetornaID(tabela: String): Integer;
begin

    try
        Result  := 0;
        try
            CriarDataSetDados;

            qryDados                    := TSQLQuery.Create(nil);
            dspDados                    := TDataSetProvider.Create(nil);
            cdsDados                    := TClientDataset.Create(nil);
            dsDados                     := TDataSource.Create(nil);
            sdsExcDados                 := TSQLDataSet.Create(nil);

            dspDados.Name                             := 'dspDados';
            dspDados.DataSet                          := qryDados;
            cdsDados.SetProvider(dspDados);
            dsDados.DataSet                           := cdsDados;
            qryDados.SQLConnection                    := conexao;
            sdsExcDados.SQLConnection                 := conexao;

            cdsDados.Close;
            qryDados.SQL.Clear;
            qryDados.SQL.Add('SELECT IDENT_CURRENT');
            qryDados.SQL.Add('('+QuotedStr(tabela)+') AS ID');
            cdsDados.Open;

        except
            on E:Exception do
            begin
                raise E.Create('Erro ao Recuperar Ultimo ID - '+E.Message);
            end;
        end;

        Result := cdsDados.FieldByName('ID').AsInteger;

    finally
        DestruirDataSetDados;
    end;


end;

procedure TProdutosDAO.SetcdsDados(const Value: TClientDataSet);
begin
  FcdsDados := Value;
end;

procedure TProdutosDAO.SetcdsEmpresas(const Value: TClientDataSet);
begin
  FcdsEmpresas := Value;
end;

procedure TProdutosDAO.SetcdsGruposProdutos(const Value: TClientDataSet);
begin
  FcdsGruposProdutos := Value;
end;

procedure TProdutosDAO.SetcdsNcms(const Value: TClientDataSet);
begin
  FcdsNcms := Value;
end;

procedure TProdutosDAO.SetcdsPesquisa(const Value: TClientDataSet);
begin
  FcdsPesquisa := Value;
end;

procedure TProdutosDAO.SetcdsProdutos(const Value: TClientDataSet);
begin
  FcdsProdutos := Value;
end;

procedure TProdutosDAO.SetcdsProdutosEmpresas(const Value: TClientDataSet);
begin
  FcdsProdutosEmpresas := Value;
end;

procedure TProdutosDAO.SetcdsUnidadesMedidas(const Value: TClientDataSet);
begin
  FcdsUnidadesMedidas := Value;
end;

procedure TProdutosDAO.Setconexao(const Value: TSQLConnection);
begin
  Fconexao := Value;
end;

procedure TProdutosDAO.SetdsDados(const Value: TDataSource);
begin
  FdsDados := Value;
end;

procedure TProdutosDAO.SetdsEmpresas(const Value: TDataSource);
begin
  FdsEmpresas := Value;
end;

procedure TProdutosDAO.SetdsGruposProdutos(const Value: TDataSource);
begin
  FdsGruposProdutos := Value;
end;

procedure TProdutosDAO.SetdsNcms(const Value: TDataSource);
begin
  FdsNcms := Value;
end;

procedure TProdutosDAO.SetdspDados(const Value: TDataSetProvider);
begin
  FdspDados := Value;
end;

procedure TProdutosDAO.SetdspEmpresas(const Value: TDataSetProvider);
begin
  FdspEmpresas := Value;
end;

procedure TProdutosDAO.SetdsPesquisa(const Value: TDataSource);
begin
  FdsPesquisa := Value;
end;

procedure TProdutosDAO.SetdspGruposProdutos(const Value: TDataSetProvider);
begin
  FdspGruposProdutos := Value;
end;

procedure TProdutosDAO.SetdspNcms(const Value: TDataSetProvider);
begin
  FdspNcms := Value;
end;

procedure TProdutosDAO.SetdspPesquisa(const Value: TDataSetProvider);
begin
  FdspPesquisa := Value;
end;

procedure TProdutosDAO.SetdspProdutos(const Value: TDataSetProvider);
begin
  FdspProdutos := Value;
end;

procedure TProdutosDAO.SetdspProdutosEmpresas(const Value: TDataSetProvider);
begin
  FdspProdutosEmpresas := Value;
end;

procedure TProdutosDAO.SetdsProdutos(const Value: TDataSource);
begin
  FdsProdutos := Value;
end;

procedure TProdutosDAO.SetdsProdutosEmpresas(const Value: TDataSource);
begin
  FdsProdutosEmpresas := Value;
end;

procedure TProdutosDAO.SetdspUnidadesMedidas(const Value: TDataSetProvider);
begin
  FdspUnidadesMedidas := Value;
end;

procedure TProdutosDAO.SetdsUnidadesMedidas(const Value: TDataSource);
begin
  FdsUnidadesMedidas := Value;
end;

procedure TProdutosDAO.SetqryDados(const Value: TSQLQuery);
begin
  FqryDados := Value;
end;

procedure TProdutosDAO.SetqryEmpresas(const Value: TSQLQuery);
begin
  FqryEmpresas := Value;
end;

procedure TProdutosDAO.SetqryGruposProdutos(const Value: TSQLQuery);
begin
  FqryGruposProdutos := Value;
end;

procedure TProdutosDAO.SetqryNcms(const Value: TSQLQuery);
begin
  FqryNcms := Value;
end;

procedure TProdutosDAO.SetqryPesquisa(const Value: TSQLQuery);
begin
  FqryPesquisa := Value;
end;

procedure TProdutosDAO.SetqryProdutos(const Value: TSQLQuery);
begin
  FqryProdutos := Value;
end;

procedure TProdutosDAO.SetqryProdutosEmpresas(const Value: TSQLQuery);
begin
  FqryProdutosEmpresas := Value;
end;

procedure TProdutosDAO.SetqryUnidadesMedidas(const Value: TSQLQuery);
begin
  FqryUnidadesMedidas := Value;
end;

procedure TProdutosDAO.SetsdsExcDados(const Value: TSQLDataSet);
begin
  FsdsExcDados := Value;
end;

procedure TProdutosDAO.SetsdsExcEmpresas(const Value: TSQLDataSet);
begin
  FsdsExcEmpresas := Value;
end;

procedure TProdutosDAO.SetsdsExcGruposProdutos(const Value: TSQLDataSet);
begin
  FsdsExcGruposProdutos := Value;
end;

procedure TProdutosDAO.SetsdsExcNcms(const Value: TSQLDataSet);
begin
  FsdsExcNcms := Value;
end;

procedure TProdutosDAO.SetsdsExcPesquisa(const Value: TSQLDataSet);
begin
  FsdsExcPesquisa := Value;
end;

procedure TProdutosDAO.SetsdsExcProdutos(const Value: TSQLDataSet);
begin
  FsdsExcProdutos := Value;
end;

procedure TProdutosDAO.SetsdsExcProdutosEmpresas(const Value: TSQLDataSet);
begin
  FsdsExcProdutosEmpresas := Value;
end;

procedure TProdutosDAO.SetsdsExcUnidadesMedidas(const Value: TSQLDataSet);
begin
  FsdsExcUnidadesMedidas := Value;
end;

function TProdutosDAO.SubstituirCaracter(Str: String; const Carac: String; subst: Char): String;
var
  x : Integer;
begin
    Result := '';

    for x := 1 to Length(str) do
      if Pos(str[x],Carac) > 0 then
      begin
          if subst <> #255 then
             Result := Result + Subst;
      end
      else Result := Result + str[x];
end;

end.
