program Produtos;

uses
  Vcl.Forms,
  uMain in 'uMain.pas' {FrmMain},
  uFrmProdutos in 'uFrmProdutos.pas' {FrmProdutos},
  uClsProdutosDAO in 'uClsProdutosDAO.pas',
  uClsEmpresasDAO in 'uClsEmpresasDAO.pas',
  uClsNcmsDAO in 'uClsNcmsDAO.pas',
  uClsUnidadesMedidasDAO in 'uClsUnidadesMedidasDAO.pas',
  uClsGruposProdutosDAO in 'uClsGruposProdutosDAO.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFrmMain, FrmMain);
  Application.Run;
end.
