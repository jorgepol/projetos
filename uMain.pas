unit uMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Menus, uFrmProdutos;

type
  TFrmMain = class(TForm)
    MainMenu1: TMainMenu;
    mnCadastros: TMenuItem;
    Produtos1: TMenuItem;
    procedure Produtos1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmMain: TFrmMain;
  FrmProdutos : TFrmProdutos;

implementation

{$R *.dfm}

procedure TFrmMain.Produtos1Click(Sender: TObject);
begin
    FrmProdutos := TFrmProdutos.Create(Self);
    try
      FrmProdutos.WindowState := wsMaximized;
      FrmProdutos.ShowModal;
    finally
      FreeAndNil(FrmProdutos);
    end;

end;

end.
