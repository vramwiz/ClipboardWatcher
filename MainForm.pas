unit MainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls,ClipboardWatcher,Vcl.Imaging.pngimage;

type
  TFormMain = class(TForm)
    Panel1: TPanel;
    Memo1: TMemo;
    Panel2: TPanel;
    Image1: TImage;
    Panel3: TPanel;
    Image2: TImage;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private 宣言 }
    FWatcher : TClipboardWatcher;
    procedure DrawPngWithCheckerBoard(Canvas: TCanvas; Png: TPngImage; const TargetRect: TRect);
    procedure OnChange(Sender: TObject; const DataTypes: TClipboardWatcherDataTypes);
  public
    { Public 宣言 }
  end;

var
  FormMain: TFormMain;

implementation

{$R *.dfm}

procedure TFormMain.FormCreate(Sender: TObject);
begin
  FWatcher := TClipboardWatcher.Create;
  FWatcher.OnChange := OnChange;
  FWatcher.Enabled := True;
end;

procedure TFormMain.FormDestroy(Sender: TObject);
begin
  FWatcher.Free;
end;

procedure TFormMain.OnChange(Sender: TObject;
  const DataTypes: TClipboardWatcherDataTypes);
begin
  if cdtText in DataTypes then Memo1.Lines.Text := FWatcher.Text;
  if cdtBitmap in DataTypes then Image1.Canvas.Draw(0,0,FWatcher.Bitmap);
  if cdtPng in DataTypes then DrawPngWithCheckerBoard(Image2.Canvas,FWatcher.Png,Image2.ClientRect);

end;

procedure TFormMain.DrawPngWithCheckerBoard(Canvas: TCanvas; Png: TPngImage;
  const TargetRect: TRect);
const
  CheckSize = 8;
var
  x, y: Integer;
  Color1, Color2: TColor;
begin
  // 市松模様の色（明るいグレー／暗いグレー）
  Color1 := $CCCCCC;
  Color2 := $999999;

  // 背景にチェッカーボードを描画
  for y := TargetRect.Top to TargetRect.Bottom - 1 do
    for x := TargetRect.Left to TargetRect.Right - 1 do
    begin
      if (((x div CheckSize) + (y div CheckSize)) mod 2 = 0) then
        Canvas.Pixels[x, y] := Color1
      else
        Canvas.Pixels[x, y] := Color2;
    end;

  // PNG を上に重ねて描画（アルファチャンネル付き）
  Png.Draw(Canvas, TargetRect);
end;



end.
