unit ClipboardWatcher;

{
  Unit Name   : ClipboardWatcher
  Description : クリップボードの変更を監視し、テキスト・ビットマップ・PNG 形式の取得と通知を行うユニットです。
                WM_CLIPBOARDUPDATE メッセージを使用し、ポーリング不要で効率よく監視します。
                OnChange イベントにより、取得成功したデータ形式の一覧が通知されます。

  Features    :
    - クリップボードの変化を自動検出（WM_CLIPBOARDUPDATE 使用）
    - テキスト・BMP・PNG 形式のデータを自動取得
    - PNG 形式は RegisterClipboardFormat により独自に対応
    - 複数イベントを連続で受けた際のデバウンス処理あり
    - データ取得後、OnChange イベントで通知

  Usage       :
    - TClipboardWatcher.Create でインスタンス生成
    - Enabled := True で監視を開始、False で停止
    - OnChange イベントを設定し、データ取得後の処理を記述
    - Text / Bitmap / Png プロパティから取得内容にアクセス可能

  Dependencies:
    - Windows, VCL.Forms, ActiveX, VCL.Clipbrd, pngimage

  Notes       :
    - GetClipboardText/GetClipboardBitmap/GetClipboardPng の順にデータ取得を試行
    - PNG は Windows 標準ではないため、CF_PNG フォーマットを手動登録
    - OpenClipboard には排他制御のため SafeOpenClipboard を使用
    - 外部アプリが大量に書き込む場合、同一内容の連続イベントを除去可能

  Author      : vramwiz
  Created     : 2025-07-10
  Updated     : 2025-07-10
}

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,System.DateUtils,Vcl.Imaging.pngimage,ExtCtrls;


const
  WM_CLIPBOARDUPDATE = $031D;
  CLIPBOARD_DEBOUNCE_MS = 500;

type
  TClipboardWatcherDataType = (cdtNone, cdtText, cdtBitmap, cdtPng);
  TClipboardWatcherDataTypes = set of TClipboardWatcherDataType;

  TClipboardWatcherChangeEvent = procedure(Sender: TObject; const DataTypes: TClipboardWatcherDataTypes) of object;

  // クリップボード監視クラス
type
  TClipboardWatcher = class(TPersistent)
  private
    FWindowHandle  : HWND;
    FTimerAgent    : TTimer;                      // 処理タイマー
    FTimerTimeout  : TTimer;                      // データの終わりを待つタイマー
    FLastTime      : TDateTime;                   // 最後に取得した時間 ※同一データ取得防止
    FEnabled       : Boolean;                     // True:変更でイベントを発生させる
    FDataTypes     : TClipboardWatcherDataTypes;  // 取得出来たデータの種類
    FPng           : TPngImage;
    FBitmap        : TBitmap;
    FText          : string;
    FOnChange      : TClipboardWatcherChangeEvent;

    function GetClipboard() : Boolean;
    function GetClipboardBitmap() : Boolean;
    function GetClipboardPng() : Boolean;
    function GetClipboardText() : Boolean;

    function SafeOpenClipboard(hWnd: HWND = 0; Retry: Integer = 10; DelayMS: Integer = 50): Boolean;

    procedure WndProc(var Msg: TMessage);
    procedure ClipboardExecute();
    procedure OnTimerAgent(Sender: TObject);
    procedure OnTimerTimeOut(Sender: TObject);
  protected
    procedure DoChange(const DataTypes: TClipboardWatcherDataTypes);
  public
    constructor Create;
    destructor Destroy; override;
    // True : クリップボード監視 False：監視を中断
    property Enabled : Boolean read FEnabled write FEnabled;
    property Text    : string    read FText;
    property Bitmap  : TBitmap   read FBitmap;
    property Png     : TPngImage read FPng;
    // クリップボード変更イベント
    property OnChange : TClipboardWatcherChangeEvent read FOnChange write FOnChange;

  end;

// 指定されたビットマップをクリップボードにコピー
procedure SetClipboardBitmap(Bmp: TBitmap);


implementation

uses
  Vcl.Clipbrd,Winapi.ActiveX;

    // PNG用クリップボード形式登録
var
  CF_PNG: UINT = 0;
  ClipboardUpdateIsSelf : Boolean;

procedure SetClipboardBitmap(Bmp: TBitmap);
const
  CF_PNG_NAME = 'PNG';
var
  CF_PNG: UINT;
  tmp: TBitmap;
  hBmp: HBITMAP;
  hDIB, hPNG: HGLOBAL;
  rowSize, dibSize: Integer;
  png: TPngImage;
  stream: TMemoryStream;
  pDIB, pBits, srcLine, dest: PByte;
  ptr: Pointer;
  x, y: Integer;
  bih: BITMAPINFOHEADER;
begin
  if not Assigned(bmp) then Exit;

  CF_PNG := RegisterClipboardFormat(CF_PNG_NAME);

  ClipboardUpdateIsSelf := True;

  // --- スキャンライン安全のため pf32bit でコピー ---
  tmp := TBitmap.Create;
  try
    tmp.PixelFormat := pf32bit;
    tmp.Width := bmp.Width;
    tmp.Height := bmp.Height;
    tmp.Canvas.Draw(0, 0, bmp);  // 描画後 ScanLine 使用可能

    // --- CF_DIB (24bit RGB) 構築 ---
    rowSize := ((tmp.Width * 3 + 3) div 4) * 4;
    dibSize := SizeOf(BITMAPINFOHEADER) + rowSize * tmp.Height;
    hDIB := GlobalAlloc(GMEM_MOVEABLE, dibSize);

    if hDIB <> 0 then
    begin
      pDIB := GlobalLock(hDIB);
      if Assigned(pDIB) then
      begin
        FillChar(bih, SizeOf(bih), 0);
        bih.biSize := SizeOf(bih);
        bih.biWidth := tmp.Width;
        bih.biHeight := tmp.Height;
        bih.biPlanes := 1;
        bih.biBitCount := 24;
        bih.biCompression := BI_RGB;

        Move(bih, pDIB^, SizeOf(bih));
        pBits := pDIB + SizeOf(bih);

        for y := tmp.Height - 1 downto 0 do
        begin
          srcLine := tmp.ScanLine[y];
          dest := pBits;
          for x := 0 to tmp.Width - 1 do
          begin
            dest^ := srcLine^;      Inc(dest); Inc(srcLine); // B
            dest^ := srcLine^;      Inc(dest); Inc(srcLine); // G
            dest^ := srcLine^;      Inc(dest); Inc(srcLine); // R
            Inc(srcLine); // Skip A
          end;
          Inc(pBits, rowSize);
        end;

        GlobalUnlock(hDIB);
      end;
    end;

    // --- CF_PNG 構築 ---
    stream := TMemoryStream.Create;
    png := TPngImage.Create;
    try
      png.Assign(tmp);
      png.SaveToStream(stream);
      stream.Position := 0;

      hPNG := GlobalAlloc(GMEM_MOVEABLE, stream.Size);
      if hPNG <> 0 then
      begin
        ptr := GlobalLock(hPNG);
        if Assigned(ptr) then
        begin
          Move(stream.Memory^, ptr^, stream.Size);
          GlobalUnlock(hPNG);
        end;
      end;
    finally
      png.Free;
      stream.Free;
    end;

    // --- 最後に CF_BITMAP を作成（破壊されないように最後！） ---
    hBmp := CopyImage(tmp.Handle, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG or LR_COPYDELETEORG);

    // --- クリップボードに登録 ---
    if OpenClipboard(0) then
    try
      EmptyClipboard;
      if hBmp <> 0 then SetClipboardData(CF_BITMAP, hBmp);
      if hDIB <> 0 then SetClipboardData(CF_DIB, hDIB);
      if hPNG <> 0 then SetClipboardData(CF_PNG, hPNG);
    finally
      CloseClipboard;
    end;

  finally
    tmp.Free;
  end;
end;



constructor TClipboardWatcher.Create;
begin
  inherited;
  FLastTime := 0;
  FWindowHandle := AllocateHWnd(WndProc);
  AddClipboardFormatListener(FWindowHandle);

  FBitmap := TBitmap.Create;
  FPng    := TPngImage.Create;

  FTimerAgent := TTimer.Create(nil);
  FTimerAgent.Enabled := False;
  FTimerAgent.Interval := 10;
  FTimerAgent.OnTimer := OnTimerAgent;

  FTimerTimeout := TTimer.Create(nil);
  FTimerTimeout.Enabled := False;
  FTimerTimeout.Interval := 30;
  FTimerTimeout.OnTimer := OnTimerTimeOut;

end;

destructor TClipboardWatcher.Destroy;
begin
  FTimerTimeout.Free;
  FTimerAgent.Free;
  FPng.Free;
  FBitmap.Free;
  RemoveClipboardFormatListener(FWindowHandle);
  DeallocateHWnd(FWindowHandle);
  inherited;
end;

procedure TClipboardWatcher.DoChange(
  const DataTypes: TClipboardWatcherDataTypes);
begin
  if Assigned(FOnChange) then FOnChange(Self,DataTypes);
end;

function TClipboardWatcher.GetClipboard() : Boolean;
begin
  result := False;
  if IsClipboardFormatAvailable(CF_BITMAP) then begin
    if not (cdtBitmap in FDataTypes) then begin
      if GetClipboardBitmap() then begin
        FDataTypes := FDataTypes + [cdtBitmap];
        result := True;
      end;
    end;
  end;
  if IsClipboardFormatAvailable(CF_PNG) then begin
    if not (cdtPng in FDataTypes) then begin
      if GetClipboardPng() then begin
        FDataTypes := FDataTypes + [cdtPng];
        result := True;
      end;
    end;
  end;
  if IsClipboardFormatAvailable(CF_UNICODETEXT) then begin
    if not (cdtText in FDataTypes) then begin
      if GetClipboardText() then begin
        FDataTypes := FDataTypes + [cdtText];
        result := True;
      end;
    end;
  end;
end;

function TClipboardWatcher.GetClipboardBitmap: Boolean;
var
  hBmp: HBITMAP;
begin
  result := False;

  if not IsClipboardFormatAvailable(CF_BITMAP) then Exit;

  hBmp := GetClipboardData(CF_BITMAP);
  if hBmp = 0 then Exit;

  hBmp := CopyImage(hBmp, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG or LR_COPYDELETEORG);
  if hBmp = 0 then Exit;

  FBitmap.Handle := 0;
  FBitmap.Handle := hBmp;
  result := True;
end;

function TClipboardWatcher.GetClipboardPng: Boolean;
var
  hData: THandle;
  pData: Pointer;
  Size: Integer;
  Stream: TMemoryStream;
begin
  Result := False;

  // WindowsでPNG形式のクリップボードフォーマットを取得（レジスタされていない場合は0）
  if CF_PNG = 0 then
    Exit;

  if not IsClipboardFormatAvailable(CF_PNG) then Exit;

  hData := GetClipboardData(CF_PNG);
  if hData = 0 then Exit;

  pData := GlobalLock(hData);
  if not Assigned(pData) then Exit;

  try
    Size := GlobalSize(hData);
    if Size = 0 then Exit;

    Stream := TMemoryStream.Create;
    try
      Stream.WriteBuffer(pData^, Size);
      Stream.Position := 0;
      FPng.LoadFromStream(Stream);
      Result := True;
    finally
      Stream.Free;
    end;
  finally
    GlobalUnlock(hData);
  end;
end;

function TClipboardWatcher.GetClipboardText: Boolean;
var
  hData: THandle;
  pData: PChar;
begin
  Result := False;
  FText := '';

  if not IsClipboardFormatAvailable(CF_UNICODETEXT) then Exit;

  hData := GetClipboardData(CF_UNICODETEXT);
  if hData = 0 then Exit;

  pData := GlobalLock(hData);
  if not Assigned(pData) then Exit;

  try
    FText := pData;
    Result := True;
  finally
    GlobalUnlock(hData);
  end;
end;

procedure TClipboardWatcher.OnTimerAgent(Sender: TObject);
begin
  FTimerAgent.Enabled := False;
  if SafeOpenClipboard(0) then
  begin
    try
      if GetClipboard() then
      begin
        FTimerTimeout.Enabled := False;
        FTimerTimeout.Enabled := True;
      end;
    finally
      CloseClipboard;
    end;
  end;
  FTimerAgent.Enabled := True;
end;

procedure TClipboardWatcher.OnTimerTimeOut(Sender: TObject);
begin
  FTimerAgent.Enabled := False;
  FTimerTimeout.Enabled := False;
  try
    DoChange(FDataTypes);
  finally
    FDataTypes := [];
  end;
end;

function TClipboardWatcher.SafeOpenClipboard(hWnd: HWND; Retry,
  DelayMS: Integer): Boolean;
var
  i: Integer;
begin
  for i := 1 to Retry do
  begin
    if OpenClipboard(hWnd) then
      Exit(True);
    Sleep(DelayMS);
  end;
  Result := False;
end;

procedure TClipboardWatcher.ClipboardExecute;
begin
  if MilliSecondsBetween(Now, FLastTime) >= CLIPBOARD_DEBOUNCE_MS then begin
    FLastTime := Now;
    if not ClipboardUpdateIsSelf then
    if FEnabled then begin
      FTimerAgent.Enabled := False;
      FTimerAgent.Enabled := True;
      FTimerTimeout.Enabled := False;
      FTimerTimeout.Enabled := True;
    end;
    //if FEnabled then DoClipboardChanged();
    ClipboardUpdateIsSelf := False;
  end;
end;


procedure TClipboardWatcher.WndProc(var Msg: TMessage);
begin
  if Msg.Msg = WM_CLIPBOARDUPDATE then begin
    ClipboardExecute();
  end;

  Msg.Result := DefWindowProc(FWindowHandle, Msg.Msg, Msg.WParam, Msg.LParam);
end;

initialization
  CF_PNG := RegisterClipboardFormat('PNG');

end.
