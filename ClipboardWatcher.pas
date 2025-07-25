unit ClipboardWatcher;

{
  Unit Name   : ClipboardWatcher
  Description : �N���b�v�{�[�h�̕ύX���Ď����A�e�L�X�g�E�r�b�g�}�b�v�EPNG �`���̎擾�ƒʒm���s�����j�b�g�ł��B
                WM_CLIPBOARDUPDATE ���b�Z�[�W���g�p���A�|�[�����O�s�v�Ō����悭�Ď����܂��B
                OnChange �C�x���g�ɂ��A�擾���������f�[�^�`���̈ꗗ���ʒm����܂��B

  Features    :
    - �N���b�v�{�[�h�̕ω����������o�iWM_CLIPBOARDUPDATE �g�p�j
    - �e�L�X�g�EBMP�EPNG �`���̃f�[�^�������擾
    - PNG �`���� RegisterClipboardFormat �ɂ��Ǝ��ɑΉ�
    - �����C�x���g��A���Ŏ󂯂��ۂ̃f�o�E���X��������
    - �f�[�^�擾��AOnChange �C�x���g�Œʒm

  Usage       :
    - TClipboardWatcher.Create �ŃC���X�^���X����
    - Enabled := True �ŊĎ����J�n�AFalse �Œ�~
    - OnChange �C�x���g��ݒ肵�A�f�[�^�擾��̏������L�q
    - Text / Bitmap / Png �v���p�e�B����擾���e�ɃA�N�Z�X�\

  Dependencies:
    - Windows, VCL.Forms, ActiveX, VCL.Clipbrd, pngimage

  Notes       :
    - GetClipboardText/GetClipboardBitmap/GetClipboardPng �̏��Ƀf�[�^�擾�����s
    - PNG �� Windows �W���ł͂Ȃ����߁ACF_PNG �t�H�[�}�b�g���蓮�o�^
    - OpenClipboard �ɂ͔r������̂��� SafeOpenClipboard ���g�p
    - �O���A�v������ʂɏ������ޏꍇ�A������e�̘A���C�x���g�������\

  �N���b�v�{�[�h�̓��e�𒼐ڎ擾�E�ݒ肷��O���[�o���֐����p�ӂ���Ă��܂��B
  `TClipboardWatcher` ���g�p�����ɁA�C�ӂ̃^�C�~���O�ŗ��p�\�ł��B

  Author      : vramwiz
  Created     : 2025-07-10
  Updated     : 2025-07-25
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

  // �N���b�v�{�[�h�Ď��N���X
type
  TClipboardWatcher = class(TPersistent)
  private
    FWindowHandle  : HWND;
    FTimerAgent    : TTimer;                      // �����^�C�}�[
    FTimerTimeout  : TTimer;                      // �f�[�^�̏I����҂^�C�}�[
    FLastTime      : TDateTime;                   // �Ō�Ɏ擾�������� ������f�[�^�擾�h�~
    FEnabled       : Boolean;                     // True:�ύX�ŃC�x���g�𔭐�������
    FDataTypes     : TClipboardWatcherDataTypes;  // �擾�o�����f�[�^�̎��
    FPng           : TPngImage;
    FBitmap        : TBitmap;
    FText          : string;
    FOnChange      : TClipboardWatcherChangeEvent;

    function GetClipboard() : Boolean;

    procedure WndProc(var Msg: TMessage);
    procedure ClipboardExecute();
    procedure OnTimerAgent(Sender: TObject);
    procedure OnTimerTimeOut(Sender: TObject);
    function GetDelayMS: Integer;
    procedure SetDelayMS(const Value: Integer);
  protected
    procedure DoChange(const DataTypes: TClipboardWatcherDataTypes);
  public
    constructor Create;
    destructor Destroy; override;
    // True : �N���b�v�{�[�h�Ď� False�F�Ď��𒆒f
    property Enabled : Boolean read FEnabled write FEnabled;
    property Text    : string    read FText;
    property Bitmap  : TBitmap   read FBitmap;
    property Png     : TPngImage read FPng;
    // �Ō�̃f�[�^����C�x���g�����܂őҋ@���鎞��
    property DelayMS : Integer   read GetDelayMS write SetDelayMS;
    // �N���b�v�{�[�h�ύX�C�x���g
    property OnChange : TClipboardWatcherChangeEvent read FOnChange write FOnChange;

  end;

// �w�肳�ꂽ�e�L�X�g���N���b�v�{�[�h�ɃR�s�[
procedure SetClipboardText(const Text: string);
// �w�肳�ꂽ�r�b�g�}�b�v���N���b�v�{�[�h�ɃR�s�[
procedure SetClipboardBitmap(Bmp: TBitmap);
// �w�肳�ꂽPNG���N���b�v�{�[�h�ɃR�s�[
procedure SetClipboardPng(Png: TPngImage);
// �N���b�v�{�[�h����r�b�g�}�b�v�擾
function GetClipboardBitmap(Bitmap : TBitmap) : Boolean;
// �N���b�v�{�[�h����Png�摜�擾
function GetClipboardPng(Png : TPngImage) : Boolean;
// �N���b�v�{�[�h����e�L�X�g�擾
function GetClipboardText(var Text : string) : Boolean;


implementation

uses
  Vcl.Clipbrd,Winapi.ActiveX;

    // PNG�p�N���b�v�{�[�h�`���o�^
var
  CF_PNG: UINT = 0;
  ClipboardUpdateIsSelf : Boolean;

//  �N���b�v�{�[�h�����S�ɊJ��
function SafeOpenClipboard(hWnd: HWND = 0; Retry: Integer = 10; DelayMS: Integer = 50): Boolean;
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
  FTimerTimeout.Interval := 100;
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

function GetClipboardBitmapSub(Bitmap  : TBitmap): Boolean;
var
  hBmp: HBITMAP;
begin
  result := False;

  if not IsClipboardFormatAvailable(CF_BITMAP) then Exit;

  hBmp := GetClipboardData(CF_BITMAP);
  if hBmp = 0 then Exit;

  hBmp := CopyImage(hBmp, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG or LR_COPYDELETEORG);
  if hBmp = 0 then Exit;

  Bitmap.Handle := 0;
  Bitmap.Handle := hBmp;
  result := True;
end;

function GetClipboardPngSub(Png     : TPngImage): Boolean;
var
  hData: THandle;
  pData: Pointer;
  Size: Integer;
  Stream: TMemoryStream;
begin
  Result := False;

  // Windows��PNG�`���̃N���b�v�{�[�h�t�H�[�}�b�g���擾�i���W�X�^����Ă��Ȃ��ꍇ��0�j
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
      Png.LoadFromStream(Stream);
      Result := True;
    finally
      Stream.Free;
    end;
  finally
    GlobalUnlock(hData);
  end;
end;

function GetClipboardTextSub(var Text    : string): Boolean;
var
  hData: THandle;
  pData: PChar;
begin
  Result := False;
  Text := '';

  if not IsClipboardFormatAvailable(CF_UNICODETEXT) then Exit;

  hData := GetClipboardData(CF_UNICODETEXT);
  if hData = 0 then Exit;

  pData := GlobalLock(hData);
  if not Assigned(pData) then Exit;

  try
    Text := pData;
    Result := True;
  finally
    GlobalUnlock(hData);
  end;
end;

function TClipboardWatcher.GetClipboard() : Boolean;
begin
  result := False;
  if IsClipboardFormatAvailable(CF_BITMAP) then begin
    if not (cdtBitmap in FDataTypes) then begin
      if GetClipboardBitmapSub(FBitmap) then begin
        FDataTypes := FDataTypes + [cdtBitmap];
        result := True;
      end;
    end;
  end;
  if IsClipboardFormatAvailable(CF_PNG) then begin
    if not (cdtPng in FDataTypes) then begin
      if GetClipboardPngSub(FPng) then begin
        FDataTypes := FDataTypes + [cdtPng];
        result := True;
      end;
    end;
  end;
  if IsClipboardFormatAvailable(CF_UNICODETEXT) then begin
    if not (cdtText in FDataTypes) then begin
      if GetClipboardTextSub(FText) then begin
        FDataTypes := FDataTypes + [cdtText];
        result := True;
      end;
    end;
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

function TClipboardWatcher.GetDelayMS: Integer;
begin
  result := FTimerTimeout.Interval;
end;

procedure TClipboardWatcher.SetDelayMS(const Value: Integer);
begin
  FTimerTimeout.Interval := Value;
end;


procedure TClipboardWatcher.WndProc(var Msg: TMessage);
begin
  if Msg.Msg = WM_CLIPBOARDUPDATE then begin
    ClipboardExecute();
  end;

  Msg.Result := DefWindowProc(FWindowHandle, Msg.Msg, Msg.WParam, Msg.LParam);
end;


function BuildClipboardDIB(bmp: TBitmap): HGLOBAL;
var
  rowSize, dibSize: Integer;
  tmp: TBitmap;
  hDIB: HGLOBAL;
  pDIB, pBits, srcLine, dest: PByte;
  x, y: Integer;
  bih: BITMAPINFOHEADER;
begin
  Result := 0;

  if not Assigned(bmp) then Exit;

  // pf32bit �ɕϊ����Ĉ��S�ɃX�L�������C������
  tmp := TBitmap.Create;
  try
    tmp.PixelFormat := pf32bit;
    tmp.Width := bmp.Width;
    tmp.Height := bmp.Height;
    tmp.Canvas.Draw(0, 0, bmp);  // �`�悵�ĕϊ�

    rowSize := ((tmp.Width * 3 + 3) div 4) * 4;
    dibSize := SizeOf(BITMAPINFOHEADER) + rowSize * tmp.Height;

    hDIB := GlobalAlloc(GMEM_MOVEABLE, dibSize);
    if hDIB = 0 then Exit;

    pDIB := GlobalLock(hDIB);
    if not Assigned(pDIB) then
    begin
      GlobalFree(hDIB);
      Exit;
    end;

    try
      // �w�b�_������
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
          Inc(srcLine); // skip alpha
        end;
        Inc(pBits, rowSize);
      end;

    finally
      GlobalUnlock(hDIB);
    end;

    Result := hDIB;

  finally
    tmp.Free;
  end;
end;

function BuildClipboardPNG(png: TPngImage): HGLOBAL;
var
  stream: TMemoryStream;
  hPNG: HGLOBAL;
  ptr: Pointer;
begin
  Result := 0;

  if not Assigned(png) then Exit;

  stream := TMemoryStream.Create;
  try
    png.SaveToStream(stream);
    stream.Position := 0;

    hPNG := GlobalAlloc(GMEM_MOVEABLE, stream.Size);
    if hPNG = 0 then Exit;

    ptr := GlobalLock(hPNG);
    if not Assigned(ptr) then
    begin
      GlobalFree(hPNG);
      Exit;
    end;

    try
      Move(stream.Memory^, ptr^, stream.Size);
    finally
      GlobalUnlock(hPNG);
    end;

    Result := hPNG;

  finally
    stream.Free;
  end;
end;

function BuildClipboardBitmap(bmp: TBitmap): HBITMAP;
begin
  Result := 0;
  if not Assigned(bmp) then Exit;

  // CopyImage �ɂ�� HBITMAP �𕡐�
  Result := CopyImage(bmp.Handle, IMAGE_BITMAP, 0, 0,
                      LR_COPYRETURNORG or LR_COPYDELETEORG);
end;

procedure SetClipboardText(const Text: string);
begin
  ClipboardUpdateIsSelf := True;
  Clipboard.Open;
  try
    Clipboard.Clear;
    Clipboard.AsText := Text;
  finally
    Clipboard.Close;
  end;
end;

procedure SetClipboardBitmap(Bmp: TBitmap);
const
  CF_PNG_NAME = 'PNG';
var
  CF_PNG: UINT;
  hBmp: HBITMAP;
  hDIB: HGLOBAL;
  hPNG: HGLOBAL;
  png: TPngImage;
begin
  if not Assigned(Bmp) then Exit;

  CF_PNG := RegisterClipboardFormat(CF_PNG_NAME);
  ClipboardUpdateIsSelf := True;

  // --- �f�[�^���� ---
  hBmp := BuildClipboardBitmap(Bmp);
  hDIB := BuildClipboardDIB(Bmp);

  // PNG�͈�x TBitmap �� TPngImage �֕ϊ����Ă���ۑ�
  png := TPngImage.Create;
  try
    png.Assign(Bmp);
    hPNG := BuildClipboardPNG(png);
  finally
    png.Free;
  end;

  // --- �N���b�v�{�[�h�ɓo�^ ---
  if OpenClipboard(0) then
  try
    EmptyClipboard;
    if hBmp <> 0 then SetClipboardData(CF_BITMAP, hBmp);
    if hDIB <> 0 then SetClipboardData(CF_DIB, hDIB);
    if hPNG <> 0 then SetClipboardData(CF_PNG, hPNG);
  finally
    CloseClipboard;
  end;
end;

procedure SetClipboardPng(Png: TPngImage);
const
  CF_PNG_NAME = 'PNG';
var
  CF_PNG: UINT;
  hBmp: HBITMAP;
  hDIB: HGLOBAL;
  hPNG: HGLOBAL;
  bmp: TBitmap;
begin
  if not Assigned(Png) then Exit;

  CF_PNG := RegisterClipboardFormat(CF_PNG_NAME);
  ClipboardUpdateIsSelf := True;

  // --- TPngImage �� TBitmap �ɕϊ� ---
  bmp := TBitmap.Create;
  try
    bmp.PixelFormat := pf32bit;
    bmp.Width := Png.Width;
    bmp.Height := Png.Height;
    bmp.Canvas.Draw(0, 0, Png);

    // --- �e�t�H�[�}�b�g�̃r���h ---
    hBmp := BuildClipboardBitmap(bmp);
    hDIB := BuildClipboardDIB(bmp);
  finally
    bmp.Free;
  end;

  // --- PNG �͂��̂܂ܕۑ��\ ---
  hPNG := BuildClipboardPNG(Png);

  // --- �N���b�v�{�[�h�ɓo�^ ---
  if OpenClipboard(0) then
  try
    EmptyClipboard;
    if hBmp <> 0 then SetClipboardData(CF_BITMAP, hBmp);
    if hDIB <> 0 then SetClipboardData(CF_DIB, hDIB);
    if hPNG <> 0 then SetClipboardData(CF_PNG, hPNG);
  finally
    CloseClipboard;
  end;
end;

// �N���b�v�{�[�h����r�b�g�}�b�v�擾
function GetClipboardBitmap(Bitmap : TBitmap) : Boolean;
begin
  Result := False;
  if not SafeOpenClipboard(0) then exit;
  try
    Result := GetClipboardBitmapSub(Bitmap);
  finally
    CloseClipboard;
  end;
end;
// �N���b�v�{�[�h����Png�摜�擾
function GetClipboardPng(Png : TPngImage) : Boolean;
begin
  Result := False;
  if not SafeOpenClipboard(0) then exit;
  try
    Result := GetClipboardPngSub(Png);
  finally
    CloseClipboard;
  end;
end;
// �N���b�v�{�[�h����e�L�X�g�擾
function GetClipboardText(var Text : string) : Boolean;
begin
  Result := False;
  if not SafeOpenClipboard(0) then exit;
  try
    Result := GetClipboardTextSub(Text);
  finally
    CloseClipboard;
  end;
end;



initialization
  CF_PNG := RegisterClipboardFormat('PNG');

end.
