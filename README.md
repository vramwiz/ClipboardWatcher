# ClipboardWatcher.pas

Delphi 向けのクリップボード監視ユニットです。  
テキスト・ビットマップ・PNG の各形式に対応し、非ポーリング型の `WM_CLIPBOARDUPDATE` メッセージでクリップボードの変更を自動検出します。

---

## ✨ 主な機能

- `WM_CLIPBOARDUPDATE` による非ポーリング型のクリップボード監視
- 対応形式：テキスト（CF_TEXT）、ビットマップ（CF_BITMAP）、PNG（CF_PNG）
- イベントベースで変更通知（`OnChange` イベント）
- 連続イベントのデバウンス処理付き（既知のWindows仕様に対応）
- PNG 形式はカスタム登録により対応

---

## 🔧 使い方

```pascal
uses
  ClipboardWatcher;

var
  Watcher: TClipboardWatcher;

procedure TForm1.FormCreate(Sender: TObject);
begin
  Watcher := TClipboardWatcher.Create;
  Watcher.OnChange :=
    procedure(Sender: TObject; const DataTypes: TClipboardWatcherDataTypes)
    begin
      if cdtText in DataTypes then
        Memo1.Lines.Text := Watcher.Text;

      if cdtBitmap in DataTypes then
        Image1.Picture.Assign(Watcher.Bitmap);

      if cdtPng in DataTypes then
        // PNG処理（保存や表示など）を実行
        ;
    end;

  Watcher.Enabled := True;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  Watcher.Free;
end;
```

---

## 📦 公開プロパティ

| プロパティ | 内容 |
|------------|------|
| `Enabled`  | True にすると監視を開始、False で停止 |
| `Text`     | 取得されたテキストデータ（string） |
| `Bitmap`   | 取得されたビットマップ（TBitmap） |
| `Png`      | 取得された PNG データ（TPngImage） |

---

## 📡 通知イベント

### `OnChange: TClipboardWatcherChangeEvent`

- シグネチャ:
  ```pascal
  TClipboardWatcherChangeEvent = procedure(Sender: TObject; const DataTypes: TClipboardWatcherDataTypes) of object;
  ```
- DataTypes には `cdtText`, `cdtBitmap`, `cdtPng` のいずれかが含まれます。

---

## 🧠 補足技術

- PNG は標準で対応していないため、`RegisterClipboardFormat('PNG')` によってフォーマットを登録
- OpenClipboard の競合対策として `SafeOpenClipboard` を使用（最大10回までリトライ）
- 500ms のデバウンスタイマー（`CLIPBOARD_DEBOUNCE_MS`）で連続通知を抑制


---

## 🧰 グローバル関数の使用例

クリップボードの内容を直接取得・設定するグローバル関数が用意されています。`TClipboardWatcher` を使用せずに、任意のタイミングで利用可能です。

### 📥 データの取得（Get）

```pascal
var
  Text: string;
  Bitmap: TBitmap;
  Png: TPngImage;
begin
  if GetClipboardText(Text) then
    ShowMessage(Text);

  Bitmap := TBitmap.Create;
  try
    if GetClipboardBitmap(Bitmap) then
      Image1.Picture.Assign(Bitmap);
  finally
    Bitmap.Free;
  end;

  Png := TPngImage.Create;
  try
    if GetClipboardPng(Png) then
      Png.SaveToFile('clipboard.png');
  finally
    Png.Free;
  end;
end;
```

### 📤 データの設定（Set）

```pascal
begin
  SetClipboardText('Hello, world!');
  SetClipboardBitmap(Image1.Picture.Bitmap);
  SetClipboardPng(MyPngImage); // TPngImage 型の変数
end;
```
---

これらの関数は内部で `SafeOpenClipboard` による排他制御を行っており、安全に動作します。失敗した場合は `False` を返し、リソースは初期化済みの状態に保たれます。

---

## 📄 ライセンス

MIT License またはプロジェクトの方針に従って自由に設定してください。

---

## 🧑‍💻 作者

Created by **vramwiz**  
Created: 2025-07-10  
Updated: 2025-07-10


