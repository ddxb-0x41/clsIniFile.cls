# clsIniFile - VBA INI ファイル操作ユーティリティ

このクラスは、VBA（Visual Basic for Applications）環境で `.ini` ファイルの読み書きを簡単に行うためのユーティリティクラスです。

## 概要

- `Section`・`Key`・`Value` による設定ファイル構造をシンプルに扱えます。
- 読み込み・追加・変更・削除はすべて**メモリ上の操作**です。
- `SaveFile` メソッドにより任意のタイミングでファイルに保存可能。
- 読み込み時、改行コードを自動検出します。

---

## インストール方法（Installation）

このクラスは、VBA（Visual Basic for Applications）上で `.ini` 形式の設定ファイルを簡易的に読み書きするためのユーティリティです。設定はすべて**メモリ上で管理され**、明示的に保存操作を行うまでファイルには書き出されません。

### 1. クラスモジュールの追加

1. Excel を開き、`Alt + F11` を押して **VBAエディター** を起動します。
2. メニューから `挿入 > クラスモジュール` を選択します。
3. プロパティウィンドウ（`F4`）で、クラス名を `clsIniFile` に変更します。
4. 提供されている `clsIniFile.cls` のコード全文を貼り付けます。

### 2. 使用準備

任意の標準モジュールで次のようにインスタンスを作成します：

```vba
Dim ini As New clsIniFile
```

この時点で、自動的に `ThisWorkbook` と**同じフォルダ・同名の `.ini` ファイル**を読み込みます。

例：  
`Book1.xlsm` → `Book1.ini`

> ※ ファイルが存在しない場合でもエラーにはならず、初期状態の空データとして内部に読み込まれます。  
> ※ 読み込んだ `.ini` ファイルの内容は、インスタンス内の辞書構造で保持されます。  
> ※ `SaveFile` メソッドを実行するまで、変更はファイルに反映されません。

---

## 使い方（Usage）

### 値の取得と設定

```vba
' SectionとKeyを事前に設定
ini.Section = "設定"
ini.Key = "ユーザー名"

' 値を設定
ini.Value = "Taro"

' 値を取得
Debug.Print ini.Value  ' => Taro
```

### Item プロパティを使用した一括指定

```vba
' 値の設定
ini.Item("設定", "パス") = "C:\Temp"

' 値の取得（この時 SectionとKeyにも格納される）
Debug.Print ini.Item("設定", "パス")  ' => C:\Temp
```

> ※ `Item("設定", "パス")` を実行すると、`.Section` と `.Key` プロパティも更新されます。

---

### セクション・キーの確認

```vba
If ini.IsExistsSection("設定") Then
    Debug.Print "セクションあり"
End If

If ini.IsExistsKey("設定", "ユーザー名") Then
    Debug.Print "キーあり"
End If
```

---

### セクションやキーの削除（メモリ上のみ）

```vba
ini.RemoveKey "設定", "ユーザー名"
ini.RemoveSection "ログ情報"
```

> ※ 削除はすべて**メモリ上の操作**であり、`SaveFile` を呼び出すまでファイルに反映されません。

---

### セクション一覧・キー一覧の取得

```vba
Dim sec As Variant
For Each sec In ini.GetSections()
    Debug.Print sec
Next

Dim key As Variant
For Each key In ini.GetSectionKeys("設定")
    Debug.Print key
Next
```

---

### ファイルへの保存

```vba
ini.SaveFile "C:\Temp\MySettings.ini"
```

#### 保存時のオプション（省略可）

```vba
ini.SaveFile "C:\Temp\MySettings.ini", adUTF8, adLF, True
```

- `文字コード`（EnumCharacterCode）：`adSHIFT_JIS`, `adUTF8`, `adUTF16`, `adASCII`
- `改行コード`（EnumLineFeedCode）：`adCRLF`, `adCR`, `adLF`
- `AddSectionUnitNewline`：セクションごとに空行を挿入するか

---

### 別のファイルを読み込む

```vba
ini.SetFile "C:\Temp\Other.ini"
```

> ※ `SetFile` 実行後はその内容が現在のメモリに上書きされます。

---

## 初期値について

- **読み込み対象ファイル**：  
  `ThisWorkbook` と同じフォルダ、同名の `.ini` ファイルを自動的に読み込みます。

- **書き込み（保存）時の初期値**：  
  - 文字コード：`Shift_JIS`（adSHIFT_JIS）  
  - 改行コード：`CRLF`（adCRLF）  
  - セクション毎に空行を追加：`True`

---

## ライセンス

MIT License
