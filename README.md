# clsIniFile - VBA INIファイル操作クラス

このクラスモジュールは、INIファイルを簡単に読み書き・管理できるVBAクラスです。  
セクション・キー単位で値の取得／設定／削除が可能で、文字コードや改行コードの指定にも対応しています。

---

## 特徴

- `.ini` ファイルの読み書きに対応
- セクション・キーごとのアクセス
- 複数文字コード (`Shift_JIS`, `UTF-8`, `UTF-16`, `ASCII`)
- 改行コードの自動検出と指定
- `ThisWorkbook` と同名の `.ini` ファイルを自動読み込み

---

## 初期化

```vba
Dim ini As New clsIniFile
```

- 初期化時に `ThisWorkbook` の `.ini` ファイルを自動読み込み  
  例：`Sample.xlsm` → `Sample.ini`

- **読み込み時の初期設定**
  - 文字コード：`Shift_JIS`
  - 改行コード：自動判別

---

## 値の取得

```vba
Dim mode As String
mode = ini.Item("SYSTEM", "MODE")

' またはプロパティを使って
ini.Section = "SYSTEM"
ini.Key = "MODE"
mode = ini.Value
```

---

## 値の設定・追加（※メモリ上のみ）

```vba
ini.Item("SYSTEM", "MODE") = "DEV"

' またはプロパティを使って
ini.Section = "SYSTEM"
ini.Key = "MODE"
ini.Value = "DEV"
```

※この時点ではファイルは更新されません。  
※ファイルへ反映するには `SaveFile` を明示的に実行してください。

---

## 削除（※メモリ上のみ）

```vba
ini.RemoveKey "SYSTEM", "MODE"
ini.RemoveSection "SYSTEM"
```

※削除もメモリ上での操作です。  
※ファイルへ反映するには `SaveFile` が必要です。

---

## 存在チェック

```vba
If ini.IsExistsSection("SYSTEM") Then
    MsgBox "セクションあり"
End If

If ini.IsExistsKey("SYSTEM", "MODE") Then
    MsgBox "キーあり"
End If
```

---

## セクション／キー一覧の取得

```vba
Dim sections As Variant
sections = ini.GetSections()

Dim keys As Variant
keys = ini.GetSectionKeys("SYSTEM")
```

---

## ファイルの保存

```vba
ini.SaveFile "C:\config.ini", adUTF8, adLF
```

### 引数（省略可能）

| 引数 | 内容 | 既定値 |
|------|------|--------|
| FilePath | 保存先パス | 必須 |
| CharacterCode | 書き込み時の文字コード | `adSHIFT_JIS` |
| LineFeedCode | 改行コード | `adCRLF` |
| AddSectionUnitNewline | セクション毎に空行を追加する | `True` |

---

## ファイルの再読み込み

```vba
ini.SetFile "C:\config.ini", adUTF8
```

- 改行コードの指定を省略すると、自動検出されます。

---

## 対応文字コード（EnumCharacterCode）

```vba
EnumCharacterCode
    adSHIFT_JIS = 932
    adUTF8      = 65001
    adUTF16     = 1200
    adASCII     = 1252
End Enum
```

---

## 対応改行コード（EnumLineFeedCode）

```vba
EnumLineFeedCode
    adCRLF = -1  ' \r\n (Windows)
    adCR   = 13  ' \r   (Mac Classic)
    adLF   = 10  ' \n   (Unix/Linux)
End Enum
```

---

## ライセンス

MIT License