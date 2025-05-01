# clsIniFile.cls
clsIniFile - INIファイル操作用VBAクラス

clsIniFile は、VBA（Visual Basic for Applications）で INI ファイルの読み書きを簡単に行うためのクラスモジュールです。セクション／キー単位での設定情報の保存・読み込みができ、ADODB を利用して文字コードや改行コードも柔軟に扱えます。

## 特徴
INI形式のファイルの読み書き対応
セクション・キーごとのデータ取得／設定が可能
UTF-8 / Shift_JIS / UTF-16 / ASCII 文字コードに対応
CRLF / CR / LF 改行コードに対応
コメントや空行の無視処理あり
自動改行コード検出機能（読み込み時）
```vb
Dim ini As New clsIniFile

' セクション・キーの指定
ini.Section = "GENERAL"
ini.Key = "USERNAME"
ini.Value = "taro"  ' 書き込み

Debug.Print ini.Value  ' 読み込み

' ファイルに保存
ini.SaveFile "C:\test.ini", adUTF8, adCRLF

' ファイルから読み込み
ini.SetFile "C:\test.ini"
```

## 主なメンバー
### プロパティ
Section: 現在操作対象のセクション
Key: 現在操作対象のキー
Value: 指定セクション・キーの値を取得または設定
Item(SectionName, KeyName): 任意のセクションとキーにアクセス

### メソッド
SetFile(FilePath, [CharacterCode], [LineFeedCode]): INIファイル読み込み
SaveFile(FilePath, [CharacterCode], [LineFeedCode], [AddSectionUnitNewline]): INIファイル保存
IsExistsSection(SectionName): セクションの存在確認
IsExistsKey(SectionName, KeyName): キーの存在確認
GetSections(): セクション名一覧取得
GetSectionKeys(SectionName): 指定セクション内のキー一覧取得
RemoveSection(SectionName): セクションの削除
RemoveKey(SectionName, KeyName): キーの削除

## 依存コンポーネント
ADODB.Stream（文字コード対応のために使用）
Scripting.FileSystemObject
