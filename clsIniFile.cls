VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--- 文字コード列挙型の定義
Enum EnumCharacterCode
    adUTF8 = 65001
    adSHIFT_JIS = 932
    adUTF16 = 1200
    adASCII = 1252
End Enum

'--- 改行コード列挙型の定義
Enum EnumLineFeedCode
    adCRLF = -1
    adCR = 13
    adLF = 10
End Enum

'--- ADODBストリーム定数と内部変数の宣言
Private Const adSaveCreateOverWrite = 2
Private Const adReadAll = -1
Private Const adTypeText = 2
Private Const adWriteLine = 1

Private FSO As Object
Private Parameters As Object
Private CharSet As Object
Private LineSeparator As Object
Private mSection As String
Private mKey As String

'--- クラス初期化処理（辞書作成と初期ファイル読み込み）
Private Sub Class_Initialize()
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Parameters = CreateObject("Scripting.Dictionary")
    Set CharSet = CreateObject("Scripting.Dictionary")
    Set LineSeparator = CreateObject("Scripting.Dictionary")
    
    With CharSet
        .Add EnumCharacterCode.adSHIFT_JIS, "Shift_JIS"
        .Add EnumCharacterCode.adUTF16, "UTF-16"
        .Add EnumCharacterCode.adUTF8, "UTF-8"
        .Add EnumCharacterCode.adASCII, "ASCII"
    End With
    
    With LineSeparator
        .Add EnumLineFeedCode.adCRLF, vbCrLf
        .Add EnumLineFeedCode.adCR, vbCr
        .Add EnumLineFeedCode.adLF, vbLf
    End With
    
    Call SetFile(FSO.BuildPath(ThisWorkbook.Path, FSO.GetBaseName(ThisWorkbook.Name) & ".ini"))
End Sub

'--- クラス終了処理（オブジェクト開放）
Private Sub Class_Terminate()
    Set FSO = Nothing
    Set Parameters = Nothing
    Set CharSet = Nothing
    Set LineSeparator = Nothing
End Sub

'--- Section/Key/Valueプロパティ（取得・設定）

Property Get Section() As String
    Section = mSection
End Property
Property Let Section(ByVal NewValue As String)
    mSection = UCase(NewValue)
End Property
Property Get Key() As String
    Key = mKey
End Property
Property Let Key(ByVal NewValue As String)
    mKey = UCase(NewValue)
End Property
Property Get Value() As String
    If Parameters.Exists(mSection) Then
        If Parameters(mSection).Exists(mKey) Then
            Value = Parameters(mSection)(mKey)
            Exit Property
        End If
    End If
    Value = ""
End Property
Property Let Value(ByVal NewValue As String)
    If mSection = "" Or mKey = "" Then Exit Property
    If Not Parameters.Exists(mSection) Then
        Parameters.Add mSection, CreateObject("Scripting.Dictionary")
    End If
    Parameters(mSection)(mKey) = NewValue
End Property

'--- 任意のSectionとKeyのValue取得・設定
Property Get Item(ByVal SectionName As String, ByVal KeyName As String) As String
    SectionName = UCase(SectionName)
    KeyName = UCase(KeyName)
    If Parameters.Exists(SectionName) Then
        If Parameters(SectionName).Exists(KeyName) Then
            Item = Parameters(SectionName)(KeyName)
            Exit Property
        End If
    End If
    Item = ""
End Property
Property Let Item(ByVal SectionName As String, ByVal KeyName As String, ByVal NewValue As String)
    SectionName = UCase(SectionName)
    KeyName = UCase(KeyName)
    If SectionName = "" Or KeyName = "" Then Exit Property
    If Not Parameters.Exists(SectionName) Then
        Parameters.Add SectionName, CreateObject("Scripting.Dictionary")
    End If
    Parameters(SectionName).Item(KeyName) = NewValue
End Property

'--- セクションとキーの存在確認
Function IsExistsSection(ByVal SectionName As String) As Boolean
    SectionName = UCase(SectionName)
    IsExistsSection = Parameters.Exists(SectionName)
End Function
Function IsExistsKey(ByVal SectionName As String, ByVal KeyName As String) As Boolean
    SectionName = UCase(SectionName)
    KeyName = UCase(KeyName)
    If IsExistsSection(SectionName) Then
        IsExistsKey = Parameters(SectionName).Exists(KeyName)
    Else
        IsExistsKey = False
    End If
End Function

'--- セクション名・キー一覧の取得
Function GetSections() As Variant
    GetSections = Parameters.keys
End Function
Function GetSectionKeys(ByVal SectionName As String) As Variant
    SectionName = UCase(SectionName)
    If Parameters.Exists(SectionName) Then
        GetSectionKeys = Parameters(SectionName).keys
    Else
        GetSectionKeys = Array()
    End If
End Function

'--- INIファイルへの保存処理
Sub SaveFile(ByVal FilePath As String, _
    Optional CharacterCode As EnumCharacterCode = EnumCharacterCode.adSHIFT_JIS, _
    Optional LineFeedCode As EnumLineFeedCode = EnumLineFeedCode.adCRLF, _
    Optional AddSectionUnitNewline As Boolean = True)
    
    On Error GoTo Catch
    
    If FSO.FileExists(FilePath) Then FSO.DeleteFile FilePath
    
    Dim SectionName As Variant
    Dim KeyName As Variant
    
    With CreateObject("ADODB.Stream")
        .Open
        .Type = adTypeText
        .CharSet = CharSet(CharacterCode)
        .LineSeparator = LineFeedCode
        
        .WriteText ";" & String(50, "="), adWriteLine
        .WriteText "; Computer Name: " & VBA.Environ("COMPUTERNAME"), adWriteLine
        .WriteText "; User Name    : " & VBA.Environ("USERNAME"), adWriteLine
        .WriteText "; Module       : clsIniFile", adWriteLine
        .WriteText "; Generated    : " & Date & " " & Time, adWriteLine
        .WriteText ";" & String(50, "="), adWriteLine
        
        For Each SectionName In Parameters.keys
            .WriteText "[" & SectionName & "]", adWriteLine
            If Parameters(SectionName).Count = 0 Then
                If Not AddSectionUnitNewline Then
                    .WriteText "", adWriteLine
                End If
            Else
                For Each KeyName In Parameters(SectionName).keys
                    .WriteText KeyName & "=" & Parameters(SectionName).Item(KeyName), adWriteLine
                Next
            End If
            If AddSectionUnitNewline Then .WriteText "", adWriteLine
        Next
        
        .SaveToFile FilePath, adSaveCreateOverWrite
        .Close
    End With
    Exit Sub
Catch:
    Debug.Print "[SaveFile Error] " & Err.Number & ": " & Err.Description & " (" & FilePath & ")"
End Sub

'--- INIファイルの読み込み処理
Function SetFile(ByVal FilePath As String, _
    Optional CharacterCode As EnumCharacterCode = EnumCharacterCode.adSHIFT_JIS, _
    Optional LineFeedCode As EnumLineFeedCode = -9999) As Boolean
    
    On Error GoTo Catch
    
    If Not (FSO.FileExists(FilePath)) Then Exit Function
    
    Dim ReadData As Variant
    Dim ReadLine As String
    Dim SectionName As String
    Dim KeyName As String
    Dim Value As String
    Dim v As Variant
    Dim i As Long
    Dim TextStream As String
    
    With CreateObject("ADODB.Stream")
        .Open
        .Type = adTypeText
        .CharSet = CharSet(CharacterCode)
        .LoadFromFile FilePath
        TextStream = .ReadText(adReadAll)
        .Close
    End With
    
    Dim AutoLineFeed As EnumLineFeedCode
    If LineFeedCode = -9999 Then
        AutoLineFeed = DetectLineFeedCode(TextStream)
    Else
        AutoLineFeed = LineFeedCode
    End If
    ReadData = Split(TextStream, LineSeparator(AutoLineFeed))
    
    For i = LBound(ReadData) To UBound(ReadData)
        ReadLine = Trim(ReadData(i))
        If Len(ReadLine) = 0 Or Left(ReadLine, 1) Like "[;#/]" Then
            ' コメントまたは空行をスキップ
        ElseIf Left(ReadLine, 1) = "[" And Right(ReadLine, 1) = "]" Then
            SectionName = UCase(Mid(ReadLine, 2, Len(ReadLine) - 2))
        ElseIf ReadLine Like "*=*" And Not SectionName = "" Then
            v = Split(ReadLine, "=", 2)
            KeyName = UCase(Trim(v(0)))
            Value = Trim(v(1))
            If Not Parameters.Exists(SectionName) Then
                Parameters.Add SectionName, CreateObject("Scripting.Dictionary")
            End If
            Parameters(SectionName).Item(KeyName) = Value
        End If
    Next
    SetFile = True
    Exit Function
Catch:
    Debug.Print "[SetFile Error] " & Err.Number & ": " & Err.Description & " (" & FilePath & ")"
    SetFile = False
End Function

'--- セクションの削除処理
Sub RemoveSection(ByVal SectionName As String)
    SectionName = UCase(SectionName)
    If Parameters.Exists(SectionName) Then
        Parameters.Remove SectionName
    End If
End Sub

'--- キーの削除処理
Sub RemoveKey(ByVal SectionName As String, ByVal KeyName As String)
    SectionName = UCase(SectionName)
    KeyName = UCase(KeyName)
    If Parameters.Exists(SectionName) Then
        If Parameters(SectionName).Exists(KeyName) Then
            Parameters(SectionName).Remove KeyName
        End If
    End If
End Sub

'--- 自動改行コード判定処理
Private Function DetectLineFeedCode(ByVal TextStream As String) As EnumLineFeedCode
    Dim LineFeedCode As Variant
    Dim MaxCount As Long: MaxCount = -1
    Dim BestLineFeedCode As EnumLineFeedCode
    Dim CurrentCount As Long
    For Each LineFeedCode In LineSeparator.keys
        CurrentCount = UBound(Split(TextStream, LineSeparator(LineFeedCode)))
        If MaxCount < CurrentCount Then
            BestLineFeedCode = LineFeedCode
            MaxCount = CurrentCount
        End If
    Next
    DetectLineFeedCode = BestLineFeedCode
End Function


