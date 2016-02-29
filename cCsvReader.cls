VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCsvReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'フォルダは以下にまとめてあるCVSファイルからFXの取引情報を引っ張ってくる。
Private Type Tables
    Table() As Variant
End Type
Private m_FolderPath As String
Private m_Total As cCsvVal
Private m_Csv() As cCsvVal
Private m_CsvTables() As Tables
Private m_TotalTable() As Variant
Property Get FolderPath()
    FolderPath = m_FolderPath
End Property
Property Get UB_csv()
    UB_csv = UBound(m_Csv)

End Property
Property Get csv(ByVal Index As Long) As cCsvVal
    Set csv = m_Csv(Index)
    
End Property
Property Get Total() As cCsvVal
    Set Total = m_Total

End Property
Property Get CsvTables(ByVal Index As Long) As Variant()
    CsvTables = m_CsvTables(Index).Table
    
End Property
Property Get TotalTable() As Variant()
    TotalTable = m_TotalTable

End Property
Private Sub Class_Initialize()
    ReDim m_Csv(0)
    Set m_Csv(0) = New cCsvVal
    Set m_Total = New cCsvVal
    
    'とりあえずTable関連も初期化
    ReDim m_CsvTables(0)
    ReDim m_CsvTables(0).Table(0, 0)
        
    ReDim m_TotalTable(0, 0)
    
End Sub
'まずは、CVSが入っているフォルダのパスを取得
Public Sub GetPath()

    'フォルダ選択ダイアグからフォルダパスを取得。キャンセルを押したらマクロ終了
    Dim Path As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Please select folder that include csv files."
        If .Show = True Then
            Path = .SelectedItems(1)
        Else
            End
        End If
    
    End With
    
    m_FolderPath = Path
    
End Sub
Public Sub GetPath_Debug()
    'デバック用
    m_FolderPath = "C:\Users\katsumi\LionFXWin\Data\csv"
    
End Sub
'フォルダ配下のcsvファイルから必要なデータを作成
Public Sub GetValue()
    'ファイル名を取得
    Dim File As String
    Dim ci As Long: ci = 0
    File = Dir(m_FolderPath & "\*csv", vbNormal)
    Do While File <> ""
        ReDim Preserve m_Csv(ci)
        Set m_Csv(ci) = New cCsvVal
        
        'ここでファイルを開いて値を設定
        Call GetValue2(File, m_Csv(ci))
        ci = ci + 1
        File = Dir()
    Loop

End Sub
Private Sub GetValue2(ByVal File As String, ByRef csv As cCsvVal)
    'とりあえずファイル名を取得しておく
    csv.FileName = File
    
    'ファイルを開いて必要なデータを取得
    Dim Buf As String
    Dim sBuf() As String
    Dim di As Long: di = 0
    
    Open m_FolderPath & "\" & File For Input As #1
                
        'データがある行までBufを移動。
        Dim n As Long
        For n = 0 To 2: Line Input #1, Buf: Next
        
        Do Until EOF(1)
            Line Input #1, Buf
            sBuf = Split(Buf, ",")
            
            'diが0以上ならばm_csvのデータをインクリメント
            If di > 0 Then csv.IncArr
            
            '値を取得
            With csv
                .Day(di) = sBuf(0)
                .Trade(di) = sBuf(16)
                .Trade_pip(di) = sBuf(14)
            End With
            di = di + 1
        Loop
    Close #1
        
    '最後に売買損益累積を計算
    csv.CalcSum
    
End Sub
Public Sub MakeTotal()
    '各シートのDataを連結したm_Totalを作成する。
    
    'とりあえずファイル名にTotalを入れておく。
    m_Total.FileName = "Total"
    
    Dim ti As Long: ti = 0
    Dim n As Long
    For n = 0 To UBound(m_Csv)
        Dim m As Long
        For m = 0 To m_Csv(n).UB_DataList
            If ti > 0 Then m_Total.IncArr
            '値を格納
            With m_Total
                .Day(ti) = m_Csv(n).Day(m)
                .Trade(ti) = m_Csv(n).Trade(m)
                .Trade_pip(ti) = m_Csv(n).Trade_pip(m)
            End With
            ti = ti + 1
        Next
    Next
    
    'Sumを作成
    m_Total.CalcSum
    
End Sub
Public Sub MakeTables()
    'CsvとTotalの2次元配列に格納する。
    Dim n As Long
    For n = 0 To UBound(m_Csv)
        '格納用の2次元入れるを宣言
        ReDim Preserve m_CsvTables(n)
        
        'CSVの値を2次元配列化
        m_CsvTables(n).Table = SetTableVal(m_CsvTables(n).Table, m_Csv(n))
    Next
    
    'Totalの値を2次元配列化
    m_TotalTable = SetTableVal(m_TotalTable, m_Total)
    
End Sub
Private Function SetTableVal(ByRef Table() As Variant, ByRef CsvVal As cCsvVal)
            
    Dim ubArr As Long
    ubArr = CsvVal.UB_DataList
        
    'Tableを配列化
    ReDim Table(ubArr, 5)
    
    Dim n As Long
    For n = 0 To ubArr
        Table(n, 0) = n + 1
        Table(n, 1) = CsvVal.Day(n)
        Table(n, 2) = CsvVal.Trade(n)
        Table(n, 3) = CsvVal.Trade_pip(n)
        Table(n, 4) = CsvVal.sumTrade(n)
        Table(n, 5) = CsvVal.sumTrade_pip(n)
    Next
    
    SetTableVal = Table
    
End Function
