VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCsvReader_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetPath(ByVal NG As Boolean)
    Dim c As cCsvReader
    Set c = New cCsvReader
    
    Call c.GetPath
        
    If c.FolderPath = "" Then
        Debug.Print "NG : cCsvReader : GetPath : No setting folder path"
        NG = True
    End If
    
    GetPath = NG
    
End Function
Public Function GetPath_Debug(ByVal NG As Boolean)
    Dim c As cCsvReader
    Set c = New cCsvReader
    
    c.GetPath_Debug
    
    If c.FolderPath = "" Then
        Debug.Print "NG : cCsvReader : GetPath_Debug : No Setting folder path"
        NG = True
    End If
    
    GetPath_Debug = NG
    
End Function
Public Function IOCheck(ByVal NG As Boolean)
    Dim c As cCsvReader
    Set c = New cCsvReader
    
    If c.UB_csv <> 0 Then
        Debug.Print "NG : cCsvReader : IOCheck : UB_csv is wrong"
        NG = True
    End If
    
    If c.csv(0).FileName <> "" Then
        Debug.Print "NG : cCsvReader : IOCheck : Csv is wrong"
        NG = True
    End If
    
    If c.Total.FileName <> "" Then
        Debug.Print "NG : cCsvReader : IOCheck : Total is wrong"
        NG = True
    End If
    
    If c.CsvTables(0)(0, 0) <> "" Then
        Debug.Print "NG : cCsvReader : IOCheck : CsvTables is wrong"
        NG = True
    End If
    
    If c.TotalTable()(0, 0) <> "" Then
        Debug.Print "NG : cCsvReader : IOCheck : TotalTable is wrong"
        NG = True
    End If
    
    IOCheck = NG
    
End Function

Public Function GetValue(ByVal NG As Boolean)

    Dim c As cCsvReader
    Set c = New cCsvReader
    
    c.GetPath_Debug
    c.GetValue
    
    Dim n As Long
    For n = 0 To c.UB_csv
        NG = BlunkCheck(NG, c.csv(n), "GetValue")
        NG = AllZeroCheck(NG, c.csv(n), "GetValue")
    Next
    
    GetValue = NG
    
End Function
Public Function MakeTotal(ByVal NG As Boolean)
    Dim c As cCsvReader
    Set c = New cCsvReader
    
    c.GetPath_Debug
    c.GetValue
    c.MakeTotal
    
    'FileNameとDayは空欄はNG
    NG = BlunkCheck(NG, c.Total, "MakeTotale")
    
    '配列がすべて0だと不自然なのでアラーム上げるようにする｡
    NG = AllZeroCheck(NG, c.Total, "MakeTotal")
        
    '次は書くcsvのDataLIstの配列数に着目して､
    '正しく結合されていることを確認する｡
    NG = CombineCheck(NG, c)

    MakeTotal = NG
    
End Function
Private Function CombineCheck(ByVal NG As Boolean, ByRef c As cCsvReader)
    'まずはTotalのデータリストのデータ数を取得
    Dim tCnt As Long
    tCnt = c.Total.UB_DataList + 1
    
    '次にcsvとデータの合計を計算
    Dim n As Long
    Dim cCnt  As Long
    For n = 0 To c.UB_csv
        cCnt = cCnt + c.csv(n).UB_DataList + 1
    Next
    
    If tCnt <> cCnt Then
        Debug.Print "NG : cCsvReader : MakeTotal : CombinheCheck is NG"
        NG = True
    End If
    
    CombineCheck = NG

End Function
Private Function BlunkCheck(ByVal NG As Boolean _
                                            , ByRef csv As cCsvVal _
                                            , ByVal mName As String)
    'Day Trade Tarade_pipは空白ではNG
    If csv.FileName = "" Then
        Debug.Print "NG : cCsvReader : " & mName & " : FileName Is Blunk"
        NG = True
    End If
    
    Dim m As Long
    For m = 0 To csv.UB_DataList
        If csv.Day(m) = #12:00:00 AM# Then
            Debug.Print "NG : cCsvReader : " & mName & " : Day is Blunk Index = " & m
            NG = True
            Exit For
        End If
    Next
    
    BlunkCheck = NG

End Function
Private Function AllZeroCheck(ByVal NG As Boolean _
                                            , ByRef csv As cCsvVal _
                                            , ByVal mName As String)
    
    Dim TradeZero As Boolean: TradeZero = False
    Dim Trade_pipZero As Boolean: Trade_pipZero = False
    Dim sumTradeZero As Boolean: sumTradeZero = False
    Dim sumTrade_pipZero As Boolean: sumTrade_pipZero = False
    
    Dim m As Long
    For m = 0 To csv.UB_DataList
        If csv.Trade(m) <> 0 Then TradeZero = True
        If csv.Trade_pip(m) <> 0 Then Trade_pipZero = True
        If csv.sumTrade(m) <> 0 Then sumTradeZero = True
        If csv.sumTrade(m) <> 0 Then sumTrade_pipZero = True
    Next
    
    If TradeZero = False Then
        Debug.Print "Warning : cCsvReader : " & mName & " : Trade : All 0 is OK?"
        NG = True
    End If
    
    If Trade_pipZero = False Then
        Debug.Print "Warning : cCsvReader : " & mName & " : Trade_pip : All 0 is OK?"
        NG = True
    End If
    
    If sumTradeZero = False Then
        Debug.Print "Warning : cCsvReader : " & mName & " : sumTrade : All 0 is OK?"
        NG = True
    End If
    
    If sumTrade_pipZero = False Then
        Debug.Print "Warning : cCsvReader : " & mName & " : sumTrade_pip : All 0 is OK?"
        NG = True
    End If
    
    AllZeroCheck = NG
    
End Function
Public Function MakeTables(ByVal NG As Boolean)
    Dim c As cCsvReader
    Set c = New cCsvReader
    
    c.GetPath_Debug
    c.GetValue
    c.MakeTotal
    c.MakeTables
    
    'Tableの要素数をチェック
    NG = ArrCountCheck(NG, c)
    
    '空欄はNG
    NG = MakeTables_BlankCheck(NG, c)
    
        
    MakeTables = NG

End Function
Private Function ArrCountCheck(ByVal NG As Boolean, ByRef c As cCsvReader)
    Dim n As Long
    For n = 0 To c.UB_csv
        If c.csv(n).UB_DataList <> UBound(c.CsvTables(n), 1) Then
            Debug.Print "NG : cCsvReader : MakeTables : Arr Count no same csv.UBDataList"
            NG = True
        End If
    Next
    
    If c.Total.UB_DataList <> UBound(c.TotalTable) Then
        Debug.Print "NG : cCsvReader : MakeTables : Arr Count no same Total.UBDataList"
        NG = True
    End If

    ArrCountCheck = NG

End Function
Private Function MakeTables_BlankCheck(ByVal NG As Boolean, ByRef c As cCsvReader)
    Dim n As Long
    For n = 0 To c.UB_csv
    Dim m As Long
    For m = 0 To c.csv(n).UB_DataList
        Dim l As Long
        For l = 0 To 4
            If c.CsvTables(n)(n, l) = "" Then
                Debug.Print "NG : cCsvReader : MakeTables : Blunk Error of CsvTable"
                NG = True
            End If
        Next
    Next
    Next
            
    Dim tn As Long
    Dim tm As Long
    For tn = 0 To UBound(c.TotalTable, 1)
    For tm = 0 To UBound(c.TotalTable, 2)
        If c.TotalTable()(tn, tm) = "" Then
            Debug.Print "NG : cCsvReader : MakeTables : Blunk Error of TotalTable"
            NG = True
        End If
    Next
    Next
    
    MakeTables_BlankCheck = NG
    
End Function

