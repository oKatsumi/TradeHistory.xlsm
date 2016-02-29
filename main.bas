Attribute VB_Name = "main"
Option Explicit
Public Sub MakeGraph()
    'csvファイルの読み込み
    Dim csv As cCsvReader
    Set csv = New cCsvReader
    
    With csv
        .GetPath
        .GetValue
        .MakeTotal
        .MakeTables
    End With
    
    Application.ScreenUpdating = False
    
    '表を作成
    Dim Inp As cInputer
    Set Inp = New cInputer
    
    Call Inp.Run(csv)
    
End Sub
Sub t()
    Dim c As Tester
    Set c = New Tester
    c.Run
End Sub
