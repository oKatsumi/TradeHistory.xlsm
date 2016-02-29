VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCSVVal_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function IOCheck(ByVal NG As Boolean)
    
    Dim c As cCsvVal
    Set c = New cCsvVal
    
    c.FileName = "Test1"
    If c.FileName <> "Test1" Then Call Err(NG, "File Name")
    
    c.Day(0) = 11111
    If c.Day(0) <> 11111 Then Call Err(NG, "Day")
    
    c.Trade(0) = 22222
    If c.Trade(0) <> 22222 Then Call Err(NG, "Trade")
    
    c.Trade_pip(0) = 33333
    If c.Trade_pip(0) <> 33333 Then Call Err(NG, "Trade_pip")
    
    c.sumTrade(0) = 44444
    If c.sumTrade(0) <> 44444 Then Call Err(NG, "sumTrade")
    
    c.sumTrade_pip(0) = 55555
    If c.sumTrade_pip(0) <> 55555 Then Call Err(NG, "sumTrade_pip")
    
    If c.UB_DataList <> 0 Then Call Err(NG, "UB_DataList")
    
        
    IOCheck = NG
    
End Function
Private Sub Err(ByRef NG As Boolean, ByVal ErrStr As String)
    Debug.Print "NG : cCsvVal : IOCheck : " & ErrStr & " is wrong"
    NG = True
    
End Sub
Public Function IncArr(ByVal NG As Boolean)
    Dim c As cCsvVal
    Set c = New cCsvVal
    
    Dim ub As Long
    ub = c.UB_DataList
    
    c.IncArr
        
    If ub + 1 <> c.UB_DataList Then
        Debug.Print "NG : cCsVVal : IncArr"
        NG = True
    End If
    
    IncArr = NG

End Function
Public Function CalcSum(ByVal NG As Boolean)

    Dim c As cCsvVal
    Set c = New cCsvVal
    
    Dim n As Long
    For n = 0 To 2
        If n > 0 Then c.IncArr
        c.Trade(n) = n
        c.Trade_pip(n) = n * 2
    Next
    
    c.CalcSum
    
    Dim NGFlag1 As Boolean: NGFlag1 = False
    If c.sumTrade(0) <> 0 Then NGFlag1 = True
    If c.sumTrade(1) <> 1 Then NGFlag1 = True
    If c.sumTrade(2) <> 3 Then NGFlag1 = True
    
    If NGFlag1 = True Then
        Debug.Print "NG : cCsvVal : CalcSum : SumTrade is NG"
        NG = True
    End If
    
    Dim NGFlag2 As Boolean: NGFlag2 = False
    If c.sumTrade_pip(0) <> 0 Then NGFlag2 = True
    If c.sumTrade_pip(1) <> 2 Then NGFlag2 = True
    If c.sumTrade_pip(2) <> 6 Then NGFlag2 = True
    
    If NGFlag2 = True Then
        Debug.Print "NG : cCsvVal : CalcSum : SumTrade_pip is NG"
        NG = True
    End If
    
    CalcSum = NG
    
End Function
