VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const LINE = "*******************"
Private sTime As Double
Private eTime As Double
Private NG As Boolean
Public Sub Run()
    Call TestStart
    Call TestSuite
    Call TestEnd
End Sub

Private Sub TestSuite()
    
    Debug.Print "cCsvVal ... "
    Dim cCsvVal_Test As New cCsvVal_Test
    NG = cCsvVal_Test.IOCheck(NG)
    NG = cCsvVal_Test.IncArr(NG)
    NG = cCsvVal_Test.CalcSum(NG)
    
    Debug.Print "cCsvReader ..."
    Dim cCsvReader_Test As New cCsvReader_Test
    NG = cCsvReader_Test.IOCheck(NG)
    'NG = cCsvReader_Test.GetPath(NG)
    NG = cCsvReader_Test.GetPath_Debug(NG)
    NG = cCsvReader_Test.GetValue(NG)
    NG = cCsvReader_Test.MakeTotal(NG)
    NG = cCsvReader_Test.MakeTables(NG)
    
    Debug.Print "cInputer ... "
    Dim cInputer_Test As New cInputer_Test
    NG = cInputer_Test.InputData(NG)
    
    
End Sub
Private Sub TestStart()
    'テストのスタート処理
    
    Dim i As Integer
    For i = 0 To 30
        Debug.Print "|"
    Next
    
    NG = False
        
    Debug.Print LINE
    Debug.Print "START"
    Debug.Print LINE
    
    sTime = Timer

End Sub
Private Sub TestEnd()
    'テストの終了処理
        
    eTime = Timer
    
    Debug.Print
    If NG = True Then
        Debug.Print LINE
        Debug.Print "RED !"
        Debug.Print LINE
    Else
        Debug.Print LINE
        Debug.Print "ALL GREEN"
        Debug.Print LINE
    End If
    
    Debug.Print "CalcTime is " & Format(eTime - sTime, "0.000") & " (sec)"
    
End Sub
