VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cInputer_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'新規にブックを作成して。エクセルにグラフを表示するクラスのテストクラス。
Option Explicit

Public Function InputData(ByVal NG As Boolean)
    
    Dim csv As cCsvReader
    Set csv = New cCsvReader

    csv.GetPath_Debug
    csv.GetValue
    csv.MakeTotal
    csv.MakeTables

    Dim Inp As cInputer
    Set Inp = New cInputer

    Call Inp.Run(csv)
        
    '目視でOK
    
    InputData = NG
    
End Function

