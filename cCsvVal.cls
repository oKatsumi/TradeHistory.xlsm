VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCsvVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_FileName As String
Private m_Day() As Date
Private m_Trade() As Long
Private m_Trade_pip() As Long
Private m_sumTrade() As Long
Private m_sumTrade_pip() As Long
Property Let FileName(ByVal FileName As String)
    m_FileName = FileName
End Property
Property Get FileName() As String
    FileName = m_FileName
End Property
Property Let Day(ByVal Index As Long, ByVal InputDay As Date)
    m_Day(Index) = InputDay

End Property
Property Get Day(ByVal Index As Long) As Date
    Day = m_Day(Index)

End Property
Property Let Trade(ByVal Index As Long, ByVal InputTrade As Long)
    m_Trade(Index) = InputTrade

End Property
Property Get Trade(ByVal Index As Long) As Long
    Trade = m_Trade(Index)

End Property
Property Let Trade_pip(ByVal Index As Long, ByVal InputTrade_pip As Long)
    m_Trade_pip(Index) = InputTrade_pip

End Property
Property Get Trade_pip(ByVal Index As Long) As Long
    Trade_pip = m_Trade_pip(Index)

End Property
Property Let sumTrade(ByVal Index As Long, ByVal InputSumTrade As Long)
    m_sumTrade(Index) = InputSumTrade

End Property
Property Get sumTrade(ByVal Index As Long) As Long
    sumTrade = m_sumTrade(Index)

End Property
Property Let sumTrade_pip(ByVal Index As Long, ByVal InputSumTrade_pip As Long)
    m_sumTrade_pip(Index) = InputSumTrade_pip

End Property
Property Get sumTrade_pip(ByVal Index As Long) As Long
    sumTrade_pip = m_sumTrade_pip(Index)

End Property
Property Get UB_DataList()
    UB_DataList = UBound(m_Day)
    
End Property
Private Sub Class_Initialize()
    'とりあえず配列関連を実態かついでにFileNameも""を入れておく
    m_FileName = ""
    ReDim m_Day(0)
    ReDim m_Trade(0)
    ReDim m_Trade_pip(0)
    ReDim m_sumTrade(0)
    ReDim m_sumTrade_pip(0)
End Sub

Public Function IncArr()
'配列のメンバ変数をインクリメントする
    Dim nextUB As Long
    nextUB = UBound(m_Day) + 1
    
    ReDim Preserve m_Day(nextUB)
    ReDim Preserve m_Trade(nextUB)
    ReDim Preserve m_Trade_pip(nextUB)
    ReDim Preserve m_sumTrade(nextUB)
    ReDim Preserve m_sumTrade_pip(nextUB)

End Function
Public Sub CalcSum()
    Dim n As Long
    Dim sumTrade As Long: sumTrade = 0
    Dim sumTrade_pip As Long: sumTrade_pip = 0
    
    For n = 0 To UBound(m_Day)
        sumTrade = sumTrade + m_Trade(n)
        sumTrade_pip = sumTrade_pip + m_Trade_pip(n)
        
        m_sumTrade(n) = sumTrade
        m_sumTrade_pip(n) = sumTrade_pip
    Next
    
End Sub
