VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cInputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'新規にブックを作成して。エクセルにグラフを表示するクラス。
Private m_Book As Workbook
Public Sub Run(ByRef c As cCsvReader)
    Workbooks.Add
    Set m_Book = ActiveWorkbook
    
    'まずはCSVのデータを入力
    Dim n As Long
    For n = 0 To c.UB_csv
        Call InputData(c.CsvTables(n) _
                                , c.csv(n).FileName _
                                , c.csv(n).UB_DataList)
    Next
    
    '次のTotalデータを入力
    Call InputData(c.TotalTable, c.Total.FileName, c.Total.UB_DataList)
    
    '最後の各シートにグラフを作成する。
    Call MakeGraph
    
    'TotalSheetをアクティベイト
    m_Book.Worksheets(c.Total.FileName).Activate
    
    'Sheet1はからシートなので削除
    Application.DisplayAlerts = False
    m_Book.Worksheets("Sheet1").Delete
    Application.DisplayAlerts = True
    
End Sub
Private Sub MakeGraph()
    Dim s As Worksheet
    For Each s In m_Book.Worksheets
        s.Activate
        'データの入っている最終行を取得
        Dim eRow As Long
        eRow = s.Range("A1").SpecialCells(xlLastCell).Row
        
        'グラフ作成(X軸="No.")
        Dim Rng As String
        s.Shapes.AddChart2(240, xlXYScatterLines).Select
        Rng = "$K$1:$K$" & eRow & ", " & "$O$1:$P$" & eRow
        Call SetChartConf(s, ActiveChart, Rng, 1, "B2")

        'グラフ作成(X軸="日付")
        s.Shapes.AddChart2(240, xlXYScatterLines).Select
        Rng = "$L$1:$L$" & eRow & ", " & "$O$1:$P$" & eRow
        Call SetChartConf(s, ActiveChart, Rng, 2, "B19")
        
    Next
End Sub
Private Sub SetChartConf(ByRef s As Worksheet _
                                    , ByRef ActChart As Chart _
                                    , ByVal Rng As String _
                                    , ByVal ChartNo As Long _
                                    , ByVal ChartPos As String)
    With ActChart
        'データ選択
        .SetSourceData Source:=Range(Rng)
        
        '汎例をグラフの下に表示
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        'グラフのスタイル変更(透明な感じのスタイル)
        .ClearToMatchStyle
        .ChartStyle = 240
        
        'チャートの色選択
        .ChartColor = 10
        
        'タイトルをシート名に変更
        .ChartTitle.Text = s.Name
        
    End With
        
    'シートの場所を変更
    s.ChartObjects("グラフ " & ChartNo).Left = s.Range(ChartPos).Left
    s.ChartObjects("グラフ " & ChartNo).Top = s.Range(ChartPos).Top
    
    'ChartNoが2の場合だけ横軸の表示を日付にする。
    If ChartNo = 2 Then
        ActiveChart.Axes(xlCategory).TickLabels.NumberFormatLocal = "m/d;@"
    End If

End Sub
Private Sub InputData(ByRef Table() As Variant _
                                , ByVal FileName As String _
                                , ByVal DataUB As Long)
    m_Book.Sheets.Add
    
    Dim s As Worksheet
    Set s = ActiveSheet
    
    'Sheet名をcsvのファイル名に変更
    s.Name = Replace(FileName, ".csv", "")
    
    'まずは表題を表示
    s.Cells(1, 11).Value = "No."
    s.Cells(1, 12).Value = "日付"
    s.Cells(1, 13).Value = "損益 (円)"
    s.Cells(1, 14).Value = "損益 (pip)"
    s.Cells(1, 15).Value = "累積損益 (円)"
    s.Cells(1, 16).Value = "累積損益 (pip)"
    
    '値を代入
    s.Range("K2:P" & DataUB + 2) = Table

    '線を引いておく
    s.Range("K1:P" & DataUB + 2).Borders.LineStyle = xlContinuous
    
    'オートフィットしておく
    s.Columns("K:P").AutoFit
        
    '表題に色を塗っておく
    s.Range("K1:P1").Interior.ColorIndex = 37 '(暗めの水色)
    
End Sub
