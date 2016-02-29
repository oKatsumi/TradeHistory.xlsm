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
'�V�K�Ƀu�b�N���쐬���āB�G�N�Z���ɃO���t��\������N���X�B
Private m_Book As Workbook
Public Sub Run(ByRef c As cCsvReader)
    Workbooks.Add
    Set m_Book = ActiveWorkbook
    
    '�܂���CSV�̃f�[�^�����
    Dim n As Long
    For n = 0 To c.UB_csv
        Call InputData(c.CsvTables(n) _
                                , c.csv(n).FileName _
                                , c.csv(n).UB_DataList)
    Next
    
    '����Total�f�[�^�����
    Call InputData(c.TotalTable, c.Total.FileName, c.Total.UB_DataList)
    
    '�Ō�̊e�V�[�g�ɃO���t���쐬����B
    Call MakeGraph
    
    'TotalSheet���A�N�e�B�x�C�g
    m_Book.Worksheets(c.Total.FileName).Activate
    
    'Sheet1�͂���V�[�g�Ȃ̂ō폜
    Application.DisplayAlerts = False
    m_Book.Worksheets("Sheet1").Delete
    Application.DisplayAlerts = True
    
End Sub
Private Sub MakeGraph()
    Dim s As Worksheet
    For Each s In m_Book.Worksheets
        s.Activate
        '�f�[�^�̓����Ă���ŏI�s���擾
        Dim eRow As Long
        eRow = s.Range("A1").SpecialCells(xlLastCell).Row
        
        '�O���t�쐬(X��="No.")
        Dim Rng As String
        s.Shapes.AddChart2(240, xlXYScatterLines).Select
        Rng = "$K$1:$K$" & eRow & ", " & "$O$1:$P$" & eRow
        Call SetChartConf(s, ActiveChart, Rng, 1, "B2")

        '�O���t�쐬(X��="���t")
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
        '�f�[�^�I��
        .SetSourceData Source:=Range(Rng)
        
        '�ė���O���t�̉��ɕ\��
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        '�O���t�̃X�^�C���ύX(�����Ȋ����̃X�^�C��)
        .ClearToMatchStyle
        .ChartStyle = 240
        
        '�`���[�g�̐F�I��
        .ChartColor = 10
        
        '�^�C�g�����V�[�g���ɕύX
        .ChartTitle.Text = s.Name
        
    End With
        
    '�V�[�g�̏ꏊ��ύX
    s.ChartObjects("�O���t " & ChartNo).Left = s.Range(ChartPos).Left
    s.ChartObjects("�O���t " & ChartNo).Top = s.Range(ChartPos).Top
    
    'ChartNo��2�̏ꍇ���������̕\������t�ɂ���B
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
    
    'Sheet����csv�̃t�@�C�����ɕύX
    s.Name = Replace(FileName, ".csv", "")
    
    '�܂��͕\���\��
    s.Cells(1, 11).Value = "No."
    s.Cells(1, 12).Value = "���t"
    s.Cells(1, 13).Value = "���v (�~)"
    s.Cells(1, 14).Value = "���v (pip)"
    s.Cells(1, 15).Value = "�ݐϑ��v (�~)"
    s.Cells(1, 16).Value = "�ݐϑ��v (pip)"
    
    '�l����
    s.Range("K2:P" & DataUB + 2) = Table

    '���������Ă���
    s.Range("K1:P" & DataUB + 2).Borders.LineStyle = xlContinuous
    
    '�I�[�g�t�B�b�g���Ă���
    s.Columns("K:P").AutoFit
        
    '�\��ɐF��h���Ă���
    s.Range("K1:P1").Interior.ColorIndex = 37 '(�Â߂̐��F)
    
End Sub