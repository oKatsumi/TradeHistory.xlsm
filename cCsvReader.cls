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
'�t�H���_�͈ȉ��ɂ܂Ƃ߂Ă���CVS�t�@�C������FX�̎���������������Ă���B
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
    
    '�Ƃ肠����Table�֘A��������
    ReDim m_CsvTables(0)
    ReDim m_CsvTables(0).Table(0, 0)
        
    ReDim m_TotalTable(0, 0)
    
End Sub
'�܂��́ACVS�������Ă���t�H���_�̃p�X���擾
Public Sub GetPath()

    '�t�H���_�I���_�C�A�O����t�H���_�p�X���擾�B�L�����Z������������}�N���I��
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
    '�f�o�b�N�p
    m_FolderPath = "C:\Users\katsumi\LionFXWin\Data\csv"
    
End Sub
'�t�H���_�z����csv�t�@�C������K�v�ȃf�[�^���쐬
Public Sub GetValue()
    '�t�@�C�������擾
    Dim File As String
    Dim ci As Long: ci = 0
    File = Dir(m_FolderPath & "\*csv", vbNormal)
    Do While File <> ""
        ReDim Preserve m_Csv(ci)
        Set m_Csv(ci) = New cCsvVal
        
        '�����Ńt�@�C�����J���Ēl��ݒ�
        Call GetValue2(File, m_Csv(ci))
        ci = ci + 1
        File = Dir()
    Loop

End Sub
Private Sub GetValue2(ByVal File As String, ByRef csv As cCsvVal)
    '�Ƃ肠�����t�@�C�������擾���Ă���
    csv.FileName = File
    
    '�t�@�C�����J���ĕK�v�ȃf�[�^���擾
    Dim Buf As String
    Dim sBuf() As String
    Dim di As Long: di = 0
    
    Open m_FolderPath & "\" & File For Input As #1
                
        '�f�[�^������s�܂�Buf���ړ��B
        Dim n As Long
        For n = 0 To 2: Line Input #1, Buf: Next
        
        Do Until EOF(1)
            Line Input #1, Buf
            sBuf = Split(Buf, ",")
            
            'di��0�ȏ�Ȃ��m_csv�̃f�[�^���C���N�������g
            If di > 0 Then csv.IncArr
            
            '�l���擾
            With csv
                .Day(di) = sBuf(0)
                .Trade(di) = sBuf(16)
                .Trade_pip(di) = sBuf(14)
            End With
            di = di + 1
        Loop
    Close #1
        
    '�Ō�ɔ������v�ݐς��v�Z
    csv.CalcSum
    
End Sub
Public Sub MakeTotal()
    '�e�V�[�g��Data��A������m_Total���쐬����B
    
    '�Ƃ肠�����t�@�C������Total�����Ă����B
    m_Total.FileName = "Total"
    
    Dim ti As Long: ti = 0
    Dim n As Long
    For n = 0 To UBound(m_Csv)
        Dim m As Long
        For m = 0 To m_Csv(n).UB_DataList
            If ti > 0 Then m_Total.IncArr
            '�l���i�[
            With m_Total
                .Day(ti) = m_Csv(n).Day(m)
                .Trade(ti) = m_Csv(n).Trade(m)
                .Trade_pip(ti) = m_Csv(n).Trade_pip(m)
            End With
            ti = ti + 1
        Next
    Next
    
    'Sum���쐬
    m_Total.CalcSum
    
End Sub
Public Sub MakeTables()
    'Csv��Total��2�����z��Ɋi�[����B
    Dim n As Long
    For n = 0 To UBound(m_Csv)
        '�i�[�p��2����������錾
        ReDim Preserve m_CsvTables(n)
        
        'CSV�̒l��2�����z��
        m_CsvTables(n).Table = SetTableVal(m_CsvTables(n).Table, m_Csv(n))
    Next
    
    'Total�̒l��2�����z��
    m_TotalTable = SetTableVal(m_TotalTable, m_Total)
    
End Sub
Private Function SetTableVal(ByRef Table() As Variant, ByRef CsvVal As cCsvVal)
            
    Dim ubArr As Long
    ubArr = CsvVal.UB_DataList
        
    'Table��z��
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