Attribute VB_Name = "memo"
Option Explicit

Private Sub ArrTset()
    Dim a As Variant
    
    
    ReDim a(1, 2)
    
    a(0, 1) = "aa"
    
    Debug.Print a(0, 1)
    
    
End Sub
Private Sub t2()
    Dim a() As String
    ReDim a(0, 1)
    ReDim Preserve a(67, 1)
    a = tf(a)
    
End Sub
Private Function tf(ByRef a() As String)
    ReDim Preserve a(0, 1)
    tf = a
End Function
