Attribute VB_Name = "modUtils"
Option Explicit

Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_SETREDRAW As Long = &HB

 
Function MakeDWord(ByVal wLo As Integer, ByVal wHi As Integer) As Long
    MakeDWord = (wHi * 65536) + (wLo And &HFFFF&)
End Function

Function LoWord(ByVal dw As Long) As Integer
    If dw And &H8000& Then
        LoWord = dw Or &HFFFF0000
    Else
        LoWord = dw And &HFFFF&
    End If
End Function

Public Function Random(ByVal lLo As Long, ByVal lHi As Long) As Long

    Random = Int(lLo + (Rnd * (lHi - lLo + 1)))

End Function
Public Sub ClearGrid(ByRef grd As MSFlexGrid)

    Dim lRows As Long
    Dim lCols As Long
    
    With grd
        For lRows = .FixedRows To .Rows - 1
            For lCols = .FixedCols To .Cols - 1
                .TextMatrix(lRows, lCols) = vbNullString
                .ColData(lCols) = 0
            Next lCols
            .RowData(lRows) = 0
        Next lRows
    End With

End Sub

Public Function Among(ByVal vItem As Variant, ParamArray avCompare() As Variant) As Boolean

    Dim i As Long
    
    Debug.Assert Not IsMissing(avCompare)
    
    Among = True
    
    For i = LBound(avCompare) To UBound(avCompare)
        If vItem = avCompare(i) Then Exit Function
    Next i
    
    Among = False
    
End Function

Public Function IsArrayValid(ByRef vArray As Variant) As Boolean

    Debug.Assert IsArray(vArray)
    
    On Error Resume Next
    
    Dim lTest As Long
    
    lTest = LBound(vArray)
    If VarType(vArray) And vbObject Then
        If Err = 0 Then
            IsArrayValid = Not (lTest = 0 And UBound(vArray) = -1)
        Else
            IsArrayValid = False
        End If
    Else
        On Error Resume Next
        IsArrayValid = (Err = 0)
    End If
    
End Function

Public Sub ListSelect(ByVal ctl As Control, ByVal lID As Long)

    Dim i As Long
    
    Debug.Assert Among(TypeName(ctl), "ListBox", "ComboBox")
    
    For i = 0 To ctl.ListCount - 1
        If ctl.ItemData(i) = lID Then
            ctl.ListIndex = i
            Exit Sub
        End If
    Next i
    
    'Item not found; deselect list
    ctl.ListIndex = -1
    
End Sub




Public Property Get Missing(Optional vMissing As Variant) As Variant

    Missing = vMissing
    
End Property

Public Sub SetRedraw(ByVal ctl As Control, ByVal bRedraw As Boolean)

    SendMessage ctl.hWnd, WM_SETREDRAW, CLng(Abs(bRedraw)), 0&
    
End Sub

Public Sub ValidateError(ByVal ctl As Control, ByVal sMsg As String)

    MsgBox sMsg, vbExclamation
    
    On Error Resume Next
    ctl.SetFocus
    
End Sub


