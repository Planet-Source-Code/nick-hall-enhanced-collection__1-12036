Attribute VB_Name = "modUtils"
Option Explicit

'****************************************
'Procedure:     modColPlus.CanConvertToLong
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          28 June 2000
'****************************************
'Description:   Returns true if variant can be converted
'to long, false otherwise
'****************************************
'Revisions:
'****************************************
Public Function CanConvertToLong(ByRef vNumber As Variant) As Boolean

    Dim bNumeric As Boolean
    
    bNumeric = False
    
    'Check to see if we've been passed a valid numeric data type
    Select Case VarType(vNumber)
    Case vbInteger, vbLong, vbByte
        'Can always be converted to a long
        CanConvertToLong = True
        Exit Function
        
    Case vbSingle, vbDouble, vbCurrency, vbDecimal, vbDate
        bNumeric = True
        
    Case vbString, vbObject
        bNumeric = IsNumeric(vNumber)
    
    End Select
    
    If Not bNumeric Then Exit Function
    
    'Check upper and lower limits
    If vNumber < MINLONG Then Exit Function
    If vNumber > MAXLONG Then Exit Function

    CanConvertToLong = True
    
End Function

'****************************************
'Procedure:     modColPlus.HashString
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 November 1999
'****************************************
'Description:   Convert a string key into a representative
'long integer.
'N.B. sKey is passed ByRef for performance reasons only and is not modified
'
'Taken from the HashCode routine in the HashTable class featured
'in Francesco Balena's article "Speed up your Apps with Data structures"
'****************************************
'Revisions:
'****************************************
Public Function HashString(ByRef sKey As String) As Long

    Dim lLastEl As Long, i As Long
    Dim alCodes() As Long
    
    'Determine number of longs needed
    lLastEl = (Len(sKey) - 1) \ 4
    
    'Call SafeArrayCreateVector function alias - this allocates array
    'descriptor and data in one go.
    'This is a bit faster than the equivalent VB line: -
    'Redim Preserve alCodes(0 to lLastEl)
    alCodes = SafeArrayCreateLongVector(vbLong, 0, lLastEl + 1)
    
    ' this also converts from Unicode to ANSI
    CopyMemory alCodes(0), ByVal sKey, Len(sKey)

    ' XOR the ANSI codes of all characters
    For i = 0 To lLastEl
        HashString = HashString Xor alCodes(i)
    Next
    
End Function

Public Function IsArrayValid(ByRef vArray As Variant) As Boolean

    Dim lMin As Long
    
    'Check to make sure this is an array
    Debug.Assert IsArray(vArray)
    
    On Error Resume Next
    lMin = LBound(vArray)
    IsArrayValid = (Err = 0)
    
End Function


'****************************************
'Procedure:     modColPlus.ValidFile
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          01 September 2000
'****************************************
'Description:   Returns true if sFilePath points to
'a valid file, false otherwise
'****************************************
'Revisions:
'****************************************
Public Function ValidFile(ByRef sFilePath As String) As Boolean

    On Error Resume Next
    
    Dim rc As Long
    
    rc = GetFileAttributes(sFilePath)
    ValidFile = (rc <> -1)
    
End Function
'****************************************
'Procedure:     modColPlus.VariantMove
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          31 August 2000
'****************************************
'Description:   Move contents of vSrc into vDest.
'At the end of this function, vDest will contain whatever
'vSrc contained and vSrc will be Empty.
'****************************************
'Revisions:
'****************************************
Public Sub VariantMove(ByRef vDest As Variant, ByRef vSrc As Variant)

    Dim vEmpty As Variant
    
    'Explicitly clear anything currently in vDest
    vDest = Empty
    
    'Copy the contents of vSrc into vDest
    CopyMemory vDest, vSrc, 16
    
    'Clear vSrc - this prevents VB from freeing the contents twice
    CopyMemory vSrc, vEmpty, 16
    
End Sub
'****************************************
'Procedure:     modColPlus.Random
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          13 July 2000
'****************************************
'Description:   Returns a random number between
'lLo and lHi, inclusive
'****************************************
'Revisions:
'****************************************
Public Function Random(ByVal lLo As Long, ByVal lHi As Long) As Long

    Random = Int(lLo + (Rnd * (lHi - lLo + 1)))

End Function

