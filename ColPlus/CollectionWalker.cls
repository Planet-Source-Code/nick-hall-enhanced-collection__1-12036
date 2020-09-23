VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCollectionWalker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IVariantWalker

Private mauListItems() As ColListItem
Private mlCurrIndex As Long
Private mlFirstItem As Long
Private mlDirection As CollectionEnumDirection
Private mCol As CCollectionPlus


'****************************************
'Procedure:     CCollectionWalker.Attach
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Attach the enumerating collection
'****************************************
'Revisions:
'****************************************
Public Sub Attach(ByVal oCol As CCollectionPlus)

    Set mCol = oCol
    With oCol
        mauListItems = .ListItems
        mlDirection = .EnumDirection
        If mlDirection = colEnumDirBackward Then
            mlFirstItem = .LastItem
        Else
            mlFirstItem = .FirstItem
        End If
    End With
    mlCurrIndex = mlFirstItem
    
End Sub



'****************************************
'Procedure:     CCollectionWalker.CurrIndex
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Set the current item in the enumeration
'****************************************
'Revisions:
'****************************************
Friend Property Let CurrIndex(ByVal NewIndex As Long)

    mlCurrIndex = NewIndex
    
End Property


'****************************************
'Procedure:     CCollectionWalker.CurrIndex
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Returns index of current item in enumeration
'****************************************
'Revisions:
'****************************************
Friend Property Get CurrIndex() As Long

    CurrIndex = mlCurrIndex
    
End Property


'****************************************
'Procedure:     CCollectionWalker.NewEnum
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Use CEnumVariant to return an IEnumVariant
'interface to the client
'****************************************
'Revisions:
'****************************************
Public Function NewEnum() As IEnumVARIANT

    Dim oEnum As CEnumVariant
    
    Set oEnum = New CEnumVariant
    oEnum.Attach Me
    Set NewEnum = oEnum
    
End Function

Private Sub Class_Terminate()

    Set mCol = Nothing
    Erase mauListItems
    
End Sub


Private Function IVariantWalker_Clone() As stdole.IEnumVARIANT

    Dim oEnum As CCollectionWalker
    
    Set oEnum = New CCollectionWalker
    
    With oEnum
        .Attach mCol
        .CurrIndex = mlCurrIndex
    End With
    
    Set IVariantWalker_Clone = oEnum.NewEnum()
    
End Function

Private Function IVariantWalker_More(v As Variant) As Boolean

    If mlCurrIndex <> 0 Then
        With mauListItems(mlCurrIndex)
            If .PreviousItem <> -1 Then
                VariantCopy v, .Value
                If mlDirection = colEnumDirBackward Then
                    mlCurrIndex = .PreviousItem
                Else
                    mlCurrIndex = .NextItem
                End If
                IVariantWalker_More = True
            End If
        End With
    End If
        
End Function


Private Sub IVariantWalker_Reset()

    mlCurrIndex = mlFirstItem
    
End Sub


Private Sub IVariantWalker_Skip(ByVal c As Long)

    Dim i As Long
    
    For i = 1 To c
        If mlCurrIndex = 0 Then Err.Raise S_FALSE
        If mauListItems(mlCurrIndex).PreviousItem = -1 Then Err.Raise S_FALSE
        If mlDirection = colEnumDirBackward Then
            mlCurrIndex = mauListItems(mlCurrIndex).PreviousItem
        Else
            mlCurrIndex = mauListItems(mlCurrIndex).NextItem
        End If
    Next i
        
End Sub


