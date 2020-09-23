VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Demo class for ColDemo project "
Option Explicit

'This is a demo class that implements the ICollectionPlusItem interface to
'demonstrate how to use a class with the CollectionPlus's Sort and Find routines

Implements ICollectionPlusItem

' Surname of person
Private msSurname As String ' Added by: Nick Hall on 07/09/2000 10:47:14 am

' First name of person
Private msFirstName As String ' Added by: Nick Hall on 07/09/2000 10:49:15 am

' Address of person
Private msAddress As String ' Added by: Nick Hall on 07/09/2000 10:49:28 am

Public Property Get Address() As String
Attribute Address.VB_Description = "Address of person"

    Address = msAddress

End Property
Public Property Let Address(ByVal NewAddress As String)

    msAddress = NewAddress

End Property

Public Property Get FirstName() As String
Attribute FirstName.VB_Description = "Person's first name"

    FirstName = msFirstName

End Property
Public Property Let FirstName(ByVal NewFirstName As String)

    msFirstName = NewFirstName

End Property

Public Property Get Name() As String
Attribute Name.VB_Description = "Displayed name of person"
Attribute Name.VB_UserMemId = 0

    Name = msSurname & ", " & msFirstName
    
End Property

Public Property Get Surname() As String
Attribute Surname.VB_Description = "Person's surname"

    Surname = msSurname

End Property
Public Property Let Surname(ByVal NewSurname As String)

    msSurname = NewSurname

End Property



Private Function ICollectionPlusItem_Compare(ByVal Item As ICollectionPlusItem, Optional ByVal CompareMethod As VbCompareMethod = 1&) As Integer

    Dim oCompPerson As CPerson
    Dim iResult As Integer
    
    'Get a reference we can work with.  If this fails then we can't compare the items anyway
    Set oCompPerson = Item
    
    'Compare by surname, first name then address
    iResult = StrComp(msSurname, oCompPerson.Surname, CompareMethod)
    If iResult = 0 Then iResult = StrComp(msFirstName, oCompPerson.FirstName, CompareMethod)
    If iResult = 0 Then iResult = StrComp(msAddress, oCompPerson.Address, CompareMethod)
    
    ICollectionPlusItem_Compare = iResult
    
End Function

Private Property Set ICollectionPlusItem_Contents(ByVal RHS As Variant)

    'Not implemented for this class
    Debug.Assert False
    
End Property

Private Property Let ICollectionPlusItem_Contents(ByVal RHS As Variant)

    Dim pb As New PropertyBag
    
    With pb
        .Contents = RHS
        msSurname = .ReadProperty("Surname", vbNullString)
        msFirstName = .ReadProperty("FirstName", vbNullString)
        msAddress = .ReadProperty("Address", vbNullString)
    End With

End Property

Private Property Get ICollectionPlusItem_Contents() As Variant

    Dim pb As New PropertyBag
    
    With pb
        .WriteProperty "Surname", msSurname, vbNullString
        .WriteProperty "FirstName", msFirstName, vbNullString
        .WriteProperty "Address", msAddress, vbNullString
        
        ICollectionPlusItem_Contents = .Contents
    End With
    
End Property


Private Function ICollectionPlusItem_IsLike(LikeArray() As Variant) As Boolean

    Dim i As Long
    Dim lCount As Long
    Dim sItem As String
    Dim sCompare As String
    Dim lUpper As Long
    Dim bLike As Boolean
    
    On Error Resume Next
    lUpper = UBound(LikeArray)
    If Err <> 0 Then Exit Function
    
    'We're expecting a maximum of three arguments in this order -
    'Surname, First name, Address.  Arguments that are blank ("") or
    'missing will be ignored
    
    For i = LBound(LikeArray) To lUpper
        If lCount = 3 Then Exit For
        lCount = lCount + 1
        
        sItem = UCase$(LikeArray(i))
        
        Select Case lCount
        Case 1
            sCompare = UCase$(msSurname)
            
        Case 2
            sCompare = UCase$(msFirstName)
            
        Case 3
            sCompare = UCase$(msAddress)
        End Select
        
        If LenB(sItem) > 0 Then
            If Not (sCompare Like sItem) Then
                bLike = False
                Exit For
            Else
                bLike = True
            End If
        End If
    Next i
    
    ICollectionPlusItem_IsLike = bLike
            
End Function

Private Property Get ICollectionPlusItem_Value() As String

    ICollectionPlusItem_Value = Me.Name
    
End Property


