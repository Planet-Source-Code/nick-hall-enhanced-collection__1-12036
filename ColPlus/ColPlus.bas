Attribute VB_Name = "modColPlus"
Option Explicit

'General purpose BAS module for constants/types used within the component

'Note: technique for using UDT's to create linked lists was taken from
'Francesco Balena's article "Speed up your Apps with Data structures" in
'the May 2000 issue of VBPJ
Public GCasts As New GCasts     'Globally accesible casting object

'Collection Errors
Public Enum EErrorCollection
    EcolBase = 8000
    EcolItemNotFound            'Specified index does not exist in the collection
    EcolUnableChangeHashSize    'Cannot change hash size when there are items in the collection
    EcolFailSave                'Could not save the collection to the specified file
    EcolFailLoad                'Could not load the collection from the specified file
    EcolUnableChangeMatchCase   'Can't change a collection to case senstive if there are items in the collection
    EcolInvalidObject           'In order to save, all objects in the collection must support ICollectionPlusItem
End Enum

'Enum to determine where new item is placed
Public Enum ColAddType
    colAddDefault = 0           'Default behaviour (item is added to end)
    colAddBefore = 1            'Item is added before the item specified by Before
    colAddAfter = 2             'Item is added after the item specified by After
    colAddIllegal = 3           'The add function specifies both Before and After parameters - this is illegal
End Enum

Public Type HashItem
    Key As String               'String key used to identify this item
    Index As Long               'Reference to a ColListItem in a separate array
    NextItem As Long            'Next index to a HashItem in a separate array
End Type

'ListItem (Variant)
Public Type ColListItem
    Key As String               'Key associated with item (can be null)
    Tag As String               'User-defined data about the item
    PreviousItem As Long        'Reference to the previous ColListItem
    NextItem As Long            'Reference to the next ColListItem
    Value As Variant            'Anything the user wishes to store in the list
End Type

'ListItem (String)
Public Type ColListItemString
    Key As String               'Key associated with item (can be null)
    Tag As String               'User-defined data about the item
    PreviousItem As Long        'Reference to the previous ListItem
    NextItem As Long            'Reference to the next ListItem
    Value As String             'Anything the user wishes to store in the list
End Type

'Types defined for saving collections

Public Enum ColItemType
    colTypeOther
    colTypeColItem
    colTypePersistObject
End Enum
    
'ColFileHeader stores the setup for the collection being saved.  Also stores number
'of items that were saved
Public Type ColFileHeader
    EnumDirection As Long
    MatchCase As Boolean
    HashSize As Long
    HashLookupChunkSize As Long
    HashLookupInitialSize As Long
    ItemChunkSize As Long
    ItemInitialSize As Long
    Count As Long
End Type

'One of these types is stored for each item in the collection.  IsColItem indicates
'that Item contains the Contents from the ICollectionPlusItem interface
Public Type ColFileItem
    ItemType As Long
    Tag As String
    Key As Variant
    Item As Variant
End Type

'Range constants
Public Const MINLONG As Long = -2147483648#
Public Const MAXLONG As Long = 2147483647

'VB-defined errors
Public Const ERR_INVALID_ARG As Long = 5
Public Const ERR_TYPE_MISMATCH As Long = 13
Public Const ERR_FILENOTFOUND As Long = 53
Public Const ERR_FILEEXISTS As Long = 58
Public Const ERR_DISKFULL As Long = 61
Public Const ERR_PERMISSIONDENIED As Long = 70
Public Const ERR_PATHFILEACCESS As Long = 75
Public Const ERR_PATHNOTFOUND As Long = 76
Public Const ERR_INVALIDFRIEND As Long = 97
Public Const ERR_INVALIDFILEFMT As Long = 321
Public Const ERR_DATAVALUEMISSING As Long = 327
Public Const ERR_INVALID_PROPERTY As Long = 380
Public Const ERR_OBJECTREQUIRED As Long = 424
Public Const ERR_KEY_EXISTS As Long = 457

'COM errors
Public Const E_NOTIMPL As Long = &H80004001

'Persistable properties
Public Const PROP_ITEM = "Item"
Public Const PROP_KEY = "Key"
Public Const PROP_COUNT = "Count"
Public Const PROP_ENUMDIR = "EnumDirection"
Public Const PROP_MATCHCASE = "MatchCase"
Public Const PROP_TAG = "Tag"
Public Const PROP_HASH_SIZE = "HashSize"
Public Const PROP_HASH_LOOKUP_CHUNK_SIZE = "HashLookupChunkSize"
Public Const PROP_HASH_LOOKUP_INIT_SIZE = "HashLookupInitialSize"
Public Const PROP_ITEM_CHUNK_SIZE = "ItemChunkSize"
Public Const PROP_ITEM_INIT_SIZE = "ItemInitialSize"

'Miscellaneous constants
Private Const COLPLUS_MSG_REPLACE = "%s"

Public Sub ErrRaise(ByVal lError As Long, ByVal sErrorSource As String, ParamArray MsgReplace() As Variant)

    Dim sError As String
    
    Debug.Assert lError <> EcolBase
    
    If lError > EcolBase Then
        'This error is defined by this class
        sError = LoadResString(lError - EcolBase + 100)
        
        If Not IsMissing(MsgReplace) Then
            Dim i As Long
            Dim lCount As Long: lCount = 1
            
            For i = LBound(MsgReplace) To UBound(MsgReplace)
                sError = Replace(Expression:=sError, Find:=COLPLUS_MSG_REPLACE & CStr(lCount), Replace:=MsgReplace(i))
                lCount = lCount + 1
            Next i
        End If
    Else
        'This is a standard error - get the standard message
        sError = Error$(lError)
    End If
    
    Err.Raise Number:=lError Or vbObjectError, Source:=sErrorSource, Description:=sError
    
End Sub

Public Sub Main()
    
    Randomize Timer
    
End Sub





'****************************************
'Procedure:     modColPlus.WriteColHeader
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          01 September 2000
'****************************************
'Description:   Write a ColFileHeader structure based on the collection
'oColPlus to the file specified by lFileNum
'****************************************
'Revisions:
'****************************************
Public Sub WriteColHeader(ByVal lFileNum As Long, ByVal oColPlus As CCollectionPlus)

    Dim udtHeader As ColFileHeader
    
    Debug.Assert lFileNum > 0
    
    'Store main data
    With oColPlus
        udtHeader.EnumDirection = .EnumDirection
        udtHeader.MatchCase = .MatchCase
        udtHeader.Count = .Count
    End With
    
    'Store settings data
    With GCasts.ICollectionPlusSettings(oColPlus)
        udtHeader.HashSize = .HashSize
        udtHeader.HashLookupInitialSize = .HashLookupInitialSize
        udtHeader.HashLookupChunkSize = .HashLookupChunkSize
        udtHeader.ItemInitialSize = .ItemInitialSize
        udtHeader.ItemChunkSize = .ItemChunkSize
    End With
    
    'Write header information to file
    Put #lFileNum, , udtHeader
    
End Sub

'****************************************
'Procedure:     modColPlus.SortCollection
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          06 July 2000
'****************************************
'Description:   Sort the items represented by the index in alIndex
'according to rules defined by oSortHelper.  Adapted from the
'SortArrayRec procedure in Bruce Mckinney's Hardcore VB
'****************************************
'Revisions:
'****************************************
Public Sub SortCollectionIndex(ByRef alIndex() As Long, ByVal oSortHelper As ICollectionPlusSortHelper, Optional ByVal vFirst As Variant, Optional ByVal vLast As Variant)

    Dim lFirst As Long
    Dim lLast As Long
    
    'Check to see if we are in recursion
    If IsMissing(vFirst) Then lFirst = LBound(alIndex) Else lFirst = vFirst
    If IsMissing(vLast) Then lLast = UBound(alIndex) Else lLast = vLast
    
    With oSortHelper
        If lFirst < lLast Then
            If lLast - lFirst = 1 Then
                If .Compare(alIndex(lFirst), alIndex(lLast)) > 0 Then
                    .Swap alIndex(lFirst), alIndex(lLast)
                End If
            Else
                Dim lLo As Long
                Dim lHigh As Long
                
                'Swap random element with the end element
                .Swap alIndex(lLast), alIndex(Random(lFirst, lLast))
                lLo = lFirst
                lHigh = lLast
                Do
                    Do While (lLo < lHigh) And (.Compare(alIndex(lLo), alIndex(lLast)) <= 0)
                        lLo = lLo + 1
                    Loop
                    
                    Do While (lHigh > lLo) And (.Compare(alIndex(lHigh), alIndex(lLast)) >= 0)
                        lHigh = lHigh - 1
                    Loop
                    
                    'If we haven't reached the pivot element then two items are out of order
                    If lLo < lHigh Then .Swap alIndex(lLo), alIndex(lHigh)
                Loop While lLo < lHigh
                
                'Restore the pivot element to its former position
                .Swap alIndex(lLo), alIndex(lLast)
                
                'Call function recursively (smaller sub division first)
                If (lLo - lFirst) < (lLast - lLo) Then
                    SortCollectionIndex alIndex, oSortHelper, lFirst, lLo - 1
                    SortCollectionIndex alIndex, oSortHelper, lLo + 1, lLast
                Else
                    SortCollectionIndex alIndex, oSortHelper, lLo + 1, lLast
                    SortCollectionIndex alIndex, oSortHelper, lFirst, lLo - 1
                End If
            End If
        End If
    End With
    
End Sub

'****************************************
'Procedure:     modColPlus.WriteColItem
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          02 October 2000
'****************************************
'Description:   Writes a ColFileItem structure containing the
'passed parameters to the specified file
'****************************************
'Revisions:
'****************************************
Public Sub WriteColItem(ByVal lFileNum As Long, ByRef vItem As Variant, ByRef vKey As Variant, ByRef sTag As String)

    On Error GoTo WriteColItemErr

    Dim udtItem As ColFileItem
    
    Debug.Assert lFileNum > 0
    
    With udtItem
        If IsObject(vItem) Then
            'Object MUST support ICollectionPlusItem interface or be persistable
            Dim pb As New PropertyBag
            
            On Error Resume Next
            
            'Try persistance first
            pb.WriteProperty PROP_ITEM, vItem
            
            If Err Then
                'Object DOES NOT support persistance; try ICollectionPlusItem interface
                Dim oItem As ICollectionPlusItem
                
                Err.Clear
                Set oItem = vItem
                If Err Then
                    On Error GoTo 0
                    'Can't do anything with this object; have to raise an error
                    Err.Raise EcolInvalidObject
                Else
                    'Object supports ICollectionPlusItem interface
                    .ItemType = colTypeColItem
                    VariantMove .Item, oItem.Contents
                End If
            Else
                'Object DOES support persistance
                .ItemType = colTypePersistObject
                VariantMove .Item, pb.Contents
            End If
            On Error GoTo WriteColItemErr
        Else
            'Non-object variant type
            .ItemType = colTypeOther
            .Item = vItem
        End If
        
        .Key = vKey
        .Tag = sTag
    End With
    
    Put #lFileNum, , udtItem
        
Exit Sub

WriteColItemErr:
    
    'Raise error back to calling collection
    Err.Raise Err.Number

End Sub


