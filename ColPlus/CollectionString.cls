VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCollectionString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Enhanced version of the Collection class, typed to Strings"
'A collection-like class that is typed specifically to strings.
'Also supports the sorting of its members and persisting to a
'PropertyBag
Option Explicit

'Implements a VB-compatible version of the standard collection interface
Implements IVBCollection

'Interface for performing sorts on collection contents
Implements ICollectionPlusSortHelper

'Implement CCollectionPlus interface to make class polymorphic with its "parent" class
Implements CCollectionPlus

'Interface for methods that are not compatible with VB5
Implements ICollectionPlus_VB5

'Interface to enable users to alter internal settings
Implements ICollectionPlusSettings

'***Constants
Private Const DFLT_HASH_SIZE As Long = 2048                     'Size of hash lookup array
Private Const DFLT_HASH_LOOKUP_INITIAL_SIZE As Long = 4096      'Initial size of hash item list
Private Const DFLT_HASH_LOOKUP_CHUNK_SIZE As Long = 1024        'Amount by which the hash array is increased
Private Const DFLT_ITEM_INITIAL_SIZE As Long = 1024             'Initial size of ColListItem array
Private Const DFLT_ITEM_CHUNK_SIZE As Long = 512                'Amount by which the ListItem array is increased
Private Const ERRSOURCE = "CollectionPlus.CCollectionString"    'Constant used for Error source argument

'*** Private members
Private malHashIndex() As Long                                  'Hash value indexes into the hash lookup array
Private mauHashLookup() As HashItem                             'Linked list containing index of ListItem in main array
Private mauListItems() As ColListItemString                     'Main linked list containing actual items that are stored

Private mlFirstItem As Long                                     'Index to the first item in the list
Private mlLastItem As Long                                      'Index to last item in the list
Private mlLastHashIndex As Long                                 'Upper bound of Hash lookup array
Private mlLastListIndex As Long                                 'Upper bound of List Item array
Private mlNextFreeHashItem As Long                              'Index to the next unallocated item in the hash array
Private mlNextFreeListItem As Long                              'Index to the next unallocated item in the list array

'Settings for hash size/list size
Private mlHashSize As Long
Private mlHashLookupInitialSize As Long
Private mlHashLookupChunkSize As Long
Private mlItemInitialSize As Long
Private mlItemChunkSize As Long

' Number of items in collection
Private mlCount As Long ' Added by: Nick Hall on 29/10/1999 9:33:33 am

' Determines whether the collection iterates forward or backwards
Private mlEnumDirection As CollectionEnumDirection ' Added by: Nick Hall on 05/07/2000 10:20:07 am

' Determines whether keys are compared case-sensitive or not
Private mbMatchCase As Boolean ' Added by: Nick Hall on 01/09/2000 2:47:36 pm

Private miAscending As Integer                                  'Determines whether items are sorted ascending or descending
Private mlCompareType As VbCompareMethod                        'Determines how items are compared

'****************************************
'Procedure:     CCollectionString.Move
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          10 October 2000
'****************************************
'Description:   Move an item in the collection
'****************************************
'Revisions:
'****************************************
Public Sub Move(ByRef Index As Variant, Optional ByRef Before As Variant, Optional ByRef After As Variant)
Attribute Move.VB_Description = "Move an item within the collection"

    Dim lCurrIndex As Long
    Dim lCompareIndex As Long
    Dim lAddFlags As ColAddType
    
    'Get index for this item
    lCurrIndex = GetItemIndex(Index)
    
    If Not IsMissing(Before) Then lAddFlags = lAddFlags Or colAddBefore
    If Not IsMissing(After) Then lAddFlags = lAddFlags Or colAddAfter
    
    'In this case, must have either Before or After.  Both or neither is illegal
    If lAddFlags = colAddDefault Or lAddFlags = colAddIllegal Then ErrRaise ERR_INVALID_ARG, ERRSOURCE
    
    'Get index of item to add before or after (generates error if can't be found)
    Select Case lAddFlags
    Case colAddBefore
        lCompareIndex = GetItemIndex(Before)
    Case colAddAfter
        lCompareIndex = GetItemIndex(After)
    End Select
    
    'Can't move an item before or after itself
    If lCompareIndex = lCurrIndex Then Exit Sub
    
    'Dettach item from its existing position
    With mauListItems(lCurrIndex)
        If mlFirstItem = lCurrIndex Then
            mlFirstItem = .NextItem
        Else
            mauListItems(.PreviousItem).NextItem = .NextItem
        End If
        
        If mlLastItem = lCurrIndex Then
            mlLastItem = .PreviousItem
        Else
            mauListItems(.NextItem).PreviousItem = .PreviousItem
        End If
    End With
    
    'Reattach in new position
    With mauListItems(lCompareIndex)
        Select Case lAddFlags
        Case colAddBefore
            If mlFirstItem = lCompareIndex Then
                'Moved item will be first
                mlFirstItem = lCurrIndex
            Else
                mauListItems(.PreviousItem).NextItem = lCurrIndex
            End If
            mauListItems(lCurrIndex).PreviousItem = .PreviousItem
            mauListItems(lCurrIndex).NextItem = lCompareIndex
            .PreviousItem = lCurrIndex
        
        Case colAddAfter
            If mlLastItem = lCompareIndex Then
                'Moved item will be last
                mlLastItem = lCurrIndex
            Else
                mauListItems(.NextItem).PreviousItem = lCurrIndex
            End If
            mauListItems(lCurrIndex).PreviousItem = lCompareIndex
            mauListItems(lCurrIndex).NextItem = .NextItem
            .NextItem = lCurrIndex
        End Select
    End With
        
End Sub

'****************************************
'Procedure:     CCollectionString.CreateNewHashIndex
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          26 September 2000
'****************************************
'Description:   Checks to make sure that no item already exists with
'specified key, then creates a hash lookup item based on the key
'to point at the actual item
'****************************************
'Revisions:
'****************************************
Private Function CreateNewHashIndex(ByVal sKey As String) As Long

    Dim lLastIndex As Long
    Dim lHashValue As Long
    
    'Upper case incoming string if we are comparing case-insensitive
    If Not mbMatchCase Then sKey = UCase$(sKey)
    
    lHashValue = HashString(sKey) Mod mlHashSize
    
    'Attempt to find this index in the hash lookup table
    'If it exists we need to raise an error
    CreateNewHashIndex = malHashIndex(lHashValue)
    Do Until CreateNewHashIndex = 0
        If mauHashLookup(CreateNewHashIndex).Key = sKey Then ErrRaise ERR_KEY_EXISTS, ERRSOURCE
        lLastIndex = CreateNewHashIndex
        CreateNewHashIndex = mauHashLookup(CreateNewHashIndex).NextItem
    Loop
    
    'Make sure enough space is allocated
    If mlNextFreeHashItem = 0 Then ExpandHashLookup
    CreateNewHashIndex = mlNextFreeHashItem
    With mauHashLookup(CreateNewHashIndex)
        .Key = sKey
    
        'Get next free item and signal end of list
        mlNextFreeHashItem = .NextItem
        .NextItem = 0
        
        If lLastIndex = 0 Then
            'Update the hash array to point to this item
            malHashIndex(lHashValue) = CreateNewHashIndex
        Else
            'Link the former last item to the new one
            mauHashLookup(lLastIndex).NextItem = CreateNewHashIndex
        End If
    End With
    
End Function


'****************************************
'Procedure:     CCollectionString.GetStringIndex
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          26 September 2000
'****************************************
'Description:   Return an index into the Item array, given
'its string key
'****************************************
'Revisions:
'****************************************
Private Function GetStringIndex(ByRef sKey As String) As Long

    Dim lHashValue As Long
    
    lHashValue = HashString(sKey) Mod mlHashSize
    
    'Attempt to find this index in the hash lookup table
    GetStringIndex = malHashIndex(lHashValue)
    Do Until GetStringIndex = 0
        If mauHashLookup(GetStringIndex).Key = sKey Then Exit Do
        GetStringIndex = mauHashLookup(GetStringIndex).NextItem
    Loop

    If GetStringIndex = 0 Then
        ErrRaise EcolItemNotFound, ERRSOURCE, sKey
    Else
        GetStringIndex = mauHashLookup(GetStringIndex).Index
    End If
    
End Function


'****************************************
'Procedure:     CCollectionString.MatchCase
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Detemines whether the collection matches keys
'case-insenstive (Default, False) or case-sensitive (True)
'****************************************
'Revisions:
'****************************************
Public Property Get MatchCase() As Boolean
Attribute MatchCase.VB_Description = "Determines whether keys are compared case-insensitive (false, default value) or case-sensitive (true)"

    MatchCase = mbMatchCase

End Property

'****************************************
'Procedure:     CCollectionString.MatchCase
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Change the way the collection compares keys
'****************************************
'Revisions:
'****************************************
Public Property Let MatchCase(ByVal NewMatchCase As Boolean)

    If NewMatchCase <> mbMatchCase Then
        If mlCount > 0 Then ErrRaise EcolUnableChangeMatchCase, ERRSOURCE
        mbMatchCase = NewMatchCase
    End If

End Property

'****************************************
'Procedure:     CCollectionString.Save
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          01 September 2000
'****************************************
'Description:   Save the contents of the collection to the file
'sepecified in FileName.  If OverwriteExisting is true, the file
'will be erased if it already exists; otherwise an error will be raised
'****************************************
'Revisions:
'****************************************
Public Sub Save(ByRef FileName As String, Optional ByVal OverwriteExisting As Boolean = False)
Attribute Save.VB_Description = "Save the contents of the collection to a file"

    On Error GoTo SaveErr
    
    Dim i As Long
    Dim lFileNum As Long
    Dim lCurrIndex As Long
    Dim bFileOpen As Boolean
    Dim vKey As Variant
    
    If ValidFile(FileName) Then
        If OverwriteExisting Then
            Kill FileName
        Else
            Err.Raise ERR_FILEEXISTS
        End If
    End If
    
    lFileNum = FreeFile
    Open FileName For Binary Access Write As lFileNum
    bFileOpen = True
    
    'Write header information to file
    WriteColHeader lFileNum, Me
    
    lCurrIndex = mlFirstItem
    For i = 1 To mlCount
        With mauListItems(lCurrIndex)
            If KeyExists(lCurrIndex) Then
                vKey = .Key
            Else
                vKey = Null
            End If
            
            WriteColItem lFileNum, .Value, vKey, .Tag
            lCurrIndex = .NextItem
        End With
    Next i
    
    bFileOpen = False
    Close #lFileNum
    
Exit Sub

SaveErr:
    If bFileOpen Then Close #lFileNum
    
    Select Case Err.Number
    Case ERR_DISKFULL, ERR_PATHFILEACCESS, ERR_PERMISSIONDENIED, ERR_PATHNOTFOUND
        ErrRaise EcolFailSave, ERRSOURCE, FileName, Err.Description
        
    Case Else
        ErrRaise Err.Number, ERRSOURCE
        
    End Select
    
End Sub

'****************************************
'Procedure:     CCollectionString.Load
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          01 September 2000
'****************************************
'Description:   Load the contents of FileName into the current
'collection.  If IgnoreDuplicates is False, it will replace any
'items in the current collection that have the same key as items
'being loaded
'****************************************
'Revisions:
'****************************************
Public Sub Load(ByRef FileName As String, Optional ByVal IgnoreDuplicates As Boolean = True)
Attribute Load.VB_Description = "Merge the previously saved collecton into the current collection"

    On Error GoTo LoadErr

    Dim bFileOpen As Boolean
    Dim i As Long
    Dim lFileNum As Long
    Dim lCount As Long
    Dim udtHeader As ColFileHeader
    Dim udtItem As ColFileItem
    
    'Attempt to open file
    lFileNum = FreeFile
    Open FileName For Binary Access Read As lFileNum
    bFileOpen = True
    
    'Read in header
    Get #lFileNum, , udtHeader
    With udtHeader
        Me.EnumDirection = .EnumDirection
        
        'Only change if there are no items in the collection
        If mlCount = 0 Then
            Me.MatchCase = .MatchCase
            GCasts.ICollectionPlusSettings(Me).ChangeSettings .HashSize, .HashLookupInitialSize, .HashLookupChunkSize, .ItemInitialSize, .ItemChunkSize
        End If
    
        lCount = .Count
    End With
    
    For i = 1 To lCount
        'Get next item in file
        Get #lFileNum, , udtItem
        With udtItem
            If IsNull(.Key) Then
                Me.Add Value:=CStr(.Item), Tag:=.Tag
            Else
                Me.Add Value:=CStr(.Item), Key:=.Key, Tag:=.Tag
            End If
        End With
    Next i

    Close #lFileNum
    
Exit Sub

LoadErr:
    If bFileOpen Then Close #lFileNum
    
    Dim lErrNumber As Long
    
    lErrNumber = Err.Number
    If lErrNumber And vbObjectError Then lErrNumber = lErrNumber And (Not vbObjectError)
    
    Select Case lErrNumber
    Case ERR_KEY_EXISTS
        If Not IgnoreDuplicates Then
            'Overwrite existing item
            With udtItem
                Me.Item(.Key) = .Item
            End With
        End If
        Resume Next
        
    Case ERR_FILENOTFOUND, ERR_PATHFILEACCESS, ERR_PERMISSIONDENIED, ERR_PATHNOTFOUND
        ErrRaise EcolFailLoad, ERRSOURCE, FileName, Err.Description
    
    Case ERR_TYPE_MISMATCH
        'This almost certainly means an invalid or corrupted file
        ErrRaise ERR_INVALIDFILEFMT, ERRSOURCE
    
    Case Else
        ErrRaise Err.Number, ERRSOURCE
    End Select

End Sub



'****************************************
'Procedure:     CCollectionString.Tag
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Return an item's tag
'****************************************
'Revisions:
'****************************************
Public Property Get Tag(ByRef Index As Variant) As String
Attribute Tag.VB_Description = "User defined data associated with item"

    Tag = mauListItems(GetItemIndex(Index)).Tag
    
End Property


'****************************************
'Procedure:     CCollectionString.Tag
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Assign an item's tag
'****************************************
'Revisions:
'****************************************
Public Property Let Tag(ByRef Index As Variant, ByRef NewTag As String)

    mauListItems(GetItemIndex(Index)).Tag = NewTag
    
End Property

Private Function CCollectionPlus_FindItems(ParamArray FindParams() As Variant) As CCollectionPlus

    'Not implemented for this class
    Err.Raise E_NOTIMPL
    
End Function

Private Sub CCollectionPlus_Load(FileName As String, Optional ByVal IgnoreDuplicates As Boolean = True, Optional ByVal ColSite As ICollectionPlusSite)

    'Note: ColSite argument is not relevant to this class, just ignore it
    Me.Load FileName, IgnoreDuplicates
    
End Sub

Private Property Let CCollectionPlus_MatchCase(ByVal RHS As Boolean)

    Me.MatchCase = RHS
    
End Property

Private Property Get CCollectionPlus_MatchCase() As Boolean

    CCollectionPlus_MatchCase = mbMatchCase
    
End Property

Private Sub CCollectionPlus_Move(Index As Variant, Optional Before As Variant, Optional After As Variant)

    Me.Move Index, Before, After
    
End Sub

Private Sub CCollectionPlus_Save(FileName As String, Optional ByVal OverwriteExisting As Boolean = False)

    Me.Save FileName, OverwriteExisting
    
End Sub



Private Property Let CCollectionPlus_Tag(Index As Variant, RHS As String)

    Me.Tag(Index) = RHS
    
End Property

Private Property Get CCollectionPlus_Tag(Index As Variant) As String
    
    CCollectionPlus_Tag = Me.Tag(Index)
    
End Property

'****************************************
'Procedure:     CCollectionString.Class_Initialize
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          06 September 2000
'****************************************
'Description:   This code should really go in the
'Class_InitProperties event but it seems some clients
'do not fire this event when creating the class.
'****************************************
'Revisions:
'****************************************
Private Sub Class_Initialize()

    'Give settings default values
    mlHashSize = DFLT_HASH_SIZE
    mlHashLookupChunkSize = DFLT_HASH_LOOKUP_CHUNK_SIZE
    mlHashLookupInitialSize = DFLT_HASH_LOOKUP_INITIAL_SIZE
    mlItemChunkSize = DFLT_ITEM_CHUNK_SIZE
    mlItemInitialSize = DFLT_ITEM_INITIAL_SIZE
    mbMatchCase = False
    
    Clear

End Sub

'****************************************
'Procedure:     CCollectionString.ICollectionPlusSettings_ChangeSettings
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          12 July 2000
'****************************************
'Description:   Change multiple settings at once.  Not all settings
'take effect immediately.  HashSize cannot be changed when there
'are items in the collection and an error is raised if this
'is attempted
'****************************************
'Revisions:
'****************************************
Private Sub ICollectionPlusSettings_ChangeSettings(Optional ByVal HashSize As Long = -1&, Optional ByVal HashLookupInitialSize As Long = -1&, Optional ByVal HashLookupChunkSize As Long = -1&, Optional ByVal ItemInitialSize As Long = -1&, Optional ByVal ItemChunkSize As Long = -1&)

    Dim bResize As Boolean
    
    If HashSize <> -1 Then
        If HashSize <> mlHashSize Then
            If mlCount > 0 Then ErrRaise EcolUnableChangeHashSize, ERRSOURCE
            If HashSize <= 0 Then ErrRaise ERR_INVALID_ARG, ERRSOURCE
            mlHashSize = HashSize
            bResize = True
        End If
    End If
    
    If HashLookupInitialSize <> -1 Then
        If HashLookupInitialSize <> mlHashLookupInitialSize Then
            If HashLookupInitialSize <= 0 Then ErrRaise ERR_INVALID_ARG, ERRSOURCE
            mlHashLookupInitialSize = HashLookupInitialSize
            bResize = True
        End If
    End If
    
    If HashLookupChunkSize <> -1 Then
        If HashLookupChunkSize <> mlHashLookupChunkSize Then
            If HashLookupChunkSize <= 0 Then ErrRaise ERR_INVALID_ARG, ERRSOURCE
            mlHashLookupChunkSize = HashLookupChunkSize
        End If
    End If
    
    If ItemChunkSize <> -1 Then
        If ItemChunkSize <> mlItemChunkSize Then
            If ItemChunkSize <= 0 Then ErrRaise ERR_INVALID_ARG, ERRSOURCE
            mlItemChunkSize = ItemChunkSize
        End If
    End If
    
    If ItemInitialSize <> -1 Then
        If ItemInitialSize <> mlItemInitialSize Then
            If ItemInitialSize <= 0 Then ErrRaise ERR_INVALID_ARG, ERRSOURCE
            mlItemInitialSize = ItemInitialSize
            bResize = True
        End If
    End If
    
    'Apply new settings immediately if there are no items in collection
    If bResize And mlCount = 0 Then Clear
    
End Sub

Private Property Let ICollectionPlusSettings_HashLookupChunkSize(ByVal RHS As Long)

    If RHS <> mlHashLookupChunkSize Then
        If RHS <= 0 Then ErrRaise ERR_INVALID_ARG, ERRSOURCE
        mlHashLookupChunkSize = RHS
    End If
    
End Property

Private Property Get ICollectionPlusSettings_HashLookupChunkSize() As Long

    ICollectionPlusSettings_HashLookupChunkSize = mlHashLookupChunkSize
    
End Property


Private Property Let ICollectionPlusSettings_HashLookupInitialSize(ByVal RHS As Long)

    If RHS <> mlHashLookupInitialSize Then
        If RHS <= 0 Then ErrRaise ERR_INVALID_ARG, ERRSOURCE
        mlHashLookupInitialSize = RHS
        If mlCount = 0 Then Clear
    End If
    
End Property

Private Property Get ICollectionPlusSettings_HashLookupInitialSize() As Long

    ICollectionPlusSettings_HashLookupInitialSize = mlHashLookupInitialSize
    
End Property

Private Property Let ICollectionPlusSettings_HashSize(ByVal RHS As Long)

    If RHS <> mlHashSize Then
        'Do not allow change if there are items currently in collection
        If mlCount > 0 Then ErrRaise EcolUnableChangeHashSize, ERRSOURCE
        If RHS <= 0 Then ErrRaise ERR_INVALID_ARG, ERRSOURCE
        mlHashSize = RHS
        Clear
    End If
    
End Property

Private Property Get ICollectionPlusSettings_HashSize() As Long

    ICollectionPlusSettings_HashSize = mlHashSize
    
End Property

Private Property Let ICollectionPlusSettings_ItemChunkSize(ByVal RHS As Long)

    If RHS <> mlItemChunkSize Then
        If RHS <= 0 Then ErrRaise ERR_INVALID_ARG, ERRSOURCE
        mlItemChunkSize = RHS
    End If
    
End Property

Private Property Get ICollectionPlusSettings_ItemChunkSize() As Long

    ICollectionPlusSettings_ItemChunkSize = mlItemChunkSize
    
End Property

Private Property Let ICollectionPlusSettings_ItemInitialSize(ByVal RHS As Long)

    If RHS <> mlItemInitialSize Then
        If RHS <= 0 Then ErrRaise ERR_INVALID_ARG, ERRSOURCE
        mlItemInitialSize = RHS
        If mlCount = 0 Then Clear
    End If
    
End Property

Private Property Get ICollectionPlusSettings_ItemInitialSize() As Long

    ICollectionPlusSettings_ItemInitialSize = mlItemInitialSize
    
End Property

'****************************************
'Procedure:     CCollectionString.Sort
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          06 July 2000
'****************************************
'Description:   Sort the collection based on set criteria
'
'The sort will be performed in three stages: -
'1.  Store indexes of collection items in array
'2.  Perform the sort operation
'3.  Update the linked list accordingly
'****************************************
'Revisions:
'****************************************
Public Sub Sort(Optional ByVal SortType As CollectionEnumSort = colEnumSortAscendingNoCase)
Attribute Sort.VB_Description = "Sorts the collection"

    Dim alIndexes() As Long
    Dim lCurrIndex As Long
    Dim lPrevIndex As Long
    Dim lLastNext As Long
    Dim i As Long
    
    'Sanity check - no point in sorting < 2 items
    If mlCount < 2 Then Exit Sub
    
    'Determine kind of sort to be made
    Select Case SortType
    Case colEnumSortAscendingCase
        miAscending = 1
        mlCompareType = vbBinaryCompare
    
    Case colEnumSortAscendingNoCase
        miAscending = 1
        mlCompareType = vbTextCompare
    
    Case colEnumSortDescendingCase
        miAscending = -1
        mlCompareType = vbBinaryCompare
    
    Case colEnumSortDescendingNoCase
        miAscending = -1
        mlCompareType = vbTextCompare
    End Select
    
    '1.  Store the indexes
    'ReDim alIndexes(1 To mlCount)
    alIndexes = SafeArrayCreateLongVector(vbLong, 1, mlCount)
    
    lCurrIndex = mlFirstItem
    For i = 1 To mlCount
        alIndexes(i) = lCurrIndex
        lCurrIndex = mauListItems(lCurrIndex).NextItem
    Next i
    
    '2.  Sort the collection
    SortCollectionIndex alIndexes, Me
    
    '3.  Recreate the list based on the sort
    
    'Get the very last 'Next' reference which is used to continue the list
    lLastNext = mauListItems(mlLastItem).NextItem
    For i = 1 To mlCount
        lCurrIndex = alIndexes(i)
        If i = 1 Then
            'This is the new first item
            mauListItems(lCurrIndex).PreviousItem = 0
            mlFirstItem = lCurrIndex
        Else
            mauListItems(lPrevIndex).NextItem = lCurrIndex
            mauListItems(lCurrIndex).PreviousItem = lPrevIndex
        End If
        
        If i = mlCount Then
            'This is the new last item
            mauListItems(lCurrIndex).NextItem = lLastNext
            mlLastItem = lCurrIndex
        End If
        
        lPrevIndex = lCurrIndex
    Next i
    
End Sub


'****************************************
'Procedure:     CCollectionString.EnumDirection
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Returns a value indicating which direction
'the collection will be enumerated
'****************************************
'Revisions:
'****************************************
Public Property Get EnumDirection() As CollectionEnumDirection
Attribute EnumDirection.VB_Description = "Determines direction in which the collection iterates when using For Each"

    EnumDirection = mlEnumDirection

End Property

'****************************************
'Procedure:     CCollectionString.EnumDirection
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Change the enumeration direction
'****************************************
'Revisions:
'****************************************
Public Property Let EnumDirection(ByVal NewEnumDirection As CollectionEnumDirection)

    If NewEnumDirection <> mlEnumDirection Then
        If NewEnumDirection >= colEnumDirForward And NewEnumDirection <= colEnumDirBackward Then
            mlEnumDirection = NewEnumDirection
        Else
            ErrRaise ERR_INVALID_ARG, ERRSOURCE
        End If
    End If

End Property






'****************************************
'Procedure:     CCollectionString.Clone
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          29 June 2000
'****************************************
'Description:   Creates a duplicate of the collection
'****************************************
'Revisions:
'****************************************
Public Function Clone() As CCollectionString
Attribute Clone.VB_Description = "Create a duplicate of the current collection"

    Dim i As Long
    Dim lCurrIndex As Long
    Dim oNewCol As CCollectionString
    
    Set oNewCol = New CCollectionString
    With oNewCol
        .EnumDirection = mlEnumDirection
        .MatchCase = mbMatchCase
    End With
    
    'Ensure storage settings match those in the current object
    GCasts.ICollectionPlusSettings(oNewCol).ChangeSettings mlHashSize, mlHashLookupInitialSize, mlHashLookupChunkSize, mlItemInitialSize, mlItemChunkSize
    
    If mlCount > 0 Then
        lCurrIndex = mlFirstItem
        For i = 1 To mlCount
            With mauListItems(lCurrIndex)
                If KeyExists(lCurrIndex) Then
                    oNewCol.Add Value:=.Value, Key:=.Key, Tag:=.Tag
                Else
                    oNewCol.Add Value:=.Value, Tag:=.Tag
                End If
                
                lCurrIndex = .NextItem
                Debug.Assert lCurrIndex <> 0
            End With
        Next i
    End If
    
    Set Clone = oNewCol
    
End Function

'****************************************
'Procedure:     CCollectionString.ExpandHashLookup
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          27 June 2000
'****************************************
'Description:   Expands the current array of hash lookup items
'****************************************
'Revisions:
'****************************************
Private Sub ExpandHashLookup()

    Dim i As Long
    Dim lNewUpper As Long
        
    'Check number of items
    If mlLastHashIndex = 0 Then
        'Default size
        lNewUpper = mlHashLookupInitialSize
    Else
        lNewUpper = mlLastHashIndex + mlHashLookupChunkSize
    End If
    
    'Update the array
    ReDim Preserve mauHashLookup(1 To lNewUpper)
    
    'Set up list
    For i = mlLastHashIndex + 1 To lNewUpper - 1
        mauHashLookup(i).NextItem = i + 1
    Next i
    
    'Set the next item field of the last index to point at the current free index
    mauHashLookup(lNewUpper).NextItem = mlNextFreeHashItem
    mlNextFreeHashItem = mlLastHashIndex + 1
    mlLastHashIndex = lNewUpper
    
End Sub


'****************************************
'Procedure:     CCollectionString.ExpandList
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          27 June 2000
'****************************************
'Description:   Expand the main list array
'****************************************
'Revisions:
'****************************************
Private Sub ExpandList()

    Dim i As Long
    Dim lNewUpper As Long
        
    'Check number of items
    If mlLastListIndex = 0 Then
        'Default size
        lNewUpper = mlItemInitialSize
    Else
        lNewUpper = mlLastListIndex + mlItemChunkSize
    End If
    
    'Update the array
    ReDim Preserve mauListItems(1 To lNewUpper)
    
    'Set up list
    For i = mlLastListIndex + 1 To lNewUpper
        mauListItems(i).NextItem = i + 1
        mauListItems(i).PreviousItem = -1
    Next i
    
    'Set the last item to point at the new available item
    mauListItems(lNewUpper).NextItem = mlNextFreeListItem
    mlNextFreeListItem = mlLastListIndex + 1
    mlLastListIndex = lNewUpper

End Sub



'****************************************
'Procedure:     CCollectionString.FirstItem
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Internal index to the first item in the collection
'****************************************
'Revisions:
'****************************************
Friend Property Get FirstItem() As Long

    FirstItem = mlFirstItem
    
End Property


'****************************************
'Procedure:     CCollectionString.KeyExists
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Determine if the given item has
'a key associated with it
'****************************************
'Revisions:
'****************************************
Private Function KeyExists(ByVal lItemIndex As Long) As Boolean

    Debug.Assert lItemIndex >= 1 And lItemIndex <= mlLastListIndex

    Dim sKey As String
    
    sKey = mauListItems(lItemIndex).Key
    If LenB(sKey) = 0 Then
        'Possible to have a zero length key - need to check hash table to see if this is so
        KeyExists = Exists(sKey)
    Else
        KeyExists = True
    End If
    
End Function


'****************************************
'Procedure:     CCollectionString.GetItemIndex
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          27 June 2000
'****************************************
'Description:   Attempts to locate referenced item and return its
'index.
'****************************************
'Revisions:
'****************************************
Private Function GetItemIndex(ByRef vIndex As Variant) As Long

    If VarType(vIndex) = vbString Then
        'Index by Key
        Dim sKey As String
        
        If mbMatchCase Then sKey = CStr(vIndex) Else sKey = UCase$(vIndex)
        GetItemIndex = GetStringIndex(sKey)
    Else
        'Assume Index by Ordinal position (needs to be convertable to a Long)
        If Not CanConvertToLong(vIndex) Then ErrRaise ERR_TYPE_MISMATCH, ERRSOURCE
        GetItemIndex = GetOrdinalIndex(vIndex)
    End If
    
End Function



'****************************************
'Procedure:     CCollectionString.GetOrdinalIndex
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          03 July 2000
'****************************************
'Description:   Find an item based on its ordinal number
'****************************************
'Revisions:
'****************************************
Private Function GetOrdinalIndex(ByVal lOrdinal As Long) As Long
    
    Dim i As Long
        
    If lOrdinal < 1 Or lOrdinal > mlCount Then ErrRaise EcolItemNotFound, ERRSOURCE, lOrdinal
    
    'Iterate though to find the desired index
    If lOrdinal > mlCount \ 2 Then
        'Iterate backwards
        GetOrdinalIndex = mlLastItem
        For i = 1 To (mlCount - lOrdinal)
            GetOrdinalIndex = mauListItems(GetOrdinalIndex).PreviousItem
        Next i
    Else
        'Iterate forwards
        GetOrdinalIndex = mlFirstItem
        For i = 1 To lOrdinal - 1
            GetOrdinalIndex = mauListItems(GetOrdinalIndex).NextItem
        Next i
    End If

End Function




'****************************************
'Procedure:     CCollectionString.Add
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          28 June 2000
'****************************************
'Description:   Add an item to the collection
'****************************************
'Revisions:
'****************************************
Public Sub Add(ByRef Value As String, Optional ByRef Key As Variant, Optional ByRef Tag As String, Optional ByRef Before As Variant, Optional ByRef After As Variant)
Attribute Add.VB_Description = "Add an item to the collection"
    
    Dim lAddFlags As ColAddType
    Dim lCompareIndex As Long
    Dim lNewListIndex As Long
    Dim lNewHashIndex As Long
    Dim lIndex As Long
    
    On Error GoTo AddErr
    
    'Check before and after parameters
    If Not IsMissing(Before) Then lAddFlags = lAddFlags Or colAddBefore
    If Not IsMissing(After) Then lAddFlags = lAddFlags Or colAddAfter
    
    'Can't add before and after
    If lAddFlags = colAddIllegal Then ErrRaise ERR_INVALID_ARG, ERRSOURCE
    
    'Create a new hash item - note the procedure will raise an error
    'if the key already exists
    If Not IsMissing(Key) Then lNewHashIndex = CreateNewHashIndex(Key)
        
    'Get index of the next free item
    If mlNextFreeListItem = 0 Then ExpandList
    lNewListIndex = mlNextFreeListItem
    mlNextFreeListItem = mauListItems(mlNextFreeListItem).NextItem
    
    'Link the hash item to the list item and attach the key to the new item
    If lNewHashIndex <> 0 Then
        mauHashLookup(lNewHashIndex).Index = lNewListIndex
        mauListItems(lNewListIndex).Key = Key
    End If
    
    Select Case lAddFlags
    Case colAddDefault          'Insert at end of list
        'Check to see if this is the first item in the list
        'If so, make new item the first item
        If mlCount = 0 Then mlFirstItem = lNewListIndex
        
        'Insert at the end
        If mlLastItem <> 0 Then
            mauListItems(lNewListIndex).PreviousItem = mlLastItem
            mauListItems(mlLastItem).NextItem = lNewListIndex
        End If
        
        'Update last item
        mlLastItem = lNewListIndex
    
    Case colAddBefore           'Before item
        'Get index of Before item
        lCompareIndex = GetItemIndex(Before)
                
        'Store reference to the comparison item's current previous item
        lIndex = mauListItems(lCompareIndex).PreviousItem
        
        'Insert the new item before the comparison item
        With mauListItems(lNewListIndex)
            .NextItem = lCompareIndex
            .PreviousItem = lIndex
        End With
        
        'Update the Before item to point to the new item
        mauListItems(lCompareIndex).PreviousItem = lNewListIndex
        
        'Check to see if Before item was first in list
        If lCompareIndex = mlFirstItem Then
            'Point at the new item instead
            mlFirstItem = lNewListIndex
        Else
            'Update previous item to point at the new item
            mauListItems(lIndex).NextItem = lNewListIndex
        End If
        
    Case colAddAfter            'After item
        'Get index of After item
        lCompareIndex = GetItemIndex(After)
        
        'Store reference to the current item's next item
        lIndex = mauListItems(lCompareIndex).NextItem
        
        'Insert the new item after the comparison item
        With mauListItems(lNewListIndex)
            .NextItem = lIndex
            .PreviousItem = lCompareIndex
        End With
        
        'Update the After item to point to the new item
        mauListItems(lCompareIndex).NextItem = lNewListIndex
        
        'Check to see if After item was last in the list
        If lCompareIndex = mlLastItem Then
            mlLastItem = lNewListIndex
        Else
            'Update next item to point back at this item
            mauListItems(lIndex).PreviousItem = lNewListIndex
        End If
    End Select
    
    'Update first item to point at 0 (indicates start of list)
    mauListItems(mlFirstItem).PreviousItem = 0
    
    'Update last item to point at the next free index
    mauListItems(mlLastItem).NextItem = mlNextFreeListItem
        
    'Update the value of the item
    With mauListItems(lNewListIndex)
        .Value = Value
        .Tag = Tag
    End With
    
    'Increment count
    mlCount = mlCount + 1

Exit Sub

AddErr:
    If lNewListIndex Then RemoveHashItem lNewListIndex
    Err.Raise Err.Number
    
End Sub


'****************************************
'Procedure:     CCollectionString.Clear
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Remove all items from the collection
'****************************************
'Revisions:
'****************************************
Public Sub Clear()
Attribute Clear.VB_Description = "Clears all contents of the collection"
    
    Erase malHashIndex
    Erase mauHashLookup
    Erase mauListItems
    
    'Reset internal pointers
    mlFirstItem = 0
    mlLastItem = 0
    mlLastHashIndex = 0
    mlLastListIndex = 0
    mlCount = 0
    mlNextFreeHashItem = 0
    mlNextFreeListItem = 0
    
    'Create main hash array
    malHashIndex = SafeArrayCreateLongVector(vbLong, 0, mlHashSize + 1)
    
End Sub



'****************************************
'Procedure:     CCollectionString.Count
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Returns number of items in collection
'****************************************
'Revisions:
'****************************************
Public Property Get Count() As Long
Attribute Count.VB_Description = "Number of items in the collection"

    Count = mlCount
    
End Property


'****************************************
'Procedure:     CCollectionString.Exists
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          30 August 2000
'****************************************
'Description:   Detemines if Index exists
'****************************************
'Revisions:
'****************************************
Public Function Exists(ByRef Index As Variant) As Boolean
Attribute Exists.VB_Description = "Determines whether an item is present in the collection"

    Dim bExists As Boolean
    
    bExists = False
    
    If VarType(Index) = vbString Then
        Dim sKey As String
        Dim lHashValue As Long
        Dim lIndex As Long
        
        'Convert key to upper case
        If Not mbMatchCase Then sKey = UCase$(Index) Else sKey = CStr(Index)
        lHashValue = HashString(sKey) Mod mlHashSize
        
        'Attempt to find this index in the hash lookup table
        lIndex = malHashIndex(lHashValue)
        Do Until lIndex = 0
            If mauHashLookup(lIndex).Key = sKey Then
                bExists = True
                Exit Do
            End If
            lIndex = mauHashLookup(lIndex).NextItem
        Loop
    Else
        'Assume numeric index
        If Not CanConvertToLong(Index) Then ErrRaise ERR_TYPE_MISMATCH, ERRSOURCE
        If Index >= 1 And Index <= mlCount Then bExists = True
    End If
    
    Exists = bExists
    
End Function


'****************************************
'Procedure:     CCollectionString.Index
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          28 June 2000
'****************************************
'Description:   Returns the numeric positon of the
'item specified by Key
'****************************************
'Revisions:
'****************************************
Public Property Get Index(ByRef Key As String) As Long
Attribute Index.VB_Description = "Returns the index of an item given its key"
    
    Dim lFindIndex As Long
    Dim lCurrIndex As Long
    Dim lCount As Long
    
    lFindIndex = GetItemIndex(Key)
    
    lCount = 1
    lCurrIndex = mlFirstItem
    Do Until lCurrIndex = lFindIndex
        lCount = lCount + 1
        lCurrIndex = mauListItems(lCurrIndex).NextItem
        
        'Should never reach the end of the available indexes
        Debug.Assert lCurrIndex <> 0
    Loop
    Index = lCount
    
End Property


'****************************************
'Procedure:     CCollectionString.Item
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Returns an item given its index
'****************************************
'Revisions:
'****************************************
Public Property Get Item(ByRef Index As Variant) As String
Attribute Item.VB_Description = "Get/Set an item in the collection"
Attribute Item.VB_UserMemId = 0

    Item = mauListItems(GetItemIndex(Index)).Value
    
End Property


'****************************************
'Procedure:     CCollectionString.Item
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Assign an item
'****************************************
'Revisions:
'****************************************
Public Property Let Item(ByRef Index As Variant, ByRef NewValue As String)
    
    mauListItems(GetItemIndex(Index)).Value = NewValue
    
End Property


'****************************************
'Procedure:     CCollectionString.Items
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Return a string array of all items
'in the collection
'****************************************
'Revisions:
'****************************************
Public Function Items() As String()
Attribute Items.VB_Description = "Returns a string array containing all the items in the collection"

    Dim i As Long
    Dim lCurrIndex As Long
    Dim asItems() As String
    
    If mlCount > 0 Then
        'Create Items array
        ReDim asItems(1 To mlCount)
        
        lCurrIndex = mlFirstItem
        For i = 1 To mlCount
            asItems(i) = mauListItems(lCurrIndex).Value
            lCurrIndex = mauListItems(lCurrIndex).NextItem
        Next i
    End If
    
    Items = asItems
    
End Function

'****************************************
'Procedure:     CCollectionString.Find
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          07 July 2000
'****************************************
'Description:   Returns a CCollectionString containing all
'items that match FindLike.
'****************************************
'Revisions:
'****************************************
Public Function Find(ByRef FindLike As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As CCollectionString
Attribute Find.VB_Description = "Returns a collection of all items matching the FindLike string"

    Dim i As Long
    Dim oFind As CCollectionString
    Dim sCurrent As String
    Dim sFindLike As String
    Dim lCurrIndex As Long
    
    Set oFind = New CCollectionString
    
    If mlCount > 0 Then
        'If we are doing a text comparison, make find string upper case
        If Compare = vbTextCompare Then
            sFindLike = UCase$(FindLike)
        Else
            sFindLike = FindLike
        End If
        
        lCurrIndex = mlFirstItem
        For i = 1 To mlCount
            With mauListItems(lCurrIndex)
                If Compare = vbTextCompare Then
                    sCurrent = UCase$(.Value)
                Else
                    sCurrent = .Value
                End If
                
                If sCurrent Like sFindLike Then
                    If KeyExists(lCurrIndex) Then
                        oFind.Add .Value, .Key, .Tag
                    Else
                        oFind.Add Value:=.Value, Tag:=.Tag
                    End If
                End If
                lCurrIndex = .NextItem
            End With
        Next i
    End If
    
    Set Find = oFind
            
End Function


'****************************************
'Procedure:     CCollectionString.Key
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Assign a key to an item
'****************************************
'Revisions:
'****************************************
Public Property Let Key(ByRef Index As Variant, ByRef NewKey As String)
Attribute Key.VB_Description = "Get/Set an item's key"
        
    Dim lIndex As Long
    
    lIndex = GetItemIndex(Index)
    
    'See if this key exists
    If Exists(NewKey) Then ErrRaise ERR_KEY_EXISTS, ERRSOURCE
    
    'Remove old key
    RemoveHashItem lIndex
    
    'Create a new hash item for the new key
    mauHashLookup(CreateNewHashIndex(NewKey)).Index = lIndex
    
    'Update item key
    mauListItems(lIndex).Key = NewKey
        
End Property


'****************************************
'Procedure:     CCollectionString.Key
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Return an item's key
'****************************************
'Revisions:
'****************************************
Public Property Get Key(ByRef Index As Variant) As String

    Key = mauListItems(GetItemIndex(Index)).Key
    
End Property


'****************************************
'Procedure:     CCollectionString.Keys
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Return an array of strings which
'contains all keys in the collection
'****************************************
'Revisions:
'****************************************
Public Function Keys() As String()
Attribute Keys.VB_Description = "Returns a string array of all the keys for the items in the collection"

    Dim i As Long
    Dim lCurrIndex As Long
    Dim asKeys() As String
    
    If mlCount > 0 Then
        'Create Keys array
        ReDim asKeys(1 To mlCount)
        
        lCurrIndex = mlFirstItem
        For i = 1 To mlCount
            asKeys(i) = mauListItems(lCurrIndex).Key
            lCurrIndex = mauListItems(lCurrIndex).NextItem
        Next i
    End If
    
    Keys = asKeys

End Function


'****************************************
'Procedure:     CCollectionString.LastItem
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Internal index to the last item in
'the collection
'****************************************
'Revisions:
'****************************************
Friend Property Get LastItem() As Long

    LastItem = mlLastItem
    
End Property


'****************************************
'Procedure:     CCollectionString.ListItems
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Returns internal array of ColListItemString
'structures for enumeration purposes
'****************************************
'Revisions:
'****************************************
Friend Function ListItems() As ColListItemString()

    ListItems = mauListItems
    
End Function


'****************************************
'Procedure:     CCollectionString.NewEnum
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Return enumerator object
'****************************************
'Revisions:
'****************************************
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Dim oEnum As CCollectionStringWalker
    
    Set oEnum = New CCollectionStringWalker
    oEnum.Attach Me
    Set NewEnum = oEnum.NewEnum
    
End Function


'****************************************
'Procedure:     CCollectionString.Remove
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          28 June 2000
'****************************************
'Description:   Remove item specified by Index
'****************************************
'Revisions:
'****************************************
Public Sub Remove(ByRef Index As Variant)
Attribute Remove.VB_Description = "Remove an item from the collection"

    Dim lItemIndex As Long
    Dim lPrevIndex As Long
    Dim lNextIndex As Long
    
    lItemIndex = GetItemIndex(Index)
    
    'Remove any associated hash value
    RemoveHashItem lItemIndex
    
    'Detach the item from the list
    With mauListItems(lItemIndex)
        lPrevIndex = .PreviousItem
        lNextIndex = .NextItem
    End With
    
    If lItemIndex = mlFirstItem Then
        'Next item will be the first
        mlFirstItem = lNextIndex
    Else
        'Update previous item to point at the next item
        mauListItems(lPrevIndex).NextItem = lNextIndex
    End If
    
    If lItemIndex = mlLastItem Then
        'Previous item will be the last
        mlLastItem = lPrevIndex
    Else
        'Update next item to point at the previous item
        mauListItems(lNextIndex).PreviousItem = lPrevIndex
    End If
    
    'Update the item itself
    With mauListItems(lItemIndex)
        .Key = vbNullString
        .NextItem = mlNextFreeListItem
        .PreviousItem = -1
        .Value = Empty
    End With
    
    'Reduce the count
    mlCount = mlCount - 1
    
    mlNextFreeListItem = lItemIndex
    
End Sub



'****************************************
'Procedure:     CCollectionString.RemoveHashItem
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          27 June 2000
'****************************************
'Description:   Removes a list item's related hash item
'****************************************
'Revisions:
'****************************************
Private Sub RemoveHashItem(ByVal lItemIndex As Long)

    Dim sKey As String
    Dim lHashIndex As Long
    Dim lPrevIndex As Long
    Dim lNextIndex As Long
    Dim lHashValue As Long
    
    If mbMatchCase Then
        sKey = mauListItems(lItemIndex).Key
    Else
        sKey = UCase$(mauListItems(lItemIndex).Key)
    End If
    
    lHashValue = HashString(sKey) Mod mlHashSize
    lHashIndex = malHashIndex(lHashValue)
    Do While lHashIndex <> 0
        If mauHashLookup(lHashIndex).Key = sKey Then Exit Do
        lPrevIndex = lHashIndex
        lHashIndex = mauHashLookup(lHashIndex).NextItem
    Loop
    
    'If item has key then there should always be an associated hash item
    Debug.Assert lHashIndex <> 0 Or LenB(sKey) = 0
    If lHashIndex = 0 Then Exit Sub
    
    'Get index to next item in the list
    lNextIndex = mauHashLookup(lHashIndex).NextItem
    
    'Remove the item from the linked list
    If lPrevIndex = 0 And lNextIndex = 0 Then
        'No more hash items with this hash value - clear
        malHashIndex(lHashValue) = 0
    Else
        If lPrevIndex <> 0 Then
            mauHashLookup(lPrevIndex).NextItem = lNextIndex
        Else
            'Update hash table to point at the new head of the list
            malHashIndex(lHashValue) = lNextIndex
        End If
    End If
    
    'Clear out item
    With mauHashLookup(lHashIndex)
        .NextItem = mlNextFreeHashItem
        .Index = 0
        .Key = vbNullString
    End With
    
    mlNextFreeHashItem = lHashIndex
            
End Sub

Private Sub CCollectionPlus_Add(Value As Variant, Optional Key As Variant, Optional Tag As String, Optional Before As Variant, Optional After As Variant)

    Me.Add Value:=CStr(Value), Key:=Key, Tag:=Tag, Before:=Before, After:=After
    
End Sub

Private Sub CCollectionPlus_Clear()

    Me.Clear
    
End Sub


Private Function CCollectionPlus_Clone() As CCollectionPlus

    Set CCollectionPlus_Clone = Me.Clone
    
End Function


Private Property Get CCollectionPlus_Count() As Long

    CCollectionPlus_Count = mlCount
    
End Property


Private Property Let CCollectionPlus_EnumDirection(ByVal RHS As CollectionEnumDirection)

    Me.EnumDirection = RHS
    
End Property

Private Property Get CCollectionPlus_EnumDirection() As CollectionEnumDirection

    CCollectionPlus_EnumDirection = mlEnumDirection
    
End Property

Private Function CCollectionPlus_Exists(Key As Variant) As Boolean

    CCollectionPlus_Exists = Me.Exists(Key)
    
End Function


Private Function CCollectionPlus_Find(FindLike As String, Optional ByVal Compare As VbCompareMethod = 0&) As CCollectionPlus

    Set CCollectionPlus_Find = Me.Find(FindLike, Compare)
    
End Function

Private Property Get CCollectionPlus_Index(Key As String) As Long

    CCollectionPlus_Index = Me.Index(Key)
    
End Property

Private Property Set CCollectionPlus_Item(Index As Variant, RHS As Variant)

    ErrRaise ERR_INVALID_PROPERTY, ERRSOURCE
    
End Property

Private Property Let CCollectionPlus_Item(Index As Variant, RHS As Variant)

    Me.Item(Index) = CStr(RHS)
    
End Property

Private Property Get CCollectionPlus_Item(Index As Variant) As Variant

    CCollectionPlus_Item = Me.Item(Index)
    
End Property


'****************************************
'Procedure:     CCollectionString.CCollectionPlus_Items
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          10 July 2000
'****************************************
'Description:   Convert String array to Variant array
'****************************************
'Revisions:
'****************************************
Private Function CCollectionPlus_Items() As Variant()

    Dim i As Long
    Dim avItems() As Variant
    Dim asItems() As String
    
    If mlCount > 0 Then
        asItems = Me.Items
        ReDim avItems(1 To mlCount)
        
        For i = 1 To mlCount
            avItems(i) = asItems(i)
        Next i
    End If
        
    CCollectionPlus_Items = avItems
    
End Function

Private Property Let CCollectionPlus_Key(Index As Variant, RHS As String)

    Me.Key(Index) = RHS
    
End Property

Private Property Get CCollectionPlus_Key(Index As Variant) As String

    CCollectionPlus_Key = Me.Key(Index)
    
End Property


Private Function CCollectionPlus_Keys() As String()

    CCollectionPlus_Keys = Me.Keys
    
End Function

Private Function CCollectionPlus_NewEnum() As IEnumVARIANT

    Set CCollectionPlus_NewEnum = Me.NewEnum
    
End Function


Private Sub CCollectionPlus_Remove(Index As Variant)

    Me.Remove Index
    
End Sub


Private Sub CCollectionPlus_Sort(Optional ByVal SortType As CollectionEnumSort = 1&)

    Me.Sort SortType
    
End Sub


'****************************************
'Procedure:     CCollectionString.Class_ReadProperties
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          04 July 2000
'****************************************
'Description:
'****************************************
'Revisions:
'****************************************
Private Sub Class_ReadProperties(PropBag As PropertyBag)

    On Error Resume Next

    Dim i As Long
    Dim lCount As Long
    Dim sKey As String
    Dim sItem As String
    Dim sTag As String
    Dim bAddKey As Boolean
    Dim lHashSize As Long
    Dim lHashLookupChunkSize As Long
    Dim lHashLookupInitialSize As Long
    Dim lItemChunkSize As Long
    Dim lItemInitialSize As Long
    
    With PropBag
        'Behaviour settings
        mlEnumDirection = .ReadProperty(PROP_ENUMDIR, colEnumDirForward)
        mbMatchCase = .ReadProperty(PROP_MATCHCASE, False)
        
        'Get storage settings
        lHashSize = .ReadProperty(PROP_HASH_SIZE, DFLT_HASH_SIZE)
        lHashLookupChunkSize = .ReadProperty(PROP_HASH_LOOKUP_CHUNK_SIZE, DFLT_HASH_LOOKUP_CHUNK_SIZE)
        lHashLookupInitialSize = .ReadProperty(PROP_HASH_LOOKUP_INIT_SIZE, DFLT_HASH_LOOKUP_INITIAL_SIZE)
        lItemChunkSize = .ReadProperty(PROP_ITEM_CHUNK_SIZE, DFLT_ITEM_CHUNK_SIZE)
        lItemInitialSize = .ReadProperty(PROP_ITEM_INIT_SIZE, DFLT_ITEM_INITIAL_SIZE)
        
        'Call to ICollectionPlusSettings interface - note this will only make changes if
        'settings were changed from the default in the persisted object.  It would not
        'be necessary to do it this way except that the InitProperties event does not seem to
        'reliably fire in certain clients.
        GCasts.ICollectionPlusSettings(Me).ChangeSettings lHashSize, lHashLookupInitialSize, lHashLookupChunkSize, lItemInitialSize, lItemChunkSize
        
        'Get number of items
        lCount = .ReadProperty(PROP_COUNT, 0)
        If lCount > 0 Then
            For i = 1 To lCount
                bAddKey = True
                Err.Clear
                sKey = .ReadProperty(PROP_KEY & CStr(i))
                If Err Then
                    'Item does not have an associated key
                    Err.Clear
                    bAddKey = False
                End If
                
                sItem = .ReadProperty(PROP_ITEM & CStr(i), vbNullString)
                
                If Err = 0 Then
                    sTag = .ReadProperty(PROP_TAG & CStr(i), vbNullString)
                    
                    'Add to the collection
                    If bAddKey Then
                        Me.Add Value:=sItem, Key:=sKey, Tag:=sTag
                    Else
                        Me.Add Value:=sItem, Tag:=sTag
                    End If
                End If
            Next i
        End If
    End With
    
End Sub

Private Sub Class_Terminate()

    Erase malHashIndex
    Erase mauHashLookup
    Erase mauListItems
    
End Sub



'****************************************
'Procedure:     CCollectionString.Class_WriteProperties
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          04 July 2000
'****************************************
'Description:
'****************************************
'Revisions:
'****************************************
Private Sub Class_WriteProperties(PropBag As PropertyBag)

    On Error Resume Next

    Dim i As Long
    Dim lCurrIndex As Long
    
    With PropBag
        'Enumeration direction
        .WriteProperty PROP_ENUMDIR, mlEnumDirection, colEnumDirForward
    
        'Collection settings
        .WriteProperty PROP_HASH_SIZE, mlHashSize, DFLT_HASH_SIZE
        .WriteProperty PROP_HASH_LOOKUP_CHUNK_SIZE, mlHashLookupChunkSize, DFLT_HASH_LOOKUP_CHUNK_SIZE
        .WriteProperty PROP_HASH_LOOKUP_INIT_SIZE, mlHashLookupInitialSize, DFLT_HASH_LOOKUP_INITIAL_SIZE
        .WriteProperty PROP_ITEM_CHUNK_SIZE, mlItemChunkSize, DFLT_ITEM_CHUNK_SIZE
        .WriteProperty PROP_ITEM_INIT_SIZE, mlItemInitialSize, DFLT_ITEM_INITIAL_SIZE
        .WriteProperty PROP_MATCHCASE, mbMatchCase, False
        
        .WriteProperty PROP_COUNT, mlCount, 0
        If mlCount > 0 Then
            lCurrIndex = mlFirstItem
            For i = 1 To mlCount
                .WriteProperty PROP_ITEM & CStr(i), mauListItems(lCurrIndex).Value, vbNullString
                .WriteProperty PROP_TAG & CStr(i), mauListItems(lCurrIndex).Tag, vbNullString
                If KeyExists(lCurrIndex) Then
                    .WriteProperty PROP_KEY & CStr(i), mauListItems(lCurrIndex).Key
                End If
                
                lCurrIndex = mauListItems(lCurrIndex).NextItem
                If lCurrIndex = 0 Then Exit For
            Next i
        End If
    End With

End Sub

Private Sub IVBCollection_Add(Item As Variant, Optional Key As Variant, Optional Before As Variant, Optional After As Variant)

    Me.Add Value:=CStr(Item), Key:=Key, Before:=Before, After:=After
    
End Sub

Private Function IVBCollection_Count() As Long

    IVBCollection_Count = mlCount
    
End Function


Private Function IVBCollection_Item(Index As Variant) As Variant

    IVBCollection_Item = Me.Item(Index)

End Function


Private Function IVBCollection_NewEnum() As stdole.IUnknown

    Set IVBCollection_NewEnum = Me.NewEnum

End Function


Private Sub IVBCollection_Remove(Index As Variant)

    Me.Remove Index
    
End Sub

Private Function ICollectionPlus_VB5_Items() As Variant

    ICollectionPlus_VB5_Items = Me.Items
    
End Function


Private Function ICollectionPlus_VB5_Keys() As Variant

    ICollectionPlus_VB5_Keys = Me.Keys
    
End Function




Private Function ICollectionPlusSortHelper_Compare(ByVal Index1 As Long, ByVal Index2 As Long) As Integer

    ICollectionPlusSortHelper_Compare = StrComp(mauListItems(Index1).Value, mauListItems(Index2).Value, mlCompareType) * miAscending
    
End Function


Private Sub ICollectionPlusSortHelper_Swap(Index1 As Long, Index2 As Long)

    Dim lTemp As Long
    
    'Swap indexes rather than items
    lTemp = Index1
    Index1 = Index2
    Index2 = lTemp
    
End Sub


