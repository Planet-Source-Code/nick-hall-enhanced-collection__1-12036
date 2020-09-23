VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCollectionPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Enhanced version of the Collection class"
'Replacement collection class, featuring dictionary-like methods
'as well as the ability to sort, persist and find items that match
'certain criteria

Option Explicit

'Implements a VB-compatible version of the standard collection interface
Implements IVBCollection

'Interface for performing sorts on collection contents
Implements ICollectionPlusSortHelper

'Interface for methods that are not compatible with VB5
Implements ICollectionPlus_VB5

'Interface to enable users to alter internal settings
Implements ICollectionPlusSettings

Public Enum CollectionEnumDirection
    colEnumDirForward
    colEnumDirBackward
End Enum

Public Enum CollectionEnumSort
    colEnumSortAscendingNoCase = 1
    colEnumSortAscendingCase
    colEnumSortDescendingNoCase
    colEnumSortDescendingCase
End Enum
    
'***Constants

Private Const DFLT_HASH_SIZE As Long = 2048                 'Size of hash lookup array
Private Const DFLT_HASH_LOOKUP_INITIAL_SIZE As Long = 4096  'Initial size of hash item list
Private Const DFLT_HASH_LOOKUP_CHUNK_SIZE As Long = 1024    'Amount by which the hash array is increased
Private Const DFLT_ITEM_INITIAL_SIZE As Long = 1024         'Initial size of ColListItem array
Private Const DFLT_ITEM_CHUNK_SIZE As Long = 512            'Amount by which the ColListItem array is increased
Private Const ERRSOURCE = "CollectionPlus.CCollectionPlus"  'Constant used for Error source argument

'*** Private members
Private malHashIndex() As Long                              'Hash value indexes into the hash lookup array
Private mauHashLookup() As HashItem                         'Linked list containing index of ColListItem in main array
Private mauListItems() As ColListItem                       'Main linked list containing actual items that are stored

Private mlFirstItem As Long                                 'Index to the first item in the list
Private mlLastItem As Long                                  'Index to last item in the list
Private mlLastHashIndex As Long                             'Upper bound of Hash lookup array
Private mlLastListIndex As Long                             'Upper bound of List Item array
Private mlNextFreeHashItem As Long                          'Index to the next unallocated item in the hash array
Private mlNextFreeListItem As Long                          'Index to the next unallocated item in the list array

'Settings for hash size/list size
Private mlHashSize As Long
Private mlHashLookupInitialSize As Long
Private mlHashLookupChunkSize As Long
Private mlItemInitialSize As Long
Private mlItemChunkSize As Long

' Number of items in collection
Private mlCount As Long ' Added by: Nick Hall on 29/10/1999 9:33:33 am

' Determine whether the collection iterates forwards or backwards
Private mlEnumDirection As CollectionEnumDirection ' Added by: Nick Hall on 05/07/2000 9:43:47 am

' If true, key hashing is case-sensitive.  Default is false
Private mbMatchCase As Boolean ' Added by: Nick Hall on 01/09/2000 2:33:29 pm

'Determines what type of sort
Private miAscending As Integer                              'Determines whether items are sorted ascending or descending
Private mlCompareType As VbCompareMethod                    'Determines how items are compared

'****************************************
'Procedure:     CCollectionPlus.CreateNewHashIndex
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
'Procedure:     CCollectionPlus.MatchCase
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Determines whether keys are compared
'case-insensitive (default,false) or case-sensitive (true)
'****************************************
'Revisions:
'****************************************
Public Property Get MatchCase() As Boolean
Attribute MatchCase.VB_Description = "Determines whether keys are compared case-insensitive (false, default value) or case-sensitive (true)"

    MatchCase = mbMatchCase

End Property

'****************************************
'Procedure:     CCollectionPlus.MatchCase
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Change the way keys are compared
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
'Procedure:     CCollectionPlus.DoFindItems
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          24 August 2000
'****************************************
'Description:   Returns a CCollectionPlus object containing
'all objects in the collection that a) support the ICollectionPlusItem
'interface and b) Returns true from the object's implementation of the
'IsLike method
'****************************************
'Revisions:
'****************************************
Private Function DoFindItems(ByRef FindParams() As Variant) As CCollectionPlus

    Dim i As Long
    Dim oItem As ICollectionPlusItem
    Dim oFindItems As CCollectionPlus
    Dim lCurrIndex As Long

    Set oFindItems = New CCollectionPlus
    
    If mlCount > 0 Then
        lCurrIndex = mlFirstItem
        
        On Error Resume Next
        For i = 1 To mlCount
            With mauListItems(lCurrIndex)
                Set oItem = .Value
                If Err Then
                    Err.Clear
                Else
                    If oItem.IsLike(FindParams) Then
                        If KeyExists(lCurrIndex) Then
                            oFindItems.Add .Value, .Key, .Tag
                        Else
                            oFindItems.Add Value:=.Value, Tag:=.Tag
                        End If
                    End If
                End If
                
                lCurrIndex = .NextItem
            End With
        Next i
    End If
    
    Set DoFindItems = oFindItems
    
End Function


'****************************************
'Procedure:     CCollectionPlus.EnumDirection
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
Attribute EnumDirection.VB_Description = "Determine whether the collection iterates forwards or backwards when used with For Each"

    EnumDirection = mlEnumDirection

End Property

'****************************************
'Procedure:     CCollectionPlus.EnumDirection
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
'Procedure:     CCollectionPlus.Clone
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          29 June 2000
'****************************************
'Description:   Creates a duplicate of the collection
'****************************************
'Revisions:
'****************************************
Public Function Clone() As CCollectionPlus
Attribute Clone.VB_Description = "Create a duplicate of the current object"

    Dim i As Long
    Dim lCurrIndex As Long
    Dim oNewCol As CCollectionPlus
    
    Set oNewCol = New CCollectionPlus
    With oNewCol
        .EnumDirection = mlEnumDirection
        .MatchCase = mbMatchCase
    End With
    
    'Update settings to match current collection
    GCasts.ICollectionPlusSettings(oNewCol).ChangeSettings mlHashSize, mlHashLookupInitialSize, mlHashLookupChunkSize, mlItemInitialSize, mlItemChunkSize
    
    If mlCount > 0 Then
        lCurrIndex = mlFirstItem
        For i = 1 To mlCount
            With mauListItems(lCurrIndex)
                'Check to see if item has key
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
'Procedure:     CCollectionPlus.ExpandHashLookup
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
'Procedure:     CCollectionPlus.ExpandList
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
'Procedure:     CCollectionPlus.Find
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          07 July 2000
'****************************************
'Description:   Returns a collection containing all
'items that match FindLike.  Note: the stored variant must
'be either a string, or convertable into a string, or an
'object that supports a string default property in order
'for this to work
'****************************************
'Revisions:
'****************************************
Public Function Find(ByRef FindLike As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As CCollectionPlus
Attribute Find.VB_Description = "Returns a collection of all items matching the FindLike string"

    Dim i As Long
    Dim sCurrent As String
    Dim sFindLike As String
    Dim lCurrIndex As Long
    Dim oFind As CCollectionPlus
    
    On Error Resume Next
    
    'Create a new collection to hold the returned items
    Set oFind = New CCollectionPlus
    
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
                
                If Err Then
                    'Ignore error and carry on with next variant
                    Err.Clear
                Else
                    If sCurrent Like sFindLike Then
                        If KeyExists(lCurrIndex) Then
                            oFind.Add .Value, .Key, .Tag
                        Else
                            oFind.Add Value:=.Value, Tag:=.Tag
                        End If
                    End If
                    lCurrIndex = .NextItem
                End If
            End With
        Next i
    End If
    
    Set Find = oFind
    
End Function

'****************************************
'Procedure:     CCollectionPlus.FindItems
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          24 August 2000
'****************************************
'Description:   Use callback interface implemented in collection
'objects to create a user-defined search of the collection.
'
'Any variants in the collection that are not objects or do not implement
'the ICollectionPlusItem interface will be ignored
'****************************************
'Revisions:
'****************************************
Public Function FindItems(ParamArray FindParams() As Variant) As CCollectionPlus
Attribute FindItems.VB_Description = "Find all items in collection that implement the ICollectionPlusItem callback interface and match the parameters provided to FindParams"

    Dim i As Long
    Dim vLocalParams() As Variant
    Dim bMissing As Boolean
    
    bMissing = IsMissing(FindParams)
    
    'Copy parameter array to local array (can't pass ParamArrays directly to other functions)
    If Not bMissing Then vLocalParams = FindParams
    Set FindItems = DoFindItems(vLocalParams)
    
    'Copy back parameters to parameter array (may have changed)
    If Not bMissing Then
        For i = LBound(vLocalParams) To UBound(vLocalParams)
            VariantMove FindParams(i), vLocalParams(i)
        Next i
    End If
    
End Function



'****************************************
'Procedure:     CCollectionPlus.FirstItem
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Return internal index to the first item
'in the collection
'****************************************
'Revisions:
'****************************************
Friend Property Get FirstItem() As Long

    FirstItem = mlFirstItem
    
End Property

'****************************************
'Procedure:     CCollectionPlus.GetStringIndex
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
'Procedure:     CCollectionPlus.GetItemIndex
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
        Dim sKey As String
        
        If mbMatchCase Then sKey = vIndex Else sKey = UCase$(vIndex)
        GetItemIndex = GetStringIndex(sKey)
    Else
        If Not CanConvertToLong(vIndex) Then ErrRaise ERR_TYPE_MISMATCH, ERRSOURCE
        GetItemIndex = GetOrdinalIndex(vIndex)
    End If
    
End Function



'****************************************
'Procedure:     CCollectionPlus.GetOrdinalIndex
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
'Procedure:     CCollectionPlus.Add
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          28 June 2000
'****************************************
'Description:   Add an item to the collection
'****************************************
'Revisions:
'****************************************
Public Sub Add(ByRef Value As Variant, Optional ByRef Key As Variant, Optional ByRef Tag As String, Optional ByRef Before As Variant, Optional ByRef After As Variant)
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
    Case colAddDefault  'Insert at end of list
        'Check to see if this is the first item in the list
        If mlCount = 0 Then mlFirstItem = lNewListIndex
        
        'Insert at the end
        If mlLastItem <> 0 Then
            mauListItems(lNewListIndex).PreviousItem = mlLastItem
            mauListItems(mlLastItem).NextItem = lNewListIndex
        End If
        
        'Update last item
        mlLastItem = lNewListIndex
    
    Case colAddBefore   'Before item
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
        
    Case colAddAfter    'After item
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
        VariantCopyInd .Value, Value
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
'Procedure:     CCollectionPlus.Clear
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Clears out all items in the collection
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
'Procedure:     CCollectionPlus.Count
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Returns the number of items in the collection
'****************************************
'Revisions:
'****************************************
Public Property Get Count() As Long
Attribute Count.VB_Description = "Number of items in the collection"

    Count = mlCount
    
End Property



'****************************************
'Procedure:     CCollectionPlus.Exists
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          30 August 2000
'****************************************
'Description:   Returns true if an item with the specified Index can be found;
'false otherwise
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
'Procedure:     CCollectionPlus.Index
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
'Procedure:     CCollectionPlus.Item
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Returns an item given it skey, or its ordinal
'****************************************
'Revisions:
'****************************************
Public Property Get Item(ByRef Index As Variant) As Variant
Attribute Item.VB_Description = "Get/Set an item in the collection"
Attribute Item.VB_UserMemId = 0

    VariantCopy Item, mauListItems(GetItemIndex(Index)).Value
        
End Property


'****************************************
'Procedure:     CCollectionPlus.Item
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Replace an object
'****************************************
'Revisions:
'****************************************
Public Property Set Item(ByRef Index As Variant, ByRef NewValue As Variant)

    Set mauListItems(GetItemIndex(Index)).Value = NewValue
    
End Property


'****************************************
'Procedure:     CCollectionPlus.Item
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Replace an item
'****************************************
'Revisions:
'****************************************
Public Property Let Item(ByRef Index As Variant, ByRef NewValue As Variant)
    
    mauListItems(GetItemIndex(Index)).Value = NewValue
    
End Property


'****************************************
'Procedure:     CCollectionPlus.Items
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Returns an array of variants containing all the
'items in the collection
'****************************************
'Revisions:
'****************************************
Public Function Items() As Variant()
Attribute Items.VB_Description = "Returns a variant array containing all the items in the collection"

    Dim i As Long
    Dim lCurrIndex As Long
    Dim avItems() As Variant
    
    If mlCount > 0 Then
        'Create Items array
        ReDim avItems(1 To mlCount)
        
        lCurrIndex = mlFirstItem
        For i = 1 To mlCount
            VariantCopy avItems(i), mauListItems(lCurrIndex).Value
            lCurrIndex = mauListItems(lCurrIndex).NextItem
        Next i
    End If
    
    Items = avItems
    
End Function


'****************************************
'Procedure:     CCollectionPlus.Key
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Replace an item's key
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
'Procedure:     CCollectionPlus.Key
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
'Procedure:     CCollectionPlus.Keys
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Returns a string array of all
'keys in the collection
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
'Procedure:     CCollectionPlus.LastItem
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Returns an internal index to the
'last item in the collection
'****************************************
'Revisions:
'****************************************
Friend Property Get LastItem() As Long

    LastItem = mlLastItem
    
End Property


'****************************************
'Procedure:     CCollectionPlus.ListItems
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Returns internal array of ColListItem
'structures for enumeration
'****************************************
'Revisions:
'****************************************
Friend Function ListItems() As ColListItem()

    ListItems = mauListItems
    
End Function


'****************************************
'Procedure:     CCollectionPlus.Load
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          01 September 2000
'****************************************
'Description:   Load a previously saved collection into the exiting one.
'If there are no items currently in the collection, also adopts saved
'settings.  If IgnoreDuplicates is set, any items using keys already in use
'by items in the current collection will be ignored; if it is not set then the
'existing values will be overwritten.
'****************************************
'Revisions:
'****************************************
Public Sub Load(ByRef FileName As String, Optional ByVal IgnoreDuplicates As Boolean = True, Optional ByVal ColSite As ICollectionPlusSite)
Attribute Load.VB_Description = "Merge the previously saved collecton into the current collection"

    On Error GoTo LoadErr

    Dim bAdd As Boolean
    Dim bFileOpen As Boolean
    Dim i As Long
    Dim lFileNum As Long
    Dim lCount As Long
    Dim udtHeader As ColFileHeader
    Dim udtItem As ColFileItem
    Dim pb As New PropertyBag
    
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
            Select Case .ItemType
            Case colTypeColItem
                If Not ColSite Is Nothing Then
                    'Raise an event to the owner of the collection to enable them
                    'to recreate the object based on saved contents.  Note the definition
                    'of Item is Object not ICollectionPlusItem.  This prevents a subtle bug
                    'whereby the IDispatch reference held in the collection points to the
                    'ICollectionPlusItem interface instead of the default interface which
                    'can cause problems
                    On Error Resume Next
                    Set .Item = ColSite.RestoreItem(.Item)
                    If Err Then
                        bAdd = False
                    Else
                        bAdd = Not (.Item Is Nothing)
                    End If
                    On Error GoTo LoadErr
                Else
                    bAdd = False
                End If
                
            Case colTypePersistObject
                pb.Contents = .Item
                Set .Item = pb.ReadProperty(PROP_ITEM)
                bAdd = True
                
            Case colTypeOther
                bAdd = True
            End Select
                
            If bAdd Then
                If IsNull(.Key) Then
                    Me.Add Value:=.Item, Tag:=.Tag
                Else
                    Me.Add Value:=.Item, Key:=.Key, Tag:=.Tag
                End If
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
                If IsObject(.Item) Then Set Me.Item(.Key) = .Item Else Me.Item(.Key) = .Item
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
'Procedure:     CCollectionPlus.KeyExists
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Determine whether a given item in the
'collection has an associated key value
'****************************************
'Revisions:
'****************************************
Private Function KeyExists(ByVal lItemIndex As Long) As Boolean

    Dim sKey As String
    
    Debug.Assert lItemIndex >= 1 And lItemIndex <= mlLastListIndex

    sKey = mauListItems(lItemIndex).Key
    If LenB(sKey) = 0 Then
        KeyExists = Exists(sKey)
    Else
        KeyExists = True
    End If
    
End Function



'****************************************
'Procedure:     CCollectionPlus.Move
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
    
    With mauListItems(lCurrIndex)
        'Dettach item from its existing position
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
    
    With mauListItems(lCompareIndex)
        'Reattach in new position
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
'Procedure:     CCollectionPlus.NewEnum
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Create an enumerator
'****************************************
'Revisions:
'****************************************
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Dim oEnum As CCollectionWalker
    
    Set oEnum = New CCollectionWalker
    oEnum.Attach Me
    Set NewEnum = oEnum.NewEnum
    
End Function


'****************************************
'Procedure:     CCollectionPlus.Remove
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
        .Tag = vbNullString
        .Value = Empty
    End With
    
    'Reduce the count
    mlCount = mlCount - 1
    
    mlNextFreeListItem = lItemIndex
    
End Sub



'****************************************
'Procedure:     CCollectionPlus.RemoveHashItem
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
    Debug.Assert lHashIndex > 0 Or LenB(sKey) = 0
    If lHashIndex = 0 Then Exit Sub
    
    'Get indexes to next and previous items in the list
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



'****************************************
'Procedure:     CCollectionPlus.Save
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
'Procedure:     CCollectionPlus.Sort
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
Attribute Sort.VB_Description = "Sort the collection"

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
'Procedure:     CCollectionPlus.Tag
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Assign an item's Tag
'****************************************
'Revisions:
'****************************************
Public Property Let Tag(ByRef Index As Variant, ByRef NewTag As String)
Attribute Tag.VB_Description = "User defined data associated with item"
    
    mauListItems(GetItemIndex(Index)).Tag = NewTag
    
End Property


'****************************************
'Procedure:     CCollectionPlus.Tag
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          09 October 2000
'****************************************
'Description:   Return an item's Tag
'****************************************
'Revisions:
'****************************************
Public Property Get Tag(ByRef Index As Variant) As String

    Tag = mauListItems(GetItemIndex(Index)).Tag
    
End Property

'****************************************
'Procedure:     CCollectionPlus.Class_Initialize
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
'Procedure:     CCollectionPlus.Class_ReadProperties
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
    Dim sTag As String
    Dim vItem As Variant
    Dim bAddKey As Boolean
    Dim lHashSize As Long
    Dim lHashLookupChunkSize As Long
    Dim lHashLookupInitialSize As Long
    Dim lItemChunkSize As Long
    Dim lItemInitialSize As Long
    
    'Behaviour settings
    With PropBag
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
                    'No key to add
                    Err.Clear
                    bAddKey = False
                End If
                
                'Move variant from property bag into local variable
                VariantMove vItem, .ReadProperty(PROP_ITEM & CStr(i))
                
                If Err = 0 Then
                    sTag = .ReadProperty(PROP_TAG & CStr(i), vbNullString)
                    
                    'Add to the collection
                    If bAddKey Then
                        Me.Add Value:=vItem, Key:=sKey, Tag:=sTag
                    Else
                        Me.Add Value:=vItem, Tag:=sTag
                    End If
                End If
            Next i
        End If
    End With
    
End Sub

Private Sub Class_Terminate()

    'Erase dynamic arrays
    Erase malHashIndex
    Erase mauHashLookup
    Erase mauListItems
    
End Sub



'****************************************
'Procedure:     CCollectionPlus.Class_WriteProperties
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
        
        'Begin writing of collection contents
        .WriteProperty PROP_COUNT, mlCount, 0
        If mlCount > 0 Then
            lCurrIndex = mlFirstItem
            For i = 1 To mlCount
                'Writing the item will fail if variant contains an object that is not
                'persistable.  Unfortunately, any error raised here will result in a
                'fatal termination of the client application.  Therefore, simply ignore
                'any items that cause errors
                Err.Clear
                
                .WriteProperty PROP_ITEM & CStr(i), mauListItems(lCurrIndex).Value
                
                If Err = 0 Then
                    If KeyExists(lCurrIndex) Then
                        .WriteProperty PROP_KEY & CStr(i), mauListItems(lCurrIndex).Key
                    End If
                    .WriteProperty PROP_TAG & CStr(i), mauListItems(lCurrIndex).Tag, vbNullString
                End If
                
                lCurrIndex = mauListItems(lCurrIndex).NextItem
                If lCurrIndex = 0 Then Exit For
            Next i
        End If
    End With

End Sub


Private Sub IVBCollection_Add(Item As Variant, Optional Key As Variant, Optional Before As Variant, Optional After As Variant)

    Me.Add Value:=Item, Key:=Key, Before:=Before, After:=After
    
End Sub

Private Function IVBCollection_Count() As Long

    IVBCollection_Count = mlCount
    
End Function


Private Function IVBCollection_Item(Index As Variant) As Variant

    VariantMove IVBCollection_Item, Me.Item(Index)

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



'****************************************
'Procedure:     CCollectionPlus.ICollectionPlusSettings_ChangeSettings
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

Private Function ICollectionPlusSortHelper_Compare(ByVal Index1 As Long, ByVal Index2 As Long) As Integer

    Dim v1 As Variant
    Dim v2 As Variant
    Dim vEmpty As Variant
    
    'Copy variants into local variables
    CopyMemory v1, mauListItems(Index1).Value, 16
    CopyMemory v2, mauListItems(Index2).Value, 16
    
    On Error Resume Next
    If IsObject(v1) And IsObject(v2) Then
        'Attempt to use ICollectionPlusItem interface to compare the two items
        ICollectionPlusSortHelper_Compare = GCasts.ICollectionPlusItem(v1).Compare(v2, mlCompareType) * miAscending
        
        If Err Then
            'Compare as strings
            Err.Clear
            ICollectionPlusSortHelper_Compare = StrComp(v1, v2, mlCompareType) * miAscending
        End If
    Else
        'This line could potentially cause an error if the variants involved do not
        'contain strings or values that can be converted to strings or objects
        'that have default properties of type string
        ICollectionPlusSortHelper_Compare = StrComp(v1, v2, mlCompareType) * miAscending
    End If
    
    'Default to 0 if error (no items will be swapped)
    If Err Then ICollectionPlusSortHelper_Compare = 0
    
    'Release local variants
    CopyMemory v1, vEmpty, 16
    CopyMemory v2, vEmpty, 16
    
End Function


Private Sub ICollectionPlusSortHelper_Swap(Index1 As Long, Index2 As Long)

    Dim lTemp As Long
    
    'Note: At this point we are swapping INDEXES not VALUES
    lTemp = Index1
    Index1 = Index2
    Index2 = lTemp
    
End Sub


