VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Collection Plus Demo"
   ClientHeight    =   6030
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   120
      Top             =   5460
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "col"
      Filter          =   "Collection + File (*.col)|*.col|All Files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6180
      TabIndex        =   1
      Top             =   5460
      Width           =   1200
   End
   Begin VB.ListBox lstItems 
      Height          =   5325
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   60
      Width           =   7275
   End
   Begin VB.Label lblDrag 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOptions 
         Caption         =   "&Load Collection"
         Index           =   1
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "&Save Collection"
         Index           =   2
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "E&xit"
         Index           =   4
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Auto Refresh"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Edit Item"
         Index           =   3
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Refresh List"
         Index           =   4
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuCollection 
      Caption         =   "&Collection"
      Begin VB.Menu mnuCollectionOptions 
         Caption         =   "&Type"
         Index           =   1
         Begin VB.Menu mnuCollectionType 
            Caption         =   "&Variant"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuCollectionType 
            Caption         =   "&String"
            Index           =   2
         End
      End
      Begin VB.Menu mnuCollectionOptions 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCollectionOptions 
         Caption         =   "Add &String..."
         Index           =   3
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuCollectionOptions 
         Caption         =   "Add &Object..."
         Index           =   4
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuCollectionOptions 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuCollectionOptions 
         Caption         =   "&Find..."
         Index           =   6
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuCollectionOptions 
         Caption         =   "Find &Items..."
         Index           =   7
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuCollectionOptions 
         Caption         =   "&Sort"
         Index           =   8
         Begin VB.Menu mnuCollectionSort 
            Caption         =   "Ascending, Case Insensitive"
            Index           =   1
         End
         Begin VB.Menu mnuCollectionSort 
            Caption         =   "Ascending, Case Sensitive"
            Index           =   2
         End
         Begin VB.Menu mnuCollectionSort 
            Caption         =   "Descending, Case Insensitive"
            Index           =   3
         End
         Begin VB.Menu mnuCollectionSort 
            Caption         =   "Descending, Case Sensitive"
            Index           =   4
         End
      End
      Begin VB.Menu mnuCollectionOptions 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuCollectionOptions 
         Caption         =   "Remove Item"
         Index           =   10
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuCollectionOptions 
         Caption         =   "Remove &All Items"
         Index           =   11
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCollectionOptions 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuCollectionOptions 
         Caption         =   "Settings..."
         Index           =   13
      End
   End
   Begin VB.Menu mnuRandom 
      Caption         =   "&Random"
      Begin VB.Menu mnuRandomOptions 
         Caption         =   "Add &String(s)"
         Index           =   1
      End
      Begin VB.Menu mnuRandomOptions 
         Caption         =   "Add &Object(s)"
         Index           =   2
      End
      Begin VB.Menu mnuRandomOptions 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuRandomOptions 
         Caption         =   "Remove Item(s)"
         Index           =   4
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IFindHelper
Implements ICollectionPlusSite

'Menu Enums
Private Enum FileMenu
    mnuFileOpen = 1
    mnuFileSave
    mnuFileExit = 4
End Enum

Private Enum ViewMenu
    mnuViewAutoRefresh = 1
    mnuViewEdit = 3
    mnuViewRefresh
End Enum

Private Enum CollectionMenu
    mnuColAddString = 3
    mnuColAddObject
    mnuColFind = 6
    mnuColFindObject
    mnuColRemove = 10
    mnuColRemoveAll
    mnuColSettings = 13
End Enum

Private Enum RandomMenu
    mnuRandomAddString = 1
    mnuRandomAddObject
    mnuRandomRemoveItems = 4
End Enum

Private Enum ColType
    colTypeVariant = 1
    colTypeString
End Enum

Private Const LB_GETSELITEMS As Long = 401
Private Const LB_ITEMFROMPOINT As Long = 425

Private mColVar As CCollectionPlus
Attribute mColVar.VB_VarHelpID = -1
Private mColString As CCollectionString

Private mlColType As ColType
Private mbAutoRefresh As Boolean
Private masListKeys() As String
Private mbRefreshed As Boolean
Private mbDragging As Boolean

Private WithEvents mfrmSettings As frmSettings
Attribute mfrmSettings.VB_VarHelpID = -1
Private WithEvents mfrmAddItem As frmAddItem
Attribute mfrmAddItem.VB_VarHelpID = -1

Private Sub ColRandomAddObject()

    Dim i As Long
    Dim lItemCount As Long
    Dim oItem As CPerson
    Dim sKey As String
    Dim sTag As String
    
    lItemCount = GetNumItems("Please enter the number of objects to add")
    If lItemCount <= 0 Then Exit Sub
    
    For i = 1 To lItemCount
        Set oItem = New CPerson
        With oItem
            .FirstName = GetRandomString(5, 20)
            .Surname = GetRandomString(5, 20)
            .Address = GetRandomString(5, 20)
        End With
        sKey = GetRandomString(5, 20)
        sTag = GetRandomString(5, 20)
        
        mColVar.Add Value:=oItem, Key:=sKey, Tag:=sTag
    Next i
    
    If mbAutoRefresh Then
        ColRefresh
    Else
        mbRefreshed = False
    End If
    SetEnabled
    
Exit Sub

ColRandomAddStringErr:

    Select Case MapError(Err)
    Case 457        'Key exists
        'Regenerate key and try again
        sKey = GetRandomString(5, 20)
        Resume
    
    Case Else
        'Unexpected
        Debug.Assert False
    End Select

End Sub

'****************************************
'Procedure:     frmMain.ColRandomAddString
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          28 September 2000
'****************************************
'Description:
'****************************************
'Revisions:
'****************************************
Private Sub ColRandomAddString()

    On Error GoTo ColRandomAddStringErr

    Dim i As Long
    Dim lItemCount As Long
    Dim sItem As String
    Dim sKey As String
    Dim sTag As String
    
    lItemCount = GetNumItems("Please enter the number of string items to add")
    If lItemCount <= 0 Then Exit Sub
    
    For i = 1 To lItemCount
        sItem = GetRandomString(5, 20)
        sKey = GetRandomString(5, 20)
        sTag = GetRandomString(5, 20)
        
        If mlColType = colTypeString Then
            mColString.Add Value:=sItem, Key:=sKey, Tag:=sTag
        Else
            mColVar.Add Value:=sItem, Key:=sKey, Tag:=sTag
        End If
    Next i
    
    If mbAutoRefresh Then
        ColRefresh
    Else
        mbRefreshed = False
    End If
    SetEnabled
    
Exit Sub

ColRandomAddStringErr:

    Select Case MapError(Err)
    Case 457        'Key exists
        'Regenerate key and try again
        sKey = GetRandomString(5, 20)
        Resume
    
    Case Else
        'Unexpected
        Debug.Assert False
    End Select
    
End Sub

Private Sub ColRandomRemoveItems()

    Dim lTotalCount As Long
    Dim lItemCount As Long
    Dim i As Long
    Dim lItem As Long
    Dim vIndex As Variant
    Dim bUseKey As Boolean
    
    lItemCount = GetNumItems("Please enter the number of items to remove")
    If lItemCount <= 0 Then Exit Sub
    
    If mlColType = colTypeString Then
        lTotalCount = mColString.Count
    Else
        lTotalCount = mColVar.Count
    End If
    
    If lItemCount > lTotalCount Then
        lItemCount = lTotalCount
    End If
    
    For i = 1 To lItemCount
        lItem = Random(1, lTotalCount)
        bUseKey = Random(-1, 0)
        
        If bUseKey Then
            If mlColType = colTypeString Then
                vIndex = mColString.Key(lItem)
            Else
                vIndex = mColVar.Key(lItem)
            End If
        Else
            vIndex = lItem
        End If
        
        If mlColType = colTypeString Then
            mColString.Remove vIndex
        Else
            mColVar.Remove vIndex
        End If
        
        lTotalCount = lTotalCount - 1
    Next i

    If mbAutoRefresh Then
        ColRefresh
    Else
        mbRefreshed = False
    End If
    
    SetEnabled
        
End Sub


'****************************************
'Procedure:     frmMain.GetItemIndex
'****************************************
'Author:        Nick Hall
'****************************************
'Date:          10 October 2000
'****************************************
'Description:   Returns the position of a list index item in the
'collection
'****************************************
'Revisions:
'****************************************
Private Function GetItemIndex(ByVal lListIndex As Long) As Long

    If mlColType = colTypeString Then
        If mColString.EnumDirection = colEnumDirBackward Then
            GetItemIndex = mColString.Count - lListIndex
        Else
            GetItemIndex = lListIndex + 1
        End If
    Else
        If mColVar.EnumDirection = colEnumDirBackward Then
            GetItemIndex = mColVar.Count - lListIndex
        Else
            GetItemIndex = lListIndex + 1
        End If
    End If
    
End Function

Private Function GetRandomString(minChars As Long, maxChars As Long) As String
    Const CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    Dim i As Long
    
    GetRandomString = String$(minChars + Rnd * (maxChars - minChars), "0")
    For i = 1 To Len(GetRandomString)
        Mid$(GetRandomString, i, 1) = Mid$(CHARS, 1 + Len(CHARS) * Rnd, 1)
    Next i
    
End Function

Private Sub ColShowAddItem(ByVal lAddType As AddType, Optional ByVal vKey As Variant)
    
    Dim bNew As Boolean
    
    bNew = IsMissing(vKey)
    
    Set mfrmAddItem = New frmAddItem
    With mfrmAddItem
        .SetDisplay lAddType, bNew
        
        Select Case mlColType
        Case colTypeVariant
            Set .Collection = mColVar
            If Not bNew Then
                .Key = CStr(vKey)
                .ItemTag = mColVar.Tag(vKey)
                If lAddType = addTypeObject Then
                    Set .Item = mColVar(vKey)
                Else
                    .Item = mColVar(vKey)
                End If
            End If
        
        Case colTypeString
            Set .Collection = mColString
            
            If Not bNew Then
                .Key = CStr(vKey)
                .ItemTag = mColString.Tag(vKey)
                .Item = mColString(vKey)
            End If
        End Select
        .Show vbModal, Me
    End With
    
    Unload mfrmAddItem
    Set mfrmAddItem = Nothing
    
End Sub

Private Sub ColEditItem()

    Dim lIndex As Long
    Dim lAddType As AddType
    
    With lstItems
        lIndex = .ListIndex
        Debug.Assert lIndex >= 0
        lAddType = .ItemData(lIndex)
    End With
    
    ColShowAddItem lAddType, masListKeys(GetItemIndex(lIndex))
    
End Sub

Private Sub ColLoadFile()

    Dim sFileName As String
    
    On Error GoTo LoadErr
    
    With dlgMain
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir
        .ShowOpen
        sFileName = .FileName
    End With
    
    If LenB(sFileName) > 0 Then
        If mlColType = colTypeString Then
            mColString.Load sFileName
        Else
            mColVar.Load FileName:=sFileName, ColSite:=Me
        End If
        
        If mbAutoRefresh Then
            ColRefresh
        Else
            mbRefreshed = False
        End If
        
        SetEnabled
    End If

Exit Sub

LoadErr:
    If Err <> cdlCancel Then
        MsgBox "Unable to load collection: error " & CStr(MapError(Err)) & vbCrLf & vbCrLf & Err.Description
    End If
    
End Sub

Private Sub ColRefresh(Optional vSelectItem As Variant)

    Dim lIndex As Long
    Dim vItem As Variant
    Dim sKey As String
    
    SetRedraw lstItems, False
        
    With lstItems
        If IsMissing(vSelectItem) Then
            sKey = KeyFromListIndex(.ListIndex)
        Else
            sKey = CStr(vSelectItem)
        End If
        
        .Clear
        
        If mlColType = colTypeString Then
            For Each vItem In mColString
                .AddItem CStr(vItem)
                .ItemData(.NewIndex) = addTypeString
            Next vItem
            masListKeys = mColString.Keys
        Else
            For Each vItem In mColVar
                .AddItem CStr(vItem)
                If IsObject(vItem) Then
                    .ItemData(.NewIndex) = addTypeObject
                Else
                    .ItemData(.NewIndex) = addTypeString
                End If
            Next vItem
            masListKeys = mColVar.Keys
        End If
        
        lIndex = ListIndexFromKey(sKey)
        .ListIndex = lIndex
        If lIndex <> -1 Then
            .Selected(lIndex) = True
        End If
    End With
    
    SetRedraw lstItems, True
    mbRefreshed = True
            
End Sub

Private Sub ColRemove()

    Dim lIndex As Long
    
    With lstItems
        If .ListCount = 0 Then Exit Sub
        For lIndex = 0 To .ListCount - 1
            If .Selected(lIndex) Then
                If mlColType = colTypeString Then
                    mColString.Remove masListKeys(lIndex + 1)
                Else
                    mColVar.Remove masListKeys(lIndex + 1)
                End If
            End If
        Next lIndex
    End With
    
    If mbAutoRefresh Then
        ColRefresh
    Else
        mbRefreshed = False
    End If
    
    SetEnabled
    
End Sub

Private Sub ColRemoveAll()

    If mlColType = colTypeString Then
        mColString.Clear
    Else
        mColVar.Clear
    End If
    
    If mbAutoRefresh Then
        ColRefresh
    Else
        mbRefreshed = False
    End If
    
    SetEnabled
    
End Sub

Private Sub ColSaveFile()

    Dim sFileName As String
    
    On Error GoTo SaveErr
    
    With dlgMain
        .Flags = cdlOFNCreatePrompt Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt
        .ShowSave
        sFileName = .FileName
    End With
    
    If LenB(sFileName) > 0 Then
        If mlColType = colTypeString Then
            mColString.Save sFileName, True
        Else
            mColVar.Save sFileName, True
        End If
    End If

Exit Sub

SaveErr:
    If Err <> cdlCancel Then
        MsgBox "Unable to save collection: error " & CStr(MapError(Err)) & vbCrLf & vbCrLf & Err.Description
    End If
    
End Sub

Private Sub ColShowFind(ByVal lFindType As FindType)

    Dim frm As frmFindItem
    
    Set frm = New frmFindItem
    With frm
        .SetDisplay lFindType
        Set .FindHelper = Me
        .Show vbModal, Me
    End With
    
    Unload frm
    Set frm = Nothing
    
End Sub

Private Sub ColShowSettings()

    Set mfrmSettings = New frmSettings
    
    With mfrmSettings
        Select Case mlColType
        Case colTypeVariant
            Set .Collection = mColVar
        Case colTypeString
            Set .Collection = mColString
        End Select
        .Show vbModal, Me
    End With
    
    Unload mfrmSettings
    Set mfrmSettings = Nothing
    
End Sub

Private Sub ColExitApp()

    Unload Me
    
End Sub

Private Function GetNumItems(ByVal sMsg As String) As Long

    Dim sResponse As String
    Dim sError As String
    
    Do
        sResponse = InputBox(sMsg & sError)
        If Not IsNumeric(sResponse) Then
            sError = vbCr & vbCr & "(please enter a numeric value)"
        Else
            Exit Do
        End If
    Loop Until LenB(sResponse) = 0
    
    GetNumItems = CLng(Val(sResponse))
    
End Function

Private Function ListIndexFromKey(ByRef sKey As String) As Long
    
    If mlColType = colTypeString Then
        'Find out if the key still exists
        With mColString
            If .Exists(sKey) Then
                If .EnumDirection = colEnumDirBackward Then
                    ListIndexFromKey = .Count - .Index(sKey)
                Else
                    ListIndexFromKey = .Index(sKey) - 1
                End If
            Else
                ListIndexFromKey = -1
            End If
        End With
    Else
        With mColVar
            If .Exists(sKey) Then
                If .EnumDirection = colEnumDirBackward Then
                    ListIndexFromKey = .Count - .Index(sKey)
                Else
                    ListIndexFromKey = .Index(sKey) - 1
                End If
            Else
                ListIndexFromKey = -1
            End If
        End With
    End If
    
End Function

Private Sub InitData()

    Set mColString = New CCollectionString
    Set mColVar = New CCollectionPlus

    mbAutoRefresh = True
    mlColType = colTypeVariant
    
    Randomize Timer
    
End Sub

Private Function KeyFromListIndex(ByVal iIndex As Integer) As String

    Dim i As Long
    Dim asKeys() As String
    Dim sFind As String
    
    If iIndex = -1 Then Exit Function
    
    With lstItems
        If .ListCount = 0 Then Exit Function
        Debug.Assert iIndex >= 0 And iIndex <= .ListCount - 1
        
        sFind = .List(iIndex)
        
        If mlColType = colTypeString Then
            With mColString
                If .Count > 0 Then
                    asKeys = mColString.Keys
                    For i = 1 To UBound(asKeys)
                        If mColString(asKeys(i)) = sFind Then
                            KeyFromListIndex = asKeys(i)
                            Exit Function
                        End If
                    Next i
                End If
            End With
        Else
            With mColVar
                If .Count > 0 Then
                    asKeys = mColVar.Keys
                    For i = 1 To UBound(asKeys)
                        If mColVar(asKeys(i)) = sFind Then
                            KeyFromListIndex = asKeys(i)
                            Exit Function
                        End If
                    Next i
                End If
            End With
        End If
    End With
    
End Function

Private Function MapError(ByVal lErrNumber As Long) As Long

    If lErrNumber And vbObjectError Then lErrNumber = lErrNumber And (Not vbObjectError)
    MapError = lErrNumber
    
End Function

Private Sub SetEnabled()

    With mnuCollectionOptions
        If mlColType = colTypeString Then
            'Disable all options to do with manuipulating objects
            .Item(mnuColAddObject).Enabled = False
            .Item(mnuColFindObject).Enabled = False
        Else
            'Enable all options to do with manipulating objects
            .Item(mnuColAddObject).Enabled = True
            .Item(mnuColFindObject).Enabled = True
        End If
    End With
    
    If mlColType = colTypeString Then
        mnuRandomOptions(mnuRandomAddObject).Enabled = False
    Else
        mnuRandomOptions(mnuRandomAddObject).Enabled = True
    End If
    
    With mnuViewOptions(mnuViewEdit)
        If mbAutoRefresh Then
            .Enabled = True
        Else
            If mbRefreshed Then
                .Enabled = True
            Else
                .Enabled = False
            End If
        End If
    End With
    
    If lstItems.ListIndex = -1 Then
        mnuViewOptions(mnuViewEdit).Enabled = False
        mnuCollectionOptions(mnuColRemove).Enabled = False
    Else
        mnuViewOptions(mnuViewEdit).Enabled = True
        mnuCollectionOptions(mnuColRemove).Enabled = True
    End If
        
End Sub

Private Sub cmdExit_Click()

    ColExitApp
    
End Sub


Private Sub Form_Load()

    InitData
    
    SetEnabled
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set mColString = Nothing
    Set mColVar = Nothing
    Erase masListKeys
    
End Sub


Private Function ICollectionPlusSite_RestoreItem(Contents As Variant) As Object

    Dim oPerson As CPerson
    
    Set oPerson = New CPerson
    ICollectionPlusItem(oPerson).Contents = Contents
    Set ICollectionPlusSite_RestoreItem = oPerson

End Function

Private Function IFindHelper_FindObject(sFirstName As String, sSurname As String, sAddress As String) As CCollectionPlus

    'Only ever called against Variant Collection
    Set IFindHelper_FindObject = mColVar.FindItems(sSurname, sFirstName, sAddress)
    
End Function

Private Function IFindHelper_FindString(sFind As String) As Variant

    If mlColType = colTypeString Then
        Set IFindHelper_FindString = mColString.Find(sFind, vbTextCompare)
    Else
        Set IFindHelper_FindString = mColVar.Find(sFind, vbTextCompare)
    End If
    
End Function


Private Sub lstItems_Click()

    SetEnabled
    
End Sub

Private Sub lstItems_DblClick()

    ColEditItem
    
End Sub


Private Sub lstItems_DragDrop(Source As Control, X As Single, Y As Single)

    Dim iPixelX As Integer
    Dim iPixelY As Integer
    Dim lOldIndex As Long
    Dim lNewIndex As Long
    Dim lParam As Long
    Dim rc As Long
    Dim sItemKey As String
    Dim sAfterKey As String
    
    lOldIndex = CLng(Source.Tag)
    
    'Get item at X,Y - need to combine integers into one long for SendMessage call
    iPixelX = ScaleX(X, vbTwips, vbPixels)
    iPixelY = ScaleY(Y, vbTwips, vbPixels)
    lParam = MakeDWord(iPixelX, iPixelY)
    
    rc = SendMessage(lstItems.hWnd, LB_ITEMFROMPOINT, 0, ByVal lParam)
    
    'Get new index from return
    lNewIndex = LoWord(rc)
    
    If lOldIndex = lNewIndex Then Exit Sub
    sItemKey = masListKeys(GetItemIndex(lOldIndex))
    sAfterKey = masListKeys(GetItemIndex(lNewIndex))
    
    'Perform the move
    If mlColType = colTypeString Then
        mColString.Move Index:=sItemKey, After:=sAfterKey
    Else
        mColVar.Move Index:=sItemKey, After:=sAfterKey
    End If
    
    If mbAutoRefresh Then
        ColRefresh
    Else
        mbRefreshed = False
    End If
    
    SetEnabled
    mbDragging = False
    
End Sub

Private Sub lstItems_KeyPress(KeyAscii As Integer)

    With lstItems
        Select Case KeyAscii
        Case vbKeySeparator, vbKeyReturn
            If .ListCount > 0 Then
                If .Selected(.ListIndex) Then
                    ColEditItem
                End If
            End If
        
        Case vbKeyBack
            ColRemove
            
        End Select
    End With
    
End Sub

Private Sub lstItems_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyDelete
        ColRemove
        
    End Select
    
End Sub

Private Sub lstItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mbDragging = False
        
End Sub

Private Sub lstItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim lSelIndex As Long
    
    If Button = vbLeftButton And Shift = 0 Then
        If (mbAutoRefresh Or mbRefreshed) And Not mbDragging Then
            With lstItems
                'Only allow dragging of one item
                If .SelCount = 1 Then
                    'Determine selected item
                    Call SendMessage(.hWnd, LB_GETSELITEMS, 1, lSelIndex)
                    
                    With lblDrag
                        .Caption = lstItems.List(lSelIndex)
                        .Left = X
                        .Top = Y
                        .Tag = CStr(lSelIndex)
                        .Visible = True
                        .Drag vbBeginDrag
                        mbDragging = True
                    End With
                End If
            End With
        End If
    End If
            
End Sub

Private Sub lstItems_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        Me.PopupMenu mnuCollection
    End If
    
End Sub

Private Sub mfrmAddItem_AddItem(Item As Variant, Key As String, Before As Variant, After As Variant, Tag As String, Cancel As Boolean)

    Dim bAdd As Boolean
    
    bAdd = True
    
    If mlColType = colTypeString Then
        If mColString.Exists(Key) Then
            If MsgBox("An item with key '" & Key & "' already exists.  Do you want to replace it?", vbYesNo + vbQuestion) = vbYes Then
                mColString(Key) = Item
            Else
                Cancel = True
            End If
            bAdd = False
        End If
        
        If bAdd Then
            mColString.Add CStr(Item), Key, Tag, Before, After
        End If
    Else
        If mColVar.Exists(Key) Then
            If MsgBox("An item with key '" & Key & "' already exists.  Do you want to replace it?", vbYesNo + vbQuestion) = vbYes Then
                If IsObject(Item) Then
                    Set mColVar(Key) = Item
                Else
                    mColVar(Key) = Item
                End If
            Else
                Cancel = True
            End If
            bAdd = False
        End If
        
        If bAdd Then
            mColVar.Add Item, Key, Tag, Before, After
        End If
    End If
        
    If mbAutoRefresh Then
        ColRefresh Key
    End If
    
    SetEnabled

End Sub

Private Sub mfrmAddItem_UpdateItem(OldItem As Variant, NewItem As Variant, OldKey As String, NewKey As String, OldTag As String, NewTag As String, Cancel As Boolean)

    If mlColType = colTypeString Then
        If OldKey <> NewKey Then
            If mColString.Exists(NewKey) Then
                MsgBox "An item with key '" & NewKey & "' already exists in this collection", vbExclamation
                Cancel = True
                Exit Sub
            Else
                mColString.Key(OldKey) = NewKey
            End If
        End If
        
        If NewItem <> OldItem Then
            mColString(NewKey) = NewItem
        End If
        
        If NewTag <> OldTag Then
            mColString.Tag(NewKey) = NewTag
        End If
    Else
        If OldKey <> NewKey Then
            If mColVar.Exists(NewKey) Then
                MsgBox "An item with key '" & NewKey & "' already exists in this collection", vbExclamation
                Cancel = True
                Exit Sub
            Else
                mColVar.Key(OldKey) = NewKey
            End If
        End If
        
        If IsObject(NewItem) Then
            If ICollectionPlusItem(NewItem).Compare(OldItem, vbTextCompare) <> 0 Then
                Set mColVar(NewKey) = NewItem
            End If
        Else
            If NewItem <> OldItem Then
                mColVar(NewKey) = NewItem
            End If
        End If
        
        If NewTag <> OldTag Then
            mColVar.Tag(NewKey) = NewTag
        End If
    End If
    
    If mbAutoRefresh Then
        ColRefresh NewKey
    End If
    
    SetEnabled

End Sub


Private Sub mfrmSettings_Changed()

    'Do a simple refresh
    If mbAutoRefresh Then
        ColRefresh
    Else
        mbRefreshed = False
    End If
    
    SetEnabled
    
End Sub

Private Sub mnuCollectionOptions_Click(Index As Integer)

    Select Case Index
    Case mnuColAddString
        ColShowAddItem addTypeString
    
    Case mnuColAddObject
        ColShowAddItem addTypeObject
        
    Case mnuColSettings
        ColShowSettings
    
    Case mnuColFind
        ColShowFind fndTypeString
    
    Case mnuColFindObject
        ColShowFind fndTypeObject
        
    Case mnuColRemove
        ColRemove
    
    Case mnuColRemoveAll
        ColRemoveAll
        
    End Select
    
End Sub

Private Sub mnuCollectionSort_Click(Index As Integer)

    Select Case mlColType
    Case colTypeVariant
        mColVar.Sort Index
    
    Case colTypeString
        mColString.Sort Index
    End Select
    
    If mbAutoRefresh Then
        ColRefresh
    Else
        mbRefreshed = False
    End If
    
    SetEnabled
    
End Sub

Private Sub mnuCollectionType_Click(Index As Integer)

    Dim i As ColType
    
    If mlColType <> Index Then
        For i = colTypeVariant To colTypeString
            mnuCollectionType(i).Checked = (i = Index)
        Next i
                
        mlColType = Index
        
        If mbAutoRefresh Then
            ColRefresh
        Else
            mbRefreshed = False
        End If
        SetEnabled
    End If
    
End Sub

Private Sub mnuFileOptions_Click(Index As Integer)

    Select Case Index
    Case mnuFileOpen
        ColLoadFile
        
    Case mnuFileSave
        ColSaveFile
        
    Case mnuFileExit
        ColExitApp
        
    End Select
    
End Sub

Private Sub mnuHelpAbout_Click()
    
    Dim ofrm As frmAbout
    
    Set ofrm = New frmAbout
    ofrm.Show vbModal, Me
    
    Unload ofrm
    Set ofrm = Nothing
    
End Sub


Private Sub mnuRandomOptions_Click(Index As Integer)

    Select Case Index
    Case mnuRandomAddString
        ColRandomAddString
        
    Case mnuRandomAddObject
        ColRandomAddObject
        
    Case mnuRandomRemoveItems
        ColRandomRemoveItems
        
    End Select
    
End Sub

Private Sub mnuViewOptions_Click(Index As Integer)

    Select Case Index
    Case mnuViewAutoRefresh
        mbAutoRefresh = Not mbAutoRefresh
        mnuViewOptions(Index).Checked = mbAutoRefresh
        SetEnabled
    
    Case mnuViewEdit
        ColEditItem
    
    Case mnuViewRefresh
        ColRefresh
        SetEnabled
        
    End Select
    
End Sub


