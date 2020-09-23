VERSION 5.00
Begin VB.Form frmAddItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Item"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraObjItem 
      Caption         =   "Item"
      Height          =   3435
      Left            =   60
      TabIndex        =   17
      Top             =   1920
      Width           =   5895
      Begin VB.TextBox txtTag 
         Height          =   330
         Index           =   1
         Left            =   1200
         TabIndex        =   7
         Top             =   2880
         Width           =   4515
      End
      Begin VB.TextBox txtAddress 
         Height          =   1230
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1140
         Width           =   4515
      End
      Begin VB.TextBox txtSurname 
         Height          =   330
         Left            =   1200
         TabIndex        =   4
         Top             =   720
         Width           =   4515
      End
      Begin VB.TextBox txtFirstName 
         Height          =   330
         Left            =   1200
         TabIndex        =   3
         Top             =   300
         Width           =   4515
      End
      Begin VB.TextBox txtKey 
         Height          =   330
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Top             =   2460
         Width           =   4515
      End
      Begin VB.Label lblGen 
         Caption         =   "Tag:"
         Height          =   375
         Index           =   8
         Left            =   180
         TabIndex        =   23
         Top             =   2940
         Width           =   495
      End
      Begin VB.Label lblGen 
         Caption         =   "Address:"
         Height          =   375
         Index           =   6
         Left            =   180
         TabIndex        =   21
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblGen 
         Caption         =   "Surname:"
         Height          =   375
         Index           =   5
         Left            =   180
         TabIndex        =   20
         Top             =   780
         Width           =   855
      End
      Begin VB.Label lblGen 
         Caption         =   "First Name:"
         Height          =   375
         Index           =   4
         Left            =   180
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblGen 
         Caption         =   "Key:"
         Height          =   375
         Index           =   3
         Left            =   180
         TabIndex        =   18
         Top             =   2520
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdAddItem 
      Caption         =   "Add Item"
      Default         =   -1  'True
      Height          =   400
      Left            =   3420
      TabIndex        =   10
      Top             =   6660
      Width           =   1200
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   400
      Left            =   4680
      TabIndex        =   11
      Top             =   6660
      Width           =   1200
   End
   Begin VB.Frame fraPosition 
      Caption         =   "Position"
      Height          =   1095
      Left            =   60
      TabIndex        =   14
      Top             =   5460
      Width           =   5835
      Begin VB.ComboBox cmbItem 
         Height          =   315
         ItemData        =   "AddItem.frx":0000
         Left            =   3300
         List            =   "AddItem.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   420
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.ComboBox cmbPosition 
         Height          =   315
         ItemData        =   "AddItem.frx":0033
         Left            =   1200
         List            =   "AddItem.frx":0040
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   420
         Width           =   1935
      End
      Begin VB.Label lblGen 
         Caption         =   "Position in collection:"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame fraStringItem 
      Caption         =   "Item"
      Height          =   1755
      Left            =   60
      TabIndex        =   12
      Top             =   60
      Width           =   5895
      Begin VB.TextBox txtTag 
         Height          =   330
         Index           =   0
         Left            =   1200
         TabIndex        =   2
         Top             =   1140
         Width           =   4515
      End
      Begin VB.TextBox txtKey 
         Height          =   330
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   4515
      End
      Begin VB.TextBox txtItem 
         Height          =   330
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   4515
      End
      Begin VB.Label lblGen 
         Caption         =   "Tag:"
         Height          =   375
         Index           =   7
         Left            =   180
         TabIndex        =   22
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblGen 
         Caption         =   "Key:"
         Height          =   375
         Index           =   2
         Left            =   180
         TabIndex        =   16
         Top             =   780
         Width           =   495
      End
      Begin VB.Label lblGen 
         Caption         =   "Item:"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmAddItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Event AddItem(ByRef Item As Variant, ByRef Key As String, ByRef Before As Variant, ByRef After As Variant, ByRef Tag As String, ByRef Cancel As Boolean)
Event UpdateItem(ByRef OldItem As Variant, ByRef NewItem As Variant, ByRef OldKey As String, ByRef NewKey As String, ByRef OldTag As String, ByRef NewTag As String, ByRef Cancel As Boolean)

Public Enum AddType
    addTypeString
    addTypeObject
End Enum

Private Enum AddPosition
    addPosDefault
    addPosBefore
    addPosAfter
End Enum

' Current collection
Private moCollection As CCollectionPlus ' Added by: Nick Hall on 12/09/2000 11:57:44 am

Private mbNew As Boolean

' Current item value
Private mvItem As Variant ' Added by: Nick Hall on 20/09/2000 9:26:51 am

' Associated key
Private msKey As String ' Added by: Nick Hall on 20/09/2000 9:27:24 am

' Determine type of item to be added
Private mlAddType As AddType ' Added by: Nick Hall on 21/09/2000 9:04:15 am

Private mbFirstTime As Boolean

' Item's associated tag
Private msItemTag As String ' Added by: Nick Hall on 28/09/2000 2:14:49 pm
Public Property Get ItemTag() As String

    ItemTag = msItemTag

End Property
Public Property Let ItemTag(ByVal NewItemTag As String)

    If msItemTag <> NewItemTag Then
        msItemTag = NewItemTag
        txtTag(mlAddType) = NewItemTag
    End If

End Property

Public Property Get AddType() As AddType

    AddType = mlAddType

End Property
Private Sub AddDisplay()

    Dim fraItem As Frame
    Dim fraBottom As Frame
    
    Select Case mlAddType
    Case addTypeString
        Set fraItem = fraStringItem
        fraObjItem.Visible = False
        cmdAddItem.Default = True
        
    Case addTypeObject
        Set fraItem = fraObjItem
        fraStringItem.Visible = False
        cmdAddItem.Default = False
        
    End Select
    
    'Position item frame
    fraItem.Top = 60

    If mbNew Then
        Me.Caption = "Add Item"
        Set fraBottom = fraPosition
        With fraPosition
            .Visible = True
            .Top = fraItem.Top + fraItem.Height + 100
        End With
        
        cmdAddItem.Caption = "Add Item"
    Else
        Me.Caption = "Update Item"
        Set fraBottom = fraItem
        fraPosition.Visible = False
        cmdAddItem.Caption = "Update Item"
    End If
    
    With cmdAddItem
        .Top = fraBottom.Top + fraBottom.Height + 100
        cmdClose.Top = .Top
        Me.Height = .Top + .Height + 500
    End With
    
End Sub

Public Property Set Item(ByVal NewItem As Variant)

    Dim oPerson As CPerson
    
    Set mvItem = NewItem
    Set oPerson = NewItem
    
    If Not oPerson Is Nothing Then
        With oPerson
            txtFirstName = .FirstName
            txtSurname = .Surname
            txtAddress = .Address
        End With
    End If
    
End Property

Public Property Get Key() As String

    Key = msKey

End Property
Public Property Let Key(ByVal NewKey As String)

    If msKey <> NewKey Then
        msKey = NewKey
        txtKey(mlAddType) = NewKey
    End If

End Property


Public Property Get Item() As Variant

    Item = mvItem

End Property
Public Property Let Item(ByVal NewItem As Variant)

    If mvItem <> NewItem Then
        mvItem = NewItem
        txtItem = NewItem
    End If

End Property

Private Sub AddFillList()

    Dim v As Variant
    
    If Not mbNew Then Exit Sub
    
    SetRedraw cmbItem, False
    
    With cmbItem
        .Clear
        
        For Each v In moCollection
            .AddItem CStr(v)
        Next v
        
        .ListIndex = -1
    End With
    
    SetRedraw cmbItem, True
        
End Sub

Private Function AddItem() As Boolean

    Dim vItem As Variant
    Dim lPosition As AddPosition
    Dim sKey As String
    Dim sTag As String
    Dim vBefore As Variant
    Dim vAfter As Variant
    Dim bCancel As Boolean
    
    'Set variant arguments as Missing by default
    vBefore = Missing
    vAfter = Missing
    
    Select Case mlAddType
    Case addTypeString
        vItem = txtItem
    
    Case addTypeObject
        Dim oPerson As CPerson
        
        Set oPerson = New CPerson
        
        With oPerson
            .FirstName = txtFirstName
            .Surname = txtSurname
            .Address = txtAddress
        End With
        Set vItem = oPerson
    End Select
    
    sKey = txtKey(mlAddType)
    sTag = txtTag(mlAddType)
    
    lPosition = cmbPosition.ListIndex
    If Among(lPosition, addPosBefore, addPosAfter) Then
        With cmbItem
            If .ListIndex = -1 Then
                ValidateError cmbItem, "Please select a relative item"
                Exit Function
            End If
            
            Select Case lPosition
            Case addPosBefore
                vBefore = moCollection.Key(.ListIndex + 1)
            
            Case addPosAfter
                vAfter = moCollection.Key(.ListIndex + 1)
            End Select
        End With
    End If
    
    bCancel = False
    RaiseEvent AddItem(vItem, sKey, vBefore, vAfter, sTag, bCancel)
    AddItem = Not bCancel
    
End Function

Public Property Get Collection() As CCollectionPlus

    Set Collection = moCollection

End Property
Public Property Set Collection(ByVal NewCollection As CCollectionPlus)

    If Not (NewCollection Is moCollection) Then
        Set moCollection = NewCollection
        
        AddFillList
    End If

End Property

Public Sub SetDisplay(ByVal lAddType As AddType, ByVal bNew As Boolean)

    If lAddType <> mlAddType Or bNew <> mbNew Or mbFirstTime Then
        mlAddType = lAddType
        mbNew = bNew
        AddDisplay
        mbFirstTime = False
    End If
    
End Sub

Private Function UpdateItem() As Boolean

    Dim vItem As Variant
    Dim sKey As String
    Dim sTag As String
    Dim bCancel As Boolean
    
    Select Case mlAddType
    Case addTypeString
        vItem = txtItem

    Case addTypeObject
        Dim oPerson As CPerson
        Set oPerson = New CPerson
        With oPerson
            .FirstName = txtFirstName
            .Surname = txtSurname
            .Address = txtAddress
        End With
        
        Set vItem = oPerson
    End Select
    
    sKey = txtKey(mlAddType)
    sTag = txtTag(mlAddType)
    bCancel = False
    
    RaiseEvent UpdateItem(mvItem, vItem, msKey, sKey, msItemTag, sTag, bCancel)
    UpdateItem = Not bCancel
    
End Function

Private Sub cmbPosition_Click()

    Dim lIndex As Long
    
    With cmbPosition
        lIndex = .ListIndex
        If lIndex = -1 Then Exit Sub
        
        If .ItemData(lIndex) = addPosDefault Then
            cmbItem.Visible = False
        Else
            cmbItem.Visible = True
        End If
    End With
    
End Sub


Private Sub cmdAddItem_Click()

    If mbNew Then
        If Not AddItem() Then Exit Sub
    Else
        If Not UpdateItem() Then Exit Sub
    End If
    
    Me.Hide
    
End Sub

Private Sub cmdClose_Click()

    Me.Hide
    
End Sub


Private Sub Form_Initialize()

    mbFirstTime = True
    
End Sub

Private Sub Form_Load()

    cmbPosition.ListIndex = 0
    
End Sub



