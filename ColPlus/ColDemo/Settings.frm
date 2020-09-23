VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Collection Settings"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   400
      Left            =   4680
      TabIndex        =   9
      Top             =   4620
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   400
      Left            =   3420
      TabIndex        =   8
      Top             =   4620
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2160
      TabIndex        =   7
      Top             =   4620
      Width           =   1200
   End
   Begin VB.Frame fraSize 
      Caption         =   "Size"
      Height          =   2655
      Left            =   60
      TabIndex        =   12
      Top             =   1860
      Width           =   5835
      Begin VB.TextBox txtItemChunk 
         Height          =   330
         Left            =   4440
         MaxLength       =   6
         TabIndex        =   6
         ToolTipText     =   "Amount by which the item lookup table increases when needed"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtItemInit 
         Height          =   330
         Left            =   4440
         MaxLength       =   6
         TabIndex        =   5
         ToolTipText     =   "Initial size of the item lookup table"
         Top             =   1620
         Width           =   1215
      End
      Begin VB.TextBox txtHashLookupChunk 
         Height          =   330
         Left            =   4440
         MaxLength       =   6
         TabIndex        =   4
         ToolTipText     =   "Amount the hash lookup table increases when needed"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtHashLookupInit 
         Height          =   330
         Left            =   4440
         MaxLength       =   6
         TabIndex        =   3
         ToolTipText     =   "Initial size of the hash lookup table"
         Top             =   780
         Width           =   1215
      End
      Begin VB.TextBox txtHashSize 
         Height          =   330
         Left            =   4440
         MaxLength       =   6
         TabIndex        =   2
         ToolTipText     =   "Size of the hash table"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblGen 
         Caption         =   "Item Chunk Size:"
         Height          =   315
         Index           =   5
         Left            =   150
         TabIndex        =   17
         Top             =   2100
         Width           =   1815
      End
      Begin VB.Label lblGen 
         Caption         =   "Item Initial Size:"
         Height          =   315
         Index           =   4
         Left            =   150
         TabIndex        =   16
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblGen 
         Caption         =   "Hash Lookup Chunk Size:"
         Height          =   315
         Index           =   3
         Left            =   150
         TabIndex        =   15
         Top             =   1260
         Width           =   1995
      End
      Begin VB.Label lblGen 
         Caption         =   "Hash Lookup Initial Size:"
         Height          =   315
         Index           =   2
         Left            =   150
         TabIndex        =   14
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblGen 
         Caption         =   "Hash Size:"
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   13
         Top             =   420
         Width           =   1815
      End
   End
   Begin VB.Frame fraBehaviour 
      Caption         =   "Behaviour Settings"
      Height          =   1695
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   5835
      Begin VB.CheckBox chkMatchCase 
         Alignment       =   1  'Right Justify
         Caption         =   "Match Case:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Determines whether or not the collection matches its keys case-sensitive"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.ComboBox cmbFillDir 
         Height          =   315
         ItemData        =   "Settings.frx":0000
         Left            =   1680
         List            =   "Settings.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Determine in which direction the collection will enumerate"
         Top             =   375
         Width           =   3915
      End
      Begin VB.Label lblGen 
         Caption         =   "Fill Direction:"
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   11
         Top             =   420
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbChanged As Boolean
Private mbInProgress As Boolean

' Collection object to change setttings on
Private moCollection As CCollectionPlus ' Added by: Nick Hall on 08/09/2000 1:57:43 pm

Event Changed()

Public Property Get Collection() As CCollectionPlus

    Set Collection = moCollection

End Property
Public Property Set Collection(ByVal NewCollection As CCollectionPlus)

    If Not moCollection Is NewCollection Then
        Set moCollection = NewCollection
        
        SettingsFill
    End If

End Property

Private Sub SetEnabled()

    cmdApply.Enabled = mbChanged
    
End Sub

Private Sub SettingsApply()

    With moCollection
        .EnumDirection = cmbFillDir.ItemData(cmbFillDir.ListIndex)
        .MatchCase = IIf(chkMatchCase = vbChecked, True, False)
    End With
    
    ICollectionPlusSettings(moCollection).ChangeSettings CLng(txtHashSize), CLng(txtHashLookupInit), CLng(txtHashLookupChunk), CLng(txtItemInit), CLng(txtItemChunk)
    
    RaiseEvent Changed
    mbChanged = False
    
    SetEnabled
    
End Sub

Private Function SettingsCancel() As Boolean

    SettingsCancel = False
    
    If mbChanged Then
        If MsgBox("Are you sure you wish to cancel (all changes will be lost)?", vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    SettingsCancel = True
    
End Function

Private Sub SettingsFill()

    'Sanity check
    If moCollection Is Nothing Then Exit Sub
    
    mbInProgress = True
    
    'Behaviour
    With moCollection
        ListSelect cmbFillDir, .EnumDirection
        chkMatchCase = IIf(.MatchCase, vbChecked, vbUnchecked)
        If .Count > 0 Then
            'Can't change match case when there are items in the collection
            chkMatchCase.Enabled = False
        End If
    End With
    
    'Size
    With ICollectionPlusSettings(moCollection)
        txtHashSize = .HashSize
        txtHashLookupInit = .HashLookupInitialSize
        txtHashLookupChunk = .HashLookupChunkSize
        txtItemInit = .ItemInitialSize
        txtItemChunk = .ItemChunkSize
    End With
    
    mbInProgress = False
    
End Sub


Private Function SettingsValid() As Boolean

    Dim lValue As Long
    
    If Not IsNumeric(txtHashSize) Then
        ValidateError txtHashSize, "Please enter a valid numeric value for hash size"
        Exit Function
    End If
    
    lValue = CLng(txtHashSize)
    If lValue <= 0 Then
        ValidateError txtHashSize, "Hash size must be greater than 0"
        Exit Function
    End If
    
    If lValue > 0 Then
        If moCollection.Count > 0 Then
            With ICollectionPlusSettings(moCollection)
                If .HashSize <> lValue Then
                    ValidateError txtHashSize, "Hash size cannot be changed when there are items in the collection"
                    txtHashSize = CStr(.HashSize)
                    Exit Function
                End If
            End With
        End If
    End If
    
    If Not IsNumeric(txtHashLookupInit) Then
        ValidateError txtHashLookupInit, "Please enter a valid numeric value for hash lookup initial size"
        Exit Function
    End If
    
    lValue = CLng(txtHashLookupInit)
    If lValue <= 0 Then
        ValidateError txtHashLookupInit, "Hash lookup initial size must be greater than 0"
        Exit Function
    End If
    
    If Not IsNumeric(txtHashLookupChunk) Then
        ValidateError txtHashLookupChunk, "Please enter a valid numeric value for hash lookup chunk size"
        Exit Function
    End If
    
    lValue = CLng(txtHashLookupChunk)
    If lValue <= 0 Then
        ValidateError txtHashSize, "Hash lookup chunk size must be greater than 0"
        Exit Function
    End If
    
    If Not IsNumeric(txtItemInit) Then
        ValidateError txtItemInit, "Please enter a valid numeric value for item initial size"
        Exit Function
    End If
    
    lValue = CLng(txtItemInit)
    If lValue <= 0 Then
        ValidateError txtItemInit, "Item initial size must be greater than 0"
        Exit Function
    End If
    
    If Not IsNumeric(txtItemChunk) Then
        ValidateError txtHashSize, "Please enter a valid numeric value for item chunk size"
        Exit Function
    End If
    
    lValue = CLng(txtItemChunk)
    If lValue <= 0 Then
        ValidateError txtItemChunk, "Item chunk size must be greater than 0"
        Exit Function
    End If
    
    SettingsValid = True
    
End Function

Private Sub chkMatchCase_Click()

    If Not mbInProgress Then
        mbChanged = True
        SetEnabled
    End If
    
End Sub


Private Sub cmbFillDir_Click()

    If Not mbInProgress Then
        mbChanged = True
        SetEnabled
    End If

End Sub


Private Sub cmdApply_Click()

    If SettingsValid() Then
        SettingsApply
    End If
    
End Sub

Private Sub cmdCancel_Click()

    If SettingsCancel() Then
        Me.Hide
    End If
    
End Sub

Private Sub cmdOK_Click()

    If SettingsValid() Then
        SettingsApply
        Me.Hide
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode <> vbFormCode Then
        Cancel = True
        
        If SettingsCancel() Then Me.Hide
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set moCollection = Nothing
    
End Sub


Private Sub txtHashLookupChunk_Change()

    If Not mbInProgress Then
        mbChanged = True
        SetEnabled
    End If

End Sub


Private Sub txtHashLookupInit_Change()

    If Not mbInProgress Then
        mbChanged = True
        SetEnabled
    End If

End Sub


Private Sub txtHashSize_Change()

    If Not mbInProgress Then
        mbChanged = True
        SetEnabled
    End If

End Sub


Private Sub txtItemChunk_Change()

    If Not mbInProgress Then
        mbChanged = True
        SetEnabled
    End If

End Sub


Private Sub txtItemInit_Change()

    If Not mbInProgress Then
        mbChanged = True
        SetEnabled
    End If

End Sub


