VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFindItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Items"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   400
      Left            =   4380
      TabIndex        =   7
      Top             =   7560
      Width           =   1200
   End
   Begin VB.Frame fraResults 
      Caption         =   "Results"
      Height          =   2895
      Left            =   60
      TabIndex        =   14
      Top             =   4560
      Width           =   5535
      Begin MSFlexGridLib.MSFlexGrid grdResults 
         Height          =   1995
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3519
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowUserResizing=   1
      End
      Begin VB.Label lblResultCount 
         Height          =   350
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   5055
      End
   End
   Begin VB.Frame fraObjFind 
      Caption         =   "Find"
      Height          =   3075
      Left            =   60
      TabIndex        =   10
      Top             =   1440
      Width           =   5535
      Begin VB.TextBox txtAddress 
         Height          =   1230
         Left            =   1140
         TabIndex        =   4
         Top             =   1260
         Width           =   4155
      End
      Begin VB.TextBox txtSurname 
         Height          =   330
         Left            =   1140
         TabIndex        =   3
         Top             =   840
         Width           =   4155
      End
      Begin VB.TextBox txtFirstName 
         Height          =   330
         Left            =   1140
         TabIndex        =   2
         Top             =   420
         Width           =   4155
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   350
         Index           =   1
         Left            =   4200
         TabIndex        =   5
         Top             =   2580
         Width           =   1095
      End
      Begin VB.Label lblGen 
         Caption         =   "Address:"
         Height          =   375
         Index           =   3
         Left            =   180
         TabIndex        =   13
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label lblGen 
         Caption         =   "Surname:"
         Height          =   375
         Index           =   2
         Left            =   180
         TabIndex        =   12
         Top             =   900
         Width           =   915
      End
      Begin VB.Label lblGen 
         Caption         =   "First Name:"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   480
         Width           =   915
      End
   End
   Begin VB.Frame fraStringFind 
      Caption         =   "Find"
      Height          =   1335
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   5535
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   350
         Index           =   0
         Left            =   4200
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtItem 
         Height          =   330
         Left            =   1140
         TabIndex        =   0
         Top             =   420
         Width           =   4155
      End
      Begin VB.Label lblGen 
         Caption         =   "Item:"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmFindItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum FindType
    fndTypeString
    fndTypeObject
End Enum

' Reference to helper object that will perform the actual find
Private moFindHelper As IFindHelper ' Added by: Nick Hall on 22/09/2000 10:35:54 am

' Determine type of find
Private mlFindType As FindType ' Added by: Nick Hall on 22/09/2000 10:37:33 am

Private mbFirstTime As Boolean
Private Sub FindDisplay()

    Dim fraItem As Frame
    
    If mlFindType = fndTypeString Then
        fraStringFind.Visible = True
        fraObjFind.Visible = False
        Set fraItem = fraStringFind
    Else
        fraObjFind.Visible = True
        fraStringFind.Visible = False
        Set fraItem = fraObjFind
    End If
    
    fraItem.Top = 60
    
    With fraResults
        .Top = fraItem.Top + fraItem.Height + 100
        cmdClose.Top = .Top + .Height + 100
    End With
    
    With cmdClose
        Me.Height = .Top + .Height + 500
    End With
    
    InitGrid
    
End Sub

Private Sub FindObjects()

    Dim i As Long
    Dim oPerson As CPerson
    Dim oResults As CCollectionPlus
    
    ClearGrid grdResults
    
    Set oResults = moFindHelper.FindObject(txtFirstName.Text, txtSurname.Text, txtAddress.Text)
    If oResults.Count > 0 Then
        With grdResults
            .Rows = oResults.Count + 1
            For Each oPerson In oResults
                i = i + 1
                .TextMatrix(i, 0) = oPerson.FirstName
                .TextMatrix(i, 1) = oPerson.Surname
                .TextMatrix(i, 2) = oPerson.Address
            Next oPerson
        End With
    Else
        grdResults.Rows = 1
    End If
    
    lblResultCount = CStr(oResults.Count) & " items found"
    
End Sub

Private Sub FindStrings()

    Dim i As Long
    Dim vItem As Variant
    Dim oResults As CCollectionPlus
    
    ClearGrid grdResults
    
    Set oResults = moFindHelper.FindString(txtItem)
    If oResults.Count > 0 Then
        With grdResults
            .Rows = oResults.Count + 1
            For Each vItem In oResults
                i = i + 1
                .TextMatrix(i, 0) = CStr(vItem)
            Next vItem
        End With
    Else
        grdResults.Rows = 1
    End If
    
    lblResultCount = CStr(oResults.Count) & " items found"
    
End Sub

Public Property Get FindType() As FindType

    FindType = mlFindType

End Property

Public Property Get FindHelper() As IFindHelper

    Set FindHelper = moFindHelper

End Property
Public Property Set FindHelper(ByVal NewFindHelper As IFindHelper)

    Set moFindHelper = NewFindHelper

End Property

Private Sub InitGrid()

    Dim i As Long
    
    With grdResults
        If mlFindType = fndTypeString Then
            .Cols = 1
            .Rows = 1
            .ColWidth(0) = 3000
            .TextMatrix(0, 0) = "Item"
        Else
            .Cols = 3
            .Rows = 1
            
            For i = 0 To 2
                .ColWidth(i) = 1500
            Next i
            
            .TextMatrix(0, 0) = "First Name"
            .TextMatrix(0, 1) = "Surname"
            .TextMatrix(0, 2) = "Address"
        End If
    End With
            
End Sub

Public Sub SetDisplay(ByVal lFindType As FindType)

    If lFindType <> mlFindType Or mbFirstTime Then
        mlFindType = lFindType
        FindDisplay
        mbFirstTime = False
    End If
        
End Sub

Private Sub cmdClose_Click()

    Me.Hide
    
End Sub

Private Sub cmdFind_Click(Index As Integer)

    Select Case Index
    Case fndTypeString
        FindStrings
    Case fndTypeObject
        FindObjects
    End Select
    
End Sub


Private Sub Form_Initialize()

    mbFirstTime = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set moFindHelper = Nothing
    
End Sub


