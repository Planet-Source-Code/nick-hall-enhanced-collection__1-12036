Attribute VB_Name = "modMain"
Option Explicit

Private mfrmMain As frmMain
Public Sub Main()

    If App.StartMode = vbSModeStandalone Then
        Set mfrmMain = New frmMain
        mfrmMain.Show
    End If
    
End Sub


Public Sub MTidyUp()

    If Not mfrmMain Is Nothing Then
        Unload mfrmMain
        Set mfrmMain = Nothing
    End If
    
End Sub


