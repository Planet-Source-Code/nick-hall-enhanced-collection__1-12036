VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEnumVariant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Taken directly, with very slight modifcation, from the version included
'with Bruce McKinney's Hardcore VB for enumerating variants

Option Explicit

' Implement VB version of IEnumVARIANT (from type library interface)
Implements IVBEnumVARIANT

Public Enum EErrorEnumVariant
    eeBaseEnumVariant = 13050   ' CEnumVariant
End Enum

Private connect As IVariantWalker

' Delegators must connect themselves so we can call back
Sub Attach(connectA As IVariantWalker)
    Set connect = connectA
End Sub

Private Sub Class_Initialize()
    ' Only executes once for life of program
    If Not EnumUpdated Then
        EnumUpdated = True
        ' There's only one v-table for the object, so modify it once
        Dim iev As IVBEnumVARIANT
        Set iev = Me
        ' Ignore item 1: QueryInterface
        ' Ignore item 2: AddRef
        ' Ignore item 3: Release
        ReplaceVtableEntry ObjPtr(iev), 4, AddressOf modColEnum.ColNext
        ReplaceVtableEntry ObjPtr(iev), 5, AddressOf modColEnum.ColSkip
        ' Ignore item 6: Reset
        ' Ignore item 7: Clone
    End If
End Sub

'' Dummy versions of implemented functions are replaced by standard module versions

Private Sub IVBEnumVARIANT_Next(ByVal cv As ULONG, v As Variant, ByVal cvFetched As Long)
    'Real implementation elsewhere
End Sub

Private Sub IVBEnumVARIANT_Skip(ByVal cv As ULONG)
    ' Dummy Skip
End Sub

' Can be implemented directly--no need to mess with the v-table
Private Sub IVBEnumVARIANT_Reset()
    connect.Reset
End Sub

Private Sub IVBEnumVARIANT_Clone(ppenum As stdole.IEnumVARIANT)

    Set ppenum = connect.Clone
    
End Sub

'' Object methods called by standard module v-table functions contain implementation

Function ClsNext(v As Variant) As Boolean
    ClsNext = connect.More(v)
End Function

Sub ClsSkip(c As Long)
    connect.Skip c
End Sub

