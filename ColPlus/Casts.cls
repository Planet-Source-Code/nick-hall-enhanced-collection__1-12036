VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GCasts"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Provides functions to cast Collection objects to other implemented interfaces"
'Global multi-use class for casting interfaces.  Note that Collection
'and IVBCollection interfaces are synonymous

Option Explicit

Public Function CCollectionPlus(ByVal ColPlus As CCollectionPlus) As CCollectionPlus
Attribute CCollectionPlus.VB_Description = "Return a collection's CollectionPlus interface"

    Set CCollectionPlus = ColPlus
    
End Function


Public Function CCollectionString(ByVal ColPlus As CCollectionString) As CCollectionString
Attribute CCollectionString.VB_Description = "Return a collection's CCollectionString interface"

    Set CCollectionString = ColPlus
    
End Function

Public Function Collection(ByVal ColPlus As IVBCollection) As IVBCollection
Attribute Collection.VB_Description = "Return a collection's Collection interface"

    Set Collection = ColPlus
    
End Function


Public Function ICollectionPlus_VB5(ByVal ColPlus As ICollectionPlus_VB5) As ICollectionPlus_VB5
Attribute ICollectionPlus_VB5.VB_Description = "Return a collection's VB5 compatible interface"

    Set ICollectionPlus_VB5 = ColPlus
    
End Function


Public Function ICollectionPlusItem(ByVal ColItem As ICollectionPlusItem) As ICollectionPlusItem
Attribute ICollectionPlusItem.VB_Description = "Return an item's ICollectionPlusItem interface"

    Set ICollectionPlusItem = ColItem
    
End Function

Public Function ICollectionPlusSettings(ByVal ColPlus As ICollectionPlusSettings) As ICollectionPlusSettings
Attribute ICollectionPlusSettings.VB_Description = "Return a collection's Settings interface"

    Set ICollectionPlusSettings = ColPlus
    
End Function


Public Function ICollectionPlusSite(ByVal ColSite As ICollectionPlusSite) As ICollectionPlusSite
Attribute ICollectionPlusSite.VB_Description = "Return an object's ICollectionPlusSite interface"

    Set ICollectionPlusSite = ColSite
    
End Function


