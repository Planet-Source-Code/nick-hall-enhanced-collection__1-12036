1.  Introduction
----------------

The CollectionPlus component consists of 2 main classes: -

CCollectionPlus - Collection of Variants;
CCollectionString - Collection of Strings

Both support the standard Collection methods/properties (Count, Add, Remove, Item) as well 
as the following: -


Method		Description

Clear 		Empties the collection;

Clone 	 	Create a duplicate of the current collection;

EnumDirection	When the collection is iterated via For-Each, the setting of this property
		determines whether the iteration will be forwards or backwards;

Exists		Determines whether or not an item is present in the collection;

Find		Returns an array of variants containing all items in the collection that 
		matched the specified string.  Supports case-sensitive/insensitive matching;

FindItems	(CCollectionPlus only) Allows custom searches of the collection.  By 
		implementing a callback interface in your classes, you can control what 
		items are returned from this function

Index		Returns the numeric position of an item within the collection

Item		As well as the normal Get functionality, you can also assign to an existing
		item in the collection

Items		Returns an array of variants containing all items in the collection

Key		Returns the key of an item within the collection

Load		Restore a collection from a file

MatchCase	Determines whether the collection matches keys case-sensitive (True) or 
		case-insensitive (default, false)

Move		Move an item within the collection

Save		Save the collection to a file

Sort		Sorts the collection (ascending/descending, case sensitive/insensitive)

Tag		Store user-defined string data against an item


2.  Interfaces and other features
---------------------------------

Interfaces implemented by both classes: -

IVBCollection		A VB-compatible version of the interface used by VB's own collection
			object.

ICollectionPlus_VB5	This interface allows VB5 users to make use of the Find, FindItems, 
			Items and Keys methods (which use VB6's ability to return arrays 
			directly from methods)

ICollectionPlusSettings	Allows you to tweak the internal settings e.g. how much space is
			reserved for the hash table


Two other interfaces are included: -

ICollectionPlusItem	A callback interface for customising the way the collection 
			sorts/searches itself.  This should be implemented by classes
			of objects that will be stored in the collection

ICollectionPlusSite	A callback interface that can be implemented by the container of
			the collection in order to allow persistance of ICollectionPlusItem
			objects

Both classes are persistable, so it is possible to write the contents of a collection to
a property bag with one statement e.g.: -

Dim pbg as New PropertyBag

'Assume mCol declared and filled earlier
pbg.WriteProperty "Collection", mCol

The CCollectionString class also implements the CCollectionPlus interface, allowing it to be
used interchangeably with the CCollectionPlus class

The component contains one other class - GCasts.  This allows you to use interfaces other
than the default one without having to declare an extra variable.  It is defined as a 
Global-Multiuse class, so you can call its methods directly e.g.

Dim v as variant

v = ICollectionPlus_VB5(mCol).Items


3.  The ColLib type library
---------------------------
The type library contains definitions for interfaces that are not directly exposed
to the client application.  It also contains some definitions for Win32 functions and
constants that are used by the CollectionPlus component.  The compiled DLL does not
need this library to be distributed.


4.  The DLL
--------------------------
The type library of the DLL has been enhanced using Matt Curland's Post Build Type Library
Modifier Add-in, included with his book Advanced Visual Basic 6 - Power techniques for
Everyday programs.  I have included the .leh file which contains details of all the edits made - these include incorporating parts of the ColLib type library and marking most of the
ByRef arguments as [in] so as to prevent data from being marshalled in both directions
when used cross-process.


5.  VB5 users
-------------
The DLL requires the MSVBVM60.dll (obtainable from http://support.microsoft.com/download/support/mslfiles/Vbrun60.exe).  
All functions work except for Items and Keys (which return arrays directly).  Use the 
ICollectionPlus_VB5 interface to access these functions.  

I have not tried to compile the code in VB5; in order to do so, you would need to change
the following things: -

a) Change the signature of the Items and Keys methods to return variants instead of arrays;
b) Code a Replace function (used in modColPlus/ErrRaise)


6.  Acknowledgements
--------------------
I would like to thank the following people whose code/ideas are incorporated within the
CollectionPlus component: -

Bruce Mckinney - Enumerating variants, sorting via interfaces & global casting modules were 
all taken from Bruce Mckinney's Hardcore Visual Basic (sadly now out of print.;

Francesco Balena - The technique for creating linked lists/hash tables from arrays of UDT's
was taken from the May 2000 issue of VBPJ ("Speed up your Apps with Data structures").
