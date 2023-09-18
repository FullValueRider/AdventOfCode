Attribute VB_Name = "Unsafe"
Option Explicit
'https://docs.microsoft.com/en-us/windows/win32/api/wtypes/ne-wtypes-varenum
' enumeration of the types tthat can be managed by a variantt



Public Enum SafeArrayfFeature
    
    FADF_AUTO = &H1             ' An array that is allocated on the stack.
    FADF_STATIC = &H2           ' An array that is statically allocated.
    FADF_EMBEDDED = &H4         ' An array that is embedded in a structure.
    FADF_FIXEDSIZE = &H10       ' An array that may not be resized or reallocated.
    FADF_RECORD = &H20          ' An array that contains records. When set, there will be a pointer to the IRecordInfo interface at negative offset 4 in the array descriptor.
    FADF_HAVEIID = &H40         ' An array that has an IID identifying interface. When set, there will be a GUID at negative offset 16 in the safe array descriptor. Flag is set only when FADF_DISPATCH or FADF_UNKNOWN is also set.
    FADF_HAVEVARTYPE = &H80     ' An array that has a variant Type. The variant type can be retrieved with SafeArrayGetVartype.
    FADF_BSTR = &H100           ' An array of BSTRs.
    FADF_UNKNOWN = &H200        ' An array of IUnknown*.
    FADF_DISPATCH = &H400       ' An array of IDispatch*.
    FADF_VARIANT = &H800        ' An array of VARIANTs.
    FADF_RESERVED = &HF008
    
End Enum

#Region "FromModCHashD"
Public Declare PtrSafe Sub BindArray Lib "kernel32" Alias "RtlMoveMemory" (PArr() As Any, pSrc As LongPtr, Optional ByVal CB As Long = 4)
Public Declare PtrSafe Function VariantCopy Lib "oleaut32" (Dst As Any, Src As Any) As Long
Public Declare PtrSafe Function VariantCopyInd Lib "oleaut32" (Dst As Any, Src As Any) As Long
Private Declare PtrSafe Function CharLowerBuffW Lib "user32" (lpsz As Any, ByVal cchLength As Long) As Long

Public LWC(-32768 To 32767) As Integer

Public Sub InitLWC()
    Dim i As Long
    For i = -32768 To 32767
        LWC(i) = i
    Next 'init the Lookup-Array to the full WChar-range
    CharLowerBuffW LWC(-32768), 65536 '<-- and convert its whole content to LowerCase-WChars
End Sub

#End Region


'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( ByVal Destination As Long, ByVal Source As Long, ByVal Length As Integer)
Public Declare PtrSafe Sub CopyMemoryToAny Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As LongPtr, ByVal Length As Long)

Public Declare PtrSafe Sub CopyAnyToMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByRef Source As Any, ByVal Length As Long)

' CopyAnyToMemory pSafeArray + 12, StrPtr (S), 4
Public Declare PtrSafe Sub GetArrayElement Lib "oleaut32" Alias "SafeArrayGetElement" (ByVal ipSAPtr As LongPtr, ByVal ipIndexesPrt As LongPtr, ByVal opDataPtr As LongPtr)
Public Declare PtrSafe Sub PutArrayElement Lib "oleaut32" Alias "SafeArrayPutElement" (ByVal ipSAUdtPtr As LongPtr, ByVal ipIndexesPrt As LongPtr, ByVal ipDataPtr As LongPtr)
Public Declare PtrSafe Function MakeArray Lib "oleaut32" Alias "SafeArrayCreate" (ByVal ipVarType As Long, ByVal ipDims As Long, ByVal ipSafeArrayBoundsPtr As LongPtr) As LongPtr

Public Declare PtrSafe Function VarPtrArray Lib "VBE7.dll" Alias "VarPtr" (Var() As Any) As LongPtr

'https://codekabinett.com/rdumps.php?Lang=2&targetDoc=windows-api-declaration-vba-64-bit
' A common pitfall - The size of user-defi ned types
' There is a common pitfall that, to my surprise, is hardly ever mentioned.

' Many API-Functions that need a user defined Type (=UDT) passed as one of their arguments expect to be
' informed about the size of that type. This usually happens either by the size being stored in a member
' inside the structure or passed as a separate argument to the function.

' Frequently developers use the Len-Function to determine the size of the type. That is incorrect,
' but it usually works on the 32-bit platform - by pure chance. Unfortunately,
' it frequently fails on the 64-bit platform.

' To understand the issue, you need to know two things about Window’s inner workings.

' The members of user-defined types are aligned sequentially in memory. One member after the other.
' Windows manages its memory in small chunks. On a 32-bit system, these chunks are always 4 bytes big.
' On a 64-bit system, these chunks have a size of 8 bytes.
' If several members of a user-defined type fit into such a chunk completely, they will be stored in
' just one of those. If a part of such a chunk is already filled and the next member in the structure
' will not fit in the remaining space, it will be put in the next chunk and the remaining space in
' the previous chunk will stay unused. This process is called padding.

' Regarding the size of user-defined types, the Windows API expects to be told the complete size the
' type occupies in memory. Including those padded areas that are empty but need to be considered to
'manage the total memory area and to determine the exact positions of each of the members of the type.

' The Len-Function adds up the size of all the members in a type, but it does not count the empty
'memory areas, which might have been created by the padding. So, the size computed by the Len-Function
'is not correct! - You need to use the LenB-Function to determine the total size of the type in memory.

' Here is a small sample to illustrate the issue:


' Public Type SmallType
'     a As Integer
'     b As Long
'     x As LongPtr
' End Type

' Public Sub testTypeSize()

'     Dim s As SmallType
    
'     Debug.Print "Len: " & Len(s)
'     Debug.Print "LenB: " & LenB(s)

' End Sub
' On 32-bit the Integer is two bytes in size but it will occupy 4 bytes in memory because the Long is put
' in the next chunk of memory. The remaining two bytes in the first chunk of memory are not used. The size
' of the members adds up to 10 bytes, but the whole Type is 12 bytes in memory.

' Memory layout of user defined type in 32-bit VBA/Windows
' On 64-bit the Integer and the Long are 6 bytes total and will fit into the first chunk together.
' The LongPtr (now 8 bytes in size) will be put into the next chunk of memory and once again the remaining
' two bytes in the first chunk of memory are not used. The size of the members adds up to 14 bytes, but the
' whole type is 16 bytes in memory.

' Memory layout of user defined type in 64-bit VBA/Windows
' So, if the underlying mechanism exists on both platforms, why is this not a problem with API calls on 32-bit? –
' It is, but simply by pure chance it is unlikely you are affected by the problem. There are relatively
' few user defined types using a datatype smaller than a DWORD (Long) As a member and I don’t know any that
' also use the size/length of the UDT structure.

Public Type ResultBufferUDT
    
    I1 As Long
   
    I2 As Long
   
    I3 As Long
   
    I4 As Long
    L1 As Long
    L2 As Long
    L3 As Long
    L4 As Long
End Type
 
'  Memory dump for the SafeArray structure including preceding hidden VT data using MS Access VBA 7
' Pos  Address Dec    Address Hex     Hex
' 0    (1185248224)   (46A573E0) >>    0h
' 1    (1185248225)   (46A573E1) >>    0h
' 2    (1185248226)   (46A573E2) >>    0h
' 3    (1185248227)   (46A573E3) >>    0h
' 4    (1185248228)   (46A573E4) >>    0h
' 5    (1185248229)   (46A573E5) >>    0h
' 6    (1185248230)   (46A573E6) >>    0h
' 7    (1185248231)   (46A573E7) >>    0h
' 8    (1185248232)   (46A573E8) >>    0h
' 9    (1185248233)   (46A573E9) >>    0h
' 10   (1185248234)   (46A573EA) >>    0h
' 11   (1185248235)   (46A573EB) >>    0h
' -------------------------------------------------------------------------------------
' Hidden DWord for VT when FADF_HAVEVARTYPE = 0x0080
' VT = 2 i.e. Integer
' 12   (1185248236)   (46A573EC) >>    2h
' 13   (1185248237)   (46A573ED) >>    0h
' 14   (1185248238)   (46A573EE) >>    0h
' 15   (1185248239)   (46A573EF) >>    0h
' -------------------------------------------------------------------------------------
' cDims = 1 i.e. One dimensional array
' 16   (1185248240)   (46A573F0) >>    1h
' 17   (1185248241)   (46A573F1) >>    0h
' -------------------------------------------------------------------------------------
' fFeatures = FADF_HAVEVARTYPE
' 18   (1185248242)   (46A573F2) >>    80h
' 19   (1185248243)   (46A573F3) >>    0h
' -------------------------------------------------------------------------------------
' cbElements = 2 i.e. element size is 2 bytes
' 20   (1185248244)   (46A573F4) >>    2h
' 21   (1185248245)   (46A573F5) >>    0h
' 22   (1185248246)   (46A573F6) >>    0h
' 23   (1185248247)   (46A573F7) >>    0h
' -------------------------------------------------------------------------------------
' cLocks
' 24   (1185248248)   (46A573F8) >>    0h
' 25   (1185248249)   (46A573F9) >>    0h
' 26   (1185248250)   (46A573FA) >>    0h
' 27   (1185248251)   (46A573FB) >>    0h
' -------------------------------------------------------------------------------------
' Padding
' 28   (1185248252)   (46A573FC) >>    0h
' 29   (1185248253)   (46A573FD) >>    0h
' 30   (1185248254)   (46A573FE) >>    0h
' 31   (1185248255)   (46A573FF) >>    0h
' -------------------------------------------------------------------------------------
' pvData
' 32   (1185248256)   (46A57400) >>    0h
' 33   (1185248257)   (46A57401) >>    97h
' 34   (1185248258)   (46A57402) >>    DCh
' 35   (1185248259)   (46A57403) >>    1Bh
' 36   (1185248260)   (46A57404) >>    0h
' 37   (1185248261)   (46A57405) >>    0h
' 38   (1185248262)   (46A57406) >>    0h
' 39   (1185248263)   (46A57407) >>    0h
' -------------------------------------------------------------------------------------
' rgsabound(0).cElements = 10
' 40   (1185248264)   (46A57408) >>    Ah
' 41   (1185248265)   (46A57409) >>    0h
' 42   (1185248266)   (46A5740A) >>    0h
' 43   (1185248267)   (46A5740B) >>    0h
' -------------------------------------------------------------------------------------
' rgsabound(0).lLbound
' 44   (1185248268)   (46A5740C) >>    0h
' 45   (1185248269)   (46A5740D) >>    0h
' 46   (1185248270)   (46A5740E) >>    0h
' 47   (1185248271)   (46A5740F) >>    0h

' Note for VBA 7 64 Bit shows the offset for SAFEARRAYBOUNDS[0...] at +24 Dec, +18 hex due to the 8-byte memory address for pvData and padding.
' I hope that helps further explain what exactly any of the SafeArray functions are performing and wherein the SafeArray structure items are located. If anyone is manually manipulating SafeArray structures be careful of any padding and getting your offsets correct.
' I'll also try and get around to your query regarding VT_I4(3) or VT_R4(4).
' There's definitely something extra going on in VBA then what the SafeArray.c functions are performing as from testing using SafeArrayDescriptorEx to create an initialized empty Integer Array the cbElements weren't set, the preceding VT was set. Strangely the VBA integer Array still worked without the cbElements set to 2 bytes for an Integer Array. When creating the SafeArrayDescriptor I now set the cbElements, manually.
' When created using the above VBA example the cbElements were set appropriately.
' In the SafeArray.c functions they aren't returning a SizeOf(VT_I4) or SizeOf(VT_R4) i.e. not supported in C, so I assume in VBA must expanding on the SafeArray.c functions and catering for its data types not covered in C.
' Someone more with more knowledge of C might be able to clarify or explain better.


' Short Version
' varType = SafeArrayGetVarType(mySafeArray);
' Long Version
' The SAFEARRAY has a features member that helps describing what's in the array

' 2.2.30.10 SAFEARRAY (archive)
' fFeatures: MUST be set to a combination of the bit flags specified in section 2.2.9.
' And then you consult:

' 2.2.9 ADVFEATUREFLAGS Advanced Feature Flags (archive)

' The following values are used in the field fFeatures of a SAFEARRAY (section 2.2.30.10) data Type.

' typedef  Enum tagADVFEATUREFLAGS
'  {
'    FADF_AUTO = 0x0001,
'    FADF_STATIC = 0x0002,
'    FADF_EMBEDDED = 0x0004,
'    FADF_FIXEDSIZE = 0x0010,
'    FADF_RECORD = 0x0020,
'    FADF_HAVEIID = 0x0040,
'    FADF_HAVEVARTYPE = 0x0080,
'    FADF_BSTR = 0x0100,
'    FADF_UNKNOWN = 0x0200,
'    FADF_DISPATCH = 0x0400,
'    FADF_VARIANT = 0x0800
'  } ADVFEATUREFLAGS;
' FADF_RECORD: The SAFEARRAY MUST contain elements of a UDT (see section 2.2.28.1)
' FADF_HAVEIID: The SAFEARRAY MUST contain MInterfacePointers elements.
' FADF_HAVEVARTYPE: If this bit flag is set, the high word of the cLocks field of the SAFEARRAY MUST contain a VARIANT type constant that describes the type of the array's elements (see sections 2.2.7 and 2.2.30.10).
' FADF_BSTR: The SAFEARRAY MUST contain an array of BSTR elements (see section 2.2.23).
' FADF_UNKNOWN: The SAFEARRAY MUST contain an array of pointers to IUnknown.
' FADF_DISPATCH: The SAFEARRAY MUST contain an array of pointers to IDispatch (see section 3.1.4).
' FADF_VARIANT: The SAFEARRAY MUST contain an array of VARIANT instances.
' So depending on the FADF, you can come up with a corresponding Variant Type:

' Feature Flag  Corresponding Variant Type
' FADF_UNKNOWN  VT_UNKNOWN
' FADF_DISPATCH VT_DISPATCH
' FADF_VARIANT  VT_VARIANT
' FADF_BSTR     VT_BSTR
' FADF_HAVEVARTYPE      SafeArrayGetVarType(mySafeArray)
' It turns out all that above work (AreSameing FADF_BSTR to VT_BSTR etc) is all wrapped up by the helper function SafeArrayGetVarType (archive):

' If FADF_HAVEVARTYPE is set, SafeArrayGetVartype returns the VARTYPE stored in the array descriptor.
' If FADF_RECORD is set, it returns VT_RECORD;
' if FADF_DISPATCH is set, it returns VT_DISPATCH;
' and if FADF_UNKNOWN is set, it returns VT_UNKNOWN.
' SafeArrayGetVartype can fail to return VT_UNKNOWN for SAFEARRAY types that are based on IUnknown. Callers should additionally check whether the SAFEARRAY type's fFeatures field has the FADF_UNKNOWN flag set.


' FADF_HAVEVARTYPE: If this bit flag is set, the high word of the cLocks field of the SAFEARRAY MUST contain a VARIANT Type constant
' that describes the Type of the array's elements (see sections 2.2.7 and 2.2.30.10). The MS documentation does claim this bizzare cLocks
' requirement. However, VBA never does this with the SAFEARRAYS it creates.
' Instead it utilzes the 4 bytes at the offset of -4 (the same byes reserved for FADF_RECORD) to store the VARIANT Type constant...
' and it does this even when fFeatures bit-field does not include FADF_HAVEVARTYPE. It's frustrating, but VBA does it diff. –


'/* Memory Layout of a SafeArray:
'*
'* -0x10: start of memory.
'* -0x10: GUID for VT_DISPATCH and VT_UNKNOWN safearrays (if FADF_HAVEIID)
'* -0x04: DWORD varianttype; (for all others, except VT_RECORD) (if FADF_HAVEVARTYPE)
'*  -0x4: IRecordInfo* iface;  (if FADF_RECORD, for VT_RECORD (can be NULL))
'*  0x00: SAFEARRAY,  i.e. starting at cDims
'*  0x10: SAFEARRAYBOUNDS[0...]
'*/

'VBA representation of the C Structs for the SafeArray object

Public Type SAFEARRAY1D
  cDims                     As Integer
  fFeatures                 As Integer
  cbElements                As Long
  cLocks                    As Long
  pvData                    As LongPtr
  cElements1D               As Long
  lLbound1D                 As Long
End Type

Public Type SafeBound                           ' offsets
    
    cElements           As Long                 ' 0
    lLbound             As Long                 ' 4
    
End Type

Public Enum SafeArrayOffset
    cDims = 0
    fFeature = 2
    cbElements = 4
    cLocks = 8
    pvdata = 12
End Enum

Public Type SafeArrayUDT            ' Offsets
    
    '@Ignore IntegerDataType
    cDims As Integer              ' 0         ' Number of Ranks in the array
    '@Ignore IntegerDataType
    fFeature As Integer              ' 2
    cbElements As Long                 ' 4         ' The size of each element in the array (in bytes)
    
    cLocks As Long                 ' 8
    pvData As Long                 ' 12        ' pointer to the start of the data
    
   SafeBound0 As SafeBound            ' 16
   SafeBound1 As SafeBound            ' 16
   SafeBound2 As SafeBound            ' 16
   SafeBound3 As SafeBound            ' 16
   SafeBound4 As SafeBound            ' 16
   SafeBound5 As SafeBound            ' 16
   SafeBound6 As SafeBound            ' 16
   SafeBound7 As SafeBound            ' 16
   SafeBound8 As SafeBound            ' 16
   SafeBound9 As SafeBound            ' 16
End Type


'
'/*
'The safe array udt specifies the number of bytes to copy for an element
'but does not describe the type of the element.  This isn't a problem when
'copying a variant element as the type is part of the variant data.  For
'other types the simplest method is to use vartype from the variant parameter
'containing the array.
'The ArrayInfo UDT allows for the type information to be provided along with the SafeArrayUdt
'*/

Public Type ArrayInfoUDT
    
    VarType As Long
    '@Ignore KeywordsUsedAsMember
    Type As Long
    SaudtPtr As LongPtr
        
    SAUdt As SafeArrayUDT
    
End Type

' ' The variant struct
' ' https://stackoverflow.com/questions/6901991/how-to-return-the-number-of-dimensions-of-a-variant-variable-passed-to-it-in-v
' Private Type ARRAY_VARIANT
'     vt As Integer
'     wReserved1 As Integer
'     wReserved2 As Integer
'     wReserved3 As Integer
'     lpSAFEARRAY As LongPtr               ' lp implied long ptr.  Type changed from long to longptr
'     data(4) As Byte
' End Type

Private Type TtagVARIANT
    
    '@Ignore IntegerDataType
    VariantType As Integer      ' described the type of value stored in the variant
    '@Ignore IntegerDataType
    Reserved1 As Integer      ' reserved
    '@Ignore IntegerDataType
    Reserved2 As Integer      ' reserved
    '@Ignore IntegerDataType
    Reserved3 As Integer      ' reserved
    ValueOrPtr As LongPtr      ' Value or pointer to the actual value
    
End Type

' Public Function SafeArrayDims(ByRef ipArrayInfo As LongPtr) As Long
'     Dim myInt As Integer
'     myInt = 0
'     CopyMemory ByVal VarPtr(myInt), ByVal ipArrayInfo + 4&, 2&
'     SafeArrayDims = CLng(myInt)
' End Function

'/*
'I 've not yet been able to make the Get and Put methods of the
'SafeArray work
'' For read/writing a Safe Array element
''--------------------------------------
'' Parameters
'' [ in] psa    pointer to safearray struct
'
'' An array descriptor created by SafeArrayCreate.
'
'' [ in] rgIndices  array of longs one for each dimension
'
'' A vector of indexes for each dimension of the array. The right-most (least significant) dimension is rgIndices[0]. The left-most dimension is stored at rgIndices[psa->cDims – 1].
'
'' [ in] pv  pointer to a variant
'
'' The data to assign to the array. The variant types VT_DISPATCH, VT_UNKNOWN, and VT_BSTR are pointers, and do not require another level of indirection.
'' from ReactOS
''HRESULT WINAPI         SafeArrayPutElement (SAFEARRAY *psa, LONG *rgIndices, void *pvData)
''HRESULT WINAPI         SafeArrayGetElement (SAFEARRAY *psa, LONG *rgIndices, void *pvData)
'
'' From
''WINOLEAUTAPI SafeArrayGetElement(SAFEARRAY *psa,LONG *rgIndices,void *pv);
''WINOLEAUTAPI SafeArrayPutElement(SAFEARRAY *psa,LONG *rgIndices,void *pv);
'
''from C:\Users\slayc\source\repos\Cpp\mingw64\include
''WINOLEAUTAPI SafeArrayGetElement(_In_ SAFEARRAY * psa, _In_reads_(_Inexpressible_(psa->cDims)) LONG * rgIndices, _Out_ void * pv);
''WINOLEAUTAPI SafeArrayPutElement(_In_ SAFEARRAY * psa, _In_reads_(_Inexpressible_(psa->cDims)) LONG * rgIndices, _In_ void * pv);
'
'' From Libre Office1
'' HRESULT WINAPI SafeArrayPutElement(SAFEARRAY*,LONG*,void*);
'' HRESULT WINAPI SafeArrayGetElement(SAFEARRAY*,LONG*,void*);
'*/




' https://stackoverflow.com/questions/24613101/vba-check-if-array-is-one-dimensional/26555865#26555865
' 32 bit version

'/* Memory Layout of a SafeArray:
'*
'* -0x10: start of memory.
'* -0x10: GUID for VT_DISPATCH and VT_UNKNOWN safearrays (if FADF_HAVEIID)
'* -0x04: DWORD varianttype; (for all others, except VT_RECORD) (if FADF_HAVEVARTYPE)
'*  -0x4: IRecordInfo* iface;  (if FADF_RECORD, for VT_RECORD (can be NULL))
'*  0x00: SAFEARRAY,  i.e. starting at cDims
'*  0x10: SAFEARRAYBOUNDS[0...]
'*/

' Public Type SafeArrayUDT            ' Offsets
    
'     cDims As Integer                ' 0         ' Number of Ranks in the array
'     fFeature As Integer             ' 2
'     cbElements As Long              ' 4         ' The size of each element in the array (in bytes)
'     cLocks As Long                  ' 8
'     pvData As Long                  ' 12        ' pointer to the start of the data
'     rgsabound As SafeBound          ' 16        ' pointer to an array of safebound equal in size to the number to CDims
    
' End Type



Public Function GetArrayInfo(ByRef ipInput As Variant) As ArrayInfoUDT
   
    Const VT_DATA_OR_PTR As Long = 8
    Dim myArrayInfo As ArrayInfoUDT
    
    myArrayInfo.Type = 0
   ' If BailOut.When(TypeInfo.IsNotArray(ipInput), IsNotArray) Then Return myArrayInfo
    
    Dim lPointer As LongPtr
    lPointer = 0
    
    ' First 2 bytes are the subtype stored in the array
    CopyMemoryToAny ByVal VarPtr(myArrayInfo.Type), ByVal VarPtr(ipInput), 2
    
    ' Type without array and reference flags
    myArrayInfo.VarType = myArrayInfo.Type And &HFF
    
    'Get the pointer to the array
    CopyMemoryToAny ByVal VarPtr(lPointer), ByVal VarPtr(ipInput) + VT_DATA_OR_PTR, 4

    'Test for Byref i.e value field is a pointter
    If (myArrayInfo.Type And VT_BYREF) <> 0 Then
        
        'Get the real address
        CopyMemoryToAny ByVal VarPtr(lPointer), ByVal lPointer, 4
        myArrayInfo.SaudtPtr = lPointer
        
    End If
    ' exit if we didn't find an allocated array
    If lPointer = 0 Then Return myArrayInfo
        
    ' now populate the safearrayUDT field of ArrayInfo
    CopyMemoryToAny ByVal VarPtr(myArrayInfo.SAUdt), ByVal lPointer, 16            'Write the safearray data.
    CopyMemoryToAny ByVal VarPtr(myArrayInfo.SAUdt.SafeBound0), ByVal lPointer + 16, myArrayInfo.SAUdt.cDims * 8
    Return myArrayInfo
    
End Function

' Public Function GetDims(VarSafeArray As Variant) As Integer
'     Dim variantType As Integer
'     Dim pointer As Long
'     Dim arrayDims As Integer

'     'The first 2 bytes of the VARIANT structure contain the type:
'     CopyMemory VarPtr(variantType), VarPtr(VarSafeArray), 2&

'     If Not (variantType And &H2000) > 0 Then
'     'It's not an array. Raise type misAreSame.
'         Err.Raise (13)
'     End If

'     'If the Variant contains an array or ByRef array, a pointer for the _
'         SAFEARRAY or array ByRef variant is located at VarPtr(VarSafeArray) + 8:
'     CopyMemory VarPtr(pointer), VarPtr(VarSafeArray) + 8, 4&

'     'If the array is ByRef, there is an additional layer of indirection through_
'     'another Variant (this is what allows ByRef calls to modify the calling scope).
'     'Thus it must be dereferenced to get the SAFEARRAY structure:
'     If (variantType And &H4000) > 0 Then 'ByRef (&H4000)
'         'dereference the pointer to pointer to get actual pointer to the SAFEARRAY
'         CopyMemory VarPtr(pointer), pointer, 4&
'     End If
'     'The pointer will be 0 if the array hasn't been initialized
'     If Not pointer = 0 Then
'         'If it HAS been initialized, we can pull the number of dimensions directly _
'             from the pointer, since it's the first member in the SAFEARRAY struct:
'         CopyMemory VarPtr(arrayDims), pointer, 2&
'         GetDims = arrayDims
'     Else
'         GetDims = 0 'Array not initialized
'     End If
' End Function




Public Function SafeArrayGetByIndex(ByVal ipIndex As Long, ByRef ipArrayInfo As ArrayInfoUDT) As Variant
     
    Dim myResult As Variant
    myResult = 0
    
    If ipArrayInfo.SAUdt.cbElements = 16 Then
        
        CopyMemoryToAny myResult, ByVal ipArrayInfo.SAUdt.pvData + ipIndex * ipArrayInfo.SAUdt.cbElements, ipArrayInfo.SAUdt.cbElements
        
        
    Else
        
        ' Copy array item into data area of Variant
        ' do we need to check for pointer in value type.
        CopyMemoryToAny myResult + 8, ByVal ipArrayInfo.SAUdt.pvData + ipIndex * ipArrayInfo.SAUdt.cbElements, ipArrayInfo.SAUdt.cbElements
        
        ' set the vartype of the myResult Variable
        CopyMemoryToAny myResult, ByVal VarPtr(ipArrayInfo.VarType), 2&
        
        
    End If
    
    Sys.Assign SafeArrayGetByIndex, myResult
    
End Function


Public Sub SafeArrayPutByIndex(ByVal ipIndex As Long, ByVal ipItem As Variant, ByRef ipArrayInfo As ArrayInfoUDT)
    
    If ipArrayInfo.SAUdt.cbElements = 16 Then
        
        CopyAnyToMemory ipArrayInfo.SAUdt.pvData + ipIndex * ipArrayInfo.SAUdt.cbElements, ipItem, ipArrayInfo.SAUdt.cbElements
        ' Do I need to set the variant type?  Testing will indicate
        ' do I need to dereference IDespath types?  e.g I'm getting a failure with oNumbers and VBType is IDespatch
    Else
        
        CopyAnyToMemory ByVal ipArrayInfo.SAUdt.pvData + ipIndex * ipArrayInfo.SAUdt.cbElements, ipItem + 8, ipArrayInfo.SAUdt.cbElements
        
        
    End If
    
End Sub

Public Sub SafeArrayPutElement(ByVal ipArrayPtr As LongPtr, ByRef ipData As Variant, ParamArray ipIndeces() As Variant)
    
    Dim myUbound As Long
    myUbound = GetArrayInfo(CVar(ipIndeces)).SAUdt.cDims
    Dim myIndeces() As Long
    ReDim myIndeces(0 To myUbound)
    
    Dim myIndex As Long
    For myIndex = LBound(myIndeces) To UBound(myIndeces)
        myIndeces (myIndex) = CLng(ipIndeces(myIndex))
    Next
    
    Dim myIndecesPtr As LongPtr
    myIndecesPtr = GetArrayInfo(CVar(myIndeces)).SAUdt.pvData
    
    PutArrayElement ipArrayPtr, ByVal myIndecesPtr, ByVal VarPtr(ipData)
    
End Sub



Public Function GetSafeArrayPtr(ByRef ipInput As Variant) As LongPtr
   
    Const VT_DATA_OR_PTR As Long = 8
    If BailOut.When(TypeInfo.IsNotAnArray(ipInput), IsNotArray) Then Exit Function
        
    Dim myArrayInfo As ArrayInfoUDT
    myArrayInfo.Type = 0
    
    Dim lPointer As LongPtr
    lPointer = 0
    
    CopyMemoryToAny ByVal VarPtr(myArrayInfo.Type), ByVal VarPtr(ipInput), 2                     'First 2 bytes are the subtype.
    myArrayInfo.VarType = myArrayInfo.Type And &HFF                                         ' Type without array and reference flags
    
    CopyMemoryToAny lPointer, ByVal VarPtr(ipInput) + VT_DATA_OR_PTR, 4            'Get the pointer.

    
    If (myArrayInfo.Type And VT_BYREF) <> 0 Then
                                                   'Test for Byref i.e value field is a pointter
        CopyMemoryToAny lPointer, ByVal lPointer, 4                       'Get the real address.
        
        
    End If
    
    GetSafeArrayPtr = lPointer
    
End Function


Public Sub testGetElement()
    
    Dim myArray As Variant
    myArray = Array("Hello", "There", "World", "Happy")
    
    Dim myData As Variant
    myData = "It works"
    
    SafeArrayPutElement myArray, myData, 2
   
End Sub

'@Ignore NonReturningFunction, ParameterNotUsed, ParameterCanBeByVal
Public Function SafeArrayGetElement(ByRef ipArray As Variant, ByRef ipData As Variant, ParamArray ipIndeces() As Variant) As Variant

    Dim myResult As ResultBufferUDT
    With myResult
        .I1 = 0
        .I2 = 0
        .I3 = 0
        .I4 = 0
        .L1 = 0
        .L2 = 0
        .L3 = 0
        .L4 = 0
    End With

    Dim myIndeces As Variant
    myIndeces = CVar(ipIndeces)

    Dim myIndecesPtr As LongPtr
    myIndecesPtr = GetArrayInfo(myIndeces).SAUdt.pvData

    Dim myVar As Variant
    myVar = 0

    Dim myVarPtr As LongPtr
    myVarPtr = VarPtr(myVar)


    Dim mySAPtr As LongPtr
    mySAPtr = GetSafeArrayPtr(ipArray)


    GetArrayElement ByVal mySAPtr, ByVal myIndecesPtr, ByVal myVarPtr
    'Debug.Print myVar
End Function
'
'Public Function SafeArrayPutElement(ByRef ipArray As Variant, ByRef ipData As Variant, ParamArray ipIndeces() As Variant) As Variant
'
'    Dim myResult As ResultBufferUDT
'    With myResult
'        .I1 = 0
'        .I2 = 0
'        .I3 = 0
'        .I4 = 0
'        .L1 = 0
'        .L2 = 0
'        .L3 = 0
'        .L4 = 0
'    End With
'
'    Dim myIndeces As Variant
'    myIndeces = CVar(ipIndeces)
'
'    Dim myIndecesPtr As LongPtr
'    myIndecesPtr = GetArrayInfo(myIndeces).SAUdt.pvData
'
'
'
'    Dim myVarPtr As LongPtr
'    myVarPtr = VarPtr(ipData)
'
'
'    Dim mySAPtr As LongPtr
'    mySAPtr = GetSafeArrayPtr(ipArray)
'
'
'    PutArrayElement ByVal mySAPtr, ByVal myIndecesPtr, ByVal myVarPtr
'    'Debug.Print myVar
'End Function


