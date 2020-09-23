VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMemComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' VB answer to the unavailable C memcmp function & RtlCompareMemory API.
'   FYI: RtlCompareMemory is only available on NT-based machines (NT4 and above)

' Tested on Win95, 98, 2K, XP Pro
' Not usable for comparing string arrays because the memory in a string array is
' not contiguous. The string array data is actually pointers to the strings, not
' the strings themselves.  However, if you wish to compare the string arrays to
' see if they actually contain the same pointers, then this would work for you too.

' The amazing thing, quite surprising to me, is that this routine has nearly
' identical speed to RtlCompareMemory (+ 1>3 ms) when the following conditions met:
'   1. In project property (Compile Tab, Optimizations button),
'           "Remove Array Bounds Check" is checked/selected
'   2. Compile the project for maximum speed
'   FYI> Do not "Remove Array Bounds Check" if you are relying on Error trapping to
'        check for uninitialized arrays or if you are using it to test for
'        array out of bounds errors. Without that optimization, the routines are
'        still very fast, but are faster with the optimization.

' This class requires the following type library:
'Reference=*\G{C878CB53-7E75-4115-BD13-EECBC9430749}#1.0#0#MemAPIs.tlb#Memory APIs

' Or uncomment the following declares
'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
'Private Type SAFEARRAY1D       ' Used to overlay array on a memory address
'    cDims      As Integer      ' Count of dimensions in this array
'    fFeatures  As Integer      ' Bitfield flags indicating attributes of the array
'    cbElements As Long         ' Byte size of each element of the array
'    cLocks     As Long         ' Number of times the array has been locked without corresponding unlock. The cLocks field is a reference count that indicates how many times the array has been locked. When there is no lock, you're not supposed to access the array data, which is located in pvData.
'    pvData     As Long         ' Pointer to the start of the array data (use only if cLocks > 0)
'    cElements  As Long         ' Count of elements in this dimension
'    lLbound    As Long         ' The lower-bounding signed index number of this dimension
'    lUbound    As Long         ' The upper-bounding signed index number of this dimension
'End Type


'Private cCmp As cMemComp
'Set cCmp = New cMemComp
'If cCmp.vbMemCompare(VarPtr(aBytes1(0)), VarPtr(aBytes2(0)), lByteCount, lPos) Then "the same"
'Set cCmp = Nothing

Public Function vbMemCompare(ByVal MemAddr1 As Long, ByVal MemAddr2 As Long, _
                              ByVal nrBytesToCompare As Long, _
                              Optional ByRef UnEqualZeroBasedBytePos As Long) As Boolean
    
    ' a flexible memcmp / RtlCompareMemory substitute for VB6
    
    ' WARNING: JUST LIKE CopyMemory, RtlCompareMemory, & MemCmp
    '   this routine does NO safety checks, it can't. The pointers you pass must be
    '   valid & number of bytes to be compared must be contiguous & not overestimated
    
    ' Requires following Declarations:
    '       Types: SAFEARRAY1D
    '       APIs:  CopyMemory
    
    ' [in] MemAddr1 :: memory address (i.e., VarPtr) to compare against MemAddr2
    ' [in] MemAddr2 :: memory address (i.e., VarPtr) to compare against MemAddr1
    ' [in] nrBytesToCompare :: contiguous bytes to be compared
    '      The bytes used starting at MemAddr1,2 must be => nrBytesToCompare
    ' [out] UnEqualLocation (Optional) :: zero based byte position where inequality
    '       occurred (if any), unmodified if function returns True (equality)
    '       See end of this routine for tips on using this value if desired
    ' [out] Return Value :: True if comparision proves identical else False
    
    ' tips on use:
    '   The 1st element to be compared in the byte arrays can be anything you like, not necessarily zero
    ' Compare 2 byte arrays: vbMemCompare(VarPtr(aByte1(0)), VarPtr(aByte1(0)), UBound(aByte1) + 1)
    ' Compare 2 Long arrays: vbMemCompare(VarPtr(aLong1(0)), VarPtr(aLong2(0)), (UBound(aLong1) + 1) * 4&)
    ' Compare mix arrays: let's say compare 1000 Longs (4 bytes each) against 4000 bytes:
    '       vbMemCompare(VarPtr(aLong(0)), VarPtr(aByte(0)), 4000&)
    ' Compare 2 DIBs same format. vbMemCompare(dibPtr1, dibPtr2, (Dib ScanWidth * Dib Height))
    ' Compare DIB against DDB bytes returned by GetDIBits:
    '       vbMemCompare(dibPtr, VarPtr(dibBits(0)), (UBound(dibBits) + 1))
    ' Compare within same array: example of comparing 1st 5000 bytes with last 5000 bytes
    '       vbMemCompare(VarPtr(aByte1(0)), VarPtr(aByte1(5000)), 5000&)
    
    ' basic sanity checks
    If nrBytesToCompare < 1 Then
        Exit Function
    ElseIf MemAddr1 = 0 Then
        Exit Function
    ElseIf MemAddr2 = 0 Then
        Exit Function
    End If
    
    Dim Index As Long ' loop variable
    Const ScanSize As Long = 8& ' the tScan1/tScan2 array types must match this size
    ' FYI: using Date or Double for the tScan array types can result in Overflows
    
    ' array overlays
    Dim tSA1 As SAFEARRAY1D, tSA2 As SAFEARRAY1D
    Dim tScan1() As Currency, tScan2() As Currency      ' 8 byte scans
    Dim tBytes1() As Byte, tBytes2() As Byte            ' 1 byte scans
    
    If nrBytesToCompare >= ScanSize Then
    
        ' set up ScanSize byte scan over the 1st passed memory pointer
        With tSA1
            .cDims = 1
            .pvData = MemAddr1
            .cbElements = ScanSize
            .cElements = (nrBytesToCompare \ ScanSize) ' truncate for now
        End With
        ' set up ScanSize byte scan over the 2nd passed memory pointer
        tSA2 = tSA1
        tSA2.pvData = MemAddr2
        
        ' overlay now
        CopyMemory ByVal ArrayPtr(tScan1), VarPtr(tSA1), 4&
        CopyMemory ByVal ArrayPtr(tScan2), VarPtr(tSA2), 4&
        
        ' compare, ScanSize bytes at a time. Wish VB had a 16, 32 or 64 variable type
        For Index = 0 To UBound(tScan1)
            If Not tScan1(Index) = tScan2(Index) Then Exit For
            ' bug out when inequality is found
        Next
        ' remove the overlays
        CopyMemory ByVal ArrayPtr(tScan1), 0&, 4&
        CopyMemory ByVal ArrayPtr(tScan2), 0&, 4&
            
        Index = Index * ScanSize ' set Index = actual byte to be checked next
        
    End If
    
    If Not Index = nrBytesToCompare Then
        ' unequal if all bytes were checked....
        ' locate exact byte position where inequality was located
        ' This also will check any bytes not compared due to non-ScanSize alignment
        
        ' set up 1 byte scan over the 1st passed memory pointer
        With tSA1
            .cDims = 1
            .cbElements = 1
            .pvData = MemAddr1 + Index  ' move memory pointer to where Index left off
            .lLbound = Index  ' adjust LBound to where Index left off
            .cElements = (nrBytesToCompare - Index) + 1 ' nr elements remaining
        End With
        ' set up 1 byte scan over the 2nd passed memory pointer
        tSA2 = tSA1
        tSA2.pvData = MemAddr2 + Index  ' move memory ponter to where Index left off
        ' overlay now
        CopyMemory ByVal ArrayPtr(tBytes1), VarPtr(tSA1), 4&
        CopyMemory ByVal ArrayPtr(tBytes2), VarPtr(tSA2), 4&
    
        ' do the comparison and/or check last n bytes
        For Index = Index To nrBytesToCompare - 1 '(max of ScanSize loops)
            If Not tBytes1(Index) = tBytes2(Index) Then Exit For
            ' bug out when inequality is found
        Next
        ' remove overlays
        CopyMemory ByVal ArrayPtr(tBytes1), 0&, 4&
        CopyMemory ByVal ArrayPtr(tBytes2), 0&, 4&

    End If
    
    ' return result(s)
    vbMemCompare = (Index = nrBytesToCompare)
    If Not vbMemCompare Then UnEqualZeroBasedBytePos = Index ' zero based position
    
    ' If you wish to identify where in your passed memory the inequality occured in
    ' relation to arrays, pixels, or memory addresses...
    
    ' This routine has no way of knowing whether you passed it a long, byte, integer
    ' array or whether you passed it memory addresses like DIB pointers. Suggest
    ' using following algos with the returned UnEqualLocation parameter.
    
    ' Arrays: Note aStartA & aStartB are array elements passed to this
    '   function (i.e., byteArrayA(0), byteArrayB(1) where aStartA=0, aStartB=1)
    
    ' Long arrays.
    '       Loc = UnEqualLocation\4 ' Calculate Loc
    '       LongA(Loc + aStartA) <> LongB(Loc + aStartB))
    
    ' Integer Arrays.
    '       Loc = UnEqualLocation\2 ' Calculate Loc
    '       IntegerA(Loc + aStartA) <> IntegerB(Loc + aStartB))
    
    ' Byte Arrays
    '       ByteA(UnEqualLocation + aStartA) <> ByteB(UnEqualLocation + aStartB))
    
    ' Date/Double/Currency arrays.
    '       Loc = UnEqualLocation\8 ' Calculate Loc
    '       DateA(Loc + aStartA) <> DateB(Loc + aStartB))
    
    ' UDTs Arrays.
    '   Yes it is possible, but not if UDT contains pointers to other memory addresses
    '   For example, VarLen string members are pointers in UDT memory & using
    '   vbMemCompare should return inequality every time since pointers won't be the same
    '       Loc = UnEqualLocation\Len(UDT) ' Calculate Loc
    '       aUDT1(Loc + aStartA) <> aUDT2(Loc + aStartB))
    '   but which member of the UDT proved inequality? Using the following you
    '   can partially determine, but maybe comparing the 2 UDT members might be best:
    '       byte within aUDT1/aUDT2 at location: UnEqualLocation-(Len(UDT)*Loc))
    
    ' DIB pointers: depending on bit depth, tweak result
    ' (remember to adjust result for bottom-up DIBs if needed)
    '   -- Pixel Colors:
    '       8 bit: UnEqualLocation is the palette index
    '       24 bit: UnEqualLocation\3 is 1st byte of the pixel
    '       32 bit: UnEqualLocation\4  is 1st byte of the pixel
    '   -- Pixel Location (DIB row & column)
    '       8 bit:  Row = UnEqualLocation\Bitmap.ScanWidth
    '               Column = UnEqualLocation-(Row * Bitmap.ScanWidth)
    '      24 bit:  Row = UnEqualLocation\Bitmap.ScanWidth
    '               Column = (UnEqualLocation-(Row * Bitmap.ScanWidth))\3
    '      32 bit:  Row = UnEqualLocation\Bitmap.ScanWidth
    '               Column = (UnEqualLocation-(Row * Bitmap.ScanWidth))\4
    
    ' Memory Pointers: simply, memPointer + UnEqualLocation
    
End Function

Private Function ArrayPtr(Arr) As Long
    Dim iDataType As Integer
    On Error GoTo UnInit
    CopyMemory iDataType, Arr, 2&                       ' get the real VarType of the argument, this is similar to VarType(), but returns also the VT_BYREF bit
    If (iDataType And vbArray) = vbArray Then           ' if a valid array was passed
        CopyMemory ArrayPtr, ByVal VarPtr(Arr) + 8&, 4& ' get the address of the SAFEARRAY descriptor stored in the second half of the Variant parameter that has received the array. Thanks to Francesco Balena.
    End If
UnInit:
End Function

Public Function vbMemCompareR(ByVal MemAddr1 As Long, ByVal MemAddr2 As Long, _
                              ByVal nrBytesToCompare As Long, _
                              Optional ByRef UnEqualZeroBasedBytePos As Long) As Boolean

    ' basic sanity checks
    If nrBytesToCompare < 1 Then
        Exit Function
    ElseIf MemAddr1 = 0 Then
        Exit Function
    ElseIf MemAddr2 = 0 Then
        Exit Function
    End If
    
    Dim Index As Long ' loop variable
    Dim Offset As Long
    Const ScanSize As Long = 8& ' the tScan1/tScan2 array types must match this size
    ' FYI: using Date or Double for the tScan array types can result in Overflows
    
    ' array overlays
    Dim tSA1 As SAFEARRAY1D, tSA2 As SAFEARRAY1D
    Dim tScan1() As Currency, tScan2() As Currency      ' 8 byte scans
    Dim tBytes1() As Byte, tBytes2() As Byte            ' 1 byte scans
    
    If nrBytesToCompare >= ScanSize Then
    
        ' set up ScanSize byte scan over the 1st passed memory pointer
        With tSA1
            .cDims = 1
            .cElements = (nrBytesToCompare \ ScanSize) ' truncate for now
             Offset = nrBytesToCompare - (.cElements * ScanSize)
            .pvData = MemAddr1 + Offset
            .cbElements = ScanSize
            .lLbound = 1
        End With
        ' set up ScanSize byte scan over the 2nd passed memory pointer
        tSA2 = tSA1
        tSA2.pvData = MemAddr2 + Offset
        
        ' overlay now
        CopyMemory ByVal ArrayPtr(tScan1), VarPtr(tSA1), 4&
        CopyMemory ByVal ArrayPtr(tScan2), VarPtr(tSA2), 4&
        
        ' compare, ScanSize bytes at a time. Wish VB had a 16, 32 or 64 variable type
        For Index = UBound(tScan1) To 1 Step -1
            If Not tScan1(Index) = tScan2(Index) Then Exit For
            ' bug out when inequality is found
        Next
        ' remove the overlays
        CopyMemory ByVal ArrayPtr(tScan1), 0&, 4&
        CopyMemory ByVal ArrayPtr(tScan2), 0&, 4&
            
        Index = Index * ScanSize ' set Index = actual byte to be checked next
        
    End If
    
    If Not (Index = 0 And Offset = 0) Then
        ' unequal if all bytes were checked....
        ' locate exact byte position where inequality was located
        ' This also will check any bytes not compared due to non-ScanSize alignment
        
        ' set up 1 byte scan over the 1st passed memory pointer
        With tSA1
            .cDims = 1
            .cbElements = 1
            .pvData = MemAddr1
            .cElements = (Index + Offset)  ' nr elements remaining
            .lLbound = 1
        End With
        ' set up 1 byte scan over the 2nd passed memory pointer
        tSA2 = tSA1
        tSA2.pvData = MemAddr2
        ' overlay now
        CopyMemory ByVal ArrayPtr(tBytes1), VarPtr(tSA1), 4&
        CopyMemory ByVal ArrayPtr(tBytes2), VarPtr(tSA2), 4&

        ' do the comparison and/or check last n bytes
        For Index = Index + Offset To 1 Step -1  '(max of ScanSize loops)
            If Not tBytes1(Index) = tBytes2(Index) Then Exit For
            ' bug out when inequality is found
        Next

        ' remove overlays
        CopyMemory ByVal ArrayPtr(tBytes1), 0&, 4&
        CopyMemory ByVal ArrayPtr(tBytes2), 0&, 4&

    End If
    
    ' return result(s)
    vbMemCompareR = (Index = 0)
    If Not vbMemCompareR Then UnEqualZeroBasedBytePos = Index - 1 ' zero based position

End Function
