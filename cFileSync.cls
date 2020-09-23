VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' cFileSync - Demonstrates the use of the Find Files API functions
' to find all files within the specified folder and all sub-folders.

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Const MAX_PATH = 260&
Private Const ALT_NAME = 14&

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As Currency
    ftLastAccessTime As Currency
    ftLastWriteTime As Currency
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternateFileName As String * ALT_NAME ' 8.3 format
End Type

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" _
    (ByVal lpFileName As Long, ByVal lpFindData As Long) As Long 'WIN32_FIND_DATA

Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileW" _
    (ByVal hFindFile As Long, ByVal lpFindData As Long) As Long 'WIN32_FIND_DATA

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const INVALID_HANDLE_VALUE = &HFFFFFFFF

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Const vbMAXOUTDW As Double = (2 ^ 32) ' 4294967296 == 4.2 Gig
Private Const MINLONG As Long = &H80000000    ' 2147483648 == 2.1 Gig

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Declare Function GetAttributes Lib "kernel32" Alias "GetFileAttributesW" _
    (ByVal lpSpec As Long) As Long

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'' CopyMemory Alias thanks to Bruce Mc :)
'Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" _
'    (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lLenB As Long)
'
''Randy Birch, VBnet.com
'Private Declare Function StrLenW Lib "kernel32" Alias "lstrlenW" _
'    (ByVal lpString As Long) As Long

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private lbs() As Long, ubs() As Long ' Non-recursive quicksort stacks
Private lStackSize As Long

Private Const BUFF_SIZE = 1024&

Private aWFD() As WIN32_FIND_DATA
Private aFiles() As String
Private aNames() As String
Private aDirs() As String
Private laIdx() As Long
Private laDirs() As Long

Private mRelative As Long
Private mRootLen As Long
Private mFileCnt As Long
Private mDirCnt As Long
Private mUBfiles As Long
Private mUBdirs As Long

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' This sub-routine demonstrates the use of the Find Files API functions
' to find all files within the specified folder and all sub-folders.
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub FindAllFiles(sPath As String)
   Dim hFind As Long, sP As String, cnt As Long, ub As Long

   If (GetAttributes(StrPtr(sPath)) And vbDirectory) = 0& Then Exit Sub

   If Right$(sPath, 1&) = "\" Then sP = sPath Else sP = sPath & "\"
   aDirs(0) = sP: mRootLen = Len(sP)
   mFileCnt = 0&: mDirCnt = 0&

   hFind = FindFirstFile(StrPtr("\\?\" & sP & "*"), VarPtr(aWFD(0)))
   Do Until (hFind = INVALID_HANDLE_VALUE)
      Do
         If (aWFD(mFileCnt).dwFileAttributes And vbDirectory) = vbDirectory Then
            If Not AscW(aWFD(mFileCnt).cFileName) = 46 Then ' <> "." or ".."
                aDirs(mDirCnt) = sP & RTrimZ(aWFD(mFileCnt).cFileName)
                mDirCnt = mDirCnt + 1&

                If mDirCnt > mUBdirs Then
                    mUBdirs = mDirCnt + BUFF_SIZE
                    ReDim Preserve aDirs(0 To mUBdirs) As String
                End If
            End If
         Else
             aNames(mFileCnt) = RTrimZ(aWFD(mFileCnt).cFileName)
             aFiles(mFileCnt) = sP & aNames(mFileCnt)
             mFileCnt = mFileCnt + 1&

             If mFileCnt > mUBfiles Then
                 mUBfiles = mFileCnt + BUFF_SIZE
                 ReDim Preserve aFiles(0 To mUBfiles) As String
                 ReDim Preserve aNames(0 To mUBfiles) As String
                 ReDim Preserve aWFD(0 To mUBfiles) As WIN32_FIND_DATA
             End If
         End If
      Loop Until (FindNextFile(hFind, VarPtr(aWFD(mFileCnt))) = 0&)

      FindClose hFind
      If cnt = mDirCnt Then Exit Do ' No more folders

      sP = aDirs(cnt) & "\"
      cnt = cnt + 1&
      hFind = FindFirstFile(StrPtr("\\?\" & sP & "*"), VarPtr(aWFD(mFileCnt)))
   Loop

   If (mFileCnt) Then
      ub = mFileCnt - 1&
      ReDim laIdx(0 To ub) As Long
      For cnt = 1& To ub
         laIdx(cnt) = cnt
      Next
      If mRelative Then
         SwapSort4i aFiles, laIdx, 0&, ub
      Else
         SwapSort4i aNames, laIdx, 0&, ub
      End If
   End If

   If (mDirCnt) Then
      ub = mDirCnt - 1&
      ReDim laDirs(0 To ub) As Long
      For cnt = 1& To ub
         laDirs(cnt) = cnt
      Next
      SwapSort4i aDirs, laDirs, 0&, ub
   End If

End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Property Get RelativePaths() As Boolean
    RelativePaths = mRelative
End Property

Public Property Let RelativePaths(ByVal lValue As Boolean)
    mRelative = lValue
End Property

Public Sub RemoveFile(ByVal idx As Long)
    Dim lp As Long
    mFileCnt = mFileCnt - 1& 'ub
    If mFileCnt Then '0 to ub
        If idx < mFileCnt Then 'lA(idx) = lA(idx + 1)...
            lp = VarPtr(laIdx(idx))
            CopyMemByV lp, lp + 4&, (mFileCnt - idx) * 4&
        End If
    End If
End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function FolderIsEmpty(ByVal FolderIdx As Long) As Boolean
   Dim tWFD As WIN32_FIND_DATA
   Dim hFind As Long
   If FolderIdx > -1& And FolderIdx < mDirCnt Then
      hFind = FindFirstFile(StrPtr("\\?\" & aDirs(laDirs(FolderIdx)) & "\*"), VarPtr(tWFD))
      If Not (hFind = INVALID_HANDLE_VALUE) Then
         FolderIsEmpty = True ' Is valid folder, default to empty
         Do
            If Not AscW(tWFD.cFileName) = 46 Then ' <> "." or ".."
               FolderIsEmpty = False
               Exit Do
            End If
         Loop Until (FindNextFile(hFind, VarPtr(tWFD)) = 0&)
         FindClose hFind
      End If
   End If
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Property Get FoldersLB() As Long
    FoldersLB = 0&
End Property

Public Property Get FoldersUB() As Long
    FoldersUB = mDirCnt - 1&
End Property

Public Property Get FilesLB() As Long
    FilesLB = 0&
End Property

Public Property Get FilesUB() As Long
    FilesUB = mFileCnt - 1&
End Property

Public Property Get FileName(ByVal idx As Long) As String
    FileName = aNames(laIdx(idx))
End Property

Public Property Get RelFileSpec(ByVal idx As Long) As String
    RelFileSpec = Mid$(aFiles(laIdx(idx)), mRootLen + 1&) ', Len(sFileSpec) - mRootLen
End Property

Public Property Get FileSpec(ByVal idx As Long) As String 'Default Property
Attribute FileSpec.VB_UserMemId = 0
    FileSpec = aFiles(laIdx(idx))
End Property

Public Property Get RootPath() As String
    RootPath = Left$(aDirs(0), mRootLen)
End Property

Public Property Get RootPathLength() As Long
    RootPathLength = mRootLen
End Property

Public Property Get RelFolderSpec(ByVal idx As Long) As String
    RelFolderSpec = Mid$(aDirs(laDirs(idx)), mRootLen + 1&) ', Len(sDir) - mRootLen
End Property

Public Property Get FolderSpec(ByVal idx As Long) As String
    FolderSpec = aDirs(laDirs(idx))
End Property

Public Property Get FileAttribs(ByVal idx As Long) As Long
    FileAttribs = aWFD(laIdx(idx)).dwFileAttributes
End Property

Public Property Get FileSize(ByVal idx As Long) As Double
    Dim dLo As Double, dHi As Double
    With aWFD(laIdx(idx))
       If .nFileSizeLow And MINLONG Then dLo = vbMAXOUTDW + .nFileSizeLow Else dLo = .nFileSizeLow
       If .nFileSizeHigh Then _
        If .nFileSizeHigh And MINLONG Then dHi = vbMAXOUTDW + .nFileSizeHigh Else dHi = .nFileSizeHigh
    End With
    FileSize = vbMAXOUTDW * dHi + dLo
End Property

Public Property Get FileCreated(ByVal idx As Long) As Currency
    FileCreated = aWFD(laIdx(idx)).ftCreationTime
End Property

Public Property Get FileLastWrite(ByVal idx As Long) As Currency
    FileLastWrite = aWFD(laIdx(idx)).ftLastWriteTime
End Property

Public Property Get FileLastAccess(ByVal idx As Long) As Currency
    FileLastAccess = aWFD(laIdx(idx)).ftLastAccessTime
End Property

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Property Get FileSizeLow(ByVal idx As Long) As Long
    FileSizeLow = aWFD(laIdx(idx)).nFileSizeLow
End Property

Public Property Get FileSizeHigh(ByVal idx As Long) As Long
    FileSizeHigh = aWFD(laIdx(idx)).nFileSizeHigh
End Property

Public Property Get FileCreatedLow(ByVal idx As Long) As Long
    CopyMemByV VarPtr(FileCreatedLow), VarPtr(aWFD(laIdx(idx)).ftCreationTime), 4&
End Property

Public Property Get FileCreatedHigh(ByVal idx As Long) As Long
    CopyMemByV VarPtr(FileCreatedHigh), VarPtr(aWFD(laIdx(idx)).ftCreationTime) + 4&, 4&
End Property

Public Property Get FileLastWriteLow(ByVal idx As Long) As Long
    CopyMemByV VarPtr(FileLastWriteLow), VarPtr(aWFD(laIdx(idx)).ftLastWriteTime), 4&
End Property

Public Property Get FileLastWriteHigh(ByVal idx As Long) As Long
    CopyMemByV VarPtr(FileLastWriteHigh), VarPtr(aWFD(laIdx(idx)).ftLastWriteTime) + 4&, 4&
End Property

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Property Let FileCRC(ByVal idx As Long, ByVal crc As Long)
    aWFD(laIdx(idx)).dwReserved0 = crc
End Property

Public Property Get FileCRC(ByVal idx As Long) As Long
    FileCRC = aWFD(laIdx(idx)).dwReserved0
End Property

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub Class_Initialize()
    mUBdirs = BUFF_SIZE
    mUBfiles = BUFF_SIZE
    ReDim aWFD(0 To mUBfiles) As WIN32_FIND_DATA
    ReDim aFiles(0 To mUBfiles) As String
    ReDim aNames(0 To mUBfiles) As String
    ReDim aDirs(0 To mUBdirs) As String
End Sub

Private Sub Class_Terminate()
    Erase aWFD() ' Clean up
    Erase aFiles()
    Erase aNames()
    Erase aDirs()
    Erase laDirs()
    Erase laIdx()
    Erase lbs()
    Erase ubs()
End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'Private Function RTrimZ(StrZ As String) As String
'    Dim lLen As Long
'    lLen = StrLenW(StrPtr(StrZ))
'    RTrimZ = LeftB$(StrZ, lLen + lLen) 'Rd
'End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub InitializeStacks(ByVal lSize As Long)
    Dim i As Long ' Allow for likely stack count + 32
    i = Log(lSize) / Log(2) + 32&  ' 10000 -> 13 + 32
    If i > lStackSize Then         ' 100000 -> 17 + 32
        ReDim lbs(0 To i) As Long  ' 1000000 -> 20 + 32
        ReDim ubs(0 To i) As Long  ' 10000000 -> 23 + 32
        lStackSize = i ' Stacks to hold pending boundries
    End If
End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub SwapSort4i(sA() As String, iA() As Long, ByVal lb As Long, ByVal ub As Long)
    ' This is my indexed non-recursive Quick-Sort, and is very very fast!
    Dim lo As Long, hi As Long
    Dim idx As Long, cnt As Long
    If ub - lb Then Else Exit Sub
    InitializeStacks ub - lb           ' Initialize pending boundary stacks
    Do: hi = (ub + lb) \ 2&            ' Get pivot index position
        idx = iA(hi): iA(hi) = iA(ub)  ' Grab current index
        lo = lb: hi = ub               ' Set bounds
        Do While (hi > lo)             ' Storm right in
            If Not (StrComp(sA(idx), sA(iA(lo))) = Ascending) Then ' (item <= lo)
                iA(hi) = iA(lo)
                hi = hi - 1&
                Do Until (hi = lo)
                    If Not (StrComp(sA(iA(hi)), sA(idx)) = Ascending) Then ' (hi <= item)
                        iA(lo) = iA(hi)
                        Exit Do
                    End If
                    hi = hi - 1&
                Loop
                If (hi = lo) Then Exit Do ' Found swaps or out of loop
            End If
            lo = lo + 1&
        Loop '-Blizzard v4i ©Rd-
        iA(hi) = idx               ' Re-assign current
        If (lb < lo - 1&) Then
            If (ub > lo + 1&) Then cnt = cnt + 1&: lbs(cnt) = lo + 1&: ubs(cnt) = ub
            ub = lo - 1&
        ElseIf (ub > lo + 1&) Then
            lb = lo + 1&
        Else
            If cnt = 0& Then Exit Do
            lb = lbs(cnt): ub = ubs(cnt): cnt = cnt - 1&
        End If
    Loop
End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Rd - crYptic but cRaZy!
