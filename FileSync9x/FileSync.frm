VERSION 5.00
Begin VB.Form frmFileSync 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " FileSync - Run junk file remover first!"
   ClientHeight    =   2985
   ClientLeft      =   4635
   ClientTop       =   3570
   ClientWidth     =   7680
   Icon            =   "FileSync.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkUseCRC 
      BackColor       =   &H00DDDDDD&
      Caption         =   "&Generate file CRCs for matching (slower)  Exact matches over-rule time stamps"
      Enabled         =   0   'False
      Height          =   240
      Left            =   330
      TabIndex        =   23
      Top             =   2640
      Width           =   5940
   End
   Begin VB.CheckBox chkRemoveEmpty 
      BackColor       =   &H00DDDDDD&
      Caption         =   "&Also remove empty folders from location B after operation is complete"
      Enabled         =   0   'False
      Height          =   240
      Left            =   330
      TabIndex        =   22
      Top             =   2385
      Value           =   1  'Checked
      Width           =   5430
   End
   Begin VB.CommandButton cmdStartOp 
      Caption         =   "&Start Operation"
      Default         =   -1  'True
      Height          =   315
      Left            =   5970
      TabIndex        =   15
      Top             =   2175
      Width           =   1515
   End
   Begin VB.HScrollBar hsbJobs 
      Height          =   270
      Left            =   6945
      Max             =   0
      TabIndex        =   3
      Top             =   210
      Width           =   480
   End
   Begin VB.CommandButton cmdIncremental 
      Caption         =   "Cr&eate..."
      Height          =   300
      Left            =   6585
      TabIndex        =   12
      Top             =   1450
      Width           =   900
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Bro&wse..."
      Height          =   300
      Index           =   2
      Left            =   6585
      TabIndex        =   9
      Top             =   1110
      Width           =   900
   End
   Begin VB.CommandButton cmdBrowse 
      BackColor       =   &H80000014&
      Caption         =   "&Browse..."
      Height          =   300
      Index           =   1
      Left            =   6585
      TabIndex        =   6
      Top             =   770
      Width           =   900
   End
   Begin VB.PictureBox Picture1 
      Height          =   225
      Left            =   8655
      ScaleHeight     =   165
      ScaleWidth      =   1095
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2310
      Width           =   1155
   End
   Begin VB.CheckBox chkKillOrphans 
      BackColor       =   &H00DDDDDD&
      Caption         =   "&Remove files in location B that are deleted/renamed in source location A"
      Height          =   240
      Left            =   330
      TabIndex        =   14
      Top             =   2130
      Width           =   5430
   End
   Begin VB.OptionButton optJob 
      BackColor       =   &H00DDDDDD&
      Caption         =   "&Copy files in A that are New or Newer than B to location C"
      Height          =   255
      Index           =   0
      Left            =   330
      TabIndex        =   0
      Top             =   135
      Value           =   -1  'True
      Width           =   4425
   End
   Begin VB.OptionButton optJob 
      BackColor       =   &H00DDDDDD&
      Caption         =   "&Delete Duplicates from location B that exist in both A and B"
      Height          =   255
      Index           =   1
      Left            =   330
      TabIndex        =   1
      Top             =   390
      Width           =   4515
   End
   Begin VB.CheckBox chkAnyLoc 
      BackColor       =   &H00DDDDDD&
      Caption         =   "&Match files within any sub-folder locations with confirmation if ambiguous"
      Height          =   240
      Left            =   330
      TabIndex        =   13
      Top             =   1875
      Width           =   5430
   End
   Begin VB.TextBox txtPaths 
      Height          =   315
      Index           =   3
      Left            =   510
      TabIndex        =   11
      Top             =   1440
      Width           =   6000
   End
   Begin VB.TextBox txtPaths 
      Height          =   315
      Index           =   2
      Left            =   510
      TabIndex        =   8
      Top             =   1095
      Width           =   6000
   End
   Begin VB.TextBox txtPaths 
      Height          =   315
      Index           =   1
      Left            =   510
      TabIndex        =   5
      Top             =   750
      Width           =   6000
   End
   Begin VB.ListBox lstDups 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   240
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5055
      Width           =   7200
   End
   Begin VB.ListBox lstExists 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   240
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3195
      Width           =   7200
   End
   Begin VB.Image imgDelJob 
      Height          =   120
      Left            =   7140
      Picture         =   "FileSync.frx":030A
      ToolTipText     =   "Click to not remember current job"
      Top             =   495
      Width           =   150
   End
   Begin VB.Image imgSwap 
      Height          =   225
      Left            =   360
      Picture         =   "FileSync.frx":037B
      ToolTipText     =   "Click to swap A and B"
      Top             =   975
      Width           =   150
   End
   Begin VB.Label lblHide 
      BackStyle       =   0  'Transparent
      Caption         =   "^ Hide ^"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   6690
      TabIndex        =   21
      Top             =   2970
      Width           =   660
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Next default backup job"
      Height          =   240
      Left            =   5070
      TabIndex        =   2
      Top             =   255
      Width           =   1785
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Duplicate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   270
      TabIndex        =   18
      Top             =   4800
      Width           =   1635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File Exists"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   270
      TabIndex        =   16
      Top             =   2955
      Width           =   1635
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   -30
      TabIndex        =   10
      Top             =   1470
      Width           =   405
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   -30
      TabIndex        =   7
      Top             =   1155
      Width           =   405
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   -30
      TabIndex        =   4
      Top             =   825
      Width           =   405
   End
End
Attribute VB_Name = "frmFileSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const sCAPT As String = " FileSync - "
Private Const sSYNC As String = "SyncJobs.ini"

Private Declare Function GetStatus Lib "user32" Alias "GetQueueStatus" _
    (ByVal uFlags As Long) As Long

Private Const QS_KEY = &H1&         ' WM_KEYUP, WM_KEYDOWN, WM_SYSKEYUP, or WM_SYSKEYDOWN
Private Const QS_MOUSEBUTTON = &H4& ' Mouse-button message (WM_LBUTTONUP, WM_RBUTTONDOWN, etc)
Private Const QS_FLAGS = QS_KEY Or QS_MOUSEBUTTON

Private Type DefaultJobs
   Id(1 To 3) As Long
End Type

Private taJobs() As DefaultJobs
Private iJobs As Long

Private saPaths() As String
Private iPaths As Long

Private fCancel As Long

Private cFindSrcFiles As cFileSync
Private cFindDstFiles As cFileSync

Private cCRC As cCRC32
Private cCmp As cMemComp

Private rCaptH As Single
Private sINI As String

Private Sub cmdStartOp_Click()
    If cmdStartOp.Caption = "&Start Operation" Then
       fCancel = 0&
       cmdStartOp.Caption = "&Stop Operation"
       If optJob(0) Then
           If UpdateJob Then AddDefaultJob ' Copy Newer
       ElseIf DeleteJob Then AddDefaultJob ' Kill Dups
       End If
       cmdStartOp.Caption = "&Start Operation"
    Else
       fCancel = -1&
    End If
End Sub

Private Function UpdateJob() As Long ' Copy New/Newer
   If InAnExe Then On Error GoTo ExitFunc
    Dim sSrc1Path As String
    Dim sSrc2Path As String
    Dim sDestPath As String
    Dim i As Long, j As Long, k As Long
    Dim fMatchRelSpec As Long
    Dim fRemoveFlag As Long
    Dim rTotal As Single
    Dim increm As Long
    Dim iUpTo As Long
    Dim iCnt As Long
    Dim iCompare As Long

    fMatchRelSpec = (chkAnyLoc = 0)

    If fMatchRelSpec Then 'chkKillOrphans.Enabled
       fRemoveFlag = (chkKillOrphans = 1)
       If fRemoveFlag Then
           If MsgBox("Are you sure you want to remove destination files no longer in source location?  ", vbOKCancel Or vbQuestion, _
                    " Delete Confirmation") = vbCancel Then Exit Function
       End If
    Else
       If MsgBox("Are you sure you want to copy matching files from any location?  ", vbOKCancel Or vbQuestion, _
                " Location Confirmation") = vbCancel Then Exit Function
    End If

    HourGlass True
    Me.Caption = sCAPT & "scanning ."

    sSrc1Path = AddSlash(txtPaths(1)) 'A
    sSrc2Path = AddSlash(txtPaths(2)) 'B

    If Not DirExists(sSrc1Path) Then GoTo ExitFunc
    If Not DirExists(sSrc2Path) Then GoTo ExitFunc

    sDestPath = AddSlash(txtPaths(3))
    If CreatePath(sDestPath) = 0& Then GoTo ExitFunc

    Me.Caption = sCAPT & "scanning . ."

    cFindSrcFiles.RelativePaths = fMatchRelSpec
    cFindSrcFiles.FindAllFiles sSrc1Path

    If GetStatus(QS_FLAGS) Then DoEvents
    If fCancel Then GoTo ExitFunc
    Me.Caption = sCAPT & "scanning . . ."

    cFindDstFiles.RelativePaths = fMatchRelSpec
    cFindDstFiles.FindAllFiles sSrc2Path

    If GetStatus(QS_FLAGS) Then DoEvents
    If fCancel Then GoTo ExitFunc

    rTotal = cFindDstFiles.FilesUB
    increm = CLng(rTotal / 100!)
    If increm = 0& Then increm = 1& 'Total < 100

    If fRemoveFlag Then ' Remove orphans from backup location
       Me.Caption = sCAPT & "removing . . . ."

       Do Until i > cFindDstFiles.FilesUB

          For j = iUpTo To cFindSrcFiles.FilesUB
             iCompare = StrComp(cFindDstFiles.RelFileSpec(i), cFindSrcFiles.RelFileSpec(j))

             If iCompare = Equal Then
                 iUpTo = j ' Sorted data so skip all prev src items from next tests
                 Exit For  ' Found match so skip to next backup item

             ElseIf iCompare = Greater Then
                 iUpTo = j ' Sorted data so skip all prev src items from next tests

             ElseIf iCompare = Lesser Then
                 j = cFindSrcFiles.FilesUB ' Sorted data so orphaned backup item
             End If
          Next j

          If j > cFindSrcFiles.FilesUB Then

             ' Remove destination files no longer in source location
             SetAttributes cFindDstFiles(i), vbNormal
             Kill cFindDstFiles(i)

             ' Delete data record to speed up process below
             cFindDstFiles.RemoveFile i
             rTotal = rTotal - 1!
             increm = CLng(rTotal / 100!)
             If increm = 0& Then increm = 1&
          Else
             i = i + 1&
          End If
          If i Mod increm = 0& Then
             Me.Caption = sCAPT & CStr(CLng(100! * i / rTotal)) & " %"
          End If
       Loop
       Me.Caption = sCAPT & "removing . . . . ."

       If GetStatus(QS_FLAGS) Then DoEvents
       If fCancel Then GoTo ExitFunc

       On Error Resume Next
       If (chkRemoveEmpty = 1) Then ' Remove empty folders
          For j = cFindDstFiles.FoldersUB To cFindDstFiles.FoldersLB Step -1
             If cFindDstFiles.FolderIsEmpty(j) Then ' Is valid empty folder
               ' Remove destination folder no longer containing fso's
               SetAttributes cFindDstFiles.FolderSpec(j), vbNormal
               RmDir cFindDstFiles.FolderSpec(j)
             End If
          Next j
       End If
       On Error GoTo 0
    End If
    If InAnExe Then On Error GoTo ExitFunc

    Me.Caption = sCAPT & "working . . . . . ."

    If GetStatus(QS_FLAGS) Then DoEvents
    If fCancel Then GoTo ExitFunc

    rTotal = cFindSrcFiles.FilesUB
    increm = CLng(rTotal / 100!)
    If increm = 0& Then increm = 1&
    iUpTo = 0&

    For i = 0& To rTotal
        For j = iUpTo To cFindDstFiles.FilesUB

           If GetStatus(QS_FLAGS) Then DoEvents
           If fCancel Then GoTo ExitFunc

           If fMatchRelSpec Then
               iCompare = StrComp(cFindSrcFiles.RelFileSpec(i), cFindDstFiles.RelFileSpec(j))
           Else
               iCompare = StrComp(cFindSrcFiles.FileName(i), cFindDstFiles.FileName(j))
           End If
           If iCompare = Equal Then

               iUpTo = j ' Sorted data so skip all prev dest items from next tests
               iCompare = CompFileTime(cFindSrcFiles.FileLastWrite(i), cFindDstFiles.FileLastWrite(j))

               If iCompare = ftNewer Then

                  If fMatchRelSpec Then

                     CreatePath sDestPath & cFindSrcFiles.RelFileSpec(i)
                     SetAttributes sDestPath & cFindSrcFiles.RelFileSpec(i), vbNormal
                     FileCopy cFindSrcFiles(i), sDestPath & cFindSrcFiles.RelFileSpec(i)
                     iCnt = iCnt + 1&

                  Else 'If cFindSrcFiles.FileSize(i) <> cFindDstFiles.FileSize(j) Then

                     Me.Height = rCaptH + 6780!
                     With cFindDstFiles
                         lstExists.Clear
                         lstExists.AddItem "File exists:"
                         AddMakeFit lstExists, .FileSpec(j)
                         lstExists.AddItem "Created:     " & FileTimeToString(.FileCreated(j))
                         lstExists.AddItem "Last Mod:    " & FileTimeToString(.FileLastWrite(j))
                         lstExists.AddItem "File Size:   " & .FileSize(j) & " bytes"
                     End With

                     With cFindSrcFiles
                         lstDups.Clear
                         lstDups.AddItem "Also exists:"
                         AddMakeFit lstDups, .FileSpec(i)
                         lstDups.AddItem "Created:     " & FileTimeToString(.FileCreated(i))
                         lstDups.AddItem "Last Mod:    " & FileTimeToString(.FileLastWrite(i))
                         lstDups.AddItem "File Size:   " & .FileSize(i) & " bytes"
                     End With

                     k = MsgBox("Do you want to transfer this file?", vbYesNoCancel)

                     If k = vbCancel Then GoTo ExitFunc

                     If k = vbYes Then

                        CreatePath sDestPath & cFindSrcFiles.RelFileSpec(i)
                        SetAttributes sDestPath & cFindSrcFiles.RelFileSpec(i), vbNormal
                        FileCopy cFindSrcFiles(i), sDestPath & cFindSrcFiles.RelFileSpec(i)
                        lstDups.AddItem "Transfered"
                        iCnt = iCnt + 1&

                     End If
                  End If
               End If

               If fMatchRelSpec Then Exit For

           ElseIf iCompare = Greater Then
               iUpTo = j ' Sorted data so skip all prev dest items from next tests

           ElseIf iCompare = Lesser Then
               j = cFindDstFiles.FilesUB ' Sorted data so new src item
           End If
        Next j

        If j > cFindDstFiles.FilesUB Then
            CreatePath sDestPath & cFindSrcFiles.RelFileSpec(i)
            FileCopy cFindSrcFiles(i), sDestPath & cFindSrcFiles.RelFileSpec(i)
            iCnt = iCnt + 1&
        End If

        If i Mod increm = 0& Then
           Me.Caption = sCAPT & CStr(CLng(100! * i / rTotal)) & " %"
        End If
    Next i
    UpdateJob = -1&

ExitFunc:
    Me.Caption = sCAPT & iCnt & " file(s) updated successfully"
    HourGlass False

    If Err Then AlertError "frmFileSync.UpdateJob", DebugPrint
End Function

Private Function DeleteJob() As Long ' Delete Duplicates
   If InAnExe Then On Error GoTo ExitFunc
    Dim sSrcPath As String
    Dim sDestPath As String
    Dim i As Long, j As Long, k As Long
    Dim fMatchRelSpec As Long
    Dim fUseCRC As Long
    Dim iCompare As Long
    Dim rTotal As Single
    Dim increm As Long
    Dim iUpTo As Long
    Dim iCnt As Long

    Dim srcA() As Byte
    Dim dstA() As Byte
    Dim iFile As Long

    fUseCRC = (chkUseCRC = 1)
    fMatchRelSpec = (chkAnyLoc = 0)

    If fMatchRelSpec Then
       If MsgBox("Are you sure you want to delete files?  ", vbOKCancel Or vbQuestion, _
                " Delete Confirmation") = vbCancel Then Exit Function
    Else
       If MsgBox("Are you sure you want to delete matching files from any location?  ", vbOKCancel Or vbQuestion, _
                " Delete Confirmation") = vbCancel Then Exit Function
    End If

    HourGlass True
    Me.Caption = sCAPT & "scanning ."

    sSrcPath = AddSlash(txtPaths(1)) 'A
    sDestPath = AddSlash(txtPaths(2)) 'B

    If Not DirExists(sSrcPath) Then GoTo ExitFunc
    If Not DirExists(sDestPath) Then GoTo ExitFunc

    Me.Caption = sCAPT & "scanning . ."

    cFindSrcFiles.RelativePaths = fMatchRelSpec
    cFindSrcFiles.FindAllFiles sSrcPath

    If GetStatus(QS_FLAGS) Then DoEvents
    If fCancel Then GoTo ExitFunc
    Me.Caption = sCAPT & "scanning . . ."

    cFindDstFiles.RelativePaths = fMatchRelSpec
    cFindDstFiles.FindAllFiles sDestPath

    If GetStatus(QS_FLAGS) Then DoEvents
    If fCancel Then GoTo ExitFunc
    Me.Caption = sCAPT & "working . . . ."

    rTotal = cFindDstFiles.FilesUB
    increm = CLng(rTotal / 100!)
    If increm = 0& Then increm = 1&

    For i = 0& To rTotal
        For j = iUpTo To cFindSrcFiles.FilesUB

           If GetStatus(QS_FLAGS) Then DoEvents
           If fCancel Then GoTo ExitFunc

           If fMatchRelSpec Then
               iCompare = StrComp(cFindDstFiles.RelFileSpec(i), cFindSrcFiles.RelFileSpec(j))
           Else
               iCompare = StrComp(cFindDstFiles.FileName(i), cFindSrcFiles.FileName(j))
           End If

           If iCompare = Equal Then
               If fMatchRelSpec Then iUpTo = j ' Sorted data so skip all prev src items from next tests

               If fUseCRC Then
                   If cFindDstFiles.FileSize(i) = cFindSrcFiles.FileSize(j) Then
                       If cFindDstFiles.FileCRC(i) = 0& Then
                           iFile = FreeFile
                           Open cFindDstFiles(i) For Binary Access Read Lock Write As #iFile
                               ReDim dstA(1 To LOF(iFile))
                               Get #iFile, 1&, dstA()
                           Close #iFile
                           cFindDstFiles.FileCRC(i) = -1&
                          'cFindDstFiles.FileCRC(i) = cCRC.CalcCRC(dstA)
                       End If

                       iFile = FreeFile
                       Open cFindSrcFiles(j) For Binary Access Read Lock Write As #iFile
                           ReDim srcA(1 To LOF(iFile))
                           Get #iFile, 1&, srcA()
                       Close #iFile
                       iCompare = cCmp.vbMemCompare(VarPtr(srcA(1)), VarPtr(dstA(1)), cFindSrcFiles.FileSize(j))
                      'iCompare = (cCRC.CalcCRC(srcA) = cFindDstFiles.FileCRC(i))
                   End If
               End If

               If iCompare Then
               ElseIf fMatchRelSpec Then
                   iCompare = CompFileTime(cFindDstFiles.FileLastWrite(i), cFindSrcFiles.FileLastWrite(j)) <> ftNewer
               Else
                   iCompare = CompFileTime(cFindDstFiles.FileLastWrite(i), cFindSrcFiles.FileLastWrite(j)) = ftEqual
               End If
               If iCompare Then

                   ' File may already be removed if 'any sub-folder' was specified
                   If PathExists(cFindDstFiles(i)) Then

                       If cFindDstFiles.FileSize(i) = cFindSrcFiles.FileSize(j) Then
                           ' Byte size match with older or equal date
                           SetAttributes cFindDstFiles(i), vbNormal
                           Kill cFindDstFiles(i)
                           iCnt = iCnt + 1&

                       Else
                           Me.Height = rCaptH + 6780!
                           With cFindSrcFiles
                               lstExists.Clear
                               lstExists.AddItem "File exists:"
                               AddMakeFit lstExists, .FileSpec(j)
                               lstExists.AddItem "Created:     " & FileTimeToString(.FileCreated(j))
                               lstExists.AddItem "Last Mod:    " & FileTimeToString(.FileLastWrite(j))
                               lstExists.AddItem "File Size:   " & .FileSize(j) & " bytes"
                           End With

                           With cFindDstFiles
                               lstDups.Clear
                               lstDups.AddItem "Also exists:"
                               AddMakeFit lstDups, .FileSpec(i)
                               lstDups.AddItem "Created:     " & FileTimeToString(.FileCreated(i))
                               lstDups.AddItem "Last Mod:    " & FileTimeToString(.FileLastWrite(i))
                               lstDups.AddItem "File Size:   " & .FileSize(i) & " bytes"
                           End With

                           k = MsgBox("Do you want to delete this file?", vbYesNoCancel)

                           If k = vbCancel Then GoTo ExitFunc

                           If k = vbYes Then
                              SetAttributes cFindDstFiles(i), vbNormal
                              Kill cFindDstFiles(i)
                              lstDups.AddItem "Deleted"
                              iCnt = iCnt + 1&
                           End If

                       End If 'FileSize
                   End If 'Still Exists
               End If 'FileTime

               If fMatchRelSpec Then Exit For

           ElseIf iCompare = Greater Then
               iUpTo = j ' Sorted data so skip all prev src items from next tests

           ElseIf iCompare = Lesser Then ' No duplicate found
               Exit For ' Sorted data so skip to next
           End If 'FileName
        Next j

        If i Mod increm = 0& Then
           Me.Caption = sCAPT & CStr(CLng(100! * i / rTotal)) & " %"
        End If
    Next i
    On Error Resume Next
    If (chkRemoveEmpty = 1) Then ' Remove empty folders
       For j = cFindDstFiles.FoldersUB To cFindDstFiles.FoldersLB Step -1
          If cFindDstFiles.FolderIsEmpty(j) Then ' Is valid empty folder
            ' Remove destination folder no longer containing fso's
            SetAttributes cFindDstFiles.FolderSpec(j), vbNormal
            RmDir cFindDstFiles.FolderSpec(j)
          End If
       Next j
    End If
    DeleteJob = -1&

ExitFunc:
    Me.Caption = sCAPT & iCnt & " duplicates removed successfully"
    HourGlass False

    If Err Then AlertError "frmFileSync.DeleteJob", DebugPrint
End Function

Private Sub hsbJobs_Change()
    Dim i As Long
    i = hsbJobs.Value - 1&
    If i < 0& Then
       hsbJobs.Min = 0
    Else
       txtPaths(1) = saPaths(taJobs(i).Id(1))
       txtPaths(2) = saPaths(taJobs(i).Id(2))
       txtPaths(3) = saPaths(taJobs(i).Id(3))
       hsbJobs.Min = 1
    End If
End Sub

Private Sub SaveDefaultJobs()
    Dim i As Long, j As Long
    Do Until j = iPaths
       SetINIKey sINI, "Paths", "Path" & j, saPaths(j)
       j = j + 1&
    Loop
    Do Until i = iJobs
       SetINIKey sINI, "Jobs", "Job" & i, taJobs(i).Id(1) & "|" & _
                                          taJobs(i).Id(2) & "|" & _
                                          taJobs(i).Id(3)
       i = i + 1&
    Loop
End Sub

Private Sub GetDefaultJobs()
    Dim s As String
    Dim i As Long, j As Long

    s = GetINIKey(sINI, "Paths", "Path0")
    Do While LenB(s)
       ReDim Preserve saPaths(-1 To iPaths) As String
       saPaths(iPaths) = s
       iPaths = iPaths + 1&
       s = GetINIKey(sINI, "Paths", "Path" & iPaths)
    Loop

    s = GetINIKey(sINI, "Jobs", "Job0")
    Do While LenB(s)
       ReDim Preserve taJobs(iJobs) As DefaultJobs
       i = InStr(s, "|")
       taJobs(iJobs).Id(1) = Left$(s, i - 1&)

       i = i + 1&
       j = InStr(i, s, "|")
       taJobs(iJobs).Id(2) = Mid$(s, i, j - i)
       taJobs(iJobs).Id(3) = Mid$(s, j + 1&)

       iJobs = iJobs + 1&
       s = GetINIKey(sINI, "Jobs", "Job" & iJobs)
    Loop
    hsbJobs.Max = iJobs
End Sub

Private Sub AddDefaultJob()
    Dim i As Long, j As Long
    Dim tJob As DefaultJobs
    For i = 1& To 3&
       If LenB(txtPaths(i)) Then
          Do Until j = iPaths
             If StrComp(txtPaths(i), saPaths(j), TextCompare) = Equal Then Exit Do
             j = j + 1&
          Loop
          If j = iPaths Then
             ReDim Preserve saPaths(-1 To iPaths) As String
             saPaths(iPaths) = txtPaths(i)
             iPaths = iPaths + 1&
          End If
          tJob.Id(i) = j
          j = 0&
       Else
          tJob.Id(i) = -1& ' saPaths(-1) == sEmpty
       End If
    Next i
    i = 0&
    Do Until i = iJobs
       If taJobs(i).Id(1) = tJob.Id(1) And _
          taJobs(i).Id(2) = tJob.Id(2) And _
          taJobs(i).Id(3) = tJob.Id(3) Then Exit Do
       i = i + 1&
    Loop
    If i = iJobs Then
       ReDim Preserve taJobs(iJobs) As DefaultJobs
       taJobs(iJobs).Id(1) = tJob.Id(1)
       taJobs(iJobs).Id(2) = tJob.Id(2)
       taJobs(iJobs).Id(3) = tJob.Id(3)
       iJobs = iJobs + 1&
    End If
    hsbJobs.Max = iJobs
End Sub

Private Sub imgDelJob_Click()
    Dim i As Long
    i = hsbJobs.Value - 1&
    If i < 0& Then Exit Sub
    If StrComp(txtPaths(1), saPaths(taJobs(i).Id(1)), TextCompare) = Equal And _
       StrComp(txtPaths(2), saPaths(taJobs(i).Id(2)), TextCompare) = Equal And _
       StrComp(txtPaths(3), saPaths(taJobs(i).Id(3)), TextCompare) = Equal Then
       iJobs = iJobs - 1&
       Do While i < iJobs
          taJobs(i).Id(1) = taJobs(i + 1).Id(1)
          taJobs(i).Id(2) = taJobs(i + 1).Id(2)
          taJobs(i).Id(3) = taJobs(i + 1).Id(3)
          SetINIKey sINI, "Jobs", "Job" & i, GetINIKey(sINI, "Jobs", "Job" & i + 1&)
          i = i + 1&
       Loop
       DelINIKey sINI, "Jobs", "Job" & i
       hsbJobs.Max = iJobs
       If iJobs Then
          hsbJobs_Change
       Else
          txtPaths(1) = vbNullString
          txtPaths(2) = vbNullString
          txtPaths(3) = vbNullString
       End If
    End If
End Sub

Private Sub imgSwap_Click()
   Dim s As String
   s = txtPaths(2)
   txtPaths(2) = txtPaths(1)
   txtPaths(1) = s
End Sub

Private Sub lblHide_Click()
   Me.Height = rCaptH + 2980!
End Sub

Private Sub optJob_Click(Index As Integer) ' Delete dups Or (delete orphans And orphans enabled)
   chkRemoveEmpty.Enabled = optJob(1) Or (chkKillOrphans = 1 And chkAnyLoc = 0)
   chkKillOrphans.Enabled = optJob(0) And (chkAnyLoc = 0) ' Copy newer And relative path
   chkUseCRC.Enabled = optJob(1)
End Sub

Private Sub chkAnyLoc_Click()              ' Delete dups Or (delete orphans And orphans enabled)
   chkRemoveEmpty.Enabled = optJob(1) Or (chkKillOrphans = 1 And chkAnyLoc = 0)
   chkKillOrphans.Enabled = optJob(0) And (chkAnyLoc = 0) ' Copy newer And relative path
End Sub

Private Sub chkKillOrphans_Click()
   chkRemoveEmpty.Enabled = (chkKillOrphans = 1) ' Delete orphans And orphans enabled
End Sub

Private Sub Form_Load()
    Set cFindSrcFiles = New cFileSync
    Set cFindDstFiles = New cFileSync
    Set cCRC = New cCRC32
    Set cCmp = New cMemComp
    Set Picture1.Font = lstExists.Font
    rCaptH = Me.Height - Me.ScaleHeight
    Me.Height = rCaptH + 2980!
    If Right$(App.Path, 1&) = "\" Then
        sINI = App.Path & sSYNC
    Else
        sINI = App.Path & "\" & sSYNC
    End If
    GetDefaultJobs
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cFindSrcFiles = Nothing
    Set cFindDstFiles = Nothing
    Set cCRC = Nothing
    Set cCmp = Nothing
    SaveDefaultJobs
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
    Dim s As String
    s = Browse(Me.hWnd, , , Display_File_Folders)
    If LenB(s) Then
        txtPaths(Index) = s
        If Index = 2 Then If LenB(txtPaths(3)) = 0& Then txtPaths(3) = s
    End If
End Sub

Private Sub cmdIncremental_Click()
    Me.Caption = sCAPT & "working"
    CreatePath AddSlash(txtPaths(3))
    Me.Caption = sCAPT & "done"
End Sub

Private Sub txtPaths_GotFocus(Index As Integer)
    txtPaths(Index).SelStart = 0&
    txtPaths(Index).SelLength = Len(txtPaths(Index))
End Sub

' Make path string fit if too long
Private Sub AddMakeFit(lst As ListBox, sEntry As String)
    Dim i As Long, r As Single, s As String
    s = sEntry: i = 13&
    r = lst.Width - 30!
    Do While Picture1.TextWidth(s) > r
        s = Left$(sEntry, 8&) & "..." & Mid$(sEntry, i)
        i = i + 1&
    Loop
    lst.AddItem s
End Sub
