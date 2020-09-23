Attribute VB_Name = "modBrowse"
Option Explicit

' ------------------------------------------------------------©Rd-
'  Browse - Opens the familiar Browse Dialog (Treeview control)
'           that displays the system folders for selection.
' ----------------------------------------------------------------

Private Type BROWSEINFO
    hWndOwner      As Long
    pidlRoot       As Long
    pszDisplayName As String
    lpszTitle      As String
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderW" (ByVal lpBI As Long) As Long 'BROWSEINFO
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListW" (ByVal pidList As Long, ByVal lpBuffer As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long

Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pIDLMem As Long)
Private Declare Sub OleInitialize Lib "ole32" (ByVal lNull As Long)

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpWndpl As WINDOWPLACEMENT) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_GETFONT = &H31&
Private Const WM_SETFONT = &H30&

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Const HWND_TOP As Long = 0
Private Const SWP_SHOWWINDOW = &H40&

Private Declare Function GetAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpSpec As Long) As Long

Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lLenB As Long)
Private Declare Sub ZeroMemByV Lib "kernel32" Alias "RtlZeroMemory" (ByVal lpDest As Long, ByVal lLenB As Long)
Private Declare Function AllocStrSpPtr Lib "oleaut32" Alias "SysAllocStringLen" (ByVal lStrPtr As Long, ByVal lLen As Long) As Long

' Messages from browse dialog:
Private Const BFFM_INITIALIZED = 1&       ' Indicates the browse dialog box has finished
                                          '  initializing. The lParam parameter is zero.
Private Const BFFM_SELCHANGED = 2&        ' Indicates the selection has changed. The lParam
                                          '  parameter contains the address of the item
                                          '  identifier list for the newly selected object.
Private Const BFFM_VALIDATEFAILED = 4&    ' If BIF_VALIDATE is specified in the ulFlags
                                          '  parameter of the BROWSEINFO structure then this
                                          '  message indicates that the user has entered an
                                          '  invalid path in the Edit Box.
' Messages to browse dialog:
Private Const BFFM_SETSTATUSTEXT = &H468& ' Display status text message.
Private Const BFFM_ENABLEOK = &H465&      ' Enable the OK button.
Private Const BFFM_SETSELECTION = &H467&  ' Selects the specified folder. The message's
                                          '  lParam is the pIDL of the folder to select if
                                          '  wParam is False (zero), or the path of the
                                          '  folder otherwise.
Private Const BFFM_SETOKTEXT = &H469&     ' Set the OK button text.

' www.mvps.org\vbnet
Private Const CSIDL_NETWORK As Long = &H12&

' Bobo Enterprises
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long

Private Const WS_EX_TRANSPARENT = &H20&
Private Const WS_CHILD = &H40000000
Private Const BS_CHECKBOX = &H2&

Private Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    y As Long
    x As Long
    Style As Long
    lpszName As String
    lpszClass As String
    dwExStyle As Long
End Type

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Const SW_NORMAL = &H1&

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC = (-4)

Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Const RDW_INVALIDATE = &H1&
Private Const WM_LBUTTONDOWN = &H201&
Private Const BM_GETCHECK = &HF0&
Private Const BM_SETCHECK = &HF1&

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

' Randy Birch, VBnet.com
Private Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long

Private Const ONE_KB As Long = &H400
Private Const MAX_PATH As Long = 260
Private Const ERROR_SUCCESS As Long = 0

Private mCheckboxTitle As String
Private mPrevWndProc As Long
Private hWndCheckbox As Long
Private hWndBrowse As Long

Public Enum BrowseInfoFlags
    Display_Full_System = &H0&   ' Select anything                          BIF_DISPLAYFULLSYSTEM
    Display_File_Folders = &H1&  ' Select local file system folders only    BIF_RETURNONLYFSDIRS
    Display_Domain_Only = &H2&   ' No network folders below domain level    BIF_DONTGOBELOWDOMAIN
    Include_Status_Text = &H4&   ' Include a status area in the dialog >>>> BIF_STATUSTEXT (not with BIF_NEWDIALOGSTYLE)
    File_System_Ancestors = &H8& ' Only return file system ancestors        BIF_RETURNFSANCESTORS
    Include_Edit_Box = &H10&     ' Edit Box allowing user to type the path  BIF_EDITBOX
    Valid_Result_Only = &H20&    ' Insist on valid result (or Cancel)  >>>> BIF_VALIDATE (use with BIF_EDITBOX or BIF_USENEWUI)
    New_Dialog_Style = &H40&     ' New ME/Win2000/XP style             >>>> BIF_NEWDIALOGSTYLE (use OleInitialize before)
    New_UI_With_Edit_Box = &H50& ' New dialog style with edit box           BIF_USENEWUI (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)
    Select_URLs_Also = &H80&     ' Display URLs also                        BIF_BROWSEINCLUDEURLS
    New_UI_W_Usage_Hint = &H140& ' Usage Hint with new dialog style    >>>> BIF_NEWDIALOGSTYLE Or BIF_UAHINT (adds Usage Hint if no EditBox)
    Select_Computers = &H1000&   ' Select computers, else grayed OK button  BIF_BROWSEFORCOMPUTER
    Select_Printers = &H2000&    ' Browsing for Printers                    BIF_BROWSEFORPRINTER
    Display_Files_Also = &H4000& ' Browsing for Everything                  BIF_BROWSEINCLUDEFILES
End Enum

''''''''''''''''''''''''''''''''''''
 Public Enum CheckboxState         '
    DontShow = -1&                 '
    Unchecked = 0&                 ' Browse() ShowCheckbox argument
    Checked = 1&                   ' BrowseCheckState public property
 End Enum                          '
''''''''''''''''''''''''''''''''''''
 Private eCheck As CheckboxState   '
''''''''''''''''''''''''''''''''''''
 #If False Then                    '
  Dim DontShow, Unchecked, Checked '
 #End If                           '
''''''''''''''''''''''''''''''''''''

' Retrieves the state of the browse checkbox on return
' Valid if Browse() ShowCheckbox arg is set to Show...
Public Property Get BrowseCheckState() As CheckboxState
    BrowseCheckState = eCheck
End Property

' ---------------------------------------------------------©Rd-
Public Function Browse(ByVal Me_hWnd As Long, Optional PromptMessage As String, Optional StartingPath As String, Optional ByVal flags As BrowseInfoFlags, Optional ByVal ShowCheckbox As CheckboxState = DontShow, Optional CheckboxMessage As String) As String
' -------------------------------------------------------------
    Dim tBrowseInfo As BROWSEINFO, lpIDList As Long
    Dim sDisplayName As String
    Dim sPromptMsg As String
    Dim sStartPath As String
    On Error GoTo ErrHandler

    ' Initialize OLE and COM if using the new dialog style
    If (flags And New_Dialog_Style) = New_Dialog_Style Then OleInitialize 0&
    If LenB(PromptMessage) Then
        sPromptMsg = PromptMessage
    Else
        sPromptMsg = "Please select a folder..."
    End If
    sDisplayName = FillNulls(ONE_KB)
    If LenB(CheckboxMessage) Then mCheckboxTitle = "  " & CheckboxMessage
    If PathExists(StartingPath) Then sStartPath = AddSlash(StartingPath)
    eCheck = ShowCheckbox

    With tBrowseInfo
        .hWndOwner = Me_hWnd
        .pszDisplayName = sDisplayName
        .lpszTitle = sPromptMsg
        .lParam = StrPtr(sStartPath)
        .lpfnCallback = GetFuncPointer(AddressOf BrowseCallback)
        .ulFlags = flags
    End With

    ' This next call issues the dialog
    lpIDList = SHBrowseForFolder(VarPtr(tBrowseInfo))

    ' If the user cancels the dialog lpIDList is zero
    If (lpIDList <> 0&) Then
        Dim sBuffer As String
        sBuffer = FillNulls(ONE_KB)

        If (SHGetPathFromIDList(lpIDList, StrPtr(sBuffer)) <> 0&) Then
            Browse = TrimNulls(sBuffer)
        End If

        ' Free the pIDL allocated by SHBrowseForFolder
        CoTaskMemFree lpIDList
        If eCheck <> DontShow Then
            SetWindowLong hWndCheckbox, GWL_WNDPROC, mPrevWndProc
            DestroyWindow hWndCheckbox
        End If
    End If
ErrHandler:
    DoEvents
End Function

' Function adopted from www.mvps.org/vbnet
Public Function BrowseNetwork(ByVal Me_hWnd As Long, Optional PromptMessage As String, Optional ByVal ShowCheckbox As Boolean, Optional CheckboxMessage As String) As String
  'returns only a valid network server or workstation (does not display shares)
   Dim tBrowseInfo As BROWSEINFO
   Dim pidlNet As Long, pidList As Long
   Dim sDisplayName As String
   Dim sPromptMsg As String
   On Error GoTo ErrHandler

  'obtain the pidl to the special folder 'network'
   If SHGetSpecialFolderLocation(Me_hWnd, CSIDL_NETWORK, pidlNet) = ERROR_SUCCESS Then

      If LenB(PromptMessage) Then
          sPromptMsg = PromptMessage
      Else
          sPromptMsg = "Please select a computer..."
      End If
      sDisplayName = FillNulls(ONE_KB)
      If LenB(CheckboxMessage) Then mCheckboxTitle = "  " & CheckboxMessage
      eCheck = ShowCheckbox

     'fill in the required members, limiting the Browse to the
     'network by specifying the returned pidl as pidlRoot
      With tBrowseInfo
         .hWndOwner = Me_hWnd
         .pidlRoot = pidlNet 'search network
         .pszDisplayName = sDisplayName
         .lpszTitle = sPromptMsg
         .lpfnCallback = GetFuncPointer(AddressOf BrowseCallback)
         .ulFlags = BrowseInfoFlags.Select_Computers
      End With

     'show the browse dialog
     pidList = SHBrowseForFolder(VarPtr(tBrowseInfo))

     'if the user cancels the dialog pidList is zero
     If (pidList <> 0&) Then
         'a server was selected. Although a valid pidl is returned,
         'SHGetPathFromIDList only returns paths to valid file system
         'objects, of which a networked machine is not. However, the
         'BROWSEINFO displayname member does contain the selected item,
         'which we return
          BrowseNetwork = TrimNulls(sDisplayName) ' tBrowseInfo.pszDisplayName = sDisplayName

          CoTaskMemFree pidList
          If eCheck <> DontShow Then
              SetWindowLong hWndCheckbox, GWL_WNDPROC, mPrevWndProc
              DestroyWindow hWndCheckbox
          End If
      End If  'If SHBrowseForFolder

      CoTaskMemFree pidlNet
   End If  'If SHGetSpecialFolderLocation
ErrHandler:
   DoEvents
End Function

Private Function BrowseCallback(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    ' Warning - Callback procedure, do not break or end
    BrowseCallback = 0& ' Default to return zero
    On Error GoTo ErrHandler
    Dim tRect As RECT
    Dim tCS As CREATESTRUCT
    Dim hFont As Long
    hWndBrowse = hWnd
    If (uMsg = BFFM_SELCHANGED) Then
        Dim sStatus As String
        sStatus = FillNulls(ONE_KB)
        If (SHGetPathFromIDList(lParam, StrPtr(sStatus)) <> 0&) Then
            ' Set the status area to the currently selected path
            SendMessage hWnd, BFFM_SETSTATUSTEXT, 0&, StrPtr(sStatus)
        End If
    ElseIf (uMsg = BFFM_INITIALIZED) Then
        CentreBrowsePos hWnd
        If (lpData <> 0&) Then
            ' wParam is True (1) since you are passing a string path.
            ' It would be False (0) if you were passing a pIDL
            SendMessage hWnd, BFFM_SETSELECTION, 1&, lpData
        End If
        If eCheck <> DontShow Then
            ' Thanks to Bobo Enterprises for checkbox code :)
            GetWindowRect hWnd, tRect
            hWndCheckbox = CreateWindowEx(WS_EX_TRANSPARENT, StrPtr("Button"), StrPtr(mCheckboxTitle), WS_CHILD Or BS_CHECKBOX, 25&, tRect.Bottom - tRect.Top - 60&, 400&, 25&, hWnd, 0&, App.hInstance, VarPtr(tCS))
            hFont = SendMessage(hWnd, WM_GETFONT, 0&, 0&)
            SendMessage hWndCheckbox, WM_SETFONT, hFont, 1&
            SendMessage hWndCheckbox, BM_SETCHECK, eCheck, 1&
            ShowWindow hWndCheckbox, SW_NORMAL
            mPrevWndProc = SetWindowLong(hWndCheckbox, GWL_WNDPROC, AddressOf CheckboxCallback)
            RedrawWindow hWnd, ByVal 0&, 0&, RDW_INVALIDATE
        End If
    ElseIf (uMsg = BFFM_VALIDATEFAILED) Then
        ' If BIF_VALIDATE is specified in the ulFlags parameter of
        ' the BROWSEINFO structure then this message indicates that
        ' the user has entered an invalid path in the Edit Box
        Beep
        BrowseCallback = -1& ' Keeps the dialog displayed
        ' Could ask to create a new folder? The lParam parameter is the
        ' address of a character buffer that contains the invalid name.
    End If
ErrHandler:
End Function

Private Function CheckboxCallback(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' Thanks to Bobo Enterprises for checkbox code :)
    If uMsg = WM_LBUTTONDOWN Then
        eCheck = Abs(SendMessage(hWndCheckbox, BM_GETCHECK, 0&, 0&) - 1&)
        SendMessage hWndCheckbox, BM_SETCHECK, eCheck, 1&
    Else
        CheckboxCallback = CallWindowProc(mPrevWndProc, hWnd, uMsg, wParam, lParam)
    End If
End Function

Private Sub CentreBrowsePos(ByVal hWndBrowse As Long)
    ' Centre the browse window
    On Error GoTo ErrHandler
    Dim Wndpl As WINDOWPLACEMENT
    Dim w As Long, h As Long

    ' First get target window's properties
    Wndpl.Length = LenB(Wndpl)

    If (GetWindowPlacement(hWndBrowse, Wndpl) <> 0&) Then
        ' Set the window's new position
        w = Wndpl.rcNormalPosition.Right - Wndpl.rcNormalPosition.Left
        h = Wndpl.rcNormalPosition.Bottom - Wndpl.rcNormalPosition.Top
        If eCheck <> DontShow Then h = h + 25&

        Wndpl.rcNormalPosition.Left = ((Screen.Width / Screen.TwipsPerPixelX) - w) \ 2&
        Wndpl.rcNormalPosition.Top = ((Screen.Height / Screen.TwipsPerPixelY) - h) \ 2&

        SetWindowPos hWndBrowse, HWND_TOP, Wndpl.rcNormalPosition.Left, _
                                           Wndpl.rcNormalPosition.Top, _
                                           w, h, SWP_SHOWWINDOW
    End If
ErrHandler:
End Sub

Private Function GetFuncPointer(ByVal hCallBack As Long) As Long
    GetFuncPointer = hCallBack
End Function

Public Function PathExists(sPath As String) As Boolean
    PathExists = Not (GetAttributes(StrPtr(sPath)) = -1&)
End Function

Public Function TrimNulls(StrZ As String) As String
    Dim lLen As Long
    lLen = StrLen(StrPtr(StrZ))
    TrimNulls = Trim$(LeftB$(StrZ, lLen + lLen)) 'Rd
End Function

'Faster sBuffer = String$(lLen, vbNullChar)
Public Function FillNulls(ByVal lLen As Long) As String  'Rd
    If (lLen > 0&) Then
        CopyMemByV VarPtr(FillNulls), VarPtr(AllocStrSpPtr(0&, lLen)), 4&
        ZeroMemByV StrPtr(FillNulls), lLen + lLen
    End If
End Function

Public Function AddSlash(sSpec As String) As String
    If (LenB(sSpec) = 0&) Then Exit Function
    If (Right$(sSpec, 1&) = "\") Then AddSlash = sSpec Else AddSlash = sSpec & "\"
End Function

'                                                                :›)
