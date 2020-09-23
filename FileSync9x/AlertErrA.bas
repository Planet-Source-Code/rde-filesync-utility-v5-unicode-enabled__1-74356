Attribute VB_Name = "AlertErr"
' ____________________________________________________________________
'  AlertError module - alerting errors when they occur.          -©Rd-
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' This module determines if this process is running within an instance
' of the VB Development Environment, or within a stand-alone executable.
'
' The error philosophy is simple:
'          - handle errors conveniently during development.
'          - log errors to a file when running as an executable.
'
' It is recommended to NOT SUPPRESS ERRORS, but to deal with errors
' within the procedure where the error occurs, helping debugging and
' assertion to happen THERE! Don't suppress, Validate! But don't get
' me wrong - every procedure should have an error handler:
'
'    If InAnExe Then On Error GoTo ErrHandler
'    ...
' ErrHandler:
'    If Err Then AlertError sProc
'
' The single advantage over conventional error raising is the automatic
' disabling of exception raising/unexpected errors when the program is
' in your end-users space.
'
' This module is also very handy for logging messages using AlertMsg,
' (automatically to the most convenient location), and so allowing for
' un-interupted run testing while recording significant events to a
' log file or the debug window.
'
' This module can *best pick* the log path for all running environments
' including when running as a compiled ActiveX component in/out of IDE.
'
' Environment enumeration thanks to Ulli.
' ___________________________________________________________________
'
' InitError
' ¯¯¯¯¯¯¯¯¯
' Optionally call InitError within an initialization event.
' Otherwise, it will be called on the first access to properties
' or procedures in the module.
'
' InitError assigns to these public read-only properties:
'
'     hWndVBE - Set to the VB IDE window handle (hWnd), or zero.
'     InVBIde - Set to True if running in the VB IDE, or False
'               if running as an EXE.
'     InAnExe - Set to True if running as an EXE, or False if
'               running in the VB IDE.
'     ExeSpec - Specifies the path and filename of the executable.
'     LogPath - Specifies the default log path used for logging.
' Environment - Enumeration identifying the running environment.
'
' Only the LogPath Property is read/write and so can be assigned
' at any time. You can also pass an empty string to reset it to
' the default path - for more on the default log path see below.
' __________________________________________________________________
'
' StackAdd and StackRemove
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' Call the StackAdd sub at the top of each procedure that utilizes
' the AlertErr module. Then call StackRemove before procedure exit.
'
' The call stack will then be included in all error logging details.
' __________________________________________________________________
'
' AlertError
' ¯¯¯¯¯¯¯¯¯¯
' Input:
'   ProcName  - A String description to identify the module
'               and routine where the error occured.
'   AlertMode - Specifies the error mode when in the IDE.
'   ExtraInfo - This optional argument can be used to alert
'               you to extra information about the error,
'               such as argument and variable values, etc.
' Output:
'   The AlertError sub-routine outputs one of the following:
'
'   If it is running in the VB IDE
' MessageBox  - Displays a MsgBox with error description.
' LogToFile   - Beeps and appends to log file in the log path.
' DebugPrint  - Beeps and prints desc to debug window (default).
' BeepOnly    - Beeps only.
' Custom      - Situational. Specially formatted message boxes?
'
'   If in an executable
' LogToFile   - Beeps and appends to log file in the log path.
'
' AlertMsg
' ¯¯¯¯¯¯¯¯
' The AlertMsg sub can be used to alert you to events of interest
' without interrupting execution, and is handy for tracking event
' sequence. AlertMsg uses the same path as the error log when
' writing to a file, but also takes an optional path parameter.
' ______________________________________________________________
'
' Example code
' ¯¯¯¯¯¯¯¯¯¯¯¯
' Call InitError:
' ___________
'  InitError
' ¯¯¯¯¯¯¯¯¯¯¯
' You can use ...
' __________________________________________
'  If InAnExe Then On Error GoTo ErrHandler
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' This allows you to proof your code as you develop, adding
' assertions and error handling as needed. Add assertions
' before the above conditional. Then add error handlers to
' deal with possible 'expected' errors, maybe using Resume.
'
' The following code creates a special case as needed.
' A CommonDialog is a good example.
' __________________________________________
'  On Error GoTo ErrHandler
'  ' code that could cause 'expected' error
'  If InAnExe Then On Error GoTo ErrHandler
'  ' more code that could raise errors
' ErrHandler:
'  If Err.Number = 'expected' Then
'      ' error is handled
'      Resume ' or Resume Next
'  ElseIf Err Then
'      AlertError sProc
'  End If
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' This allows you to identify and handle possible scenarios
' ('expected' errors), but raises 'unexpected' errors right
' where they occur. If an 'unexpected' error occurs when the
' project is compiled the error will be logged.
'
' In some cases you need to remove 'If InAnExe Then' and revert
' to a less immediate solution. A Callback comes to mind:
' __________________________________________
'  On Error GoTo ErrHandler
'  ' call-back code
' ErrHandler:
'  sProc = Me.Name & ".CallbackFunc"
'  If Err Then
'      AlertError sProc, DebugPrint, "lParam = " & lParam
'  End If
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' _______________________________________________________________
'
' Details
' ¯¯¯¯¯¯¯
' The hWndVBE property is set to the VB IDE window handle, which
' is handy when more than one instance of the IDE is running.
'
' When running in a compiled executable the ExeSpec property
' specifies the path and file name of the parent executable, but
' when in the IDE it contains the path and name of the VB exe.
'
' By default, the log file is written to the path obtained from
' ExeSpec only when in an exe, and the App.Path property is used
' in the IDE.
'
' If you are using this module in a compiled ActiveX component
' running in another client project *as a compiled executable*
' the ExeSpec property will identify the app path of the client,
' and the App.Path property will specify the location of your
' component. In this case the ExeSpec path is used by default.
' This proves much more useful when debugging the client exe.
'
' If you are using this module in an ActiveX component running
' in another client project *in the IDE* the ExeSpec property
' will specify the path of the VB exe, and the App.Path property
' will specify the location of your component (not the App.Path
' property of the client project). In this case your components
' App.Path property is used by default.
'
' You can over-ride the default log path by passing an optional
' path parameter to InitError to be used as the log path. If
' you do specify the log path it will be used in all running
' environments, not just in the IDE (can be reset, see below).
'
' Remember, according to this philosophy, your component could
' still use 'If InAnExe Then On Error GoTo ErrHandler' which will
' raise errors to the client during their development process for
' invalid arguments and other assertions (data types, ranges, etc).
' Using assertions and raising errors in the IDE = easy debugging!
'
' The log file is named App.EXEName & "_Error.log" for error
' logging, and App.EXEName & "_Msg.log" for AlertMsg. You can
' optionally specify the name (without extension) to be used in
' place of App.EXEName when calling InitError.
'
' InitError can be re-called to specify another log path and/or
' file name prefix without re-testing the running environment,
' and omitting either parameter will reset the default log path
' or name according to the logic as described above.
'
' Also, the LogPath property can be assigned another log path at
' any time. You can also assign an empty string to the LogPath
' property to reset it to the default log path. Clear as mud?
' ________________________________________________________________
'
' Compiled ActiveX Components:
'
' If you're using this module in a compiled ActiveX component then
' printing to the Debug window is not available, so AlertError and
' AlertMsg will default to logging to file.
'
' Also, a message box may be inappropriate when running in a client
' project even in the IDE, so the following flag can be used to log
' all MessageBox errors to file once your component is compiled:
'
' Components raise msgbox errors during dev only
Private Const COMPONENT_NO_MSGBOX_TO_CLIENT As Boolean = False
'
' Final note - This flag is relevant only when using:
'
' On Error GoTo ErrHandler
'
' When using:
'
' If InAnExe Then On Error GoTo ErrHandler
'
' all component errors will be raised to the client when in
' the IDE, with no interception by this module.
' ______________________________________________________________

Public Enum eAlertMode
    DebugPrint
    LogToFile
    MessageBox
    BeepOnly
    Custom
End Enum

Public Enum eEnvironment
    EnvironIDE = 1         ' Project in the IDE
    EnvironCompiled = 2    ' Compiled executable
    EnvironCompiledIDE = 3 ' Compiled component in IDE
End Enum

' Now includes the full GetClientSpec module code and supports vb5/6

Private Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpEnumFunc As Long, ByRef lParam As Long) As Long
Private Declare Function GetWindowClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nBufLen As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long ' ©Rd
Private Declare Function GetModuleHandleZ Lib "kernel32" Alias "GetModuleHandleA" (ByVal hNull As Long) As Long
Private Declare Function GetAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpSpec As String) As Long

'Randy Birch, VBnet.com
Private Declare Function StrLenW Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long

Private Const INVALID_FILE_ATTRIBUTES = (-1)
Private Const GWL_HINSTANCE = (-6)
Private Const MAX_PATH As Long = 260

Private saStack() As String
Private icStack As Long

Private maVBIDEs() As Long
Private mhWndVBE As Long
Private mInVBIDE As Boolean
Private mInAnExe As Boolean
Private mExeSpec As String

Private mLogPath As String
Private mExePath As String
Private mEXEName As String
Private mEnviron As eEnvironment
Private mfInit As Boolean

Option Explicit

' ___________________________________________________________
' PUBLIC SUB: InitError - First property access calls here
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' Assigns globals: VBIDE hWnd or zero, run mode props set.
'
' You can over-ride the default log path by passing the optional
' path parameter to be used as the log path. If you do specify the
' log path it will be used in all running environments, not just
' in the IDE, but can be reset on the run at any time.
'
' The log file is named App.EXEName & "_Error.log" for error
' logging, and App.EXEName & "_Msg.log" for AlertMsg. You can
' optionally specify a name to be used in place of App.EXEName.
'
' InitError can be re-called to specify another log path without
' re-testing the running environment, and omitting the log path
' parameter will reset the log path to best pick for environment.
'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Public Sub InitError(Optional sLogPath As String, Optional sLogFileName As String)
Attribute InitError.VB_Description = "Initialization sub; sets public read-only properties hWndVBE, InVBIde, InAnExe, ExeSpec and LogPath."
   On Error GoTo Fail
   If (mfInit = False) Then
      mfInit = True 'Set Props first time
      mhWndVBE = GetVBIdeHandle 'VBE instance
      mExeSpec = GetClientSpec  'Full spec and exe
      mExePath = RTrimChr(mExeSpec) 'Full path to exe

      If (mhWndVBE = 0) Then
         mEnviron = EnvironCompiled
         mInAnExe = True 'In An Exe

      ElseIf (App.StartMode = vbSModeAutomation) Then
         mEnviron = EnvironCompiledIDE
         mInVBIDE = True 'In Component in IDE

      Else
         mEnviron = EnvironIDE
         mInVBIDE = True 'In IDE
      End If
   End If

   If LenB(sLogFileName) = 0 Then
      mEXEName = App.EXEName
   Else
      mEXEName = RTrimChr(sLogFileName, ".")
   End If

   If FolderExists(sLogPath) Then
      ' Remove trailing backslash if present
      If (Right$(sLogPath, 1) = "\") Then
         mLogPath = RTrimChr(sLogPath)
      Else
         mLogPath = sLogPath
      End If
   Else
      If (mInAnExe) Then
         mLogPath = mExePath
      Else
         mLogPath = App.Path
      End If
   End If
Fail:
End Sub

' ___________________________________________________________
' PUBLIC SUB: StackAdd
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' Call the StackAdd sub at the top of each procedure that
' utilizes the AlertErr module.
'
' The call stack will be included in all error logging.
'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Public Sub StackAdd(Module_ProcName As String)
    ReDim Preserve saStack(icStack) As String
    saStack(icStack) = Module_ProcName
    icStack = icStack + 1&
End Sub

' ___________________________________________________________
' PUBLIC SUB: StackRemove
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' Each procedure that calls the StackAdd sub must also call
' StackRemove before procedure exit.
'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Public Sub StackRemove()
    If icStack Then icStack = icStack - 1&
End Sub

' ___________________________________________________________
' PUBLIC SUB: AlertError - Logs automatically when in an Exe
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' The ProcName argument can be used to name the module and
' procedure where the error occured.
'
' When in the IDE this sub handles errors according to the
' AlertMode argument, which defaults to DebugPrint if omitted.
' If in an executable it automatically defaults to logging.
'
' The optional ExtraInfo argument can be used to alert you
' to pertinent information about the error, such as argument
' and variable values, and other state data.
'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Public Sub AlertError(ProcName As String, Optional ByVal AlertMode As eAlertMode = DebugPrint, Optional ExtraInfo As String)
Attribute AlertError.VB_Description = "When in the IDE handles errors according to the optional AlertMode argument, which defaults to DebugPrint if omitted. If in an executable it automatically defaults to logging."
    Dim Num As Long, Src As String, Desc As String

    Num = Err.Number
    Src = Err.Source
    If (Erl = 0) Then
        Desc = Err.Description
    Else
        Desc = "Error on line " & Erl & vbCrLf & Err.Description
    End If

    If LenB(ExtraInfo) Then Desc = Desc & vbCrLf & ExtraInfo
    If icStack Then Desc = Desc & StackRead

    On Error GoTo Fail
    If mfInit Then Else InitError

    If (mInAnExe) Then
        AlertMode = LogToFile
    ElseIf (mEnviron = EnvironCompiledIDE) Then
        ' If a compiled ActiveX component in another vb project
        ' then substitute debug.print with log to file.
        If AlertMode = DebugPrint Then
            AlertMode = LogToFile
        ElseIf AlertMode = MessageBox Then
            ' If raise msgbox errors during dev only
            If COMPONENT_NO_MSGBOX_TO_CLIENT Then
                AlertMode = LogToFile
        End If: End If
    End If

    Select Case AlertMode
            Case MessageBox
                MsgBox ProcName & " error!" & vbCr & vbCr & _
                       "Error #" & Num & " - " & Desc, _
                       vbExclamation, "Error #" & Num

            Case DebugPrint
                Debug.Print " ------- "; Format$(Now, "h:nn:ss"); " -------"
                Debug.Print ProcName; " error!"
                Debug.Print "Error #"; Num; " - "; Desc
                Debug.Print "                       * * * * * ERROR * * * * *"
                                         Beep
            Case LogToFile
                Dim i As Integer: i = FreeFile()
                Open mLogPath & "\" & mEXEName & "_Error.log" For Append Shared As #i
                    Print #i, Src; " error log ";
                    Print #i, Format$(Now, "h:nn:ss am/pm mmmm d, yyyy")
                    Print #i, ProcName; " error!"
                    Print #i, "Error #"; Num; " - "; Desc
                    Print #i, " * * * * * * * * * * * * * * * * * * *"
                Close #i
                                            Beep
            Case BeepOnly
                ' Beep me only
                                              Beep
            Case Else
                ' Do nothing. Specially formatted messages?
    End Select
Fail:
End Sub

' ___________________________________________________________
' PUBLIC SUB: AlertMsg - Logs automatically when in an Exe
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' This sub can be used to alert you to pertinent information
' about the the running app without interrupting execution.
'
' When in the IDE this sub handles messages according to the
' AlertMode argument, which defaults to DebugPrint if omitted.
' If in an executable it automatically writes to a log file.
'
' The log file path can be over-ridden by the optional path
' parameter.
'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Public Sub AlertMsg(Msg As String, Optional ByVal AlertMode As eAlertMode = DebugPrint, Optional sLogPath As String, Optional ByVal bBeep As Boolean, Optional ByVal bVerbose As Boolean)
Attribute AlertMsg.VB_Description = "When in the IDE handles messages according to the optional AlertMode argument, which defaults to DebugPrint if omitted. If in an executable it automatically defaults to logging."
    On Error GoTo Fail
    If mfInit Then Else InitError
    If (mInAnExe) Then
        AlertMode = LogToFile
    ElseIf (mEnviron = EnvironCompiledIDE) Then
        ' If a compiled ActiveX component in another vb project
        ' then substitute debug.print with log to file.
        If AlertMode = DebugPrint Then
            AlertMode = LogToFile
        ElseIf AlertMode = MessageBox Then
            ' If raise msgbox errors during dev only
            If COMPONENT_NO_MSGBOX_TO_CLIENT Then
                AlertMode = LogToFile
        End If: End If
    End If
    Dim sFile As String
    If FolderExists(sLogPath) Then
        ' Add trailing backslash if missing
        If (Right$(sLogPath, 1) = "\") Then
            sFile = sLogPath & mEXEName
        Else
            sFile = sLogPath & "\" & mEXEName
        End If
    Else
        sFile = mLogPath & "\" & mEXEName
    End If
    Select Case AlertMode
            Case MessageBox
                MsgBox Msg, vbInformation, " Message..."

            Case DebugPrint
                If bVerbose Then Debug.Print " ------- "; Format$(Now, "h:nn:ss"); " -------"
                Debug.Print Msg
                If bVerbose Then Debug.Print "                       * * * * * MSG * * * * *"
                             If bBeep Then Beep
            Case LogToFile
                Dim i As Integer: i = FreeFile()
                Open sFile & "_Msg.log" For Append Shared As #i
                    If bVerbose Then Print #i, Format$(Now, "h:nn:ss am/pm mmmm d, yyyy")
                    Print #i, Msg
                    If bVerbose Then Print #i, " * * * * * * * * * * * * * * * * * * *"
                Close #i
                             If bBeep Then Beep
            Case BeepOnly
                             If bBeep Then Beep
            Case Else
                ' Do nothing. Specially formatted messages?
    End Select
Fail:
End Sub

' ___________________________________________________________
' PUBLIC PROPERTY: Environment
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' Property to easily identify the running environment.
'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Public Property Get Environment() As eEnvironment
    If mfInit Then Else InitError
    Environment = mEnviron
End Property

' ___________________________________________________________
' PUBLIC PROPERTY: hWndVBE
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' Set to the VB IDE window handle (hWnd), or zero.
'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Property Get hWndVBE() As Long
Attribute hWndVBE.VB_Description = "Set to the VB IDE window handle (hWnd), or zero if running as an executable."
   If mfInit Then Else InitError
   hWndVBE = mhWndVBE
End Property

' ___________________________________________________________
' PUBLIC PROPERTY: InVBIde
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' Set to True if running in the VB IDE, or False
' if running as an EXE.
'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Property Get InVBIde() As Boolean
Attribute InVBIde.VB_Description = "Set to True if running in the VB IDE, or False if running as an EXE."
    If mfInit Then Else InitError
    InVBIde = mInVBIDE
End Property

' ___________________________________________________________
' PUBLIC PROPERTY: InAnExe
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' Set to True if running as an EXE, or False if
' running in the VB IDE.
'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Property Get InAnExe() As Boolean
Attribute InAnExe.VB_Description = "Set to True if running as an EXE, or False if running in the VB IDE."
    If mfInit Then Else InitError
    InAnExe = mInAnExe
End Property

' ___________________________________________________________
' PUBLIC PROPERTY: ExeSpec
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' Specifies the path and filename of the executable.
'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Property Get ExeSpec() As String
Attribute ExeSpec.VB_Description = "Specifies the path and filename of the executable."
    If mfInit Then Else InitError
    ExeSpec = mExeSpec
End Property

' ___________________________________________________________
' PUBLIC PROPERTY: LogPath
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' Specifies the default log path used for logging.
'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Property Get LogPath() As String
Attribute LogPath.VB_Description = "Specifies the default log path used for logging."
    If mfInit Then Else InitError
    LogPath = mLogPath
End Property

Property Let LogPath(sLogPath As String)
   If FolderExists(sLogPath) Then
      ' Remove trailing backslash if present
      If (Right$(sLogPath, 1) = "\") Then
         mLogPath = RTrimChr(sLogPath)
      Else
         mLogPath = sLogPath
      End If
   Else
      If (mInAnExe) Then
         mLogPath = mExePath
      Else
         mLogPath = App.Path
      End If
   End If
End Property

' ___________________________________________________________
' PRIVATE FUNCTION: GetVBIdeHandle                      -©Rd-
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' If running within an instance of the VB IDE GetVBIdeHandle
' returns the window handle (hWnd) of the Main VB window.
'
' If running as a stand-alone executable the GetVBIdeHandle
' function returns zero.
'
' Returns: VB's window handle (hWnd), zero otherwise.
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function GetVBIdeHandle() As Long
    On Error GoTo ErrHandler
    Dim rc As Long, nVBIDEs As Long

    ' Search all current thread windows for the VB IDE main window
    rc = EnumThreadWindows(GetCurrentThreadId, AddressOf CallBackIDE, nVBIDEs)

    ' If the IDE is running
    If (nVBIDEs) Then
        Dim VBProcessID As Long, MeProcessID As Long, i As Long

        ' Get this components's Process ID
        MeProcessID = GetCurrentProcessId

        For i = 1 To nVBIDEs
            ' Get VB's Process ID
            rc = GetWindowThreadProcessId(maVBIDEs(i), VBProcessID)

            ' If running in the same process
            If (VBProcessID = MeProcessID) Then
                GetVBIdeHandle = maVBIDEs(i) ' ©Rd
                Exit Function
            End If
        Next i
    End If
ErrHandler:
End Function

' ___________________________________________________________
' PRIVATE FUNCTION: CallBackIDE                         -©Rd-
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' This is a support function for the GetVBIdeHandle function.
'
' This is a Call-Back function called by the EnumThreadWindows
' API function (used in GetVBIdeHandle above).
'
' It receives the handle of each window, and if the handle is
' the Main VB IDE window it is added to the maVBIDEs array.
'
' Assigns ByRef: The lCount parameter indicating the number
'          of VB IDE's currently running, zero otherwise.
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function CallBackIDE(ByVal hWnd As Long, ByRef lCount As Long) As Long
    On Error GoTo ErrHandler
    ' Default to Enum the next window
    CallBackIDE = 1
    ' If it's a VB IDE instance
    If (GetClassName(hWnd) = "IDEOwner") Then
        lCount = lCount + 1
        ReDim Preserve maVBIDEs(1 To lCount) As Long
        ' Record the window handle
        maVBIDEs(lCount) = hWnd
    End If
    Exit Function
ErrHandler:
    ' On error cancel callback
    CallBackIDE = 0
End Function

' ___________________________________________________________
' PRIVATE FUNCTION: StackRead
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' This is a support function that returns the call stack.
'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function StackRead() As String
    Dim i As Long
    Do While icStack - i
       i = i + 1&
       StackRead = StackRead & vbCrLf & "Caller: " & saStack(icStack - i)
    Loop
End Function

' ___________________________________________________________
' PRIVATE FUNCTION: GetClassName                        -©Rd-
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' This is a support function for the CallBackIDE function.
'
' This function returns the class name of the window whose
' handle is passed as the hWnd argument.
'
' Returns: The window's class name.
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function GetClassName(ByVal hWnd As Long) As String
    On Error GoTo ErrHandler
    GetClassName = "unknown"
    Dim ClassName As String, BufLength As Long
    ' Allow ample length for the class name
    BufLength = MAX_PATH
    ClassName = String$(BufLength, vbNullChar)

    If (GetWindowClassName(hWnd, ClassName, BufLength)) Then
        GetClassName = TrimNull(ClassName)
    End If
ErrHandler:
End Function

' ___________________________________________________________
' PUBLIC FUNCTION: GetClientSpec                        -©Rd-
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' This is a support function for the InitError sub-routine.
'
' This function returns the path and name of the file used
' to create the calling process.
'
' Returns: A fully-qualified path and name.
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Public Function GetClientSpec() As String
    On Error GoTo ErrHandler
    Dim sModName As String, hInst As Long, rc As Long
    ' Get the application hInstance. By passing NULL, GetModuleHandle
    ' returns a handle to the file used to create the calling process.
    hInst = GetWindowLong(GetModuleHandleZ(0&), GWL_HINSTANCE)
    ' Get the module file name
    sModName = String$(MAX_PATH, vbNullChar)
    rc = GetModuleFileName(hInst, sModName, MAX_PATH)
    GetClientSpec = TrimNull(sModName)
ErrHandler:
    ' Return empty string on error
End Function

' ___________________________________________________________
' PUBLIC FUNCTION: RTrimChr
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' This function removes from sStr the first occurrence from
' the right of the specified character(s) and everything
' following it, and returns just the path up to but not
' including the specified character(s).
'
' It always searches from right to left starting at the end
' of sStr. If the character(s) does not exist in sStr then
' the whole of sStr is returned and lRetPos is set to
' Len(sStr) + 1.
'
' If sChar is omitted it defaults to a backslash.
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Public Function RTrimChr(sStr As String, Optional sChar As String = "\", Optional ByRef lRetPos As Long, _
                                  Optional ByVal eCompare As VbCompareMethod = vbBinaryCompare) As String
    On Error GoTo ErrHandler
    Dim lPos As Long
    ' Default to return the passed string
    lRetPos = Len(sStr) + 1&
    If LenB(sChar) Then
        lPos = InStr(1&, sStr, sChar, eCompare)
        Do Until lPos = 0&
            lRetPos = lPos
            lPos = InStr(lRetPos + 1&, sStr, sChar, eCompare)
        Loop
    End If
    ' Return sStr w/o sChar and any following substring
    RTrimChr = LeftB$(sStr, lRetPos + lRetPos - 2&)
ErrHandler:
End Function

' ___________________________________________________________
' PUBLIC FUNCTION: TrimNull
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' This function extracts the string from the null terminated
' string passed to it.
'
' Returns: The string of characters up to the first null
'          (ASCII 0) character.
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Public Function TrimNull(StrZ As String) As String
    Dim lLen As Long
    lLen = StrLenW(StrPtr(StrZ))
    TrimNull = LeftB$(StrZ, lLen + lLen) 'Rd
End Function

' ___________________________________________________________
' PUBLIC FUNCTION: FolderExists
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' This function tests the specified path to see if it
' is an existing directory.
'
' Returns: True if the specified path is a valid directory,
'          or False otherwise.
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Public Function FolderExists(sPath As String) As Boolean
    Dim Attribs As Long: Attribs = GetAttributes(sPath)
    If Not (Attribs = INVALID_FILE_ATTRIBUTES) Then
        FolderExists = ((Attribs And vbDirectory) = vbDirectory)
    End If
End Function
'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯    :›)
