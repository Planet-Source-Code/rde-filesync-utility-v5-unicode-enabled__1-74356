Attribute VB_Name = "modFileSync"
Option Explicit

Private Type SYSTEMTIME
    wYear As Integer         ' Specifies the current year.
    wMonth As Integer        ' Specifies the current month; January = 1, February = 2, and so on.
    wDayOfWeek As Integer    ' Specifies the current day of the week; Sunday = 0, Monday = 1, and so on.
    wDay As Integer          ' Specifies the current day of the month.
    wHour As Integer         ' Specifies the current hour.
    wMinute As Integer       ' Specifies the current minute.
    wSecond As Integer       ' Specifies the current second.
    wMilliseconds As Integer ' Specifies the current millisecond.
End Type

Private Declare Function FileTimeToSystemTime Lib "kernel32" _
    (lpFileTime As Currency, lpSystemTime As SYSTEMTIME) As Long

Private Declare Function SystemTimeToVariantTime Lib "oleaut32" _
    (lpSystemTime As SYSTEMTIME, vtime As Date) As Long

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" _
    (lpFileTime As Currency, lpLocalFileTime As Currency) As Long

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Enum eCompareFileTime
    ftOlder = -1&  ' First file time is less than second file time
    ftEqual = 0&   ' First file time is equal to second file time
    ftNewer = 1&   ' First file time is greater than second file time
End Enum

#If False Then
    Dim ftOlder, ftEqual, ftNewer
#End If

Private Declare Function CompareFileTime Lib "kernel32" _
  (pThisFileTime As Currency, pThanFileTime As Currency) As eCompareFileTime

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Const INVALID_FILE_ATTRIBUTES As Long = &HFFFFFFFF

Public Enum vbFileAttributes
    vbInvalidFile = INVALID_FILE_ATTRIBUTES
    vbNormal = &H0&         ' The file or directory has no other attributes set. This attribute is valid only if used alone.
    vbReadOnly = &H1&       ' The file or directory is read-only. Applications can read the file but cannot write to it or delete it. In the case of a directory, applications cannot delete it.
    vbHidden = &H2&         ' The file or directory is hidden. It is not included in an ordinary directory listing.
    vbSystem = &H4&         ' The file or directory is part of, or is used exclusively by, the operating system.
    vbVolume = &H8&         ' The name specified is used as the volume label for the current medium.
    vbDirectory = &H10&     ' The name specified is a directory or folder.
    vbArchive = &H20&       ' File has changed since last backup. Use this attribute to mark files for backup or removal.
    vbEncrypted = &H40&     ' The file or directory is encrypted. For a file, this means that all data streams are encrypted. For a directory, this means that encryption is the default for newly created files and subdirectories.
    vbNormalAttr = &H80&    ' FILE_ATTRIBUTE_NORMAL = 128
    vbTemporary = &H100&    ' The file is being used for temporary storage. File systems attempt to keep all of the data in memory for quicker access rather than flushing the data back to mass storage. A temporary file should be deleted by the application as soon as it is no longer needed.
    vbSparseFile = &H200&   ' The file is a sparse file.
    vbReparsePoint = &H400& ' The file has an associated reparse point.
    vbCompressed = &H800&   ' The file or directory is compressed. For a file, this means that all of the data in the file is compressed. For a directory, this means that compression is the default for newly created files and subdirectories.
    vbOffline = &H1000&     ' The data of the file is not immediately available. Indicates that the file data has been physically moved to offline storage.
End Enum

#If False Then
    Dim vbInvalidFile, vbNormal, vbReadOnly, vbHidden, vbSystem, vbVolume, vbDirectory, vbArchive, vbEncrypted, vbNormalAttr, vbTemporary, vbSparseFile, vbReparsePoint, vbCompressed, vbOffline
#End If

Public Declare Function GetAttributes Lib "kernel32" Alias "GetFileAttributesA" _
    (ByVal lpSpec As String) As vbFileAttributes

Public Declare Function SetAttributes Lib "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpSpec As String, ByVal dwAttributes As Long) As vbFileAttributes

Public Const DIR_SEP As String = "\"
Private OrigPointer As Long

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Sub HourGlass(Optional ByVal bOn As Boolean = True)
    If bOn Then ' 0 = vbDefault
        If Screen.MousePointer <> vbHourglass Then
            ' Save pointer and set hourglass
            OrigPointer = Screen.MousePointer
            Screen.MousePointer = vbHourglass
        End If
    Else
        ' Restore pointer
        Screen.MousePointer = OrigPointer
    End If
End Sub

'-----------------------------------------------------------
' Creates the specified path if it doesn't already exist.

' IMPORTANT: The specified path must contain a trailing
' backslash character, and an optional appended (ignored)
' folder or filename after the last backslash character.

' Returns: 2 if created, 1 if existed, 0 if error.
'-----------------------------------------------------------
Public Function CreatePath(sPath As String) As Long
    If (LenB(sPath) = 0&) Then Exit Function
    Dim sTemp As String, idx As Integer

    On Error GoTo FailedCreatePath
    If PathExists(sPath) Then
        CreatePath = 1&
    Else
        ' Set Idx to the first backslash
        idx = InStr(1&, sPath, DIR_SEP)

        Do ' Loop and make each subdir of the path separately
            idx = InStr(idx + 1&, sPath, DIR_SEP)
            If (idx <> 0&) Then
                sTemp = Left$(sPath, idx - 1&)
                ' Determine if this directory already exists
                If DirExists(sTemp) Then
                    CreatePath = 1&
                Else
                    ' We must create this directory
                    MkDir sTemp
                    CreatePath = 2&
                End If
            End If
        Loop Until idx = 0&
    End If
    Exit Function
FailedCreatePath:
    CreatePath = 0&
End Function

Public Function DirExists(sPath As String) As Boolean
    Dim Attribs As vbFileAttributes
    Attribs = GetAttributes(sPath)
    If (Attribs <> vbInvalidFile) Then
        DirExists = ((Attribs And vbDirectory) = vbDirectory)
    End If
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function CompFileTime(pThisFileTime As Currency, pThanFileTime As Currency) As eCompareFileTime
    CompFileTime = CompareFileTime(pThisFileTime, pThanFileTime)
End Function

Function PrettyTime(ByVal TimeValue As Date) As String 'Formatted like 7:56 PM
    PrettyTime = Format$(Hour(TimeValue) & ":" & Minute(TimeValue), "h:mm AM/PM")
End Function

Public Function FileTimeToString(ByVal cFileTime As Currency) As String
    On Error GoTo HandleIt
    Dim tSysTime As SYSTEMTIME
    Dim dteFile As Date
    ' Convert file time to local time zone
    FileTimeToLocalFileTime cFileTime, cFileTime
    ' Convert the file time to system time format
    FileTimeToSystemTime cFileTime, tSysTime
    ' Convert the system time to variant Date format
    SystemTimeToVariantTime tSysTime, dteFile
    ' Format the Date into string format
    FileTimeToString = Replace$(Format$(dteFile, "ddddd hh:mm AM/PM"), " 0", "  ")
HandleIt:
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

