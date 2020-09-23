Attribute VB_Name = "modIni"
Option Explicit

'The GetPrivateProfileString function is not case-sensitive
Private Declare Function ReadINIString Lib "kernel32" Alias "GetPrivateProfileStringW" _
    (ByVal lpSectName As Long, ByVal lpKeyName As Long, ByVal lpDefault As Long, _
     ByVal lpRetString As Long, ByVal nSize As Long, ByVal lpFileName As Long) As Long

Private Declare Function WriteINIString Lib "kernel32" Alias "WritePrivateProfileStringW" _
    (ByVal lpSectName As Long, ByVal lpKeyName As Long, ByVal lpString As Long, _
     ByVal lpFileName As Long) As Long

Private Declare Function ReadINISection Lib "kernel32" Alias "GetPrivateProfileSectionW" _
    (ByVal lpSectName As Long, ByVal lpRetString As Long, ByVal nSize As Long, _
     ByVal lpFileName As Long) As Long

Private Const lCHUNK As Long = 256

'fSuccess = SetINIKey("Vbaddin.ini", "Add-Ins32", "MyAddin.Connect", 3)
Public Function SetINIKey(sFile As String, sSection As String, _
                          sKey As String, sValue As String) As Boolean
    SetINIKey = WriteINIString(StrPtr(sSection), StrPtr(sKey), StrPtr(sValue), StrPtr(sFile))
End Function

'sValue = GetINIKey("Vbaddin.ini", "Add-INS32", "MyAddIn.CoNNect")
Public Function GetINIKey(sFile As String, sSection As String, _
                          sKey As String, Optional sDefault As String) As String
    Dim lLen As Long, lBuf As Long, sVal As String
    Do: lBuf = lBuf + lCHUNK
        sVal = String$(lBuf, vbNullChar)
        lLen = ReadINIString(StrPtr(sSection), StrPtr(sKey), StrPtr(sDefault), StrPtr(sVal), lBuf, StrPtr(sFile))
    Loop While (lLen = lBuf - 1&)
    GetINIKey = LeftB$(sVal, lLen + lLen)
End Function

'fSuccess = DelINIKey("Settings.ini", "Settings", "Server Path")
Public Function DelINIKey(sFile As String, sSection As String, _
                          sKey As String) As Boolean
    DelINIKey = WriteINIString(StrPtr(sSection), StrPtr(sKey), 0&, StrPtr(sFile))
End Function

'fSuccess = DelINISection("Settings.ini", "Settings")
Public Function DelINISection(sFile As String, sSection As String) As Boolean
    DelINISection = WriteINIString(StrPtr(sSection), 0&, 0&, StrPtr(sFile))
End Function

'sKeysValues = GetINISection("Settings.ini", "Settings")
Public Function GetINISection(sFile As String, sSection As String) As String
  ' The keys and values for the specified section are null-delimited,
  ' and terminated by a final null character.
  ' Returns: key 1=value 1{NULL}key 2=value 2{NULL}key 3=value 3{NULL}
    Dim lLen As Long, lBuf As Long, sBuff As String
    Do: lBuf = lBuf + lCHUNK
        sBuff = String$(lBuf, vbNullChar)
        lLen = ReadINISection(StrPtr(sSection), StrPtr(sBuff), lBuf, StrPtr(sFile))
    Loop While (lLen = lBuf - 2&)
    GetINISection = LeftB$(sBuff, lLen + lLen)
End Function                                                   ' Rd :)
