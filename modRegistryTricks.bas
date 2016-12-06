Attribute VB_Name = "modRegistryTricks"
#Const FinalMode = 1
Option Explicit

Public stealth_stage As Long
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const ERROR_SUCCESS = 0
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_ARENA_TRASHED = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259
Private Const KEY_QUERY_VALUE = &H1
Private Const REG_OPTION_NON_VOLATILE = 0

Private Const REG_OPTION_BACKUP_RESTORE = 4 ' open for backup or restore
Private Const REG_OPTION_VOLATILE = 1 ' Key is not preserved when system is rebooted

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10

Private Const KEY_SET_VALUE = &H2
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = (KEY_READ)
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))



Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwindex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpName As String, ByRef hSubKey As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpName As String, ByVal lpName As String, ByRef cbName As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long



'Private Declare Function CopyFile Lib "kernel32" _
  Alias "CopyFileA" (ByVal lpExistingFileName As String, _
  ByVal lpNewFileName As String, ByVal bFailIfExists As Long) _
  As Long
  
'Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public StealthFilename As String
Public StealthVersion As String
Public StealthErrors As Long

Public stealthLog() As String

Public stealthIDselected As Integer

Private Function GetCurrentDate()
   ' FormatDateTime formats Date in long date.
   GetCurrentDate = FormatDateTime(Date, 1)
End Function
Public Function RemoveProgramFromInstallationList(lngGroup As Long, strRegistryPath As String, strToCompare As String, strName As String) As String
  On Error GoTo goterr
        Dim strRes As String
        Dim hSubKey As Long
        Dim rc As Long
        Dim idx As Long
        Dim lpValues() As Variant
        Dim dname As String
        Dim KeyValType As Long
        Dim lpBuffer As String
        Dim lbuflen As Long
        Dim lpValueName As String
        Dim lpcbvaluename As Long
        Dim lpcbData As Long
        Dim lResult As Long
        Dim keyname As String
        Dim toDelete As String
        Dim ltoDelete As String
        Dim lStrName As String
        Dim strComplete As String
        hSubKey = 0
        lStrName = LCase(strName)
       ' List1.Clear
         keyname = strRegistryPath
        rc = RegOpenKey(lngGroup, keyname, hSubKey)
        
         If rc <> ERROR_SUCCESS Then
          GoTo GetKeyError
         End If
          idx = 0
        While rc = ERROR_SUCCESS
            lpValueName = Space$(400)
            lpcbvaluename = 400
           rc = RegEnumKey(hSubKey, idx, lpValueName, lpcbvaluename)
           If rc = ERROR_SUCCESS Then
             lpValueName = Left$(lpValueName, lpcbvaluename)
             lpValueName = Left(lpValueName, InStr(lpValueName, Chr(0)) - 1)
             dname = strRegistryPath & lpValueName
                If RegOpenKeyEx(lngGroup, dname, 0, KEY_QUERY_VALUE, lResult) = ERROR_SUCCESS Then
                     If (RegQueryValueEx(lResult, strToCompare, 0, KeyValType, ByVal 0, lbuflen) = ERROR_SUCCESS) Then
                      lpBuffer = String(lbuflen - 1, " ")
                      If RegQueryValueEx(lResult, strToCompare, 0, KeyValType, ByVal lpBuffer, lbuflen) = ERROR_SUCCESS Then
                       toDelete = lpBuffer
                       ltoDelete = LCase(toDelete)
                      End If
                     End If
                End If
            RegCloseKey (lResult)
            If lStrName = ltoDelete Then
                strComplete = keyname & lpValueName & "\"
                strRes = DeleteKey(lngGroup, strComplete)
                If strRes = "OK" Then
                    strRes = ""
                End If
                RemoveProgramFromInstallationList = strRes
                Exit Function
            End If
            idx = idx + 1
           End If
        Wend
GetKeyError:
         RegCloseKey (hSubKey)
         RemoveProgramFromInstallationList = ""
         Exit Function
goterr:
         RemoveProgramFromInstallationList = "ERROR: at RemoveProgramFromInstallationList, code " & CStr(Err.Number) & " : " & Err.Description
End Function




 'Delete this key.
Private Function DeleteKey(ByVal section As Long, ByVal key_name _
    As String) As String
    On Error GoTo goterr
Dim pos As Integer
Dim parent_key_name As String
Dim parent_hKey As Long
Dim strRes As String

    If Right$(key_name, 1) = "\" Then key_name = _
        Left$(key_name, Len(key_name) - 1)

    ' Delete the key's subkeys.
    strRes = DeleteSubkeys(section, key_name)
    If strRes <> "OK" Then
        DeleteKey = strRes
        Exit Function
    End If

    ' Get the parent's name.
    pos = InStrRev(key_name, "\")
    If pos = 0 Then
        ' This is a top-level key.
        ' Delete it from the section.
        RegDeleteKey section, key_name
    Else
        ' This is not a top-level key.
        ' Find the parent key.
        parent_key_name = Left$(key_name, pos - 1)
        key_name = Mid$(key_name, pos + 1)

        ' Open the parent key.
        If RegOpenKeyEx(section, _
            parent_key_name, _
            0&, KEY_ALL_ACCESS, parent_hKey) <> _
                ERROR_SUCCESS _
        Then
            DeleteKey = "Error opening parent key"
            Exit Function
        Else
            ' Delete the key from its parent.
            RegDeleteKey parent_hKey, key_name

            ' Close the parent key.
            RegCloseKey parent_hKey
        End If
    End If
    DeleteKey = "OK"
    Exit Function
goterr:
    DeleteKey = "ERROR: at DeleteKey, code " & CStr(Err.Number) & " : " & Err.Description
End Function

' Delete all the key's subkeys.
Private Function DeleteSubkeys(ByVal section As Long, ByVal _
    key_name As String) As String
    On Error GoTo goterr
Dim hKey As Long
Dim subkeys As Collection
Dim subkey_num As Long
Dim Length As Long
Dim subkey_name As String

    ' Open the key.
    If RegOpenKeyEx(section, key_name, _
        0&, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS _
    Then
        DeleteSubkeys = "Error opening key '" & key_name & "'"
        Exit Function
    End If

    ' Enumerate the subkeys.
    Set subkeys = New Collection
    subkey_num = 0
    Do
        ' Enumerate subkeys until we get an error.
        Length = 256
        subkey_name = Space$(Length)
        If RegEnumKey(hKey, subkey_num, _
            subkey_name, Length) _
                <> ERROR_SUCCESS Then Exit Do
        subkey_num = subkey_num + 1

        subkey_name = Left$(subkey_name, InStr(subkey_name, _
            Chr$(0)) - 1)
        subkeys.Add subkey_name
    Loop
    
    ' Recursively delete the subkeys and their subkeys.
    For subkey_num = 1 To subkeys.Count
        ' Delete the subkey's subkeys.
        DeleteSubkeys section, key_name & "\" & _
            subkeys(subkey_num)

        ' Delete the subkey.
        RegDeleteKey hKey, subkeys(subkey_num)
    Next subkey_num

    ' Close the key.
    RegCloseKey hKey
    DeleteSubkeys = "OK"
    Exit Function
goterr:
    DeleteSubkeys = "ERROR: at DeleteSubkeys, code " & CStr(Err.Number) & " : " & Err.Description
End Function


' Get the key information for this key and
' its subkeys.
Private Function GetKeyInfo(ByVal section As Long, ByVal _
    key_name As String, ByVal indent As Integer) As String
Dim subkeys As Collection
Dim subkey_values As Collection
Dim subkey_num As Integer
Dim subkey_name As String
Dim subkey_value As String
Dim Length As Long
Dim hKey As Long
Dim txt As String
Dim subkey_txt As String

    Set subkeys = New Collection
    Set subkey_values = New Collection

    If Right$(key_name, 1) = "\" Then key_name = _
        Left$(key_name, Len(key_name) - 1)

    ' Open the key.
    If RegOpenKeyEx(section, _
        key_name, _
        0&, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS _
    Then
        GetKeyInfo = "Error opening key."
        Exit Function
    End If

    ' Enumerate the subkeys.
    subkey_num = 0
    Do
        ' Enumerate subkeys until we get an error.
        Length = 256
        subkey_name = Space$(Length)
        If RegEnumKey(hKey, subkey_num, _
            subkey_name, Length) _
                <> ERROR_SUCCESS Then Exit Do
        subkey_num = subkey_num + 1
        
        subkey_name = Left$(subkey_name, InStr(subkey_name, _
            Chr$(0)) - 1)
        subkeys.Add subkey_name
    
        ' Get the subkey's value.
        Length = 256
        subkey_value = Space$(Length)
        If RegQueryValue(hKey, subkey_name, _
            subkey_value, Length) _
            <> ERROR_SUCCESS _
        Then
            subkey_values.Add "Error"
        Else
            ' Remove the trailing null character.
            subkey_value = Left$(subkey_value, Length - 1)
            subkey_values.Add subkey_value
        End If
    Loop
    
    ' Close the key.
    If RegCloseKey(hKey) <> ERROR_SUCCESS Then
        GetKeyInfo = "Error closing key."
        Exit Function
    End If

    ' Recursively get information on the keys.
    For subkey_num = 1 To subkeys.Count
        subkey_txt = GetKeyInfo(section, key_name & "\" & _
            subkeys(subkey_num), indent + 2)
        txt = txt & Space(indent) & _
            subkeys(subkey_num) & _
            ": " & subkey_values(subkey_num) & vbCrLf & _
            subkey_txt
    Next subkey_num

    GetKeyInfo = txt
End Function


Public Function BlackdFileCopy(strFrom As String, strTo As String) As String
    On Error GoTo goterr
    Dim fs As Scripting.FileSystemObject
    Dim fol As Scripting.Folder
    Dim fil As Scripting.Folder
    Set fs = New Scripting.FileSystemObject
    If fs.FileExists(strTo) = True Then
        fs.DeleteFile strTo, True
    End If
    fs.CopyFile strFrom, strTo, True
    BlackdFileCopy = ""
    Exit Function
goterr:
        BlackdFileCopy = "ERROR: System was not able to do the copy" & vbCrLf & _
        "ERROR CODE: " & CStr(Err.Number) & vbCrLf & _
        "ERROR DESCRIPTION: " & Err.Description & vbCrLf & _
        "FROM: " & strFrom & vbCrLf & _
        "TO: " & strTo
End Function

Public Function BlackdFileExistCheck(strTo As String) As Boolean
    On Error GoTo goterr
    Dim fs As Scripting.FileSystemObject
    Dim fol As Scripting.Folder
    Dim fil As Scripting.Folder
    Set fs = New Scripting.FileSystemObject
    If fs.FileExists(strTo) = False Then
        BlackdFileExistCheck = False
    Else
         BlackdFileExistCheck = True
    End If
    Exit Function
goterr:
        MsgBox "ERROR: Filesystem critical error." & vbCrLf & _
        "ERROR CODE: " & CStr(Err.Number) & vbCrLf & _
        "ERROR DESCRIPTION: " & Err.Description, vbCritical + vbOKOnly, "ERROR"
        End

End Function

Public Function RandomFileName() As String
    Dim strType1(0 To 20) As String
    Dim strType2(0 To 4) As String
    Dim lngS As Long
    Dim i As Long
    Dim optV As Long
    Dim strName As String
    Dim strNew As String
    Dim lngNew As Long
    Randomize
    strType1(0) = "b"
    strType1(1) = "c"
    strType1(2) = "d"
    strType1(3) = "f"
    strType1(4) = "g"
    strType1(5) = "h"
    strType1(6) = "j"
    strType1(7) = "k"
    strType1(8) = "l"
    strType1(9) = "m"
    strType1(10) = "n"
    strType1(11) = "p"
    strType1(12) = "q"
    strType1(13) = "r"
    strType1(14) = "s"
    strType1(15) = "t"
    strType1(16) = "v"
    strType1(17) = "w"
    strType1(18) = "x"
    strType1(19) = "y"
    strType1(20) = "z"
    
    strType2(0) = "a"
    strType2(1) = "e"
    strType2(2) = "i"
    strType2(3) = "o"
    strType2(4) = "u"
   
    strName = ""
    lngS = randomNumberBetween(3, 9)
    optV = randomNumberBetween(0, 1)
    For i = 1 To lngS
      If optV = 0 Then
        optV = 1
        lngNew = randomNumberBetween(0, 20)
        strNew = strType1(lngNew)
      Else
        optV = 0
        lngNew = randomNumberBetween(0, 4)
        strNew = strType2(lngNew)
      End If
      strName = strName & strNew
    Next i
    strName = strName & ".exe"
    RandomFileName = strName
End Function

Public Function CopyMyselfTo(strNewName As String) As String
    On Error GoTo goterr
    Dim blnok As Boolean
    Dim strbase As String
    Dim strFrom As String
    Dim strTo As String
    strbase = App.Path
    If Right$(strbase, 1) <> "\" Then
        strbase = strbase & "\"
    End If
    strFrom = strbase & "blackdproxy.exe"
    strTo = strbase & strNewName
    CopyMyselfTo = BlackdFileCopy(strFrom, strTo)
    Exit Function
goterr:
    If Err.Number = 70 Then
        CopyMyselfTo = "WINDOWS VISTA: a part of the stealth process should be done manually! :" & vbCrLf & _
        "FOR YOUR SAFETY..." & vbCrLf & _
        "1. Close Blackd Proxy" & vbCrLf & _
        "2. Copy blackdproxy.exe to " & strNewName & vbCrLf & _
        "3. From now always execute " & strNewName & " instead blackdproxy.exe"
    Else
        CopyMyselfTo = "ERROR: System was not able to do the copy" & vbCrLf & _
        "ERROR CODE: " & CStr(Err.Number) & vbCrLf & _
        "ERROR DESCRIPTION: " & Err.Description & vbCrLf & _
        "FROM: " & strFrom & vbCrLf & _
        "TO: " & strTo
    End If
End Function
