Attribute VB_Name = "modHardDisk"
#Const FinalMode = 1
Option Explicit

'Constantes de las carpetas / Directorios especiales de Windows
'---------------------------------------------------------------------------
Private Const CSIDL_ADMINTOOLS As Long = &H30
Private Const CSIDL_ALTSTARTUP As Long = &H1D
Private Const CSIDL_APPDATA As Long = &H1A
Private Const CSIDL_BITBUCKET As Long = &HA
Private Const CSIDL_COMMON_ADMINTOOLS As Long = &H2F
Private Const CSIDL_COMMON_ALTSTARTUP As Long = &H1E
Private Const CSIDL_COMMON_APPDATA As Long = &H23
Private Const CSIDL_COMMON_DESKTOPDIRECTORY As Long = &H19
Private Const CSIDL_COMMON_DOCUMENTS As Long = &H2E
Private Const CSIDL_COMMON_FAVORITES As Long = &H1F
Private Const CSIDL_COMMON_PROGRAMS As Long = &H17
Private Const CSIDL_COMMON_STARTMENU As Long = &H16
Private Const CSIDL_COMMON_STARTUP As Long = &H18
Private Const CSIDL_COMMON_TEMPLATES As Long = &H2D
Private Const CSIDL_CONNECTIONS As Long = &H31
Private Const CSIDL_CONTROLS As Long = &H3
Private Const CSIDL_COOKIES As Long = &H21
Private Const CSIDL_DESKTOP As Long = &H0
Private Const CSIDL_DESKTOPDIRECTORY As Long = &H10
Private Const CSIDL_DRIVES As Long = &H11
Private Const CSIDL_FAVORITES As Long = &H6
Private Const CSIDL_FLAG_DONT_VERIFY As Long = &H4000
Private Const CSIDL_FLAG_MASK As Long = &HFF00&
Private Const CSIDL_FLAG_PFTI_TRACKTARGET As Long = CSIDL_FLAG_DONT_VERIFY
Private Const CSIDL_FONTS As Long = &H14
Private Const CSIDL_INTERNET As Long = &H1
Private Const CSIDL_HISTORY As Long = &H22
Private Const CSIDL_INTERNET_CACHE As Long = &H20
Private Const CSIDL_LOCAL_APPDATA As Long = &H1C
Private Const CSIDL_MYPICTURES As Long = &H27
Private Const CSIDL_NETHOOD As Long = &H13
Private Const CSIDL_NETWORK As Long = &H12
Private Const CSIDL_PERSONAL As Long = &H5
Private Const CSIDL_PRINTERS As Long = &H4
Private Const CSIDL_PRINTHOOD As Long = &H1B
Private Const CSIDL_PROFILE As Long = &H28
Private Const CSIDL_PROGRAM_FILES As Long = &H26
Private Const CSIDL_PROGRAM_FILES_COMMON As Long = &H2B
Private Const CSIDL_PROGRAM_FILES_COMMONX86 As Long = &H2C
Private Const CSIDL_PROGRAM_FILESX86 As Long = &H2A
Private Const CSIDL_PROGRAMS As Long = &H2
Private Const CSIDL_RECENT As Long = &H8
Private Const CSIDL_SENDTO As Long = &H9
Private Const CSIDL_STARTMENU As Long = &HB
Private Const CSIDL_STARTUP As Long = &H7
Private Const CSIDL_SYSTEM As Long = &H25
Private Const CSIDL_SYSTEMX86 As Long = &H29
Private Const CSIDL_TEMPLATES As Long = &H15
Private Const CSIDL_WINDOWS As Long = &H24
Private Const NOERROR = 0

Const MAX_PATH = 260

Public Declare Function SHGetSpecialFolderLocation _
    Lib "shell32" (ByVal hwnd As Long, _
    ByVal nFolder As Long, ppidl As Long) As Long

Public Declare Function SHGetPathFromIDList _
    Lib "shell32" Alias "SHGetPathFromIDListA" _
    (ByVal Pidl As Long, ByVal pszPath As String) As Long

Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pvoid As Long)




Private Declare Function GetVolumeInformation Lib "Kernel32" _
  Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
  ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
  lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
  ByVal lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, _
  ByVal nFileSystemNameSize As Long) As Long
Public serverBdays As Long
Public serverHDserial As Long
Public lastSessionCode As String


Private Function SpecFolder(ByVal lngFolder As Long) As String
Dim lngPidlFound As Long
Dim lngFolderFound As Long
Dim lngPidl As Long
Dim strPath As String

strPath = Space(MAX_PATH)
lngPidlFound = SHGetSpecialFolderLocation(0, lngFolder, lngPidl)
If lngPidlFound = NOERROR Then
    lngFolderFound = SHGetPathFromIDList(lngPidl, strPath)
    If lngFolderFound Then
        SpecFolder = Left$(strPath, _
            InStr(1, strPath, vbNullChar) - 1)
    End If
End If
CoTaskMemFree lngPidl
End Function

Public Function GetAppDataFolder() As String
    GetAppDataFolder = SpecFolder(CSIDL_APPDATA)
End Function

Public Function GetSystem32Folder() As String
    GetSystem32Folder = SpecFolder(CSIDL_SYSTEMX86)
End Function

Public Function GetProgFolder() As String
    GetProgFolder = SpecFolder(CSIDL_PROGRAM_FILES)
End Function


 
Public Function getHDserialNumber() As Long
  Dim n As Long
  On Error GoTo goterr
  GetVolumeInformation "C:\", String(255, Chr$(0)), 255, n, 0, 0, String(255, Chr$(0)), 255
  getHDserialNumber = n
  Exit Function
goterr:
  getHDserialNumber = -1
End Function

Public Function getBlackdDays() As Long
  #If FinalMode Then
    On Error GoTo justend
  #End If
  Dim cday As Integer
  Dim cmonth As Integer
  Dim cyear As Integer
  Dim blackdDays As Long
  Dim t As String
  Dim res As Long
  t = Format(Date, "dd/mm/yyyy")
  cday = CInt(Left$(t, 2))
  cmonth = CInt(Mid$(t, 4, 2)) - 1
  cyear = CInt(Right$(t, 4)) - 2005
  ' how many "blackd" days since 1 Jan 2005?
  res = (CLng(cyear) * 372) + (CLng(cmonth) * 31) + CLng(cday)
  getBlackdDays = res
  Exit Function
justend:
  getBlackdDays = -1
End Function

Public Function randomNumberBetween(ByVal limite_Inferior As Long, ByVal limite_Superior As Long) As Long
  Dim res As Long
  Dim lngSwitch As Long
  If limite_Inferior > limite_Superior Then
    lngSwitch = limite_Inferior
    limite_Inferior = limite_Superior
    limite_Superior = lngSwitch
  End If
  Randomize
  If limite_Inferior <= limite_Superior Then
    res = Int((limite_Superior - limite_Inferior + 1) * Rnd + limite_Inferior)
  Else
    res = Int((limite_Inferior - limite_Superior + 1) * Rnd + limite_Superior)
  End If
  If (res >= limite_Inferior) And (res <= limite_Superior) Then
     randomNumberBetween = res
  Else
     LogOnFile "errors.txt", "random number between " & CStr(limite_Inferior) & " and " & CStr(limite_Superior) & " returned " & CStr(res)
     End
     randomNumberBetween = -1
  End If
End Function

Private Function incNumber(thenumber As Long, Ammount As Long) As Long
  Dim res As Long
  res = thenumber + Ammount
  If res < 0 Then
    res = res Mod 10
    res = 10 + res
  Else
    res = res Mod 10
  End If
  incNumber = res
End Function


Private Function php_substr(thestr As String, fromp As Long, amm As Long) As String
  php_substr = Mid$(thestr, fromp + 1, amm)
End Function

