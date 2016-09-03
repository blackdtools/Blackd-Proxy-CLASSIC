Attribute VB_Name = "modMC"
#Const FinalMode = 1
Option Explicit

' The special multiclient address:
'Public Const adrMulticlient = &H502BB5
' The bytes that do the trick:
'Public Const multiclientByte1 = &H90
'Public Const multiclientByte2 = &H90

' The special multiclient address:
Public adrMulticlient As Long
' The bytes that do the trick:
Public multiclientByte1 As Byte
Public multiclientByte2 As Byte


'***********************
'* Win32 Constants . . .
'***********************
Private Const INFINITE As Long = &HFFFF
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const WINAPI_TRUE = 1
Private Const PROCESS_TERMINATE = 1
Private Const CREATE_SUSPENDED As Long = &H4


Private Const STARTF_USESHOWWINDOW = &H1
Private Enum enSW
SW_HIDE = 0
SW_NORMAL = 1
SW_MAXIMIZE = 3
SW_MINIMIZE = 6
End Enum

Private Type PROCESS_INFORMATION
hProcess As Long
hThread As Long
dwProcessId As Long
dwThreadId As Long
End Type

Private Type STARTUPINFO
cb As Long
lpReserved As Long
lpDesktop As Long
lpTitle As Long
dwX As Long
dwY As Long
dwXSize As Long
dwYSize As Long
dwXCountChars As Long
dwYCountChars As Long
dwFillAttribute As Long
dwFlags As Long
wShowWindow As Integer
cbReserved2 As Integer
lpReserved2 As Byte
hStdInput As Long
hStdOutput As Long
hStdError As Long
End Type

Private Type SECURITY_ATTRIBUTES
nLength As Long
lpSecurityDescriptor As Long
bInheritHandle As Long
End Type

Private Enum enPriority_Class
NORMAL_PRIORITY_CLASS = &H20
IDLE_PRIORITY_CLASS = &H40
HIGH_PRIORITY_CLASS = &H80
End Enum

Private Const GW_HWNDFIRST& = 0
Private Const HWND_NOTOPMOST& = -2
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOMOVE& = &H2
Private Const SWP_NOSIZE& = &H1

Private Const PROCESS_VM_READ = (&H10)
Private Const PROCESS_VM_WRITE = (&H20)
Private Const PROCESS_VM_OPERATION = (&H8)
Private Const PROCESS_QUERY_INFORMATION = (&H400)
Private Const PROCESS_READ_WRITE_QUERY = PROCESS_VM_READ + PROCESS_VM_WRITE + PROCESS_VM_OPERATION + PROCESS_QUERY_INFORMATION

Private Declare Function GetCurrentProcess _
                                                    Lib "kernel32" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function SetThreadPriority Lib "kernel32" _
                                                       (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long

Private Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)

Private Const THREAD_BASE_PRIORITY_LOWRT As Long = 15 ' value that gets a thread to LowRealtime-1
Private Const THREAD_BASE_PRIORITY_MAX As Long = 2 ' maximum thread base priority boost
Private Const THREAD_BASE_PRIORITY_MIN As Long = -2 ' minimum thread base priority boost
Private Const THREAD_BASE_PRIORITY_IDLE As Long = -15 ' value that gets a thread to idle

Private Enum ThreadPriority
    THREAD_PRIORITY_LOWEST = -2
    THREAD_PRIORITY_BELOW_NORMAL = -1
    THREAD_PRIORITY_NORMAL = 0
    THREAD_PRIORITY_HIGHEST = 2
    THREAD_PRIORITY_ABOVE_NORMAL = 1
    THREAD_PRIORITY_TIME_CRITICAL = 15 ' THREAD_BASE_PRIORITY_LOWRT
    THREAD_PRIORITY_IDLE = -15 'THREAD_BASE_PRIORITY_IDLE
End Enum


Private Declare Function TerminateProcess Lib "kernel32" Alias "Terminate Process" ( _
 ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Declare Function CreateProcess Lib "kernel32" _
         Alias "CreateProcessA" _
         (ByVal lpApplicationName As String, _
         ByVal lpCommandLine As String, _
         lpProcessAttributes As Any, _
         lpThreadAttributes As Any, _
         ByVal bInheritHandles As Long, _
         ByVal dwCreationFlags As Long, _
         lpEnvironment As Any, _
         ByVal lpCurrentDriectory As String, _
         lpStartupInfo As STARTUPINFO, _
         lpProcessInformation As PROCESS_INFORMATION) As Long



 
 Private Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
 

 
Public Function LaunchTibia(ByVal strTibiaPath, withMC As Boolean) As String
      Dim prev As String
      Dim loc1 As String
      Dim fs As scripting.FileSystemObject
      Set fs = New scripting.FileSystemObject
      If strTibiaPath = "" Then
        loc1 = ""
      Else
        If Right$(strTibiaPath, 1) = "\" Then
          loc1 = strTibiaPath & "Tibia.exe"
        Else
          loc1 = strTibiaPath & "\Tibia.exe"
        End If
      End If
      If fs.FileExists(loc1) = False Then
       prev = TibiaExePath
       TibiaExePath = autoGetTibiaFolder()
       If TibiaExePath = "" Then
        TibiaExePath = prev
       End If
       strTibiaPath = TibiaExePath
      End If
    If withMC = True Then
        LaunchTibia = LaunchTibiaMC(strTibiaPath, useDynamicOffset)
    Else
        LaunchTibia = LaunchFileNormalWay(strTibiaPath)
    End If
End Function

Public Function autoGetTibiaFolder(Optional ByVal ParTibiaFolder As String = "") As String
    On Error GoTo gotErr
    Dim tpath As String
    If ParTibiaFolder = "" Then
        If DefaultTibiaFolder = "" Then
            ParTibiaFolder = "Tibia"
        Else
            ParTibiaFolder = DefaultTibiaFolder
        End If
    End If
    tpath = ""
    Dim fs As scripting.FileSystemObject
    Set fs = New scripting.FileSystemObject
    tpath = GetProgFolder()
    If Right$(tpath, 1) <> "\" Then
        tpath = tpath & "\"
    End If
    tpath = tpath & ParTibiaFolder & "\"
    autoGetTibiaFolder = tpath
    Exit Function
gotErr:
    autoGetTibiaFolder = ""
End Function

Public Function autoGetMagebotFolder() As String
    On Error GoTo gotErr
    Dim tpath As String
    Dim fs As scripting.FileSystemObject
    Set fs = New scripting.FileSystemObject
    tpath = GetProgFolder()
    If Right$(tpath, 1) <> "\" Then
        tpath = tpath & "\"
    End If
    tpath = tpath & "Magebot\"
    autoGetMagebotFolder = tpath
    Exit Function
gotErr:
    autoGetMagebotFolder = ""
End Function

'Public Function autoGetMagebotExe() As String
'    On Error GoTo goterr
'    autoGetMagebotExe = autoGetFileContaining(MagebotPath, "magebot")
'    Exit Function
'goterr:
'    autoGetMagebotExe = ""
'End Function

Public Function autoGetFileContaining(strPath As String, strCriteria) As String
    On Error GoTo gotErr
    Dim tpath As String
    Dim sName As String
    Dim lPos As Long
    Dim fs As scripting.FileSystemObject
    Dim fol As scripting.Folder
    Dim fil As scripting.File
    Set fs = New scripting.FileSystemObject
    Set fol = fs.GetFolder(strPath)
    For Each fil In fol.Files
        sName = fil.name
        lPos = InStr(1, sName, strCriteria, vbTextCompare)
        If lPos > 0 Then
            autoGetFileContaining = sName
            Exit Function
        End If
    Next fil
    autoGetFileContaining = ""
    Exit Function
gotErr:
    autoGetFileContaining = ""
End Function



Public Function LaunchFileNormalWay(ByVal strTibiaPath As String, Optional strFile As String = "tibia.exe") As String
    Dim pInfo As PROCESS_INFORMATION
    Dim sInfo As STARTUPINFO
    Dim sNull As String
    Dim lSuccess As Long
    Dim lRetValue As Long
    Dim b1 As Byte
    Dim b2 As Byte
    Dim TibiaProcHandle As Long
    
      Dim loc1 As String

      If strTibiaPath = "" Then
        loc1 = ""
      Else
        If Right$(strTibiaPath, 1) = "\" Then
          loc1 = strTibiaPath
        Else
          loc1 = strTibiaPath & "\"
        End If
      End If
    
    'b1 = multiclientByte1
    'b2 = multiclientByte2
    'sInfo.cb = Len(sInfo)
    GetStartupInfo sInfo
    
    ' create tibia process , and pause it at same time
    lSuccess = CreateProcess(sNull, _
                                 loc1 & strFile, _
                                 ByVal 0&, _
                                 ByVal 0&, _
                                 1&, _
                                 0&, _
                                 ByVal 0&, _
                                 loc1, _
                                 sInfo, _
                                 pInfo)
    If lSuccess = 0 Then
        LaunchFileNormalWay = "Failed to execute " & strTibiaPath & strFile
        Exit Function
    End If
    ' success in creation. Now we can handle the paused process
    
    'TibiaProcHandle = pInfo.hProcess
    ' give that tibia process a little touch of magic only in its memory (file is not modified)

    'WriteProcessMemory TibiaProcHandle, adrMulticlient, b1, 1, 0&
    'WriteProcessMemory TibiaProcHandle, adrMulticlient + 1, b2, 1, 0&
    
    ' now tibia can start executing, play!
    'lRetValue = ResumeThread(pInfo.hThread)
    LaunchFileNormalWay = ""
End Function
