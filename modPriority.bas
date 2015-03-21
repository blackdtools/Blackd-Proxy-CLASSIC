Attribute VB_Name = "modPriority"
#Const FinalMode = 0
Option Explicit

' (c) Copyright 2003 Andrew Novick.
' You may use this code in your projects, including projects
' that you sell so long as there is substantial additional
' content. All other rights including rights to publication
' are reserved.
 

' Win32 API declarations
Public Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
End Type

Private Declare Function GetWindowModuleFileName Lib "user32.dll" (ByVal hwnd As Long, ByVal pszFileName As String, ByVal cchFileNameMax As Long) As Long


Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long


Private Declare Function GetVersionEx Lib "Kernel32" _
 Alias "GetVersionExA" (LpVersionInformation _
 As OSVERSIONINFO) As Long
Private Declare Function GetCurrentProcess _
                                                    Lib "Kernel32" () As Long
Private Declare Function GetCurrentProcessId Lib "Kernel32" () As Long
Private Declare Function GetCurrentThread Lib "Kernel32" () As Long
Private Declare Function GetCurrentThreadId Lib "Kernel32" () As Long
Private Declare Function SetThreadPriority Lib "Kernel32" _
                                                       (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function GetThreadPriority Lib "Kernel32" (ByVal hThread As Long) As Long

Private Const THREAD_BASE_PRIORITY_LOWRT As Long = 15 ' value that gets a thread to LowRealtime-1
Private Const THREAD_BASE_PRIORITY_MAX As Long = 2 ' maximum thread base priority boost
Private Const THREAD_BASE_PRIORITY_MIN As Long = -2 ' minimum thread base priority boost
Private Const THREAD_BASE_PRIORITY_IDLE As Long = -15 ' value that gets a thread to idle

Public Enum ThreadPriority
    THREAD_PRIORITY_LOWEST = -2
    THREAD_PRIORITY_BELOW_NORMAL = -1
    THREAD_PRIORITY_NORMAL = 0
    THREAD_PRIORITY_HIGHEST = 2
    THREAD_PRIORITY_ABOVE_NORMAL = 1
    THREAD_PRIORITY_TIME_CRITICAL = 15 ' THREAD_BASE_PRIORITY_LOWRT
    THREAD_PRIORITY_IDLE = -15 'THREAD_BASE_PRIORITY_IDLE
End Enum

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Private Declare Function SetPriorityClass Lib "Kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetPriorityClass Lib "Kernel32" (ByVal hProcess As Long) As Long

Public Declare Function GetLastError _
    Lib "Kernel32" () As Long
Public Declare Function FormatMessage _
    Lib "Kernel32" _
    Alias "FormatMessageA" _
   (ByVal dwFlags As Long, _
    lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    Arguments As Long) As Long

Public Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000

' Used by the OpenProcess API call
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_SET_INFORMATION As Long = &H200

' Used by SetPriorityClass
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const BELOW_NORMAL_PRIORITY_CLASS = 16384
Private Const ABOVE_NORMAL_PRIORITY_CLASS = 32768
Private Const IDLE_PRIORITY_CLASS = &H40
Private Const HIGH_PRIORITY_CLASS = &H80
Private Const REALTIME_PRIORITY_CLASS = &H100

Public Enum ProcessPriorities
    ppidle = IDLE_PRIORITY_CLASS
    ppbelownormal = BELOW_NORMAL_PRIORITY_CLASS
    ppAboveNormal = ABOVE_NORMAL_PRIORITY_CLASS
    ppNormal = NORMAL_PRIORITY_CLASS
    ppHigh = HIGH_PRIORITY_CLASS
    ppRealtime = REALTIME_PRIORITY_CLASS
End Enum

Public MyPriorityID As Long
Public TibiaPriorityID As Long
Public PriorityErrors As String

Public Function GetDllErrorMessage(lErrNum As Long) As String
    Dim sError As String * 500
    Dim lErrMsg As Long
    lErrMsg = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, _
                ByVal 0&, lErrNum, 0, sError, Len(sError), 0)
    GetDllErrorMessage = Trim(sError)
End Function

Public Function ProcessPriorityName(ByVal priority As ProcessPriorities) As String

Dim sName As String

Select Case priority

    Case ppidle
        sName = "Idle"

    Case ppbelownormal
        sName = "Below Normal"

    Case ppNormal
        sName = "Normal"

    Case ppAboveNormal
        sName = "Above Normal"

    Case ppHigh
        sName = "High"

    Case ppRealtime
        sName = "Realtime"

    Case Else
        sName = "Unknown:" & CStr(priority)

End Select

ProcessPriorityName = sName

End Function

Public Function ProcessPriorityGet(Optional ByVal ProcessID As Long, Optional ByVal hwnd As Long) As Long

    ' Gets the process priority identified by an Id, a hWnd

    '  or if not identified, then the current process

    Dim hProc As Long
    Const fdwAccess As Long = PROCESS_QUERY_INFORMATION

    ' If not passed a PID, then find value from hWnd.

    If ProcessID = 0 Then
        If hwnd <> 0 Then
            Call GetWindowThreadProcessId(hwnd, ProcessID)
        Else
            ProcessID = GetCurrentProcessId()
        End If
    End If

    '   Need to open process with simple query rights,

    ' get the current setting, and close handle.
    hProc = OpenProcess(fdwAccess, 0&, ProcessID)
    ProcessPriorityGet = GetPriorityClass(hProc)

    Call CloseHandle(hProc)

End Function

Public Function ProcessPrioritySet( _
                    Optional ByVal ProcessID As Long, _
                    Optional ByVal hwnd As Long, _
                    Optional ByVal priority As ProcessPriorities = NORMAL_PRIORITY_CLASS _
                    ) As Boolean

    Dim hProc As Long
    Dim lonRes As Long
    Dim res As Boolean
    Dim tmp As Long
    #If FinalMode Then
    On Error GoTo returnValue
    #End If
    Const fdwAccess1 As Long = PROCESS_QUERY_INFORMATION Or PROCESS_SET_INFORMATION
    Const fdwAccess2 As Long = PROCESS_QUERY_INFORMATION
    res = False
    ' If not passed a PID, then find value from hWnd.

    If ProcessID = 0 Then
        If hwnd <> 0 Then
           GetWindowThreadProcessId hwnd, ProcessID
           If ProcessID = 0 Then
             PriorityErrors = PriorityErrors & " ; GetWindowThreadProcessId FAILED"
           End If
        Else
            ProcessID = GetCurrentProcessId()
            If ProcessID = 0 Then
              PriorityErrors = PriorityErrors & " ; GetCurrentProcessId FAILED"
            End If
        End If
    End If
     PriorityErrors = PriorityErrors & " ; ProcessID = " & CStr(ProcessID)
    ' Need to open process with setinfo rights.
    hProc = OpenProcess(fdwAccess1, 0&, ProcessID)
    If hProc Then
        ' Attempt to set new priority.
        PriorityErrors = PriorityErrors & " ; hProc = " & CStr(hProc)
        lonRes = SetPriorityClass(hProc, priority)
        If lonRes = 0 Then
          PriorityErrors = PriorityErrors & " ; SetPriorityClass FAILED : " & GetDllErrorMessage(Err.LastDllError)
          res = False
        Else
          res = True
        End If
    Else
        PriorityErrors = PriorityErrors & " ; OpenProcess FAILED : " & GetDllErrorMessage(Err.LastDllError)
        ' Weren't allowed to setinfo, so just open to
        ' enable return of current priority setting.
        hProc = OpenProcess(fdwAccess2, 0&, ProcessID)
    End If

    ' Get current/new setting.
    'ProcessPrioritySet = GetPriorityClass(hProc)
    ' Clean up.
    Call CloseHandle(hProc)
    ProcessPrioritySet = res
    Exit Function
returnValue:
    PriorityErrors = PriorityErrors & " ; Unexpected error at ProcessPrioritySet : " & Err.Description
    ProcessPrioritySet = False
End Function

Public Function ProcIDFromhWnd(ByVal hwnd As Long) As Long
    Dim idProc As Long
    Call GetWindowThreadProcessId(hwnd, idProc)
    ProcIDFromhWnd = idProc
End Function
Public Function ProcFromProcID(idProc As Long) As Long
    ProcFromProcID = OpenProcess(PROCESS_QUERY_INFORMATION Or _
                                 PROCESS_VM_READ, 0, idProc)
End Function

Public Function SetMyOwnPriority( _
 Optional ByVal ProcessPriority As ProcessPriorities = ppNormal _
 ) As Boolean
  Dim hThread As Long
  Dim rc As Long
  Dim res As Boolean
  #If FinalMode Then
  On Error GoTo endBadly
  #End If
  res = False
  res = ProcessPrioritySet(, , ProcessPriority)
  If res = False Then
    PriorityErrors = PriorityErrors & " ; ProcessPrioritySet FAILED"
  End If
  SetMyOwnPriority = res
  Exit Function
endBadly:
  PriorityErrors = PriorityErrors & " ; Unexpected error at ProcessPriority : " & Err.Description
  SetMyOwnPriority = False
End Function

Public Function SetProcessPriorityByHwnd(hwndPar As Long, _
 Optional ByVal ProcessPriority As ProcessPriorities = ppNormal _
 ) As Boolean
  Dim rc As Long
  Dim res As Boolean
  #If FinalMode Then
  On Error GoTo returnValue
  #End If
  res = False
  res = ProcessPrioritySet(, hwndPar, ProcessPriority)
  If res = False Then
    PriorityErrors = PriorityErrors & " ; ProcessPrioritySet FAILED"
  End If
  SetProcessPriorityByHwnd = res
  Exit Function
returnValue:
  PriorityErrors = PriorityErrors & " ; Unexpected error at SetProcessPriorityByHwnd : " & Err.Description
  SetProcessPriorityByHwnd = False
End Function

Public Function UpdateMyPriority() As Boolean
  Dim pok As Boolean
  #If FinalMode Then
  On Error GoTo endBadly
  #End If
  PriorityErrors = "UpdateMyPriority() was called"
  pok = False
  Select Case MyPriorityID
  Case 0
    pok = SetMyOwnPriority(ppidle)
  Case 1
    pok = SetMyOwnPriority(ppbelownormal)
  Case 2
    pok = SetMyOwnPriority(ppNormal)
  Case 3
    pok = SetMyOwnPriority(ppAboveNormal)
  Case 4
    pok = SetMyOwnPriority(ppHigh)
  Case 5
    pok = SetMyOwnPriority(ppRealtime)
  Case Else
    MyPriorityID = 2
    pok = SetMyOwnPriority(ppNormal)
  End Select
  If pok = False Then
    PriorityErrors = PriorityErrors & " ; SetMyOwnPriority FAILED"
    frmAdvanced.lblMessage.Caption = "FAILED TO CHANGE CPU PRIORITIES ! (1)"
    frmAdvanced.lblMessage.ForeColor = &HFF&
  Else
    frmAdvanced.lblMessage.Caption = UCase("Successfully CHANGED CPU PRIORITIES")
    frmAdvanced.lblMessage.ForeColor = &HFF00&
  End If
  UpdateMyPriority = pok
  Exit Function
endBadly:
  PriorityErrors = PriorityErrors & " ; Unexpected error at UpdateMyPriority : " & Err.Description
  UpdateMyPriority = False
End Function

Public Function UpdateTibiaPriority() As Boolean
  Dim tibiaclient As Long
  'Dim hWndDesktop As Long
  Dim pok As Boolean
  #If FinalMode Then
  On Error GoTo endBadly
  #End If
  PriorityErrors = "UpdateTibiaPriority() was called"
  pok = True
  'hWndDesktop = GetDesktopWindow()
  tibiaclient = 0
  Do
    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      PriorityErrors = PriorityErrors & " ; No tibia clients found"
      Exit Do
    Else
      PriorityErrors = PriorityErrors & " ; Changing priority to client " & CStr(tibiaclient)
      'change priority
      Select Case TibiaPriorityID
      Case 0
        pok = SetProcessPriorityByHwnd(tibiaclient, ppidle)
      Case 1
        pok = SetProcessPriorityByHwnd(tibiaclient, ppbelownormal)
      Case 2
        pok = SetProcessPriorityByHwnd(tibiaclient, ppNormal)
      Case 3
        pok = SetProcessPriorityByHwnd(tibiaclient, ppAboveNormal)
      Case 4
        pok = SetProcessPriorityByHwnd(tibiaclient, ppHigh)
      Case 5
        pok = SetProcessPriorityByHwnd(tibiaclient, ppRealtime)
      Case Else
        TibiaPriorityID = 2
        pok = SetProcessPriorityByHwnd(tibiaclient, ppNormal)
      End Select
      If pok = False Then
        PriorityErrors = PriorityErrors & " ; SetProcessPriorityByHwnd FAILED"
        GoTo justend
      End If
    End If
  Loop
justend:
  If pok = False Then
    frmAdvanced.lblMessage.Caption = "FAILED TO CHANGE CPU PRIORITIES ! (2)"
    frmAdvanced.lblMessage.ForeColor = &HFF&
  Else
    frmAdvanced.lblMessage.Caption = UCase("Successfully CHANGED CPU PRIORITIES")
    frmAdvanced.lblMessage.ForeColor = &HFF00&
  End If
   UpdateTibiaPriority = pok
   Exit Function
endBadly:
   PriorityErrors = PriorityErrors & " ; Unexpected error at UpdateTibiaPriority : " & Err.Description
   UpdateTibiaPriority = False
End Function
