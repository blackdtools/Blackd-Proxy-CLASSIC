Attribute VB_Name = "modProcessKiller"
Option Explicit

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Type LUID
   LowPart As Long
   HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   pLuid As LUID
   Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   TheLuid As LUID
   Attributes As Long
End Type


Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'Purpose   :    Terminates a process given a process ID or a the handle to a form.
'Inputs    :    [lProcessID]          The process ID (or PID) to terminate.
'          [lHwndWindow]          Any window handle belonging to the application.
'Outputs   :    Returns True on success.
'Author    :    Andrew Baker
'Date      :    28/04/2001
'Notes     :    In WIN NT, click the "Processes" tab in the "Task Manager"
'          to see the process ID (or PID) for an application.
'          Must specify either lHwndWindow or lProcessID.
'          Equivalent to pressing Alt+Ctrl+Del then "End Task"

Public Function ProcessTerminate(Optional lProcessID As Long, Optional lHwndWindow As Long) As Boolean
   Dim lhwndProcess As Long
   Dim lExitCode As Long
   Dim lRetVal As Long
   Dim lhThisProc As Long
   Dim lhTokenHandle As Long
   Dim tLuid As LUID
   Dim tTokenPriv As TOKEN_PRIVILEGES, tTokenPrivNew As TOKEN_PRIVILEGES
   Dim lBufferNeeded As Long
   
   Dim PROCESS_ALL_ACCESS As Long
   Dim PROCESS_TERMINATE As Long
   PROCESS_ALL_ACCESS = &H1F0FFF
   PROCESS_TERMINATE = &H1
   Const ANYSIZE_ARRAY = 1, TOKEN_ADJUST_PRIVILEGES = &H20
   Const TOKEN_QUERY = &H8, SE_DEBUG_NAME As String = "SeDebugPrivilege"
   Const SE_PRIVILEGE_ENABLED = &H2

   On Error Resume Next
   If lHwndWindow Then
       'Get the process ID from the window handle
       lRetVal = GetWindowThreadProcessId(lHwndWindow, lProcessID)
   End If
   
   If lProcessID Then
       'Give Kill permissions to this process
       lhThisProc = GetCurrentProcess
       
       OpenProcessToken lhThisProc, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, lhTokenHandle
       LookupPrivilegeValue "", SE_DEBUG_NAME, tLuid
       'Set the number of privileges to be change
       tTokenPriv.PrivilegeCount = 1
       tTokenPriv.TheLuid = tLuid
       tTokenPriv.Attributes = SE_PRIVILEGE_ENABLED
       'Enable the kill privilege in the access token of this process
       AdjustTokenPrivileges lhTokenHandle, False, tTokenPriv, Len(tTokenPrivNew), tTokenPrivNew, lBufferNeeded

       'Open the process to kill
       lhwndProcess = OpenProcess(PROCESS_TERMINATE, 0, lProcessID)
   
       If lhwndProcess Then
         'Obtained process handle, kill the process
         ProcessTerminate = CBool(TerminateProcess(lhwndProcess, lExitCode))
         Call CloseHandle(lhwndProcess)
       End If
   End If
   On Error GoTo 0
End Function
