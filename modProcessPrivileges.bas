Attribute VB_Name = "modProcessPrivileges"
Option Explicit

' Constants used for various API calls. Refer to MSDN for detailed
' information about what these constants mean.

Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const ANYSIZE_ARRAY = 1
'Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const SE_DEBUG_NAME As String = "SeDebugPrivilege"
Private Const SE_PRIVILEGE_ENABLED = &H2

' Structures used with various API calls.
' Refer to MSDN for detailed information
' about what these structures are, and how they are used.

Private Type LARGE_INTEGER
   LowPart As Long
   HighPart As Long
End Type

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
   Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

' Refer to the MSDN for detailed information on
' all of these API calls.


Private Declare Function GetLastError Lib "Kernel32" () As Long

Private Declare Function FormatMessage Lib "Kernel32" Alias _
   "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, _
   ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
   ByVal lpBuffer As String, ByVal nSize As Long, _
   Arguments As Long) As Long

Private Declare Function CloseHandle Lib "Kernel32" _
   (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcess Lib "Kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" _
   (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, _
TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias _
   "LookupPrivilegeValueA" (ByVal lpSystemName As String, _
   ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" _
   (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
   NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
   PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32" _
   (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
   ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "Kernel32" _
   (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
   
Public Function GetDllErrorDescription(ByVal lngCode As Long) As String

Dim sError As String * 500
Dim lErrMsg As Long

lErrMsg = FormatMessage(&H1000, ByVal 0&, lngCode, 0, sError, Len(sError), 0)
GetDllErrorDescription = Trim$(sError)

End Function
   
Public Sub SetAllPrivilegesForMe()
On Error GoTo goterr
         Dim hProcess As Long           ' Handle to your current process
         Dim hToken As Long             ' Handle to your process token.
         Dim lPrivilege As Long         ' Privilege to enable/disable
         Dim iPrivilegeflag As Boolean  ' Flag whether to enable/disable
                                        ' the privilege of concern.
         Dim lResult As Long            ' Result call of various APIs (long)
         Dim bResult As Boolean         ' Result call of various APIs (boolean)
         
         ' get our current process handle
         hProcess = GetCurrentProcess

         ' open the tokens for this process
         lResult = OpenProcessToken(hProcess, TOKEN_ADJUST_PRIVILEGES Or _
                                   TOKEN_QUERY, hToken)

         ' if OpenProcessToken fails, the return result is zero, test for
         ' success here

         If (lResult = 0) Then
            Debug.Print "Error: Unable To Open Process Token : " & Err.LastDllError & " : " & GetDllErrorDescription(Err.LastDllError)
            CloseHandle (hToken)
            Exit Sub
         End If

         ' Now that you have the token for this process, you want to set
         ' the SE_DEBUG_NAME privilege.

         bResult = SetPrivilege(hToken, SE_DEBUG_NAME, True)

         ' Make sure you could set the privilege on this token

         If (bResult = False) Then
            Debug.Print "Error : Could Not Set SeDebug Privilege on Token Handle"
            CloseHandle (hToken)
            Exit Sub
         Else
            'Debug.Print "Blackd Proxy have full power now."
         End If
         ' PROBLEM SOLVED!
goterr:
         Exit Sub
End Sub
Public Function SetPrivilege(ByRef hToken As Long, ByVal Privilege As String, ByVal bSetFlag As Boolean) As Boolean

         Dim TP As TOKEN_PRIVILEGES          ' Used in getting the current
                                             ' token privileges
         Dim TPPrevious As TOKEN_PRIVILEGES  ' Used in setting the new
                                             ' token privileges
         Dim LUID As LUID                    ' Stores the Local Unique
                                             ' Identifier - refer to MSDN
         Dim cbPrevious As Long              ' Previous size of the
                                               ' TOKEN_PRIVILEGES structure
         Dim lResult As Long                 ' Result of various API calls
         Dim lastDLLerrorID As Long
         ' Grab the size of the TOKEN_PRIVILEGES structure,
         ' used in making the API calls.
         cbPrevious = Len(TP)

         ' Grab the LUID for the request privilege.
         lResult = LookupPrivilegeValue("", Privilege, LUID)

         ' If LoopupPrivilegeValue fails, the return result will be zero.
         ' Test to make sure that the call succeeded.
         If (lResult = 0) Then
            SetPrivilege = False
            Exit Function
         End If

         ' Set up basic information for a call.
         ' You want to retrieve the current privileges
         ' of the token under concern before you can modify them.
         TP.PrivilegeCount = 1
         TP.Privileges(0).pLuid = LUID
         TP.Privileges(0).Attributes = 0
         SetPrivilege = lResult

         ' You need to acquire the current privileges first
         'lResult = AdjustTokenPrivileges(hToken, -1, TP, Len(TP), _
                                        TPPrevious, cbPrevious)
lResult = AdjustTokenPrivileges(hToken, False, TP, Len(TP), TPPrevious, cbPrevious)
         ' If AdjustTokenPrivileges fails, the return result is zero,
         ' test for success.
         If (lResult = 0) Then
            lastDLLerrorID = GetLastError()
            Debug.Print "dll failed with error " & lastDLLerrorID & " : " & GetDllErrorDescription(lastDLLerrorID)
            SetPrivilege = False
            Exit Function
         End If

         ' Now you can set the token privilege information
         ' to what the user is requesting.
         TPPrevious.PrivilegeCount = 1
         TPPrevious.Privileges(0).pLuid = LUID

         ' either enable or disable the privilege,
         ' depending on what the user wants.
         Select Case bSetFlag
            Case True: TPPrevious.Privileges(0).Attributes = _
                       TPPrevious.Privileges(0).Attributes Or _
                       (SE_PRIVILEGE_ENABLED)
            Case False: TPPrevious.Privileges(0).Attributes = _
                        TPPrevious.Privileges(0).Attributes Xor _
                        (SE_PRIVILEGE_ENABLED And _
                        TPPrevious.Privileges(0).Attributes)
         End Select

         ' Call adjust the token privilege information.
         lResult = AdjustTokenPrivileges(hToken, False, TPPrevious, _
                                        cbPrevious, TP, cbPrevious)

         ' Determine your final result of this function.
         If (lResult = 0) Then
            ' You were not able to set the privilege on this token.
            SetPrivilege = False
            Exit Function
         Else
            ' You managed to modify the token privilege
            SetPrivilege = True
            Exit Function
         End If

End Function

