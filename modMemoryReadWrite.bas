Attribute VB_Name = "modMemoryReadWrite"
#Const FinalMode = 1
Option Explicit

Private Enum enPriority_Class
NORMAL_PRIORITY_CLASS = &H20
IDLE_PRIORITY_CLASS = &H40
HIGH_PRIORITY_CLASS = &H80
End Enum

Private Enum enSW
SW_HIDE = 0
SW_NORMAL = 1
SW_MAXIMIZE = 3
SW_MINIMIZE = 6
End Enum

'***********************
'* Win32 Constants . . .
'***********************
Private Const MEM_PRIVATE& = &H20000
Private Const MEM_COMMIT& = &H1000


Private Const INFINITE As Long = &HFFFF
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const TH32CS_SNAPMODULE As Long = 8&
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const WINAPI_TRUE = 1
Private Const PROCESS_TERMINATE = 1
Private Const CREATE_SUSPENDED As Long = &H4

Private Const GW_HWNDFIRST& = 0
Private Const HWND_NOTOPMOST& = -2
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOMOVE& = &H2
Private Const SWP_NOSIZE& = &H1

Private Const PROCESS_VM_READ = (&H10)
Private Const PROCESS_VM_WRITE = (&H20)
Private Const PROCESS_VM_OPERATION = (&H8)
Private Const PROCESS_QUERY_INFORMATION = (&H400)
Private Const PROCESS_QUERY_LIMITED_INFORMATION = (&H1000)
Private Const PROCESS_READ_WRITE_QUERY = PROCESS_VM_READ + PROCESS_VM_WRITE + PROCESS_VM_OPERATION + PROCESS_QUERY_INFORMATION
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Const GW_HWNDNEXT As Long = 2
Private Const MAX_PATH As Long = 260

Private Const PAGE_EXECUTE_READWRITE As Long = &H40&

Private Const STARTF_USESHOWWINDOW = &H1


Private Type MODULEINFO
   lpBaseOfDLL As Long
   SizeOfImage As Long
   EntryPoint As Long
End Type

Private Type PROCESS_MEMORY_COUNTERS
   cb As Long
   PageFaultCount As Long
   PeakWorkingSetSize As Long
   WorkingSetSize As Long
   QuotaPeakPagedPoolUsage As Long
   QuotaPagedPoolUsage As Long
   QuotaPeakNonPagedPoolUsage As Long
   QuotaNonPagedPoolUsage As Long
   PagefileUsage As Long
   PeakPagefileUsage As Long
End Type

Private Type MODULEENTRY32W
dwSize As Long
th32ModuleID As Long
th32ProcessID As Long
GlblcntUsage As Long
ProccntUsage As Long
modBaseAddr As Long
modBaseSize As Long
hModule As Long
szModule(511) As Byte
szExePath(519) As Byte
End Type

Private Type PSAPI_WS_WATCH_INFORMATION
   FaultingPc As Long
   FaultingVa As Long
End Type

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


Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type MEMORY_BASIC_INFORMATION ' 28 bytes
    baseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    State As Long
    Protect As Long
    lType As Long
End Type

Private Type SYSTEM_INFO ' 36 Bytes
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Private Type TypeOffsetInfo
    pid As Long
    offset As Long
End Type



Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
   
Private Declare Function WriteProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Private Declare Function VirtualQueryEx& Lib "Kernel32" (ByVal hProcess As Long, lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long)

Private Declare Sub GetSystemInfo Lib "Kernel32" (lpSystemInfo As SYSTEM_INFO)
   
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long


' Private Declare Function VirtualProtectEx Lib "Kernel32" (ByVal hProcess As Long, ByRef lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long



Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long


Private Declare Sub GetStartupInfo Lib "Kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)

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


Private Declare Function TerminateProcess Lib "Kernel32" Alias "Terminate Process" ( _
 ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Declare Function CreateProcess Lib "Kernel32" _
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
 


Public useDynamicOffset As String
Public useDynamicOffsetBool As Boolean
Public tibiaModuleRegionSize As Long

Public OffsetsCache() As TypeOffsetInfo
Private NextOffset As Long
Private OffsetsCacheSize As Long
Public Function ResetOffsetCache(ByVal parOffsetsCacheSize As Long)
    Dim i As Long
    OffsetsCacheSize = parOffsetsCacheSize
    ReDim OffsetsCache(1 To parOffsetsCacheSize)
    For i = 1 To parOffsetsCacheSize
       OffsetsCache(i).pid = -1
       OffsetsCache(i).offset = 0
    Next i
    NextOffset = 1
End Function

'Public Function getProcessBase(ByVal hProcess As Long, ByVal expectedRegionSize As Long, Optional PIDinsteadHp As Boolean = False) As Long
'    On Error GoTo goterr
'    Dim lpMem As Long, ret As Long, lLenMBI As Long
'    Dim lWritten As Long, CalcAddress As Long, lPos As Long
'    Dim sBuffer As String
'    Dim sSearchString As String, sReplaceString As String
'    Dim si As SYSTEM_INFO
'    Dim mbi As MEMORY_BASIC_INFORMATION
'    Dim realH As Long
'    Dim pid As Long
'    Dim res As Long
'    If PIDinsteadHp = True Then
'        res = GetWindowThreadProcessId(hProcess, pid)
'        realH = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
'        hProcess = realH
'    End If
'    Call GetSystemInfo(si)
'    lpMem = si.lpMinimumApplicationAddress
'    lLenMBI = Len(mbi)
'    ' Scan memory
'    Do While lpMem < si.lpMaximumApplicationAddress
'        mbi.RegionSize = 0
'        ret = VirtualQueryEx(hProcess, ByVal lpMem, mbi, lLenMBI)
'        If ret = lLenMBI Then
'            If (mbi.State = MEM_COMMIT) Then
''           Debug.Print "BaseAddress=" & Hex(mbi.BaseAddress)
''           Debug.Print "AllocationBase=" & Hex(mbi.AllocationBase)
''           Debug.Print "RegionSize=" & Hex(mbi.RegionSize)
'           If (mbi.RegionSize = expectedRegionSize) Then ' this is the interesting region
'                res = mbi.AllocationBase
''                Debug.Print Hex(mbi.AllocationProtect) ' should be = 80
''                Debug.Print Hex(mbi.AllocationBase)
''                Debug.Print Hex(mbi.BaseAddress)
''                Debug.Print Hex(mbi.Protect)
'                'Debug.Print "The correct result is " & CStr(res)
'                ' the new result
'               ' Debug.Print "The new result is " & CStr(getProcessBase2(hProcess, expectedRegionSize, PIDinsteadHp))
'
'                 If PIDinsteadHp = True Then
'                   CloseHandle hProcess
'                End If
'               getProcessBase = res
'               Exit Function
'           End If
'
'           End If
'           lpMem = mbi.BaseAddress + mbi.RegionSize
'        Else
'           Exit Do
'        End If
'    Loop
'    If PIDinsteadHp = True Then
'       CloseHandle hProcess
'    End If
'goterr:
'    getProcessBase = 0
'End Function




Public Function getProcessBase(ByVal hProcess As Long, ByVal expectedRegionSize As Long, Optional PIDinsteadHp As Boolean = False) As Long
    On Error GoTo goterr
    ' expectedRegionSize is used again
    Dim lpMem As Long, ret As Long, lLenMBI As Long
    Dim lWritten As Long, CalcAddress As Long, lPos As Long
    Dim sBuffer As String
    Dim sSearchString As String, sReplaceString As String
    Dim si As SYSTEM_INFO
    Dim mbi As MEMORY_BASIC_INFORMATION
    Dim realH As Long
    Dim pid As Long
    Dim res As Long
    If PIDinsteadHp = True Then
        res = GetWindowThreadProcessId(hProcess, pid)
        realH = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
        hProcess = realH
    End If
    Call GetSystemInfo(si)
    lpMem = si.lpMinimumApplicationAddress
    lLenMBI = Len(mbi)
    ' Scan memory
    Do While lpMem < si.lpMaximumApplicationAddress
        mbi.RegionSize = 0
        ret = VirtualQueryEx(hProcess, ByVal lpMem, mbi, lLenMBI)
        If ret = lLenMBI Then
           If (mbi.State = MEM_COMMIT) Then
                If mbi.AllocationProtect = &H80 Then
                If mbi.baseAddress - mbi.AllocationBase = &H1000 Then
                If mbi.Protect = &H20 Then
                If (mbi.RegionSize = expectedRegionSize) Then
                    res = mbi.AllocationBase
                    'Debug.Print "The new result is " & CStr(res)
                    If PIDinsteadHp = True Then
                      CloseHandle hProcess
                    End If
                    getProcessBase = res
                    Exit Function
                End If
                End If
                End If
                End If
           End If
           lpMem = mbi.baseAddress + mbi.RegionSize
        Else
           Exit Do
        End If
    Loop
    If PIDinsteadHp = True Then
       CloseHandle hProcess
    End If
goterr:
    getProcessBase = 0
End Function

Public Function getProcessOffset(ByVal hProcess As Long, ByVal pid As Long) As Long
    On Error GoTo goterr
    Dim lpMem As Long, ret As Long, lLenMBI As Long
    Dim lWritten As Long, CalcAddress As Long, lPos As Long
    Dim sBuffer As String
    Dim sSearchString As String, sReplaceString As String
    Dim si As SYSTEM_INFO
    Dim mbi As MEMORY_BASIC_INFORMATION
    Dim realH As Long

   Dim res As Long
   Dim i As Long
   Dim theOffset As Long
   If useDynamicOffsetBool = False Then
     getProcessOffset = 0
     Exit Function
   End If
   
     For i = 1 To OffsetsCacheSize
          If (pid = OffsetsCache(i).pid) Then
             getProcessOffset = OffsetsCache(i).offset
             Exit Function
          End If
     Next i
    
    Call GetSystemInfo(si)
    lpMem = si.lpMinimumApplicationAddress
    lLenMBI = Len(mbi)
    
    Do While lpMem < si.lpMaximumApplicationAddress
        mbi.RegionSize = 0
        ret = VirtualQueryEx(hProcess, ByVal lpMem, mbi, lLenMBI)
        If ret = lLenMBI Then
            If (mbi.State = MEM_COMMIT) Then
               If (mbi.RegionSize = tibiaModuleRegionSize) Then ' this is the interesting region
                   res = mbi.AllocationBase - &H400000
                   OffsetsCache(NextOffset).pid = pid
                   OffsetsCache(NextOffset).offset = res
                   NextOffset = NextOffset + 1
                   If NextOffset > OffsetsCacheSize Then
                     NextOffset = 1
                   End If
                   getProcessOffset = res
                   Exit Function
               End If
           End If
           lpMem = mbi.baseAddress + mbi.RegionSize
        Else
           Exit Do
        End If
    Loop
goterr:
    getProcessOffset = 0
End Function

Public Function Memory_ReadByte(ByVal address As Long, ByVal process_Hwnd As Long, _
 Optional absoluteAddress As Boolean = False) As Byte
  
   ' Declare some variables we need
   Dim pid As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   Dim valbuffer As Byte   ' Byte
   
   Dim res As Long
   
   Dim offset As Long

   
    
   ' First get a handle to the "game" window
   If (process_Hwnd = 0) Then Exit Function
   
   ' We can now get the pid
   GetWindowThreadProcessId process_Hwnd, pid
   
   

   ' Use the pid to get a Process Handle
    'phandle = OpenProcess(PROCESS_VM_READ, False, pid)

   phandle = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid) ' more powerfull
   If (phandle = 0) Then
     Debug.Print "Error " & CStr(Err.LastDllError) & ": " & GetDllErrorDescription(Err.LastDllError)
     Exit Function
   End If
   
   '1
   'offset = 0
   If ((useDynamicOffsetBool = True) And (absoluteAddress = False)) Then
     offset = getProcessOffset(phandle, process_Hwnd)
     address = address + offset
   End If
   
   
   ' Read Long
   res = ReadProcessMemory(phandle, address, valbuffer, 1, 0&)
   
   ' Return
   Memory_ReadByte = valbuffer
   
   ' Close the Process Handle
   CloseHandle phandle
  
End Function
Public Function Memory_ReadLong(ByVal address As Long, ByVal process_Hwnd As Long, _
 Optional absoluteAddress As Boolean = False) As Long
  
   ' Declare some variables we need
   Dim pid As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   Dim valbuffer As Long   ' Long
   
   Dim offset As Long
    
   ' First get a handle to the "game" window
   If (process_Hwnd = 0) Then Exit Function
   
   ' We can now get the pid
   GetWindowThreadProcessId process_Hwnd, pid
   
   ' Use the pid to get a Process Handle
   'phandle = OpenProcess(PROCESS_VM_READ, False, pid)
   
   phandle = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid) ' more powerfull
   If (phandle = 0) Then Exit Function
   
   '2
   offset = 0
   If ((useDynamicOffsetBool = True) And (absoluteAddress = False)) Then
     offset = getProcessOffset(phandle, process_Hwnd)
     address = address + offset
   End If
   
   ' Read Long
   ReadProcessMemory phandle, address, valbuffer, 4, 0&
       
   ' Return
   Memory_ReadLong = valbuffer
   
   ' Close the Process Handle
   CloseHandle phandle
  
End Function
Public Function Memory_Analyze1(ByVal StartAddress As Long, ByVal BytesToRead As Long, ByVal Stringify As Boolean, _
                                ByVal StringMinLen As Long, ByVal process_Hwnd As Long, Optional absoluteAddress As Boolean = False) As String
' Declare some variables we need
    Dim PID As Long         ' Used to hold the Process Id
    Dim phandle As Long     ' Holds the Process Handle
    Dim ByteBuf As Byte   ' Byte
    Dim res As String
    Dim offset As Long
    Dim i As Long
    Dim LastBytesRead As Long
    Dim tmpstr As String

    On Error GoTo goterr

    ' First get a handle to the "game" window
    If (process_Hwnd = 0) Then Exit Function

    ' We can now get the pid
    GetWindowThreadProcessId process_Hwnd, PID



    ' Use the pid to get a Process Handle
    'phandle = OpenProcess(PROCESS_VM_READ, False, pid)

    phandle = OpenProcess(PROCESS_READ_WRITE_QUERY, False, PID)    ' more powerfull
    If (phandle = 0) Then
        Debug.Print "Error " & CStr(Err.LastDllError) & ": " & GetDllErrorDescription(Err.LastDllError)
        Exit Function
    End If

    '1
    'offset = 0
    If ((useDynamicOffsetBool = True) And (absoluteAddress = False)) Then
        offset = getProcessOffset(phandle, process_Hwnd)
        StartAddress = StartAddress + offset
    End If
    ' Read string

    For i = 1 To BytesToRead Step 1
        LastBytesRead = ReadProcessMemory(phandle, StartAddress + i, ByteBuf, 1, 0&)
        If LastBytesRead <> 1 Then
            GoTo goterr
            'err.raise?
        End If
        '&H20 to &H7E - http://www.asciitable.com/
        If Stringify And ByteBuf >= &H20 And ByteBuf <= &H7E Then
            tmpstr = tmpstr & Chr(ByteBuf)
        Else
            If Stringify And Len(tmpstr) >= StringMinLen Then
                res = res & " " & tmpstr & " " & GoodHex(ByteBuf)
                tmpstr = ""
            Else
                If Stringify And Len(tmpstr) > 0 Then
                    res = res & " " & Hexarize(tmpstr) & GoodHex(ByteBuf)    ' Hexarize ends with " "
                    tmpstr = ""
                Else
                    res = res & " " & GoodHex(ByteBuf)
                End If
            End If
        End If
    Next i
exitwhile:
    If Stringify And Len(tmpstr) >= StringMinLen Then
        res = res & " " & tmpstr
        tmpstr = ""
    Else
        If Stringify And Len(tmpstr) > 0 Then
            res = res & " " & RTrim(Hexarize(tmpstr))
            tmpstr = ""
        End If
    End If


    ' Close the Process Handle
    CloseHandle phandle
    Memory_Analyze1 = res
    Exit Function
goterr:
    '???
    Memory_Analyze1 = res & "... after reading " & CStr(i) & " bytes, got an error reading at memory location (decimal) " & CStr(StartAddress + i) & " :  Err.Number: " & _
                      CStr(Err.Number) & " Err.Description: " & Err.Description & " Err.LastDllError: " & CStr(Err.LastDllError)
    If phandle <> 0 Then
        CloseHandle phandle
    End If
End Function
Public Function Memory_BlackdAddressToFinalAdddress(ByVal address As Long, ByVal process_Hwnd As Long)
   Dim pid As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   Dim offset As Long
   Dim res As Long
   Dim lasterr As Long
   Dim numberw As Long
   Dim NewProtection As Long
   Dim OldProtection As Long
   Dim readedb As Long
   OldProtection = 0
   ' First get a handle to the "game" window
   If (process_Hwnd = 0) Then
        Memory_BlackdAddressToFinalAdddress = 0
        Exit Function
   End If
   
   ' We can now get the pid
   GetWindowThreadProcessId process_Hwnd, pid
   
   ' Use the pid to get a Process Handle
   phandle = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)

   If (phandle = 0) Then
        Memory_BlackdAddressToFinalAdddress = 0
        Exit Function
   End If
   
  
   offset = 0
   If (useDynamicOffsetBool = True) Then
     offset = getProcessOffset(phandle, process_Hwnd)
     address = address + offset
   End If
   Memory_BlackdAddressToFinalAdddress = address
End Function


'Public Sub Memory_WriteByteFORCE(ByVal Address As Long, ByRef thebytes() As Byte, rsize As Long, ByVal process_Hwnd As Long, _
' Optional absoluteAddress As Boolean = False)
'
'   'Declare some variables we need
'   Dim pid As Long         ' Used to hold the Process Id
'   Dim phandle As Long     ' Holds the Process Handle
'   Dim offset As Long
'   Dim res As Long
'   Dim lasterr As Long
'   Dim numberw As Long
'   Dim NewProtection As Long
'   Dim OldProtection As Long
'   Dim readedb As Long
'   OldProtection = 0
'   ' First get a handle to the "game" window
'   If (process_Hwnd = 0) Then Exit Sub
'
'   ' We can now get the pid
'   GetWindowThreadProcessId process_Hwnd, pid
'
'   ' Use the pid to get a Process Handle
'
'     phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
'
'   If (phandle = 0) Then Exit Sub
'
'   '3
'   offset = 0
'   If ((useDynamicOffsetBool = True) And (absoluteAddress = False)) Then
'     offset = getProcessOffset(phandle, process_Hwnd)
'     Address = Address + offset
'   End If
'
'   VirtualProtectEx phandle, Address, rsize, PAGE_EXECUTE_READWRITE, OldProtection
'
'     ' Write bytes
'     res = WriteProcessMemory(phandle, Address, thebytes(0), rsize, 0)
'     If res = 0 Then
'      lasterr = GetLastError()
'      Debug.Print "error " & CStr(lasterr) & ": " & GetDllErrorDescription(lasterr)
'     End If
'
'   VirtualProtectEx phandle, Address, rsize, OldProtection, NewProtection
'
'   ' Close the Process Handle
'   CloseHandle phandle
'
'End Sub
 
Public Sub Memory_WriteByte(ByVal address As Long, ByVal valbuffer As Byte, ByVal process_Hwnd As Long, _
 Optional absoluteAddress As Boolean = False)

   'Declare some variables we need
   Dim pid As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   Dim offset As Long
   Dim res As Long
   Dim lasterr As Long
   
   ' First get a handle to the "game" window
   If (process_Hwnd = 0) Then Exit Sub
   
   ' We can now get the pid
   GetWindowThreadProcessId process_Hwnd, pid
   
   ' Use the pid to get a Process Handle
   
     phandle = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)

   If (phandle = 0) Then Exit Sub
   
   '3
   offset = 0
   If ((useDynamicOffsetBool = True) And (absoluteAddress = False)) Then
     offset = getProcessOffset(phandle, process_Hwnd)
     address = address + offset
   End If
   
   ' Write byte
  
     res = WriteProcessMemory(phandle, address, valbuffer, 1, 0&)
  

'   If res = 0 Then
'    lasterr = GetLastError()
'    Debug.Print "error " & CStr(lasterr) & ": " & GetDllErrorDescription(lasterr)
'   End If
   
   
   ' Close the Process Handle
   CloseHandle phandle

End Sub
Public Sub Memory_WriteLong(ByVal address As Long, ByVal valbuffer As Long, ByVal process_Hwnd As Long, _
 Optional absoluteAddress As Boolean = False)

   'Declare some variables we need
   Dim pid As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   Dim offset As Long
   
   ' First get a handle to the "game" window
   If (process_Hwnd = 0) Then Exit Sub
   
   ' We can now get the pid
   GetWindowThreadProcessId process_Hwnd, pid
   
   ' Use the pid to get a Process Handle
   phandle = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
   If (phandle = 0) Then Exit Sub
   
   '4
   offset = 0
   If ((useDynamicOffsetBool = True) And (absoluteAddress = False)) Then
     offset = getProcessOffset(phandle, process_Hwnd)
     address = address + offset
   End If
   
   ' Write Long
   WriteProcessMemory phandle, address, valbuffer, 4, 0&
   
   ' Close the Process Handle
   CloseHandle phandle

End Sub


Public Function LaunchTibiaMC(ByVal strTibiaPath As String, Optional useDynamicOffset As String = "no") As String
    Dim pInfo As PROCESS_INFORMATION
    Dim sInfo As STARTUPINFO
    Dim sNull As String
    Dim lSuccess As Long
    Dim lRetValue As Long
    Dim b1 As Byte
    Dim b2 As Byte
    Dim TibiaProcHandle As Long
    Dim theBase As Long
    Dim theOffset As Long
    Dim resWrite As Long
    Dim resRead As Long
    Dim strCurrentIP As String
    Dim valbuffer As Byte
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
    
    
    b1 = multiclientByte1
    b2 = multiclientByte2
    'sInfo.cb = Len(sInfo)
    GetStartupInfo sInfo
    
    ' create tibia process , and pause it at same time
    lSuccess = CreateProcess(sNull, _
                                 loc1, _
                                 ByVal 0&, _
                                 ByVal 0&, _
                                 1&, _
                                 CREATE_SUSPENDED, _
                                 ByVal 0&, _
                                 strTibiaPath, _
                                 sInfo, _
                                 pInfo)
    If lSuccess = 0 Then
        LaunchTibiaMC = "Failed to execute " & strTibiaPath & "tibia.exe"
        Exit Function
    End If
    ' success in creation. Now we can handle the paused process
    
    TibiaProcHandle = pInfo.hProcess
    
    
    ' give that tibia process a little touch of magic only in its memory (file is not modified)

    If useDynamicOffset = "yes" Then
      
    
      theBase = getProcessBase(TibiaProcHandle, tibiaModuleRegionSize)
     
      'theBase = getProcessBase2(TibiaProcHandle, tibiaModuleRegionSize)
      
      theOffset = theBase - &H400000
    Else
      theOffset = 0
    End If

    If TibiaVersionLong <= 772 Then
      Select Case TibiaVersionLong
      Case 772
        adrMulticlient = CLng("&H4DA6E5")
      Case 760
        adrMulticlient = CLng("&H44DE45")
      Case 740
        adrMulticlient = CLng("&H4DA6E5")
      End Select
      b1 = &HEB
     resWrite = WriteProcessMemory(TibiaProcHandle, adrMulticlient + theOffset, b1, 1, 0&)
     If (resWrite = 0) Then
       Debug.Print "Memory Error. Unable to write mc byte."
     End If
    Else
      WriteProcessMemory TibiaProcHandle, adrMulticlient + theOffset, b1, 1, 0&
      WriteProcessMemory TibiaProcHandle, adrMulticlient + theOffset + 1, b2, 1, 0&
    End If
    
 
    lRetValue = ResumeThread(pInfo.hThread)

    LaunchTibiaMC = ""
End Function

