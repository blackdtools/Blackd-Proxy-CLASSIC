Attribute VB_Name = "modTibia11Functions"
' Change to Tibia11allowed = 1 to allow Tibia 11 configs
' Don't worry! If you are a programmer capable to compile this code
' then you are authorized to use Tibia 11 configs even if you didn't purchase gold.
' However, you should not share it with other people.
#Const Tibia11allowed = 0
#Const FinalMode = 1
#Const DebugConEvents = 0
Option Explicit
#If Tibia11allowed = 1 Then
    Public Const Tibia11allowed As Boolean = True
#Else
    Public Const Tibia11allowed As Boolean = False
#End If

#If DebugConEvents = 1 Then
    Public Const cteDebugConEvents As Boolean = True
#Else
    Public Const cteDebugConEvents As Boolean = False
#End If

'Public Const defaultGameServerEnd As String = "-lb.ciproxy.com"

Public Const CTE_NOT_CONNECTED As Integer = 0
Public Const CTE_LOGIN_CHAR_SELECTION As Integer = 1
Public Const CTE_CONNECTING As Integer = 2
Public Const CTE_GAME_CONNECTED As Integer = 3
    

'***********************
'* Win32 Constants . . .
'***********************
Private Type POINTAPI 'Type to hold coordinates
    x As Long
    y As Long
End Type

Private Const TH32CS_SNAPMODULE As Long = &H8
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



Private Const MEM_PRIVATE& = &H20000
Private Const MEM_COMMIT& = &H1000


Private Const INFINITE As Long = &HFFFF
Private Const TH32CS_SNAPPROCESS As Long = &H2

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

Private Const PAGE_READONLY As Long = &H2&
Private Const PAGE_EXECUTE_READWRITE As Long = &H40&

Private Const STARTF_USESHOWWINDOW = &H1

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_VISIBLE = &H10000000
Private Const WS_EX_APPWINDOW = &H40000
Private Type HWND_TEXT
 Window_Handle As Long
 Window_Title As String
End Type
Private ColCounter As Long
Private HandleTextCollection() As HWND_TEXT

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
    Offset As Long
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


Public Type TibiaServerEntry
    Id As Long
    name As String
    url As String
    port As Long
    this_register_adr As Long
    name_adr As Long
    url_adr As Long
    port_adr As Long
    url2 As String
    url2_adr As Long
   ' rawbytes() As Byte
End Type

Public Type TibiaCharListEntry
    Id As Long
    name As String
    server As String
    entry_address As Long
    name_address As Long
End Type



'Private Declare Function VirtualProtectEx Lib "Kernel32" (ByVal hProcess As Long, ByRef lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetClassName Lib "user32" _
   Alias "GetClassNameA" _
   (ByVal hWnd As Long, _
   ByVal lpClassName As String, _
   ByVal nMaxCount As Long) As Long
   
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
(ByVal hWnd As Long, ByVal wIndx As Long) As Long

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" _
 (ByVal hWnd As Long) As Long
 
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
 (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
 
Private Declare Function EnumWindows Lib "user32" _
 (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)

Public Declare Sub GetCursorPos Lib "user32" (lpPoint As POINTAPI)

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "Kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Module32FirstW Lib "Kernel32" (ByVal hSnapshot As Long, ByRef uModule As Any) As Long
Private Declare Function Module32NextW Lib "Kernel32" (ByVal hSnapshot As Long, ByRef uModule As Any) As Long

Private Declare Function WriteProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Private Declare Function VirtualQueryEx& Lib "Kernel32" (ByVal hProcess As Long, lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long)

Private Declare Sub GetSystemInfo Lib "Kernel32" (lpSystemInfo As SYSTEM_INFO)
   
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long

Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long


Public moduleDictionary As Scripting.Dictionary
Public mainTibiaHandle As Scripting.Dictionary
Public objWMIService As Object
Public adrServerList_PortOffset As Long
Public MAXDATTILESpath As String
Public subTibiaVersionLong As Long
Public useAntiDDoS As Boolean
Public useFirewall As Boolean

Public lastFillAdrName As String
Public lastFillAdrValue As String
Public lastFillSize As String

Public Function QMemory_ReadNBytes(ByVal pid As Long, ByVal finalAddress As Long, ByRef Rbuff() As Byte) As Long
    Dim usize As Long
    Dim tibiaHandle As Long
    Dim readtotal As Long
    On Error GoTo goterr
    readtotal = 0
    usize = UBound(Rbuff) + 1
    If (usize < 1) Then
        Exit Function
    End If
    tibiaHandle = OpenProcess(PROCESS_VM_READ, 0, pid)
    ReadProcessMemory tibiaHandle, finalAddress, Rbuff(0), usize, readtotal
    CloseHandle (tibiaHandle)
    If (readtotal = 0) Then
        QMemory_ReadNBytes = -1
    Else
        QMemory_ReadNBytes = 0
    End If
    Exit Function
goterr:
    QMemory_ReadNBytes = -1
    Debug.Print ("Error at QMemory_ReadNBytes:" & Err.Description)
End Function
    
Public Function QMemory_ReadString(ByVal pid As Long, ByVal address As Long, Optional maxSize As Long = 2048) As String
    Dim msg_size As Long
    Dim msg_offset As Long
    Dim msg_start As Long
    Dim msg_lastc As Long
    Dim msg_refs As Long
    Dim allbytes() As Byte
    Dim b As Byte
    Dim res As String
    Dim i As Long
    Dim auxRes As Long
    res = ""
    msg_refs = QMemory_Read4Bytes(pid, address)
    If (msg_refs > 1000) Then
        ' This QString was moved to a different place
        QMemory_ReadString = QMemory_ReadString(pid, msg_refs, maxSize)
        Exit Function
    End If
    msg_size = QMemory_Read4Bytes(pid, address + 4) ' Size of the QString
    msg_offset = QMemory_Read4Bytes(pid, address + 12) ' Offset that we should use to find the String start
    msg_start = address + msg_offset
    If msg_size > maxSize Then ' Only read up to maxSize characters
        msg_size = maxSize
    End If
    msg_lastc = msg_size - 1
    If msg_lastc < 0 Then
        QMemory_ReadString = ""
        Exit Function
    End If
    ReDim allbytes((msg_size * 2) - 1)
    auxRes = QMemory_ReadNBytes(pid, msg_start, allbytes)
    If (auxRes = -1) Then
        QMemory_ReadString = ""
        Exit Function
    End If
    For i = 0 To msg_lastc
        res = res & Chr(allbytes(i * 2))
    Next i
    QMemory_ReadString = res
End Function

Public Function QMemory_WriteNBytes(ByVal pid As Long, ByVal finalAddress As Long, ByRef newValue() As Byte) As Long
    Dim tibiaHandle As Long
    Dim res As Long
    Dim lpNumberOfBytesWritten As Long
    Dim usize As Long
    lpNumberOfBytesWritten = 0
    On Error GoTo goterr
    usize = UBound(newValue) + 1
    tibiaHandle = OpenProcess(PROCESS_READ_WRITE_QUERY, 0, pid)
    If tibiaHandle = -1 Then
        QMemory_WriteNBytes = -1
        Exit Function
    End If
    res = WriteProcessMemory(tibiaHandle, finalAddress, newValue(0), usize, lpNumberOfBytesWritten)
    If (res = 1) Then
        CloseHandle (tibiaHandle)
        QMemory_WriteNBytes = 0
    Else
        CloseHandle (tibiaHandle)
        QMemory_WriteNBytes = -1
    End If
    Exit Function
goterr:
    QMemory_WriteNBytes = -1
End Function

Public Function QMemory_Write2Bytes(ByVal pid As Long, ByVal finalAddress As Long, newValue As Long) As Long
    Dim tibiaHandle As Long
    Dim res As Long
    Dim lpNumberOfBytesWritten As Long
    Dim Rbuff(1) As Byte
    lpNumberOfBytesWritten = 0
    On Error GoTo goterr
    Rbuff(0) = LowByteOfLong(newValue)
    Rbuff(1) = HighByteOfLong(newValue)
    tibiaHandle = OpenProcess(PROCESS_READ_WRITE_QUERY, 0, pid)
    res = WriteProcessMemory(tibiaHandle, finalAddress, Rbuff(0), 2, lpNumberOfBytesWritten)
    If (res = 1) Then
        CloseHandle (tibiaHandle)
        QMemory_Write2Bytes = 0
    Else
        CloseHandle (tibiaHandle)
        QMemory_Write2Bytes = -1
    End If
    Exit Function
goterr:
    QMemory_Write2Bytes = -1
End Function

'Public Function QMemory_WriteRSA(ByVal pid As Long, ByRef newHexKey As String) As Long
'    Dim realAddress As Long
'    Dim i As Long
'    Dim RSA_bytes() As Byte
'    Dim lastC As Long
'    Dim res As Long
'    Dim resF As Long
'    Dim writeChr As String
'    Dim byteChr As Byte
'    Dim resVirtual As Long
'    realAddress = ReadCurrentAddress(pid, adrRSAhex, -1, False)
'    Debug.Print "Current address of RSA= " & Hex(realAddress)
'    lastC = Len(newHexKey) - 1
'    ReDim RSA_bytes(lastC)
'    For i = 0 To lastC
'        writeChr = Mid$(newHexKey, i + 1, 1)
'        byteChr = ConvStrToByte(writeChr)
'        RSA_bytes(i) = byteChr
'    Next i
'    resVirtual = VirtualProtectEx(pid, realAddress, 256, PAGE_EXECUTE_READWRITE, PAGE_READONLY)
'    res = QMemory_WriteNBytes(pid, realAddress, RSA_bytes)
'    If (res = 0) Then
'        resF = 0
'    Else
'        resF = -1
'    End If
'    resVirtual = VirtualProtectEx(pid, realAddress, 256, PAGE_READONLY, PAGE_EXECUTE_READWRITE)
'    QMemory_WriteRSA = resF
'    Exit Function
'End Function

Public Function ModifyQString(ByVal pid As Long, ByVal address As Long, ByRef newText As String) As Long
        Dim msg_maxsize As Long
        Dim msg_offset As Long
        Dim new_size As Long
        Dim msg_start As Long
        Dim res As Long
        Dim allbytes() As Byte
        Dim i As Long
        On Error GoTo goterr
        new_size = Len(newText)
        msg_offset = QMemory_Read4Bytes(pid, address + 12)
        msg_maxsize = QMemory_Read4Bytes(pid, address + 8)
        If (new_size > msg_maxsize) Then
            ModifyQString = -1
            Exit Function
        End If
        msg_start = address + msg_offset
        ReDim allbytes((new_size * 2) - 1)
        For i = 0 To new_size - 1
            allbytes(i * 2) = Asc(Mid(newText, i + 1, 1))
            allbytes(1 + (i * 2)) = &H0
        Next i
        res = QMemory_WriteNBytes(pid, msg_start, allbytes)
        If (res = 0) Then
            res = QMemory_Write4Bytes(pid, address + 4, new_size)
            ModifyQString = res
        Else
            res = -1
        End If
        ModifyQString = res
        Exit Function
goterr:
        ModifyQString = -1
    End Function
    
    
 
Public Function QMemory_ReadDouble(ByVal pid As Long, ByVal finalAddress As Long) As Double
    Dim Rbuff(7) As Byte
    Dim d As Double
    Dim auxRes As Long
    auxRes = QMemory_ReadNBytes(pid, finalAddress, Rbuff)
    If (auxRes = -1) Then
        QMemory_ReadDouble = -1
        Exit Function
    End If
    CopyMemory d, Rbuff(0), LenB(d)
    QMemory_ReadDouble = d
    Exit Function
goterr:
    QMemory_ReadDouble = -1
End Function

    
Public Function QMemory_Read4Bytes(ByVal pid As Long, ByVal finalAddress As Long) As Long
    Dim res As Long
    Dim tibiaHandle As Long
    On Error GoTo goterr
    tibiaHandle = OpenProcess(PROCESS_VM_READ, 0, pid)
    ReadProcessMemory tibiaHandle, finalAddress, res, 4, 0
    CloseHandle (tibiaHandle)
    QMemory_Read4Bytes = res
    Exit Function
goterr:
    QMemory_Read4Bytes = -1
End Function

Public Function QMemory_Read2Bytes(ByVal pid As Long, ByVal finalAddress As Long) As Long
    Dim Rbuff(1) As Byte
    Dim tibiaHandle As Long
    On Error GoTo goterr
    tibiaHandle = OpenProcess(PROCESS_VM_READ, 0, pid)
    ReadProcessMemory tibiaHandle, finalAddress, Rbuff(0), 2, 0
    CloseHandle (tibiaHandle)
    QMemory_Read2Bytes = GetTheLong(Rbuff(0), Rbuff(1))
    Exit Function
goterr:
    QMemory_Read2Bytes = -1
End Function

Public Function QMemory_Read1Byte(ByVal pid As Long, ByVal finalAddress As Long) As Byte
    Dim Rbuff As Byte
    Dim tibiaHandle As Long
    On Error GoTo goterr
    tibiaHandle = OpenProcess(PROCESS_VM_READ, 0, pid)
    ReadProcessMemory tibiaHandle, finalAddress, Rbuff, 1, 0
    CloseHandle (tibiaHandle)
    QMemory_Read1Byte = Rbuff
    Exit Function
goterr:
    QMemory_Read1Byte = &HFF
End Function
    
Public Function QMemory_Write4Bytes(ByVal pid As Long, ByVal finalAddress As Long, ByVal newValue As Long) As Long
    Dim tibiaHandle As Long
    Dim res As Long
    Dim lpNumberOfBytesWritten As Long
    lpNumberOfBytesWritten = 0
    On Error GoTo goterr
    tibiaHandle = OpenProcess(PROCESS_READ_WRITE_QUERY, 0, pid)
    res = WriteProcessMemory(tibiaHandle, finalAddress, newValue, 4, lpNumberOfBytesWritten)
    If (res = 1) Then
        CloseHandle (tibiaHandle)
        QMemory_Write4Bytes = 0
    Else
        CloseHandle (tibiaHandle)
        QMemory_Write4Bytes = -1
    End If
    Exit Function
goterr:
    QMemory_Write4Bytes = -1
End Function





Private Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
 Dim Title As String
 Dim r As Long
 
 r = GetWindowTextLength(hWnd)
 Title = Space(r)
 GetWindowText hWnd, Title, r + 1
 
 'Add to type array
 ColCounter = ColCounter + 1
 ReDim Preserve HandleTextCollection(ColCounter)
 HandleTextCollection(ColCounter).Window_Handle = hWnd
 HandleTextCollection(ColCounter).Window_Title = Title
 
 EnumWindowsProc = True
 
End Function
Private Sub EnumAllWindows()
 
 ColCounter = 0
 ReDim HandleTextCollection(0)

 EnumWindows AddressOf EnumWindowsProc, ByVal 0&
 
 'At this point, HandleTextCollection() array holds the handles and window titles
 'of all windows enumerated

End Sub
 Private Function IsNothing(ByRef objParm As Object) As Boolean
        IsNothing = IIf(objParm Is Nothing, True, False)
    End Function
    
Public Function CheckIfEnumDone()
 On Error GoTo goterr
 If (HandleTextCollection(0).Window_Handle = &H0) Then
 CheckIfEnumDone = True
 Else
 CheckIfEnumDone = True
 End If
 Exit Function
goterr:
 CheckIfEnumDone = False
End Function

Public Function Get_MainWindowHandle_from_ProcessID_and_class(ByVal AppPID As Long, ByRef expectedClass As String, _
  Optional ByVal doNewEnum As Boolean = True) As Long
    Dim AppPID_HWND() As HWND_TEXT 'Will store only handles related to AppPID
    Dim CNT As Long
    Dim HND As Long
    Dim n As Long
    Dim r As Long
    Dim i As Long
    Dim TaskID As Long
    Dim TheHandle As Long
    Dim test_thread_id As Long
    Dim testPID As Long
    
    TheHandle = -1
    If (doNewEnum = True) Then
       Call EnumAllWindows ' Do Enum of all current windows
    ElseIf CheckIfEnumDone() = False Then
       Call EnumAllWindows ' enum was not done yet. An initial enum is mandatory.
    End If
    
    CNT = 0
    ReDim AppPID_HWND(0)
    For i = 1 To UBound(HandleTextCollection())
        HND = HandleTextCollection(i).Window_Handle
        n = GetWindowThreadProcessId(HND, TaskID)
        If TaskID = AppPID Then
            'Handle matches AppPID
            CNT = CNT + 1
            ReDim Preserve AppPID_HWND(CNT)
            AppPID_HWND(CNT).Window_Handle = HND
            AppPID_HWND(CNT).Window_Title = HandleTextCollection(i).Window_Title
        End If
    Next i
    'Search through all the handles related to AppPID
    For i = 1 To UBound(AppPID_HWND())
        
            If GetWindowClass(AppPID_HWND(i).Window_Handle) = expectedClass Then
                TheHandle = AppPID_HWND(i).Window_Handle
                Exit For
            End If
      
    Next i
    Get_MainWindowHandle_from_ProcessID_and_class = TheHandle
End Function


Private Function GetWindowClass(ByVal hWnd As Long) As String
  Dim sClass As String
  If hWnd = 0 Then
    GetWindowClass = ""
  Else
    sClass = Space$(256)
    GetClassName hWnd, sClass, 255
    GetWindowClass = Left$(sClass, InStr(sClass, vbNullChar) - 1)
  End If
End Function

Public Sub GetAllBaseAddressesAndRegionSizes(ByRef expectedName As String, ByRef expectedClass As String)
    'Debug.Print "GetAllBaseAddressesAndRegionSizes called"
    Dim procmodule_name As String
    Dim procmodule_base As Long
    'Dim procmodule_size As Long ' << size not needed for now
    Dim ubproc As Long
    Dim mainWindowHandle As Long
    Dim currentPID As Long
    Dim i As Long
    On Error GoTo goterr
    Dim items As Object
    Dim item As Object
    'Dim count As Long
    Dim uModule As MODULEENTRY32W, lModuleSnapshot&
    If (moduleDictionary Is Nothing) Then
        Set moduleDictionary = New Scripting.Dictionary
    End If
    ' EnumAllWindows will obtain a snapshoot of all current windows references.
    Call EnumAllWindows ' Here we will only call  EnumAllWindows once, to save time.
    Set items = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & expectedName & ".exe'", , 32)
    For Each item In items
        currentPID = item.ProcessID
        
        ' Then we obtain mainWindowHandle from Proccess ID
        ' Using the optional parameter with value = False
        ' we save time because it will not repeat EnumAllWindows
        mainWindowHandle = Get_MainWindowHandle_from_ProcessID_and_class(currentPID, expectedClass, False)
        If (Not (mainWindowHandle = 0)) Then
            'count = count + 1
            'Debug.Print "Found Tibia PID: " & CStr(currentPID) & " main handle = " & CStr(mainWindowHandle)
            mainTibiaHandle(currentPID) = mainWindowHandle
            lModuleSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, currentPID)
            uModule.dwSize = LenB(uModule)
            If lModuleSnapshot > 0 Then
                If Module32FirstW(lModuleSnapshot, uModule) <> 0 Then
                    Do
                        procmodule_base = uModule.modBaseAddr
                        procmodule_name = Left$(uModule.szModule, InStr(uModule.szModule, Chr(0)) - 1)
                        ' procmodule_size = uModule.modBaseSize
                        moduleDictionary(procmodule_name & CStr(currentPID)) = procmodule_base
                        'Debug.Print procmodule_name & CStr(currentPID) & "=" & Hex(procmodule_base)
                    Loop Until (Module32NextW(lModuleSnapshot, uModule) = 0)
                End If
            End If
        End If
    Next
   ' Debug.Print "Total clients found = " & count
    Exit Sub
goterr:
    Debug.Print ("Error: Unexpected error - " & Err.Description)
End Sub



Public Function GetTibiaPIDs(ByRef expectedName As String, ByRef expectedClass As String, _
ByRef CurrentTibiaPids() As Long) As Long
    Dim ubproc As Long
    Dim i As Long
    On Error GoTo goterr
    Dim items As Object
    Dim item As Object
    Dim last As Long
    ReDim CurrentTibiaPids(0)
    CurrentTibiaPids(0) = -1
    last = -1
    Dim uModule As MODULEENTRY32W, lModuleSnapshot&
    Set items = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & expectedName & ".exe'", , 32)
    For Each item In items
        last = last + 1
        ReDim Preserve CurrentTibiaPids(last)
        CurrentTibiaPids(last) = item.ProcessID
        'Debug.Print "Found Tibia PID: " & CStr(item.ProcessID)
    Next
    GetTibiaPIDs = last + 1
    'Debug.Print "Total clients found = " & (last + 1)
    Exit Function
goterr:
    Debug.Print ("Error: Unexpected error - " & Err.Description)
     CurrentTibiaPids(0) = -1
    GetTibiaPIDs = 0
End Function

Public Function arrayToString(ByRef arr() As Byte) As String
    Dim isize As Long
    Dim res As String
    Dim i As Long
    On Error GoTo goterr
    isize = UBound(arr)
    res = GoodHex(arr(0))
    For i = 1 To isize
        res = res & " " & GoodHex(arr(i))
    Next i
    arrayToString = res
    Exit Function
goterr:
    arrayToString = ""
End Function



Private Sub fillCollectionDictionary(ByRef pid As Long, ByVal adrCurrentItem As Long, ByVal adrSTARTER_ITEM As Long, _
                                     ByRef dict As Scripting.Dictionary, _
                                     ByRef totalItems As Long, ByVal currentDepth As Long, _
                                     ByRef maxDepth As Long, ByRef bytesPerElement As Long, _
                                     Optional ByRef maxValidKeyID As Long = -1, _
                                     Optional ByVal addBaseAddress As Boolean = False)
    On Error GoTo goterr
        Dim Id As Long
        Dim auxRes As Long
        If (adrCurrentItem = 0) Then
            Dim errMsg As String
            errMsg = "Critical fail at fillCollectionDictionary. Tibia version " & CStr(TibiaVersionLong) & " Please report to daniel@blackdtools.com Failed to gather collection for address " & lastFillAdrName & " = " & lastFillAdrValue & " with size " & CStr(lastFillSize)
            Debug.Print errMsg
            LogOnFile "errors.txt", errMsg
            End
        End If
        Id = QMemory_Read4Bytes(pid, adrCurrentItem + &H10)
        If (maxValidKeyID > -1) Then
            If (Id > maxValidKeyID) Then
                Exit Sub
            End If
        End If
        If (dict.Exists(Id) = False) Then
            Dim tmp() As Byte
            If (addBaseAddress) Then
                ReDim tmp(4 + bytesPerElement - 1)
            Else
                ReDim tmp(bytesPerElement - 1)
            End If
            auxRes = QMemory_ReadNBytes(pid, adrCurrentItem, tmp)
            If (auxRes = -1) Then
                Exit Sub
            End If
            If (addBaseAddress) Then
                CopyMemory tmp(bytesPerElement), adrCurrentItem, 4
            End If
            dict(Id) = tmp
            ' Debug.Print ("Key #" & CStr(id) & " found at " & Hex(adrCurrentItem))
        End If
        If (currentDepth < maxDepth) Then
            Dim p0, p1, p2 As Long
            p0 = QMemory_Read4Bytes(pid, adrCurrentItem)
            p1 = QMemory_Read4Bytes(pid, adrCurrentItem + 4)
            p2 = QMemory_Read4Bytes(pid, adrCurrentItem + 8)
            If (Not (p0 = adrSTARTER_ITEM)) Then
                fillCollectionDictionary pid, p0, adrSTARTER_ITEM, dict, totalItems, currentDepth + 1, maxDepth, bytesPerElement, maxValidKeyID, addBaseAddress
            End If
            If (Not (p1 = adrSTARTER_ITEM)) Then
                fillCollectionDictionary pid, p1, adrSTARTER_ITEM, dict, totalItems, currentDepth + 1, maxDepth, bytesPerElement, maxValidKeyID, addBaseAddress
            End If
            If (Not (p2 = adrSTARTER_ITEM)) Then
                fillCollectionDictionary pid, p2, adrSTARTER_ITEM, dict, totalItems, currentDepth + 1, maxDepth, bytesPerElement, maxValidKeyID, addBaseAddress
            End If
        End If
        Exit Sub
goterr:
        Debug.Print ("Something failed: " + Err.Description)

End Sub

Public Sub ReadTibia11Collection(ByRef pid As Long, adrPath As AddressPath, _
                                      ByRef bytesPerElement As Long, _
                                      ByRef dict As Scripting.Dictionary, _
                                      Optional ByVal directAddress As Long = -1, _
                                      Optional ByVal customDepth As Long = -1, _
                                      Optional ByVal maxValidKeyID As Long = -1, _
                                      Optional ByVal addBaseAddress As Boolean = False)
                                    
    Dim adrCOLLECTION_START As Long
    Dim totalItems As Long
    Dim adrSTARTER_ITEM As Long
 
    Dim maxDepth As Long
    Dim p0, p1, p2 As Long
    Set dict = New Scripting.Dictionary
    If (directAddress > -1) Then
        adrCOLLECTION_START = directAddress
    Else
        adrCOLLECTION_START = ReadCurrentAddress(pid, adrPath, -1, False)
    End If
    If (customDepth > -1) Then
        maxDepth = customDepth
    Else
        totalItems = QMemory_Read4Bytes(pid, adrCOLLECTION_START + 4)
        If (totalItems = 0) Then
            Exit Sub
        End If
        maxDepth = Math.Round(Math.Sqr(totalItems))
    End If
    lastFillAdrName = adrPath.name
    lastFillAdrValue = adrPath.rawString
    lastFillSize = totalItems
    adrSTARTER_ITEM = QMemory_Read4Bytes(pid, adrCOLLECTION_START)
    p0 = QMemory_Read4Bytes(pid, adrSTARTER_ITEM)
    p1 = QMemory_Read4Bytes(pid, adrSTARTER_ITEM + 4)
    p2 = QMemory_Read4Bytes(pid, adrSTARTER_ITEM + 8)
    fillCollectionDictionary pid, p0, adrSTARTER_ITEM, dict, totalItems, 0, maxDepth, bytesPerElement, maxValidKeyID, addBaseAddress
    fillCollectionDictionary pid, p1, adrSTARTER_ITEM, dict, totalItems, 0, maxDepth, bytesPerElement, maxValidKeyID, addBaseAddress
    fillCollectionDictionary pid, p2, adrSTARTER_ITEM, dict, totalItems, 0, maxDepth, bytesPerElement, maxValidKeyID, addBaseAddress
End Sub


Private Function fillCollectionDictionaryMIN(ByRef pid As Long, ByVal adrCurrentItem As Long, _
                                     ByVal adrPrev As Long, _
                                     ByRef dict As Scripting.Dictionary, _
                                     ByVal currentDepth As Long, _
                                     ByRef maxDepth As Long, ByRef adrRoot As Long) As Long
    On Error GoTo goterr
        Dim Id As Long
        Dim c0, c1, c2 As Long
        Dim Count As Long
        Count = 1
        Id = QMemory_Read4Bytes(pid, adrCurrentItem + &H10)
        If (dict.Exists(Id) = False) Then
            dict(Id) = adrCurrentItem
            If (currentDepth < maxDepth) Then
                  Dim p0, p1, p2 As Long
                  p0 = QMemory_Read4Bytes(pid, adrCurrentItem)
                  p1 = QMemory_Read4Bytes(pid, adrCurrentItem + 4)
                  p2 = QMemory_Read4Bytes(pid, adrCurrentItem + 8)
                  If Not ((p0 = adrPrev) Or (p0 = adrRoot)) Then
                    c0 = fillCollectionDictionaryMIN(pid, p0, adrCurrentItem, dict, currentDepth + 1, maxDepth, adrRoot)
                    Count = Count + c0
                  End If
                  If Not ((p1 = adrPrev) Or (p1 = adrRoot)) Then
                    c1 = fillCollectionDictionaryMIN(pid, p1, adrCurrentItem, dict, currentDepth + 1, maxDepth, adrRoot)
                    Count = Count + c1
                  End If
                  If Not ((p2 = adrPrev) Or (p2 = adrRoot)) Then
                   c2 = fillCollectionDictionaryMIN(pid, p2, adrCurrentItem, dict, currentDepth + 1, maxDepth, adrRoot)
                   Count = Count + c2
                  End If
            End If
        End If
        fillCollectionDictionaryMIN = Count
        Exit Function
goterr:
        fillCollectionDictionaryMIN = 0
        Debug.Print ("Something failed: " + Err.Description)
End Function
Public Sub ReadTibia11CollectionMIN(ByRef pid As Long, adrPath As AddressPath, _
                                      ByRef dict As Scripting.Dictionary)
                                    
    Dim adrCOLLECTION_START As Long
    Dim totalItems As Long
    Dim adrSTARTER_ITEM As Long
    Dim maxDepth As Long
    Dim p0, p1, p2 As Long
    Dim item As Variant
    Dim iterCount As Long
    Dim c0, c1, c2 As Long
    Dim adrRoot As Long
    iterCount = 0
    Set dict = New Scripting.Dictionary
    adrCOLLECTION_START = ReadCurrentAddress(pid, adrPath, -1, False)
    totalItems = QMemory_Read4Bytes(pid, adrCOLLECTION_START + 4)
    If (totalItems = 0) Then
        Exit Sub
    End If
    maxDepth = totalItems
    adrSTARTER_ITEM = QMemory_Read4Bytes(pid, adrCOLLECTION_START)
    adrRoot = adrSTARTER_ITEM
    p0 = QMemory_Read4Bytes(pid, adrSTARTER_ITEM)
    p1 = QMemory_Read4Bytes(pid, adrSTARTER_ITEM + 4)
    p2 = QMemory_Read4Bytes(pid, adrSTARTER_ITEM + 8)
    c0 = fillCollectionDictionaryMIN(pid, p0, adrSTARTER_ITEM, dict, 0, maxDepth, adrRoot)
    c1 = fillCollectionDictionaryMIN(pid, p1, adrSTARTER_ITEM, dict, 0, maxDepth, adrRoot)
    c2 = fillCollectionDictionaryMIN(pid, p2, adrSTARTER_ITEM, dict, 0, maxDepth, adrRoot)
    iterCount = c0 + c1 + c2
    Debug.Print "Collection size " & CStr(totalItems) & " took " & CStr(iterCount) & " iterations"
    '  For Each item In dict
    '   Debug.Print CStr(item) & " (" & CStr(Hex(item)) & ") found at " & CStr(Hex(dict(item)))
    '  Next item
End Sub
' Works ok. Just need to read full collection
'Public Function FindCollectionItemByKey(ByRef pid As Long, adrPath As AddressPath, ByRef keyToSearch As Long, ByRef dict As Scripting.Dictionary, Optional ByVal reloadDictionary As Boolean = True) As Long
'    Dim res As Long
'    If reloadDictionary Then
'        Set dict = New Scripting.Dictionary
'        ReadTibia11CollectionMIN pid, adrPath, dict
'    End If
'
'    If (dict.Exists(keyToSearch) = False) Then
'        Debug.Print "This key was not in this collection: " & CStr(keyToSearch)
'        res = -1
'    Else
'        res = dict(keyToSearch)
'    End If
'    FindCollectionItemByKey = res
'End Function


Public Function FindCollectionItemByKey(ByRef pid As Long, adrPath As AddressPath, ByRef keyToSearch As Long) As Long
    Dim res As Long
    Dim adrCOLLECTION_START As Long
    Dim totalItems As Long
    Dim pLeft As Long
    Dim pRight As Long
    Dim val0 As Long
    Dim val1 As Long
    Dim val2 As Long
    Dim maxDepth As Long
    Dim adrSTARTER_ITEM As Long
    Dim p(5) As Long
    Dim isGoal As Boolean
    Dim isFail As Boolean
    Dim currentDepth As Long
    Const rootval As Long = -1
    currentDepth = 1
    adrCOLLECTION_START = ReadCurrentAddress(pid, adrPath, -1, False)
    totalItems = QMemory_Read4Bytes(pid, adrCOLLECTION_START + 4)
    If (totalItems = 0) Then
        FindCollectionItemByKey = -1
        Exit Function
    End If
    maxDepth = 1 + (totalItems / 2)
    adrSTARTER_ITEM = QMemory_Read4Bytes(pid, adrCOLLECTION_START)
    p(0) = QMemory_Read4Bytes(pid, adrSTARTER_ITEM)
    p(1) = QMemory_Read4Bytes(pid, adrSTARTER_ITEM + 4)
    p(2) = QMemory_Read4Bytes(pid, adrSTARTER_ITEM + 8)
    p(3) = p(0)
    p(4) = p(1)
    p(5) = p(2)
    q_nextIteration pid, adrSTARTER_ITEM, keyToSearch, p, pLeft, pRight, isGoal, isFail
    If (isGoal) Then
        res = pLeft
    End If
    If (isFail) Then
        currentDepth = maxDepth + 1
    End If
    If (isGoal = False) And (isFail = False) Then
        'Debug.Print "..."
        Do
            currentDepth = currentDepth + 1
            p(0) = QMemory_Read4Bytes(pid, pLeft)
            p(1) = QMemory_Read4Bytes(pid, pLeft + 4)
            p(2) = QMemory_Read4Bytes(pid, pLeft + 8)
            p(3) = QMemory_Read4Bytes(pid, pRight)
            p(4) = QMemory_Read4Bytes(pid, pRight + 4)
            p(5) = QMemory_Read4Bytes(pid, pRight + 8)
            q_nextIteration pid, adrSTARTER_ITEM, keyToSearch, p, pLeft, pRight, isGoal, isFail
            If (isGoal) Then
                res = pLeft
                Exit Do
            End If
            If (isFail) Then
                currentDepth = maxDepth + 1
                Exit Do
            End If
            ' Debug.Print "..."
        Loop Until currentDepth > maxDepth
    End If
    If (currentDepth > maxDepth) Then
      ' Debug.Print "WARNING at FindCollectionItemByKey: Key not found (With size=" & CStr(totalItems) & ") Key = " & CStr(keyToSearch) & " (" & CStr(Hex(keyToSearch)) & ")"
      FindCollectionItemByKey = -1
    Else
      ' Debug.Print "(With size=" & CStr(totalItems) & ") Key " & CStr(keyToSearch) & " (" & CStr(Hex(keyToSearch)) & ") found after " & CStr(currentDepth) & " iterations at " & CStr(Hex(res))
      FindCollectionItemByKey = res
    End If
End Function

Private Sub q_nextIteration(ByRef pid As Long, ByRef adrSTARTER_ITEM As Long, _
 ByRef keyToSearch As Long, ByRef p() As Long, _
 ByRef pLeft As Long, ByRef pRight As Long, _
 ByRef isGoal As Boolean, ByRef isFail As Boolean)
    Const rootval_min As Long = -2147483648#
    Const rootval_max As Long = 2147483647
    Dim val_min(5) As Long
    Dim val_max(5) As Long
    Dim best_min_v As Long
    Dim best_min_i As Long
    Dim best_max_v As Long
    Dim best_max_i As Long
    Dim i As Long
    isGoal = False
    isFail = False
    best_min_i = 0
    best_min_v = rootval_min
    best_max_i = 0
    best_max_v = rootval_max
    For i = 0 To 5
        If (p(i) = adrSTARTER_ITEM) Then
            val_min(i) = rootval_min
            val_max(i) = rootval_max
        Else
            val_min(i) = QMemory_Read4Bytes(pid, p(i) + &H10)
            val_max(i) = val_min(i)
        End If
    Next i
    'Debug.Print "q_nextIteration options: " & CStr(val(0)) & "," & CStr(val(1)) & "," & CStr(val(2)) & "," & CStr(val(3)) & "," & CStr(val(4)) & "," & CStr(val(5))
    For i = 0 To 5
        If (val_min(i) = keyToSearch) Then
               ' Debug.Print "Goal found at q_nextIteration p(" & CStr(i) & ")"
                isGoal = True
                pLeft = p(i)
                pRight = p(i)
                Exit Sub
        End If
        ' Pick best left
        If (val_min(i) > best_min_v) And (val_min(i) <= keyToSearch) Then
            best_min_v = val_min(i)
            best_min_i = i
        End If
        ' Pick best right
        If (val_max(i) < best_max_v) And (val_max(i) >= keyToSearch) Then
            best_max_v = val_max(i)
            best_max_i = i
        End If
    Next i
    pLeft = p(best_min_i)
    pRight = p(best_max_i)
    'Debug.Print "Next iteration will search between " & CStr(val_min(best_min_i)) & " and " & CStr(val_max(best_max_i))
    If (val_min(best_min_i) > keyToSearch) Or (val_max(best_max_i) < keyToSearch) Then
        isFail = True
    End If
End Sub

Public Function TibiaClientConnectionStatus(ByRef pid As Long) As Long
       Dim auxAdr As Long
       Dim auxVal As Double
       Dim pixels As Long
       Dim currentCharName As String
       currentCharName = ReadCurrentCharName(pid)
       If Not (currentCharName = "") Then
           TibiaClientConnectionStatus = CTE_GAME_CONNECTED ' 3 - Game connected
           Exit Function
       End If
       auxAdr = ReadCurrentAddress(pid, adrSelectedItem_height, -1, False)
       If (auxAdr = -1) Then
           TibiaClientConnectionStatus = CTE_NOT_CONNECTED
           Exit Function
       End If
       auxVal = QMemory_ReadDouble(pid, auxAdr)
       pixels = Math.Round(auxVal)
       Select Case (pixels)
           Case 18
               TibiaClientConnectionStatus = CTE_NOT_CONNECTED ' 0 - Not connected
               Exit Function
           Case 16
               TibiaClientConnectionStatus = CTE_LOGIN_CHAR_SELECTION ' 1 - User at Character Selection
               Exit Function
           Case 100
               TibiaClientConnectionStatus = CTE_CONNECTING ' 2 - Connecting... (to characer selection or to game)
               Exit Function
           Case 14
               TibiaClientConnectionStatus = CTE_GAME_CONNECTED ' 3 - Game connected
               Exit Function
           Case Else
               Debug.Print ("status code =" & CStr(pixels))
               TibiaClientConnectionStatus = CTE_NOT_CONNECTED
               Exit Function
       End Select
End Function

Public Function ReadCurrentCharName(ByRef pid As Long) As String
    Dim auxAdr As Long
    Dim strRes As String
    auxAdr = ReadCurrentAddress(pid, adrSelectedCharName, -1, True)
    strRes = QMemory_ReadString(pid, auxAdr)
    ReadCurrentCharName = ""
End Function

Public Function GetProcessIdByAdrConnected_TibiaQ() As Long
    Dim tibia_pids() As Long
    Dim totalPids As Long
    Dim i As Long
    Dim connectionStatus As Long
    Dim foundCount As Long
    Dim res As Long
    foundCount = 0
    GetAllBaseAddressesAndRegionSizes tibiamainname, tibiaclassname
    totalPids = GetTibiaPIDs(tibiamainname, tibiaclassname, tibia_pids)
    If (totalPids <= 0) Then
        GetProcessIdByAdrConnected_TibiaQ = -1
        Exit Function
    End If
    For i = 0 To totalPids - 1
        connectionStatus = TibiaClientConnectionStatus(tibia_pids(i))
        If (connectionStatus = CTE_CONNECTING) Then
            foundCount = foundCount + 1
            res = tibia_pids(i)
        End If
    Next i
    If (foundCount = 1) Then
        GetProcessIdByAdrConnected_TibiaQ = res
    ElseIf (foundCount > 1) Then
        GetProcessIdByAdrConnected_TibiaQ = -2
    Else
        If (totalPids = 1) Then
            GetProcessIdByAdrConnected_TibiaQ = tibia_pids(0) ' can't be other pid
        Else
            GetProcessIdByAdrConnected_TibiaQ = -1
        End If
    End If
End Function

Public Function QMemory_ReadStringP(ByRef pid As Long, adrPath As AddressPath, Optional maxSize As Long = 2048) As String
   Dim adrAux As Long

   adrAux = ReadCurrentAddress(pid, adrPath, -1, True)
   If (adrAux) = -1 Then
        QMemory_ReadStringP = -1
   Else
        QMemory_ReadStringP = QMemory_ReadString(pid, adrAux, maxSize)
   End If
End Function



Public Function ReadTibia11CharList(ByVal pid As Long, ByRef res() As TibiaCharListEntry) As Long
    Dim i As Integer
    Dim tmpElement As TibiaCharListEntry
    Dim adrCOLLECTION_START As Long
    Dim adrCharList As Long
    Dim adrType As Long
    Dim adrCharListStart As Long
    Dim resSize As Long
    Dim bytesPerElement As Long
    Dim lastI As Long
    Dim tmpAdr As Long
    adrCOLLECTION_START = ReadCurrentAddress(pid, adrServerList_CollectionStart, -1, False)
    adrCharList = QMemory_Read4Bytes(pid, adrCOLLECTION_START + 8)
    'Debug.Print ("Charlist at " & Hex(adrCharList))
    adrType = QMemory_Read4Bytes(pid, adrCharList)
    If (adrType = -1) Then
        ReadTibia11CharList = -1
        Exit Function
    End If
    If (adrType > 1000) Then
        adrCharList = QMemory_Read4Bytes(pid, adrType)
    End If
    resSize = QMemory_Read4Bytes(pid, adrCharList + &HC)
    adrCharListStart = adrCharList + &H10
    If (resSize = 0) Then
        ReadTibia11CharList = -1
        Exit Function
    End If
    bytesPerElement = resSize * 4
    Dim tmp() As Byte
    ReDim tmp(bytesPerElement - 1)
    QMemory_ReadNBytes pid, adrCharListStart, tmp
    Dim tamStruct As Long
    tamStruct = 60
    Dim resStruct() As Byte
    ReDim res(resSize - 1)
    lastI = resSize - 1
    For i = 0 To resSize - 1
        tmpElement.Id = i
        tmpElement.entry_address = BitConverter_ToInt32(tmp, 4 * i)
        ReDim resStruct(tamStruct - 1)
        QMemory_ReadNBytes pid, tmpElement.entry_address, resStruct
        tmpAdr = QMemory_Read4Bytes(pid, tmpElement.entry_address + &H10)
        tmpElement.name_address = tmpAdr
        tmpElement.name = QMemory_ReadString(pid, tmpAdr)
        tmpAdr = QMemory_Read4Bytes(pid, tmpElement.entry_address + &H14)
        tmpElement.server = QMemory_ReadString(pid, tmpAdr)
        res(i) = tmpElement
    Next i
    ReadTibia11CharList = resSize - 1
End Function

Public Sub RestoreAllCharlists()
    Dim tibia_pids() As Long
    Dim totalPids As Long
    totalPids = GetTibiaPIDs(tibiamainname, tibiaclassname, tibia_pids)
    LastNumTibiaClients = totalPids
    If totalPids = 0 Then
        ' Debug.Print "Tibia 11 clients not found (0)"
        Exit Sub
    End If
    Dim serverList() As TibiaServerEntry
    Dim i As Long
    Dim j As Long
    Dim readResult As Long
    Dim a As Long
    Dim strServerName As String
    Dim strURLtrans As String
    Dim strDefaultServer As String
    Dim realDomain As String
    Dim realPort As Long
  
    For j = 0 To UBound(tibia_pids)
        readResult = ReadTibia11ServerList(tibia_pids(j), adrServerList_CollectionStart, serverList)
        If readResult = 0 Then
            For i = 0 To UBound(serverList)
                strServerName = serverList(i).name
                realDomain = GetGameServerDOMAIN1(strServerName)
                realPort = GetGameServerPort(strServerName)
                If Not (realDomain = "") Then
                    'Debug.Print "Restoring " & Hex(serverList(i).url_adr) & " (" & strServerName & ") to " & realDomain & ":" & realPort
                    ModifyQString tibia_pids(j), serverList(i).url_adr, realDomain
                    QMemory_Write4Bytes tibia_pids(j), serverList(i).port_adr, realPort
                End If
                If (serverList(i).url2_adr > 0) Then
                    realDomain = GetGameServerDOMAIN2(strServerName)
                    If Not (realDomain = "") Then
                        'Debug.Print "Restoring " & Hex(serverList(i).url_adr) & " (" & strServerName & ") to " & realDomain & ":" & realPort
                        ModifyQString tibia_pids(j), serverList(i).url2_adr, realDomain
                    
                    End If
                End If
            Next i
            Debug.Print "PID " & CStr(tibia_pids(j)) & ": RESTORED server list."
         End If
    Next j
    
End Sub


Public Function ReadTibia11ServerList(ByRef pid As Long, ByRef adrPath As AddressPath, _
 ByRef res() As TibiaServerEntry, Optional ByVal stopIfPort As Long = -1) As Long
    Dim tmpRes As Scripting.Dictionary
    Dim resSize As Long
    Dim tmpElement As TibiaServerEntry
    Dim i As Long
    Dim item As Variant
    Dim Key As Long
    Dim val() As Byte
    Dim adrCOLLECTION_START As Long
    Dim adrSTARTER_ITEM As Long
    Dim totalItems As Long
    Dim auxAdr As Long
    Dim firstChar As String
    Dim currentPort As Long
    Dim cte_bytesPerRegister As Long
    Const altIPpos As Long = &H20
    cte_bytesPerRegister = adrServerList_PortOffset + 4
    On Error GoTo goterr
    If stopIfPort = -1 Then
        ReadTibia11Collection pid, adrPath, cte_bytesPerRegister, tmpRes, , , , True
    Else
        adrCOLLECTION_START = ReadCurrentAddress(pid, adrPath, -1, False)
        If (adrCOLLECTION_START = -1) Then
            ReadTibia11ServerList = -1
            Exit Function
        End If
        totalItems = QMemory_Read4Bytes(pid, adrCOLLECTION_START + 4)
        If (totalItems = 0) Then
            ReadTibia11ServerList = -1
            Exit Function
        End If
        adrSTARTER_ITEM = QMemory_Read4Bytes(pid, adrCOLLECTION_START)
        If (auxAdr = -1) Then
            ReadTibia11ServerList = -1
            Exit Function
        End If
        
        auxAdr = QMemory_Read4Bytes(pid, adrSTARTER_ITEM)
        If (auxAdr = -1) Then
            ReadTibia11ServerList = -1
            Exit Function
        End If
        
        currentPort = QMemory_Read4Bytes(pid, auxAdr + &H20)
        If (stopIfPort = 7171) Then
            auxAdr = QMemory_Read4Bytes(pid, auxAdr + &H1C)
            If (auxAdr = -1) Then
                ReadTibia11ServerList = -1
                Exit Function
            End If
            firstChar = QMemory_ReadString(pid, auxAdr, 1)
            If (firstChar = "1") Then
                ReadTibia11ServerList = -2
                Exit Function
            End If
        Else
            If (currentPort = stopIfPort) Then
                ReadTibia11ServerList = -2
                Exit Function
            End If
        End If
        ReadTibia11Collection pid, adrPath, cte_bytesPerRegister, tmpRes, adrCOLLECTION_START, , , True
    End If
    resSize = tmpRes.Count
    If (resSize = 0) Then
        ReadTibia11ServerList = -1
        Exit Function
    End If
    ReDim res(resSize - 1)
    i = 0
    For Each item In tmpRes
        Key = item
        val = tmpRes(Key)
        tmpElement.Id = Key
        'tmpElement.rawbytes = val
        tmpElement.name_adr = BitConverter_ToInt32(val, &H18)
        tmpElement.url_adr = BitConverter_ToInt32(val, &H1C)
        tmpElement.name = QMemory_ReadString(pid, tmpElement.name_adr)
        tmpElement.url = QMemory_ReadString(pid, tmpElement.url_adr)
        tmpElement.port = BitConverter_ToInt32(val, adrServerList_PortOffset)
        If (adrServerList_PortOffset > &H20) Then
            tmpElement.url2_adr = BitConverter_ToInt32(val, &H20)
            tmpElement.url2 = QMemory_ReadString(pid, tmpElement.url2_adr)
        Else
            tmpElement.url2_adr = 0
            tmpElement.url2 = ""
        End If
        tmpElement.this_register_adr = BitConverter_ToInt32(val, cte_bytesPerRegister)  ' trick (we left base address in last 4 bytes)
        tmpElement.port_adr = tmpElement.this_register_adr + adrServerList_PortOffset
        res(tmpElement.Id) = tmpElement
        i = i + 1
    Next
    ReadTibia11ServerList = 0
    Exit Function
goterr:
    ReadTibia11ServerList = -1
End Function

Public Function BitConverter_ToInt32(ByRef arr() As Byte, ByRef pos As Long) As Long
    Dim l As Long
    CopyMemory l, arr(pos), 4
    BitConverter_ToInt32 = l
End Function
Public Function BitConverter_ToInt8(ByRef arr() As Byte, ByRef pos As Long) As Byte
    BitConverter_ToInt8 = arr(pos)
End Function

Private Sub closeTibia11Client(ByVal pid As Long)


Dim bRes As Boolean
Dim tibiaHandle As Long
On Error GoTo goterr

    If mainTibiaHandle.Exists(pid) Then ' retrieve directly from our dictionary
        tibiaHandle = mainTibiaHandle(pid)
    Else
        tibiaHandle = Get_MainWindowHandle_from_ProcessID_and_class(pid, tibiaclassname) ' slow procedure
    End If
        bRes = ProcessTerminate(pid, tibiaHandle)
        Exit Sub
goterr:
        Debug.Print "WARNING: Unable to close Tibia PID " & CStr(pid)

End Sub

Public Function IsFormLoaded(fForm As Form) As Boolean
On Error GoTo Err_Proc

Dim x As Integer

For x = 0 To Forms.Count - 1
If (Forms(x) Is fForm) Then
IsFormLoaded = True
Exit Function
End If
Next x

IsFormLoaded = False


Exit Function

Err_Proc:

IsFormLoaded = True

End Function

Public Sub RedirectAllServersHere(Optional ByVal debugMode As Integer = 0) ' needs to be called often to capture all connections
    If (confirmedExit = True) Then
        ' already closing bot
        Exit Sub
    End If
    
    Dim tibia_pids() As Long
    Dim totalPids As Long
    totalPids = GetTibiaPIDs(tibiamainname, tibiaclassname, tibia_pids)
    LastNumTibiaClients = totalPids
    If totalPids = 0 Then
        ' Debug.Print "Tibia 11 clients not found (0)"
        Exit Sub
    End If
    Dim newPort As Long
    Dim serverList() As TibiaServerEntry
    Dim i As Long
    Dim j As Long
    Dim readResult As Long
    Dim a As Long
    Dim strServerName As String
    Dim strURLtrans As String
    Dim strDefaultServer As String
    newPort = CLng(frmMain.sckClientGame(0).LocalPort)
    If (newPort = 0) Then
        Exit Sub
    End If
    For j = 0 To UBound(tibia_pids)
  
        If (debugMode = 1) Then
            readResult = ReadTibia11ServerList(tibia_pids(j), adrServerList_CollectionStart, serverList)
            If (readResult = 0) Then
                For i = 0 To UBound(serverList)
                  strServerName = serverList(i).name
              
                  Debug.Print CStr(serverList(i).name & ": " & serverList(i).url & " port " & serverList(i).port & " (" & Hex(serverList(i).this_register_adr) & ") DDoS safe url=" & serverList(i).url2 & " (" & Hex(serverList(i).url2_adr) & ") referenced at " & Hex(serverList(i).url_adr))
                Next i
            End If
            Dim lastCharIndex As Integer
            Dim charList() As TibiaCharListEntry
            lastCharIndex = CInt(ReadTibia11CharList(tibia_pids(0), charList))
            If lastCharIndex > -1 Then
                Dim strNick As String
                Dim strServer As String
                Dim strDomain As String
                Dim aa As Integer
                For aa = 0 To lastCharIndex
                   strNick = charList(aa).name
                   strServer = charList(aa).server
                   strDomain = GetGameServerDOMAIN1(strServer)
                   Debug.Print "> " & strNick & " > " & strServer & " > " & strDomain
                Next aa
                Debug.Print "debug ok"
                End
            End If
        Else
            readResult = ReadTibia11ServerList(tibia_pids(j), adrServerList_CollectionStart, serverList, newPort)
             Select Case (readResult)
             Case 0
               
                For i = 0 To UBound(serverList)
                    strServerName = serverList(i).name
                    If (serverList(i).url = "127.0.0.1") Then ' already modified
                        If (GetGameServerDOMAIN1(strServerName) = "") Then
                            'showWarning = True
                           ' strDefaultServer = LCase(strServerName) & defaultGameServerEnd
                           ' AddGameServer strServerName, "127.0.0.1:" & 7171, strDefaultServer
                           closeTibia11Client tibia_pids(j)
                           If IsFormLoaded(frmMain) Then
                              frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Tibia client PID " & CStr(tibia_pids(j)) & " closed for security reasons (unable to handle previously modified client list). No big problem. Please just launch Tibia client again."
                           End If
                           Exit For
                        End If
                        ' no need to write 127.0.0.1 again
                    Else
                        If (GetGameServerDOMAIN1(strServerName) = "") Then
                            AddGameServer strServerName, "127.0.0.1:" & serverList(i).port, serverList(i).url
                            If (serverList(i).url2_adr > 0) Then
                                AddGameServer2 strServerName, serverList(i).url2
                            End If
                        End If
                        ' only if url is different than 127.0.0.1 then we update it to 127.0.0.1
                        ModifyQString tibia_pids(j), serverList(i).url_adr, "127.0.0.1"
                        If (serverList(i).url2_adr > 0) Then
                           ModifyQString tibia_pids(j), serverList(i).url2_adr, "127.0.0.1"
                        End If
                    End If
                    QMemory_Write4Bytes tibia_pids(j), serverList(i).port_adr, newPort ' we always need to update our port
                Next i
                If cteDebugConEvents = True Then
                    LogConEvent "PID " & CStr(tibia_pids(j)) & ": MODIFIED server list. Blackd Proxy is now ready here."
                End If

             Case -1
               '  Debug.Print "PID " & CStr(tibia_pids(j)) & ": NO SERVER LIST YET"
              Case -2
                 'Debug.Print "PID " & CStr(tibia_pids(j)) & ": ALREADY modified"
                 
             End Select
        End If
    Next j
    
 
End Sub
Public Sub BuildCharListForTibiaQ(ByVal idConnection As Integer, ByRef selName As String, ByRef listPos As Integer)
    Dim a As Integer
    Dim strNick As String
    Dim strServer As String
    Dim strPort As Long
    Dim strDomain As String
    Dim lastCharIndex As Integer
    Dim charList() As TibiaCharListEntry
    selName = ""
    listPos = -1
    lastCharIndex = CInt(ReadTibia11CharList(ProcessID(idConnection), charList))
    If lastCharIndex = -1 Then
       Exit Sub
    End If
    selName = QMemory_ReadStringP(ProcessID(idConnection), adrSelectedCharName_afterCharList)
    If selName = "" Then
       Exit Sub
    End If
    ResetCharList2 idConnection
    For a = 0 To lastCharIndex
        strNick = charList(a).name
        strServer = charList(a).server
        If ((useAntiDDoS = True) And (TibiaVersionLong >= 1104)) Then
            strDomain = GetGameServerDOMAIN2(strServer)
        Else
            strDomain = GetGameServerDOMAIN1(strServer)
        End If
        If (strDomain = "127.0.0.1") Then
            selName = ""
            listPos = -1
            Exit Sub
        End If
        strPort = GetGameServerPort(strServer)
        AddCharServer2 idConnection, strNick, strServer, "127", "0", "0", "1", strPort, strDomain
        If (strNick = selName) Then
            listPos = a
        End If
    Next a
    If (listPos = -1) Then
        ResetCharList2 idConnection
    End If
    If cteDebugConEvents = True Then
        LogConEvent "Selected char position = " & CStr(listPos)
    End If
End Sub



Public Function MAKELPARAM(ByVal wLow As Long, ByVal wHigh As Long) As Long

        MAKELPARAM = LoWord(wLow) Or (&H10000 * LoWord(wHigh))

End Function


Public Function LoWord(ByVal lDWord As Long) As Long

        If lDWord And &H8000& Then

            LoWord = lDWord Or &HFFFF0000

        Else

            LoWord = lDWord And &HFFFF&

        End If

End Function

Public Sub SendClickToTibia11(ByRef pid As Long, ByVal x As Long, ByVal y As Long)
    Dim coordinates As Long
    Dim tibiaHandle As Long
    Dim res As Long
    Dim currentPosition As POINTAPI
    GetCursorPos currentPosition
    coordinates = MAKELPARAM(x, y)
    If mainTibiaHandle.Exists(pid) Then ' retrieve directly from our dictionary
        tibiaHandle = mainTibiaHandle(pid)
    Else
        tibiaHandle = Get_MainWindowHandle_from_ProcessID_and_class(pid, tibiaclassname) ' slow procedure
    End If
    If tibiaHandle = -1 Then
        Debug.Print "Unable to get the main handle from pid " & CStr(pid)
        Exit Sub
    End If
    res = SendMessage(tibiaHandle, WM_MOUSEMOVE, 0&, coordinates)
    res = SendMessage(tibiaHandle, WM_LBUTTONDOWN, 1&, coordinates)
    res = SendMessage(tibiaHandle, WM_MOUSEMOVE, 1&, coordinates)
    res = SendMessage(tibiaHandle, WM_LBUTTONUP, 0&, coordinates)
    coordinates = MAKELPARAM(currentPosition.x, currentPosition.y)
    res = SendMessage(tibiaHandle, WM_LBUTTONUP, 0&, coordinates)
End Sub

Public Sub SafeMemoryMoveXYZ_Tibia11(ByRef idConnection As Integer, Px As Long, Py As Long, Pz As Long)
    Const maxError As Long = 10 ' allows clicking even if it slightly moved (for high level fast chars)
    Dim currentMinimapMinX As Long
    Dim currentMinimapMinY As Long
    Dim currentMinimapPixelsX As Long
    Dim currentMinimapPixelsY As Long
    Dim currentMinimapZ As Long
    Dim currentPointSize As Single
    Dim corner_posx As Long
    Dim corner_posy As Long
    Dim safecheck_width As Long
    Dim safecheck_height As Long
    Dim pid As Long
    Dim retryCount As Long
    Dim retryCount2 As Long
    Dim clickDone As Boolean
    Dim prevSize As Single
    Dim halfX As Long
    Dim halfY As Long
    Dim goodCheck As Boolean
    Dim goalRawX As Long
    Dim goalRawY As Long
    Dim clickX As Long
    Dim clickY As Long
    Dim precisionDifX As Long
    Dim precisionDifY As Long
    pid = ProcessID(idConnection)
    safecheck_width = ReadCurrentAddressDOUBLE(pid, adrMiniMapRect_Width_Double, -1)
    safecheck_height = ReadCurrentAddressDOUBLE(pid, adrMiniMapRect_Height_Double, -1)
    If Not (safecheck_width = 172) Then
        Debug.Print "Unexpected Minimap width"
        Exit Sub
    End If
    If Not (safecheck_height = 113) Then
        Debug.Print "Unexpected Minimap width"
        Exit Sub
    End If
    corner_posx = ReadCurrentAddressDOUBLE(pid, adrGameRect_Width_Double, -1)
    corner_posy = ReadCurrentAddressDOUBLE(pid, adrMiniMapRect_Y_Double, -1)
    
    retryCount = 0
    Do
        currentPointSize = ReadCurrentAddressFLOAT(pid, adrMiniMapDisplay_Zoom_PointSize1_Float, -1)
        If (currentPointSize = -1) Then
            Debug.Print "Minimap address bug!"
            Exit Sub
        End If
        If currentPointSize < 1 Then
            Debug.Print "requires Zoom +"
            prevSize = currentPointSize
            SendClickToTibia11 pid, corner_posx + 135, corner_posy + 80
            retryCount2 = 0
            Do
                DoEvents
                wait 10
                retryCount2 = retryCount2 + 1
                If (retryCount2 > 10) Then
                    Debug.Print "Minimap zoom+ click failed!"
                    Exit Sub
                End If
                currentPointSize = ReadCurrentAddressFLOAT(pid, adrMiniMapDisplay_Zoom_PointSize1_Float, -1)
            Loop Until Not (currentPointSize = prevSize)
        ElseIf currentPointSize > 1 Then
            Debug.Print "requires Zoom -"
            prevSize = currentPointSize
            SendClickToTibia11 pid, corner_posx + 135, corner_posy + 60
            retryCount2 = 0
            Do
                DoEvents
                wait 10
                retryCount2 = retryCount2 + 1
                If (retryCount2 > 10) Then
                    Debug.Print "Minimap zoom- click failed!"
                    Exit Sub
                End If
                currentPointSize = ReadCurrentAddressFLOAT(pid, adrMiniMapDisplay_Zoom_PointSize1_Float, -1)
            Loop Until Not (currentPointSize = prevSize)
        End If
        retryCount = retryCount + 1
        If retryCount > 5 Then
            Debug.Print "Unexpected problem with minimap zoom"
            Exit Sub
        End If
    Loop Until currentPointSize = 1
    'Debug.Print "Now Zoom level is OK"
    clickDone = False
    retryCount = 0
    Do
        currentMinimapZ = ReadCurrentAddress(pid, adrMiniMapDisplay_Z, -1, True)
        If (currentMinimapZ = -1) Then
            Debug.Print "Minimap address bug!"
            Exit Sub
        End If
        If Not (currentMinimapZ = myZ(idConnection)) Then
            If (clickDone = False) Then
                Debug.Print "Requires centre..."
                SendClickToTibia11 pid, corner_posx + 150, corner_posy + 100
                clickDone = True
            End If
            DoEvents
            wait 10
        End If
        retryCount = retryCount + 1
        If retryCount > 10 Then
            Debug.Print "Minimap centre click failed!"
            Exit Sub
        End If
        If Not (currentMinimapZ = myZ(idConnection)) Then
            currentMinimapZ = ReadCurrentAddress(pid, adrMiniMapDisplay_Z, -1, True)
        End If
    Loop Until (currentMinimapZ = myZ(idConnection))
   ' Debug.Print "Now minimap Z is OK"
 
    clickDone = False
    retryCount = 0
    Do
        currentMinimapMinX = ReadCurrentAddress(pid, adrMiniMapDisplay_MinX, -1, True)
        currentMinimapMinY = ReadCurrentAddress(pid, adrMiniMapDisplay_MinY, -1, True)
        currentMinimapPixelsX = ReadCurrentAddress(pid, adrMiniMapDisplay_SizeX, -1, True)
        currentMinimapPixelsY = ReadCurrentAddress(pid, adrMiniMapDisplay_SizeY, -1, True)
        halfX = (currentMinimapPixelsX - 1) / 2
        halfY = (currentMinimapPixelsY - 1) / 2
        precisionDifX = myX(idConnection) - currentMinimapMinX - halfX
        precisionDifY = myY(idConnection) - currentMinimapMinY - halfY
        goodCheck = (Math.Abs(precisionDifX) < maxError) And (Math.Abs(precisionDifY) < maxError)
        If goodCheck = False Then
            If (clickDone = False) Then
                Debug.Print "Requires centre..."
                SendClickToTibia11 pid, corner_posx + 150, corner_posy + 100
                clickDone = True
            End If
            DoEvents
            wait 10
        End If
        retryCount = retryCount + 1
        If retryCount > 10 Then
            Debug.Print "Minimap centre click failed!"
            Exit Sub
        End If
        If goodCheck = False Then
            currentMinimapMinX = ReadCurrentAddress(pid, adrMiniMapDisplay_MinX, -1, True)
            currentMinimapMinY = ReadCurrentAddress(pid, adrMiniMapDisplay_MinY, -1, True)
            currentMinimapPixelsX = ReadCurrentAddress(pid, adrMiniMapDisplay_SizeX, -1, True)
            currentMinimapPixelsY = ReadCurrentAddress(pid, adrMiniMapDisplay_SizeY, -1, True)
            halfX = (currentMinimapPixelsX - 1) / 2
            halfY = (currentMinimapPixelsY - 1) / 2
            precisionDifX = myX(idConnection) - currentMinimapMinX - halfX
            precisionDifY = myY(idConnection) - currentMinimapMinY - halfY
            goodCheck = (Math.Abs(precisionDifX) < maxError) And (Math.Abs(precisionDifY) < maxError)
        End If
    Loop Until goodCheck
    If Not ((precisionDifX = 0) And (precisionDifY = 0)) Then
        'Debug.Print "Map was slightly moved: " & CStr(precisionDifX) & "," & CStr(precisionDifY)
    End If
    goalRawX = Px - myX(idConnection) + halfX + precisionDifX
    goalRawY = Py - myY(idConnection) + halfY + precisionDifY
    If (goalRawX > (currentMinimapPixelsX - 1)) Or (goalRawY > (currentMinimapPixelsY - 1)) Then
        Debug.Print "SafeMemoryMoveXYZ_Tibia11 FAIL: Required position is out of bounds"
        Exit Sub
    End If
    clickX = corner_posx + 9 + goalRawX
    clickY = corner_posy + 6 + goalRawY
    SendClickToTibia11 pid, clickX, clickY
  '  Debug.Print "done click at " & CStr(clickX) & "," & CStr(clickY)
End Sub
