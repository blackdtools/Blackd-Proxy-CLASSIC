Attribute VB_Name = "modTibia11Functions"
' Change to Tibia11allowed = 1 to allow Tibia 11 configs
' Don't worry! If you are a programmer capable to compile this code
' then you are authorized to use Tibia 11 configs even if you didn't purchase gold.
' However, you should not share it with other people.
#Const Tibia11allowed = 0
#Const FinalMode = 1
Option Explicit
#If Tibia11allowed = 1 Then
    Public Const Tibia11allowed As Boolean = True
#Else
    Public Const Tibia11allowed As Boolean = False
#End If


Public Const defaultGameServerEnd As String = "-lb.ciproxy.com"

Public Const CTE_NOT_CONNECTED As Integer = 0
Public Const CTE_LOGIN_CHAR_SELECTION As Integer = 1
Public Const CTE_CONNECTING As Integer = 2
Public Const CTE_GAME_CONNECTED As Integer = 3
    

'***********************
'* Win32 Constants . . .
'***********************

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
    id As Long
    name As String
    url As String
    port As Long
    this_register_adr As Long
    name_adr As Long
    url_adr As Long
    port_adr As Long
   ' rawbytes() As Byte
End Type

Public Type TibiaCharListEntry
    id As Long
    name As String
    server As String
    entry_address As Long
    name_address As Long
End Type



'Private Declare Function VirtualProtectEx Lib "Kernel32" (ByVal hProcess As Long, ByRef lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long


Private Declare Function GetClassName Lib "user32" _
   Alias "GetClassNameA" _
   (ByVal hwnd As Long, _
   ByVal lpClassName As String, _
   ByVal nMaxCount As Long) As Long
   
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
(ByVal hwnd As Long, ByVal wIndx As Long) As Long

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" _
 (ByVal hwnd As Long) As Long
 
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
 (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
 
Private Declare Function EnumWindows Lib "user32" _
 (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)
    
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Module32FirstW Lib "kernel32" (ByVal hSnapshot As Long, ByRef uModule As Any) As Long
Private Declare Function Module32NextW Lib "kernel32" (ByVal hSnapshot As Long, ByRef uModule As Any) As Long

Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Private Declare Function VirtualQueryEx& Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long)

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
   
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Public moduleDictionary As scripting.Dictionary
Public objWMIService As Object


Public Function QMemory_ReadNBytes(ByVal pid As Long, ByVal finalAddress As Long, ByRef Rbuff() As Byte) As Long
    Dim usize As Long
    Dim TibiaHandle As Long
    Dim readtotal As Long
    On Error GoTo gotErr
    readtotal = 0
    usize = UBound(Rbuff) + 1
    If (usize < 1) Then
        Exit Function
    End If
    TibiaHandle = OpenProcess(PROCESS_VM_READ, 0, pid)
    ReadProcessMemory TibiaHandle, finalAddress, Rbuff(0), usize, readtotal
    CloseHandle (TibiaHandle)
    If (readtotal = 0) Then
        QMemory_ReadNBytes = -1
    Else
        QMemory_ReadNBytes = 0
    End If
    Exit Function
gotErr:
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
    Dim TibiaHandle As Long
    Dim res As Long
    Dim lpNumberOfBytesWritten As Long
    Dim usize As Long
    lpNumberOfBytesWritten = 0
    On Error GoTo gotErr
    usize = UBound(newValue) + 1
    TibiaHandle = OpenProcess(PROCESS_READ_WRITE_QUERY, 0, pid)
    If TibiaHandle = -1 Then
        QMemory_WriteNBytes = -1
        Exit Function
    End If
    res = WriteProcessMemory(TibiaHandle, finalAddress, newValue(0), usize, lpNumberOfBytesWritten)
    If (res = 1) Then
        CloseHandle (TibiaHandle)
        QMemory_WriteNBytes = 0
    Else
        CloseHandle (TibiaHandle)
        QMemory_WriteNBytes = -1
    End If
    Exit Function
gotErr:
    QMemory_WriteNBytes = -1
End Function

Public Function QMemory_Write2Bytes(ByVal pid As Long, ByVal finalAddress As Long, newValue As Long) As Long
    Dim TibiaHandle As Long
    Dim res As Long
    Dim lpNumberOfBytesWritten As Long
    Dim Rbuff(1) As Byte
    lpNumberOfBytesWritten = 0
    On Error GoTo gotErr
    Rbuff(0) = LowByteOfLong(newValue)
    Rbuff(1) = HighByteOfLong(newValue)
    TibiaHandle = OpenProcess(PROCESS_READ_WRITE_QUERY, 0, pid)
    res = WriteProcessMemory(TibiaHandle, finalAddress, Rbuff(0), 2, lpNumberOfBytesWritten)
    If (res = 1) Then
        CloseHandle (TibiaHandle)
        QMemory_Write2Bytes = 0
    Else
        CloseHandle (TibiaHandle)
        QMemory_Write2Bytes = -1
    End If
    Exit Function
gotErr:
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
        On Error GoTo gotErr
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
gotErr:
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
gotErr:
    QMemory_ReadDouble = -1
End Function

    
Public Function QMemory_Read4Bytes(ByVal pid As Long, ByVal finalAddress As Long) As Long
    Dim res As Long
    Dim TibiaHandle As Long
    On Error GoTo gotErr
    TibiaHandle = OpenProcess(PROCESS_VM_READ, 0, pid)
    ReadProcessMemory TibiaHandle, finalAddress, res, 4, 0
    CloseHandle (TibiaHandle)
    QMemory_Read4Bytes = res
    Exit Function
gotErr:
    QMemory_Read4Bytes = -1
End Function

Public Function QMemory_Read2Bytes(ByVal pid As Long, ByVal finalAddress As Long) As Long
    Dim Rbuff(1) As Byte
    Dim TibiaHandle As Long
    On Error GoTo gotErr
    TibiaHandle = OpenProcess(PROCESS_VM_READ, 0, pid)
    ReadProcessMemory TibiaHandle, finalAddress, Rbuff(0), 2, 0
    CloseHandle (TibiaHandle)
    QMemory_Read2Bytes = GetTheLong(Rbuff(0), Rbuff(1))
    Exit Function
gotErr:
    QMemory_Read2Bytes = -1
End Function

Public Function QMemory_Read1Byte(ByVal pid As Long, ByVal finalAddress As Long) As Byte
    Dim Rbuff As Byte
    Dim TibiaHandle As Long
    On Error GoTo gotErr
    TibiaHandle = OpenProcess(PROCESS_VM_READ, 0, pid)
    ReadProcessMemory TibiaHandle, finalAddress, Rbuff, 1, 0
    CloseHandle (TibiaHandle)
    QMemory_Read1Byte = Rbuff
    Exit Function
gotErr:
    QMemory_Read1Byte = &HFF
End Function
    
Public Function QMemory_Write4Bytes(ByVal pid As Long, ByVal finalAddress As Long, ByVal newValue As Long) As Long
    Dim TibiaHandle As Long
    Dim res As Long
    Dim lpNumberOfBytesWritten As Long
    lpNumberOfBytesWritten = 0
    On Error GoTo gotErr
    TibiaHandle = OpenProcess(PROCESS_READ_WRITE_QUERY, 0, pid)
    res = WriteProcessMemory(TibiaHandle, finalAddress, newValue, 4, lpNumberOfBytesWritten)
    If (res = 1) Then
        CloseHandle (TibiaHandle)
        QMemory_Write4Bytes = 0
    Else
        CloseHandle (TibiaHandle)
        QMemory_Write4Bytes = -1
    End If
    Exit Function
gotErr:
    QMemory_Write4Bytes = -1
End Function





Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
 Dim Title As String
 Dim r As Long
 
 r = GetWindowTextLength(hwnd)
 Title = Space(r)
 GetWindowText hwnd, Title, r + 1
 
 'Add to type array
 ColCounter = ColCounter + 1
 ReDim Preserve HandleTextCollection(ColCounter)
 HandleTextCollection(ColCounter).Window_Handle = hwnd
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
 On Error GoTo gotErr
 If (HandleTextCollection(0).Window_Handle = &H0) Then
 CheckIfEnumDone = True
 Else
 CheckIfEnumDone = True
 End If
 Exit Function
gotErr:
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


Private Function GetWindowClass(ByVal hwnd As Long) As String
  Dim sClass As String
  If hwnd = 0 Then
    GetWindowClass = ""
  Else
    sClass = Space$(256)
    GetClassName hwnd, sClass, 255
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
    On Error GoTo gotErr
    Dim items As Object
    Dim item As Object
    'Dim count As Long
    Dim uModule As MODULEENTRY32W, lModuleSnapshot&
    If (moduleDictionary Is Nothing) Then
        Set moduleDictionary = New scripting.Dictionary
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
           ' Debug.Print "Found Tibia PID: " & CStr(currentPID)
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
gotErr:
    Debug.Print ("Error: Unexpected error - " & Err.Description)
End Sub



Public Function GetTibiaPIDs(ByRef expectedName As String, ByRef expectedClass As String, _
ByRef CurrentTibiaPids() As Long) As Long
    Dim ubproc As Long
    Dim i As Long
    On Error GoTo gotErr
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
gotErr:
    Debug.Print ("Error: Unexpected error - " & Err.Description)
     CurrentTibiaPids(0) = -1
    GetTibiaPIDs = 0
End Function

Public Function arrayToString(ByRef arr() As Byte) As String
    Dim isize As Long
    Dim res As String
    Dim i As Long
    On Error GoTo gotErr
    res = "Array ="
    isize = UBound(arr)
    For i = 0 To isize
        res = res & " " & GoodHex(arr(i))
    Next i
    arrayToString = res
    Exit Function
gotErr:
    arrayToString = ""
End Function



Private Sub fillCollectionDictionary(ByRef pid As Long, ByVal adrCurrentItem As Long, ByVal adrSTARTER_ITEM As Long, _
                                     ByRef dict As scripting.Dictionary, _
                                     ByRef totalItems As Long, ByVal currentDepth As Long, _
                                     ByRef maxDepth As Long, ByRef bytesPerElement As Long, _
                                     Optional ByRef maxValidKeyID As Long = -1, _
                                     Optional ByVal addBaseAddress As Boolean = False)
    On Error GoTo gotErr
        Dim id As Long
        Dim auxRes As Long
        id = QMemory_Read4Bytes(pid, adrCurrentItem + &H10)
        If (maxValidKeyID > -1) Then
            If (id > maxValidKeyID) Then
                Exit Sub
            End If
        End If
        If (dict.Exists(id) = False) Then
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
            dict(id) = tmp
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
gotErr:
        Debug.Print ("Something failed: " + Err.Description)

End Sub

Public Sub ReadTibia11Collection(ByRef pid As Long, adrPath As AddressPath, _
                                      ByRef bytesPerElement As Long, _
                                      ByRef dict As scripting.Dictionary, _
                                      Optional ByVal directAddress As Long = -1, _
                                      Optional ByVal customDepth As Long = -1, _
                                      Optional ByVal maxValidKeyID As Long = -1, _
                                      Optional ByVal addBaseAddress As Boolean = False)
                                    
    Dim adrCOLLECTION_START As Long
    Dim totalItems As Long
    Dim adrSTARTER_ITEM As Long
 
    Dim maxDepth As Long
    Dim p0, p1, p2 As Long
    Set dict = New scripting.Dictionary
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
    adrSTARTER_ITEM = QMemory_Read4Bytes(pid, adrCOLLECTION_START)
    p0 = QMemory_Read4Bytes(pid, adrSTARTER_ITEM)
    p1 = QMemory_Read4Bytes(pid, adrSTARTER_ITEM + 4)
    p2 = QMemory_Read4Bytes(pid, adrSTARTER_ITEM + 8)
    fillCollectionDictionary pid, p0, adrSTARTER_ITEM, dict, totalItems, 0, maxDepth, bytesPerElement, maxValidKeyID, addBaseAddress
    fillCollectionDictionary pid, p1, adrSTARTER_ITEM, dict, totalItems, 0, maxDepth, bytesPerElement, maxValidKeyID, addBaseAddress
    fillCollectionDictionary pid, p2, adrSTARTER_ITEM, dict, totalItems, 0, maxDepth, bytesPerElement, maxValidKeyID, addBaseAddress
End Sub

' I am investigating this at this moment. This function does not work yet
Public Function FindCollectionItemByKey(ByRef pid As Long, adrPath As AddressPath, ByRef keyToSearch As Long) As Long
    Const maxDepth As Long = 10
    Dim currentDepth As Long
    Dim adrCOLLECTION_START As Long
    Dim adrCurrentItem As Long
    Dim p0, p1, p2 As Long
    Dim lowKey, midKey, highKey As Long
    Dim res As Long
    Dim keyDiffP0, keyDiffP1, keyDiffP2 As Long
    adrCOLLECTION_START = ReadCurrentAddress(pid, adrPath, -1, False)
    adrCurrentItem = QMemory_Read4Bytes(pid, adrCOLLECTION_START)
    currentDepth = 1
     Debug.Print "Now searching " & CStr(keyToSearch) & "..."
    Do
        p0 = QMemory_Read4Bytes(pid, adrCurrentItem)
        p1 = QMemory_Read4Bytes(pid, adrCurrentItem + 4)
        p2 = QMemory_Read4Bytes(pid, adrCurrentItem + 8)
        lowKey = QMemory_Read4Bytes(pid, p0 + &H10)
        midKey = QMemory_Read4Bytes(pid, p1 + &H10)
        highKey = QMemory_Read4Bytes(pid, p2 + &H10)
        Debug.Print "Iteration #" & CStr(currentDepth) & ": [" & CStr(lowKey) & "," & CStr(midKey) & "," & CStr(highKey) & "]"
        If (keyToSearch = lowKey) Then
            Debug.Print "OK: Key " & CStr(keyToSearch) & " found at p0"
            res = p0
            Exit Do
        End If
        If (keyToSearch = midKey) Then
            Debug.Print "OK: Key " & CStr(keyToSearch) & " found at p1"
            res = p1
            Exit Do
        End If
        If (keyToSearch = highKey) Then
            Debug.Print "OK: Key " & CStr(keyToSearch) & " found at p2"
            res = p2
            Exit Do
        End If
        If (keyToSearch < lowKey) Then
            Debug.Print "FAIL: Key " & CStr(keyToSearch) & " not found in this collection"
            res = -1
            Exit Do
        End If
        If (keyToSearch > highKey) Then
            Debug.Print "FAIL: Key " & CStr(keyToSearch) & " not found in this collection"
            res = -1
            Exit Do
        End If
        adrCurrentItem = p1
        
        currentDepth = currentDepth + 1
    Loop Until (currentDepth > maxDepth) ' we set a max number of loops, just in case
    FindCollectionItemByKey = res
End Function
Public Function BitConverter_ToInt16(ByRef arr() As Byte, ByRef pos As Long) As Long
    Dim i As Integer
    CopyMemory i, arr(pos), 2
    BitConverter_ToInt16 = CLng(i)
End Function

Public Function BitConverter_ToInt32(ByRef arr() As Byte, ByRef pos As Long) As Long
    Dim l As Long
    CopyMemory l, arr(pos), 4
    BitConverter_ToInt32 = l
End Function
    
Public Function ReadTibia11ServerList(ByRef pid As Long, ByRef adrPath As AddressPath, _
 ByRef res() As TibiaServerEntry, Optional ByVal stopIfPort As Long = -1) As Long
    Dim tmpRes As scripting.Dictionary
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
    Const cte_bytesPerRegister As Long = &H24
    On Error GoTo gotErr
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
        tmpElement.id = Key
        'tmpElement.rawbytes = val
        tmpElement.name_adr = BitConverter_ToInt32(val, &H18)
        tmpElement.url_adr = BitConverter_ToInt32(val, &H1C)
        tmpElement.name = QMemory_ReadString(pid, tmpElement.name_adr)
        tmpElement.url = QMemory_ReadString(pid, tmpElement.url_adr)
        tmpElement.port = BitConverter_ToInt32(val, &H20)
        tmpElement.this_register_adr = BitConverter_ToInt32(val, cte_bytesPerRegister)  ' trick (we left base address in last 4 bytes)
        tmpElement.port_adr = tmpElement.this_register_adr + &H20
        res(tmpElement.id) = tmpElement
        i = i + 1
    Next
    ReadTibia11ServerList = 0
    Exit Function
gotErr:
    ReadTibia11ServerList = -1
End Function

Public Function ReadCurrentCharName(ByRef pid As Long) As String
    Dim auxAdr As Long
    Dim strRes As String
    auxAdr = ReadCurrentAddress(pid, adrSelectedCharName, -1, True)
    strRes = QMemory_ReadString(pid, auxAdr)
    ReadCurrentCharName = ""
End Function
    
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
       ' Debug.Print(Hex(auxAdr))

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
           Case Else ' maybe in a menu
               Debug.Print ("status code =" & CStr(pixels))
               TibiaClientConnectionStatus = CTE_NOT_CONNECTED
               Exit Function
       End Select
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
        tmpElement.id = i
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
                realDomain = GetGameServerDOMAIN(strServerName)
                realPort = GetGameServerPort(strServerName)
                If Not (realDomain = "") Then
                    'Debug.Print "Restoring " & Hex(serverList(i).url_adr) & " (" & strServerName & ") to " & realDomain & ":" & realPort
                    ModifyQString tibia_pids(j), serverList(i).url_adr, realDomain
                    QMemory_Write4Bytes tibia_pids(j), serverList(i).port_adr, realPort
                End If
            Next i
            Debug.Print "PID " & CStr(tibia_pids(j)) & ": RESTORED server list."
         End If
    Next j
    
End Sub

Public Sub RedirectAllServersHere(Optional ByVal debugMode As Integer = 0) ' needs to be called often to capture all connections
    If (debugMode = 1) Then
      '  ConvertToPlainBitmap "C:\BlackdProxyCLASSIC\Blackd-Proxy-CLASSIC\test.png", "C:\BlackdProxyCLASSIC\Blackd-Proxy-CLASSIC\test.bmp"
        Exit Sub
    End If
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
    Dim showWarning As Boolean
    newPort = CLng(frmMain.sckClientGame(0).LocalPort)
  
    
    
    For j = 0 To UBound(tibia_pids)
        If (debugMode = 1) Then
'            readResult = ReadTibia11ServerList(tibia_pids(j), adrServerList_CollectionStart, serverList)
'            If (readResult = 0) Then
'                For i = 0 To UBound(serverList)
'                  strServerName = serverList(i).name
'                  strURLtrans = GetGameServerDOMAIN(strServerName)
'                  Debug.Print CStr(serverList(i).name & ": " & serverList(i).url & " (" & strURLtrans & ") referenced at " & Hex(serverList(i).url_adr))
'                Next i
'            End If
            Dim itemAdr As Long
            Const keyToFind As Long = 11
            ' I am investigating this at this moment. This function does not work yet
            itemAdr = FindCollectionItemByKey(tibia_pids(j), adrServerList_CollectionStart, keyToFind)
            Debug.Print "key " & CStr(keyToFind) & " found at " & CStr(Hex(itemAdr))
        Else
            readResult = ReadTibia11ServerList(tibia_pids(j), adrServerList_CollectionStart, serverList, newPort)
             Select Case (readResult)
             Case 0
                showWarning = False
                For i = 0 To UBound(serverList)
                    strServerName = serverList(i).name
                    If (serverList(i).url = "127.0.0.1") Then ' already modified
                        If (GetGameServerDOMAIN(strServerName) = "") Then
                            showWarning = True
                            strDefaultServer = LCase(strServerName) & defaultGameServerEnd
                            AddGameServer strServerName, "127.0.0.1:" & 7171, strDefaultServer
                           
                        End If
                        ' no need to write 127.0.0.1 again
                    Else
                        If (GetGameServerDOMAIN(strServerName) = "") Then
                            AddGameServer strServerName, "127.0.0.1:" & serverList(i).port, serverList(i).url
                        End If
                        ' only if url is different than 127.0.0.1 then we update it to 127.0.0.1
                        ModifyQString tibia_pids(j), serverList(i).url_adr, "127.0.0.1"
                    End If
                    QMemory_Write4Bytes tibia_pids(j), serverList(i).port_adr, newPort ' we always need to update our port
                Next i
                Debug.Print "PID " & CStr(tibia_pids(j)) & ": MODIFIED server list. Blackd Proxy is now ready here."
                If (showWarning) Then
                    Debug.Print "WARNING: Reusing previous char list. Supposing servers = <servername>" & defaultGameServerEnd
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
        strDomain = GetGameServerDOMAIN(strServer)
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
    Debug.Print "Selected char position = " & CStr(listPos)
End Sub
