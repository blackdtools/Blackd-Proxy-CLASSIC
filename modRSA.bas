Attribute VB_Name = "modRSA"
#Const FinalMode = 1
Option Explicit

Public adrRSA As Long
'Public Const RLserverRSAkey As String = "124710459426827943004376449897985582167801707960697037164044904862948569380850421396904597686953877022394604239428185498284169068581802277612081027966724336319448537811441719076484340922854929273517308661370727105382899118999403808045846444647284499123164879035103627004668521005328367415259939915284902061793"

' confirmed key still used in tibia 10.98
Public Const RLserverRSAkey1075 As String = "132127743205872284062295099082293384952776326496165507967876361843343953435544496682053323833394351797728954155097012103928360786959821132214473291575712138800495033169914814069637740318278150290733684032524174782740134357629699062987023311132821016569775488792221429527047321331896351555606801473202394175817"
Public Const RLserverRSAkey1075_HEX As String = "BC27F992A96B8E2A43F4DFBE1CEF8FD51CF43D2803EE34FBBD8634D8B4FA32F7D9D9E159978DD29156D62F4153E9C5914263FC4986797E12245C1A6C4531EFE48A6F7C2EFFFFF18F2C9E1C504031F3E4A2C788EE96618FFFCEC2C3E5BFAFAF743B3FC7A872EE60A52C29AA688BDAF8692305312882F1F66EE9D8AEB7F84B1949"
Public Const OTserverRSAkey As String = "109120132967399429278860960508995541528237502902798129123468757937266291492576446330739696001110603907230888610072655818825358503429057592827629436413108566029093628212635953836686562675849720620786279431090218017681061521755056710823876476444260558147179707119674283982419152118103759076030616683978566631413"
' n
Public Const OTserverRSAkey_HEX As String = "9B646903B45B07AC956568D87353BD7165139DD7940703B03E6DD079399661B4A837AA60561D7CCB9452FA0080594909882AB5BCA58A1A1B35F8B1059B72B1212611C6152AD3DBB3CFBEE7ADC142A75D3D75971509C321C5C24A5BD51FD460F01B4E15BEB0DE1930528A5D3F15C1E3CBF5C401D6777E10ACAAB33DBE8D5B7FF5"

Public Const OTserverRSAkeyPRIVATE As String = "46730330223584118622160180015036832148732986808519344675210555262940258739805766860224610646919605860206328024326703361630109888417839241959507572247284807035235569619173792292786907845791904955103601652822519121908367187885509270025388641700821735345222087940578381210879116823013776808975766851829020659073"
' d
Public Const OTserverRSAkeyPRIVATE_HEX As String = "428bd3b5346daf71a761106f71a43102f8c857d6549c54660bb6378b52b0261399de8ce648bac410e2ea4e0a1ced1fac2756331220ca6db7ad7b5d440b7828865856e7aa6d8f45837feee9b4a3a0aa21322a1e2ab75b1825e786cf81a28a8a09a1e28519db64ff9baf311e850c2bfa1fb7b08a056cc337f7df443761aefe8d81"


Public Const CommonRSAkeyExponent = "65537"
' e
Public Const CommonRSAkeyExponent_HEX = "010001"

Public WARNING_USING_OTSERVER_RSA As Boolean
'
Private Const MAX_PATH As Long = 260

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const TH32CS_SNAPMODULE As Long = 8&
Private Const TH32CS_SNAPMODULE32 As Long = 10&
Private Const INVALID_HANDLE_VALUE As Long = -1

' typedef struct tagMODULEENTRY32 {
'  DWORD   dwSize;
'  DWORD   th32ModuleID;
'  DWORD   th32ProcessID;
'  DWORD   GlblcntUsage;
'  DWORD   ProccntUsage;
'  BYTE    *modBaseAddr;
'  DWORD   modBaseSize;
'  HMODULE hModule;
'  TCHAR   szModule[MAX_MODULE_NAME32 + 1];
'  TCHAR   szExePath[MAX_PATH];
' } MODULEENTRY32, *PMODULEENTRY32;

Private Type MODULEENTRY32
 dwSize As Long
 th32ModuleID As Long
  th32ProcessID As Long
  GlblcntUsage As Long
  ProccntUsage As Long
  modBaseAddr As Long
  modBaseSize As Long
  hModule As Long
'  TCHAR   szModule[MAX_MODULE_NAME32 + 1];
  szModule As String * 256 ' hope this is right...
'  TCHAR   szExePath[MAX_PATH];
  szExePath As String * MAX_PATH 'and this..
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

Private Declare Function CreateToolhelp32Snapshot Lib "Kernel32" _
   (ByVal lFlags As Long, ByVal lProcessID As Long) As Long

Private Declare Function Module32First Lib "Kernel32" _
   (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "Kernel32" _
   (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long



Public Declare Function ProcessFirst Lib "Kernel32" _
    Alias "Process32First" _
   (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

Public Declare Function ProcessNext Lib "Kernel32" _
    Alias "Process32Next" _
   (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

Private Declare Sub CloseHandle Lib "Kernel32" _
   (ByVal hPass As Long)

Private Function GetMainModuleAddress(ByVal process_Hwnd As Long, ByRef MainModuleAddress As Long, ByRef MainModuleSize As Long) As Boolean
  Dim hSnapshot As Long
  Dim uHandle As MODULEENTRY32
  Dim foo As Long
  Dim pid As Long
  GetWindowThreadProcessId process_Hwnd, pid
  uHandle.dwSize = Len(uHandle) ' DO NOT use Len$ here!
  hSnapshot = CreateToolhelp32Snapshot(24, pid) '24=TH32CS_SNAPMODULE | TH32CS_SNAPMODULE32
  If (hSnapshot = INVALID_HANDLE_VALUE) Then
    Debug.Print "CreateToolhelp32Snapshot failed on pid " & CStr(pid) & " ...TODO: use GetLastError() for more info about why it failed"
    GetMainModuleAddress = False
    Exit Function
  End If
  foo = Module32First(hSnapshot, uHandle)
  If (foo = 0) Then
    Debug.Print "Module32First failed on pid " & CStr(pid) & " ...TODO: use GetLastError() for more info about why it failed"
    CloseHandle (hSnapshot)
    GetMainModuleAddress = False
    Exit Function
  End If
  CloseHandle (hSnapshot)
  MainModuleAddress = uHandle.modBaseAddr
  MainModuleSize = uHandle.modBaseSize
  GetMainModuleAddress = True
End Function



Public Sub AutoUpdateRSA(ByVal pid As Long)
  On Error GoTo goterr
  Dim pg As Integer
  Dim i As Long
  Dim b As Byte
  Dim sb As String
  Dim s As String
  Dim si As Integer
 ' Dim sc As String
  Dim maxsi As Integer
  Dim backupi As Long
  Dim reskey As String
  Dim TibiaExeModuleAddress As Long
  Dim TibiaExeModuleSize As Long
  Dim TibiaExeModuleEnd As Long

   frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Trying to autoupdate adrRSA..."
   If (GetMainModuleAddress(pid, TibiaExeModuleAddress, TibiaExeModuleSize) = False) Then
     frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "FAIL ... Error at AutoUpdateRSA, GetMainModuleAddress failed.."
     adrRSA = 0
     'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "FAIL ... Error at AutoUpdateRSA (" & CStr(Err.Number) & ") : " & Err.Description
     Exit Sub
  End If
  TibiaExeModuleEnd = TibiaExeModuleAddress + TibiaExeModuleSize
  reskey = ""
  pg = 0
  maxsi = 1
  si = 1
  'sc = Mid$(RLserverRSAkey, si, 1)
  sb = ""
  ' i = &H500000
  i = TibiaExeModuleAddress
  Do
     b = Memory_ReadByte(i, pid)
     sb = Chr(b)
     If (IsNumeric(sb)) Then
    ' If (sb = sc) Then
    reskey = reskey & sb
       si = si + 1
       If (si = 2) Then
         backupi = i
       End If
       If (si > maxsi) Then
         maxsi = si
         'Debug.Print ("New record (" & maxsi - 1 & ") at &H" & Hex(backupi)) & " : " & reskey
         If maxsi - 1 = 309 Then
           adrRSA = backupi
           frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "SUCCESS!! Found RSA key at &H" & (Hex(adrRSA)) & " : " & reskey
           Exit Sub
         End If
       End If
     Else
       reskey = ""
       If (si > 1) Then
         i = backupi
         si = 1
       End If
     End If
    ' sc = Mid$(RLserverRSAkey, si, 1)
     pg = pg + 1
     If (pg >= 10000) Then
       frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & Hex(i) & " Searching RSA key for this tibia client..."
       pg = 0
     End If
     i = i + 1
     DoEvents
  'Loop Until i >= &HA00000
  Loop Until i >= TibiaExeModuleEnd
   frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "FAIL ... MEMORY SCAN COMPLETED WITHOUT RESULTS"
   Exit Sub
   
goterr:
  adrRSA = 0
   frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "FAIL ... Error at AutoUpdateRSA (" & CStr(Err.Number) & ") : " & Err.Description
  Exit Sub
End Sub
Public Sub TryToUpdateRSA(ByVal process_Hwnd As Long, ByVal strKey As String, Optional fixRSA As Boolean = False)
' at this moment fixRSA =true only will works at Windows XP (or in any Window if ASLR is disabled)
    Dim i As Long
    Dim writeChr As String
    Dim currByteAdr As Long
    Dim byteChr As Byte
    Dim byteChrR As Byte
    Dim RSA_bytes(308) As Byte
    Dim b1 As Byte
    Dim b2 As Byte
    Dim b3 As Byte
    Dim b4 As Byte
    Dim res As Long
    Dim realAddress As Long
    fixRSA = False ' @Programmers: you can uncomment this to obtain adrRSA in old clients. ASLR should be disabled with Microsoft EMET
    Debug.Print "adrRSA=" & Hex(adrRSA)
    
    If fixRSA = True Then
      If adrRSA = 0 Then
        AutoUpdateRSA (process_Hwnd)
        If adrRSA = 0 Then
           Debug.Print ("Failed to obtain RSA address. ASLR was  enabled so it was not possible to obtain it.")
        Else
           Debug.Print ("Obtained RSA key = &H" & Hex(adrRSA))
        End If
      End If
    End If

    If adrRSA = 0 Then
        Exit Sub
    End If
    If process_Hwnd = -1 Then
        Exit Sub
    End If
   
    realAddress = Memory_BlackdAddressToFinalAdddress(adrRSA, process_Hwnd)
    If (realAddress = 0) Then
      Exit Sub
    End If
   
    For i = 0 To 308
      writeChr = Mid$(strKey, i + 1, 1)
      byteChr = ConvStrToByte(writeChr)
      RSA_bytes(i) = byteChr
    Next i
   
    res = BlackdForceWrite(realAddress, RSA_bytes(0), 309, process_Hwnd)
    Debug.Print "RSA key changed"

    WARNING_USING_OTSERVER_RSA = True
End Sub

Public Sub ModifyTibiaRSAs()
  Dim tibiaclient As Long
  tibiaclient = 0
  Do
    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
        Exit Do
    Else
        TryToUpdateRSA tibiaclient, OTserverRSAkey
    End If
  Loop
End Sub
