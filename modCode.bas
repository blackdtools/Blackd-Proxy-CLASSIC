Attribute VB_Name = "modCode"
#Const FinalMode = 1
#Const AllowPowerCommands = 1
#Const cte_DownloadMode = 0 ' 0 = winsock ; 1 = inet
Option Explicit

Public Const FIX_addConfigPaths As String = _
"config1033,config1034,config1035,config1036,config1037,config1038,config1039,config1040,config1041,config1050,config1051,config1051preview,config1052,config1052preview,config1053,config1053preview,config1054,config1055,config1056,config1057,config1058,config1059,config1060,config1061,config1062,config1063,config1064,config1070,config1071,config1072,config1073,config1074,config1075,config1076,config1077,config1078,config1079,config1080,config1081,config1082,config1090,config1091,config1092,config1093,config1094,config1095,config1096,config1097,config1098,config1099,config1100,config1101,config1102,config1103"

Public Const FIX_addConfigVersions As String = _
"10.33,10.34,10.35,10.36,10.37,10.38,10.39,10.4,10.41,10.5,10.51,10.51 preview,10.52,10.52 preview,10.53,10.53 preview,10.54,10.55,10.56,10.57,10.58,10.59,10.60,10.61,10.62,10.63,10.64,10.70,10.71,10.72,10.73,10.74,10.75,10.76,10.77,10.78,10.79,10.80,10.81,10.82,10.90,10.91,10.92,10.93,10.94,10.95,10.96,10.97,10.98,10.99,11.00,11.01,11.02,11.03"

Public Const FIX_addConfigVersionsLongs As String = _
"1033,1034,1035,1036,1037,1038,1039,1040,1041,1050,1051,1051,1052,1052,1053,1053,1054,1055,1056,1057,1058,1059,1060,1061,1062,1063,1064,1070,1071,1072,1073,1074,1075,1076,1077,1078,1079,1080,1081,1082,1090,1091,1092,1093,1094,1095,1096,1097,1098,1099,1100,1101,1102,1103"

Public Const FIX_highestTibiaVersionLong As String = "1103"
Public Const FIX_TibiaVersionDefaultString As String = "10.99"
Public Const FIX_TibiaVersionDefaultLong As String = "1099"
Public Const FIX_TibiaVersionForceString As String = "10.99"

Public Const EQUIPMENT_SLOTS As Long = 11 ' new slot since Tibia 9.54
Public Const SLOT_AMMUNITION As Long = 10
Public Const SLOT_RIGHTHAND As Long = 5
Public Const SLOT_LEFTHAND As Long = 6
Public Const SLOT_BACKPACK As Long = 3

Public Const DropDelayerConst As Long = 3 ' turns to wait before doing a step of the Drop process

Public Const cte_initHP = 10000
Public Const cte_initMANA = 10000
Public Const localstr As String = "127.0.0.1"
Public Const longsecretkey = "pfiwmvjgjikdfzasdruieopqwfhgkvvbnmklpofufrhufhuhsqaewftswgyguuhbvxhchufudhgoipopeqwiueifhjhsfdzvvcdvhfhfruyiurtuiuwfewqweffswqdepoffr"

Public Type TypeInitialPacket
    packet() As Byte
    mustSend As Boolean
End Type
Type url
    Scheme As String
    Host As String
    port As Long
    uri As String
    Query As String
End Type


' LEVELSPY - XRAY



'======Constants=======
'API constants
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
'Statusbar
Const STATUSBAR_DURATION = 50
'Levelspy
Const LEVELSPY_NOP_DEFAULT = 49451
Const LEVELSPY_ABOVE_DEFAULT = 7
Const LEVELSPY_BELOW_DEFAULT = 2
Const LEVELSPY_MIN = 0
Const LEVELSPY_MAX = 7
'name spy
Const NAMESPY_NOP_DEFAULT = 19573
Const NAMESPY_NOP2_DEFAULT = 17013
'z-axis
Const Z_AXIS_DEFAULT = 7 'default ground level

Public adrAccount As Long

Public timeToRetryOpenDepot() As Long

Public LastCharServerIndex As Integer
'====booleans====
Public bLevelSpy() As Boolean

Public LEVELSPY_NOP As Long
Public LEVELSPY_ABOVE As Long
Public LEVELSPY_BELOW As Long
' name spy
Public NAMESPY_NOP As Long
Public NAMESPY_NOP2 As Long
' full light
Public LIGHT_NOP As Long
Public LIGHT_AMOUNT As Long
' player
Public PLAYER_Z As Long

Public RedSquare As Long

Public conEventLog As String
Public Const RETRYDELAY = 10000 ' in ms
'Public Const TOOSLOWLOGINSERVER_MS = 500 ' MS
Public Const MaxTimeWithoutServerPackets = 45000 'in ms


Public Const sndAsync = &H1
Public Const sndLoop = &H8
Public Const sndNoStop = &H10


      'constants required by Shell_NotifyIcon API call:
      Public Const NIM_ADD = &H0
      Public Const NIM_MODIFY = &H1
      Public Const NIM_DELETE = &H2
      Public Const NIF_MESSAGE = &H1
      Public Const NIF_ICON = &H2
      Public Const NIF_TIP = &H4
      Public Const WM_MOUSEMOVE = &H200
      Public Const WM_LBUTTONDOWN = &H201     'Button down
      Public Const WM_LBUTTONUP = &H202       'Button up
      Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Public Const WM_RBUTTONDOWN = &H204     'Button down
      Public Const WM_RBUTTONUP = &H205       'Button up
      Public Const WM_RBUTTONDBLCLK = &H206   'Double-click


'for help file
Public Const SW_NORMAL = 1






Public Const RuneMakerOptions_activated_default = False
Public Const RuneMakerOptions_autoEat_default = False
Public Const RuneMakerOptions_ManaFluid_default = False
Public Const RuneMakerOptions_autoLogoutAnyFloor_default = False
Public Const RuneMakerOptions_autoLogoutCurrentFloor_default = False
Public Const RuneMakerOptions_autoLogoutOutOfRunes_default = False
Public Const RuneMakerOptions_autoWaste_default = False
Public Const RuneMakerOptions_msgSound_default = False
Public Const RuneMakerOptions_msgSound2_default = False
Public Const RuneMakerOptions_firstActionText_default = "exura"
Public Const RuneMakerOptions_firstActionMana_default = 25
Public Const RuneMakerOptions_LowMana_default = 100
Public Const RuneMakerOptions_secondActionText_default = "adura vita"
Public Const RuneMakerOptions_secondActionMana_default = 400
Public Const RuneMakerOptions_secondActionSoulpoints_default = 3

'custom ng var
Public Const healingCheatsOptions_sdmax_default = False
Public Const healingCheatsOptions_antipush_default = False
Public Const healingCheatsOptions_pmax_default = False
Public Const healingCheatsOptions_htarget_default = False
Public Const healingCheatsOptions_exaustEat_default = 0
Public Const healingCheatsOptions_HouseX_default = 0
Public Const healingCheatsOptions_HouseY_default = 0

'custom ng healing
Public Const healingCheatsOptions_exaust_default = False
Public Const healingCheatsOptions_txtSpellhi_default = "exura gran"
Public Const healingCheatsOptions_txtSpelllo_default = "exura vita"
Public Const healingCheatsOptions_txtPot_default = ""
Public Const healingCheatsOptions_txtMana_default = ""
Public Const healingCheatsOptions_txtHealthhi_default = "0"
Public Const healingCheatsOptions_txtHealthlo_default = "0"
Public Const healingCheatsOptions_txtHealpot_default = "0"
Public Const healingCheatsOptions_txtManapot_default = "0"
Public Const healingCheatsOptions_txtManahi_default = "70"
Public Const healingCheatsOptions_txtManalo_default = "160"
Public Const healingCheatsOptions_Combo1_default = "Health Potion"
Public Const healingCheatsOptions_Combo2_default = "Mana Potion"

'custom ng extras
Public Const extrasOptions_txtSpell_default = "exura"
Public Const extrasOptions_txtMana_default = "25"
Public Const extrasOptions_txtSSA_default = "0"
Public Const extrasOptions_cmbHouse_default = "North"
Public Const extrasOptions_chkMana_default = False
Public Const extrasOptions_chkDanger_default = False
Public Const extrasOptions_chkPM_default = False
Public Const extrasOptions_chkEat_default = False
Public Const extrasOptions_chkautoUtamo_default = False
Public Const extrasOptions_chkautoHur_default = False
Public Const extrasOptions_chkautogHur_default = False
Public Const extrasOptions_chkAFK_default = False
Public Const extrasOptions_chkGold_default = False
Public Const extrasOptions_chkPlat_default = False
Public Const extrasOptions_chkDash_default = False
Public Const extrasOptions_chkMW_default = False
Public Const extrasOptions_chkSSA_default = False
Public Const extrasOptions_chkTitle_default = False
Public Const extrasOptions_chkHouse_default = False

'custom ng persistent
Public Const persistentOptions_txtHk1_default = "100"
Public Const persistentOptions_txtHk2_default = "100"
Public Const persistentOptions_txtHk3_default = "100"
Public Const persistentOptions_txtHk4_default = "100"
Public Const persistentOptions_txtHk5_default = "100"
Public Const persistentOptions_txtHk6_default = "100"
Public Const persistentOptions_txtHk7_default = "100"
Public Const persistentOptions_txtHk8_default = "100"
Public Const persistentOptions_txtHk9_default = "100"
Public Const persistentOptions_txtHk10_default = "100"
Public Const persistentOptions_txtHk11_default = "100"
Public Const persistentOptions_txtExiva1_default = ""
Public Const persistentOptions_txtExiva2_default = ""
Public Const persistentOptions_txtExiva3_default = ""
Public Const persistentOptions_txtExiva4_default = ""
Public Const persistentOptions_txtExiva5_default = ""
Public Const persistentOptions_txtExiva6_default = ""
Public Const persistentOptions_txtExiva7_default = ""
Public Const persistentOptions_txtExiva8_default = ""
Public Const persistentOptions_txtExiva9_default = ""
Public Const persistentOptions_txtExiva10_default = ""
Public Const persistentOptions_txtExiva11_default = ""
Public Const persistentOptions_chkExiva1_default = False
Public Const persistentOptions_chkExiva2_default = False
Public Const persistentOptions_chkExiva3_default = False
Public Const persistentOptions_chkExiva4_default = False
Public Const persistentOptions_chkExiva5_default = False
Public Const persistentOptions_chkExiva6_default = False
Public Const persistentOptions_chkExiva7_default = False
Public Const persistentOptions_chkExiva8_default = False
Public Const persistentOptions_chkExiva9_default = False
Public Const persistentOptions_chkExiva10_default = False
Public Const persistentOptions_chkExiva11_default = False
Public Const persistentOptions_persistvar1_default = 0
Public Const persistentOptions_persistvar2_default = 0
Public Const persistentOptions_persistvar3_default = 0
Public Const persistentOptions_persistvar4_default = 0
Public Const persistentOptions_persistvar5_default = 0
Public Const persistentOptions_persistvar6_default = 0
Public Const persistentOptions_persistvar7_default = 0
Public Const persistentOptions_persistvar8_default = 0
Public Const persistentOptions_persistvar9_default = 0
Public Const persistentOptions_persistvar10_default = 0
Public Const persistentOptions_persistvar11_default = 0

'custom ng aimbot
Public Const aimbotOptions_chkSDcombo_default = False
Public Const aimbotOptions_chkUEcombo_default = False
Public Const aimbotOptions_txtLeader_default = ""
Public Const aimbotOptions_txtCombo_default = "exevo gran mas vis"

Public Const MAXLOGINMEMORY = 500
Public Const HIGHEST_ITEM_BPSLOT = 99
Private Const GW_HWNDFIRST& = 0
Private Const HWND_NOTOPMOST& = -2
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOMOVE& = &H2
Private Const SWP_NOSIZE& = &H1

Public Const PROCESS_VM_READ = (&H10)
Public Const PROCESS_VM_WRITE = (&H20)
Public Const PROCESS_VM_OPERATION = (&H8)
Public Const PROCESS_QUERY_INFORMATION = (&H400)
Public Const PROCESS_READ_WRITE_QUERY = PROCESS_VM_READ + PROCESS_VM_WRITE + PROCESS_VM_OPERATION + PROCESS_QUERY_INFORMATION


Public Type SYSTEMTIME
wYear As Integer
wMonth As Integer
wDayOfWeek As Integer
wDay As Integer
wHour As Integer
wMinute As Integer
wSecond As Integer
wMilliseconds As Integer
End Type

Public Declare Sub GetSystemTime Lib "Kernel32" _
(lpSystemTime As SYSTEMTIME)


'for help fire
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'for ontop
Public Declare Function GetWindow& Lib "user32" _
    (ByVal hWnd As Long, ByVal wCmd As Long)
Public Declare Function SetWindowPos& Lib "user32" _
    (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long)
    
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'Find a child window with a given class name and caption
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'Get the handle of the desktop window
'Public Declare Function GetDesktopWindow Lib "user32" () As Long
'To read / write ini
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As _
String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As _
Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Private Declare Sub CopyMemory Lib "Kernel32" _
Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal _
Length As Long)

'get classname
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

' get windows with current focus
Public Declare Function GetForegroundWindow Lib "user32" () As Long

' get the caption of the windows
' Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
  (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetWindowText Lib "user32" _
 Alias "SetWindowTextA" (ByVal hWnd As Long, _
 ByVal lpString As String) As Long



#If Win32 Then
  Public Declare Function GetTickCount Lib "Kernel32" () As Long
#Else
  Public Declare Function GetTickCount Lib "user" () As Long
#End If

' tray icon

      Public Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hWnd As Long) As Long
      Public Declare Function Shell_NotifyIcon Lib "shell32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'user defined type required by Shell_NotifyIcon API call
      Public Type NOTIFYICONDATA
       cbSize As Long
       hWnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hIcon As Long
       szTip As String * 64
      End Type
      
Public Type TypeOneListItem
  CharacterName As String
  ServerName As String
  serverIP1 As Byte
  serverIP2 As Byte
  serverIP3 As Byte
  serverIP4 As Byte
  serverPort As Long
End Type
Public Type TypeCharacterList
  numItems As Integer
  pointer As Integer ' write from here if memory is full
  item(1 To MAXLOGINMEMORY) As TypeOneListItem 'memory for up character - server relations
End Type
      
Public Type TibiaTileStr
    str As String
    num As Long
End Type
      
      
Public Type TypeOneListItem2
  CharacterName As String
  ServerName As String
  serverIP1 As Byte
  serverIP2 As Byte
  serverIP3 As Byte
  serverIP4 As Byte
  serverPort As Long
  serverDOMAIN As String
End Type
Public Type TypeCharacterList2
  numItems As Integer
  item(0 To MAXLOGINMEMORY) As TypeOneListItem2 'memory for up character - server relations
End Type
Public Type TypeBuffer
  numbytes As Long
  packet() As Byte
End Type


Public Type TypeItem
  t1 As Byte
  t2 As Byte
  t3 As Byte ' amount
  t4 As Byte ' new strange byte, tibia 9.9
End Type
Public Type TypeBackpack
  open As Boolean
  cap As Long
  used As Long
  name As String
  item(0 To HIGHEST_ITEM_BPSLOT) As TypeItem
End Type

Public Type TypeRuneMakerOptions
  activated As Boolean
  autoEat As Boolean
  ManaFluid As Boolean
  autoLogoutAnyFloor As Boolean
  autoLogoutCurrentFloor As Boolean
  autoLogoutOutOfRunes As Boolean
  autoWaste As Boolean
  msgSound As Boolean
  msgSound2 As Boolean
  firstActionText As String
  firstActionMana As Long
  LowMana As Long
  secondActionText As String
  secondActionMana As Long
  secondActionSoulpoints As Long
End Type

'custom ng
Public Type TypehealingCheatsOptions
exaustEat As Long
sdmax As Boolean
HouseX As Long
HouseY As Long
htarget As Boolean
antipush As Boolean
pmax As Boolean
exaust As Boolean
txtSpellhi As String
txtSpelllo As String
txtPot As String
txtMana As String
txtHealthhi As String
txtHealthlo As String
txtHealpot As String
txtManapot As String
txtManahi As String
txtManalo As String
Combo1 As String
Combo2 As String
End Type

'custom ng extras
Public Type TypeextrasOptions
txtSpell As String
txtMana As String
txtSSA As String
cmbHouse As String
chkMana As Boolean
chkDanger As Boolean
chkPM As Boolean
chkEat As Boolean
chkautoUtamo As Boolean
chkautoHur As Boolean
chkautogHur As Boolean
chkAFK As Boolean
chkGold As Boolean
chkPlat As Boolean
chkDash As Boolean
chkMW As Boolean
chkSSA As Boolean
chkHouse As Boolean
chkTitle As Boolean
End Type

'custom ng persistent
Public Type TypepersistentOptions
txtHk1 As String
txtHk2 As String
txtHk3 As String
txtHk4 As String
txtHk5 As String
txtHk6 As String
txtHk7 As String
txtHk8 As String
txtHk9 As String
txtHk10 As String
txtHk11 As String
txtExiva1 As String
txtExiva2 As String
txtExiva3 As String
txtExiva4 As String
txtExiva5 As String
txtExiva6 As String
txtExiva7 As String
txtExiva8 As String
txtExiva9 As String
txtExiva10 As String
txtExiva11 As String
chkExiva1 As Boolean
chkExiva2 As Boolean
chkExiva3 As Boolean
chkExiva4 As Boolean
chkExiva5 As Boolean
chkExiva6 As Boolean
chkExiva7 As Boolean
chkExiva8 As Boolean
chkExiva9 As Boolean
chkExiva10 As Boolean
chkExiva11 As Boolean
persistvar1 As Long
persistvar2 As Long
persistvar3 As Long
persistvar4 As Long
persistvar5 As Long
persistvar6 As Long
persistvar7 As Long
persistvar8 As Long
persistvar9 As Long
persistvar10 As Long
persistvar11 As Long
End Type

'custom ng aimbot
Public Type TypeaimbotOptions
chkSDcombo As Boolean
chkUEcombo As Boolean
txtLeader As String
txtCombo As String
End Type

Public gISIDE As Boolean
Public TrainerTimer1 As Long
Public TrainerTimer2 As Long
Public initialRuneBackpack() As Byte
Public FirstExecute As Boolean
Public DoingMainLoop() As Boolean
Public DoingMainLoopLogin() As Boolean
Public SendingSpecialOutfit() As Boolean
Public RuneMakerOptions() As TypeRuneMakerOptions
'custom ng
Public healingCheatsOptions() As TypehealingCheatsOptions
Public extrasOptions() As TypeextrasOptions
Public persistentOptions() As TypepersistentOptions
Public aimbotOptions() As TypeaimbotOptions
'
Public ConnectionBuffer() As TypeBuffer
Public ConnectionBufferLogin() As TypeBuffer


Public CharacterList As TypeCharacterList
Public CharacterList2() As TypeCharacterList2

Public shouldOpenErrorsTXTfolder As Boolean

Public Connected() As Boolean
Public nextLight() As String
Public GameConnected() As Boolean
Public MustCheckFirstClientPacket() As Boolean
Public LastNumTibiaClients As Long
' ips are stored as an array of bytes starting in next positions:
' (values loaded at frmMain.load from file)
Public memLoginServer() As Long

' port numbers are stored as long in next positions:
' (values loaded at frmMain.load from file)
Public MemPortLoginServer() As Long
Public XTEAoption As Long
Public LoginServerStartPointer As Long
Public LoginServerStep As Long
Public HostnamePointerOffset As Long
Public IPAddressPointerOffset As Long
Public PortOffset As Long
Public proxyChecker As Long
Public tibiaEntryServer As String




Public fakemessagesLevel As Long
Public NeedToIgnoreFirstGamePacket() As Boolean
Public ClosedBoard As Boolean
Public CanceledBoard As Boolean
Public VisibleAdvancedOptions As Boolean
Public LightIntesityHex As String
Public BlockUnload As Integer
Public MapWantedOnTop As Boolean
Public Backpack() As TypeBackpack
Public bpIDselected As Long
Public runemakerIDselected As Long
'custom ng
Public healingIDselected As Long
Public extrasIDselected As Long
Public persistentIDselected As Long
Public aimbotIDselected As Long
Public blnShowAdvancedOptions2 As Long
'
Public LoadWasCompleted As Boolean
Public MAXCLIENTS As Long
'Public UseRealTibiaDatInLatestTibiaVersion As Boolean
Public HIGHEST_BP_ID As Long
Public blnShowAdvancedOptions As Long
Public posSpamActivated() As Boolean
Public posSpamChannelB1() As Byte
Public posSpamChannelB2() As Byte
Public getSpamActivated() As Boolean
Public getSpamChannelB1() As Byte
Public getSpamChannelB2() As Byte
Public makingRune() As Boolean
Public fastIDreason As Integer
Public fastCounter As Long
Public executingCavebot() As Boolean
Public SpeedDist As Long
Public GotKillOrderTargetID() As Double
Public GotKillOrder() As Boolean
Public GotKillOrderTargetName() As String
Public AllowUHpaused() As Boolean
Public SpamAutoFastHeal() As Boolean
Public nextFastHeal() As Long
Public logoutAllowed() As Long
Public IgnoreServer() As Boolean
Public FirstCharInCharList() As String
Public NoHealingNextTurn() As Boolean
Public DropDelayerTurn() As Long
Public IamAdmin As Boolean

Public lngNextScreenshotNumber As Long

'tileIDs will change with new tibia version
'runes
Public tileID_Blank As Long
Public tileID_WallBugItem As Long
Public tileID_SD As Long
Public tileID_HMM As Long
Public tileID_Explosion As Long
Public tileID_IH As Long
Public tileID_UH As Long

Public tileID_fireball As Long
Public tileID_stalagmite As Long
Public tileID_icicle As Long

'items
Public tileID_Bag As Long
Public tileID_Backpack As Long
Public tileID_Oracle As Long
Public tileID_FishingRod As Long
Public tileID_Rope As Long
Public tileID_LightRope As Long
Public tileID_Shovel As Long
Public tileID_LightShovel As Long

'water
Public tileID_waterEmpty As Long
Public tileID_waterWithFish As Long
Public tileID_waterEmptyEnd As Long
Public tileID_waterWithFishEnd As Long

Public TimesWarnedAboutRelog As Long

' blocking objects
Public tileID_blockingBox As Long

' to up floor
Public tileID_rampToNorth As Long
Public tileID_rampToSouth As Long
Public tileID_ladderToUp As Long
Public tileID_holeInCelling As Long
Public tileID_stairsToUp As Long
Public tileID_woodenStairstoUp As Long

Public tileID_desertRamptoUp As Long

Public tileID_rampToRightCycMountain As Long
Public tileID_rampToLeftCycMountain As Long

Public tileID_jungleStairsToNorth As Long
Public tileID_jungleStairsToLeft As Long


' to down
Public tileID_grassCouldBeHole As Long
Public tileID_pitfall As Long
Public tileID_openHole As Long
Public tileID_openHole2 As Long
Public tileID_trapdoor As Long
Public tileID_trapdoor2 As Long
Public tileID_sewerGate As Long
Public tileID_stairsToDown As Long
Public tileID_stairsToDown2 As Long
Public tileID_woodenStairstoDown As Long
Public tileID_rampToDown As Long
Public tileID_closedHole As Long
Public tileID_desertLooseStonePile As Long
Public tileID_OpenDesertLooseStonePile As Long
Public tileID_trapdoorKazordoon As Long
Public tileID_stairsToDownKazordoon As Long
Public tileID_stairsToDownThais As Long
Public tileID_down1 As Long
Public tileID_down2 As Long
Public tileID_down3 As Long

'FOOD
Public tileID_firstFoodTileID As Long
Public tileID_lastFoodTileID As Long
Public tileID_firstMushroomTileID As Long
Public tileID_lastMushroomTileID As Long


'FIELD RANGE1
Public tileID_firstFieldRangeStart As Long
Public tileID_firstFieldRangeEnd As Long
Public tileID_secondFieldRangeStart As Long
Public tileID_secondFieldRangeEnd As Long

Public tileID_campFire1 As Long
Public tileID_campFire2 As Long

'WALKABLE FIELDS
Public tileID_walkableFire1 As Long
Public tileID_walkableFire2 As Long
Public tileID_walkableFire3 As Long

'inside depot item
Public tileID_depotChest As Long

' flasks mana
Public tileID_flask As Long

Public tileID_health_potion As Long
Public tileID_strong_health_potion As Long
Public tileID_small_health_potion As Long
Public tileID_great_health_potion As Long
Public tileID_mana_potion As Long
Public tileID_strong_mana_potion As Long
Public tileID_great_mana_potion As Long

Public tileID_ultimate_health_potion As Long
Public tileID_great_spirit_potion As Long

Public byteNothing As Byte
Public byteMana As Byte
Public byteLife As Byte

'tray icon


Public nid As NOTIFYICONDATA


Public Antibanmode As Long
'runemaker
Public lock_chkActivate As Boolean
Public lock_chkFood  As Boolean
Public lock_chkManaFluid As Boolean
Public lock_chkLogoutDangerAny As Boolean
Public lock_chkLogoutDangerCurrent As Boolean
Public lock_chkLogoutOutRunes As Boolean
Public lock_chkWaste As Boolean
Public lock_chkmsgSound As Boolean
Public lock_chkmsgSound2 As Boolean
'custom ng
Public lock_chkMana As Boolean
Public lock_chkDanger As Boolean
Public lock_chkPM As Boolean
Public lock_chkEat As Boolean
Public lock_chkautoUtamo As Boolean
Public lock_chkautoHur As Boolean
Public lock_chkautogHur As Boolean
Public lock_chkAFK As Boolean
Public lock_chkGold As Boolean
Public lock_chkPlat As Boolean
Public lock_chkDash As Boolean
Public lock_chkMW As Boolean
Public lock_chkSSA As Boolean
Public lock_chkTitle As Boolean
Public lock_chkHouse As Boolean
Public lock_chkExiva1 As Boolean
Public lock_chkExiva2 As Boolean
Public lock_chkExiva3 As Boolean
Public lock_chkExiva4 As Boolean
Public lock_chkExiva5 As Boolean
Public lock_chkExiva6 As Boolean
Public lock_chkExiva7 As Boolean
Public lock_chkExiva8 As Boolean
Public lock_chkExiva9 As Boolean
Public lock_chkExiva10 As Boolean
Public lock_chkExiva11 As Boolean
Public lock_chkSDcombo As Boolean
Public lock_chkUEcombo As Boolean
'

'login servers


Public serverLogoutMessage As String

Public NumberOfLoginServers As Long
Public trueLoginServer() As String
Public trueLoginPort() As String

Public PREFEREDLOGINSERVER As String
Public PREFEREDLOGINPORT As String

Public publicDebugMode As Boolean

Public runeTurn() As Integer

Public PUSHDELAYTIMES As Long '9 by default

Public TibiaVersion As String
Public TibiaVersionLong As Long

Public LoadingStarted As Boolean
Public CornerMessage As String
Public CornerColor As Long
Public returnValue As VbMsgBoxResult
Public BlueAuraDelay As Long
Public ReconnectionStage() As Long
Public ReconnectionPacket() As TypeBuffer

Public var_expleft() As String
Public var_nextlevel() As String
Public var_exph() As String
Public var_timeleft() As String
Public var_played() As String
Public var_expgained() As String
Public var_lf() As String
Public ExivaExpPlace As String
Public thisShouldNotBeLoading As Integer

Public firstValidOutfit As Long
Public lastValidOutfit As Long

Public configPath As String
' Public VarProtection2 As Long
Public extremeDebugMode As Boolean
Public reconnectionRetryCount() As Long
Public nextReconnectionRetry() As Long

Public LimitedToServer As String
Public GLOBAL_RUNEHEAL_HP As Long
Public gotDictErr As Long
Public RecordLogin As Boolean
'Public VarProtection3 As Long
Public CurrBlackdServer As String
Public CurrBlackdServer_folder As String
'Public VarProtection4 As Long

Public ValueOfUservar As Scripting.Dictionary  ' A dictionary Uservar (string) -> name (string)
Public lastUsedChannelID() As String

Public mustSendFirstWhenConnected() As TypeInitialPacket
'Public VarProtection5 As Long
Public lastRecChannelID() As String


Public fakemessagesLevel1 As Byte
'Public VarProtection6 As Long
Public fakemessagesLevel2 As Byte

'Public VarProtection1 As Long
'Public VarProtection7 As Long

Public confirmedExit As Boolean

Public tibiaclassname As String
Public tibiamainname As String

Public LastFasterLogin As String
Public AlreadyCheckingFasterLogin As String

'blackd ng custom var
Public MustUnload As Boolean



Public ProcessidIPrelations As Scripting.Dictionary
Public ProcessidAccountRelations As Scripting.Dictionary
Public IgnoredCreatures As Scripting.Dictionary
Public ConnectionSignal() As Boolean
Public TOOSLOWLOGINSERVER_MS As Long

Public usingPriorities() As Boolean
Public broadcastIDselected As Long
Public currentBroadcastIndex As Long
Public BroadcastDelay1 As Long
Public BroadcastDelay2 As Long
Public BroadcastMC As Long

Public LAST_BATTLELISTPOS As Long

Public CurrentTibiaDatPath As String
Public CurrentTibiaDatDATE As Date
Public CurrentTibiaDatVERSION As Long
Public CurrentTibiaDatFILE As String
Public MyErrorDate As Date

Public configOverrideByCommand As Boolean
Public dateErrDescription As String

Public DefaultTibiaFolder As String

Public OVERWRITE_CONFIGPATH As String
Public OVERWRITE_CLIENT_PATH As String
Public OVERWRITE_MAPS_PATH As String
Public OVERWRITE_OT_MODE As Boolean
Public OVERWRITE_OT_IP As String
Public OVERWRITE_OT_PORT As Long
Public OVERWRITE_SHOWAGAIN As Boolean

Public MemoryProtectedMode As Boolean
Public ForceDisableEncryption As Boolean
Public CloseLoginServerAfterCharList As Boolean

Public Function MemoryChangeFloor(idConnection As Integer, relfloornumber As String) As Long 'receives mc id and relative floor increase desired
    On Error GoTo goterr
    Dim floornumber As Long
    Dim pid As Long
    Dim relChange As Long
    Dim ammountOfChanges As Long
    Dim i As Long
    If (TibiaVersionLong >= 1100) Then
        MemoryChangeFloor = -1
        Exit Function
    End If
    'MemoryChangeFloor = -1 ' not working yet
    'Exit Function
    If IsNumeric(relfloornumber) = False Then
        MemoryChangeFloor = -1 'failure (bad parameters)
        Exit Function
    End If
    relChange = CLng(relfloornumber)
    ammountOfChanges = Abs(relChange)
    levelSpy_Off idConnection
    If ammountOfChanges > 0 Then
        Call WriteNops(idConnection, LEVELSPY_NOP, 2)
        'Initialize Level spying
        LevelSpy_Init idConnection
        'Set boolean
        bLevelSpy(idConnection) = True
        
        'full light
        Call WriteNops(idConnection, LIGHT_NOP, 2)
        Call writeBytes(idConnection, LIGHT_AMOUNT, 255, 1)
        
        
    End If
    For i = 1 To ammountOfChanges
        If relChange > 0 Then
            levelSpy_Down idConnection
        Else
            levelSpy_Up idConnection
        End If
    Next i
    MemoryChangeFloor = 0 ' sucess
    Exit Function
goterr:
    MemoryChangeFloor = -1 ' failure (unknown)
End Function




Public Sub levelSpy_Off(idConnection As Integer)
'disable level spying by restoring default values
Call writeBytes(idConnection, LEVELSPY_NOP, LEVELSPY_NOP_DEFAULT, 2)
Call writeBytes(idConnection, LEVELSPY_ABOVE, LEVELSPY_ABOVE_DEFAULT, 1)
Call writeBytes(idConnection, LEVELSPY_BELOW, LEVELSPY_BELOW_DEFAULT, 1)
'Set boolean
bLevelSpy(idConnection) = False
End Sub
 
Public Sub WriteNops(idConnection As Integer, address As Long, Nops As Integer)

'Get Process Handle
Dim ProcessHandle As Long
GetProcessIDs idConnection
ProcessHandle = ProcessID(idConnection)

'Write Memory
Dim i, j As Integer
i = 0: j = 0
For i = 1 To Nops
Const nop = &H90
Memory_WriteByte address + j, nop, ProcessHandle
j = j + 1
Next i
'Close process handle

End Sub

Private Sub writeBytes(idConnection As Integer, address As Long, value As Long, byteS As Integer)
'Get Process Handle
Dim ProcessHandle As Long
GetProcessIDs idConnection
ProcessHandle = ProcessID(idConnection)
'write to memory
If byteS = 1 Then
  'Debug.Print "Writting 1 byte [" & CStr(ProcessHandle) & "] at address & " & CStr(Address) & " :" & CStr(Value)
  Memory_WriteByte address, CByte(value), ProcessHandle
Else
  'Debug.Print "Writting 2 byte [" & CStr(ProcessHandle) & "] at address & " & CStr(Address) & " :" & CStr(Value)
  Memory_WriteByte address, LowByteOfLong(value), ProcessHandle
  Memory_WriteByte address + 1, HighByteOfLong(value), ProcessHandle
End If
End Sub

'Initialize level spying
Public Sub LevelSpy_Init(idConnection As Integer)
'Get player Z
Dim playerZ As Integer
playerZ = readBytes(idConnection, PLAYER_Z, 1)
'Set levelspy to current level
If (playerZ <= Z_AXIS_DEFAULT) Then
    'Above ground
    Call writeBytes(idConnection, LEVELSPY_ABOVE, Z_AXIS_DEFAULT - playerZ, 1)
    Call writeBytes(idConnection, LEVELSPY_BELOW, LEVELSPY_BELOW_DEFAULT, 1)
Else
    'Below Ground
    Call writeBytes(idConnection, LEVELSPY_ABOVE, LEVELSPY_ABOVE_DEFAULT, 1)
    Call writeBytes(idConnection, LEVELSPY_BELOW, LEVELSPY_BELOW_DEFAULT, 1)
End If
End Sub

'Increase spy level
Public Sub levelSpy_Up(idConnection As Integer)
'Levelspy must be on

'Get player z
Dim playerZ As Integer
playerZ = readBytes(idConnection, PLAYER_Z, 1)
'Ground level
Dim groundLevel As Long
groundLevel = 0
If playerZ <= Z_AXIS_DEFAULT Then
    groundLevel = LEVELSPY_ABOVE ' above ground
Else
    groundLevel = LEVELSPY_BELOW ' below ground
End If
    
'Get Current level
Dim currentLevel As Integer
currentLevel = readBytes(idConnection, groundLevel, 1)
If currentLevel >= LEVELSPY_MAX Then
    Call writeBytes(idConnection, groundLevel, LEVELSPY_MIN, 1) ' Loop back to start
Else
    Call writeBytes(idConnection, groundLevel, currentLevel + 1, 1) ' increase spy level
    
'Set statusbar
'setStatusBar ("Level Spy: Up")
End If
End Sub

'Decrease spy level
Public Sub levelSpy_Down(idConnection As Integer)
'Levelspy must be on
If bLevelSpy(idConnection) = False Then
'setStatusBar ("Please Enable Level Spy first!")
Exit Sub
End If
'Get player z
Dim playerZ As Integer
playerZ = readBytes(idConnection, PLAYER_Z, 1)
'Ground level
Dim groundLevel As Long
groundLevel = 0
If playerZ <= Z_AXIS_DEFAULT Then
    groundLevel = LEVELSPY_ABOVE ' above ground
Else
    groundLevel = LEVELSPY_BELOW ' below ground
End If
    
'Get Current level
Dim currentLevel As Integer
currentLevel = readBytes(idConnection, groundLevel, 1)
If currentLevel <= LEVELSPY_MIN Then
    Call writeBytes(idConnection, groundLevel, LEVELSPY_MAX, 1) ' Loop back to start
Else
    Call writeBytes(idConnection, groundLevel, currentLevel - 1, 1) ' increase spy level
    
'Set statusbar
'setStatusBar ("Level Spy: Down")
End If
End Sub

Public Function readBytes(idConnection As Integer, address As Long, byteS As Integer) As Long
'Get Process Handle
Dim ProcessHandle As Long
Dim b1 As Byte
Dim b2 As Byte
GetProcessIDs idConnection
ProcessHandle = ProcessID(idConnection)

'read memory
Dim Buffer As Long
Buffer = 0
If byteS = 1 Then
    readBytes = Memory_ReadByte(address, ProcessHandle)
Else
    b1 = Memory_ReadByte(address, ProcessHandle)
    b2 = Memory_ReadByte(address + 1, ProcessHandle)
    readBytes = GetTheLong(b1, b2)
End If
'Call ReadProcessMemory(processHandle, Address, buffer, bytes, 0)
'Close handle
'CloseHandle (processHandle)
'readBytes = buffer
End Function

Public Sub AddProcessIdIPrelation(strIP As String, strProcessID As Long)
  ' add item to dictionary
  ProcessidIPrelations.item(strIP) = strProcessID
End Sub
Public Sub ResetProcessidIPrelations()
  On Error GoTo goterr
  Dim a As Long
  a = 0
  ProcessidIPrelations.RemoveAll
  Exit Sub
goterr:
  a = -1
End Sub
Public Function GetProcessIdFromIP(strIP As String) As Long
  ' get the name from an ID
  Dim aRes As Long
  Dim res As Boolean
  If ProcessidIPrelations.Exists(strIP) = True Then
    GetProcessIdFromIP = ProcessidIPrelations.item(strIP)
  Else
    GetProcessIdFromIP = 0 ' error
  End If
End Function



'Public Sub AddProcessIdAccountRelation(strAccount As String, strProcessID As Long)
'  ' add item to dictionary
'  ProcessidAccountRelations.item(strAccount) = strProcessID
'End Sub
'Public Sub ResetProcessidAccountRelations()
'  ProcessidAccountRelations.RemoveAll
'End Sub
'Public Function GetProcessIdFromAccount(strAccount As String) As Long
'  ' get the name from an ID
'  Dim aRes As Long
'  Dim res As Boolean
'  If ProcessidIPrelations.Exists(strAccount) = True Then
'    GetProcessIdFromAccount = ProcessidIPrelations.item(strAccount)
'  Else
'    GetProcessIdFromAccount = 0 ' error
'  End If
'End Function

Public Sub OverwriteOnFileSimple(file_name As String, strText As String)
  Dim fn As Integer
  Dim writeThis As String
  Dim a As Long
  On Error GoTo ignoreit
  a = 0
  fn = FreeFile
    writeThis = strText
  Open App.Path & "\" & file_name For Output As #fn
    Print #fn, writeThis
  Close #fn

  Exit Sub
ignoreit:
  a = -1
End Sub

Public Sub AddwriteOnFileSimple(file_name As String, strText As String)
  Dim fn As Integer
  Dim writeThis As String
  Dim a As Long
  On Error GoTo ignoreit
  a = 0
  fn = FreeFile
    writeThis = strText
  Open App.Path & "\" & file_name For Append As #fn
    Print #fn, writeThis
  Close #fn

  Exit Sub
ignoreit:
  a = -1
End Sub

Public Sub AddUserVar(ByVal strUservar As String, ByVal strValue As String)
  On Error GoTo goterr
  ' add item to dictionary
  Dim res As Boolean
  ValueOfUservar.item(strUservar) = strValue
  Exit Sub
goterr:
  LogOnFile "errors.txt", "Get error at AddUserVar : " & Err.Description
End Sub

Public Function GetUserVar(ByVal strUservar As String) As String
  On Error GoTo goterr
  ' get the IPandport from server name
  Dim aRes As String
  Dim res As Boolean
  If ValueOfUservar.Exists(strUservar) = True Then
    GetUserVar = ValueOfUservar.item(strUservar)
  Else
    GetUserVar = ""
  End If
  Exit Function
goterr:
  LogOnFile "errors.txt", "Got error at AddUserVar : " & Err.Description
  GetUserVar = ""
End Function

Public Sub ChangeGLOBAL_RUNEHEAL_HP(newValue As Long)
  Dim i As Integer
  Dim aRes As Long
  Dim oldVal As Long
  oldVal = GLOBAL_RUNEHEAL_HP
  frmHardcoreCheats.lblHPvalue.Caption = CStr(newValue) & " %"
  GLOBAL_RUNEHEAL_HP = newValue
  If frmHardcoreCheats.scrollHP.value <> newValue Then
    frmHardcoreCheats.scrollHP.value = newValue
  End If
  If oldVal <> GLOBAL_RUNEHEAL_HP Then
  For i = 1 To MAXCLIENTS
    If (GameConnected(i) = True) And (ReconnectionStage(i) = 0) And (sentWelcome(i) = True) Then
      aRes = SendLogSystemMessageToClient(i, "BlackdProxy: The autoruneheal was changed to " & CStr(GLOBAL_RUNEHEAL_HP) & " %")
      DoEvents
    End If
  Next i
  End If
End Sub

Public Sub enLight(i As Integer)
  Dim inRes As Integer
  Dim cPacket() As Byte
  #If FinalMode Then
  On Error GoTo ignoreit
  #End If
  ' 8D 8E 20 AE 01 0F D7 00 00 00 00 00 00 00 00 00
  inRes = GetCheatPacket(cPacket, "07 00 8D " & IDstring(i) & " " & LightIntesityHex & " " & nextLight(i))
  frmMain.UnifiedSendToClientGame i, cPacket
  DoEvents
  Exit Sub
ignoreit:
  '..
End Sub
         
' For internal use.
' Return the class name of the specified window
' Example: MsgBox GetWindowClass(Me.hWnd)
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
          
Public Sub ConfigurePath(phwnd As Long, isfrmmain As Boolean)
 Dim res As String
 res = BrowseForFolder(phwnd, "Select tibia map folder (usually on " & cte_automapfolder & ")")
 If res <> "" Then
    If ((TibiaVersionLong >= 800) And (LCase(Right$(res, 7)) <> "automap")) Then
        'MsgBox "The folder you selected is not valid", vbOKOnly + vbExclamation, "Please do this first"
        Exit Sub
    End If
   TibiaPath = res
   If isfrmmain = True Then
     frmMain.txtTibiaPath.Text = res
   End If
 End If
End Sub

Public Function TryAutoPath() As String
    On Error GoTo cantdoit
    Const ParTibiaFolder As String = "Tibia"
    Dim strAppdata As String
    Dim strProposal As String
    Dim strProp2 As String
    Dim fs As Scripting.FileSystemObject
        
    If TibiaVersionLong >= 1100 Then
        Set fs = New Scripting.FileSystemObject
        strAppdata = GetLocalApplicationDataFolder()
        strProposal = strAppdata & "\" & ParTibiaFolder & "\packages\Tibia\minimap"
     
        If fs.FolderExists(strProposal) = True Then
            Set fs = Nothing
            TryAutoPath = strProposal
            Exit Function
        Else
            Set fs = Nothing
            TryAutoPath = ""
            Exit Function
        End If
    ElseIf TibiaVersionLong >= 800 Then
        Set fs = New Scripting.FileSystemObject
        strAppdata = GetAppDataFolder()
         strProposal = strAppdata & "\" & ParTibiaFolder & "\Automap"
        strProp2 = strAppdata & "\" & ParTibiaFolder
        If fs.FolderExists(strProposal) = True Then
            Set fs = Nothing
            TryAutoPath = strProposal
            Exit Function
        ElseIf fs.FolderExists(strProp2) = True Then
            fs.CreateFolder strProposal
            Set fs = Nothing
            TryAutoPath = strProposal
        Else
            Set fs = Nothing
            TryAutoPath = ""
            Exit Function
        End If
    Else
        TryAutoPath = TibiaExePath
    End If
    Exit Function
cantdoit:
    TryAutoPath = ""
End Function

Public Sub givePathMsg(thehwnd As Long)
    Dim trythis As String
    If (TibiaPath = "") Or ((TibiaVersionLong >= 800) And (LCase(Right$(TibiaPath, 7)) <> "automap")) Then
        If ((TibiaVersionLong < 800) And (TibiaPath <> "")) Then
            Exit Sub
        End If
        trythis = TryAutoPath()
        If (trythis = "") Then
            MsgBox "Select tibia map folder (usually on " & cte_automapfolder & " )" & vbCrLf & vbCrLf & _
            "What to do if you don't see the folder:" & vbCrLf & _
            "1. Play Tibia 8.00+ at least one time. Then close Tibia. This will make the folder." & vbCrLf & _
            "2. Unhide special folders : folder options > view > check 'show hidden files and folders' " & vbCrLf & _
            "3. Restart blackd proxy so the folder browser gets updated and after that you should be able to browse it at " & _
            vbCrLf & cte_automapfolder & vbCrLf & vbCrLf & _
            "Note that the exact path depends on your windows user name!", vbOKOnly + vbExclamation, "Please do this first"
            ConfigurePath thehwnd, False
            Exit Sub
        Else
            'MsgBox "Blackd Proxy detected that your maps are now here:" & vbCrLf & _
             trythis, vbOKOnly + vbInformation, "Just for your information"
            TibiaPath = trythis
        End If
    End If
End Sub

Public Function ValidateTibiaPath(str As String) As String
  Dim res As String
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  If TibiaVersionLong >= 1100 Then
    If TibiaPath = "" Then
        res = ""
    ElseIf LCase(Right(str, 7)) <> "minimap" Then
        res = "PATH NOT CONFIGURED! USE THIS BUTTON TO BROWSE -->"
    Else
        res = str
    End If
    ValidateTibiaPath = res
  ElseIf TibiaVersionLong >= 800 Then
    If TibiaPath = "" Then
    res = ""
    ElseIf LCase(Right(str, 7)) <> "automap" Then
    res = "PATH NOT CONFIGURED! USE THIS BUTTON TO BROWSE -->"
    Else
    res = str
    End If
    ValidateTibiaPath = res
  Else
  res = str
  
'  Dim fso As Scripting.FileSystemObject
'  Set fso = New Scripting.FileSystemObject
'  If (fso.FileExists(str & "\tibia.exe") = True) Or (str = "PATH NOT CONFIGURED! USE THIS BUTTON TO BROWSE -->") Then
'    res = str
'  Else
'    If TibiaPath = "" Then
'       res = "PATH NOT CONFIGURED! USE THIS BUTTON TO BROWSE -->"
'    ElseIf (fso.FileExists(TibiaPath & "\tibia.exe") = True) Then
'       res = TibiaPath
'    Else
'      res = ""
'    End If
'  End If
  ValidateTibiaPath = res
  
  End If
  Exit Function
goterr:
  ValidateTibiaPath = "PATH NOT CONFIGURED! USE THIS BUTTON TO BROWSE -->"
End Function

Public Function Hexarize(strinput As String) As String
  Dim strByte As String
  Dim res As String
  res = ""
  While Len(strinput) > 0
    strByte = Left(strinput, 1)
    strinput = Right(strinput, Len(strinput) - 1)
    res = res & GoodHex(Asc(strByte)) & " "
  Wend
  Hexarize = res
End Function

' url encodes a string
'warning: unicode is untested
'creds: http://www.vbforums.com/showthread.php?334645-Winsock-Making-HTTP-POST-GET-Requests
Function URLEncode(ByVal str As String) As String
        Dim intLen As Integer
        Dim x As Integer
        Dim curChar As Long
                Dim newStr As String
                intLen = Len(str)
        newStr = ""
                        For x = 1 To intLen
            curChar = Asc(Mid$(str, x, 1))
            
            If (curChar < 48 Or curChar > 57) And _
                (curChar < 65 Or curChar > 90) And _
                (curChar < 97 Or curChar > 122) Then
                                newStr = newStr & "%" & Hex(curChar)
            Else
                newStr = newStr & Chr(curChar)
            End If
        Next x
        
        URLEncode = newStr
End Function

Public Function Hexarize2(strinput As String) As String
  Dim strByte As String
  Dim res As String
  Dim bcount As Long
  bcount = 0
  res = ""
  While Len(strinput) > 0
    strByte = Left(strinput, 1)
    strinput = Right(strinput, Len(strinput) - 1)
    res = res & GoodHex(Asc(strByte)) & " "
    bcount = bcount + 1
  Wend
  res = GoodHex(LowByteOfLong(bcount)) & " " & GoodHex(HighByteOfLong(bcount)) & " " & res
  Hexarize2 = res
End Function



Public Sub ToggleTopmost(ByVal hWindow As Long, b As Boolean)
  Dim hw As Long
  If b = False Then
    SetWindowPos hWindow, HWND_NOTOPMOST, 0, 0, 0, 0, _
     SWP_NOMOVE Or SWP_NOSIZE
  Else
    SetWindowPos hWindow, HWND_TOPMOST, 0, 0, 0, 0, _
     SWP_NOMOVE Or SWP_NOSIZE
  End If
End Sub


Public Sub AddCharServer(charName As String, ServerName As String, serverIP1 As Byte, _
 serverIP2 As Byte, serverIP3 As Byte, serverIP4 As Byte, serverPort As Long)
  Dim nextPlace As Integer
  Dim currentPlace As Integer
  Dim i As Integer
  '3 cases
  ' case 1: character name already on list, update info
  currentPlace = 0
  For i = 1 To CharacterList.numItems
    If CharacterList.item(i).CharacterName = charName Then
       currentPlace = i
       Exit For
    End If
  Next i
  If currentPlace = 0 Then
    ' case 2: character name not on list, list empty, add it
    nextPlace = CharacterList.numItems + 1
    If nextPlace <= MAXLOGINMEMORY Then
      CharacterList.numItems = nextPlace
      currentPlace = nextPlace
    Else
      ' case 3: character name not on list, but list is full
      ' overwrite the pointer position
      currentPlace = CharacterList.pointer
      CharacterList.pointer = CharacterList.pointer + 1
      If CharacterList.pointer = MAXLOGINMEMORY + 1 Then
        CharacterList.pointer = 1
      End If
    End If
  End If

  CharacterList.item(currentPlace).CharacterName = charName
  CharacterList.item(currentPlace).ServerName = ServerName
  CharacterList.item(currentPlace).serverIP1 = serverIP1
  CharacterList.item(currentPlace).serverIP2 = serverIP2
  CharacterList.item(currentPlace).serverIP3 = serverIP3
  CharacterList.item(currentPlace).serverIP4 = serverIP4
  CharacterList.item(currentPlace).serverPort = serverPort
  
End Sub
Public Sub ResetCharServer()
  CharacterList.numItems = 0
  CharacterList.pointer = 1
End Sub

Public Function GetCharListPosition(ByRef packet() As Byte, ByRef selectedcharacter As String) As Integer
  ' get the list position of the selected character
  #If FinalMode Then
  On Error GoTo returnTheResult
  #End If
  Dim res As Integer
  Dim lon As Long
  Dim i As Long
  res = -1 'error
  If packet(2) <> &HA Then
    res = 0
    GoTo returnTheResult 'this is not a character list packet
  End If
  lon = GetTheLong(packet(12), packet(13))
  selectedcharacter = ""
  For i = 14 To 13 + lon
    selectedcharacter = selectedcharacter & Chr(packet(i))
  Next i
  ' compare through the character list
  res = 0
  For i = 1 To MAXLOGINMEMORY
    If selectedcharacter = CharacterList.item(i).CharacterName Then
      res = i
      Exit For
    End If
  Next i
returnTheResult:
  GetCharListPosition = res
End Function



Public Function PacketIPchange(ByRef packet() As Byte) As Integer
  Dim lon As Long
  Dim motd As Long
  Dim numChars As Long
  Dim lonCName As Long
  Dim lonSName As Long
  Dim i As Integer
  Dim j As Integer
  Dim pos As Long
  Dim servName As String
  Dim servIP1 As Byte
  Dim servIP2 As Byte
  Dim servIP3 As Byte
  Dim servIP4 As Byte
  Dim servPort As Long
  Dim charName As String
  Dim hb As Byte
  Dim lb As Byte
  Dim res As Integer
  #If FinalMode Then
 On Error GoTo returnTheResult
  #End If
  'OverwriteOnFile "test.txt", frmMain.showAsStr2(packet, 0)
  res = -1 'error
  If frmMain.chckAlter.value = 0 Then
    res = 1
    GoTo returnTheResult 'proxy user don't want to change this packet
  End If
  If packet(2) <> &H14 Then
    'MsgBox "1.received " & GoodHex(packet(2)), vbOKOnly + vbInformation, "DEBUG"
    res = 0
    GoTo returnTheResult 'this is not a list of character packet
  End If
  lon = GetTheLong(packet(0), packet(1))
  motd = GetTheLong(packet(3), packet(4))
  numChars = CLng(packet(motd + 6))
  pos = motd + 7
  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Client can select this characters:"
  For i = 1 To numChars 'read all the characters on the list
    lonCName = GetTheLong(packet(pos), packet(pos + 1))
    pos = pos + 2
    charName = ""
    For j = 1 To lonCName
      charName = charName & Chr(packet(pos))
      pos = pos + 1
    Next j
    lonSName = GetTheLong(packet(pos), packet(pos + 1))
    pos = pos + 2
    'read server name of character i
    servName = ""
    For j = 1 To lonSName
      servName = servName & Chr(packet(pos))
      pos = pos + 1
    Next j
    ' save IP
    servIP1 = packet(pos)
    servIP2 = packet(pos + 1)
    servIP3 = packet(pos + 2)
    servIP4 = packet(pos + 3)
    ' change IP
    packet(pos) = 127
    packet(pos + 1) = 0
    packet(pos + 2) = 0
    packet(pos + 3) = 1
    ' save port
    servPort = CLng(packet(pos + 5)) * 256 + CLng(packet(pos + 4))
    ' change port
    ' split the port into high and low bytes
    hb = HighByteOfLong(CLng(frmMain.txtClientGameP.Text))
    lb = LowByteOfLong(CLng(frmMain.txtClientGameP.Text))
    packet(pos + 4) = lb
    packet(pos + 5) = hb
    pos = pos + 6
    ' add the relation of character name - server data in the list :
   ' frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & charName
    AddCharServer charName, servName, servIP1, servIP2, servIP3, servIP4, servPort
  Next i
  ' 10000 days premium account mirage ;)
  ' packet(lon) = &H10
  ' packet(lon + 1) = &H27
  res = 1
returnTheResult:
  PacketIPchange = res
End Function



Public Sub AddCharServer2(idConnection As Integer, charName As String, ServerName As String, serverIP1 As Byte, _
 serverIP2 As Byte, serverIP3 As Byte, serverIP4 As Byte, serverPort As Long, Optional ByVal serverDOMAIN As String = "")
  Dim nextPlace As Integer
  Dim currentPlace As Integer
  If CharacterList2(idConnection).numItems = MAXLOGINMEMORY Then
    frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "CHARLIST ERROR: TOO MANY CHARACTERS" & vbCrLf
    Exit Sub
  End If
  currentPlace = CharacterList2(idConnection).numItems
  CharacterList2(idConnection).item(currentPlace).CharacterName = charName
  CharacterList2(idConnection).item(currentPlace).ServerName = ServerName
  CharacterList2(idConnection).item(currentPlace).serverIP1 = serverIP1
  CharacterList2(idConnection).item(currentPlace).serverIP2 = serverIP2
  CharacterList2(idConnection).item(currentPlace).serverIP3 = serverIP3
  CharacterList2(idConnection).item(currentPlace).serverIP4 = serverIP4
  CharacterList2(idConnection).item(currentPlace).serverPort = serverPort
  CharacterList2(idConnection).item(currentPlace).serverDOMAIN = serverDOMAIN
  CharacterList2(idConnection).numItems = CharacterList2(idConnection).numItems + 1
End Sub
Public Sub ResetCharList2(idConnection As Integer)
  CharacterList2(idConnection).numItems = 0
End Sub







Public Sub DisableBoardButtons()
  frmCheats.cmdToAscii.enabled = False
  frmCheats.cmdToHex.enabled = False
  frmCheats.cmdCountBytes.enabled = False
  frmCheats.cmdOpenBoard.enabled = False
  frmCheats.cmdSendHex.enabled = False
  frmCavebot.cmdLoadCopyPaste.enabled = False
End Sub

Public Sub EnableBoardButtons()
  frmCheats.cmdToAscii.enabled = True
  frmCheats.cmdToHex.enabled = True
  frmCheats.cmdCountBytes.enabled = True
  frmCheats.cmdOpenBoard.enabled = True
  frmCheats.cmdSendHex.enabled = True
  frmCavebot.cmdLoadCopyPaste.enabled = True
End Sub

Public Sub wait(ByVal dblMilliseconds As Double)
  ' (just for tests)
  Dim dblStart As Double
  Dim dblEnd As Double
  Dim dblTickCount As Double
  dblTickCount = GetTickCount()
  dblStart = GetTickCount()
  dblEnd = GetTickCount() + dblMilliseconds
  Do
    DoEvents
    dblTickCount = GetTickCount()
  Loop Until dblTickCount > dblEnd Or dblTickCount < dblStart
End Sub


Public Function CountTibiaWindows() As Long
  Dim tibiaclient As Long
  'Dim hWndDesktop As Long
  Dim countt As Long
  countt = 0
  'hWndDesktop = GetDesktopWindow()
  tibiaclient = 0
  Do

    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      Exit Do
    Else
      countt = countt + 1
    End If
   
  Loop
  CountTibiaWindows = countt
End Function
Public Function ConvStrToByte(str As String) As Byte
  'converts an 1 character string into a byte
  Dim res As Byte
  Dim cad As String
  cad = "&H" & Hex(Asc(str))
  res = CLng(cad)
  ConvStrToByte = res
End Function

Public Function readMemoryString(tibiaclient As Long, memPos As Long, Optional maxread As Long = 255, Optional absoluteA As Boolean = False) As String
    Dim b As Byte
    Dim strString As String
    Dim posM As Long
    Dim i As Long

    strString = ""
    posM = 0
    i = 0
    Do
        If i = maxread Then
          readMemoryString = strString
          Exit Function
        End If
        b = Memory_ReadByte(memPos + posM, tibiaclient, absoluteA)
        posM = posM + 1
        i = i + 1
        If b <> 0 Then
            strString = strString & Chr(b)
        End If
    Loop Until b = &H0
    readMemoryString = strString
End Function

Private Function getFirstAvailableIP() As String
    Dim usedIP(1 To 254) As Boolean
    Dim i As Long
    Dim lngIP As Long
    Dim strIDpart As String
    Dim strRes As String
    Dim tibiaclient As Long
    Dim strCurrentIP As String
    For i = 1 To 254
        usedIP(i) = False
    Next i
    Do
        
        tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
        If tibiaclient = 0 Then
            Exit Do
        Else
            strCurrentIP = readMemoryString(tibiaclient, memLoginServer(1))
            If Left$(strCurrentIP, 8) = "127.0.0." Then
                strIDpart = Right$(strCurrentIP, Len(strCurrentIP) - 8)
                lngIP = CLng(strIDpart)
                usedIP(lngIP) = True
            End If
        End If
    Loop
    For i = 1 To 254
        If usedIP(i) = False Then
            strRes = "127.0.0." & CStr(i)
            getFirstAvailableIP = strRes
            Exit Function
        End If
    Next i
    getFirstAvailableIP = "127.0.0.1"
End Function


Public Function ReadyToChangeTibiaIP() As Boolean
    On Error GoTo goterr
    Dim theubound As Long
    theubound = UBound(memLoginServer)
    ReadyToChangeTibiaIP = True
    Exit Function
goterr:
    ReadyToChangeTibiaIP = False
End Function

Public Function readProxyChecker(tibiaclient As Long) As String
  ' only for debugs - to ensure that we are reading correct code
  Dim res As String
  Dim bb As Byte
  Dim i As Integer
  i = 0
  res = ""
  Do
    bb = Memory_ReadByte(proxyChecker + i, tibiaclient, False)
    i = i + 1
    res = res & GoodHex(bb) & " "
  Loop Until bb = &HC3
  'res = Trim$(res)
  'Debug.Print res
  ' In tibia 10.11 correct code is this:
  ' 8B 40 20 80 38 00 74 18 80 78 01 00 74 12 80 78 02 00 74 0C 80 78 03 00 74 06 B8 01 00 00 00 C3
  readProxyChecker = res
End Function

Public Sub nullifyProxyChecker(tibiaclient As Long)
  ' modifies the tibia.exe code at memory so it does not detect that we are connecting through proxy
  Dim bb As Byte
  Dim i As Integer
  Dim packet() As Byte
  Dim ign As Integer
  Dim lastI As Integer
  If TibiaVersionLong >= 1012 Then
  Exit Sub
  End If
  
  ign = GetCheatPacket(packet, "8B 40 20 83 38 00 B8 00 00 00 00 0F 95 D0 C3")
  lastI = UBound(packet)
  For i = 0 To lastI
    bb = packet(i)
   ' Debug.Print GoodHex(bb)
    Memory_WriteByte proxyChecker + i, bb, tibiaclient, False
  Next i
 ' Debug.Print "OK"
End Sub
Public Function OBTAINmemLoginServer(tibiaclient As Long, nm As Integer) As Long
  Dim adrStruct As Long
  Dim incUnit As Long
  Dim adrPointer As Long
  Dim CRC As String
  Dim incTotal As Long
  adrStruct = Memory_ReadLong(LoginServerStartPointer, tibiaclient, False)
  incUnit = LoginServerStep
  incTotal = HostnamePointerOffset + ((nm - 1) * incUnit)
  adrPointer = Memory_ReadLong(adrStruct + incTotal, tibiaclient, True)
  
  #If FinalMode = 0 Then
  CRC = readMemoryString(tibiaclient, adrPointer, 255, True)
  Debug.Print CRC
  #End If
  OBTAINmemLoginServer = adrPointer
End Function

Public Function OBTAINmemIPNumber(tibiaclient As Long, nm As Integer) As Long
  Dim adrStruct As Long
  Dim incUnit As Long
  Dim adrPointer As Long
  Dim incTotal As Long
  Dim CRC As String
  adrStruct = Memory_ReadLong(LoginServerStartPointer, tibiaclient, False)
  incUnit = LoginServerStep
  incTotal = IPAddressPointerOffset + ((nm - 1) * incUnit)
  adrPointer = Memory_ReadLong(adrStruct + incTotal, tibiaclient, True)
  CRC = CLng(Memory_ReadByte(adrPointer + 0, tibiaclient, True)) & "." & _
        CLng(Memory_ReadByte(adrPointer + 1, tibiaclient, True)) & "." & _
        CLng(Memory_ReadByte(adrPointer + 2, tibiaclient, True)) & "." & _
        CLng(Memory_ReadByte(adrPointer + 3, tibiaclient, True))
  #If FinalMode = 0 Then
  Debug.Print "Original numeric IP = " & CRC
  #End If
  ' if you didn't log yet then crc = 0.0.0.0
  ' else you get something
  ' in Tibia 10.11 you would see ip from Amazon like 54.212.249.103
  
  OBTAINmemIPNumber = adrPointer
End Function


Public Function OBTAINmemLoginPort(tibiaclient As Long, nm As Integer) As Long
  Dim adrStruct As Long
  Dim incUnit As Long
 
  Dim CRC As String
  Dim incTotal As Long
  Dim b1 As Byte
  Dim b2 As Byte
  Dim tport As Long
  Dim adrPort As Long
  Dim debugi As Long
  adrStruct = Memory_ReadLong(LoginServerStartPointer, tibiaclient, False)

  incUnit = LoginServerStep
  incTotal = PortOffset + ((nm - 1) * incUnit)
  adrPort = adrStruct + incTotal

  
  #If FinalMode = 0 Then
  b1 = Memory_ReadByte(adrPort, tibiaclient, True)
  b2 = Memory_ReadByte(adrPort + 1, tibiaclient, True)
  tport = GetTheLong(b1, b2)
  CRC = CStr(tport)
  Debug.Print CRC
  #End If
 
  OBTAINmemLoginPort = adrPort
End Function

Public Sub ModifyTibiaIPs2()
  ' Change login ips for Tibia 10.11 and higher
  Dim tibiaclient As Long
  Dim MemHERELoginServer As Long
  Dim MemHEREPortLoginServer As Long
  Dim i As Integer
  Dim j As Integer
  Dim num As Integer
  Dim num2 As Integer
  Dim posM As Integer
  Dim writeStr As String
  Dim writeChr As String
  Dim strIP As String
  Dim strCurrentIP As String
  Dim strCurrentPORT As String
  Dim lngFinalPort As Long
  Dim adrServer As Long
  Dim portb1 As Byte
  Dim portb2 As Byte
  Dim debugStr As String
  Dim MemHEREipNumber As Long

  
  If ReadyToChangeTibiaIP() = False Then
    Debug.Print "Warning: Blackd Proxy was not ready to modify Tibia Ips yet"
    Exit Sub
  End If
  'hWndDesktop = GetDesktopWindow()
  ResetProcessidIPrelations
  tibiaclient = 0

  Do

    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      Exit Do
    Else
      If NumberOfLoginServers = 0 Then
           NumberOfLoginServers = 10
      End If
      'debugSTR = readProxyChecker(tibiaclient)
      nullifyProxyChecker (tibiaclient)
      
      strIP = "127.0.0.1"

      For i = 1 To NumberOfLoginServers
        posM = 0
 

        MemHERELoginServer = OBTAINmemLoginServer(tibiaclient, i)
        'MemHEREPortLoginServer = MemPortLoginServer(i)
        MemHEREipNumber = OBTAINmemIPNumber(tibiaclient, i)
        
        
        ' write this string
        writeStr = strIP
        For j = 1 To Len(writeStr)
          writeChr = Left(writeStr, 1)
          writeStr = Right(writeStr, Len(writeStr) - 1)
          Memory_WriteByte MemHERELoginServer + posM, ConvStrToByte(writeChr), tibiaclient, True
          posM = posM + 1
        Next j
        AddProcessIdIPrelation strIP, tibiaclient
        ' 2x bytes 00
        Memory_WriteByte MemHERELoginServer + posM, &H0, tibiaclient, True
        posM = posM + 1
        Memory_WriteByte MemHERELoginServer + posM, &H0, tibiaclient, True
        posM = posM + 1
        
        ' write numeric ip 127.0.0.1
        Memory_WriteByte MemHEREipNumber + 0, &H7F, tibiaclient, True ' 127.
        Memory_WriteByte MemHEREipNumber + 1, &H0, tibiaclient, True  ' 0.
        Memory_WriteByte MemHEREipNumber + 2, &H0, tibiaclient, True  ' 0.
        Memory_WriteByte MemHEREipNumber + 3, &H1, tibiaclient, True  ' 1
        
        ' port
        If TibiaVersionLong >= 841 Then
            lngFinalPort = frmMain.SckClient(0).LocalPort
        Else
            lngFinalPort = CLng(frmMain.txtClientLoginP.Text)
        End If
        
        MemHEREPortLoginServer = OBTAINmemLoginPort(tibiaclient, i)
        portb1 = LowByteOfLong(lngFinalPort)
        portb2 = HighByteOfLong(lngFinalPort)
        Memory_WriteByte MemHEREPortLoginServer, portb1, tibiaclient, True
        Memory_WriteByte MemHEREPortLoginServer + 1, portb2, tibiaclient, True
      Next i
    End If
  Loop
End Sub

Public Sub ModifyTibiaIPs()
  Dim tibiaclient As Long
  'Dim hWndDesktop As Long
  Dim MemHERELoginServer As Long
  Dim MemHEREPortLoginServer As Long
  Dim i As Integer
  Dim j As Integer
  Dim num As Integer
  Dim num2 As Integer
  Dim posM As Integer
  Dim writeStr As String
  Dim writeChr As String
  Dim strIP As String
  Dim strCurrentIP As String
  Dim strCurrentPORT As String
  Dim lngFinalPort As Long
  
  If TibiaVersionLong >= 1011 Then
   ' new login system
   ModifyTibiaIPs2
  End If
  
  If ReadyToChangeTibiaIP() = False Then
    Debug.Print "Warning: Blackd Proxy was not ready to modify Tibia Ips yet"
    Exit Sub
  End If
  'hWndDesktop = GetDesktopWindow()
  ResetProcessidIPrelations
  tibiaclient = 0

  Do

    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      Exit Do
    Else
      'change IP and port in memory
      'SetAllPrivilegesForPID tibiaclient
      
      ' punto b - OK
      strCurrentIP = readMemoryString(tibiaclient, memLoginServer(1), 255, False)
      'Debug.Print "Original IP=" & strCurrentIP
      strCurrentPORT = CStr(Memory_ReadLong(MemPortLoginServer(1), tibiaclient))
     ' Debug.Print "Original Port=" & strCurrentPORT
      
      If NumberOfLoginServers = 0 Then
           NumberOfLoginServers = 5
      End If
      If AlternativeBinding = 0 Then
        strIP = "127.0.0.1"
      Else
        strCurrentIP = readMemoryString(tibiaclient, memLoginServer(1))
        If Left$(strCurrentIP, 4) = "127." Then
          strIP = strCurrentIP
        Else
          strIP = getFirstAvailableIP()
        End If
      End If
      
      For i = 1 To NumberOfLoginServers
        posM = 0
 

        MemHERELoginServer = memLoginServer(i)
        MemHEREPortLoginServer = MemPortLoginServer(i)
        
        
        
        ' write this string
        writeStr = strIP
        For j = 1 To Len(writeStr)
          writeChr = Left(writeStr, 1)
          writeStr = Right(writeStr, Len(writeStr) - 1)
          Memory_WriteByte MemHERELoginServer + posM, ConvStrToByte(writeChr), tibiaclient
          posM = posM + 1
        Next j
        AddProcessIdIPrelation strIP, tibiaclient
        ' 2x bytes 00
        Memory_WriteByte MemHERELoginServer + posM, &H0, tibiaclient
        posM = posM + 1
        Memory_WriteByte MemHERELoginServer + posM, &H0, tibiaclient
        posM = posM + 1
        ' port
        If TibiaVersionLong >= 841 Then
            lngFinalPort = frmMain.SckClient(0).LocalPort
        Else
            lngFinalPort = CLng(frmMain.txtClientLoginP.Text)
        End If
        Memory_WriteLong MemHEREPortLoginServer, lngFinalPort, tibiaclient
      Next i
    End If
  Loop
End Sub




Public Function buildIPstring(int1 As Integer, int2 As Integer, int3 As Integer, int4 As Integer) As String
  buildIPstring = CStr(int1) & "." & CStr(int2) & "." & CStr(int3) & "." & CStr(int4)
End Function
Public Function GetTheLong(byte1 As Byte, byte2 As Byte) As Long
  'get the long from 2 consecutive bytes in a tibia packet
  Dim res As Long
  res = CLng(byte2) * 256 + CLng(byte1)
  GetTheLong = res
End Function


Public Function longToBytes(ByRef byteArray() As Byte, ByVal thelong As Long) As Byte()
    CopyMemory byteArray(0), ByVal VarPtr(thelong), Len(thelong)
End Function

Public Function FourBytesLong(byte1 As Byte, byte2 As Byte, byte3 As Byte, byte4 As Byte) As Long
  Dim res As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  If byte4 = &HFF Then
    ' should not happen
    Debug.Print "WARNING: bad call to FourBytesLong"
    res = -1
  Else
    res = CLng(byte4) * 16777216 + CLng(byte3) * 65536 + CLng(byte2) * 256 + CLng(byte1)
  End If
  FourBytesLong = res
  Exit Function
goterr:
  FourBytesLong = -1
End Function

Public Function FourBytesDouble(byte1 As Byte, byte2 As Byte, byte3 As Byte, byte4 As Byte) As Double
  Dim res As Double
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  res = CDbl(byte4) * 16777216 + CDbl(byte3) * 65536 + CDbl(byte2) * 256 + CDbl(byte1)
  FourBytesDouble = res
  Exit Function
goterr:
  FourBytesDouble = -1
End Function

Public Function RewriteWithLocalIPcharPos(ByVal pos As Long) As Byte
  
  Dim res As Byte
  Dim letr As String
  If pos > Len(localstr) Then
     res = &H0
  Else
    letr = Mid$(localstr, pos, 1)
    res = AscB(letr)
  End If
  
 
  RewriteWithLocalIPcharPos = res
End Function

Public Function NewCharListChanger(ByRef packet() As Byte, ByRef packetNew() As Byte, ByVal idConnection As Integer, ByVal strIP As String, bstart As Long, ByRef pos As Long) As Integer
  Const rewriteEnabled As Boolean = True ' set false for debug purposes. Packet should only be copied in newpacket. Proxy will not really connect.
  On Error GoTo goterr:
  Dim ti As Long
  Dim newServerName As String
  Dim ultimoN As Long
  Dim numChars As Long
  Dim i As Long
  Dim j As Long
  Dim lonSName As Long
  Dim serID As Long
  Dim lonCName As Long
  Dim totalServidores As Long
  Dim b1Local As Byte
  Dim b2Local As Byte
  Dim lLocal As Long
  Dim charName As String
  Dim newServerDomain As String
  Dim tmpB As Byte
  Dim newServerPort As Long
  Dim gamep As Long
  Dim hb As Byte
  Dim lb As Byte
  Dim lon As Long
  Dim adder As Long
  Dim servDOMAIN As String
  Dim tmpU As Long
  Dim newSize As Long
  Dim modSize As Long
  Dim aleat As Long
  Dim servName As String
  Dim servPort As Long
  Dim servIP1 As Byte
  Dim servIP2 As Byte
  Dim servIP3 As Byte
  Dim servIP4 As Byte
  Dim finalAfterPos As Long
  If (Not (packet(pos) = &H64)) Then
    Debug.Print "NewCharListChanger FAIL: Expected &H64 at packet(pos)"
    NewCharListChanger = -1
    Exit Function
  End If
  
  lLocal = Len(localstr)
  b1Local = HighByteOfLong(lLocal)
  b2Local = LowByteOfLong(lLocal)
  gamep = CLng(frmMain.sckClientGame(0).LocalPort)
  servDOMAIN = ""
  adder = bstart - 2
  lon = GetTheLong(packet(0 + adder), packet(1 + adder))
  finalAfterPos = adder + GetTheLong(packet(0 + adder), packet(1 + adder))
     
  pos = pos + 2
  ReDim Preserve packetNew(pos)
  For ti = 1 To pos
    packetNew(ti) = packet(ti)
  Next ti
  pos = pos - 1
      
  '  Debug.Print frmMain.showAsStr(packet, True)
  ' Debug.Print frmMain.showAsStr(packetNEW, True)
  ultimoN = pos
  totalServidores = CLng(packet(pos))
  ReDim loadedServers(totalServidores - 1)
  ReDim loadedPorts(totalServidores - 1)
  ReDim loadedDomains(totalServidores - 1)
  For i = 1 To totalServidores
    pos = pos + 1
    ReDim Preserve packetNew(ultimoN + 1)
    ultimoN = ultimoN + 1
    packetNew(ultimoN) = packet(pos)
    serID = CLng(packet(pos))
    pos = pos + 1
    lonCName = GetTheLong(packet(pos), packet(pos + 1))
        
    ReDim Preserve packetNew(ultimoN + 1)
    ultimoN = ultimoN + 1
    packetNew(ultimoN) = packet(pos)
        
    ReDim Preserve packetNew(ultimoN + 1)
    ultimoN = ultimoN + 1
    packetNew(ultimoN) = packet(pos + 1)
    pos = pos + 2
        
        
       
    'Debug.Print frmMain.showAsStr(packetNEW, True)
        
    newServerName = ""
    For j = 1 To lonCName
      newServerName = newServerName & Chr(packet(pos))
          
      ReDim Preserve packetNew(ultimoN + 1)
      ultimoN = ultimoN + 1
      packetNew(ultimoN) = packet(pos)
          
      pos = pos + 1
    Next j
    lonSName = GetTheLong(packet(pos), packet(pos + 1))
        
        
    If (rewriteEnabled = True) Then
      ReDim Preserve packetNew(ultimoN + 1)
      ultimoN = ultimoN + 1
      packetNew(ultimoN) = b2Local
            
      ReDim Preserve packetNew(ultimoN + 1)
      ultimoN = ultimoN + 1
      packetNew(ultimoN) = b1Local
      pos = pos + 2
    Else
      ReDim Preserve packetNew(ultimoN + 1)
      ultimoN = ultimoN + 1
      packetNew(ultimoN) = packet(pos)
            
      ReDim Preserve packetNew(ultimoN + 1)
      ultimoN = ultimoN + 1
      packetNew(ultimoN) = packet(pos + 1)
      pos = pos + 2
    End If
      'read server name of character i
      newServerDomain = ""
      For j = 1 To lonSName
          
          newServerDomain = newServerDomain & Chr(packet(pos))
          If (rewriteEnabled = True) Then
          
              tmpB = RewriteWithLocalIPcharPos(j)
              If tmpB = &H0 Then
                 'packet(pos) = tmpB ' commented UNSURE...
                 
                 
    '            ReDim Preserve packetNEW(ultimoN + 1) ' borrar luego
    '            ultimoN = ultimoN + 1
    '            packetNEW(ultimoN) = packet(pos)
                
                
              Else
              
                'packet(pos) = tmpB
                'ReDim Preserve packetNEW(ultimoN + 1)
                'ultimoN = ultimoN + 1
                'packetNEW(ultimoN) = packet(pos)
                
               ' packet(pos) = tmpB
                ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = tmpB
                
              End If
          Else
                ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = packet(pos)
          End If
          pos = pos + 1
        Next j
        newServerPort = GetTheLong(packet(pos), packet(pos + 1))
        loadedServers(serID) = newServerName
        loadedPorts(serID) = newServerPort
        loadedDomains(serID) = newServerDomain
        
        If (rewriteEnabled = True) Then
              hb = HighByteOfLong(gamep)
              lb = LowByteOfLong(gamep)
        
              ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = lb
    
              ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = hb
        Else
              ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = packet(pos)
    
              ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = packet(pos + 1)
        End If
        If (Left$(newServerDomain, 4) = "127.") Then
            NewCharListChanger = 0
            Exit Function
        End If
       ' Debug.Print newServerName & " = " & newServerDomain & ":" & newServerPort
        
        AddGameServer newServerName, "127.0.0.1:" & newServerPort, newServerDomain
        pos = pos + 2
        
        ReDim Preserve packetNew(ultimoN + 1)
        ultimoN = ultimoN + 1
        packetNew(ultimoN) = packet(pos)
      Next i
      pos = pos + 1
      
      ' We don't need to care about the rest, just copy all and we have the packet ready
      ' Note1 : this includes last part of subpacket &H0C (char -> server index in internal list)
    
        tmpU = lon + 8 - pos
  ReDim Preserve packetNew(ultimoN + tmpU)
  For ti = 1 To tmpU
    packetNew(ultimoN + ti) = packet(pos + ti - 1)
  Next ti
  ultimoN = ultimoN + tmpU
  
  
  
  
  ' Fill packet with correct trash
  newSize = UBound(packetNew) - 7

  hb = HighByteOfLong(newSize)
  lb = LowByteOfLong(newSize)
  packetNew(6) = lb
  packetNew(7) = hb
  

 
  
  newSize = newSize + 6
  hb = HighByteOfLong(newSize)
  lb = LowByteOfLong(newSize)
  packetNew(0) = lb
  packetNew(1) = hb


  modSize = (newSize + 4) Mod 8
  aleat = 0
  If modSize > 0 Then
    aleat = (8 - modSize)
  End If


  For ti = 1 To aleat
      ReDim Preserve packetNew(ultimoN + 1)
      ultimoN = ultimoN + 1
      If rewriteEnabled = True Then
        packetNew(ultimoN) = &H0
      Else
        packetNew(ultimoN) = packet(ultimoN)
      End If
  Next ti
  
  newSize = UBound(packetNew) - 1
  hb = HighByteOfLong(newSize)
  lb = LowByteOfLong(newSize)
  packetNew(0) = lb
  packetNew(1) = hb
  
  
  
  numChars = CLng(packet(pos))
  pos = pos + 1
  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Client can select this characters:"
  ResetCharList2 idConnection
  For i = 1 To numChars 'read all the characters on the list
    serID = CLng(packet(pos))
    pos = pos + 1
    lonCName = GetTheLong(packet(pos), packet(pos + 1))
    pos = pos + 2
    charName = ""
    For j = 1 To lonCName
      charName = charName & Chr(packet(pos))
      pos = pos + 1
    Next j
    servName = loadedServers(serID)
    servPort = loadedPorts(serID)
    servDOMAIN = loadedDomains(serID)


   ' servPort = GetGameServerPort(servName)
   ' servDOMAIN = GetGameServerDOMAIN(servName)
   
    AddCharServer2 idConnection, charName, servName, servIP1, servIP2, servIP3, servIP4, servPort, servDOMAIN
   ' Debug.Print charName & "-> server #" & CStr(serID) & " (" & servName & ") = " & servDOMAIN & ":" & servPort
  Next i
  
  
   pos = finalAfterPos

  '  Debug.Print frmMain.showAsStr(packetNEW, True)
  '  Debug.Print "OK"
     
  NewCharListChanger = 1
  Exit Function
goterr:
  Debug.Print ("FATAL ERROR AT NewCharListChanger!!!")
  NewCharListChanger = -1
End Function

Public Function PacketIPchange6(ByRef packet() As Byte, ByVal idConnection As Integer, ByVal strIP As String, Optional bstart As Long = 2) As Integer
  Const rewriteEnabled As Boolean = True ' set false for debug purposes. Packet should only be copied in newpacket. Proxy will not really connect.
  Dim lon As Long
  Dim motd As Long
  Dim numChars As Long
  Dim lonCName As Long
  Dim lonSName As Long
  Dim i As Integer
  Dim j As Integer
  Dim pos As Long
  Dim servName As String
  Dim servIP1 As Byte
  Dim servIP2 As Byte
  Dim servIP3 As Byte
  Dim servIP4 As Byte
  Dim servPort As Long
  Dim charName As String
  Dim hb As Byte
  Dim lb As Byte
  Dim res As Integer
  Dim adder As Long
  Dim serverIPport As String
  Dim lngIPid As Long
  Dim gamep As Long
  Dim doingdebugHere As Boolean
  Dim tipoBloque As Byte
  Dim servDOMAIN As String
  Dim totalServidores As Long
  Dim newServerName As String
  Dim newServerDomain As String
  Dim newServerPort As Long
  Dim loadedServers() As String
  Dim loadedPorts() As String
  Dim loadedDomains() As String
  Dim zz As Long
  Dim serID As Long
  Dim packetNew() As Byte
  Dim ultimoN As Long
  Dim ti As Long
  Dim lLocal As Long
  Dim b1Local As Byte
  Dim b2Local As Byte
  Dim tmpB As Byte
  Dim tmpU As Long
  Dim newSize As Long
  Dim modSize As Long
  Dim aleat As Long
  Dim realL As Long
  Dim ultimoTI As Long
  Dim strangeNewThingLen As Long
  Dim debugChain As String
  Dim pType As Byte
  Dim initialPos As Long
  Dim debugLon1 As Long
  Dim showDebug As Boolean
  Dim debugReasons As String
  Dim lastGoodPos As Long
  Dim mobName As String
  Dim expectMore As Boolean
  Dim finalAfterPos As Long
  Dim lastpType As Byte
  Dim lontmp1 As Long
  Dim charlistWasParsed As Boolean
  Dim theInc As Long
  Dim fillstart As Long
  Dim nRes As Integer
  Dim fillend As Long
  Dim packetsdif As Long
  On Error GoTo returnTheResult
  charlistWasParsed = False
    'Debug.Print "PacketIPchange6 ORIGINAL>" & frmMain.showAsStr(packet, True)
  debugChain = ""
  lLocal = Len(localstr)
  b1Local = HighByteOfLong(lLocal)
  b2Local = LowByteOfLong(lLocal)
  ultimoN = 0
  gamep = CLng(frmMain.sckClientGame(0).LocalPort)
  servDOMAIN = ""
  'LogOnFile "gotthem.txt", frmMain.showAsStr2(packet, 0) & vbCrLf
  res = -1 'error
  adder = bstart - 2
 ' If packet(2 + adder) <> &H28 Then
  '  Debug.Print "This is not a list of character Tibia 10.91+ packet... Received type = " & GoodHex(packet(2 + adder))
  '  res = -1
   ' GoTo returnTheResult 'this is not a list of character packet
  'End If
  
  If frmMain.chckAlter.value = 0 Then
    res = 0
    GoTo returnTheResult 'proxy user don't want to change this packet
  End If
  lon = GetTheLong(packet(0 + adder), packet(1 + adder))
    
  ultimoN = 2 + adder - 1
  ReDim packetNew(ultimoN)
  For ti = 0 To ultimoN
    packetNew(ti) = packet(ti)
  Next ti
  initialPos = ultimoN + 1
  pos = initialPos
  debugReasons = ""
  lastpType = &HFF
  finalAfterPos = adder + GetTheLong(packet(0 + adder), packet(1 + adder))
  packetsdif = 0
  Do
    initialPos = pos
    lastGoodPos = pos
    mobName = ""
    expectMore = True
    pType = packet(pos)
    debugChain = debugChain & " " & GoodHex(pType)
    'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "! EVAL : " & GoodHex(pType)
    Select Case pType ' type of subpacket
    Case &H28
     ' Debug.Print "&H28 - AUTH"
      pos = pos + 1
      lontmp1 = GetTheLong(packet(pos), packet(pos + 1))
      pos = pos + 2 + lontmp1
      
      theInc = 3 + lontmp1
      fillstart = UBound(packetNew) + 1
      fillend = fillstart + theInc - 1
      ReDim Preserve packetNew(fillend)
      ultimoN = ultimoN + theInc
      packetsdif = ultimoN - pos + 1
      For ti = fillstart To fillend
        packetNew(ti) = packet(ti - packetsdif)
      Next ti
       
       
      'Debug.Print frmMain.showAsStr(packet, True)
      'Debug.Print frmMain.showAsStr(packetNEW, True)
      'Debug.Print "OK"
    Case &H14
     ' Debug.Print "&H14 - MOTD"
      pos = pos + 1
      lontmp1 = GetTheLong(packet(pos), packet(pos + 1))
      pos = pos + 2 + lontmp1
      
      theInc = 3 + lontmp1
      fillstart = UBound(packetNew) + 1
      fillend = fillstart + theInc - 1
      ReDim Preserve packetNew(fillend)
      ultimoN = ultimoN + theInc
      packetsdif = ultimoN - pos + 1
      For ti = fillstart To fillend
        packetNew(ti) = packet(ti - packetsdif)
      Next ti
      
      
      'Debug.Print frmMain.showAsStr(packet, True)
      'Debug.Print frmMain.showAsStr(packetNEW, True)
      'Debug.Print "OK"
    Case &HC
      '  Debug.Print "&H0C - 0C ??"
      pos = pos + 2
    Case &H64
     ' Debug.Print "&H64 - CHAR LIST"
      nRes = NewCharListChanger(packet, packetNew, idConnection, strIP, bstart, pos)
      If (nRes = 1) Then
        res = 1
        frmMain.UnifiedSendToClient idConnection, packetNew, False, True
      Else
        res = 0
        frmMain.UnifiedSendToClient idConnection, packet, False, True
      End If
      LastCharServerIndex = idConnection
      PacketIPchange6 = res
      Exit Function
    Case Else
      ' should not happen, unless protocol get updated
       ' LogOnFile "errors.txt", "WARNING IN PACKET" & frmMain.showAsStr2(packet, 0) & vbCrLf & "UNKNOWN PTYPE : " & GoodHex(pType)
       Debug.Print ("UNKNOWN PTYPE : " & GoodHex(pType) & " ; PACKET= " & frmMain.showAsStr2(packet, 0))
       debugReasons = debugReasons & vbCrLf & " [ UNKNOWN PTYPE : " & GoodHex(pType) & " ] "
       showDebug = True
       expectMore = False
    End Select
    If pos = finalAfterPos Then
      expectMore = False
    ElseIf pos > finalAfterPos Then
      debugLon1 = pos - finalAfterPos
      debugReasons = debugReasons & vbCrLf & " [ BAD EVAL IN PTYPE : " & GoodHex(pType) & " ; overread of +" & CStr(debugLon1) & _
      " bytes. Last good position=" & CStr(lastGoodPos) & " ]"
      showDebug = True
      expectMore = False
    End If
    lastpType = pType
  Loop Until expectMore = False
  
  

  'Debug.Print "OLD PCKT>" & frmMain.showAsStr(packet, True)
  'Debug.Print "NEW PCKT>" & frmMain.showAsStr(packetNEW, True)
  If (charlistWasParsed) Then
    res = 1
    frmMain.UnifiedSendToClient idConnection, packetNew, False, True
  Else
    res = 0
    frmMain.UnifiedSendToClient idConnection, packet, False, True
  End If
  LastCharServerIndex = idConnection
returnTheResult:
  'LogOnFile "gotthem.txt", "AFTER (" & CStr(res) & ")> " & frmMain.showAsStr2(packet, 0) & vbCrLf
  PacketIPchange6 = res
End Function


Public Function PacketIPchange5(ByRef packet() As Byte, ByVal idConnection As Integer, ByVal strIP As String, Optional bstart As Long = 2) As Integer
  Const rewriteEnabled As Boolean = True ' set false for debug purposes. Packet should only be copied in newpacket. Proxy will not really connect.
  Dim lon As Long
  Dim motd As Long
  Dim numChars As Long
  Dim lonCName As Long
  Dim lonSName As Long
  Dim i As Integer
  Dim j As Integer
  Dim pos As Long
  Dim servName As String
  Dim servIP1 As Byte
  Dim servIP2 As Byte
  Dim servIP3 As Byte
  Dim servIP4 As Byte
  Dim servPort As Long
  Dim charName As String
  Dim hb As Byte
  Dim lb As Byte
  Dim res As Integer
  Dim adder As Long
  Dim serverIPport As String
  Dim lngIPid As Long
  Dim gamep As Long
  Dim doingdebugHere As Boolean
  Dim tipoBloque As Byte
  Dim servDOMAIN As String
  Dim totalServidores As Long
  Dim newServerName As String
  Dim newServerDomain As String
  Dim newServerPort As Long
  Dim loadedServers() As String
  Dim loadedPorts() As String
  Dim loadedDomains() As String
  Dim zz As Long
  Dim serID As Long
  Dim packetNew() As Byte
  Dim ultimoN As Long
  Dim ti As Long
  Dim lLocal As Long
  Dim b1Local As Byte
  Dim b2Local As Byte
  Dim tmpB As Byte
  Dim tmpU As Long
  Dim newSize As Long
  Dim modSize As Long
  Dim aleat As Long
  Dim realL As Long
  Dim ultimoTI As Long
  Dim strangeNewThingLen As Long
  On Error GoTo returnTheResult
  
   ' Debug.Print "PacketIPchange5 ORIGINAL>" & frmMain.showAsStr(packet, True)
    
  lLocal = Len(localstr)
  b1Local = HighByteOfLong(lLocal)
  b2Local = LowByteOfLong(lLocal)
  ultimoN = 0
  gamep = CLng(frmMain.sckClientGame(0).LocalPort)
  servDOMAIN = ""
  'LogOnFile "gotthem.txt", frmMain.showAsStr2(packet, 0) & vbCrLf
  res = -1 'error
  adder = bstart - 2
  If packet(2 + adder) <> &H28 Then
    Debug.Print "This is not a list of character Tibia 10.74+ packet... Received type = " & GoodHex(packet(2 + adder))
    res = -1
    GoTo returnTheResult 'this is not a list of character packet
  End If
  strangeNewThingLen = GetTheLong(packet(adder + 3), packet(adder + 4))


  If frmMain.chckAlter.value = 0 Then
    res = 0
    GoTo returnTheResult 'proxy user don't want to change this packet
  End If

    If packet(adder + 5 + strangeNewThingLen) <> &H14 Then
    Debug.Print "This is not a list of character packet... Received type at " & CStr(adder + 5 + strangeNewThingLen) & " = " & GoodHex(packet(adder + 5 + strangeNewThingLen))
    res = -1
    GoTo returnTheResult 'this is not a list of character packet
    End If

 
  lon = GetTheLong(packet(0 + adder), packet(1 + adder))
'motd = GetTheLong(packet(3 + adder), packet(4 + adder))
 motd = GetTheLong(packet(3 + adder + 3 + strangeNewThingLen), packet(4 + adder + 3 + strangeNewThingLen))
 
 ' pos = motd + 5 + adder + 2 ' 2 new bytes after motd since tibia 10.61 :
pos = motd + 5 + adder + 2 + 3 + strangeNewThingLen
  
  tipoBloque = packet(pos)
  pos = pos + 1

  
  ultimoN = pos
    ReDim packetNew(ultimoN)
  For ti = 0 To ultimoN
    packetNew(ti) = packet(ti)
    ultimoTI = ti
  Next ti


    ti = ultimoTI


 
  If tipoBloque = &H64 Then
    totalServidores = CLng(packet(pos))
    ReDim loadedServers(totalServidores - 1)
    ReDim loadedPorts(totalServidores - 1)
    ReDim loadedDomains(totalServidores - 1)
    For i = 1 To totalServidores
    
        pos = pos + 1
        ReDim Preserve packetNew(ultimoN + 1)
        ultimoN = ultimoN + 1
        packetNew(ultimoN) = packet(pos)

        
        serID = CLng(packet(pos))
        pos = pos + 1
        lonCName = GetTheLong(packet(pos), packet(pos + 1))
        
        ReDim Preserve packetNew(ultimoN + 1)
        ultimoN = ultimoN + 1
        packetNew(ultimoN) = packet(pos)
        
        ReDim Preserve packetNew(ultimoN + 1)
        ultimoN = ultimoN + 1
        packetNew(ultimoN) = packet(pos + 1)
        pos = pos + 2
        
        newServerName = ""
        For j = 1 To lonCName
          newServerName = newServerName & Chr(packet(pos))
          
          ReDim Preserve packetNew(ultimoN + 1)
          ultimoN = ultimoN + 1
          packetNew(ultimoN) = packet(pos)
          
          pos = pos + 1
        Next j
        lonSName = GetTheLong(packet(pos), packet(pos + 1))
        
        
         If (rewriteEnabled = True) Then
            ReDim Preserve packetNew(ultimoN + 1)
            ultimoN = ultimoN + 1
            packetNew(ultimoN) = b2Local
            
            ReDim Preserve packetNew(ultimoN + 1)
            ultimoN = ultimoN + 1
            packetNew(ultimoN) = b1Local
            pos = pos + 2
        Else
            ReDim Preserve packetNew(ultimoN + 1)
            ultimoN = ultimoN + 1
            packetNew(ultimoN) = packet(pos)
            
            ReDim Preserve packetNew(ultimoN + 1)
            ultimoN = ultimoN + 1
            packetNew(ultimoN) = packet(pos + 1)
            pos = pos + 2
        End If
        'read server name of character i
        newServerDomain = ""
        For j = 1 To lonSName
          
          newServerDomain = newServerDomain & Chr(packet(pos))
          If (rewriteEnabled = True) Then
          
              tmpB = RewriteWithLocalIPcharPos(j)
              If tmpB = &H0 Then
                 'packet(pos) = tmpB ' commented UNSURE...
                 
                 
    '            ReDim Preserve packetNEW(ultimoN + 1) ' borrar luego
    '            ultimoN = ultimoN + 1
    '            packetNEW(ultimoN) = packet(pos)
                
                
              Else
              
                'packet(pos) = tmpB
                'ReDim Preserve packetNEW(ultimoN + 1)
                'ultimoN = ultimoN + 1
                'packetNEW(ultimoN) = packet(pos)
                
               ' packet(pos) = tmpB
                ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = tmpB
                
              End If
          Else
                ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = packet(pos)
          End If
          pos = pos + 1
        Next j
        newServerPort = GetTheLong(packet(pos), packet(pos + 1))
        loadedServers(serID) = newServerName
        loadedPorts(serID) = newServerPort
        loadedDomains(serID) = newServerDomain
        
        If (rewriteEnabled = True) Then
              hb = HighByteOfLong(gamep)
              lb = LowByteOfLong(gamep)
        
              ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = lb
    
              ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = hb
        Else
              ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = packet(pos)
    
              ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = packet(pos + 1)
        End If
        
        AddGameServer newServerName, "127.0.0.1:" & newServerPort, newServerDomain
        pos = pos + 2
        
        ReDim Preserve packetNew(ultimoN + 1)
        ultimoN = ultimoN + 1
        packetNew(ultimoN) = packet(pos)
    Next i
    
     
    pos = pos + 1
  End If
  
 
  tmpU = lon + 8 - pos
  ReDim Preserve packetNew(ultimoN + tmpU)
  For ti = 1 To tmpU
    packetNew(ultimoN + ti) = packet(pos + ti - 1)
  Next ti

  ultimoN = ultimoN + tmpU
  newSize = UBound(packetNew) - 7

  hb = HighByteOfLong(newSize)
  lb = LowByteOfLong(newSize)
  packetNew(6) = lb
  packetNew(7) = hb
  


  newSize = newSize + 6
  hb = HighByteOfLong(newSize)
  lb = LowByteOfLong(newSize)
  packetNew(0) = lb
  packetNew(1) = hb

'
  modSize = (newSize + 4) Mod 8
  aleat = 0
  If modSize > 0 Then
    aleat = (8 - modSize)
  End If


  For ti = 1 To aleat
      ReDim Preserve packetNew(ultimoN + 1)
      ultimoN = ultimoN + 1
      If rewriteEnabled = True Then
        packetNew(ultimoN) = &H0
      Else
        packetNew(ultimoN) = packet(ultimoN)
      End If
  Next ti
  
  newSize = UBound(packetNew) - 1
  hb = HighByteOfLong(newSize)
  lb = LowByteOfLong(newSize)
  packetNew(0) = lb
  packetNew(1) = hb
  



 ' Debug.Print "NEW SIZE=" & GoodHex(lb) & " " & GoodHex(hb)
  
  numChars = CLng(packet(pos))
  pos = pos + 1
  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Client can select this characters:"
  ResetCharList2 idConnection
  For i = 1 To numChars 'read all the characters on the list
    serID = CLng(packet(pos))
    pos = pos + 1
    lonCName = GetTheLong(packet(pos), packet(pos + 1))
    pos = pos + 2
    charName = ""
    For j = 1 To lonCName
      charName = charName & Chr(packet(pos))
      pos = pos + 1
    Next j
    servName = loadedServers(serID)
    servPort = loadedPorts(serID)
    servDOMAIN = loadedDomains(serID)


   ' servPort = GetGameServerPort(servName)
   ' servDOMAIN = GetGameServerDOMAIN(servName)
   
    AddCharServer2 idConnection, charName, servName, servIP1, servIP2, servIP3, servIP4, servPort, servDOMAIN
    'Debug.Print charName & "-> server #" & CStr(serID) & " (" & servName & ") = " & servDOMAIN & ":" & servPort
  Next i
  
  'Debug.Print frmMain.showAsStr(packet, True)
  

 ' Debug.Print "NEW PCKT>" & frmMain.showAsStr(packetNEW, True)
 
  frmMain.UnifiedSendToClient idConnection, packetNew, False, True

  
  LastCharServerIndex = idConnection
  res = 1
returnTheResult:
  'LogOnFile "gotthem.txt", "AFTER (" & CStr(res) & ")> " & frmMain.showAsStr2(packet, 0) & vbCrLf
  PacketIPchange5 = res
End Function


Public Function PacketIPchange5b(ByRef packet() As Byte, ByVal idConnection As Integer, ByVal strIP As String, Optional bstart As Long = 2) As Integer
  Const rewriteEnabled As Boolean = True ' set false for debug purposes. Packet should only be copied in newpacket. Proxy will not really connect.
  Dim lon As Long
  Dim motd As Long
  Dim numChars As Long
  Dim lonCName As Long
  Dim lonSName As Long
  Dim i As Integer
  Dim j As Integer
  Dim pos As Long
  Dim servName As String
  Dim servIP1 As Byte
  Dim servIP2 As Byte
  Dim servIP3 As Byte
  Dim servIP4 As Byte
  Dim servPort As Long
  Dim charName As String
  Dim hb As Byte
  Dim lb As Byte
  Dim res As Integer
  Dim adder As Long
  Dim serverIPport As String
  Dim lngIPid As Long
  Dim gamep As Long
  Dim doingdebugHere As Boolean
  Dim tipoBloque As Byte
  Dim servDOMAIN As String
  Dim totalServidores As Long
  Dim newServerName As String
  Dim newServerDomain As String
  Dim newServerPort As Long
  Dim loadedServers() As String
  Dim loadedPorts() As String
  Dim loadedDomains() As String
  Dim zz As Long
  Dim serID As Long
  Dim packetNew() As Byte
  Dim ultimoN As Long
  Dim ti As Long
  Dim lLocal As Long
  Dim b1Local As Byte
  Dim b2Local As Byte
  Dim tmpB As Byte
  Dim tmpU As Long
  Dim newSize As Long
  Dim modSize As Long
  Dim aleat As Long
  Dim realL As Long
  Dim ultimoTI As Long
  Dim strangeNewThingLen As Long
  Dim with28 As Boolean

  On Error GoTo returnTheResult
  
    'Debug.Print "PacketIPchange5b ORIGINAL>" & frmMain.showAsStr(packet, True)
    
  lLocal = Len(localstr)
  b1Local = HighByteOfLong(lLocal)
  b2Local = LowByteOfLong(lLocal)
  ultimoN = 0
  gamep = CLng(frmMain.sckClientGame(0).LocalPort)
  servDOMAIN = ""
  'LogOnFile "gotthem.txt", frmMain.showAsStr2(packet, 0) & vbCrLf
  res = -1 'error
  adder = bstart - 2


  If frmMain.chckAlter.value = 0 Then
    res = 0
    GoTo returnTheResult 'proxy user don't want to change this packet
  End If
  If packet(adder + 2) <> &H14 Then
    Debug.Print "This is not a list of character packet... " & CStr(adder + 2) & " Received type = " & GoodHex(packet(adder + 2))
    res = -1
    GoTo returnTheResult 'this is not a list of character packet
  End If

 
  lon = GetTheLong(packet(0 + adder), packet(1 + adder))
motd = GetTheLong(packet(3 + adder), packet(4 + adder))
 'motd = GetTheLong(packet(3 + adder + 3 + strangeNewThingLen), packet(4 + adder + 3 + strangeNewThingLen))
   with28 = True
  If packet(adder + 5 + motd) <> &H28 Then
    with28 = False
      If packet(adder + 5 + motd) <> &H64 Then
        Debug.Print "This is not a list of character packet... " & CStr(adder + 5 + motd) & " Received type = " & GoodHex(packet(adder + 5 + motd))
        res = -1
        GoTo returnTheResult 'this is not a list of character packet
      End If
  End If
  If with28 = True Then
    strangeNewThingLen = GetTheLong(packet(adder + 5 + motd + 1), packet(adder + 5 + motd + 2))
    pos = adder + 5 + motd + 3 + strangeNewThingLen
  Else
    pos = adder + 5 + motd
  End If
  tipoBloque = packet(pos)
  pos = pos + 1

  
  ultimoN = pos
    ReDim packetNew(ultimoN)
  For ti = 0 To ultimoN
    packetNew(ti) = packet(ti)
    ultimoTI = ti
  Next ti


    ti = ultimoTI

 
  If tipoBloque = &H64 Then
    totalServidores = CLng(packet(pos))
    ReDim loadedServers(totalServidores - 1)
    ReDim loadedPorts(totalServidores - 1)
    ReDim loadedDomains(totalServidores - 1)
    For i = 1 To totalServidores
    
        pos = pos + 1
        ReDim Preserve packetNew(ultimoN + 1)
        ultimoN = ultimoN + 1
        packetNew(ultimoN) = packet(pos)

        
        serID = CLng(packet(pos))
        pos = pos + 1
        lonCName = GetTheLong(packet(pos), packet(pos + 1))
        
        ReDim Preserve packetNew(ultimoN + 1)
        ultimoN = ultimoN + 1
        packetNew(ultimoN) = packet(pos)
        
        ReDim Preserve packetNew(ultimoN + 1)
        ultimoN = ultimoN + 1
        packetNew(ultimoN) = packet(pos + 1)
        pos = pos + 2
        
        newServerName = ""
        For j = 1 To lonCName
          newServerName = newServerName & Chr(packet(pos))
          
          ReDim Preserve packetNew(ultimoN + 1)
          ultimoN = ultimoN + 1
          packetNew(ultimoN) = packet(pos)
          
          pos = pos + 1
        Next j
        lonSName = GetTheLong(packet(pos), packet(pos + 1))
        
        
         If (rewriteEnabled = True) Then
            ReDim Preserve packetNew(ultimoN + 1)
            ultimoN = ultimoN + 1
            packetNew(ultimoN) = b2Local
            
            ReDim Preserve packetNew(ultimoN + 1)
            ultimoN = ultimoN + 1
            packetNew(ultimoN) = b1Local
            pos = pos + 2
        Else
            ReDim Preserve packetNew(ultimoN + 1)
            ultimoN = ultimoN + 1
            packetNew(ultimoN) = packet(pos)
            
            ReDim Preserve packetNew(ultimoN + 1)
            ultimoN = ultimoN + 1
            packetNew(ultimoN) = packet(pos + 1)
            pos = pos + 2
        End If
        'read server name of character i
        newServerDomain = ""
        For j = 1 To lonSName
          
          newServerDomain = newServerDomain & Chr(packet(pos))
          If (rewriteEnabled = True) Then
          
              tmpB = RewriteWithLocalIPcharPos(j)
              If tmpB = &H0 Then
                 'packet(pos) = tmpB ' commented UNSURE...
                 
                 
    '            ReDim Preserve packetNEW(ultimoN + 1) ' borrar luego
    '            ultimoN = ultimoN + 1
    '            packetNEW(ultimoN) = packet(pos)
                
                
              Else
              
                'packet(pos) = tmpB
                'ReDim Preserve packetNEW(ultimoN + 1)
                'ultimoN = ultimoN + 1
                'packetNEW(ultimoN) = packet(pos)
                
               ' packet(pos) = tmpB
                ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = tmpB
                
              End If
          Else
                ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = packet(pos)
          End If
          pos = pos + 1
        Next j
        newServerPort = GetTheLong(packet(pos), packet(pos + 1))
        loadedServers(serID) = newServerName
        loadedPorts(serID) = newServerPort
        loadedDomains(serID) = newServerDomain
        
        If (rewriteEnabled = True) Then
              hb = HighByteOfLong(gamep)
              lb = LowByteOfLong(gamep)
        
              ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = lb
    
              ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = hb
        Else
              ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = packet(pos)
    
              ReDim Preserve packetNew(ultimoN + 1)
                ultimoN = ultimoN + 1
                packetNew(ultimoN) = packet(pos + 1)
        End If
        
        AddGameServer newServerName, "127.0.0.1:" & newServerPort, newServerDomain
        pos = pos + 2
        
        ReDim Preserve packetNew(ultimoN + 1)
        ultimoN = ultimoN + 1
        packetNew(ultimoN) = packet(pos)
    Next i
    
     
    pos = pos + 1
  End If
  
 
  tmpU = lon + 8 - pos
  ReDim Preserve packetNew(ultimoN + tmpU)
  For ti = 1 To tmpU
    packetNew(ultimoN + ti) = packet(pos + ti - 1)
  Next ti

  ultimoN = ultimoN + tmpU
  newSize = UBound(packetNew) - 7

  hb = HighByteOfLong(newSize)
  lb = LowByteOfLong(newSize)
  packetNew(6) = lb
  packetNew(7) = hb
  


  newSize = newSize + 6
  hb = HighByteOfLong(newSize)
  lb = LowByteOfLong(newSize)
  packetNew(0) = lb
  packetNew(1) = hb

'
  modSize = (newSize + 4) Mod 8
  aleat = 0
  If modSize > 0 Then
    aleat = (8 - modSize)
  End If


  For ti = 1 To aleat
      ReDim Preserve packetNew(ultimoN + 1)
      ultimoN = ultimoN + 1
      If rewriteEnabled = True Then
        packetNew(ultimoN) = &H0
      Else
        packetNew(ultimoN) = packet(ultimoN)
      End If
  Next ti
  
  newSize = UBound(packetNew) - 1
  hb = HighByteOfLong(newSize)
  lb = LowByteOfLong(newSize)
  packetNew(0) = lb
  packetNew(1) = hb
  



 ' Debug.Print "NEW SIZE=" & GoodHex(lb) & " " & GoodHex(hb)
  
  numChars = CLng(packet(pos))
  pos = pos + 1
  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Client can select this characters:"
  ResetCharList2 idConnection
  For i = 1 To numChars 'read all the characters on the list
    serID = CLng(packet(pos))
    pos = pos + 1
    lonCName = GetTheLong(packet(pos), packet(pos + 1))
    pos = pos + 2
    charName = ""
    For j = 1 To lonCName
      charName = charName & Chr(packet(pos))
      pos = pos + 1
    Next j
    servName = loadedServers(serID)
    servPort = loadedPorts(serID)
    servDOMAIN = loadedDomains(serID)


   ' servPort = GetGameServerPort(servName)
   ' servDOMAIN = GetGameServerDOMAIN(servName)
   
    AddCharServer2 idConnection, charName, servName, servIP1, servIP2, servIP3, servIP4, servPort, servDOMAIN
    'Debug.Print charName & "-> server #" & CStr(serID) & " (" & servName & ") = " & servDOMAIN & ":" & servPort
  Next i
  
  'Debug.Print frmMain.showAsStr(packet, True)
  

 'Debug.Print "NEW PCKT>" & frmMain.showAsStr(packetNEW, True)
 
  frmMain.UnifiedSendToClient idConnection, packetNew, False, True

  
  LastCharServerIndex = idConnection
  res = 1
returnTheResult:
  'LogOnFile "gotthem.txt", "AFTER (" & CStr(res) & ")> " & frmMain.showAsStr2(packet, 0) & vbCrLf
  PacketIPchange5b = res
End Function








Public Function PacketIPchange4(ByRef packet() As Byte, ByVal idConnection As Integer, ByVal strIP As String, Optional bstart As Long = 2) As Integer
  Dim lon As Long
  Dim motd As Long
  Dim numChars As Long
  Dim lonCName As Long
  Dim lonSName As Long
  Dim i As Integer
  Dim j As Integer
  Dim pos As Long
  Dim servName As String
  Dim servIP1 As Byte
  Dim servIP2 As Byte
  Dim servIP3 As Byte
  Dim servIP4 As Byte
  Dim servPort As Long
  Dim charName As String
  Dim hb As Byte
  Dim lb As Byte
  Dim res As Integer
  Dim adder As Long
  Dim serverIPport As String
  Dim lngIPid As Long
  Dim gamep As Long
  Dim doingdebugHere As Boolean
  Dim tipoBloque As Byte
  Dim servDOMAIN As String
  Dim totalServidores As Long
  Dim newServerName As String
  Dim newServerDomain As String
  Dim newServerPort As Long
  Dim loadedServers() As String
  Dim loadedPorts() As String
  Dim loadedDomains() As String
  Dim zz As Long
  Dim serID As Long
  Dim packetNew() As Byte
  Dim ultimoN As Long
  Dim ti As Long
  Dim lLocal As Long
  Dim b1Local As Byte
  Dim b2Local As Byte
  Dim tmpB As Byte
  Dim tmpU As Long
  Dim newSize As Long
  Dim modSize As Long
  Dim aleat As Long
  Dim realL As Long
  Dim ultimoTI As Long
  'On Error GoTo returnTheResult
  lLocal = Len(localstr)
  b1Local = HighByteOfLong(lLocal)
  b2Local = LowByteOfLong(lLocal)
  ultimoN = 0
  gamep = CLng(frmMain.sckClientGame(0).LocalPort)
  servDOMAIN = ""
 ' LogOnFile "gotthem.txt", frmMain.showAsStr2(packet, 0) & vbCrLf
  res = -1 'error
  'If (UseCrackd = True) Then
  '  adder = 2
  'Else
    adder = bstart - 2
  'End If
  Debug.Print "Receiving packet from login server..."
  If frmMain.chckAlter.value = 0 Then
    res = 1
    GoTo returnTheResult 'proxy user don't want to change this packet
  End If
  If packet(2 + adder) <> &H14 Then
    Debug.Print "This is not a list of character packet... Received type = " & GoodHex(packet(2 + adder))
    res = -1
    GoTo returnTheResult 'this is not a list of character packet
  End If

 If (TibiaVersionLong >= 1061) Then
  lon = GetTheLong(packet(0 + adder), packet(1 + adder))
  motd = GetTheLong(packet(3 + adder), packet(4 + adder))
  
  pos = motd + 5 + adder + 2 ' 2 new bytes after motd since tibia 10.61 :

  
  tipoBloque = packet(pos)
  pos = pos + 1
  ReDim packetNew(pos)
  
  ultimoN = pos - 2
  For ti = 0 To ultimoN
    packetNew(ti) = packet(ti)
    ultimoTI = ti
  Next ti
  ultimoN = ultimoN + 2


  ' 2 new bytes after motd are 0C 01
  ' Meaning? no idea! Just copy and ignore them
    ti = ultimoTI
    ti = ti + 1
    packetNew(ti) = packet(ti)
    ti = ti + 1
    packetNew(ti) = packet(ti)
 Else
   lon = GetTheLong(packet(0 + adder), packet(1 + adder))
  motd = GetTheLong(packet(3 + adder), packet(4 + adder))
  
  pos = motd + 5 + adder
  
  
  tipoBloque = packet(pos)
  pos = pos + 1
  ReDim packetNew(pos)
  
  ultimoN = pos
  For ti = 0 To ultimoN
    packetNew(ti) = packet(ti)
    ultimoTI = ti
  Next ti
 
 End If
 
 
  If tipoBloque = &H64 Then
    totalServidores = CLng(packet(pos))
    ReDim loadedServers(totalServidores - 1)
    ReDim loadedPorts(totalServidores - 1)
    ReDim loadedDomains(totalServidores - 1)
    For i = 1 To totalServidores
    
        pos = pos + 1
        ReDim Preserve packetNew(ultimoN + 1)
        ultimoN = ultimoN + 1
        packetNew(ultimoN) = packet(pos)
         
        serID = CLng(packet(pos))
        
        pos = pos + 1

        lonCName = GetTheLong(packet(pos), packet(pos + 1))
        
        ReDim Preserve packetNew(ultimoN + 1)
        ultimoN = ultimoN + 1
        packetNew(ultimoN) = packet(pos)
        
        ReDim Preserve packetNew(ultimoN + 1)
        ultimoN = ultimoN + 1
        packetNew(ultimoN) = packet(pos + 1)
        pos = pos + 2
        
        newServerName = ""
        For j = 1 To lonCName
          newServerName = newServerName & Chr(packet(pos))
          
          ReDim Preserve packetNew(ultimoN + 1)
          ultimoN = ultimoN + 1
          packetNew(ultimoN) = packet(pos)
          
          pos = pos + 1
        Next j
        lonSName = GetTheLong(packet(pos), packet(pos + 1))
        
        
        ReDim Preserve packetNew(ultimoN + 1)
        ultimoN = ultimoN + 1
        packetNew(ultimoN) = b2Local
        
        ReDim Preserve packetNew(ultimoN + 1)
        ultimoN = ultimoN + 1
        packetNew(ultimoN) = b1Local
        pos = pos + 2
        
        'read server name of character i
        newServerDomain = ""
        For j = 1 To lonSName
          
          newServerDomain = newServerDomain & Chr(packet(pos))
          tmpB = RewriteWithLocalIPcharPos(j)
          If tmpB = &H0 Then
             packet(pos) = tmpB
'            ReDim Preserve packetNEW(ultimoN + 1) ' borrar luego
'            ultimoN = ultimoN + 1
'            packetNEW(ultimoN) = packet(pos)
            
            
          Else
          
            packet(pos) = tmpB
            
            ReDim Preserve packetNew(ultimoN + 1)
            ultimoN = ultimoN + 1
            packetNew(ultimoN) = packet(pos)

            
          End If
          pos = pos + 1
        Next j
        newServerPort = GetTheLong(packet(pos), packet(pos + 1))
        loadedServers(serID) = newServerName
        loadedPorts(serID) = newServerPort
        loadedDomains(serID) = newServerDomain
        
   
        hb = HighByteOfLong(gamep)
        lb = LowByteOfLong(gamep)
        packet(pos) = lb
        packet(pos + 1) = hb
        
        
        
          ReDim Preserve packetNew(ultimoN + 1)
            ultimoN = ultimoN + 1
            packetNew(ultimoN) = packet(pos)

          ReDim Preserve packetNew(ultimoN + 1)
            ultimoN = ultimoN + 1
            packetNew(ultimoN) = packet(pos + 1)

        
        AddGameServer newServerName, "127.0.0.1:" & newServerPort, newServerDomain
        pos = pos + 2
        
        ReDim Preserve packetNew(ultimoN + 1)
        ultimoN = ultimoN + 1
        packetNew(ultimoN) = packet(pos)
    Next i
    
     
    pos = pos + 1
  End If
  
 
  tmpU = lon + 8 - pos
  ReDim Preserve packetNew(ultimoN + tmpU)
  For ti = 1 To tmpU
    packetNew(ultimoN + ti) = packet(pos + ti - 1)
  Next ti

  ultimoN = ultimoN + tmpU
  newSize = UBound(packetNew) - 7

  hb = HighByteOfLong(newSize)
  lb = LowByteOfLong(newSize)
  packetNew(6) = lb
  packetNew(7) = hb
  


  newSize = newSize + 6
  hb = HighByteOfLong(newSize)
  lb = LowByteOfLong(newSize)
  packetNew(0) = lb
  packetNew(1) = hb
  
'  Debug.Print "O> " & frmMain.showAsStr(packet, True)
'  Debug.Print "N> " & frmMain.showAsStr(packetNEW, True)
'  Debug.Print "OK"
'
  modSize = (newSize + 4) Mod 8
  aleat = 0
  If modSize > 0 Then
    aleat = (8 - modSize)
  End If


  For ti = 1 To aleat
      ReDim Preserve packetNew(ultimoN + 1)
      ultimoN = ultimoN + 1
      packetNew(ultimoN) = &H0
  Next ti
  
  newSize = UBound(packetNew) - 1
  hb = HighByteOfLong(newSize)
  lb = LowByteOfLong(newSize)
  packetNew(0) = lb
  packetNew(1) = hb
  
'  Debug.Print "O> " & frmMain.showAsStr(packet, True)
'  Debug.Print "N> " & frmMain.showAsStr(packetNEW, True)
'  Debug.Print "OK"


 ' Debug.Print "NEW SIZE=" & GoodHex(lb) & " " & GoodHex(hb)
  
  numChars = CLng(packet(pos))
  pos = pos + 1
  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Client can select this characters:"
  ResetCharList2 idConnection
  For i = 1 To numChars 'read all the characters on the list
    serID = CLng(packet(pos))
    pos = pos + 1
    lonCName = GetTheLong(packet(pos), packet(pos + 1))
    pos = pos + 2
    charName = ""
    For j = 1 To lonCName
      charName = charName & Chr(packet(pos))
      pos = pos + 1
    Next j
    servName = loadedServers(serID)
    servPort = loadedPorts(serID)
    servDOMAIN = loadedDomains(serID)


   ' servPort = GetGameServerPort(servName)
   ' servDOMAIN = GetGameServerDOMAIN(servName)
   
    AddCharServer2 idConnection, charName, servName, servIP1, servIP2, servIP3, servIP4, servPort, servDOMAIN
    'Debug.Print charName & "-> server #" & CStr(serID) & " (" & servName & ") = " & servDOMAIN & ":" & servPort
  Next i
  
  'Debug.Print frmMain.showAsStr(packet, True)
  
  'Debug.Print "ORIGINAL>" & frmMain.showAsStr(packet, True)
  'Debug.Print "NEW PCKT>" & frmMain.showAsStr(packetNEW, True)
 
  frmMain.UnifiedSendToClient idConnection, packetNew, False, True

  
  LastCharServerIndex = idConnection
  res = 1
returnTheResult:
  'LogOnFile "gotthem.txt", "AFTER (" & CStr(res) & ")> " & frmMain.showAsStr2(packet, 0) & vbCrLf
  PacketIPchange4 = res
End Function


Public Function PacketIPchange3(ByRef packet() As Byte, ByVal idConnection As Integer, ByVal strIP As String, Optional bstart As Long = 2) As Integer
  Dim lon As Long
  Dim motd As Long
  Dim numChars As Long
  Dim lonCName As Long
  Dim lonSName As Long
  Dim i As Integer
  Dim j As Integer
  Dim pos As Long
  Dim servName As String
  Dim servIP1 As Byte
  Dim servIP2 As Byte
  Dim servIP3 As Byte
  Dim servIP4 As Byte
  Dim servPort As Long
  Dim charName As String
  Dim hb As Byte
  Dim lb As Byte
  Dim res As Integer
  Dim adder As Long
  Dim serverIPport As String
  Dim lngIPid As Long
  Dim gamep As Long
  Dim doingdebugHere As Boolean
  
  On Error GoTo returnTheResult
 ' LogOnFile "gotthem.txt", frmMain.showAsStr2(packet, 0) & vbCrLf
  res = -1 'error
  'If (UseCrackd = True) Then
  '  adder = 2
  'Else
    adder = bstart - 2
  'End If
  'Debug.Print "got a char packet"
  If frmMain.chckAlter.value = 0 Then
    res = 1
    GoTo returnTheResult 'proxy user don't want to change this packet
  End If
  If packet(2 + adder) <> &H14 Then
    'MsgBox "2.received " & GoodHex(packet(2)), vbOKOnly + vbInformation, "DEBUG"
    res = -1
    GoTo returnTheResult 'this is not a list of character packet
  End If
  lon = GetTheLong(packet(0 + adder), packet(1 + adder))
  motd = GetTheLong(packet(3 + adder), packet(4 + adder))
  numChars = CLng(packet(motd + 6 + adder))
  pos = motd + 7 + adder
  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Client can select this characters:"
  ResetCharList2 idConnection
  For i = 1 To numChars 'read all the characters on the list
    lonCName = GetTheLong(packet(pos), packet(pos + 1))
    pos = pos + 2
    charName = ""
    For j = 1 To lonCName
      charName = charName & Chr(packet(pos))
      pos = pos + 1
    Next j
    lonSName = GetTheLong(packet(pos), packet(pos + 1))
    pos = pos + 2
    'read server name of character i
    servName = ""
    For j = 1 To lonSName
      servName = servName & Chr(packet(pos))
      pos = pos + 1
    Next j
    ' save IP
    servIP1 = packet(pos)
    servIP2 = packet(pos + 1)
    servIP3 = packet(pos + 2)
    servIP4 = packet(pos + 3)
    'Debug.Print servName & "=" & CLng(servIP1) & "." & CLng(servIP2) & "." & CLng(servIP3) & "." & CLng(servIP4)
    ' change IP
    If AlternativeBinding = 0 Then
        lngIPid = 1
    Else
        If Left$(strIP, 8) = "127.0.0." Then
            lngIPid = CLng(Right$(strIP, Len(strIP) - 8))
        Else
            lngIPid = 1
        End If
    End If
    'temp
    doingdebugHere = False
    If doingdebugHere = False Then
    packet(pos) = 127
    packet(pos + 1) = 0
    packet(pos + 2) = 0
    packet(pos + 3) = lngIPid
    ' save port
    servPort = CLng(packet(pos + 5)) * 256 + CLng(packet(pos + 4))
    'Debug.Print "Original game PORT=" & CLng(servPort)
    ' change port
    ' split the port into high and low bytes
    
    If TibiaVersionLong >= 841 Then
       gamep = frmMain.sckClientGame(0).LocalPort
    Else
       gamep = CLng(frmMain.txtClientGameP.Text)
    End If
    hb = HighByteOfLong(gamep)
    lb = LowByteOfLong(gamep)
    packet(pos + 4) = lb
    packet(pos + 5) = hb
    End If
    pos = pos + 6
    If (TibiaVersionLong >= 971) Then
      pos = pos + 1 ' skip strange byte
    End If
    ' add the relation of character name - server data in the list :
   ' frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & charName
    serverIPport = fixThreeDigits(servIP1) & "." & fixThreeDigits(servIP2) & "." & _
     fixThreeDigits(servIP3) & "." & fixThreeDigits(servIP4) & ":" & CStr(servPort)
    AddGameServer servName, serverIPport
    AddCharServer2 idConnection, charName, servName, servIP1, servIP2, servIP3, servIP4, servPort
    'Debug.Print charName & "->" & serverIPport
  Next i
  'setCharListPosition2 idConnection
  ' 10000 days premium account mirage ;)
  ' packet(lon) = &H10
  ' packet(lon + 1) = &H27
  LastCharServerIndex = idConnection
  res = 1
returnTheResult:
  'LogOnFile "gotthem.txt", "AFTER (" & CStr(res) & ")> " & frmMain.showAsStr2(packet, 0) & vbCrLf
  PacketIPchange3 = res
End Function






Public Function PacketIPchange2(ByRef packet() As Byte, ByVal idConnection As Integer, ByVal strIP As String, Optional bstart As Long = 2) As Integer
  Dim lon As Long
  Dim motd As Long
  Dim numChars As Long
  Dim lonCName As Long
  Dim lonSName As Long
  Dim i As Integer
  Dim j As Integer
  Dim pos As Long
  Dim servName As String
  Dim servIP1 As Byte
  Dim servIP2 As Byte
  Dim servIP3 As Byte
  Dim servIP4 As Byte
  Dim servPort As Long
  Dim charName As String
  Dim hb As Byte
  Dim lb As Byte
  Dim res As Integer
  Dim adder As Long
  Dim serverIPport As String
  Dim lngIPid As Long
  Dim gamep As Long
  Dim doingdebugHere As Boolean
  On Error GoTo returnTheResult
  'LogOnFile "gotthem.txt", frmMain.showAsStr2(packet, 0) & vbCrLf
  res = -1 'error
  'If (UseCrackd = True) Then
  '  adder = 2
  'Else
    adder = bstart - 2
  'End If
  'Debug.Print "got a char packet"
  If frmMain.chckAlter.value = 0 Then
    res = 1
    GoTo returnTheResult 'proxy user don't want to change this packet
  End If
  If packet(2 + adder) <> &H14 Then
    'MsgBox "2.received " & GoodHex(packet(2)), vbOKOnly + vbInformation, "DEBUG"
    res = -1
    GoTo returnTheResult 'this is not a list of character packet
  End If
  lon = GetTheLong(packet(0 + adder), packet(1 + adder))
  motd = GetTheLong(packet(3 + adder), packet(4 + adder))
  numChars = CLng(packet(motd + 6 + adder))
  pos = motd + 7 + adder
  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Client can select this characters:"
  ResetCharList2 idConnection
  For i = 1 To numChars 'read all the characters on the list
    lonCName = GetTheLong(packet(pos), packet(pos + 1))
    pos = pos + 2
    charName = ""
    For j = 1 To lonCName
      charName = charName & Chr(packet(pos))
      pos = pos + 1
    Next j
    lonSName = GetTheLong(packet(pos), packet(pos + 1))
    pos = pos + 2
    'read server name of character i
    servName = ""
    For j = 1 To lonSName
      servName = servName & Chr(packet(pos))
      pos = pos + 1
    Next j
    ' save IP
    servIP1 = packet(pos)
    servIP2 = packet(pos + 1)
    servIP3 = packet(pos + 2)
    servIP4 = packet(pos + 3)
    'Debug.Print servName & "=" & CLng(servIP1) & "." & CLng(servIP2) & "." & CLng(servIP3) & "." & CLng(servIP4)
    ' change IP
    If AlternativeBinding = 0 Then
        lngIPid = 1
    Else
        If Left$(strIP, 8) = "127.0.0." Then
            lngIPid = CLng(Right$(strIP, Len(strIP) - 8))
        Else
            lngIPid = 1
        End If
    End If
    'temp
    doingdebugHere = False
    If doingdebugHere = False Then
    packet(pos) = 127
    packet(pos + 1) = 0
    packet(pos + 2) = 0
    packet(pos + 3) = lngIPid
    ' save port
    servPort = CLng(packet(pos + 5)) * 256 + CLng(packet(pos + 4))
    'Debug.Print "Original game PORT=" & CLng(servPort)
    ' change port
    ' split the port into high and low bytes
    
    If TibiaVersionLong >= 841 Then
       gamep = frmMain.sckClientGame(0).LocalPort
    Else
       gamep = CLng(frmMain.txtClientGameP.Text)
    End If
    hb = HighByteOfLong(gamep)
    lb = LowByteOfLong(gamep)
    packet(pos + 4) = lb
    packet(pos + 5) = hb
    End If
    pos = pos + 6
    If (TibiaVersionLong >= 971) Then
      pos = pos + 1 ' skip strange byte
    End If
    ' add the relation of character name - server data in the list :
   ' frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & charName
    serverIPport = fixThreeDigits(servIP1) & "." & fixThreeDigits(servIP2) & "." & _
     fixThreeDigits(servIP3) & "." & fixThreeDigits(servIP4) & ":" & CStr(servPort)
    AddGameServer servName, serverIPport
    AddCharServer2 idConnection, charName, servName, servIP1, servIP2, servIP3, servIP4, servPort
    'Debug.Print charName & "->" & serverIPport
  Next i
  'setCharListPosition2 idConnection
  ' 10000 days premium account mirage ;)
  ' packet(lon) = &H10
  ' packet(lon + 1) = &H27
  LastCharServerIndex = idConnection
  res = 1
returnTheResult:
  'LogOnFile "gotthem.txt", "AFTER (" & CStr(res) & ")> " & frmMain.showAsStr2(packet, 0) & vbCrLf
  PacketIPchange2 = res
End Function


Public Function GoodHex(b As Byte) As String
  Dim res As String
  res = Hex(b)
  If Len(res) = 1 Then
    GoodHex = "0" & res 'add a zero if VB conversion only return 1 character
  Else
    GoodHex = res
  End If
End Function
Public Function HighByteOfLong(address As Long) As Byte
  Dim h As Byte
  h = CByte(address \ 256) ' high byte
  HighByteOfLong = h
End Function

Public Function LowByteOfLong(address As Long) As Byte
  Dim h As Byte
  Dim l As Byte
  h = CByte(address \ 256)
  l = CByte(address - (CLng(h) * 256)) ' low byte
  LowByteOfLong = l
End Function

Public Function FromHexToDec(str As String) As Byte
  Dim res As Byte
  ' converts 1 character string
  ' to a byte
  res = 16 'reserved to error
  Select Case str
  Case "0"
    res = 0
  Case "1"
    res = 1
  Case "2"
    res = 2
  Case "3"
    res = 3
  Case "4"
    res = 4
  Case "5"
    res = 5
  Case "6"
    res = 6
  Case "7"
    res = 7
  Case "8"
    res = 8
  Case "9"
    res = 9
  Case "A", "a"
    res = 10
  Case "B", "b"
    res = 11
  Case "C", "c"
    res = 12
  Case "D", "d"
    res = 13
  Case "E", "e"
    res = 14
  Case "F", "f"
    res = 15
  End Select
  FromHexToDec = res
End Function

'Public Sub GenerateSelectionPacket(ByRef genPacket() As Byte, ByRef lastServerName As String)
 '   Dim l As Integer
 '   Dim c As String
   ' Dim i As Integer
   ' l = Len(lastServerName)
    'ReDim genPacket(l)
    'For i = 0 To l - 1
     '   c = Mid(lastServerName, i + 1, 1)
      '  genPacket(i) = Asc(c)
    'Next i
    'genPacket(l) = &HA
'End Sub
Public Function ConvSToAscii(ByRef b() As Byte) As String
  ' converts an 2 character string that represents a byte
  ' to a 1 ascii character (1 character string)
  Dim res As String
  Dim i As Integer
  #If FinalMode Then
  On Error GoTo endF
  #End If
  res = ""
  For i = 0 To UBound(b)
  If Not (b(i) = &HA) Then
    res = res & Chr(b(i))
    End If
  Next i
endF:
  ConvSToAscii = res
End Function
Public Function ConvBToAscii(b As Byte) As String
  ' converts an 2 character string that represents a byte
  ' to a 1 ascii character (1 character string)
  Dim res As String
  #If FinalMode Then
  On Error GoTo endF
  #End If
  res = "?"
  res = Chr(b)

endF:
  ConvBToAscii = res
End Function

Public Function ConvToAscii(str As String) As String
  ' converts an 2 character string that represents a byte
  ' to a 1 ascii character (1 character string)
  Dim res As String
  Dim intRes As Byte
  Dim leftChr As String
  Dim rightChr As String
  Dim leftVal As Byte
  Dim rightVal As Byte
  #If FinalMode Then
  On Error GoTo endF
  #End If
  res = "?"
  leftChr = Left(str, 1)
  rightChr = Right(str, 1)
  leftVal = FromHexToDec(leftChr)
  rightVal = FromHexToDec(rightChr)
  If leftVal = 16 Or rightVal = 16 Then
    res = "?"
  Else
    res = Chr(leftVal * 16 + rightVal)
  End If
endF:
  ConvToAscii = res
End Function

Public Sub SafeCastCheatStringSPACES(ByRef strFunction As String, ByVal idConnection As Integer, ByVal strinput As String, Optional ByVal withDelay As Long = 0)
    Dim res As Integer
    Dim conv As Byte
    Dim strByte As String
    Dim leftChr As String
    Dim rightChr As String
    Dim leftVal As Byte
    Dim rightVal As Byte
    Dim i As Long
    Dim packet() As Byte
    Dim ub As Long
    Dim nby As Long
    Dim valid1 As Integer
    Dim valid2 As Integer
    #If FinalMode Then
    On Error GoTo endF
    #End If
    res = -1
    i = 2
    ReDim packet(2)
    ' analyze the string
    While Len(strinput) > 0
      If Left(strinput, 1) = " " Or Left(strinput, 1) = vbCr Or Left(strinput, 1) = vbLf Then
        strinput = Right(strinput, Len(strinput) - 1)
      Else
        ' get the byte from the 2 characters
        strByte = Left(strinput, 2)
        strinput = Right(strinput, Len(strinput) - 2)
        leftChr = Left(strByte, 1)
        rightChr = Right(strByte, 1)
        leftVal = FromHexToDec(leftChr)
        rightVal = FromHexToDec(rightChr)
        If leftVal = 16 Or rightVal = 16 Then
          GoTo endF ' error
        Else
          conv = leftVal * 16 + rightVal
        End If
        ' add a byte to the packet
        ReDim Preserve packet(i)
        packet(i) = conv
        i = i + 1
      End If
    Wend
    ub = UBound(packet) - 1
    packet(0) = LowByteOfLong(ub)
    packet(1) = HighByteOfLong(ub)
   ' Debug.Print "Function " & strFunction & " >> " & frmMain.showAsStr(packet, True)
    If (withDelay > 0) Then
        wait withDelay
    End If
    frmMain.UnifiedSendToServerGame idConnection, packet, True
    DoEvents
endF:
  res = 0

End Sub

Public Sub SafeCastCheatString(ByRef strFunction As String, ByVal idConnection As Integer, ByVal strinput As String, Optional ByVal withDelay As Long = 0)
    Dim res As Integer
    Dim conv As Byte
    Dim strByte As String
    Dim leftChr As String
    Dim rightChr As String
    Dim leftVal As Byte
    Dim rightVal As Byte
    Dim i As Long
    Dim packet() As Byte
    Dim ub As Long
    Dim nby As Long
    Dim valid1 As Integer
    Dim valid2 As Integer
    #If FinalMode = 1 Then
    On Error GoTo endF
    #End If
    res = -1
    nby = (Len(strinput) + 1) / 3
    ReDim packet(nby + 1)
    nby = nby - 1
    For i = 0 To nby
      leftChr = Mid$(strinput, 1 + (3 * i), 1)
      rightChr = Mid$(strinput, 2 + (3 * i), 1)
      Select Case leftChr
        Case "0"
          valid1 = 0
        Case "1"
          valid1 = 1
        Case "2"
          valid1 = 2
        Case "3"
          valid1 = 3
        Case "4"
          valid1 = 4
        Case "5"
          valid1 = 5
        Case "6"
          valid1 = 6
        Case "7"
          valid1 = 7
        Case "8"
          valid1 = 8
        Case "9"
          valid1 = 9
        Case "A", "a"
          valid1 = 10
        Case "B", "b"
          valid1 = 11
        Case "C", "c"
          valid1 = 12
        Case "D", "d"
          valid1 = 13
        Case "E", "e"
          valid1 = 14
        Case "F", "f"
          valid1 = 15
        Case " ", vbLf
          #If FinalMode = 0 Then
          Debug.Print "Caught spaces at " & strFunction
          #End If
          SafeCastCheatStringSPACES strFunction, idConnection, strinput
          Exit Sub
        Case Else
          valid1 = 16
      End Select
      Select Case rightChr
        Case "0"
          valid2 = 0
        Case "1"
          valid2 = 1
        Case "2"
          valid2 = 2
        Case "3"
          valid2 = 3
        Case "4"
          valid2 = 4
        Case "5"
          valid2 = 5
        Case "6"
          valid2 = 6
        Case "7"
          valid2 = 7
        Case "8"
          valid2 = 8
        Case "9"
          valid2 = 9
        Case "A", "a"
          valid2 = 10
        Case "B", "b"
          valid2 = 11
        Case "C", "c"
          valid2 = 12
        Case "D", "d"
          valid2 = 13
        Case "E", "e"
          valid2 = 14
        Case "F", "f"
          valid2 = 15
        Case " ", vbLf
          #If FinalMode = 0 Then
          Debug.Print "Caught spaces at " & strFunction
          #End If
          SafeCastCheatStringSPACES strFunction, idConnection, strinput, withDelay
          Exit Sub
        Case Else
          valid2 = 16
      End Select
      If ((valid1 = 16) Or (valid2 = 16)) Then
        GoTo endF ' error
      Else
        conv = valid1 * 16 + valid2
        packet(i + 2) = conv
      End If
    Next i
    ub = UBound(packet) - 1
    packet(0) = LowByteOfLong(ub)
    packet(1) = HighByteOfLong(ub)
    'Debug.Print "Function " & strFunction & " >> " & frmMain.showAsStr(packet, True)
    If (withDelay > 0) Then
        wait withDelay
    End If
    frmMain.UnifiedSendToServerGame idConnection, packet, True
    DoEvents
    Exit Sub
endF:
Debug.Print "ERROR"
  res = 0

End Sub

Public Function GetCheatPacket(ByRef packet() As Byte, strinput As String) As Integer
  Dim res As Integer
  Dim conv As Byte
  Dim strByte As String
  Dim leftChr As String
  Dim rightChr As String
  Dim leftVal As Byte
  Dim rightVal As Byte
  Dim i As Long
  #If FinalMode Then
  On Error GoTo endF
  #End If
  res = -1
  i = 0
  ' analyze the string
  While Len(strinput) > 0
    If Left(strinput, 1) = " " Or Left(strinput, 1) = vbCr Or Left(strinput, 1) = vbLf Then
      strinput = Right(strinput, Len(strinput) - 1)
    Else
      ' get the byte from the 2 characters
      strByte = Left(strinput, 2)
      strinput = Right(strinput, Len(strinput) - 2)
      leftChr = Left(strByte, 1)
      rightChr = Right(strByte, 1)
      leftVal = FromHexToDec(leftChr)
      rightVal = FromHexToDec(rightChr)
      If leftVal = 16 Or rightVal = 16 Then
        GoTo endF ' error
      Else
        conv = leftVal * 16 + rightVal
      End If
      ' add a byte to the packet
      ReDim Preserve packet(i)
      packet(i) = conv
      i = i + 1
    End If
  Wend
  res = 0
endF:
  GetCheatPacket = res
End Function

Public Function GetCharListPositionPre(idConnection As Integer, ByRef selectedcharacter As String) As Integer
  ' get the list position of the selected character
  On Error GoTo returnTheResult
  Dim theindex As Byte
  Dim theindexLNG As Long
  theindexLNG = ReadCurrentAddress(ProcessID(idConnection), adrSelectedCharIndex, -1, True)

  GetCharListPositionPre = CInt(theindexLNG)
  Exit Function
returnTheResult:
  GetCharListPositionPre = -1
End Function

Public Sub setCharListPosition2(idConnection As Integer)
  ' get the list position of the selected character
  ' and also return selectedcharacter byref
  ' (?)
  Dim realADR As Long
  realADR = ReadCurrentAddress(ProcessID(idConnection), adrSelectedCharIndex, 0, False)
  If realADR <> 0 Then
    Memory_WriteByte realADR, &HFF, ProcessID(idConnection)
  End If
End Sub


Public Function GetCharListPosition2(idConnection As Integer, ByRef selectedcharacter As String) As Integer
  ' get the list position of the selected character
  ' and also return selectedcharacter byref
  On Error GoTo returnTheResult
  Dim theindex As Byte
  Dim theindexLNG As Long
  Dim gtc As Long
  Dim maxgtc As Long
  maxgtc = GetTickCount() + 10000
  
'  theindex = &HFF
'  While (theindex = &HFF)
'    theindex = Memory_ReadByte(adrSelectedCharIndex, ProcessID(idConnection))
'    gtc = GetTickCount()
'    If gtc > maxgtc Then
'        GetCharListPosition2 = -1
'        Exit Function
'    End If
'    DoEvents
'  Wend
'wait (500)
  If ProcessID(idConnection) = -1 Then
    GetCharListPosition2 = -1
    Exit Function
  End If
    theindexLNG = ReadCurrentAddress(ProcessID(idConnection), adrSelectedCharIndex, -1, True)
    If theindexLNG = -1 Then
        GetCharListPosition2 = -1
        Exit Function
    Else
        theindex = CByte(theindexLNG)
    End If
'    gtc = GetTickCount()
  If theindex >= CharacterList2(idConnection).numItems Then
    GetCharListPosition2 = -1
    Exit Function
  End If
  selectedcharacter = CharacterList2(idConnection).item(theindex).CharacterName
  
  
  
  GetCharListPosition2 = CInt(theindex)
  Exit Function
returnTheResult:
  GetCharListPosition2 = -1
End Function






Public Function GiveServerError(str As String, idConnection As Integer) As Long
  Dim cheatpacket() As Byte
  Dim longP As Long
  Dim longSend As Long
  Dim hb As Byte
  Dim lb As Byte
  Dim pos As Integer
  Dim i As Integer
  Dim strCad As String
  Dim chCad As String
  Dim aRes As Long
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  aRes = -1
  If GameConnected(idConnection) = False Then
    ' not a valid ID, do nothing
    GiveServerError = aRes
    Exit Function
  End If
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & CharacterName(idConnection) & " closed with this message: " & str
  GiveServerError = 0
  Exit Function ' giving problems lately....
  longSend = Len(str)
  ReDim cheatpacket(longSend + 4)
  longP = 3 + longSend
  hb = HighByteOfLong(CLng(longP))
  lb = LowByteOfLong(CLng(longP))
  cheatpacket(0) = lb
  cheatpacket(1) = hb
  cheatpacket(2) = &H14 'server error
  pos = 3
  ' message - lenght
  hb = HighByteOfLong(CLng(longSend))
  lb = LowByteOfLong(CLng(longSend))
  cheatpacket(pos) = lb
  pos = pos + 1
  cheatpacket(pos) = hb
  pos = pos + 1
  ' message - text
  strCad = str
  For i = 1 To longSend
    chCad = Left(strCad, 1)
    strCad = Right(strCad, Len(strCad) - 1)
    cheatpacket(pos) = Asc(chCad)
    pos = pos + 1
  Next i
  ' send the packet
  frmMain.UnifiedSendToClientGame idConnection, cheatpacket
  GiveServerError = aRes
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at GiveServerError #"
  frmMain.DoCloseActions idConnection
  DoEvents
  GiveServerError = -1
End Function

Public Function SendChannelMessage(idConnection As Integer, strSend As String, _
 channelB1 As Byte, channelB2 As Byte) As Long
  Dim cheatpacket() As Byte
  Dim longSend As Long
  Dim longP As Long
  Dim i As Integer
  Dim strCad As String
  Dim chCad As String
  Dim totalL As Long
  Dim hb As Byte
  Dim lb As Byte
  Dim pos As Long
  longSend = Len(strSend)
  totalL = 7 + longSend

  ReDim cheatpacket(totalL)
  longP = totalL - 1
  hb = HighByteOfLong(CLng(longP))
  lb = LowByteOfLong(CLng(longP))
  ' 4C 00 96 05 00 00 46 00 74 65 73 74 69 6E 67 20 2E 2E 2E 20 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61 61
  cheatpacket(0) = lb
  cheatpacket(1) = hb
  cheatpacket(2) = &H96 'message
  If TibiaVersionLong < 820 Then
    cheatpacket(3) = &H5  'to channel
  Else
    cheatpacket(3) = &H7  'to channel
  End If
  cheatpacket(4) = channelB1 'channel ID
  cheatpacket(5) = channelB2
  pos = 6
  hb = HighByteOfLong(CLng(longSend))
  lb = LowByteOfLong(CLng(longSend))
  cheatpacket(pos) = lb
  pos = pos + 1
  cheatpacket(pos) = hb
  pos = pos + 1
  ' from - text
  strCad = strSend
  For i = 1 To longSend
    chCad = Left(strCad, 1)
    strCad = Right(strCad, Len(strCad) - 1)
    cheatpacket(pos) = Asc(chCad)
    pos = pos + 1
  Next i
  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & frmMain.showAsStr2(cheatpacket, True)
  
  frmMain.UnifiedSendToServerGame idConnection, cheatpacket, True
  SendChannelMessage = 0
  
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at SendChannelMessage #"
  frmMain.DoCloseActions idConnection
  DoEvents
  SendChannelMessage = -1
End Function
Public Function GiveChannelMessage(idConnection As Integer, strSend As String, strFrom As String, _
 channelB1 As Byte, channelB2 As Byte) As Long
  Dim cheatpacket() As Byte
  Dim longP As Long
  Dim longSend As Long
  Dim longFrom As Long
  Dim totalL As Long
  Dim hb As Byte
  Dim lb As Byte
  Dim pos As Integer
  Dim i As Integer
  Dim strCad As String
  Dim chCad As String
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  
  longSend = Len(strSend)
  longFrom = Len(strFrom)
  ' totalL = num of bytes -1 (because array index start from 0)
  totalL = 9 + longSend + longFrom
  If TibiaVersionLong > 760 Then
    totalL = totalL + 4
  End If
  If TibiaVersionLong >= 773 Then
    totalL = totalL + 2
  End If
  ReDim cheatpacket(totalL)
  ' packet header = num of bytes -2 , that is totalL -1
  longP = totalL - 1
  hb = HighByteOfLong(CLng(longP))
  lb = LowByteOfLong(CLng(longP))
  cheatpacket(0) = lb
  cheatpacket(1) = hb
  cheatpacket(2) = &HAA 'message
  If TibiaVersionLong > 760 Then
    cheatpacket(3) = &H0
    cheatpacket(4) = &H0
    cheatpacket(5) = &H0
    cheatpacket(6) = &H0
    pos = 7
  Else
    pos = 3
  End If
  ' from - lenght
  hb = HighByteOfLong(CLng(longFrom))
  lb = LowByteOfLong(CLng(longFrom))
  cheatpacket(pos) = lb
  pos = pos + 1
  cheatpacket(pos) = hb
  pos = pos + 1
  ' from - text
  strCad = strFrom
  For i = 1 To longFrom
    chCad = Left(strCad, 1)
    strCad = Right(strCad, Len(strCad) - 1)
    cheatpacket(pos) = Asc(chCad)
    pos = pos + 1
  Next i
  If TibiaVersionLong >= 773 Then
    cheatpacket(pos) = fakemessagesLevel1
    pos = pos + 1
    cheatpacket(pos) = fakemessagesLevel2
    pos = pos + 1
    totalL = totalL + 2
  End If
  cheatpacket(pos) = oldmessage_H5 ' to channel
  pos = pos + 1
  cheatpacket(pos) = channelB1 ' channel ID byte 1
  pos = pos + 1
  cheatpacket(pos) = channelB2 ' channel ID byte 2
  pos = pos + 1
  ' message - lenght
  hb = HighByteOfLong(CLng(longSend))
  lb = LowByteOfLong(CLng(longSend))
  cheatpacket(pos) = lb
  pos = pos + 1
  cheatpacket(pos) = hb
  pos = pos + 1
  ' message - text
  strCad = strSend
  For i = 1 To longSend
    chCad = Left(strCad, 1)
    strCad = Right(strCad, Len(strCad) - 1)
    cheatpacket(pos) = Asc(chCad)
    pos = pos + 1
  Next i
  ' send the packet
  
  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & frmMain.showAsStr2(cheatpacket, True)
  
  frmMain.UnifiedSendToClientGame idConnection, cheatpacket

  
  GiveChannelMessage = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at GiveChannelMessage #"
  frmMain.DoCloseActions idConnection
  DoEvents
  GiveChannelMessage = -1
End Function

Public Function SendMessageToClient(idConnection As Integer, strSend As String, strFrom As String) As Long
  Dim cheatpacket() As Byte
  Dim longP As Long
  Dim longSend As Long
  Dim longFrom As Long
  Dim totalL As Long
  Dim hb As Byte
  Dim lb As Byte
  Dim pos As Integer
  Dim i As Integer
  Dim strCad As String
  Dim chCad As String
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  
  If frmStealth.chkStealthMessages.value = 1 Then
    stealthLog(idConnection) = stealthLog(idConnection) & vbCrLf & TibiaTimestamp() & strFrom & " [" & CStr(fakemessagesLevel) & "]: " & Replace(strSend, vbLf, vbCrLf)
    If idConnection = stealthIDselected Then
        frmStealth.UpdateValues
    End If
    'SendMessageToClient = 0
    'Exit Function
  End If
  
  
  longSend = Len(strSend)
  longFrom = Len(strFrom)
  ' totalL = num of bytes -1 (because array index start from 0)
  totalL = 7 + longSend + longFrom
  If TibiaVersionLong > 760 Then
    totalL = totalL + 4
  End If
  If TibiaVersionLong >= 773 Then
    totalL = totalL + 2
  End If
  ReDim cheatpacket(totalL)
  ' packet header = num of bytes -2 , that is totalL -1
  longP = totalL - 1
  hb = HighByteOfLong(CLng(longP))
  lb = LowByteOfLong(CLng(longP))
  cheatpacket(0) = lb
  cheatpacket(1) = hb
  cheatpacket(2) = &HAA 'message
  If TibiaVersionLong > 760 Then
    cheatpacket(3) = &H0
    cheatpacket(4) = &H0
    cheatpacket(5) = &H0
    cheatpacket(6) = &H0
    pos = 7
  Else
    pos = 3
  End If
  ' from - lenght
  hb = HighByteOfLong(CLng(longFrom))
  lb = LowByteOfLong(CLng(longFrom))
  cheatpacket(pos) = lb
  cheatpacket(pos + 1) = hb
  pos = pos + 2
  ' from - text
  strCad = strFrom
  For i = 1 To longFrom
    chCad = Left(strCad, 1)
    strCad = Right(strCad, Len(strCad) - 1)
    cheatpacket(pos) = Asc(chCad)
    pos = pos + 1
  Next i
    
  If TibiaVersionLong >= 773 Then
    cheatpacket(pos) = fakemessagesLevel1
    pos = pos + 1
    cheatpacket(pos) = fakemessagesLevel2
    pos = pos + 1
  End If
  cheatpacket(pos) = oldmessage_H4
  pos = pos + 1
  ' message - lenght
  hb = HighByteOfLong(CLng(longSend))
  lb = LowByteOfLong(CLng(longSend))
  cheatpacket(pos) = lb
  pos = pos + 1
  cheatpacket(pos) = hb
  pos = pos + 1
  ' message - text
  strCad = strSend
  For i = 1 To longSend
    chCad = Left(strCad, 1)
    strCad = Right(strCad, Len(strCad) - 1)
    cheatpacket(pos) = Asc(chCad)
    pos = pos + 1
  Next i
  ' send the packet
  frmMain.UnifiedSendToClientGame idConnection, cheatpacket
  SendMessageToClient = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at SendMessageToClient #"
  frmMain.DoCloseActions idConnection
  DoEvents
  SendMessageToClient = -1
End Function

Public Function SendSystemMessageToClient(idConnection As Integer, strSend As String) As Long
' &HB4 , &H14 -login
'        &H17 - temp
      ' system message (msgtype,lenght and message)
      'lonN = GetTheLong(packet(pos + 2), packet(pos + 3))
     ' pos = pos + 4 + lonN
  Dim cheatpacket() As Byte
  Dim longP As Long
  Dim longSend As Long
  Dim longFrom As Long
  Dim totalL As Long
  Dim hb As Byte
  Dim lb As Byte
  Dim pos As Integer
  Dim i As Integer
  Dim strCad As String
  Dim chCad As String
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  ' 8.61: 1F 00 B4 14 1B 00 59 6F 75 20 63 61 6E 6E 6F 74 20 75 73 65 20 74 68 69 73 20 6F 62 6A 65 63 74 2E
  
  If (idConnection = 0) Then
    SendSystemMessageToClient = -1
    Exit Function
  End If
  
  If frmStealth.chkStealthMessages.value = 1 Then
    stealthLog(idConnection) = stealthLog(idConnection) & vbCrLf & TibiaTimestamp() & "(sysmessage) " & Replace(strSend, vbLf, vbCrLf)
    If idConnection = stealthIDselected Then
        frmStealth.UpdateValues
    End If
    SendSystemMessageToClient = 0
    Exit Function
  End If
  
  If GameConnected(idConnection) = True Then
  ' If GameConnected(idconnection) = True And sentFirstPacket(idconnection) = True Then
  longSend = Len(strSend)
  ' totalL = num of bytes -1 (because array index start from 0)
  totalL = 5 + longSend
  ReDim cheatpacket(totalL)
  ' packet header = num of bytes -2 , that is totalL -1
  longP = totalL - 1
  hb = HighByteOfLong(CLng(longP))
  lb = LowByteOfLong(CLng(longP))
  cheatpacket(0) = lb
  cheatpacket(1) = hb
  cheatpacket(2) = &HB4 ' system message
  If TibiaVersionLong >= 1036 Then
    cheatpacket(3) = &H14 ' temporal
  ElseIf TibiaVersionLong >= 872 Then
    cheatpacket(3) = &H13 ' temporal
  ElseIf TibiaVersionLong >= 861 Then
    cheatpacket(3) = &H14 ' temporal
  ElseIf TibiaVersionLong >= 840 Then
    cheatpacket(3) = &H1A ' temporal
  ElseIf TibiaVersionLong >= 820 Then
    cheatpacket(3) = &H19 ' temporal
  Else
    cheatpacket(3) = &H17 ' temporal
  End If
  pos = 4
  ' message - lenght
  hb = HighByteOfLong(CLng(longSend))
  lb = LowByteOfLong(CLng(longSend))
  cheatpacket(pos) = lb
  pos = pos + 1
  cheatpacket(pos) = hb
  pos = pos + 1
  ' message - text
  strCad = strSend
  For i = 1 To longSend
    chCad = Left(strCad, 1)
    strCad = Right(strCad, Len(strCad) - 1)
    cheatpacket(pos) = Asc(chCad)
    pos = pos + 1
  Next i
  ' send the packet
  
  frmMain.UnifiedSendToClientGame idConnection, cheatpacket
  End If
  SendSystemMessageToClient = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & _
   idConnection & " failed to use SendSystemMessageToClient : " & strSend & " . Err number = " & _
   CStr(Err.Number) & " ; Err description = " & Err.Description
  'frmMain.DoCloseActions idConnection
  'DoEvents
  SendSystemMessageToClient = -1
End Function


Public Function SendCustomSystemMessageToClient(idConnection As Integer, strSend As String, thecolor As Byte) As Long
' &HB4 , &H14 -login
'        &H17 - temp
      ' system message (msgtype,lenght and message)
      'lonN = GetTheLong(packet(pos + 2), packet(pos + 3))
     ' pos = pos + 4 + lonN
  Dim cheatpacket() As Byte
  Dim longP As Long
  Dim longSend As Long
  Dim longFrom As Long
  Dim totalL As Long
  Dim hb As Byte
  Dim lb As Byte
  Dim pos As Integer
  Dim i As Integer
  Dim strCad As String
  Dim chCad As String
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  If (idConnection = 0) Then
    SendCustomSystemMessageToClient = -1
    Exit Function
  End If
  
  If frmStealth.chkStealthMessages.value = 1 Then
    stealthLog(idConnection) = stealthLog(idConnection) & vbCrLf & TibiaTimestamp() & Replace(strSend, vbLf, vbCrLf)
    If idConnection = stealthIDselected Then
        frmStealth.UpdateValues
    End If
    SendCustomSystemMessageToClient = 0
    Exit Function
  End If
  
  If GameConnected(idConnection) = True And sentFirstPacket(idConnection) = True Then
  longSend = Len(strSend)
  ' totalL = num of bytes -1 (because array index start from 0)
  totalL = 5 + longSend
  ReDim cheatpacket(totalL)
  ' packet header = num of bytes -2 , that is totalL -1
  longP = totalL - 1
  hb = HighByteOfLong(CLng(longP))
  lb = LowByteOfLong(CLng(longP))
  cheatpacket(0) = lb
  cheatpacket(1) = hb

  cheatpacket(2) = &HB4 ' system message
  If TibiaVersionLong >= 872 Then
    If (thecolor - 1) < 0 Then
        cheatpacket(3) = 0 '  type and color
    Else
        cheatpacket(3) = thecolor - 1 '  type and color
    End If
  ElseIf TibiaVersionLong >= 861 Then
    If (thecolor - 3) < 0 Then
        cheatpacket(3) = 0 '  type and color
    Else
        cheatpacket(3) = thecolor - 3 '  type and color
    End If
  ElseIf TibiaVersionLong >= 840 Then
    cheatpacket(3) = thecolor + 3 '  type and color
  ElseIf TibiaVersionLong >= 820 Then
    cheatpacket(3) = thecolor + 2 '  type and color
  Else
    cheatpacket(3) = thecolor '  type and color
  End If
  pos = 4
  ' message - lenght
  hb = HighByteOfLong(CLng(longSend))
  lb = LowByteOfLong(CLng(longSend))
  cheatpacket(pos) = lb
  pos = pos + 1
  cheatpacket(pos) = hb
  pos = pos + 1
  ' message - text
  strCad = strSend
  For i = 1 To longSend
    chCad = Left(strCad, 1)
    strCad = Right(strCad, Len(strCad) - 1)
    cheatpacket(pos) = Asc(chCad)
    pos = pos + 1
  Next i
  ' send the packet
  
  frmMain.UnifiedSendToClientGame idConnection, cheatpacket
  End If
  SendCustomSystemMessageToClient = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at SendCustomSystemMessageToClient #"
  frmMain.DoCloseActions idConnection
  DoEvents
  SendCustomSystemMessageToClient = -1
End Function
Public Function SendLogSystemMessageToClient(idConnection As Integer, strSend As String) As Long
' &HB4 , &H14 -login
'        &H17 - temp
      ' system message (msgtype,lenght and message)
      'lonN = GetTheLong(packet(pos + 2), packet(pos + 3))
     ' pos = pos + 4 + lonN
  Dim cheatpacket() As Byte
  Dim longP As Long
  Dim longSend As Long
  Dim longFrom As Long
  Dim totalL As Long
  Dim hb As Byte
  Dim lb As Byte
  Dim pos As Integer
  Dim i As Integer
  Dim strCad As String
  Dim chCad As String
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  If (idConnection = 0) Then
    SendLogSystemMessageToClient = -1
    Exit Function
  End If
  
  If frmStealth.chkStealthMessages.value = 1 Then
    stealthLog(idConnection) = stealthLog(idConnection) & vbCrLf & TibiaTimestamp() & Replace(strSend, vbLf, vbCrLf)
    If idConnection = stealthIDselected Then
        frmStealth.UpdateValues
    End If
    SendLogSystemMessageToClient = 0
    Exit Function
  End If
  
  If GameConnected(idConnection) = True And sentFirstPacket(idConnection) = True Then
  longSend = Len(strSend)
  ' totalL = num of bytes -1 (because array index start from 0)
  totalL = 5 + longSend
  ReDim cheatpacket(totalL)
  ' packet header = num of bytes -2 , that is totalL -1
  longP = totalL - 1
  hb = HighByteOfLong(CLng(longP))
  lb = LowByteOfLong(CLng(longP))
  cheatpacket(0) = lb
  cheatpacket(1) = hb
  cheatpacket(2) = &HB4 ' system message
  If TibiaVersionLong >= 1036 Then
  cheatpacket(3) = &H11 ' log
  ElseIf TibiaVersionLong >= 872 Then
  cheatpacket(3) = &H10 ' log
  ElseIf TibiaVersionLong >= 861 Then
  cheatpacket(3) = &H11 ' log
  ElseIf TibiaVersionLong >= 840 Then
  cheatpacket(3) = &H17 ' log
  ElseIf TibiaVersionLong >= 820 Then
  cheatpacket(3) = &H16 ' log
  Else
  cheatpacket(3) = &H14 ' log
  End If
  pos = 4
  ' message - lenght
  hb = HighByteOfLong(CLng(longSend))
  lb = LowByteOfLong(CLng(longSend))
  cheatpacket(pos) = lb
  pos = pos + 1
  cheatpacket(pos) = hb
  pos = pos + 1
  ' message - text
  strCad = strSend
  For i = 1 To longSend
    chCad = Left(strCad, 1)
    strCad = Right(strCad, Len(strCad) - 1)
    cheatpacket(pos) = Asc(chCad)
    pos = pos + 1
  Next i
  ' send the packet
  
  frmMain.UnifiedSendToClientGame idConnection, cheatpacket
  End If
  SendLogSystemMessageToClient = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at SendLogSystemMessageToClient #"
  'frmMain.DoCloseActions idConnection
  DoEvents
  SendLogSystemMessageToClient = -1
End Function

Public Function TibiaTimestamp() As String
    Dim strRes As String
    strRes = Format(Time, "hh:mm") & " "
    TibiaTimestamp = strRes
End Function
Public Function GiveGMmessage(idConnection As Integer, strSend As String, strFrom As String) As Long
  Dim cheatpacket() As Byte
  Dim longP As Long
  Dim longSend As Long
  Dim longFrom As Long
  Dim totalL As Long
  Dim hb As Byte
  Dim lb As Byte
  Dim pos As Integer
  Dim i As Integer
  Dim strCad As String
  Dim chCad As String
  'If (1 = 0) Then
  '  GiveGMmessage = 0 ' NO GM MESSAGES ! XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
  'Exit Function ' NO GM MESSAGES ! XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
  'End If
  

  #If FinalMode Then
  On Error GoTo errclose
  #End If
  If (idConnection = 0) Then
    GiveGMmessage = -1
    Exit Function
  End If
  If frmStealth.chkStealthMessages.value = 1 Then
    stealthLog(idConnection) = stealthLog(idConnection) & vbCrLf & TibiaTimestamp() & "(**RED MSG!***) " & strFrom & " [" & CStr(fakemessagesLevel) & "]: " & Replace(strSend, vbLf, vbCrLf)
    If idConnection = stealthIDselected Then
        frmStealth.UpdateValues
    End If
    GiveGMmessage = 0
    Exit Function
  End If
  longSend = Len(strSend)
  longFrom = Len(strFrom)
  ' totalL = num of bytes -1 (because array index start from 0)
  totalL = 7 + longSend + longFrom
  If TibiaVersionLong > 760 Then
  totalL = totalL + 4
  End If
  If TibiaVersionLong >= 773 Then ' NEW STRANGE THING
    totalL = totalL + 2
  End If
  ReDim cheatpacket(totalL)
  ' packet header = num of bytes -2 , that is totalL -1
  longP = totalL - 1
  hb = HighByteOfLong(CLng(longP))
  lb = LowByteOfLong(CLng(longP))
  cheatpacket(0) = lb
  cheatpacket(1) = hb
  cheatpacket(2) = &HAA 'message
  If TibiaVersionLong > 760 Then
    cheatpacket(3) = &H0
    cheatpacket(4) = &H0
    cheatpacket(5) = &H0
    cheatpacket(6) = &H0
    pos = 7
  Else
    pos = 3
  End If
  ' from - lenght
  hb = HighByteOfLong(CLng(longFrom))
  lb = LowByteOfLong(CLng(longFrom))
  cheatpacket(pos) = lb
  cheatpacket(pos + 1) = hb
  pos = pos + 2
  ' from - text
  strCad = strFrom
  For i = 1 To longFrom
    chCad = Left(strCad, 1)
    strCad = Right(strCad, Len(strCad) - 1)
    cheatpacket(pos) = Asc(chCad)
    pos = pos + 1
  Next i
  If TibiaVersionLong >= 773 Then ' NEW STRANGE THING
    cheatpacket(pos) = fakemessagesLevel1
    pos = pos + 1
    cheatpacket(pos) = fakemessagesLevel2
    pos = pos + 1
  End If
  'cheatpacket(pos) = &H9
  ' ! tests 10.36

  cheatpacket(pos) = oldmessage_H9
  'cheatpacket(pos) = &HC
  pos = pos + 1
  ' message - lenght
  hb = HighByteOfLong(CLng(longSend))
  lb = LowByteOfLong(CLng(longSend))
  cheatpacket(pos) = lb
  pos = pos + 1
  cheatpacket(pos) = hb
  pos = pos + 1
  ' message - text
  strCad = strSend
  For i = 1 To longSend
    chCad = Left(strCad, 1)
    strCad = Right(strCad, Len(strCad) - 1)
    cheatpacket(pos) = Asc(chCad)
    pos = pos + 1
  Next i
  ' send the packet

  frmMain.UnifiedSendToClientGame idConnection, cheatpacket
  GiveGMmessage = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at GiveGMmessage #"
  frmMain.DoCloseActions idConnection
  DoEvents
  GiveGMmessage = -1
End Function

Public Sub LogStatusOnFile(file_name As String)
  Dim fn As Integer
  Dim a As Long
  Dim writeThis As String
  On Error GoTo ignoreit
  a = 0
  fn = FreeFile
  Open App.Path & "\" & file_name For Append As #fn
    writeThis = vbCrLf & "ADITIONAL DETAILS:" & vbCrLf
    Print #fn, writeThis
    writeThis = "TibiaVersionLong=" & CStr(TibiaVersionLong)
    Print #fn, writeThis
    writeThis = "TibiaVersion=" & TibiaVersion
    Print #fn, writeThis
    writeThis = "MAXCLIENTS=" & CStr(MAXCLIENTS)
    Print #fn, writeThis
    writeThis = "highestDatTile=" & CStr(highestDatTile)
    Print #fn, writeThis
    writeThis = "Usecrackd=" & BooleanAsStr(UseCrackd)
    Print #fn, writeThis
    writeThis = "Option1=" & BooleanAsStr(frmMain.TrueServer1.value)
    Print #fn, writeThis
    writeThis = "Option2=" & BooleanAsStr(frmMain.TrueServer2.value)
    Print #fn, writeThis
    writeThis = "Option3=" & BooleanAsStr(frmMain.TrueServer3.value)
    Print #fn, writeThis
  Close #fn
  Exit Sub
ignoreit:
  a = -1
End Sub
Public Function GetMyAppDataFolder() As String
    On Error GoTo goterr
    Dim base As String
    Dim fullPath As String
    Dim fs As Scripting.FileSystemObject
    base = GetAppDataFolder()
    fullPath = base & "\Blackd Proxy"
    Set fs = New Scripting.FileSystemObject
    If fs.FolderExists(fullPath) = False Then
       fs.CreateFolder (fullPath)
    End If
    GetMyAppDataFolder = fullPath
    Exit Function
goterr:
    GetMyAppDataFolder = App.Path
End Function
Public Sub LogOnFile(file_name As String, strText As String)
  Dim fn As Integer
  Dim errheader As String
  Dim writeThis As String
  Dim a As Long
  On Error GoTo ignoreit
  a = 0
  fn = FreeFile
  If file_name = "errors.txt" Then
    openErrorsTXTfolder
    errheader = "[" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & " using version " & ProxyVersion & " , with config.ini v" & CStr(TibiaVersionLong) & " ] "
    If TibiaVersionLong < highestTibiaVersionLong Then
        If frmMain.TrueServer3.value = True Then
           errheader = errheader & "[Playing OTserver>> " & frmMain.ForwardGameTo.Text & ":" & frmMain.txtServerLoginP.Text & "] "
        Else
           errheader = errheader & "[Trying to play real server with old config] "
        End If
    Else
       errheader = errheader & "[Playing real servers] "
    End If
    writeThis = errheader & strText
    frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & writeThis
  Else
    writeThis = strText
  End If
  If (file_name = "errors.txt") Then
    Dim mySafeSolder As String
    mySafeSolder = GetMyAppDataFolder()
    Open mySafeSolder & "\" & file_name For Append As #fn
  Else
    If Len(file_name) > 4 Then
      If Left$(file_name, 4) = "log_" Then
        Open App.Path & "\mylogs\" & file_name For Append As #fn
      Else
        Open App.Path & "\" & file_name For Append As #fn
      End If
    Else
    Open App.Path & "\" & file_name For Append As #fn
  End If
  End If
    Print #fn, writeThis
  Close #fn
  If file_name = "errors.txt" Then
    If thisShouldNotBeLoading = 1 Then
      'custom ng
      'frmMenu.Caption = "ERROR - Check errors.txt for details"
    End If
  End If
  Exit Sub
ignoreit:
  a = -1
End Sub

Public Sub OverwriteOnFile(file_name As String, strText As String)
  Dim fn As Integer
  Dim errheader As String
  Dim writeThis As String
  Dim a As Long
  On Error GoTo ignoreit
  a = 0
  fn = FreeFile
  If file_name = "errors.txt" Then
    errheader = "[" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & " using version " & ProxyVersion & " , with config.ini v" & CStr(TibiaVersionLong) & " ] "
    writeThis = errheader & strText
  Else
    writeThis = strText
  End If
  If (file_name = "errors.txt") Then
    Dim mySafeSolder As String
    mySafeSolder = GetMyAppDataFolder()
    Open mySafeSolder & "\" & file_name For Append As #fn
  Else
  Open App.Path & "\" & file_name For Output As #fn
  End If
    Print #fn, writeThis
  Close #fn
  If file_name = "errors.txt" Then
    LogStatusOnFile "errors.txt"
    'custom ng
    'frmMenu.Caption = "ERROR - Check errors.txt for details"
  End If
  Exit Sub
ignoreit:
  a = -1
End Sub

Public Function fromBooleanToStr(b As Boolean) As String
  Dim res As String
  If b = True Then
    res = "TRUE"
  Else
    res = "FALSE"
  End If
  fromBooleanToStr = res
End Function

Public Function MyHexPosition(idConnection As Integer) As String
  Dim res As String
  Dim b1 As Byte
  Dim b2 As Byte
  res = ""
  b1 = LowByteOfLong(myX(idConnection))
  b2 = HighByteOfLong(myX(idConnection))
  res = GoodHex(b1) & " " & GoodHex(b2)
  b1 = LowByteOfLong(myY(idConnection))
  b2 = HighByteOfLong(myY(idConnection))
  res = res & " " & GoodHex(b1) & " " & GoodHex(b2)
  b1 = CLng(myZ(idConnection))
  res = res & " " & GoodHex(b1)
  MyHexPosition = res
End Function
Public Function GetHexStrFromPosition(x As Long, y As Long, z As Long) As String
  Dim res As String
  Dim b1 As Byte
  Dim b2 As Byte
  res = ""
  b1 = LowByteOfLong(x)
  b2 = HighByteOfLong(x)
  res = GoodHex(b1) & " " & GoodHex(b2)
  b1 = LowByteOfLong(y)
  b2 = HighByteOfLong(y)
  res = res & " " & GoodHex(b1) & " " & GoodHex(b2)
  b1 = CLng(z)
  res = res & " " & GoodHex(b1)
  GetHexStrFromPosition = res
End Function

Public Function MyStackPos(idConnection As Integer) As Byte
  'return current stackpos for given idconnection
  Dim i As Long
  Dim res As Byte
  res = &HFF
  For i = 1 To 10
    If Matrix(0, 0, myZ(idConnection), idConnection).s(i).dblID = myID(idConnection) Then
      res = CByte(i)
      Exit For
    End If
  Next i
  MyStackPos = res
End Function

Public Function FirstPersonStackPos(idConnection As Integer) As Byte
  ' return current stackpos of first person in central square
  ' (to cast UH there)
  Dim i As Long
  Dim res As Byte
  res = &HFF
  For i = 0 To 10
    If Matrix(0, 0, myZ(idConnection), idConnection).s(i).dblID <> 0 Then
      res = CByte(i)
      Exit For
    End If
  Next i
  FirstPersonStackPos = res
End Function


Public Function UseIH(idConnection As Integer) As Long
  Dim fRes As TypeSearchItemResult2
  Dim cPacket() As Byte
  Dim myS As Byte
  Dim aRes As Long
  Dim res As Long
  Dim inRes As Integer
  Dim sCheat As String
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  res = 0
  GoTo forcefordebug
  If myMagLevel(idConnection) < 1 Then 'can't use IH
       If (frmHardcoreCheats.chkRuneAlarm.value = 1) Then
         If PlayTheDangerSound = False Then
           ChangePlayTheDangerSound True
           aRes = GiveGMmessage(idConnection, "Warning: You are low hp!", "BlackdProxy")
           DoEvents
         End If
       End If
    UseIH = -1
    Exit Function
  End If
forcefordebug:
   fRes = SearchItem(idConnection, LowByteOfLong(tileID_IH), HighByteOfLong(tileID_IH))  'search IH
   
If (frmHardcoreCheats.chkTotalWaste.value = True) Then 'And (TibiaVersionLong >= 773)) Then
  GoTo justdoit
End If
 
    
          If fRes.foundCount > 0 Then
            myS = FirstPersonStackPos(idConnection)
            If myS < &HFF Then
              res = 0
              aRes = SendSystemMessageToClient(idConnection, CStr(fRes.foundCount) & " IHs found - Casting one from bp ID " & _
               CStr(CLng(fRes.bpID)) & " slot " & CStr(CLng(fRes.slotID)) & " to stackpos " & GoodHex(myS))
               
              ' 11 00 83 FF FF ...
              SafeCastCheatString "UseIH1", idConnection, "83 FF FF " & GoodHex(&H40 + fRes.bpID) & " 00 " & _
                  GoodHex(fRes.slotID) & " " & FiveChrLon(tileID_IH) & " " & GoodHex(fRes.slotID) & " " & MyHexPosition(idConnection) & " 63 00 " & GoodHex(myS)

            Else
              res = -1
              If PlayTheDangerSound = False Then
                ChangePlayTheDangerSound True
                aRes = GiveGMmessage(idConnection, "Unable to cast IHs here. Try moving or reloging!", "BlackdProxy")
                DoEvents
                res = -1
                GoTo lastcheck
              End If
            End If
          Else ' NEW
justdoit:
      If ((frmHardcoreCheats.chkEnhancedCheats.value = True) Or (frmHardcoreCheats.chkTotalWaste.value = True)) Then 'And (TibiaVersionLong >= 773)) Then
              ' 0d 00 84 ...
              sCheat = "0D 00 84 FF FF 00 00 00 " & GoodHex(LowByteOfLong(tileID_IH)) & _
               " " & GoodHex(HighByteOfLong(tileID_IH)) & " 00 " & _
               SpaceID(myID(idConnection))
              SafeCastCheatString "UseIH2", idConnection, sCheat
                res = 0
                GoTo lastcheck
            Else
                GoTo lastcheck
            End If
          End If

lastcheck:
    If (frmHardcoreCheats.chkRuneAlarm.value = 1) And (CInt(frmHardcoreCheats.txtAlarmUHs.Text) > fRes.foundCount) Then
      If PlayTheDangerSound = False Then
        ChangePlayTheDangerSound True
        If fRes.foundCount = 0 Then
            aRes = GiveGMmessage(idConnection, "Can't find IHs, open new bp of IHs!", "BlackdProxy")
        Else
            aRes = GiveGMmessage(idConnection, "Warning: You are low of IHs!", "BlackdProxy")
        End If
        DoEvents
        UseIH = -1
        Exit Function
      End If

    Else
        UseIH = res
    End If
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at UseIH #"
  frmMain.DoCloseActions idConnection
  DoEvents
  UseIH = -1
End Function

Public Function UseFastIH(idConnection As Integer) As Long
  Dim fRes As TypeSearchItemResult2
  Dim cPacket() As Byte
  Dim myS As Byte
  Dim aRes As Long
  Dim inRes As Integer
  Dim sCheat As String
  Dim SpecialSource As Boolean
  Dim res As Long
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  res = 0
  If myMagLevel(idConnection) < 1 Then 'can't use IH
    UseFastIH = -1
    Exit Function
  End If
  
      If ((frmHardcoreCheats.chkEnhancedCheats.value = True) Or (frmHardcoreCheats.chkTotalWaste.value = True)) Then 'And (TibiaVersionLong >= 773)) Then
    SpecialSource = True
   Else
    SpecialSource = False
   End If
   fRes = SearchFirstItem(idConnection, LowByteOfLong(tileID_IH), HighByteOfLong(tileID_IH))  'search IH
     
   If SpecialSource = False Then

          If fRes.foundCount > 0 Then
            myS = FirstPersonStackPos(idConnection)
            If myS < &HFF Then
              'aRes = SendSystemMessageToClient(idConnection, CStr(fRes.foundCount) & " IHs found - Casting one from bp ID " & _
               CStr(CLng(fRes.bpID)) & " slot " & CStr(CLng(fRes.slotID)) & " to stackpos " & GoodHex(myS))
              sCheat = "83 FF FF " & GoodHex(&H40 + fRes.bpID) & " 00 " & _
               GoodHex(fRes.slotID) & " " & FiveChrLon(tileID_IH) & " " & GoodHex(fRes.slotID) & " " & MyHexPosition(idConnection) & " 63 00 " & GoodHex(myS)
              SafeCastCheatString "UseFastIH1", idConnection, sCheat
            Else
              If PlayTheDangerSound = False Then
                ChangePlayTheDangerSound True
                aRes = GiveGMmessage(idConnection, "Unable to cast IHs here. Try moving or reloging!", "BlackdProxy")
                DoEvents
                res = -1
                GoTo lastcheck
              End If
            End If
          Else
             GoTo lastcheck
          End If
        Else ' NEW
          sCheat = "84 FF FF 00 00 00 " & GoodHex(LowByteOfLong(tileID_IH)) & _
           " " & GoodHex(HighByteOfLong(tileID_IH)) & " 00 " & _
           SpaceID(myID(idConnection))
          SafeCastCheatString "UseFastIH2", idConnection, sCheat
        End If
  res = 0

lastcheck:
    If (frmHardcoreCheats.chkRuneAlarm.value = 1) And (CInt(frmHardcoreCheats.txtAlarmUHs.Text) > fRes.foundCount) Then
      If PlayTheDangerSound = False Then
        ChangePlayTheDangerSound True
        If fRes.foundCount = 0 Then
            aRes = GiveGMmessage(idConnection, "Can't find IHs, open new bp of IHs!", "BlackdProxy")
        Else
            aRes = GiveGMmessage(idConnection, "Warning: You are low of IHs!", "BlackdProxy")
        End If
        DoEvents
        UseFastIH = -1
        Exit Function
      End If

    Else
        UseFastIH = res
    End If
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at UseFastIH #"
  frmMain.DoCloseActions idConnection
  DoEvents
  UseFastIH = -1
End Function

Public Function UseUH(idConnection As Integer) As Long
  Dim fRes As TypeSearchItemResult2
  Dim cPacket() As Byte
  Dim myS As Byte
  Dim aRes As Long
  Dim inRes As Integer
  Dim sCheat As String
  Dim res As Long
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  res = 0
  If myMagLevel(idConnection) < 4 Then 'can't use UH, try IH
    UseUH = UseIH(idConnection)
    Exit Function
  End If
fRes = SearchItem(idConnection, LowByteOfLong(tileID_UH), HighByteOfLong(tileID_UH))  'search UH
If (frmHardcoreCheats.chkTotalWaste.value = True) Then ' And (TibiaVersionLong >= 773)) Then
  GoTo justdoit
End If

       
          If fRes.foundCount > 0 Then
            myS = FirstPersonStackPos(idConnection)
            If myS < &HFF Then
              res = 0
              aRes = SendSystemMessageToClient(idConnection, CStr(fRes.foundCount) & " UHs found - Casting one from bp ID " & _
               CStr(CLng(fRes.bpID)) & " slot " & CStr(CLng(fRes.slotID)) & " to stackpos " & GoodHex(myS))
              sCheat = "83 FF FF " & GoodHex(&H40 + fRes.bpID) & " 00 " & _
               GoodHex(fRes.slotID) & " " & FiveChrLon(tileID_UH) & " " & GoodHex(fRes.slotID) & " " & MyHexPosition(idConnection) & " 63 00 " & GoodHex(myS)
             '  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "SENT: " & scheat
              SafeCastCheatString "UseUH1", idConnection, sCheat
            Else ' NEW
              res = -1
              If PlayTheDangerSound = False Then
                ChangePlayTheDangerSound True
                aRes = GiveGMmessage(idConnection, "Unable to cast UHs here. Try moving or reloging!", "BlackdProxy")
                DoEvents
                res = -1
                GoTo lastcheck
              End If
            End If
          Else
justdoit:
            If ((frmHardcoreCheats.chkEnhancedCheats.value = True) Or (frmHardcoreCheats.chkTotalWaste.value = True)) Then ' And (TibiaVersionLong >= 773)) Then
               sCheat = "84 FF FF 00 00 00 " & GoodHex(LowByteOfLong(tileID_UH)) & _
                " " & GoodHex(HighByteOfLong(tileID_UH)) & " 00 " & _
                SpaceID(myID(idConnection))
                SafeCastCheatString "UseUH2", idConnection, sCheat
                res = 0
                GoTo lastcheck
            Else
                GoTo lastcheck
            End If
          End If

lastcheck:
    If (frmHardcoreCheats.chkRuneAlarm.value = 1) And (CInt(frmHardcoreCheats.txtAlarmUHs.Text) > fRes.foundCount) Then
      If PlayTheDangerSound = False Then
        ChangePlayTheDangerSound True
        If fRes.foundCount = 0 Then
            aRes = GiveGMmessage(idConnection, "Can't find UHs, open new bp of UHs!", "BlackdProxy")
        Else
            aRes = GiveGMmessage(idConnection, "Warning: You are low of UHs!", "BlackdProxy")
        End If
        DoEvents
        UseUH = -1
        Exit Function
      End If

    Else
        UseUH = res
    End If
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at UseUH #"
  frmMain.DoCloseActions idConnection
  DoEvents
  UseUH = -1
End Function
Public Function UseFastUH(idConnection As Integer) As Long
  Dim fRes As TypeSearchItemResult2
  Dim cPacket() As Byte
  Dim myS As Byte
  Dim aRes As Long
  Dim inRes As Integer
  Dim sCheat As String
  Dim SpecialSource As Boolean
  Dim res As Long
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  res = 0
  If myMagLevel(idConnection) < 4 Then 'can't use UH, try IH
    UseFastUH = UseFastIH(idConnection)
    Exit Function
  End If
      If ((frmHardcoreCheats.chkEnhancedCheats.value = True) Or (frmHardcoreCheats.chkTotalWaste.value = True)) Then ' And (TibiaVersionLong >= 773)) Then
    SpecialSource = True
   Else
    SpecialSource = False
   End If
   
   fRes = SearchFirstItem(idConnection, LowByteOfLong(tileID_UH), HighByteOfLong(tileID_UH))  'search UH
   
   If SpecialSource = False Then
       
          If fRes.foundCount > 0 Then
            myS = FirstPersonStackPos(idConnection)
            If myS < &HFF Then
              'aRes = SendSystemMessageToClient(idConnection, CStr(fRes.foundCount) & " UHs found - Casting one from bp ID " & _
               CStr(CLng(fRes.bpID)) & " slot " & CStr(CLng(fRes.slotID)) & " to stackpos " & GoodHex(myS))
              sCheat = "83 FF FF " & GoodHex(&H40 + fRes.bpID) & " 00 " & _
               GoodHex(fRes.slotID) & " " & FiveChrLon(tileID_UH) & " " & GoodHex(fRes.slotID) & " " & MyHexPosition(idConnection) & " 63 00 " & GoodHex(myS)
              SafeCastCheatString "UseFastUH1", idConnection, sCheat
              
            Else
              If PlayTheDangerSound = False Then
                ChangePlayTheDangerSound True
                aRes = GiveGMmessage(idConnection, "Unable to cast UHs here. Try moving or reloging!", "BlackdProxy")
                DoEvents
                res = -1
                GoTo lastcheck
              End If
            End If
          Else
            GoTo lastcheck
          End If
Else ' NEW
          sCheat = "84 FF FF 00 00 00 " & GoodHex(LowByteOfLong(tileID_UH)) & _
           " " & GoodHex(HighByteOfLong(tileID_UH)) & " 00 " & _
           SpaceID(myID(idConnection))
          SafeCastCheatString "UseFastUH2", idConnection, sCheat
End If
  res = 0

lastcheck:
    If (frmHardcoreCheats.chkRuneAlarm.value = 1) And (CInt(frmHardcoreCheats.txtAlarmUHs.Text) > fRes.foundCount) Then
      If PlayTheDangerSound = False Then
        ChangePlayTheDangerSound True
        If fRes.foundCount = 0 Then
            aRes = GiveGMmessage(idConnection, "Can't find UHs, open new bp of UHs!", "BlackdProxy")
        Else
            aRes = GiveGMmessage(idConnection, "Warning: You are low of UHs!", "BlackdProxy")
        End If
        DoEvents
        UseFastUH = -1
        Exit Function
      End If

    Else
        UseFastUH = res
    End If
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at UseFastUH #"
  frmMain.DoCloseActions idConnection
  DoEvents
  UseFastUH = -1
End Function
Public Function CatchFish(idConnection As Integer) As Long
  ' nothing
  ' 5D 0D = fishing rod
  Dim baseID As Long
  Dim addID As Long
  Dim fishCount As Long
  Dim fishThis As Long
  Dim fishX As Long
  Dim fishY As Long
  Dim fishz As Long
  Dim x As Long
  Dim y As Long
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim tileSTR As String
  Dim aRes As Long
  Dim inRes As Integer
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  ' en la version 772 se podia ya usar la caña desde cualquier lado???
If ((frmHardcoreCheats.chkEnhancedCheats.value = True) Or (frmHardcoreCheats.chkTotalWaste.value = True)) Then 'Or (TibiaVersionLong >= 773)) Then
GoTo justdoit
End If
  If mySlot(idConnection, SLOT_AMMUNITION).t1 = LowByteOfLong(tileID_FishingRod) And _
   mySlot(idConnection, SLOT_AMMUNITION).t2 = HighByteOfLong(tileID_FishingRod) Then
justdoit:
    fishCount = 0
    For x = -7 To 7
      For y = -5 To 5
        baseID = GetTheLong(Matrix(y, x, myZ(idConnection), idConnection).s(0).t1, _
         Matrix(y, x, myZ(idConnection), idConnection).s(0).t2)
        addID = GetTheLong(Matrix(y, x, myZ(idConnection), idConnection).s(1).t1, _
         Matrix(y, x, myZ(idConnection), idConnection).s(1).t2)
        If DatTiles(baseID).haveFish = True And addID = 0 Then
          fishCount = fishCount + 1
        End If
      Next y
    Next x
    If fishCount = 0 Then
      aRes = SendSystemMessageToClient(idConnection, "No fish left in in your screen!!")
    Else
      
      fishThis = CLng(Int((fishCount * Rnd) + 1))   ' random 1-fishcount
      fishCount = 0
      fishX = myX(idConnection)
      fishY = myY(idConnection)
      fishz = myZ(idConnection)
      tileSTR = ""
      For x = -7 To 7
        For y = -5 To 5
          baseID = GetTheLong(Matrix(y, x, myZ(idConnection), idConnection).s(0).t1, _
           Matrix(y, x, myZ(idConnection), idConnection).s(0).t2)
          addID = GetTheLong(Matrix(y, x, myZ(idConnection), idConnection).s(1).t1, _
           Matrix(y, x, myZ(idConnection), idConnection).s(1).t2)
          If DatTiles(baseID).haveFish = True And addID = 0 Then
            fishCount = fishCount + 1
            If fishCount = fishThis Then
              tileSTR = GoodHex(Matrix(y, x, myZ(idConnection), idConnection).s(0).t1) & _
               " " & GoodHex(Matrix(y, x, myZ(idConnection), idConnection).s(0).t2)
              fishX = fishX + x
              fishY = fishY + y
            End If
          End If
        Next y
      Next x
      ' cast fishing rod
      ' 11 00 83 FF FF 0A 00 00 5D 0D 00 39 7D BD 7D 07 59 02 00
      aRes = SendSystemMessageToClient(idConnection, CStr(fishCount) & " fish left in your screen. Fishing number " & CStr(fishThis))
      DoEvents
If ((frmHardcoreCheats.chkEnhancedCheats.value = True) Or (frmHardcoreCheats.chkTotalWaste.value = True)) Then 'And (TibiaVersionLong >= 773)) Then
      sCheat = "11 00 83 FF FF 00 00 00 " & FiveChrLon(tileID_FishingRod) & " 00 " & GetHexStrFromPosition(fishX, fishY, fishz) & " " & tileSTR & " 00"
 
Else
      sCheat = "11 00 83 FF FF 0A 00 00 " & FiveChrLon(tileID_FishingRod) & " 00 " & GetHexStrFromPosition(fishX, fishY, fishz) & " " & tileSTR & " 00"
   End If
     ' SendSystemMessageToClient idConnection, sCheat
      inRes = GetCheatPacket(cPacket, sCheat)
      frmMain.UnifiedSendToServerGame idConnection, cPacket, True
    End If
  Else
    aRes = SendLogSystemMessageToClient(idConnection, "First equip the fishing rod in your ammo position. ")
  End If
  CatchFish = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at CatchFish #"
  frmMain.DoCloseActions idConnection
  DoEvents
  CatchFish = -1
End Function

Public Function FiveChrLon(num As Long) As String
  FiveChrLon = GoodHex(LowByteOfLong(num)) & " " & GoodHex(HighByteOfLong(num))
End Function

Public Function GetTheByteFromTwoChr(str As String) As Byte
  On Error GoTo goterr
  Dim b1 As Byte
  Dim b2 As Byte
  Dim res As Long
  res = -1
  If Len(str) > 1 Then
    b1 = FromHexToDec(Mid(str, 1, 1))
    b2 = FromHexToDec(Mid(str, 2, 1))
    res = (CLng(b2)) + (CLng(b1) * 16)
  End If
  GetTheByteFromTwoChr = CByte(res)
  Exit Function
goterr:
  GetTheByteFromTwoChr = 0
End Function
Public Function GetTheLongFromFiveChr(str As String) As Long
  On Error GoTo goterr
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim b4 As Byte
  Dim res As Long
  res = -1
  
  If Len(str) > 4 Then
    b1 = FromHexToDec(Mid(str, 1, 1))
    b2 = FromHexToDec(Mid(str, 2, 1))
    b3 = FromHexToDec(Mid(str, 4, 1))
    b4 = FromHexToDec(Mid(str, 5, 1))
    res = (CLng(b2)) + (CLng(b1) * 16) + (CLng(b4) * 256) + (CLng(b3) * 4096)
  End If
  GetTheLongFromFiveChr = res
  Exit Function
goterr:
  GetTheLongFromFiveChr = -1 'new in 8.21 +
End Function

Public Function StringToHexString(str As String) As String
  Dim res As String
  Dim lonS As Long
  Dim cPart As String
  Dim bPart As Byte
  Dim i As Long
  res = ""
  lonS = Len(str)
  If lonS > 0 Then
    cPart = Mid(str, 1, 1)
    bPart = AscB(cPart)
    res = GoodHex(bPart)
  End If
  For i = 2 To lonS
     cPart = Mid(str, i, 1)
    bPart = AscB(cPart)
    res = res & " " & GoodHex(bPart)
  Next i
  StringToHexString = res
End Function

Public Function RevealAll(idConnection As Integer) As Long
  Dim x As Long
  Dim y As Long
  Dim s As Long
  Dim z As Long
  Dim startz As Long
  Dim endz As Long
  Dim totalLen As Long
  Dim sendStr As String
  Dim currStr As String
  Dim subF As Long
  Dim addS As String
  Dim showx As Long
  Dim showy As Long
  Dim cPacket() As Byte
  Dim mobName As String
  Dim aRes As Long
  Dim mobID As Double
  Dim inRes As Integer
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  totalLen = 0
  ' work with a local z
  z = myZ(idConnection)
  ' two cases: you are underground (z>7) or not
  If (z > 7) Then
    startz = z - 2
    endz = MinV(15, z + 2) ' there is a special case on the most deep. This deal with that
  Else
    startz = 0
    endz = 7
  End If
  For z = startz To endz
    For y = -6 To 7
      If y < -5 Then
        showy = -5
      ElseIf y > 5 Then
        showy = 5
      Else
        showy = y
      End If
      For x = -8 To 9
        If x < -7 Then
          showx = -7
        ElseIf x > 7 Then
          showx = 7
        Else
          showx = x
        End If
        For s = 0 To 10
          mobID = Matrix(y, x, z, idConnection).s(s).dblID
          If mobID > 0 Then
            mobName = GetNameFromID(idConnection, mobID)
            subF = z - myZ(idConnection)
            If subF > 0 Then
              mobName = "+" & CStr(z - myZ(idConnection)) & " " & mobName
            ElseIf subF < 0 Then
              mobName = CStr(z - myZ(idConnection)) & " " & mobName
            End If
            If Len(mobName) > 9 Then
              mobName = Left(mobName, 9)
            End If
            ' 84 3C 7D AD 7D 08 D7 06 00 57 65 68 65 79 21
             currStr = "84 " & GetHexStrFromPosition(myX(idConnection) + showx, myY(idConnection) + showy, myZ(idConnection)) & _
              " D7 " & FiveChrLon(Len(mobName)) & " " & StringToHexString(mobName) & " "
             sendStr = sendStr & currStr
             totalLen = totalLen + 9 + Len(mobName)
          End If
        Next s
      Next x
    Next y
  Next z
  If totalLen = 0 Then
    aRes = SendSystemMessageToClient(idConnection, "Nobody on track")
  Else

    addS = FiveChrLon(totalLen)
    sendStr = addS & " " & sendStr
   ' aRes = SendMessageToClient(idConnection, "Revealing all : " & sendStr, "TEST")
    inRes = GetCheatPacket(cPacket, sendStr)
    frmMain.UnifiedSendToClientGame idConnection, cPacket
    'aRes = SendMessageToClient(idConnection, "SIZE: " & CStr(UBound(cPacket)), "TEST")
  End If
  RevealAll = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at RevealAll #"
  frmMain.DoCloseActions idConnection
  DoEvents
  RevealAll = -1
End Function

Public Function RevealAll2(idConnection As Integer) As Long
  Dim x As Long
  Dim y As Long
  Dim s As Long
  Dim z As Long
  Dim startz As Long
  Dim endz As Long
  Dim totalLen As Long
  Dim sendStr As String
  Dim currStr As String
  Dim subF As Long
  Dim addS As String
  Dim showx As Long
  Dim showy As Long
  Dim cPacket() As Byte
  Dim mobName As String
  Dim aRes As Long
  Dim mobID As Double
  Dim inRes As Integer
  Dim collectedTotal As Long
  Dim realTotal As Long
  Dim location As String
  Dim lenLocation As String
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  collectedTotal = 0
  totalLen = 0
  ' work with a local z
  aRes = SendCustomSystemMessageToClient(idConnection, "--- exiva all results: ---", &HB)
  DoEvents
  z = myZ(idConnection)
  ' two cases: you are underground (z>7) or not
  If (z > 7) Then
    startz = z - 2
    endz = MinV(15, z + 2) ' there is a special case on the most deep. This deal with that
  Else
    startz = 0
    endz = 7
  End If
  For z = startz To endz
    For y = -6 To 7
      If y < -5 Then
        showy = -5
      ElseIf y > 5 Then
        showy = 5
      Else
        showy = y
      End If
      For x = -8 To 9
        If x < -7 Then
          showx = -7
        ElseIf x > 7 Then
          showx = 7
        Else
          showx = x
        End If
        For s = 0 To 10
          mobID = Matrix(y, x, z, idConnection).s(s).dblID
          If mobID > 0 Then
            mobName = GetNameFromID(idConnection, mobID) & " [" & GetHPFromID(idConnection, mobID) & "% hp]"
            
            If x < 0 And y < 0 Then
              location = "to the north-west"
            ElseIf x < 0 And y > 0 Then
               location = "to the south-west."
            ElseIf x > 0 And y < 0 Then
              location = "to the north-east."
            ElseIf x > 0 And y > 0 Then
              location = "to the south-east."
            ElseIf x = 0 And y < 0 Then
              location = "to the north."
            ElseIf x = 0 And y > 0 Then
              location = "to the south."
            ElseIf x < 0 And y = 0 Then
              location = "to the west"
            ElseIf x > 0 And y = 0 Then
              location = "to the east"
            Else
              location = "standing next to you."
            End If
            subF = z - myZ(idConnection)
            If subF > 0 Then
              If location = "standing next to you." Then
                location = mobName & " is below you. (+" & CStr(z - myZ(idConnection)) & ")"
              Else
                location = mobName & " is on a lower level (+" & CStr(z - myZ(idConnection)) & ") " & location
              End If
              mobName = "(+" & CStr(z - myZ(idConnection)) & ") " & mobName
            ElseIf subF < 0 Then
              If location = "standing next to you." Then
                location = mobName & " is above you. (" & CStr(z - myZ(idConnection)) & ")"
              Else
                location = mobName & " is on a higher level (" & CStr(z - myZ(idConnection)) & ") " & location
              End If
              mobName = "(" & CStr(z - myZ(idConnection)) & ") " & mobName
            Else
              
              
              location = mobName & " is " & location
            End If
            collectedTotal = collectedTotal + 1
            realTotal = realTotal + 1
            'If Len(mobname) > 9 Then
            '  mobname = Left(mobname, 9)
            'End If
            aRes = SendCustomSystemMessageToClient(idConnection, location, &HB)
            DoEvents
            ' GAMESERVER1<( hex ) 1C 00 AA 00 00 00 00 07 00 61 20 73 68 65 65 70 00 00 10 3C 7D CF 7D 07 04 00 4D 61 65 68
             currStr = "AA 00 00 00 00 01 00 2E 00 00 " & GoodHex(oldmessage_H10) & " " & GetHexStrFromPosition(myX(idConnection) + showx, myY(idConnection) + showy, myZ(idConnection)) & " " & _
                FiveChrLon(Len(mobName)) & " " & StringToHexString(mobName)
             If sendStr = "" Then
               sendStr = currStr
             Else
               sendStr = sendStr & " " & currStr
             End If
             totalLen = totalLen + 18 + Len(mobName)
             If collectedTotal = 10 Then
                addS = FiveChrLon(totalLen)
                sendStr = addS & " " & sendStr
                'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "TEST>" & sendStr
                inRes = GetCheatPacket(cPacket, sendStr)
                frmMain.UnifiedSendToClientGame idConnection, cPacket
                totalLen = 0
                collectedTotal = 0
                DoEvents
             End If
          End If
        Next s
      Next x
    Next y
  Next z
  If totalLen = 0 Then
    ' do nothing
  Else

                addS = FiveChrLon(totalLen)
                sendStr = addS & " " & sendStr
                'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "TEST>" & sendStr
                inRes = GetCheatPacket(cPacket, sendStr)
                frmMain.UnifiedSendToClientGame idConnection, cPacket
                totalLen = 0
                DoEvents

  End If
  If realTotal > 10 Then
    aRes = SendCustomSystemMessageToClient(idConnection, "--- Found " & CStr(realTotal) & " on track! CAN'T DISPLAY ALL ON SCREEN ---", &HB)
    DoEvents
  Else
    aRes = SendCustomSystemMessageToClient(idConnection, "--- Found " & CStr(realTotal) & " on track. ---", &HB)
    DoEvents
  End If
  RevealAll2 = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at RevealAll #"
  frmMain.DoCloseActions idConnection
  DoEvents
  RevealAll2 = -1
End Function


Public Function EvalClientMessage(ByVal idConnection As Integer, ByRef packet() As Byte, pos As Long) As Long
  Dim res As Long
  Dim mtype As Byte
  Dim lonPerson As Long
  Dim lonM As Long
  Dim msg As String
  Dim lMsg As String
  Dim i As Long
  Dim keyChar As String
  Dim keyChar2
  Dim rightpart As String
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim b4 As Byte
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim aRes As Long
  Dim mcid As Integer
  Dim tempID As Long
  Dim lonO As Long
  Dim inRes As Integer
  Dim tmpStr As String
  Dim channelB1 As Byte
  Dim channelB2 As Byte
  Dim tmpmsg2 As String
  Dim upacket As Long
  Dim mtypeComp As Byte
  Dim tmpID As Double
  Dim aRes2 As Long
  Dim typeSay As Byte
  Dim typeTell As Byte
  Dim typeChannel As Byte
  Dim typeTell2 As Byte
  Dim typeChannel2 As Byte
  Dim typeReport As Byte
  Dim typeNPCTell As Byte
  Dim finalModeVar As Boolean
  Dim continueRest As Boolean
  finalModeVar = False
  #If FinalMode Then
  finalModeVar = True
  #End If
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  If TibiaVersionLong < 872 Then
    typeSay = &H1
    typeTell = &H4
    typeChannel = &H5
    typeTell2 = &H6
    typeChannel2 = &H7
    typeReport = &H8
    typeNPCTell = &HFF
  ElseIf TibiaVersionLong < 1036 Then
    typeSay = &H1
    typeTell = &H5
    typeChannel = &H7
    typeTell2 = &HFF
    typeChannel2 = &HFF
    typeReport = &H8
    typeNPCTell = &HB
  Else
    typeSay = &H1
    typeTell = &H5
    typeChannel = &H7
    typeTell2 = &HFF
    typeChannel2 = &HFF
    typeReport = &H8
    typeNPCTell = &HC 'SAYINTRADE
  End If
  res = 0
  mtype = packet(pos + 1)
  pos = pos + 2
  msg = ""
  '  0A 00 96 05 00 00 04 00 2D 67 65 74
  
  ' 44 65 6C 6C 20 4E 61 79 61 6D
  Select Case mtype
    Case typeSay ' say
      lonM = GetTheLong(packet(pos), packet(pos + 1))
      pos = pos + 2
      For i = 1 To lonM
        msg = msg & Chr(packet(pos))
        pos = pos + 1
      Next i
    Case typeNPCTell ' npc tell
      lonM = GetTheLong(packet(pos), packet(pos + 1))
      pos = pos + 2
      For i = 1 To lonM
        msg = msg & Chr(packet(pos))
        pos = pos + 1
      Next i
     ' Debug.Print msg
    Case typeTell ' tell
      lonPerson = GetTheLong(packet(pos), packet(pos + 1))
      pos = pos + 2
      For i = 1 To lonPerson
        tmpmsg2 = tmpmsg2 & Chr(packet(pos))
        pos = pos + 1
      Next i
      
      
     
      upacket = UBound(packet)
      If pos <= upacket Then
        lonM = GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 2
        For i = 1 To lonM
          msg = msg & Chr(packet(pos))
          pos = pos + 1
        Next i
        'Debug.Print msg
      Else
        msg = tmpmsg2
        ' tibia 8.2+
        ' "trade"
        ' 09 00 96 04 05 00 74 72 61 64 65
      End If
      

    
    Case typeChannel ' channel
      channelB1 = packet(pos)
      channelB2 = packet(pos + 1)
      lastUsedChannelID(idConnection) = GoodHex(channelB1) & " " & GoodHex(channelB2)
      pos = pos + 2
      lonM = GetTheLong(packet(pos), packet(pos + 1))
      pos = pos + 2
      For i = 1 To lonM
        msg = msg & Chr(packet(pos))
        pos = pos + 1
      Next i
      
    Case typeTell2 'tell2
        lonPerson = GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 2
        For i = 1 To lonPerson
            tmpmsg2 = tmpmsg2 & Chr(packet(pos))
            pos = pos + 1
        Next i
        lonM = GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 2
        For i = 1 To lonM
            msg = msg & Chr(packet(pos))
            pos = pos + 1
        Next i
    Case typeChannel2 ' channel2
      channelB1 = packet(pos)
      channelB2 = packet(pos + 1)
      lastUsedChannelID(idConnection) = GoodHex(channelB1) & " " & GoodHex(channelB2)
      pos = pos + 2
      lonM = GetTheLong(packet(pos), packet(pos + 1))
      pos = pos + 2
      For i = 1 To lonM
        msg = msg & Chr(packet(pos))
        pos = pos + 1
      Next i
    Case typeReport ' say in report to gm
      lonM = GetTheLong(packet(pos), packet(pos + 1))
      pos = pos + 2
      For i = 1 To lonM
        msg = msg & Chr(packet(pos))
        pos = pos + 1
      Next i
      
  End Select
  lMsg = LCase(msg)
  ' position spamer
  If TibiaVersionLong < 820 Then
    mtypeComp = &H5
  Else
    mtypeComp = &H7
  End If
  If posSpamActivated(idConnection) = False Then
    If (lMsg = "-pos") Then
      If (mtype = mtypeComp) Then
        posSpamChannelB1(idConnection) = channelB1
        posSpamChannelB2(idConnection) = channelB2
        posSpamActivated(idConnection) = True
        aRes = GiveChannelMessage(idConnection, "Now sending positions to channel " & GoodHex(channelB1) & " " & GoodHex(channelB2) & Chr(10) & "DON'T CLOSE ANY CHANNEL WHILE IT!" & Chr(10) & "-pos again to disable", "Blackd", channelB1, channelB2)
        DoEvents
        EvalClientMessage = 1
        Exit Function
      Else
        posSpamChannelB1(idConnection) = &HFF
        posSpamChannelB2(idConnection) = &HFF
        posSpamActivated(idConnection) = False
        aRes = GiveGMmessage(idConnection, "To activate position spamer you should write -pos in a CHANNEL!", "Blackd")
        DoEvents
        EvalClientMessage = 1
        Exit Function
      End If
    End If
  Else
    If (lMsg = "-pos") Then
      posSpamChannelB1(idConnection) = &HFF
      posSpamChannelB2(idConnection) = &HFF
      posSpamActivated(idConnection) = False
      aRes = GiveGMmessage(idConnection, "Stoped position spam." & Chr(10) & "-pos in a channel to activate it again.", "Blackd")
      DoEvents
      EvalClientMessage = 1
      Exit Function
    End If
  End If
  
  If getSpamActivated(idConnection) = False Then
    If (lMsg = "-get") Then
      If (mtype = mtypeComp) Then
        getSpamChannelB1(idConnection) = channelB1
        getSpamChannelB2(idConnection) = channelB2
        getSpamActivated(idConnection) = True
        aRes = GiveChannelMessage(idConnection, "Now reading positions from channel " & GoodHex(channelB1) & " " & GoodHex(channelB2) & Chr(10) & "DON'T CLOSE ANY CHANNEL WHILE IT!" & Chr(10) & "-get again to disable", "Blackd", channelB1, channelB2)
        DoEvents
        EvalClientMessage = 1
        Exit Function
      Else
        getSpamChannelB1(idConnection) = &HFF
        getSpamChannelB2(idConnection) = &HFF
        getSpamActivated(idConnection) = False
        aRes = GiveGMmessage(idConnection, "To activate position reader you should write -get in a CHANNEL!", "Blackd")
        DoEvents
        EvalClientMessage = 1
        Exit Function
      End If
    End If
  Else
    If (lMsg = "-get") Then
      getSpamChannelB1(idConnection) = &HFF
      getSpamChannelB2(idConnection) = &HFF
      getSpamActivated(idConnection) = False
      aRes = GiveGMmessage(idConnection, "Stoped reading positions." & Chr(10) & "-get in a channel to activate it again.", "Blackd")
      DoEvents
      EvalClientMessage = 1
      Exit Function
    End If
  End If
  
  'custom ng custom var
  
  'unequip ring
  If (lMsg = "exiva unequipr") Then
    aRes = ExecuteInTibia("Exiva > 78 FF FF 09 00 00 $hex-equiped-item:09$ 00 FF FF 03 00 00 01", idConnection, True)
    DoEvents
    EvalClientMessage = 1
    Exit Function
  End If
  
  'equip ering
  If (lMsg = "exiva energyring") Then
    aRes = ExecuteInTibia("exiva #EB 0B 09", idConnection, True)
    DoEvents
    EvalClientMessage = 1
    Exit Function
  End If
  
  'sdmax
  If (lMsg = "exiva sdmax") Then
    If healingCheatsOptions(idConnection).sdmax = False Then
       healingCheatsOptions(idConnection).sdmax = True
       aRes = SendCustomSystemMessageToClient(idConnection, "SDMAX ON", &HB)
    Else
        healingCheatsOptions(idConnection).sdmax = False
        aRes = SendCustomSystemMessageToClient(idConnection, "SDMAX OFF", &HB)
    End If
    EvalClientMessage = 1
    Exit Function
  End If
  
  'attack target
  If (lMsg = "exiva target") Then
    If healingCheatsOptions(idConnection).htarget = False Then
        healingCheatsOptions(idConnection).htarget = True
        aRes = ExecuteInTibia("exiva < B4 15 $hex-tibiastr:Hold Target ON$", idConnection, True)
        aRes = SendCustomSystemMessageToClient(idConnection, "Hold Target ON", &HB)
    Else
        healingCheatsOptions(idConnection).htarget = False
        aRes = ExecuteInTibia("exiva < B4 15 $hex-tibiastr:Hold Target OFF$", idConnection, True)
        aRes = SendCustomSystemMessageToClient(idConnection, "Hold Target OFF", &HB)
        GotKillOrder(idConnection) = False
    End If
    EvalClientMessage = 1
  Exit Function
  End If
  
  'antipush gold
  If (lMsg = "exiva antipush") Then
      If healingCheatsOptions(idConnection).antipush = False Then
      healingCheatsOptions(idConnection).antipush = True
      aRes = ExecuteInTibia("exiva drop D7 0B 01", idConnection, True)
      aRes = ExecuteInTibia("exiva < B4 15 $hex-tibiastr:Antipush ON$", idConnection, True)
      aRes = SendCustomSystemMessageToClient(idConnection, "Antipush ON", &HB)
      Else
      healingCheatsOptions(idConnection).antipush = False
      aRes = ExecuteInTibia("exiva < B4 15 $hex-tibiastr:Antipush OFF$", idConnection, True)
      aRes = SendCustomSystemMessageToClient(idConnection, "Antipush OFF", &HB)
      End If
  EvalClientMessage = 1
  Exit Function
  End If

  'pushmax
  If (lMsg = "exiva pushmax") Then
      If healingCheatsOptions(idConnection).pmax = False Then
        healingCheatsOptions(idConnection).pmax = True
        aRes = ExecuteInTibia("exiva < B4 15 $hex-tibiastr:Pushmax ON$", idConnection, True)
        aRes = SendCustomSystemMessageToClient(idConnection, "Pushmax ON", &HB)
      Else
        healingCheatsOptions(idConnection).pmax = False
        RemoveSpamOrder idConnection, 2
        aRes = StartPush(idConnection, "stoppush")
        aRes = ExecuteInTibia("exiva < B4 15 $hex-tibiastr:Pushmax OFF$", idConnection, True)
        aRes = SendCustomSystemMessageToClient(idConnection, "Pushmax OFF", &HB)
        DoEvents
      End If
    EvalClientMessage = 1
    Exit Function
  End If
  
  
  If Len(msg) > 6 Then
    If Left(lMsg, 6) = "exiva " Then
      keyChar = Mid(lMsg, 7, 1)
      Select Case keyChar
      #If AllowPowerCommands Then
      Case "_"
         res = 1
         rightpart = Right(lMsg, Len(lMsg) - 6)
         aRes = storeVar(idConnection, rightpart)
      Case ">"
         res = 1
         If (Mid$(lMsg, 8, 1) = ">") Then 'no autoheader
            rightpart = Right(lMsg, Len(lMsg) - 8)
            aRes = sendString(idConnection, rightpart, True, False)
         Else 'autoheader
            rightpart = Right(lMsg, Len(lMsg) - 7)
            aRes = sendString(idConnection, rightpart, True, True)
         End If
      Case "<"
         res = 1
         rightpart = Right(msg, Len(msg) - 7)
         aRes = sendString(idConnection, rightpart, False, True)
      Case "#"
        rightpart = Right(lMsg, Len(lMsg) - 7)
        If Len(rightpart) = 8 Then
          res = 1
          b1 = FromHexToDec(Mid(rightpart, 1, 1)) * 16 + FromHexToDec(Mid(rightpart, 2, 1))
          b2 = FromHexToDec(Mid(rightpart, 4, 1)) * 16 + FromHexToDec(Mid(rightpart, 5, 1))
          b3 = FromHexToDec(Mid(rightpart, 7, 1)) * 16 + FromHexToDec(Mid(rightpart, 8, 1))
          aRes = MoveItemToEquip(idConnection, b1, b2, b3)
        End If
      Case "!"
        continueRest = True
        rightpart = Right(lMsg, Len(lMsg) - 7)
        If Len(rightpart) = 5 Then
          res = 1
          b1 = FromHexToDec(Mid(rightpart, 1, 1)) * 16 + FromHexToDec(Mid(rightpart, 2, 1))
          b2 = FromHexToDec(Mid(rightpart, 4, 1)) * 16 + FromHexToDec(Mid(rightpart, 5, 1))
          tempID = GetTheLong(b1, b2)
          If TibiaVersionLong >= 1058 Then
            '  6E 00 25 0B FF 03 00 62 61 67 08 00 01 00 04 00 00 00 04 25 0B FF 05 0B FF B1 0D FF 0A 01 FF 01
            '  6E 00 25 0B FF 03 00 62 61 67 08 00 01 00 01 00 00 00 01 D7 0B FF 10
            If tempID > highestDatTile Then
              If finalModeVar = True Then
                aRes = GiveGMmessage(idConnection, "Received invalid tile number ( " & CStr(tempID) & " ) Max tile number=" & CStr(highestDatTile), "Blackd")
                DoEvents
                continueRest = False
              Else
                ' suppose haveExtraByte=false
                sCheat = "6E 00 " & FiveChrLon(tileID_Bag) & " FF 03 00 62 61 67 08 00 01 00 01 00 00 00 01 " & GoodHex(b1) & " " & GoodHex(b2) & " FF"
                sCheat = "16 00 " & sCheat
              End If
            Else
              If DatTiles(tempID).haveExtraByte = True Then
                 sCheat = "6E 00 " & FiveChrLon(tileID_Bag) & " FF 03 00 62 61 67 08 00 01 00 01 00 00 00 01 " & GoodHex(b1) & " " & GoodHex(b2) & " FF 01"
                 sCheat = "17 00 " & sCheat
              Else
                 sCheat = "6E 00 " & FiveChrLon(tileID_Bag) & " FF 03 00 62 61 67 08 00 01 00 01 00 00 00 01 " & GoodHex(b1) & " " & GoodHex(b2) & " FF"
                 sCheat = "16 00 " & sCheat
              End If

            End If
          ElseIf (TibiaVersionLong >= 990) Then
            '  6E 00 25 0B FF 03 00 62 61 67 08 00 01 01 0E FF 01
            '  6E 00 XX XX FF 03 00 62 61 67 08 00 01 01 0E FF 01
            sCheat = "6E 00 " & FiveChrLon(tileID_Bag) & " FF 03 00 62 61 67 08 00 01 " & GoodHex(b1) & " " & GoodHex(b2) & " FF 01"
            sCheat = "11 00 " & sCheat
          Else
          
            sCheat = "6E 00 " & FiveChrLon(tileID_Bag) & " 03 00 62 61 67 08 00 01 " & GoodHex(b1) & " " & GoodHex(b2)
            sCheat = "0E 00 " & sCheat
          End If
          If continueRest = True Then
            inRes = GetCheatPacket(cPacket, sCheat)
            frmMain.UnifiedSendToClientGame idConnection, cPacket
          End If
        ElseIf Len(rightpart) = 8 Then
          res = 1
          b1 = FromHexToDec(Mid(rightpart, 1, 1)) * 16 + FromHexToDec(Mid(rightpart, 2, 1))
          b2 = FromHexToDec(Mid(rightpart, 4, 1)) * 16 + FromHexToDec(Mid(rightpart, 5, 1))
          b3 = FromHexToDec(Mid(rightpart, 7, 1)) * 16 + FromHexToDec(Mid(rightpart, 8, 1))
          tempID = GetTheLong(b1, b2)
           
          If TibiaVersionLong >= 1058 Then
           If tempID > highestDatTile Then
              If finalModeVar = True Then
                aRes = GiveGMmessage(idConnection, "Received invalid tile number ( " & CStr(tempID) & " ) Max tile number=" & CStr(highestDatTile), "Blackd")
                DoEvents
                continueRest = False
              Else
                ' suppose haveExtraByte=false
                   sCheat = "6E 00 " & FiveChrLon(tileID_Bag) & " FF 03 00 62 61 67 08 00 01 00 01 00 00 00 01 " & GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(b3)
                  sCheat = "16 00 " & sCheat
              End If
            Else
              If DatTiles(tempID).haveExtraByte = True Then
               sCheat = "6E 00 " & FiveChrLon(tileID_Bag) & " FF 03 00 62 61 67 08 00 01 00 01 00 00 00 01 " & GoodHex(b1) & " " & GoodHex(b2) & " FF " & GoodHex(b3)
               sCheat = "17 00 " & sCheat
              Else
                sCheat = "6E 00 " & FiveChrLon(tileID_Bag) & " FF 03 00 62 61 67 08 00 01 00 01 00 00 00 01 " & GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(b3)
               sCheat = "16 00 " & sCheat
              End If

            End If
          ElseIf (TibiaVersionLong >= 990) Then
             sCheat = "6E 00 " & FiveChrLon(tileID_Bag) & " FF 03 00 62 61 67 08 00 01 " & GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(b3) & " 01"
             tempID = GetTheLong(b1, b2)
             sCheat = "11 00 " & sCheat
          Else
             sCheat = "6E 00 " & FiveChrLon(tileID_Bag) & " 03 00 62 61 67 08 00 01 " & GoodHex(b1) & " " & GoodHex(b2)
             tempID = GetTheLong(b1, b2)
             sCheat = "0F 00 " & sCheat & " " & GoodHex(b3)
          End If
          If continueRest = True Then
            inRes = GetCheatPacket(cPacket, sCheat)
            frmMain.UnifiedSendToClientGame idConnection, cPacket
          End If
        ElseIf Len(rightpart) = 11 Then
          res = 1
          b1 = FromHexToDec(Mid(rightpart, 1, 1)) * 16 + FromHexToDec(Mid(rightpart, 2, 1))
          b2 = FromHexToDec(Mid(rightpart, 4, 1)) * 16 + FromHexToDec(Mid(rightpart, 5, 1))
          b3 = FromHexToDec(Mid(rightpart, 7, 1)) * 16 + FromHexToDec(Mid(rightpart, 8, 1))
          b4 = FromHexToDec(Mid(rightpart, 10, 1)) * 16 + FromHexToDec(Mid(rightpart, 11, 1))
          
           If TibiaVersionLong >= 1058 Then
            ' WARNING: 4 bytes tile will fail if haveextrabyte = false
            
            '  6E 00 25 0B FF 03 00 62 61 67 08 00 01 00 04 00 00 00 04 25 0B FF 05 0B FF B1 0D FF 0A 01 FF 01
            '  6E 00 25 0B FF 03 00 62 61 67 08 00 01 00 01 00 00 00 01 D7 0B FF 10
             sCheat = "6E 00 " & FiveChrLon(tileID_Bag) & " FF 03 00 62 61 67 08 00 01 00 01 00 00 00 01 " & GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(b3) & " " & GoodHex(b4)
             sCheat = "17 00 " & sCheat
          ElseIf (TibiaVersionLong >= 990) Then
             sCheat = "6E 00 " & FiveChrLon(tileID_Bag) & " FF 03 00 62 61 67 08 00 01 " & GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(b3) & " " & GoodHex(b4)
             tempID = GetTheLong(b1, b2)
             sCheat = "11 00 " & sCheat
          Else
            sCheat = "6E 00 " & FiveChrLon(tileID_Bag) & " 03 00 62 61 67 08 00 01 " & GoodHex(b1) & " " & GoodHex(b2)
            tempID = GetTheLong(b1, b2)
            sCheat = "10 00 " & sCheat & " " & GoodHex(b3) & " " & GoodHex(b4)
          End If
          inRes = GetCheatPacket(cPacket, sCheat)
          frmMain.UnifiedSendToClientGame idConnection, cPacket
        End If
      Case "0" ' cast SD
        res = 1
        rightpart = Right(msg, Len(lMsg) - 7)
        aRes = SendAimbot(parseVars(idConnection, rightpart), idConnection, LowByteOfLong(tileID_SD), HighByteOfLong(tileID_SD))
        DoEvents
        If frmHardcoreCheats.chkColorEffects.value = 1 Then
          nextLight(idConnection) = "FD"
          enLight idConnection
        End If
      Case "1" ' cast HMM
        res = 1
        rightpart = Right(msg, Len(lMsg) - 7)
        aRes = SendAimbot(parseVars(idConnection, rightpart), idConnection, LowByteOfLong(tileID_HMM), HighByteOfLong(tileID_HMM))
        DoEvents
        If frmHardcoreCheats.chkColorEffects.value = 1 Then
          nextLight(idConnection) = "FD"
          enLight idConnection
        End If
      Case "2" ' cast Explosion
        res = 1
        rightpart = Right(msg, Len(lMsg) - 7)
        aRes = SendAimbot(parseVars(idConnection, rightpart), idConnection, LowByteOfLong(tileID_Explosion), HighByteOfLong(tileID_Explosion))
        DoEvents
        If frmHardcoreCheats.chkColorEffects.value = 1 Then
          nextLight(idConnection) = "FD"
          enLight idConnection
        End If
      Case "3" ' cast IH
        res = 1
        rightpart = Right(msg, Len(lMsg) - 7)
        aRes = SendAimbot(parseVars(idConnection, rightpart), idConnection, LowByteOfLong(tileID_IH), HighByteOfLong(tileID_IH))
        DoEvents
        If frmHardcoreCheats.chkColorEffects.value = 1 Then
          nextLight(idConnection) = "04"
          enLight idConnection
        End If
      Case "4" ' cast UH
        res = 1
        rightpart = Right(msg, Len(lMsg) - 7)
        aRes = SendAimbot(parseVars(idConnection, rightpart), idConnection, LowByteOfLong(tileID_UH), HighByteOfLong(tileID_UH))
        DoEvents
        If frmHardcoreCheats.chkColorEffects.value = 1 Then
          nextLight(idConnection) = "04"
          enLight idConnection
        End If
      Case "5" ' cast SD
        res = 1
        rightpart = Right(msg, Len(lMsg) - 7)
        aRes = SendMobAimbot(parseVars(idConnection, rightpart), idConnection, LowByteOfLong(tileID_SD), HighByteOfLong(tileID_SD))
        DoEvents
        If frmHardcoreCheats.chkColorEffects.value = 1 Then
          nextLight(idConnection) = "FD"
          enLight idConnection
        End If
      Case "6" ' cast HMM
        res = 1
        rightpart = Right(msg, Len(lMsg) - 7)
        aRes = SendMobAimbot(parseVars(idConnection, rightpart), idConnection, LowByteOfLong(tileID_HMM), HighByteOfLong(tileID_HMM))
          DoEvents
        If frmHardcoreCheats.chkColorEffects.value = 1 Then
          nextLight(idConnection) = "FD"
          enLight idConnection
        End If
      Case "7" ' cast Explosion
        res = 1
        rightpart = Right(msg, Len(lMsg) - 7)
        aRes = SendMobAimbot(parseVars(idConnection, rightpart), idConnection, LowByteOfLong(tileID_Explosion), HighByteOfLong(tileID_Explosion))
        DoEvents
        If frmHardcoreCheats.chkColorEffects.value = 1 Then
          nextLight(idConnection) = "FD"
          enLight idConnection
        End If
      Case "8" ' cast IH
        res = 1
        rightpart = Right(msg, Len(lMsg) - 7)
        aRes = SendMobAimbot(parseVars(idConnection, rightpart), idConnection, LowByteOfLong(tileID_IH), HighByteOfLong(tileID_IH))
        DoEvents
        If frmHardcoreCheats.chkColorEffects.value = 1 Then
          nextLight(idConnection) = "04"
          enLight idConnection
        End If
      Case "9" ' cast UH
        res = 1
        rightpart = Right(msg, Len(lMsg) - 7)
        aRes = SendMobAimbot(parseVars(idConnection, rightpart), idConnection, LowByteOfLong(tileID_UH), HighByteOfLong(tileID_UH))
        DoEvents
        If frmHardcoreCheats.chkColorEffects.value = 1 Then
          nextLight(idConnection) = "04"
          enLight idConnection
        End If
        

        
        
        
      Case "+" ' all MC cast
        If Len(msg) > 7 Then
          keyChar2 = LCase(Mid(lMsg, 8, 1))
          rightpart = Right(msg, Len(lMsg) - 8)
          If (rightpart = "") Then
            rightpart = LCase(currTargetName(idConnection))
          End If
          Select Case keyChar2
          Case "0" ' cast SD
            res = 1
            For mcid = 1 To MAXCLIENTS
              If GameConnected(mcid) = True And GotPacketWarning(mcid) = False And sentFirstPacket(mcid) = True Then
                aRes = SendAimbot(parseVars(idConnection, rightpart), mcid, LowByteOfLong(tileID_SD), HighByteOfLong(tileID_SD))
                DoEvents
                If frmHardcoreCheats.chkColorEffects.value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              End If
            Next mcid
          Case "1" ' cast HMM
            res = 1
            For mcid = 1 To MAXCLIENTS
              If GameConnected(mcid) = True And GotPacketWarning(mcid) = False And sentFirstPacket(mcid) = True Then
                aRes = SendAimbot(parseVars(idConnection, rightpart), mcid, LowByteOfLong(tileID_HMM), HighByteOfLong(tileID_HMM))
                DoEvents
                If frmHardcoreCheats.chkColorEffects.value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              End If
            Next mcid
          Case "2" ' cast Explosion
            res = 1
            For mcid = 1 To MAXCLIENTS
              If GameConnected(mcid) = True And GotPacketWarning(mcid) = False And sentFirstPacket(mcid) = True Then
                aRes = SendAimbot(parseVars(idConnection, rightpart), mcid, LowByteOfLong(tileID_Explosion), HighByteOfLong(tileID_Explosion))
                DoEvents
                If frmHardcoreCheats.chkColorEffects.value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              End If
            Next mcid
          Case "3" ' cast IH
            res = 1
            For mcid = 1 To MAXCLIENTS
              If GameConnected(mcid) = True And GotPacketWarning(mcid) = False And sentFirstPacket(mcid) = True Then
                aRes = SendAimbot(parseVars(idConnection, rightpart), mcid, LowByteOfLong(tileID_IH), HighByteOfLong(tileID_IH))
                DoEvents
                If frmHardcoreCheats.chkColorEffects.value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              End If
            Next mcid
          Case "4" ' cast UH
            res = 1
            For mcid = 1 To MAXCLIENTS
              If GameConnected(mcid) = True And GotPacketWarning(mcid) = False And sentFirstPacket(mcid) = True Then
                aRes = SendAimbot(parseVars(idConnection, rightpart), mcid, LowByteOfLong(tileID_UH), HighByteOfLong(tileID_UH))
                DoEvents
                If frmHardcoreCheats.chkColorEffects.value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              End If
            Next mcid
          Case "5" ' cast SD
            res = 1
            For mcid = 1 To MAXCLIENTS
              If GameConnected(mcid) = True And GotPacketWarning(mcid) = False And sentFirstPacket(mcid) = True Then
                aRes = SendMobAimbot(parseVars(idConnection, rightpart), mcid, LowByteOfLong(tileID_SD), HighByteOfLong(tileID_SD))
                DoEvents
                If frmHardcoreCheats.chkColorEffects.value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              End If
            Next mcid
          Case "6" ' cast HMM
            res = 1
            For mcid = 1 To MAXCLIENTS
              If GameConnected(mcid) = True And GotPacketWarning(mcid) = False And sentFirstPacket(mcid) = True Then
                aRes = SendMobAimbot(parseVars(idConnection, rightpart), mcid, LowByteOfLong(tileID_HMM), HighByteOfLong(tileID_HMM))
                DoEvents
                If frmHardcoreCheats.chkColorEffects.value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              End If
            Next mcid
          Case "7" ' cast Explosion
            res = 1
            For mcid = 1 To MAXCLIENTS
              If GameConnected(mcid) = True And GotPacketWarning(mcid) = False And sentFirstPacket(mcid) = True Then
                aRes = SendMobAimbot(parseVars(idConnection, rightpart), mcid, LowByteOfLong(tileID_Explosion), HighByteOfLong(tileID_Explosion))
                DoEvents
                If frmHardcoreCheats.chkColorEffects.value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              End If
            Next mcid
          Case "8" ' cast IH
            res = 1
            For mcid = 1 To MAXCLIENTS
              If GameConnected(mcid) = True And GotPacketWarning(mcid) = False And sentFirstPacket(mcid) = True Then
                aRes = SendMobAimbot(parseVars(idConnection, rightpart), mcid, LowByteOfLong(tileID_IH), HighByteOfLong(tileID_IH))
                DoEvents
                If frmHardcoreCheats.chkColorEffects.value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              End If
            Next mcid
          Case "9" ' cast UH
            res = 1
            For mcid = 1 To MAXCLIENTS
              If GameConnected(mcid) = True And GotPacketWarning(mcid) = False And sentFirstPacket(mcid) = True Then
                aRes = SendMobAimbot(parseVars(idConnection, rightpart), mcid, LowByteOfLong(tileID_UH), HighByteOfLong(tileID_UH))
                DoEvents
                If frmHardcoreCheats.chkColorEffects.value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              End If
            Next mcid
          Case "a" ' type A : Say (text)"
            res = 1
            For mcid = 1 To MAXCLIENTS
              If GameConnected(mcid) = True And GotPacketWarning(mcid) = False And sentFirstPacket(mcid) = True Then
                aRes = ExecuteInTibia(rightpart, mcid, True)
                DoEvents
              End If
            Next mcid
          Case "b" ' cast fireball
            res = 1
            For mcid = 1 To MAXCLIENTS
              If GameConnected(mcid) = True And GotPacketWarning(mcid) = False And sentFirstPacket(mcid) = True Then
                aRes = SendMobAimbot(parseVars(idConnection, rightpart), mcid, LowByteOfLong(tileID_fireball), HighByteOfLong(tileID_fireball))
                DoEvents
                If frmHardcoreCheats.chkColorEffects.value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              End If
            Next mcid
          Case "c" ' cast stalagmite
            res = 1
            For mcid = 1 To MAXCLIENTS
              If GameConnected(mcid) = True And GotPacketWarning(mcid) = False And sentFirstPacket(mcid) = True Then
                aRes = SendMobAimbot(parseVars(idConnection, rightpart), mcid, LowByteOfLong(tileID_stalagmite), HighByteOfLong(tileID_stalagmite))
                DoEvents
                If frmHardcoreCheats.chkColorEffects.value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              End If
            Next mcid
          Case "d" ' cast icicle
            res = 1
            For mcid = 1 To MAXCLIENTS
              If GameConnected(mcid) = True And GotPacketWarning(mcid) = False And sentFirstPacket(mcid) = True Then
                aRes = SendMobAimbot(parseVars(idConnection, rightpart), mcid, LowByteOfLong(tileID_icicle), HighByteOfLong(tileID_icicle))
                DoEvents
                If frmHardcoreCheats.chkColorEffects.value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              End If
            Next mcid
            
            
            
            
          End Select
        End If
        #End If
      Case Else
      
      
      
        rightpart = Right(lMsg, Len(lMsg) - 6)
        If rightpart = "line" Then
          res = 1
          aRes = SendLogSystemMessageToClient(idConnection, "Current cavebot line = " & CStr(exeLine(idConnection)))
          DoEvents
        #If AllowPowerCommands Then
        
       ElseIf rightpart = "a" Then ' cast fireball
            res = 1
            aRes = SendMobAimbot("", idConnection, LowByteOfLong(tileID_fireball), HighByteOfLong(tileID_fireball))
            DoEvents
            If frmHardcoreCheats.chkColorEffects.value = 1 Then
              nextLight(idConnection) = "FD"
              enLight idConnection
            End If
       ElseIf rightpart = "b" Then ' cast stalagmite
            res = 1
            aRes = SendMobAimbot("", idConnection, LowByteOfLong(tileID_stalagmite), HighByteOfLong(tileID_stalagmite))
            DoEvents
            If frmHardcoreCheats.chkColorEffects.value = 1 Then
              nextLight(idConnection) = "04"
              enLight idConnection
            End If
       ElseIf rightpart = "c" Then ' cast icicle
            res = 1
            aRes = SendMobAimbot("", idConnection, LowByteOfLong(tileID_icicle), HighByteOfLong(tileID_icicle))
            DoEvents
            If frmHardcoreCheats.chkColorEffects.value = 1 Then
              nextLight(idConnection) = "04"
              enLight idConnection
            End If
        ElseIf rightpart = "life_fluid" Then
            res = 1
            aRes = UseFluid(idConnection, byteLife)
            
        ElseIf rightpart = "health_potion" Then
            res = 1
            aRes = UsePotion(idConnection, tileID_health_potion)
        ElseIf rightpart = "strong_health_potion" Then
            res = 1
            aRes = UsePotion(idConnection, tileID_strong_health_potion)
        ElseIf rightpart = "great_health_potion" Then
            res = 1
            aRes = UsePotion(idConnection, tileID_great_health_potion)
        ElseIf rightpart = "small_health_potion" Then
            res = 1
            aRes = UsePotion(idConnection, tileID_small_health_potion)
        ElseIf rightpart = "mana_fluid" Then
            res = 1
            aRes = UseFluid(idConnection, byteMana)
        ElseIf rightpart = "mana_potion" Then
            res = 1
            aRes = UsePotion(idConnection, tileID_mana_potion)
        ElseIf rightpart = "mana_potion" Then
            res = 1
            aRes = UsePotion(idConnection, tileID_mana_potion)
        ElseIf rightpart = "strong_mana_potion" Then
            res = 1
            aRes = UsePotion(idConnection, tileID_strong_mana_potion)
        ElseIf rightpart = "great_mana_potion" Then
            res = 1
            aRes = UsePotion(idConnection, tileID_great_mana_potion)
        ElseIf rightpart = "ultimate_health_potion" Then
            res = 1
            aRes = UsePotion(idConnection, tileID_ultimate_health_potion)
        ElseIf rightpart = "great_spirit_potion" Then
            res = 1
            aRes = UsePotion(idConnection, tileID_great_spirit_potion)
        ElseIf rightpart = "uh" Then
          res = 1
          aRes = UseUH(idConnection)
          If aRes = 0 Then
            DoEvents
          End If
          If frmHardcoreCheats.chkColorEffects.value = 1 Then
            nextLight(idConnection) = "04"
            enLight idConnection
          End If
        ElseIf rightpart = "ih" Then
          res = 1
          aRes = UseIH(idConnection)
          If aRes = 0 Then
          DoEvents
          End If
          If frmHardcoreCheats.chkColorEffects.value = 1 Then
            nextLight(idConnection) = "04"
            enLight idConnection
          End If
        ElseIf rightpart = "mana" Then
          res = 1
          aRes = UseFluid(idConnection, byteMana)
          DoEvents
          If frmHardcoreCheats.chkColorEffects.value = 1 Then
            nextLight(idConnection) = "04"
            enLight idConnection
          End If
        ElseIf rightpart = "screenshot" Then
          res = 1
          GetScreenshot frmScreenshot, getScreenshotname()
          DoEvents
        ElseIf rightpart = "expreset" Then
          res = 1
          aRes = expReset(idConnection)
          DoEvents
        ElseIf rightpart = "testsound" Then
          res = 1
          ChangePlayTheDangerSound True
          frmRunemaker.ChkDangerSound.value = 1
          aRes = SendLogSystemMessageToClient(idConnection, "Testing the sound. Danger alarm activated. Deactivate with exiva cancel")
          DoEvents
        ElseIf rightpart = "testding" Then
          res = 1
          PlayMsgSound = True
        ElseIf rightpart = "fastuh" Then
          res = 1
          aRes = UseFastUH(idConnection)
          DoEvents
        ElseIf rightpart = "blueaura" Then
          res = 1
          If GetSpamOrderPosition(idConnection, 3) = 0 Then
            AddSpamOrder idConnection, 3
            aRes = GiveGMmessage(idConnection, "Casting fast UHs each " & CStr(BlueAuraDelay) & " mseconds", "BLUE AURA ACTIVATED")
            DoEvents
          Else
            RemoveSpamOrder idConnection, 3
            aRes = GiveGMmessage(idConnection, "Back to normal autouh status", "BLUE AURA STOPPED")
            DoEvents
          End If
          DoEvents
        ElseIf rightpart = "pos" Then
          res = 1
          aRes = SendLogSystemMessageToClient(idConnection, "Your position is : " & myX(idConnection) & "," & myY(idConnection) & "," & myZ(idConnection))
          DoEvents
        ElseIf rightpart = "exp" Then
          res = 1
          aRes = GiveExpInfo(idConnection, frmHardcoreCheats.txtExivaExpFormat.Text)
          DoEvents
        ElseIf rightpart = "usex" Then
          res = 1
          aRes = CLng(GiveExpInfo(idConnection, frmHardcoreCheats.txtExivaExpFormat.Text))
          DoEvents
        ElseIf rightpart = "fish" Then
          res = 1
          aRes = CatchFish(idConnection)
          DoEvents
        ElseIf rightpart = "save" Then
          res = 1
          SaveCharSettings idConnection
          DoEvents
        ElseIf rightpart = "turbo" Then
          res = 1
          aRes = DoTurbo(idConnection)
          DoEvents
        ElseIf rightpart = "addmove" Then
          res = 1
          AddCavebotMove
          aRes = SendLogSystemMessageToClient(idConnection, "Added cavebot script line: move " & myX(idConnection) & "," & myY(idConnection) & "," & myZ(idConnection))
          DoEvents
        ElseIf rightpart = "dictionary" Then
          res = 1
          aRes = PrintDictionary(idConnection)
          DoEvents
        ElseIf rightpart = "version" Then
          res = 1
          'ProxyVersion
          aRes = SendLogSystemMessageToClient(idConnection, "Blackd proxy version: " & ProxyVersion)
          DoEvents
        ElseIf rightpart = "state" Then
          res = 1
          aRes = SendLogSystemMessageToClient(idConnection, "(For debug) Internal state:")
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "cavebotCurrentTargetPriority(idConnection) = " & _
           CStr(cavebotCurrentTargetPriority(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "myExp(idConnection) = " & _
           CStr(myExp(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "myHP(idConnection) = " & _
           CStr(myHP(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "myMaxHP(idConnection) = " & _
           CStr(myMaxHP(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "myMaxMana(idConnection) = " & _
           CStr(myMaxMana(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "myMana(idConnection) = " & _
           CStr(myMana(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "myMagLevel(idConnection) = " & _
           CStr(myMagLevel(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "mySoulpoints(idConnection) = " & _
           CStr(mySoulpoints(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "gotPacketWarning(idConnection) = " & _
           BooleanAsStr(GotPacketWarning(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "sentWelcome(idConnection) = " & _
           BooleanAsStr(sentWelcome(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "autoLoot(idConnection) = " & _
           BooleanAsStr(autoLoot(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "DangerGM(idConnection) = " & _
           BooleanAsStr(DangerGM(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "DangerPK(idConnection) = " & _
           BooleanAsStr(DangerPK(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "GetTickCount() = " & _
           CStr(GetTickCount()))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "lootTimeExpire(idConnection) = " & _
           CStr(lootTimeExpire(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "onDepotPhase(idConnection) = " & _
           CStr(onDepotPhase(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "cavebotOnDanger(idConnection) = " & _
           CStr(cavebotOnDanger(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "makingRune(idConnection) = " & _
           BooleanAsStr(makingRune(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "moveRetry(idconnection) = " & _
          CStr(moveRetry(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "ignoreNext(idconnection) = " & _
          CStr(ignoreNext(idConnection)))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "GetTickCount() = " & _
          CStr(GetTickCount()))
          DoEvents
          aRes = SendLogSystemMessageToClient(idConnection, "cavebotLenght(idconnection) = " & _
          CStr(cavebotLenght(idConnection)))
          DoEvents
        ElseIf rightpart = "bomb" Then
          res = 1
          aRes = ExecuteMagebomb(idConnection, "")
          DoEvents
        ElseIf rightpart = "all" Then
          res = 1
          If TibiaVersionLong >= 773 Then
            aRes = RevealAll2(idConnection)
            DoEvents
          Else
            aRes = RevealAll(idConnection)
            DoEvents
          End If
          If frmHardcoreCheats.chkColorEffects.value = 1 Then
            nextLight(idConnection) = "1F"
            enLight idConnection
          End If
        ElseIf Left$(rightpart, 6) = "phone " Then
          res = 1
          TelephoneCall Right$(rightpart, Len(rightpart) - 6)
        ElseIf rightpart = "relog" Then
          res = 1
          If (TibiaVersionLong >= 841) Then
            aRes = GiveGMmessage(idConnection, "Sorry, automatic reconnection is not possible since Tibia 8.41", "BlackdProxy")
            DoEvents
          Else
            If (ReconnectionStage(idConnection) = 0) Or (ReconnectionStage(idConnection) = 3) Then
              If ((TimesWarnedAboutRelog = 0) And (Antibanmode = 1)) Then
                  TimesWarnedAboutRelog = TimesWarnedAboutRelog + 1
                  aRes = GiveGMmessage(idConnection, "Using this kind of relog now you risk for a ban or deletion, specially after server save, server kicks and game updates. If you accept the risk then type the command again, else relogin from start, using acc/password.", "BlackdProxy")
                  DoEvents
              Else
                  StartReconnection idConnection
              End If
            End If
          End If
        ElseIf rightpart = "turn0" Then
          res = 1
          TurnMe idConnection, 0
        ElseIf rightpart = "turn1" Then
          res = 1
          TurnMe idConnection, 1
        ElseIf rightpart = "turn2" Then
          res = 1
          TurnMe idConnection, 2
        ElseIf rightpart = "turn3" Then
          res = 1
          TurnMe idConnection, 3
        ElseIf rightpart = "openbp" Then
          res = 1
          aRes = openBP(idConnection)
        ElseIf rightpart = "close" Then
          res = 1
          If ReconnectionStage(idConnection) > 0 Then
            ReconnectionStage(idConnection) = 10
          End If
          frmMain.DoCloseActions idConnection
          frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "(Stoping alarm because it was a desired close)"
          ChangePlayTheDangerSound False
          ReconnectionStage(idConnection) = 0
        ElseIf Left$(rightpart, 4) = "log " Then
          res = 1
          LogOnFile "log_" & CharacterName(idConnection) & ".txt", parseVars(idConnection, Right$(rightpart, Len(rightpart) - 4))
        ElseIf Left$(rightpart, 5) = "plot " Then
          res = 1
          aRes = PlotPosition(idConnection, Right(rightpart, Len(rightpart) - 5))
        ElseIf Left$(rightpart, 5) = "drop " Then
          res = 1
          aRes = DropItemOnGround(idConnection, Right(rightpart, Len(rightpart) - 5))
        ElseIf Left$(rightpart, 18) = "useitemwithamount:" Then
          res = 1
          aRes = UseItemWithAmount(idConnection, Right(rightpart, Len(rightpart) - 18))
        ElseIf Left$(rightpart, 14) = "useitemonname:" Then
          res = 1
          aRes = UseItemOnName(idConnection, Right(rightpart, Len(rightpart) - 14))
        ElseIf Left$(rightpart, 5) = "speed" Then
          res = 1
          aRes = ChangeSpeed(idConnection, Right(rightpart, Len(rightpart) - 5))
        ElseIf Left(rightpart, 5) = "+push" Then
          res = 1
          If Len(rightpart) > 4 Then
            rightpart = Right(rightpart, Len(rightpart) - 5)
            For mcid = 1 To MAXCLIENTS
              If GameConnected(mcid) = True And GotPacketWarning(mcid) = False And sentFirstPacket(mcid) = True Then
                aRes = StartPush(idConnection, rightpart)
              End If
            Next mcid
          Else
            For mcid = 1 To MAXCLIENTS
              If GameConnected(mcid) = True And GotPacketWarning(mcid) = False And sentFirstPacket(mcid) = True Then
                aRes = StartPush(idConnection, "")
              End If
            Next mcid
          End If
          
        ElseIf Left(rightpart, 2) = "b:" Then
            res = 1
            rightpart = Right(rightpart, Len(rightpart) - 2)
            aRes = SendMobAimbot(parseVars(idConnection, rightpart), idConnection, LowByteOfLong(tileID_fireball), HighByteOfLong(tileID_fireball))
            DoEvents
            If frmHardcoreCheats.chkColorEffects.value = 1 Then
              nextLight(idConnection) = "FD"
              enLight idConnection
            End If
        ElseIf Left(rightpart, 2) = "c:" Then
            res = 1
            rightpart = Right(rightpart, Len(rightpart) - 2)
            aRes = SendMobAimbot(parseVars(idConnection, rightpart), idConnection, LowByteOfLong(tileID_stalagmite), HighByteOfLong(tileID_stalagmite))
            DoEvents
            If frmHardcoreCheats.chkColorEffects.value = 1 Then
              nextLight(idConnection) = "FD"
              enLight idConnection
            End If
        ElseIf Left(rightpart, 2) = "d:" Then
            res = 1
            rightpart = Right(rightpart, Len(rightpart) - 2)
            aRes = SendMobAimbot(parseVars(idConnection, rightpart), idConnection, LowByteOfLong(tileID_icicle), HighByteOfLong(tileID_icicle))
            DoEvents
            If frmHardcoreCheats.chkColorEffects.value = 1 Then
              nextLight(idConnection) = "FD"
              enLight idConnection
            End If
        ElseIf Left(rightpart, 4) = "push" Then
          res = 1
          If Len(rightpart) > 4 Then
            rightpart = Right(rightpart, Len(rightpart) - 5)
            aRes = StartPush(idConnection, rightpart)
          Else
            aRes = StartPush(idConnection, "")
          End If
        ElseIf Left(rightpart, 4) = "load" Then
          res = 1
          If Len(rightpart) > 4 Then
            rightpart = Right(rightpart, Len(rightpart) - 5)
            tmpStr = LoadCharSettings(idConnection, rightpart)
          Else
            tmpStr = LoadCharSettings(idConnection)
          End If
          If tmpStr = "" Then
              aRes = SendLogSystemMessageToClient(idConnection, "Load sucesfull")
              DoEvents
          Else
              aRes = GiveGMmessage(idConnection, tmpStr, "BlackdProxy")
              DoEvents
          End If
        ElseIf Left(rightpart, 9) = "autocombo" Then
          res = 1
          If Len(rightpart) > 9 Then
            rightpart = Right$(rightpart, Len(rightpart) - 9)
            aRes = DoAutocombo(idConnection, rightpart)
          Else
            aRes = DoAutocombo(idConnection, "")
          End If
        ElseIf Left(rightpart, 5) = "bomb " Then
          res = 1
          If Len(rightpart) > 5 Then
            rightpart = Right(rightpart, Len(rightpart) - 5)
            aRes = ExecuteMagebomb(idConnection, rightpart)
          Else
            aRes = ExecuteMagebomb(idConnection, "")
          End If
        ElseIf Left(rightpart, 5) = "sayt:" Then
          res = 1
          If Len(rightpart) > 5 Then
            rightpart = Right(rightpart, Len(rightpart) - 5)
            aRes = SayInTrade(idConnection, rightpart)
          Else
            aRes = SayInTrade(idConnection, "")
          End If
        ElseIf Left(rightpart, 5) = "sell:" Then
          res = 1
          If Len(rightpart) > 5 Then
            rightpart = Right(rightpart, Len(rightpart) - 5)
            aRes = SellInTrade(idConnection, rightpart)
          Else
            aRes = SellInTrade(idConnection, "")
          End If
        ElseIf Left(rightpart, 4) = "buy:" Then
          res = 1
          If Len(rightpart) > 4 Then
            rightpart = Right(rightpart, Len(rightpart) - 4)
            aRes = BuyInTrade(idConnection, rightpart)
          Else
            aRes = BuyInTrade(idConnection, "")
          End If
        ElseIf Left(rightpart, 4) = "view" Then
          res = 1
          If TibiaVersionLong >= 841 Then
            aRes = SendLogSystemMessageToClient(idConnection, "Unsafe command since Tibia 8.41 , use exiva xray instead (and this is still risky)")
              DoEvents
          Else
            If Len(rightpart) > 4 Then
              rightpart = Right(rightpart, Len(rightpart) - 4)
              aRes = ViewFloor(idConnection, rightpart)
              If aRes = -1 Then
                aRes = SendLogSystemMessageToClient(idConnection, "Incorrect parameter for exiva viewfloor")
                DoEvents
              End If
            End If
          End If
        ElseIf Left(rightpart, 4) = "xray" Then
          res = 1
          If Len(rightpart) > 4 Then
            rightpart = Right$(rightpart, Len(rightpart) - 4)
            aRes = MemoryChangeFloor(idConnection, rightpart)
            If aRes = -1 Then
              aRes = SendLogSystemMessageToClient(idConnection, "Incorrect parameter for exiva viewfloor")
              DoEvents
            End If
          End If
        ElseIf Left(rightpart, 4) = "kill" Then
          res = 1
          rightpart = Right(rightpart, Len(rightpart) - 4)
          aRes = ProcessKillOrder(idConnection, rightpart)
        ElseIf Left(rightpart, 5) = "mcit " Then
          res = 1
          rightpart = Right(rightpart, Len(rightpart) - 5)
          For mcid = 1 To MAXCLIENTS
            If GameConnected(mcid) = True Then
              aRes = ExecuteInTibia(rightpart, mcid, True)
            End If
          Next mcid
        ElseIf Left(rightpart, 6) = "check " Then
          res = 1
          rightpart = Right(rightpart, Len(rightpart) - 6)
          aRes = IngameCheck(idConnection, rightpart)
        ElseIf Left(rightpart, 6) = "ignore" Then
            res = 1
            
            If cavebotEnabled(idConnection) = False Then
              tmpID = currTargetID(idConnection)
            Else
              tmpID = lastAttackedID(idConnection)
            End If
            
            If tmpID = 0 Then
              aRes = SendLogSystemMessageToClient(idConnection, "Warning: No target detected. Can't add that to ignore list!!")
              DoEvents
            Else
              aRes2 = AddIgnoredcreature(idConnection, tmpID)
              aRes = MeleeAttack(idConnection, 0)
              lastAttackedID(idConnection) = 0
              'If publicDebugMode = True Then
              If aRes2 = 0 Then
                aRes = SendLogSystemMessageToClient(idConnection, "Creature ID #" & CStr(tmpID) & _
                 " ( " & GetNameFromID(idConnection, tmpID) & " ) will be ignored")
                DoEvents
              End If
              'End If
            End If

        ElseIf Left(rightpart, 12) = "resetignores" Then
          res = 1
          RemoveAllIgnoredcreature idConnection
          aRes = SendLogSystemMessageToClient(idConnection, "List of ignored creatures ids is now reseted and clean")
          DoEvents
        ElseIf Left(rightpart, 6) = "outfit" Then
          res = 1
          tmpStr = Right(rightpart, Len(rightpart) - 6)
          SendOutfit idConnection, tmpStr
          DoEvents
        ElseIf rightpart = "pause" Then
          res = 1
          For mcid = 1 To MAXCLIENTS
            DangerPK(idConnection) = False
            DangerGM(idConnection) = False
            DangerPlayer(idConnection) = False
            RemoveSpamOrder idConnection, 1 'remove  auto UH
            UHRetryCount(idConnection) = 0
            logoutAllowed(idConnection) = 0
          Next mcid
          ChangePlayTheDangerSound False
          aRes = ChangePauseStatus(idConnection, True, False)
        ElseIf rightpart = "pause-" Then
          res = 1
          For mcid = 1 To MAXCLIENTS
            DangerPK(idConnection) = False
            DangerGM(idConnection) = False
            DangerPlayer(idConnection) = False
            RemoveSpamOrder idConnection, 1 'remove  auto UH
            UHRetryCount(idConnection) = 0
            logoutAllowed(idConnection) = 0
          Next mcid
          ChangePlayTheDangerSound False
          aRes = ChangePauseStatus(idConnection, True, True)
        ElseIf rightpart = "dance" Then
       ' Debug.Print "antiidle"
          res = 1
          aRes = randomNumberBetween(0, 3)
          tmpStr = "exiva turn" & CStr(aRes)
          tempID = GetTickCount() + 100
          AddSchedule idConnection, tmpStr, tempID
          aRes = aRes + 1
          tmpStr = "exiva turn" & CStr(aRes Mod 4)
          tempID = tempID + 300
          AddSchedule idConnection, tmpStr, tempID
          aRes = aRes + 1
          tmpStr = "exiva turn" & CStr(aRes Mod 4)
          tempID = tempID + 300
          AddSchedule idConnection, tmpStr, tempID
          aRes = aRes + 1
          tmpStr = "exiva turn" & CStr(aRes Mod 4)
          tempID = tempID + 300
          AddSchedule idConnection, tmpStr, tempID
        ElseIf rightpart = "play" Then
          res = 1
          DangerPlayer(idConnection) = False
          aRes = ChangePauseStatus(idConnection, False, False)
        #End If
        ElseIf Left(rightpart, 5) = "jump " Then
          res = 1
          'delete partial progress on current instruction
          fishCounter(idConnection) = 0
          onDepotPhase(idConnection) = 0
          'jump to given cavebot script line
         ' exeLine(idConnection) = CLng(Right(rightpart, Len(rightpart) - 5))
           updateExeLine idConnection, CLng(Right(rightpart, Len(rightpart) - 5)), False
          aRes = SendLogSystemMessageToClient(idConnection, "Jumped to script line " & CStr(exeLine(idConnection)))
          DoEvents
        ElseIf rightpart = "cancel" Then
          res = 1
          For mcid = 1 To MAXCLIENTS
            DangerPK(mcid) = False
            DangerGM(mcid) = False
            DangerPlayer(mcid) = False
            LogoutTimeGM(mcid) = 0
            moveRetry(mcid) = 0
            RemoveSpamOrder mcid, 1
            UHRetryCount(mcid) = 0
            logoutAllowed(mcid) = 0
          Next mcid
          ChangePlayTheDangerSound False
          
          aRes = SendLogSystemMessageToClient(idConnection, "Switched off all alarms. All scheduled close have been canceled.")
          DoEvents
        ElseIf rightpart = "other" Then
          res = 0
        End If
      End Select
    End If
  End If
  EvalClientMessage = res
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " # lost connection at EvalClientMessage . Number : " & Err.Number & " Description : " & Err.Description & " Source: " & Err.Source
  frmMain.DoCloseActions idConnection
  DoEvents
  EvalClientMessage = -1
End Function

Public Function ApplyHardcoreCheats(ByRef packet() As Byte, ByVal idConnection As Integer, Optional forceEval As Boolean = False) As Integer
  Dim res As Integer
  Dim lon As Long
  Dim Target As String
  Dim tellMsg As String
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim i As Long
  Dim lastB As Long
  Dim tileID As Long
  'Dim fRes As TypeSearchItemResult
  Dim myS As Byte
  Dim aRes As Long
  Dim tmpID As Double
  Dim tByte As Byte
  Dim isDamageRune As Boolean
  Dim currHP As Long
  #If FinalMode Then
  On Error GoTo badError
  #End If
  res = 0
  
  If TrialVersion = True Then
    If sentWelcome(idConnection) = False Or GotPacketWarning(idConnection) = True Then
      ApplyHardcoreCheats = res
      Exit Function
    End If
  End If
    
  If GameConnected(idConnection) = True Then
    lastB = UBound(packet)
    lon = GetTheLong(packet(0), packet(1))
    tByte = packet(2)
    Select Case tByte
    Case &H1E
      If publicDebugMode = True Then
        aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Client sent ping reply 1E")
        DoEvents
      End If
    Case &H14
      logoutAllowed(idConnection) = 20000 + GetTickCount() ' disable reconnection 20 sec
    Case &HD3
      If SendingSpecialOutfit(idConnection) = True Then ' sending special outfit
        res = 1
        If TibiaVersionLong > 760 Then
          aRes = SendOutfit4(idConnection, packet(3), packet(4), packet(5), packet(6), packet(7), packet(8))
        Else
          aRes = SendOutfit2(idConnection, packet(3), packet(4), packet(5), packet(6), packet(7))
        End If
        DoEvents
      End If
    Case &H96
      If lastB > 8 Then ' command
        If (frmHardcoreCheats.chkApplyCheats.value = 1) Then
            If ((frmStealth.chkStealthCommands.value = 0) Or (forceEval = True)) Then
                res = EvalClientMessage(idConnection, packet, 2)
            Else
                res = 0
            End If
        End If
      End If
    Case &H82 ' click something (maybe bed)
      logoutAllowed(idConnection) = 6000 + GetTickCount() ' allow logout for 6 seconds
    Case &H79
    ' 04 00 79 C4 1E 00
    ' new since tibia 8.21+
      If frmCheats.chkInspectTileID.value = 1 Then ' identify
        'add tile info by tell
        aRes = SendLogSystemMessageToClient(idConnection, "You see tile ID " & _
         GoodHex(packet(3)) & " " & GoodHex(packet(4)) & "  with info " & GetTileInfoString(packet(3), packet(4)))
   
        DoEvents
      End If
    Case &H8C
      If frmCheats.chkInspectTileID.value = 1 Then ' identify
        'add tile info by tell
        aRes = SendLogSystemMessageToClient(idConnection, "You see tile ID " & _
         GoodHex(packet(8)) & " " & GoodHex(packet(9)) & "  with info " & GetTileInfoString(packet(8), packet(9)))
   
        DoEvents
      End If
    Case &HA1 ' melee attack
      tmpID = FourBytesDouble(packet(3), packet(4), packet(5), packet(6))
      If tmpID <> 0 Then
        currTargetID(idConnection) = tmpID
        currTargetName(idConnection) = GetNameFromID(idConnection, tmpID)
        'aRes = SendLogSystemMessageToClient(idConnection, "You selected this target : " & currTargetName(idConnection))
        'DoEvents
      End If
      If TrainerOptions(idConnection).misc_stoplowhp = 1 Then
         currHP = GetHPFromID(idConnection, tmpID)
         If (currHP < TrainerOptions(idConnection).stoplowhpHP) Then
           'currTargetID(idconnection) = 0
           'currTargetName(idconnection) = ""
           aRes = SendLogSystemMessageToClient(idConnection, "BlackdProxy: WARNING > " & GetNameFromID(idConnection, tmpID) & " is " & GetHPFromID(idConnection, tmpID) & "% hp")
           DoEvents
          ' res = -1 'ignore attack order
         End If
      End If
    Case &H9E
        doingTrade(idConnection) = False
    Case &H84
      'rune attack
      tileID = GetTheLong(packet(8), packet(9))
      Select Case tileID
      Case tileID_IH
        isDamageRune = False
      Case tileID_UH
        isDamageRune = False
      Case Else
        isDamageRune = True
      End Select
      If (isDamageRune = True) Then
        tmpID = FourBytesDouble(packet(11), packet(12), packet(13), packet(14))
        If tmpID <> 0 Then
          currTargetID(idConnection) = tmpID
          currTargetName(idConnection) = GetNameFromID(idConnection, tmpID)
          'aRes = SendLogSystemMessageToClient(idConnection, "You selected this target : " & currTargetName(idConnection))
          'DoEvents
        End If
        If TrainerOptions(idConnection).misc_stoplowhp = 1 Then
          If GetHPFromID(idConnection, tmpID) < TrainerOptions(idConnection).stoplowhpHP Then
            'currTargetID(idconnection) = 0
            'currTargetName(idconnection) = ""
            aRes = SendLogSystemMessageToClient(idConnection, "BlackdProxy: WARNING > " & GetNameFromID(idConnection, tmpID) & " is " & GetHPFromID(idConnection, tmpID) & "% hp")
            DoEvents
            'res = -1 'ignore attack order
          End If
        End If
      End If
    Case Else
      ' nothing
    End Select
  Else 'close

  End If
  ApplyHardcoreCheats = res
  Exit Function
badError:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "got UNEXPECTED ERROR IN ID " & idConnection & " (" & Err.Description & ") when client sent a packet . Description: " & Err.Description
  ApplyHardcoreCheats = 0
End Function
Public Function SendAimbot(Target As String, idConnection As Integer, runeB1 As Byte, runeB2 As Byte) As Long
  Dim aRes As Long
  Dim lTarget As String
  Dim lSquare As String
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim thing As String
  Dim fRes As TypeSearchItemResult2
  Dim myS As Byte
  Dim x As Long
  Dim y As Long
  Dim s As Long
  Dim tileID As Long
  Dim tmpID As Double
  Dim inRes As Integer
  Dim SpecialSource As Boolean
  Dim isDamageRune As Boolean
  Dim percent As Long
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  tileID = GetTheLong(runeB1, runeB2)
  Select Case tileID
  Case tileID_SD
    thing = "SDs"
    isDamageRune = True
  Case tileID_HMM
    thing = "HMMs"
    isDamageRune = True
  Case tileID_Explosion
    thing = "Explosions"
    isDamageRune = True
  Case tileID_IH
    thing = "IHs"
    isDamageRune = False
  Case tileID_UH
    thing = "UHs"
    isDamageRune = False
  Case tileID_fireball
    thing = "Fireballs"
    isDamageRune = True
  Case tileID_stalagmite
    thing = "Stalagmites"
    isDamageRune = True
  Case tileID_icicle
    thing = "Icicles"
    isDamageRune = True

  Case Else
    thing = "item " & GoodHex(runeB1) & " " & GoodHex(runeB2)
    isDamageRune = False
  End Select
  If (isDamageRune = True) Then
    If frmHardcoreCheats.chkProtectedShots.value = 1 Then
      percent = 100 * ((myHP(idConnection) / myMaxHP(idConnection)))
      If (percent < GLOBAL_RUNEHEAL_HP) Then
        aRes = SendLogSystemMessageToClient(idConnection, "BlackdProxy: Your shot have been blocked for safety (low hp)")
        DoEvents
        SendAimbot = -2
        Exit Function
      End If
    End If
  End If
  
  
  Dim lLastTargetName As String
  lLastTargetName = LCase(currTargetName(idConnection)) 'currTargetName=name of last target, even if there is no target atm..
  SpecialSource = False
  If (frmHardcoreCheats.chkTotalWaste.value = True) Then 'And (TibiaVersionLong >= 773)) Then
    SpecialSource = True
  End If
 ' aRes = SendMessageToClient(idConnection, "Casting on " & target & " ;)", "GM BlackdProxy")
  ' search the rune
  fRes = SearchItem(idConnection, runeB1, runeB2)  'search thing
  If fRes.foundCount = 0 Then
     If (SpecialSource = False) And (Not (frmHardcoreCheats.chkEnhancedCheats = True)) Then 'And (TibiaVersionLong >= 773))) Then
       aRes = SendSystemMessageToClient(idConnection, "can't find " & thing & ", open new bp of " & thing & "!")
       SendAimbot = 0
       Exit Function
     Else
       SpecialSource = True
     End If
  End If
  If (TibiaVersionLong < 760) Then
    myS = MyStackPos(idConnection)
  Else
    myS = FirstPersonStackPos(idConnection)
  End If
  ' search yourself
  If myS = &HFF Then
    aRes = SendLogSystemMessageToClient(idConnection, "Your map is out of sync, can't use " & thing & "!")
    SendAimbot = 0
    Exit Function
  End If
  ' search the person
  lTarget = LCase(Target)
  Dim RedSquareID As Long
  RedSquareID = ReadRedSquare(idConnection)
  For y = -6 To 7
    For x = -8 To 9
      For s = 1 To 10
        tmpID = Matrix(y, x, myZ(idConnection), idConnection).s(s).dblID
        If tmpID = 0 Then
          lSquare = ""
        Else
          lSquare = LCase(GetNameFromID(idConnection, tmpID))
        End If
        If (Len(lTarget) <> 0 And lSquare = lTarget) Or (Len(lTarget) = 0 And RedSquareID <> 0 And RedSquareID = tmpID) Or (((RedSquareID = 0 Or Not HPOfID(idConnection).Exists(RedSquareID)) Or GetHPFromID(idConnection, CDbl(RedSquareID)) = 0) And Len(lTarget) = 0 And Len(lLastTargetName) <> 0 And lLastTargetName = lSquare) Then
        '0D 00 84 FF FF 40 00 00 40 0C 00 CB DD 01 40
          If SpecialSource = True Then
               sCheat = "83 FF FF 00 00 00 " & GoodHex(runeB1) & " " & GoodHex(runeB2) & " " & _
               "00 " & GetHexStrFromPosition(myX(idConnection) + x, myY(idConnection) + y, myZ(idConnection)) & _
               " 63 00 " & GoodHex(CByte(s))
               SafeCastCheatString "SendAimbot1", idConnection, sCheat
          Else
               aRes = SendSystemMessageToClient(idConnection, CStr(fRes.foundCount) & " " & thing & " found - Casting one from bp ID " & _
                CStr(CLng(fRes.bpID)) & " slot " & CStr(CLng(fRes.slotID)))
               DoEvents 'added in 8.71
               sCheat = "83 FF FF " & GoodHex(&H40 + fRes.bpID) & " 00 " & _
               GoodHex(fRes.slotID) & " " & GoodHex(runeB1) & " " & GoodHex(runeB2) & " " & _
               GoodHex(fRes.slotID) & " " & GetHexStrFromPosition(myX(idConnection) + x, myY(idConnection) + y, myZ(idConnection)) & _
               " 63 00 " & GoodHex(CByte(s))
              SafeCastCheatString "SendAimbot2", idConnection, sCheat
          End If
          SendAimbot = 0
          Exit Function
        End If
      Next s
    Next x
  Next y
  aRes = SendSystemMessageToClient(idConnection, "Sorry, " & Target & " is not on BlackdProxy track")
  SendAimbot = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at SendAimbot #"
  frmMain.DoCloseActions idConnection
  DoEvents
  SendAimbot = -1
End Function
Public Function FindCreatureByName(ByVal Target As String, idConnection As Integer, ByRef foundX As Long, ByRef foundY As Long, ByRef foundZ As Long) As Boolean
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim s As Long
    Dim tmpName As String
    Dim tmpID As Double
    foundX = -1
    foundY = -1
    foundZ = -1
    If (Len(Target) = 0) Then
        FindCreatureByName = False
        Exit Function
    End If
    Target = LCase(Target)
    
  For z = -1 To 1 'just 1 floor below, 1 floor above, and current floor, try to save some c
  For y = -6 To 7
    For x = -8 To 9
      For s = 1 To 10
        tmpID = Matrix(y, x, myZ(idConnection) + z, idConnection).s(s).dblID
        If tmpID = 0 Then
        '...
        Else
          tmpName = LCase(GetNameFromID(idConnection, tmpID))
          If (tmpName = Target) Then
            'found it!
            foundX = myX(idConnection) + x
            foundY = myY(idConnection) + y
            foundZ = myZ(idConnection) + z
            FindCreatureByName = True
            Exit Function
          End If
        End If
        Next s
    Next x
  Next y
  Next z
  
    'failed to find creature
    FindCreatureByName = False
End Function


Public Function SendMobAimbot(Target As String, idConnection As Integer, runeB1 As Byte, runeB2 As Byte) As Long
  Dim aRes As Long
  Dim lTarget As String
  Dim lSquare As String
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim thing As String
  Dim fRes As TypeSearchItemResult2
  Dim myS As Byte
  Dim x As Long
  Dim y As Long
  Dim s As Long
  Dim tileID As Long
  Dim tmpID As Double
  Dim inRes As Integer
  Dim SpecialSource As Boolean
  Dim isDamageRune As Boolean
  Dim percent As Long
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  tileID = GetTheLong(runeB1, runeB2)
  Select Case tileID
  Case tileID_SD
    thing = "SDs"
    isDamageRune = True
  Case tileID_HMM
    thing = "HMMs"
    isDamageRune = True
  Case tileID_Explosion
    thing = "Explosions"
    isDamageRune = True
  Case tileID_IH
    thing = "IHs"
    isDamageRune = False
  Case tileID_UH
    thing = "UHs"
    isDamageRune = False
  Case tileID_fireball
    thing = "Fireballs"
    isDamageRune = True
  Case tileID_stalagmite
    thing = "Stalagmites"
    isDamageRune = True
  Case tileID_icicle
    thing = "Icicles"
    isDamageRune = True
    
  Case Else
    thing = "runes"
    isDamageRune = False
  End Select
  If (isDamageRune = True) Then
    If frmHardcoreCheats.chkProtectedShots.value = 1 Then
      percent = 100 * ((myHP(idConnection) / myMaxHP(idConnection)))
      If (percent < GLOBAL_RUNEHEAL_HP) Then
        aRes = SendLogSystemMessageToClient(idConnection, "BlackdProxy: Your shot have been blocked for safety (low hp)")
        DoEvents
        SendMobAimbot = -2
        Exit Function
      End If
    End If
  End If
  
 ' aRes = SendMessageToClient(idConnection, "Casting on " & target & " ;)", "GM BlackdProxy")
  ' search the rune
      If ((frmHardcoreCheats.chkEnhancedCheats.value = True) Or (frmHardcoreCheats.chkTotalWaste.value = True)) Then 'And (TibiaVersionLong >= 773)) Then
    SpecialSource = True
   Else
    SpecialSource = False
    fRes = SearchItem(idConnection, runeB1, runeB2)  'search thing


  If fRes.foundCount = 0 Then
     aRes = SendSystemMessageToClient(idConnection, "can't find " & thing & ", open new bp of " & thing & "!")
     DoEvents
     SendMobAimbot = 0
     Exit Function
  End If
  End If
  If (TibiaVersionLong < 760) Then
    myS = MyStackPos(idConnection)
  Else
    myS = FirstPersonStackPos(idConnection)
  End If
  ' search yourself
  If myS = &HFF Then
    aRes = SendLogSystemMessageToClient(idConnection, "Your map is out of sync, can't use " & thing & "!")
    SendMobAimbot = 0
    Exit Function
  End If
  
  If Target = "" Then ' use last targeted
    If SpecialSource = False Then
            aRes = SendSystemMessageToClient(idConnection, CStr(fRes.foundCount) & " " & thing & " found - Casting one from bp ID " & _
           CStr(CLng(fRes.bpID)) & " slot " & CStr(CLng(fRes.slotID)))
         ' GetHexStrFromPosition(myX(idConnection)
          sCheat = "84 FF FF " & GoodHex(&H40 + fRes.bpID) & " 00 " & _
               GoodHex(fRes.slotID) & " " & GoodHex(runeB1) & " " & GoodHex(runeB2) & " " & _
               GoodHex(fRes.slotID) & " " & SpaceID(currTargetID(idConnection))
         
          SafeCastCheatString "SendMobAimbot1", idConnection, sCheat
          Exit Function
    Else
          sCheat = "84 FF FF 00 00 00 " & GoodHex(runeB1) & " " & GoodHex(runeB2) & " 00 " & SpaceID(currTargetID(idConnection))
        
          SafeCastCheatString "SendMobAimbot2", idConnection, sCheat
    End If
  End If
  ' search the creature
  lTarget = LCase(Target)
  For y = -6 To 7
    For x = -8 To 9
      For s = 0 To 10
        tmpID = Matrix(y, x, myZ(idConnection), idConnection).s(s).dblID
        If tmpID = 0 Then
          lSquare = ""
        Else
          lSquare = LCase(GetNameFromID(idConnection, tmpID))
        End If
        If lSquare = lTarget Then
            If SpecialSource = False Then
        '0D 00 84 FF FF 40 00 00 40 0C 00 CB DD 01 40
              aRes = SendSystemMessageToClient(idConnection, CStr(fRes.foundCount) & " " & thing & " found - Casting one from bp ID " & _
              CStr(CLng(fRes.bpID)) & " slot " & CStr(CLng(fRes.slotID)))
        
              sCheat = "84 FF FF " & GoodHex(&H40 + fRes.bpID) & " 00 " & _
                  GoodHex(fRes.slotID) & " " & GoodHex(runeB1) & " " & GoodHex(runeB2) & " " & _
                  GoodHex(fRes.slotID) & " " & SpaceID(Matrix(y, x, myZ(idConnection), idConnection).s(s).dblID)
            
              SafeCastCheatString "SendMobAimbot3", idConnection, sCheat
              Exit Function
          Else
             sCheat = "84 FF FF 00 00 00 " & GoodHex(runeB1) & " " & GoodHex(runeB2) & " 00 " & SpaceID(Matrix(y, x, myZ(idConnection), idConnection).s(s).dblID)
             SafeCastCheatString "SendMobAimbot4", idConnection, sCheat
             SendMobAimbot = 0
            Exit Function
          End If
        End If
      Next s
    Next x
  Next y
  aRes = SendSystemMessageToClient(idConnection, "Sorry, " & Target & " is not on BlackdProxy track")
  SendMobAimbot = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at SendMobAimbot #"
  frmMain.DoCloseActions idConnection
  DoEvents
  SendMobAimbot = -1
End Function

Public Function SpaceID(dblID As Double) As String
Dim dbl1 As Byte
Dim dbl2 As Byte
Dim dbl3 As Byte
Dim dbl4 As Byte
Dim tmpID As Long
Dim res As String
tmpID = dblID
dbl1 = dblID \ 16777216
dblID = dblID - (CLng(dbl1) * 16777216)
dbl2 = dblID \ 65536
dblID = dblID - (CLng(dbl2) * 65536)
dbl3 = dblID \ 256
dblID = dblID - (CLng(dbl3) * 256)
dbl4 = dblID
res = GoodHex(dbl4) & " " & GoodHex(dbl3) & " " & GoodHex(dbl2) & " " & GoodHex(dbl1)
dblID = tmpID
SpaceID = res
End Function



Public Function CastSpell(idConnection As Integer, spellString As String) As Long
  Dim cPacket() As Byte
  Dim lonS As Long
  Dim totalL As Long
  Dim limL As Long
  Dim i As Long
  Dim j As Long
  Dim aRes As Long
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  If publicDebugMode = True Then
    aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Casting spell : " & spellString)
    DoEvents
  End If
  lonS = Len(spellString)
  limL = 5 + lonS
  ReDim cPacket(limL)
  totalL = 4 + lonS
  cPacket(0) = LowByteOfLong(totalL)
  cPacket(1) = HighByteOfLong(totalL)
  cPacket(2) = &H96
  cPacket(3) = &H1
  cPacket(4) = LowByteOfLong(lonS)
  cPacket(5) = HighByteOfLong(lonS)
  j = 1
  For i = 6 To limL
    cPacket(i) = CByte(Asc(Mid(spellString, j, 1)))
    j = j + 1
  Next i
  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & ">" & frmMain.showAsStr2(cpacket, True)
  frmMain.UnifiedSendToServerGame idConnection, cPacket, True
  DoEvents
  CastSpell = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at CastSpell #"
  frmMain.DoCloseActions idConnection
  DoEvents
  CastSpell = -1
End Function

Public Function MoveItemToRightHand(idConnection As Integer, b1 As Byte, b2 As Byte, b3 As Byte, bpIDpar As Byte, slotIDpar As Byte, checkAmmo As Boolean) As Long
  '0F 00 78 FF FF 40 00 06 0D 0C 06 FF FF 05 00 00 01
  Dim i As Long
  Dim j As Long
  Dim cPacket() As Byte
  Dim aRes As Long
  Dim bpID As Byte
  Dim slotID As Byte
  Dim resF As TypeSearchItemResult2
  Dim sCheat As String
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  If ((b1 = 0) And (b2 = 0)) Then
    MoveItemToRightHand = -1
    Exit Function
  End If
  If Not ((mySlot(idConnection, SLOT_RIGHTHAND).t1 = 0) And (mySlot(idConnection, SLOT_RIGHTHAND).t2 = 0)) Then
    MoveItemToRightHand = -1
    Exit Function
  End If
  bpID = bpIDpar
  slotID = slotIDpar
  If b1 = &H0 And b2 = &H0 Then
    MoveItemToRightHand = 0
    Exit Function
  End If
  If checkAmmo = True Then
    If mySlot(idConnection, SLOT_AMMUNITION).t1 = b1 And mySlot(idConnection, SLOT_AMMUNITION).t2 = b2 And mySlot(idConnection, SLOT_AMMUNITION).t3 >= b3 Then
      If publicDebugMode = True Then
        aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Moving item " & GoodHex(b1) & " " & GoodHex(b2) & " to right hand from ammo")
        DoEvents
      End If
      sCheat = "78 FF FF 0A 00 00 " & GoodHex(b1) & " " & GoodHex(b2) & " 00 FF FF 05 00 00 "
'      ReDim cPacket(16)
'      cPacket(0) = LowByteOfLong(15)
'      cPacket(1) = HighByteOfLong(15)
'      cPacket(2) = &H78
'      cPacket(3) = &HFF
'      cPacket(4) = &HFF
'      cPacket(5) = &HA ' from ammo
'      cPacket(6) = &H0
'      cPacket(7) = &H0
'      cPacket(8) = b1
'      cPacket(9) = b2
'      cPacket(10) = &H0
'      cPacket(11) = &HFF
'      cPacket(12) = &HFF
'      cPacket(13) = &H5 'right hand
'      cPacket(14) = &H0
'      cPacket(15) = &H0
'      If TibiaVersionLong >= 860 Then
'        If b3 = &H0 Then
'           cPacket(16) = &H1
'        Else
'           cPacket(16) = b3 ' amount
'        End If
'      Else
'        cPacket(16) = b3 ' amount
'      End If
      If TibiaVersionLong >= 860 Then
        If b3 = &H0 Then
          sCheat = sCheat & " 01"
        Else
          sCheat = sCheat & " " & GoodHex(b3)
        End If
      Else
        sCheat = sCheat & " " & GoodHex(b3)
      End If
       
      'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & ">" & frmMain.showAsStr2(cpacket, True)
      'frmMain.UnifiedSendToServerGame idConnection, cPacket, True
      'DoEvents

      SafeCastCheatString "MoveItemToRightHand1", idConnection, sCheat
      MoveItemToRightHand = 0
      Exit Function
    End If
    resF = SearchItemWithAmount(idConnection, b1, b2, b3)
    If resF.foundCount = 0 Then
      MoveItemToRightHand = 0
      Exit Function
    Else
      bpID = resF.bpID
      slotID = resF.slotID
    End If
  End If
  If publicDebugMode = True Then
    aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Moving item " & GoodHex(b1) & " " & GoodHex(b2) & " to right hand from <bpID " & GoodHex(bpID) & " slotID" & GoodHex(bpID) & ">")
    DoEvents
  End If
  If (bpID = &HFF) Then
    aRes = SendLogSystemMessageToClient(idConnection, "BlackdProxy: Out of blank runes")
    DoEvents
    MoveItemToRightHand = -1
    Exit Function
  End If
  sCheat = "78 FF FF " & GoodHex(&H40 + bpID) & " 00 " & GoodHex(slotID) & " " & GoodHex(b1) & " " & GoodHex(b2) & " " & _
   GoodHex(slotID) & " FF FF 05 00 00"
'  ReDim cPacket(16)
'  cPacket(0) = LowByteOfLong(15)
'  cPacket(1) = HighByteOfLong(15)
'  cPacket(2) = &H78
'  cPacket(3) = &HFF
'  cPacket(4) = &HFF
'  cPacket(5) = &H40 + bpID
'  cPacket(6) = &H0
'  cPacket(7) = slotID
'  cPacket(8) = b1
'  cPacket(9) = b2
'  cPacket(10) = slotID
'  cPacket(11) = &HFF
'  cPacket(12) = &HFF
'  cPacket(13) = &H5 'right hand
'  cPacket(14) = &H0
'  cPacket(15) = &H0
'  cPacket(16) = b3 ' amount
'  'If ((TibiaVersionLong >= 860) And (cPacket(16) = 0)) Then
'  If (cPacket(16) = 0) Then
'    cPacket(16) = &H1
'  End If
'  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & ">" & frmMain.showAsStr2(cpacket, True)
'  Debug.Print frmMain.showAsStr2(cPacket, True)
'  frmMain.UnifiedSendToServerGame idConnection, cPacket, True
'  DoEvents
      If TibiaVersionLong >= 860 Then
        If b3 = &H0 Then
          sCheat = sCheat & " 01"
        Else
          sCheat = sCheat & " " & GoodHex(b3)
        End If
      Else
        sCheat = sCheat & " " & GoodHex(b3)
      End If
  SafeCastCheatString "MoveItemToRightHand2", idConnection, sCheat
  
  MoveItemToRightHand = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at MoveItemToRightHand(" & _
   CStr(idConnection) & "," & GoodHex(b1) & "," & GoodHex(b2) & "," & _
   GoodHex(b3) & "," & GoodHex(bpIDpar) & "," & GoodHex(slotIDpar) & "," & BooleanAsStr(checkAmmo) & ")"
  LogOnFile "errors.txt", "WARNING: " & Err.Description & " at MoveItemToRightHand(" & _
   CStr(idConnection) & "," & GoodHex(b1) & "," & GoodHex(b2) & "," & _
   GoodHex(b3) & "," & GoodHex(bpIDpar) & "," & GoodHex(slotIDpar) & "," & BooleanAsStr(checkAmmo) & ")"
  'frmMain.DoCloseActions idconnection
  'DoEvents
  MoveItemToRightHand = -1
End Function

Public Function MoveItemToLeftHand(idConnection As Integer, b1 As Byte, b2 As Byte, b3 As Byte, bpIDpar As Byte, slotIDpar As Byte, checkAmmo As Boolean) As Long
  '0F 00 78 FF FF 40 00 06 0D 0C 06 FF FF 05 00 00 01
  Dim i As Long
  Dim j As Long
  Dim cPacket() As Byte
  Dim aRes As Long
  Dim bpID As Byte
  Dim slotID As Byte
  Dim sCheat As String
  Dim resF As TypeSearchItemResult2
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  If ((b1 = 0) And (b2 = 0)) Then
    MoveItemToLeftHand = -1
    Exit Function
  End If
  If Not ((mySlot(idConnection, SLOT_LEFTHAND).t1 = 0) And (mySlot(idConnection, SLOT_LEFTHAND).t2 = 0)) Then
    MoveItemToLeftHand = -1
    Exit Function
  End If
  bpID = bpIDpar
  slotID = slotIDpar
  If b1 = &H0 And b2 = &H0 Then
    MoveItemToLeftHand = 0
    Exit Function
  End If
  If checkAmmo = True Then
    If mySlot(idConnection, SLOT_AMMUNITION).t1 = b1 And mySlot(idConnection, SLOT_AMMUNITION).t2 = b2 And mySlot(idConnection, SLOT_AMMUNITION).t3 >= b3 Then
      If publicDebugMode = True Then
        aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Moving item " & GoodHex(b1) & " " & GoodHex(b2) & " to left hand from ammo")
        DoEvents
      End If
      sCheat = "78 FF FF 0A 00 00 " & GoodHex(b1) & " " & GoodHex(b2) & " 00 FF FF 06 00 00 "

      If TibiaVersionLong >= 860 Then
        If b3 = &H0 Then
          sCheat = sCheat & " 01"
        Else
          sCheat = sCheat & " " & GoodHex(b3)
        End If
      Else
        sCheat = sCheat & " " & GoodHex(b3)
      End If
      SafeCastCheatString "MoveItemToLeftHand1", idConnection, sCheat
      
      MoveItemToLeftHand = 0
      Exit Function
    End If
    resF = SearchItemWithAmount(idConnection, b1, b2, b3)
    If resF.foundCount = 0 Then
      MoveItemToLeftHand = 0
      Exit Function
    Else
      bpID = resF.bpID
      slotID = resF.slotID
    End If
  End If
  
  
  If publicDebugMode = True Then
    aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Moving item " & GoodHex(b1) & " " & GoodHex(b2) & " to left hand from <bpID " & GoodHex(bpID) & " slotID" & GoodHex(bpID) & ">")
    DoEvents
  End If
  If (bpID = &HFF) Then
    aRes = SendLogSystemMessageToClient(idConnection, "BlackdProxy: Out of blank runes")
    DoEvents
    MoveItemToLeftHand = -1
    Exit Function
  End If
  sCheat = "78 FF FF " & GoodHex(&H40 + bpID) & " 00 " & GoodHex(slotID) & " " & GoodHex(b1) & " " & GoodHex(b2) & " " & _
   GoodHex(slotID) & " FF FF 06 00 00"

      If TibiaVersionLong >= 860 Then
        If b3 = &H0 Then
          sCheat = sCheat & " 01"
        Else
          sCheat = sCheat & " " & GoodHex(b3)
        End If
      Else
        sCheat = sCheat & " " & GoodHex(b3)
      End If
  SafeCastCheatString "MoveItemToLeftHand2", idConnection, sCheat
  MoveItemToLeftHand = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at MoveItemToLeftHand(" & _
   CStr(idConnection) & "," & GoodHex(b1) & "," & GoodHex(b2) & "," & _
   GoodHex(b3) & "," & GoodHex(bpIDpar) & "," & GoodHex(slotIDpar) & "," & BooleanAsStr(checkAmmo) & ")"
  LogOnFile "errors.txt", "WARNING: " & Err.Description & " at MoveItemToLeftHand(" & _
   CStr(idConnection) & "," & GoodHex(b1) & "," & GoodHex(b2) & "," & _
   GoodHex(b3) & "," & GoodHex(bpIDpar) & "," & GoodHex(slotIDpar) & "," & BooleanAsStr(checkAmmo) & ")"
  'frmMain.DoCloseActions idconnection
  'DoEvents
  MoveItemToLeftHand = -1
End Function

Public Function EatFood(idConnection As Integer, b1 As Byte, b2 As Byte, bpID As Byte, slotID As Byte, Optional ByVal onlyFood As Boolean = True) As Long
 ' 0A 00 82 FF FF 40 00 01 C0 0D 01 00
 Dim i As Long
  Dim j As Long
  Dim cPacket() As Byte
  Dim aRes As Long
  Dim tileID As Long
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  If ((Backpack(idConnection, bpID).item(slotID).t1 = b1) And (Backpack(idConnection, bpID).item(slotID).t2 = b2)) Then
    tileID = GetTheLong(b1, b2)
    If onlyFood = True Then
        If DatTiles(tileID).isFood = False Then ' not food
            If publicDebugMode = True Then
              aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Not food. Safe mode blocked eat food attempt from <bpID " & GoodHex(bpID) & " slotID" & GoodHex(bpID) & ">")
              DoEvents
            End If
            EatFood = 0
            Exit Function
        End If
    End If
  Else ' food no longer there
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Food no longer there. Safe mode blocked eat food attempt from <bpID " & GoodHex(bpID) & " slotID" & GoodHex(bpID) & ">")
      DoEvents
    End If
    EatFood = 0
    Exit Function
  End If
  If publicDebugMode = True Then
    aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Eating food " & GoodHex(b1) & " " & GoodHex(b2) & " from <bpID " & GoodHex(bpID) & " slotID" & GoodHex(bpID) & ">")
    DoEvents
  End If
  ReDim cPacket(11)
  cPacket(0) = LowByteOfLong(10)
  cPacket(1) = HighByteOfLong(10)
  cPacket(2) = &H82
  cPacket(3) = &HFF
  cPacket(4) = &HFF
  cPacket(5) = &H40 + bpID
  cPacket(6) = &H0
  cPacket(7) = slotID
  cPacket(8) = b1
  cPacket(9) = b2
  cPacket(10) = slotID
  cPacket(11) = &H0 ' action =eat?
  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & ">" & frmMain.showAsStr2(cpacket, True)
  frmMain.UnifiedSendToServerGame idConnection, cPacket, True
  DoEvents
  EatFood = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at EatFood #"
  frmMain.DoCloseActions idConnection
  DoEvents
  EatFood = -1
End Function
Public Sub UpdateExpVars(idConnection As Integer)
  Dim NextLevel As Double
  Dim Level As Double
  Dim Experience As Double
  Dim ExpH As Double
  Dim TimeLoged As Double
  Dim Tsec As Double
  Dim Tmin As Double
  Dim Thour As Double
  Dim Ttemp As Double
  Dim AverageS As Double
  Dim EstimatedLeft As Double
  Dim Esec As Double
  Dim Emin As Double
  Dim Ehour As Double
  Dim IniExp As Double
  Dim ExpGain As Double
  Dim substr As String
  Dim aRes As Long
  Dim sRes As String
  Dim dblSol As Double
  Dim dblSol2 As Double
  On Error Resume Next
  sRes = ""
  var_expleft(idConnection) = "?"
  var_nextlevel(idConnection) = "?"
  var_exph(idConnection) = "?"
  var_timeleft(idConnection) = "?"
  var_played(idConnection) = "?"
  var_expgained(idConnection) = "?"
  IniExp = CDbl(myInitialExp(idConnection)) ' your exp when you made the login
  Level = CDbl(myLevel(idConnection) + 1) ' your current level + 1
  Experience = CDbl(myExp(idConnection)) ' your current experience
  ExpGain = Experience - IniExp ' your exp gain
  ' HERE IS THE MAIN FORMULA OF TIBIA EXP :
  NextLevel = Round(50 * Level * (Level * (Level / 3 - 2) + 17 / 3) - 200 - Experience)
  ' Now we get the seconds since las login time :
  TimeLoged = CDbl((GetTickCount() - myInitialTickCount(idConnection)) / 1000)
  ' Lets translate it to hours : minutes : seconds ...
  Thour = TimeLoged \ 3600
  Ttemp = TimeLoged - (Thour * 3600)
  Tmin = Ttemp \ 60
  Ttemp = Ttemp - (Tmin * 60)
  Tsec = Round(Ttemp)
  ' From last 2 things we also can get the average exp/h in this session :
  dblSol = TimeLoged / 3600
  If dblSol > 0 Then
    AverageS = Round((Experience - IniExp) / dblSol)
  Else
    AverageS = 0
  End If
  ' And also lets see the estimated time left to level up :
  If AverageS = 0 Then
    ' Avoid X\0
    EstimatedLeft = 0
  Else
    dblSol2 = (AverageS / 3600)
    If dblSol2 = 0 Then
        EstimatedLeft = 0
    Else
        EstimatedLeft = Round(NextLevel / (AverageS / 3600))
    End If
  End If
  ' Lets translate it to hours : minutes : seconds ...
  Ehour = EstimatedLeft \ 3600
  Ttemp = EstimatedLeft - (Ehour * 3600)
  Emin = Ttemp \ 60
  Ttemp = Ttemp - (Emin * 60)
  Esec = Round(Ttemp)
  If EstimatedLeft = 0 Then
    ' if you never get exp , then you never will level up!
    substr = "<NEVER>"
  Else
    ' all ok, show results
    substr = Ehour & "h " & Emin & "m " & Esec & "s"
  End If
  var_expleft(idConnection) = CStr(NextLevel)
  var_nextlevel(idConnection) = CStr(Level)
  var_exph(idConnection) = AverageS
  var_timeleft(idConnection) = substr
  var_played(idConnection) = Thour & "h " & Tmin & "m " & Tsec & "s"
  var_expgained(idConnection) = CStr(ExpGain)
  Exit Sub
goterr:
  substr = ""
End Sub


Public Function parseVars(idConnection As Integer, formatStr As String) As String
Dim pos As Long
Dim lastp As Long
Dim resP As String
Dim cc As String
Dim varn As String
Dim validvar As Boolean
Dim eUpdated As Boolean
Dim theTranslation As String

Dim tmpStr As String
Dim currLine As String
Dim tmpbln As Boolean
Dim currLineNumber As Long
Dim bracketCount As Long
Dim firstBracket As Long
Dim part2 As String
Dim part3 As String
Dim varnSwitch As String
Dim strEspecial As String

Dim OriginalformatStr As String
' first deal with bracket content, recursively { ... }

OriginalformatStr = formatStr

eUpdated = False
resP = ""
lastp = Len(formatStr)
pos = 1
bracketCount = 0
While pos <= lastp
  cc = Mid(formatStr, pos, 1)
  If cc = "{" Then
    bracketCount = bracketCount + 1
    If bracketCount = 1 Then
      firstBracket = pos
    End If
  End If
  If cc = "}" Then
    bracketCount = bracketCount - 1
    If (bracketCount = 0) Then
      part2 = Mid$(OriginalformatStr, firstBracket + 1, pos - firstBracket - 1)
      part3 = Right$(OriginalformatStr, Len(formatStr) - pos)
      parseVars = parseVars(idConnection, resP & parseVars(idConnection, part2) & part3)
      Exit Function
    End If
  End If
  If (bracketCount = 0) Then
    resP = resP & cc
  End If
  pos = pos + 1
Wend
' then deal with variables $something$
pos = 1
resP = ""
varn = ""
While pos <= lastp
  cc = Mid(formatStr, pos, 1)
  If (cc = "$") Then
    varn = ""
    pos = pos + 1
    validvar = False
    While (pos <= lastp) And (validvar = False)
      cc = Mid(formatStr, pos, 1)
      If cc = "$" Then
        validvar = True
      Else
        varn = varn & cc
      End If
      pos = pos + 1
    Wend
    If validvar = True Then
    varnSwitch = LCase(varn)
    Select Case varnSwitch
    Case "" ' $$ translates to $
      theTranslation = "$"
    Case "expleft" ' $expleft$ = exp left to next level
      If (eUpdated = False) Then
        UpdateExpVars idConnection
        eUpdated = True
      End If
      theTranslation = var_expleft(idConnection)
    Case "nextlevel" ' $nextlevel$ = your level + 1
      If (eUpdated = False) Then
        UpdateExpVars idConnection
        eUpdated = True
      End If
      theTranslation = var_nextlevel(idConnection)
    Case "exph" ' $exph$ = your exp/h
      If (eUpdated = False) Then
        UpdateExpVars idConnection
        eUpdated = True
      End If
      theTranslation = var_exph(idConnection)
    Case "timeleft" ' $timeleft$ = estimated time left for level advance
      If (eUpdated = False) Then
        UpdateExpVars idConnection
        eUpdated = True
      End If
      theTranslation = var_timeleft(idConnection)
    Case "played" ' $played$ = time played this session
      If (eUpdated = False) Then
        UpdateExpVars idConnection
        eUpdated = True
      End If
      theTranslation = var_played(idConnection)
    Case "expgained" ' $expgained$ = exp gained this session
      If (eUpdated = False) Then
        UpdateExpVars idConnection
        eUpdated = True
      End If
      theTranslation = var_expgained(idConnection)
    Case "charactername" ' $charactername$ = your char name
      theTranslation = CharacterName(idConnection)
    Case "broadcast" ' $broadcast$ = current broadcast destination
      If ((currentBroadcastIndex > -1) And (currentBroadcastIndex < frmBroadcast.lstList.ListCount)) Then
         theTranslation = frmBroadcast.lstList.List(currentBroadcastIndex)
      Else
        theTranslation = "-nobody"
      End If
    Case "lastcheckresult"
      theTranslation = lastIngameCheck(idConnection)
    Case "lastchecktileid"
      theTranslation = lastIngameCheckTileID(idConnection)
    Case "lastsender" ' $lastsender$ = your last sender
      theTranslation = var_lastsender(idConnection)
    Case "lastmsg" ' $lastmsg$ = your last msg
      theTranslation = FixLastMSG(idConnection)
    Case "lf" ' $lf$ = line feed character (if possible)
      theTranslation = var_lf(idConnection)
    Case "myhp"
      theTranslation = CStr(myHP(idConnection))
    
    Case "myhppercent"
      theTranslation = CStr(Round((myHP(idConnection) / myMaxHP(idConnection)) * 100))
    Case "mymanapercent"
        If myMaxMana(idConnection) = 0 Then
            theTranslation = "100"
        Else
            theTranslation = CStr(Round((myMana(idConnection) / myMaxMana(idConnection)) * 100))
       End If
    Case "mymana"
      theTranslation = CStr(myMana(idConnection))
    Case "mycap"
      theTranslation = CStr(myCap(idConnection))
    Case "mystamina"
      theTranslation = CStr(myStamina(idConnection))
    Case "mylevel"
      theTranslation = CStr(myLevel(idConnection))
    Case "mysoulpoints"
      theTranslation = CStr(mySoulpoints(idConnection))
    Case "myexp"
      theTranslation = CStr(myExp(idConnection))
    Case "lastpkname"
      theTranslation = DangerPKname(idConnection)
    Case "lastgmname"
      theTranslation = DangerGMname(idConnection)
    Case "lasthpchange"
      theTranslation = CStr(lastHPchange(idConnection))
    Case "date"
      theTranslation = CStr(Format(Date, "dd/mm/yyyy"))
    Case "time"
      theTranslation = CStr(Format(Time, "hh:mm:ss"))
    Case "shorttime"
      theTranslation = CStr(Format(Time, "hh:mm"))
    Case "hex-myid"
      theTranslation = IDofName(idConnection, "", 1)
    Case "hex-lastattackedid"
      theTranslation = IDofName(idConnection, "", 2)
    Case "cavebottimewithsametarget"
      'If cavebotEnabled(idConnection) = True Then
     '   theTranslation = 0
     ' Else
        theTranslation = CStr(GetTickCount() - CavebotTimeWithSameTarget(idConnection))
     ' End If
    Case "lastattackedid"
      theTranslation = CStr(lastAttackedID(idConnection))
    Case "bestenemy"
      theTranslation = TellBestEnemy(idConnection)
    Case "bestenemyhp"
      theTranslation = CStr(TellBestEnemyHP(idConnection))
    Case "bestenemyid"
      theTranslation = CStr(TellBestEnemyID(idConnection))
    Case "hex-bestenemyid"
      theTranslation = DoubleToStr(CStr(TellBestEnemyID(idConnection)))
    Case "myx"
      theTranslation = CStr(myX(idConnection))
    Case "myy"
      theTranslation = CStr(myY(idConnection))
    Case "myz"
      theTranslation = CStr(myZ(idConnection))
    Case "comboorder"
      theTranslation = CStr(frmHardcoreCheats.txtOrder.Text)
    Case "comboleader"
      theTranslation = CStr(frmHardcoreCheats.txtRemoteLeader.Text)
    Case "lastusedchannelid"
      theTranslation = lastUsedChannelID(idConnection)
    Case "lastrecchannelid"
      theTranslation = lastRecChannelID(idConnection)
    Case "hex-currenttargetid"
      theTranslation = SpaceID(ReadRedSquare(idConnection))
    Case Else ' user variables
      If Left(varn, 1) = "_" Then
        strEspecial = "(" & CStr(idConnection) & ")" ' local variables use internal prefix (idconnection)
        If Len(varn) > 1 Then
          If Mid(varn, 2, 1) = "_" Then
            strEspecial = "" ' global variables have no internal prefix
          End If
        End If
        theTranslation = GetUserVar(varn & strEspecial)
      ' predefined "functions"
      ElseIf Len(varn) > 4 Then
        If (Left$(varn, 6) = "check:") Then
          theTranslation = CStr(CheckVariableCondition(Right$(varn, (Len(varn) - 6))))
        ElseIf (Left$(varn, 7) = "istrue:") Then
          theTranslation = CStr(CheckIsTrue(Right$(varn, (Len(varn) - 7))))
        ElseIf (Left$(varn, 11) = "countitems:") Then
          theTranslation = CStr(CountTheItemsForUser(idConnection, Right$(varn, (Len(varn) - 11))))
        ElseIf (Left$(varn, 11) = "hpofhex-id:") Then
          theTranslation = CStr(SafeGiveIDinfo(idConnection, Right$(varn, (Len(varn) - 11)), 1)) 'give HP
        ElseIf (Left$(varn, 12) = "dirofhex-id:") Then
          theTranslation = CStr(SafeGiveIDinfo(idConnection, Right$(varn, (Len(varn) - 12)), 2)) 'give Direction
        ElseIf (Left$(varn, 13) = "numbertohex1:") Then
          theTranslation = HexConverter(Right$(varn, (Len(varn) - 13)), 1)
        ElseIf (Left$(varn, 13) = "numbertohex2:") Then
          theTranslation = HexConverter(Right$(varn, (Len(varn) - 13)), 2)
        ElseIf (Left$(varn, 13) = "hex1tonumber:") Then
          theTranslation = HexConverter(Right$(varn, (Len(varn) - 13)), 3)
        ElseIf (Left$(varn, 13) = "hex2tonumber:") Then
          theTranslation = HexConverter(Right$(varn, (Len(varn) - 13)), 4)
        ElseIf (Left$(varn, 13) = "randomnumber:") Then
          theTranslation = ProcessVarRandom(Right$(varn, (Len(varn) - 13)))
        ElseIf (Left$(varn, 13) = "numericalexp:") Then
          theTranslation = CStr(NumericValueOfExpresion(Right$(varn, (Len(varn) - 13))))
        ElseIf (Left$(varn, 10) = "statusbit:") Then
          theTranslation = GetStatusBit(idConnection, Right$(varn, (Len(varn) - 10)))
        ElseIf (Left$(varn, 17) = "hex-equiped-item:") Then
          theTranslation = equipmentInfo(idConnection, 1, Right$(varn, (Len(varn) - 17)))
        ElseIf (Left$(varn, 20) = "hex-equiped-ammount:") Then
          theTranslation = equipmentInfo(idConnection, 3, Right$(varn, (Len(varn) - 20)))
        ElseIf (Left$(varn, 28) = "hex-equiped-ammount-special:") Then
          theTranslation = equipmentInfo(idConnection, 2, Right$(varn, (Len(varn) - 28)))
        ElseIf (Left$(varn, 20) = "num-equiped-ammount:") Then
          theTranslation = equipmentInfo(idConnection, 4, Right$(varn, (Len(varn) - 20)))
        ElseIf (Left$(varn, 13) = "hex-idofname:") Then
          theTranslation = IDofName(idConnection, Right$(varn, (Len(varn) - 13)), 0)
        ElseIf (Left$(varn, 13) = "nameofhex-id:") Then
          theTranslation = NameOfHexID(idConnection, Right$(varn, (Len(varn) - 13)))
        ElseIf (Left$(varn, 10) = "urlencode:") Then
          theTranslation = URLEncode(Right$(varn, (Len(varn) - 10)))
        ElseIf (Left$(varn, 13) = "hex-tibiastr:") Then
          theTranslation = Hexarize2(Right$(varn, (Len(varn) - 13)))
        ElseIf (Left$(varn, 8) = "httpget:") Then
          theTranslation = frmMain.HTTPGet(Right$(varn, (Len(varn) - 8)))
        ElseIf (Left$(varn, 13) = "randomlineof:") Then 'special variable
          theTranslation = GetRandomLineOf(Right$(varn, (Len(varn) - 13)))
        ElseIf (Left$(varn, 19) = "pksonrelativefloor:") Then 'special variable
          theTranslation = CStr(CountOnFloor(idConnection, Right$(varn, (Len(varn) - 19)), True, False))
        ElseIf (Left$(varn, 19) = "gmsonrelativefloor:") Then 'special variable
          theTranslation = CStr(CountOnFloor(idConnection, Right$(varn, (Len(varn) - 19)), False, True))
        ElseIf (Left$(varn, 25) = "pksandgmsonrelativefloor:") Then 'special variable
          theTranslation = CStr(CountOnFloor(idConnection, Right$(varn, (Len(varn) - 25)), True, True))
        ElseIf (Left$(varn, 28) = "meleetargetsonrelativefloor:") Then 'special variable
          theTranslation = CStr(CountOnFloor(idConnection, Right$(varn, (Len(varn) - 28)), False, False, True))
        ElseIf (Left$(varn, 18) = "useitemwithamount:") Then 'HHBCODE
          theTranslation = CStr(UseItemWithAmount(idConnection, Right$(varn, (Len(varn) - 18))))
        ElseIf (Left$(varn, 13) = "nlineoflabel:") Then 'special variable
          tmpStr = LCase(Right$(varn, (Len(varn) - 12)))
          tmpbln = False
          currLineNumber = 0
          Do
            currLine = LCase(GetStringFromIDLine(idConnection, currLineNumber))
            If currLine = tmpStr Then
              theTranslation = CStr(currLineNumber)
              tmpbln = True
              Exit Do
            End If
            currLineNumber = currLineNumber + 1
          Loop Until ((currLine = "") Or (currLine = "?"))
          If tmpbln = False Then
            theTranslation = "1000000"
          End If
        Else
          theTranslation = "<" & varn & ">"
        End If
      Else
        theTranslation = "<" & varn & ">"
      End If
    End Select
    Else
      theTranslation = "<" & varn & " ... ?"
    End If
     resP = resP & theTranslation
  Else
    resP = resP & cc
    pos = pos + 1
  End If
Wend
parseVars = resP
End Function
Public Sub initInitialPacket(idConnection As Integer)
    ReDim mustSendFirstWhenConnected(idConnection).packet(0)
    mustSendFirstWhenConnected(idConnection).mustSend = False
End Sub
Public Function GiveExpInfo(idConnection As Integer, formatStr As String) As Long
  Dim aRes As Long
  Dim theMessage As String
  Dim mycolorm As Byte
  Dim leftp As String
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  'mycolorm = &H13 ' white
  leftp = Left$(frmHardcoreCheats.cmbWhere.Text, 2)
  mycolorm = CByte(CLng(leftp))
  ' mycolorm = &H16 ' green
  ' UpdateExpVars idConnection
  If mycolorm = &H14 Then ' clasic
    var_lf(idConnection) = " "
  Else
    var_lf(idConnection) = vbLf
  End If
  theMessage = parseVars(idConnection, formatStr)
  aRes = SendCustomSystemMessageToClient(idConnection, theMessage, mycolorm)
  DoEvents
  GiveExpInfo = 0
  Exit Function
goterr:
  GiveExpInfo = -1
End Function

Public Function StartPush(idConnection As Integer, Target As String) As Long
  Dim x As Long
  Dim y As Long
  Dim s As Long
  Dim tmpName As String
  Dim tmpID As Double
  Dim lname As String
  Dim aRes As Long
  If Target = "" Then
    RemoveSpamOrder idConnection, 2 'remove auto push
    aRes = SendLogSystemMessageToClient(idConnection, "Auto pushing DISABLED")
    DoEvents
    pushTarget(idConnection) = 0
    aRes = StopFollowTarget(idConnection)
    DoEvents
    StartPush = 0
    Exit Function
  Else
    lname = LCase(Target)
    pushTarget(idConnection) = 0
    For y = -6 To 7
      For x = -8 To 9
        For s = 0 To 10
          tmpID = Matrix(y, x, myZ(idConnection), idConnection).s(s).dblID
          If tmpID > 0 Then
            tmpName = LCase(GetNameFromID(idConnection, tmpID))
            If tmpName = lname Then
              pushTarget(idConnection) = tmpID
            End If
          End If
        Next s
      Next x
    Next y
    If pushTarget(idConnection) = 0 Then
      RemoveSpamOrder idConnection, 2 'remove auto push
      aRes = StopFollowTarget(idConnection)
      DoEvents
      aRes = SendLogSystemMessageToClient(idConnection, "Auto pushing DISABLED - not found: " & Target)
      DoEvents
    Else
      AddSpamOrder idConnection, 2 'add auto push
      aRes = SendLogSystemMessageToClient(idConnection, "Auto pushing ENABLED : " & Target & " (ID " & CStr(pushTarget(idConnection)) & ")")
      DoEvents
      aRes = followTarget(idConnection, pushTarget(idConnection))
      ' doevents included in followtarget
    End If
  End If
  StartPush = 0
End Function

Public Function StartPush2(idConnection As Integer) As Long
  Dim x As Long
  Dim y As Long
  Dim s As Long
  Dim tmpName As String
  Dim tmpID As Double
  Dim lname As String
  Dim aRes As Long
  Dim Target As String
'endF:
  Target = currTargetName(idConnection)
  If Target = "stoppush" Then
    RemoveSpamOrder idConnection, 2 'remove auto push
    aRes = SendLogSystemMessageToClient(idConnection, "Auto pushing DISABLED")
    DoEvents
    pushTarget(idConnection) = 0
    aRes = StopFollowTarget(idConnection)
    DoEvents
    StartPush2 = 0
    Exit Function
  Else
    lname = LCase(Target)
    pushTarget(idConnection) = 0
    For y = -6 To 7
      For x = -8 To 9
        For s = 0 To 10
          tmpID = Matrix(y, x, myZ(idConnection), idConnection).s(s).dblID
          If tmpID > 0 Then
            tmpName = LCase(GetNameFromID(idConnection, tmpID))
            If tmpName = lname Then
              pushTarget(idConnection) = tmpID
            End If
          End If
        Next s
      Next x
    Next y
    If pushTarget(idConnection) = 0 Then
      RemoveSpamOrder idConnection, 2 'remove auto push
      aRes = StopFollowTarget(idConnection)
      DoEvents
      aRes = SendLogSystemMessageToClient(idConnection, "Auto pushing DISABLED - not found: " & Target)
      DoEvents
    Else
      AddSpamOrder idConnection, 2 'add auto push
      aRes = SendLogSystemMessageToClient(idConnection, "Auto pushing ENABLED : " & Target & " (ID " & CStr(pushTarget(idConnection)) & ")")
      DoEvents
      GoTo endF:
endF:
  Target = currTargetName(idConnection)
      'aRes = followTarget(idConnection, pushTarget(idConnection))
      ' doevents included in followtarget
    End If
  End If
  StartPush2 = 0
End Function


Public Function StopFollowTarget(idConnection As Integer) As Long
  Dim cPacket(6) As Byte
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  cPacket(0) = &H5
  cPacket(1) = &H0
  cPacket(2) = &HA2
  cPacket(3) = &H0
  cPacket(4) = &H0
  cPacket(5) = &H0
  cPacket(6) = &H0
 ' frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & frmMain.showAsStr2(cPacket, True)
  frmMain.UnifiedSendToServerGame idConnection, cPacket, True
  DoEvents
  StopFollowTarget = 0
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at StopFollowTarget #"
  frmMain.DoCloseActions idConnection
  DoEvents
  StopFollowTarget = -1
End Function



Public Function DoTurbo(idConnection As Integer) As Long
  Dim cPacket(3) As Byte
  Dim curDirPlayer As Byte
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  cPacket(0) = &H1
  cPacket(1) = &H0
  ' 00 = north ; 01 = right ; 02 = south ; 03 = left
  curDirPlayer = GetDirectionFromID(idConnection, myID(idConnection))
  Select Case curDirPlayer
  Case &H0
    cPacket(2) = &H65
  Case &H1
    cPacket(2) = &H66
  Case &H2
    cPacket(2) = &H67
  Case &H3
    cPacket(2) = &H68
  End Select
   
 ' frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & frmMain.showAsStr2(cPacket, True)
  frmMain.UnifiedSendToServerGame idConnection, cPacket, True
  DoEvents
  DoTurbo = 0
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at DoTurbo #"
  frmMain.DoCloseActions idConnection
  DoEvents
  DoTurbo = -1
End Function

Public Function followTarget(idConnection As Integer, targetID As Double) As Long
  Dim lTarget As Double
  Dim x As Long
  Dim y As Long
  Dim s As Long
  Dim tmpID As Double
  Dim aRes As Long
  Dim sPacket As String
  Dim cPacket() As Byte
  Dim inRes As Integer
  Dim posibleCount As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  lTarget = targetID
         ' aRes = SendLogSystemMessageToClient(idConnection, "Searching ID: " & CStr(lTarget))
        '  DoEvents

  ' follow ID
  '05 00 A2 EA 20 A4 00
   For y = -6 To 7
    For x = -8 To 9
      For s = 0 To 10
        tmpID = Matrix(y, x, myZ(idConnection), idConnection).s(s).dblID
        If tmpID = lTarget Then
          sPacket = "05 00 A2 " & SpaceID(tmpID)
         ' frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & sPacket
          inRes = GetCheatPacket(cPacket, sPacket)
          frmMain.UnifiedSendToServerGame idConnection, cPacket, True
          DoEvents
          followTarget = 0
          Exit Function
        End If
      Next s
    Next x
  Next y
  followTarget = 0
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at FollowTarget #"
  frmMain.DoCloseActions idConnection
  DoEvents
  followTarget = -1
End Function
Public Function DoPush(idConnection As Integer) As Long
  Dim aRes As Long
  Dim iniX As Long
  Dim iniY As Long
  Dim iniS As Long
  Dim desx As Long
  Dim desy As Long
  Dim optIni As Long
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim tileID As Long
  Dim foundTarget As Boolean
  Dim moveToHere As Long
  Dim posibleCount As Integer
  Dim squareFree As Boolean
  Dim discarded(1 To 8) As Boolean
  Dim randomsDone As Long
  Dim isValid As Boolean
  Dim cPacket(16) As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim b4 As Byte
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  foundTarget = False
  optIni = 0
  For i = 1 To 10
    If Matrix(0, -1, myZ(idConnection), idConnection).s(i).dblID = pushTarget(idConnection) Then
       iniX = -1
       iniY = 0
       iniS = i
       foundTarget = True
       optIni = 1
    End If
  Next i
  
  If foundTarget = False Then
  For i = 1 To 10
    If Matrix(0, 1, myZ(idConnection), idConnection).s(i).dblID = pushTarget(idConnection) Then
       iniX = 1
       iniY = 0
       iniS = i
       foundTarget = True
       optIni = 2
    End If
  Next i
  End If
  
  If foundTarget = False Then
  For i = 1 To 10
    If Matrix(-1, 0, myZ(idConnection), idConnection).s(i).dblID = pushTarget(idConnection) Then
       iniX = 0
       iniY = -1
       iniS = i
       foundTarget = True
       optIni = 3
    End If
  Next i
  End If
  
  If foundTarget = False Then
  For i = 1 To 10
    If Matrix(1, 0, myZ(idConnection), idConnection).s(i).dblID = pushTarget(idConnection) Then
       iniX = 0
       iniY = 1
       iniS = i
       foundTarget = True
       optIni = 4
    End If
  Next i
  End If
  
  If foundTarget = False Then
  For i = 1 To 10
    If Matrix(-1, -1, myZ(idConnection), idConnection).s(i).dblID = pushTarget(idConnection) Then
       iniX = -1
       iniY = -1
       iniS = i
       foundTarget = True
       optIni = 5
    End If
  Next i
  End If
  
  If foundTarget = False Then
  For i = 1 To 10
    If Matrix(1, 1, myZ(idConnection), idConnection).s(i).dblID = pushTarget(idConnection) Then
       iniX = 1
       iniY = 1
       iniS = i
       foundTarget = True
       optIni = 6
    End If
  Next i
  End If
  
  If foundTarget = False Then
  For i = 1 To 10
    If Matrix(-1, 1, myZ(idConnection), idConnection).s(i).dblID = pushTarget(idConnection) Then
       iniX = 1
       iniY = -1
       iniS = i
       foundTarget = True
       optIni = 7
    End If
  Next i
  End If
  
  If foundTarget = False Then
  For i = 1 To 10
    If Matrix(1, -1, myZ(idConnection), idConnection).s(i).dblID = pushTarget(idConnection) Then
       iniX = -1
       iniY = 1
       iniS = i
       foundTarget = True
       optIni = 8
    End If
  Next i
  End If
  
  If optIni = 0 Then 'target not found
    DoPush = 0
    Exit Function
  End If
  
  For i = 1 To 8
    discarded(i) = False
  Next i
  randomsDone = 0
  
reRoll:
  moveToHere = CLng(Int((8 * Rnd) + 1))
  While discarded(moveToHere) = True
    moveToHere = moveToHere + 1
    If moveToHere = 9 Then
      moveToHere = 1
    End If
  Wend
  isValid = True
  Select Case moveToHere
  Case 1
    If optIni = 3 Or optIni = 4 Or optIni = 5 Or optIni = 8 Then
      i = -1
      j = 0
      For k = 0 To 10
        tileID = GetTheLong(Matrix(j, i, myZ(idConnection), idConnection).s(k).t1, _
         Matrix(j, i, myZ(idConnection), idConnection).s(k).t2)
        If tileID = 0 Then
          Exit For
        ElseIf DatTiles(tileID).blocking = True Or tileID = 97 Then
          isValid = False
          Exit For
        End If
      Next k
      If iniX = -i Or iniY = -j Then 'invalid movement
        isValid = False
      End If
    Else
      isValid = False
    End If
  Case 2
    If optIni = 3 Or optIni = 4 Or optIni = 6 Or optIni = 7 Then
      i = 1
      j = 0
      For k = 0 To 10
        tileID = GetTheLong(Matrix(j, i, myZ(idConnection), idConnection).s(k).t1, _
         Matrix(j, i, myZ(idConnection), idConnection).s(k).t2)
        If tileID = 0 Then
          Exit For
        ElseIf DatTiles(tileID).blocking = True Or tileID = 97 Then
          isValid = False
          Exit For
        End If
      Next k
      If iniX = -i Or iniY = -j Then 'invalid movement
        isValid = False
      End If
    Else
      isValid = False
    End If
  Case 3
    If optIni = 1 Or optIni = 2 Or optIni = 5 Or optIni = 7 Then
      i = 0
      j = -1
      For k = 0 To 10
        tileID = GetTheLong(Matrix(j, i, myZ(idConnection), idConnection).s(k).t1, _
         Matrix(j, i, myZ(idConnection), idConnection).s(k).t2)
        If tileID = 0 Then
          Exit For
        ElseIf DatTiles(tileID).blocking = True Or tileID = 97 Then
          isValid = False
          Exit For
        End If
      Next k
      If iniX = -i Or iniY = -j Then 'invalid movement
        isValid = False
      End If
    Else
      isValid = False
    End If
  Case 4
    If optIni = 1 Or optIni = 2 Or optIni = 6 Or optIni = 8 Then
      i = 0
      j = 1
      For k = 0 To 10
        tileID = GetTheLong(Matrix(j, i, myZ(idConnection), idConnection).s(k).t1, _
         Matrix(j, i, myZ(idConnection), idConnection).s(k).t2)
        If tileID = 0 Then
          Exit For
        ElseIf DatTiles(tileID).blocking = True Or tileID = 97 Then
          isValid = False
          Exit For
        End If
      Next k
      If iniX = -i Or iniY = -j Then 'invalid movement
        isValid = False
      End If
    Else
      isValid = False
    End If
  Case 5
    isValid = False
  Case 6
    isValid = False
  Case 7
    isValid = False
  Case 8
    isValid = False
  End Select
  randomsDone = randomsDone + 1
  If isValid = False Then
    discarded(moveToHere) = True
    If randomsDone = 8 Then
     
       DoPush = 0
       Exit Function
    Else
      GoTo reRoll
    End If
  Else ' valid destination
    ' do the move
    b4 = 10
      b1 = Matrix(iniY, iniX, myZ(idConnection), idConnection).s(10).t1
      b2 = Matrix(iniY, iniX, myZ(idConnection), idConnection).s(10).t2
      b3 = Matrix(iniY, iniX, myZ(idConnection), idConnection).s(10).t3
    For k = 1 To 9
    b1 = Matrix(iniY, iniX, myZ(idConnection), idConnection).s(k + 1).t1
    b2 = Matrix(iniY, iniX, myZ(idConnection), idConnection).s(k + 1).t2
    tileID = GetTheLong(b1, b2)
    If (b1 = 0 And b2 = 0) Or (DatTiles(tileID).notMoveable = True) Then
      b1 = Matrix(iniY, iniX, myZ(idConnection), idConnection).s(k).t1
      b2 = Matrix(iniY, iniX, myZ(idConnection), idConnection).s(k).t2
      b3 = Matrix(iniY, iniX, myZ(idConnection), idConnection).s(k).t3
      If b3 = 0 Then
        b3 = 1
      End If
      b4 = k
      Exit For
    End If
    Next k
    If b1 = &H61 And b2 = &H0 Then
      b1 = &H63
    End If

              iniX = myX(idConnection) + iniX
              iniY = myY(idConnection) + iniY
              desx = myX(idConnection) + i
              desy = myY(idConnection) + j
              'SEND HERE
              'aRes = SendLogSystemMessageToClient(idConnection, "Moving from " & iniX & "," & iniY & "," & iniS & " to " & desX & "," & desY)
              'DoEvents


              cPacket(0) = &HF
              cPacket(1) = &H0
              cPacket(2) = &H78
              cPacket(3) = LowByteOfLong(iniX)
              cPacket(4) = HighByteOfLong(iniX)
              cPacket(5) = LowByteOfLong(iniY)
              cPacket(6) = HighByteOfLong(iniY)
              cPacket(7) = CByte(myZ(idConnection))
              cPacket(8) = b1
              cPacket(9) = b2
              ' cPacket(10) = CByte(iniS)
              cPacket(10) = b4 'last stack pos with item
              cPacket(11) = LowByteOfLong(desx)
              cPacket(12) = HighByteOfLong(desx)
              cPacket(13) = LowByteOfLong(desy)
              cPacket(14) = HighByteOfLong(desy)
              cPacket(15) = CByte(myZ(idConnection))
              cPacket(16) = b3
              
              'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & frmMain.showAsStr2(cPacket, True)
              frmMain.UnifiedSendToServerGame idConnection, cPacket, True
              DoEvents
              
              'aRes = SendLogSystemMessageToClient(idConnection, "Moving from " & iniX & "," & iniY & "," & iniS & " to " & desX & "," & desY)
              'DoEvents
              DoPush = 0
              Exit Function
  End If
  

  DoPush = 0
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at DoPush #"
  frmMain.DoCloseActions idConnection
  DoEvents
  DoPush = -1
End Function

Public Function ViewFloor(idConnection As Integer, strFloor As String) As Long
  Dim floor As Long
  Dim aRes As Long
  Dim inRes As Integer
  Dim sCheat As String
  Dim cPacket() As Byte
  Dim tileID As Long
  Dim x As Integer
  Dim y As Integer
  Dim s As Integer
  Dim s2 As Integer
  Dim totalLen As Long
  Dim squareCount As Long
  Dim squareLim As Long
  Dim drawNext As Boolean
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  floor = CLng(strFloor) + myZ(idConnection)
  If floor < 0 Or floor > 15 Then
    ViewFloor = -1
  End If
 ' squareLim = 200
  squareCount = 0
  sCheat = ""
  totalLen = 0
  
  For y = -6 To 7
  For x = -8 To 9
  drawNext = True
  For s = 0 To 10
   tileID = GetTheLong(Matrix(y, x, myZ(idConnection), idConnection).s(s).t1, Matrix(y, x, myZ(idConnection), idConnection).s(s).t2)
   If tileID = 97 Then
     'GoTo drawGot
     tileID = GetTheLong(Matrix(y, x, floor, idConnection).s(0).t1, Matrix(y, x, floor, idConnection).s(0).t2)
     If tileID <> 0 Then
      sCheat = sCheat & " 69 " & FiveChrLon(myX(idConnection) + x) & " " & FiveChrLon(myY(idConnection) + y) & " " & _
       GoodHex(CByte(myZ(idConnection))) & " "
      totalLen = totalLen + 6
   For s2 = 0 To 10
      If s2 = 0 Then
        If DatTiles(tileID).blocking = False Then
          tileID = GetTheLong(Matrix(y, x, floor, idConnection).s(0).t1, Matrix(y, x, floor, idConnection).s(0).t2)
        Else
         tileID = GetTheLong(Matrix(y, x, myZ(idConnection), idConnection).s(0).t1, Matrix(y, x, myZ(idConnection), idConnection).s(0).t2)
        End If
      Else
        tileID = GetTheLong(Matrix(y, x, myZ(idConnection), idConnection).s(s2).t1, Matrix(y, x, myZ(idConnection), idConnection).s(s2).t2)
      End If
      If tileID = 0 Then
        Exit For
      ElseIf tileID = 97 Then
        sCheat = sCheat & "63 00 " & SpaceID(Matrix(y, x, myZ(idConnection), idConnection).s(s2).dblID) & " 02 "
        totalLen = totalLen + 7
      Else
        If DatTiles(tileID).haveExtraByte = True Then
          sCheat = sCheat & FiveChrLon(tileID) & " " & GoodHex(Matrix(y, x, myZ(idConnection), idConnection).s(s2).t3) & " "
          totalLen = totalLen + 3
        Else
          sCheat = sCheat & FiveChrLon(tileID) & " "
          totalLen = totalLen + 2
        End If
      End If
    Next s2
    sCheat = sCheat & "00 FF"
    totalLen = totalLen + 2
   ' GoTo drawGot
     End If
     drawNext = False
     Exit For
   End If
  Next s
  
  If drawNext = True Then
  sCheat = sCheat & " 69 " & FiveChrLon(myX(idConnection) + x) & " " & FiveChrLon(myY(idConnection) + y) & " " & _
   GoodHex(CByte(myZ(idConnection))) & " "
  totalLen = totalLen + 6
  tileID = GetTheLong(Matrix(y, x, floor, idConnection).s(0).t1, Matrix(y, x, floor, idConnection).s(0).t2)
  If tileID = 0 Then
    sCheat = sCheat & "01 FF"
    totalLen = totalLen + 2
  Else
    For s = 0 To 10
      tileID = GetTheLong(Matrix(y, x, floor, idConnection).s(s).t1, Matrix(y, x, floor, idConnection).s(s).t2)
      If tileID = 0 Then
        Exit For
      ElseIf tileID = 97 Then
        sCheat = sCheat & "63 00 " & SpaceID(Matrix(y, x, floor, idConnection).s(s).dblID) & " 02 "
        totalLen = totalLen + 7
      Else
        If DatTiles(tileID).haveExtraByte = True Then
          sCheat = sCheat & FiveChrLon(tileID) & " " & GoodHex(Matrix(y, x, floor, idConnection).s(s).t3) & " "
          totalLen = totalLen + 3
        Else
          sCheat = sCheat & FiveChrLon(tileID) & " "
          totalLen = totalLen + 2
        End If
      End If
    Next s
    sCheat = sCheat & "00 FF"
    totalLen = totalLen + 2
  End If
  End If
  'squareCount = squareCount + 1
 ' If squareCount >= squareLim Then
 '   GoTo drawGot
 ' End If
  Next x
  Next y
drawGot:
  
  sCheat = FiveChrLon(totalLen) & sCheat
  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & sCheat
  inRes = GetCheatPacket(cPacket, sCheat)
  frmMain.UnifiedSendToClientGame idConnection, cPacket
  DoEvents
  IgnoreServer(idConnection) = True
  waitThisMs2 (2000)
  IgnoreServer(idConnection) = False
  StartReconnection idConnection
  ViewFloor = 0
  Exit Function
goterr:
  ViewFloor = -1
End Function

Public Function GameInspect(idConnection As Integer, x As Long, y As Long, z As Long) As Long
  Dim aRes As Long
  Dim useX As Long
  Dim useY As Long
  Dim useZ As Long
  Dim useS As Byte
  Dim ts As Byte
  Dim tileID As Long
  Dim continue As Boolean
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim inRes As Integer
  Dim zdif As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  '09 00 8C 3C 7D AA 7D 08 E8 09 02
  'aRes = SendLogSystemMessageToClient(idConnection, "Inspecting " & X & "," & Y & "," & Z)
 ' DoEvents
  continue = True
  ' inspect players
  For ts = 1 To 10
    tileID = GetTheLong(Matrix(y, x, z, idConnection).s(ts).t1, Matrix(y, x, z, idConnection).s(ts).t2)
    If tileID = 97 Then
      useS = ts
      continue = False
      Exit For
    End If
  Next ts
  ' inspect items
  If continue = True Then
  For ts = 1 To 10
    tileID = GetTheLong(Matrix(y, x, z, idConnection).s(ts).t1, Matrix(y, x, z, idConnection).s(ts).t2)
    If DatTiles(tileID).alwaysOnTop = False And tileID <> 0 Then
      useS = ts
      continue = False
      Exit For
    End If
  Next ts
  End If
  ' inspect ontop
  If continue = True Then
  For ts = 1 To 10
    tileID = GetTheLong(Matrix(y, x, z, idConnection).s(ts).t1, Matrix(y, x, z, idConnection).s(ts).t2)
    If DatTiles(tileID).alwaysOnTop = True Then
      useS = ts
      continue = False
      Exit For
    End If
  Next ts
  End If
  ' inspect ground
  If continue = True Then
      tileID = GetTheLong(Matrix(y, x, z, idConnection).s(0).t1, Matrix(y, x, z, idConnection).s(0).t2)
      useS = 0
  End If
  zdif = myZ(idConnection) - z
  useX = myX(idConnection) + x + zdif
  useY = myY(idConnection) + y + zdif
  useZ = z
  '09 00 8C 3C 7D AA 7D 08 E8 09 02
  sCheat = "8C " & FiveChrLon(useX) & " " & FiveChrLon(useY) & " " & GoodHex(CByte(useZ)) & " " & FiveChrLon(tileID) & " " & GoodHex(useS)
 ' frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & sCheat
 
 SafeCastCheatString "GameInspect1", idConnection, sCheat
 

  GameInspect = 0
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at GameInspect #"
  frmMain.DoCloseActions idConnection
  DoEvents
  GameInspect = -1
End Function

Public Sub DoRandomMove(idConnection As Integer)
  Dim aRes As Long
  Dim sCheat As String
  Dim moveType As Byte
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  
'  cPacket(0) = &H1
  'cPacket(1) = &H0
  moveType = CByte(Int((4 * Rnd) + 101))
  sCheat = GoodHex(moveType)
  'cPacket(2) = moveType
    If publicDebugMode = True Then
    Select Case moveType
    Case &H65
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] semitraped - doing random step : north")
      DoEvents
    Case &H66
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] semitraped - doing random step : right")
      DoEvents
    Case &H67
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] semitraped - doing random step : south")
      DoEvents
    Case &H68
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] semitraped - doing random step : left")
      DoEvents
    Case &H6A
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] semitraped - doing random step : north + right")
      DoEvents
    Case &H6B
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] semitraped - doing random step : south + right")
      DoEvents
    Case &H6C
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] semitraped - doing random step : south + left")
      DoEvents
    Case &H6D
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] semitraped - doing random step : north + left")
      DoEvents
    End Select
  End If
  If GameConnected(idConnection) = True And sentFirstPacket(idConnection) = True Then
    SafeCastCheatString "DoRandomMove1", idConnection, sCheat
  End If
  Exit Sub
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at DoRandomMove #"
  frmMain.DoCloseActions idConnection
End Sub

Public Sub DoManualMove(idConnection As Integer, moveType As Byte)
  Dim aRes As Long
  Dim cPacket(2) As Byte
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  cPacket(0) = &H1
  cPacket(1) = &H0
  cPacket(2) = moveType
  If publicDebugMode = True Then
    Select Case moveType
    Case &H65
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Force step : north")
      DoEvents
    Case &H66
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Force step : right")
      DoEvents
    Case &H67
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Force step : south")
      DoEvents
    Case &H68
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Force step : left")
      DoEvents
    Case &H6A
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Force step : north + right")
      DoEvents
    Case &H6B
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Force step : south + right")
      DoEvents
    Case &H6C
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Force step : south + left")
      DoEvents
    Case &H6D
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Force step : north + left")
      DoEvents
    End Select
  End If
  If GameConnected(idConnection) = True And sentFirstPacket(idConnection) = True Then
    frmMain.UnifiedSendToServerGame idConnection, cPacket, True
    DoEvents
  End If
  Exit Sub
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at DoManualMove #"
  frmMain.DoCloseActions idConnection
End Sub




Public Function ValidateOutfitNumber(strNumber As String) As Long
  Dim intNumber As Long
  Dim evNumber As Long
  On Error GoTo goterr
  intNumber = -1
  evNumber = CLng(strNumber)
  If (evNumber < firstValidOutfit) Or (evNumber > lastValidOutfit) Then
    ValidateOutfitNumber = -1
    Exit Function
  End If
  Select Case evNumber
    Case 10, 11, 12, 20, 46, 47, 72, 77, 93, 96, 97, 98, 135
      intNumber = -1
    Case Else
      intNumber = evNumber
  End Select
  ValidateOutfitNumber = intNumber
  Exit Function
goterr:
  ValidateOutfitNumber = -1
End Function

Public Function NewValidateOutfitNumber(strNumber As String) As Long
  Dim intNumber As Long
  Dim evNumber As Long
  On Error GoTo goterr
  intNumber = -1
  evNumber = CLng(strNumber)
  If (evNumber < firstValidOutfit) Or (evNumber > lastValidOutfit) Then
    NewValidateOutfitNumber = -1
    Exit Function
  End If
  Select Case evNumber
    Case 135
      intNumber = -1
    Case Else
      intNumber = evNumber
  End Select
  NewValidateOutfitNumber = intNumber
  Exit Function
goterr:
  NewValidateOutfitNumber = -1
End Function



Public Function SendOutfit(idConnection As Integer, strNumber As String) As Long
  Dim cPacket() As Byte
  Dim aRes As Long
  Dim intNumber As Long
  Dim sCheat As String
  Dim inRes As Integer
  Dim b1 As Byte
  Dim b2 As Byte
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  If TibiaVersionLong <= 750 Then
      '08 00 C8 02 00 00 00 00 02 86
    '  ReDim cPacket(9)
    '  cPacket(0) = &H8
    '  cPacket(1) = &H0
    '  cPacket(2) = &HC8
    '  cPacket(3) = &H2
    '  cPacket(4) = &H0
    '  cPacket(5) = &H0
    '  cPacket(6) = &H0
    '  cPacket(7) = &H0
    '  cPacket(8) = &H2
    '  cPacket(9) = &H86
      sCheat = "C8 02 00 00 00 00 02 86"
      SendingSpecialOutfit(idConnection) = True
      SafeCastCheatString "SendOutfit1-1", idConnection, sCheat
    '  frmMain.UnifiedSendToClientGame idConnection, cPacket
    '  DoEvents
  Else
    If strNumber = "" Then
      aRes = SendLogSystemMessageToClient(idConnection, "Sorry, since Tibia 7.55 you should use: exiva outfit <number from " & firstValidOutfit & " to " & lastValidOutfit & ">  example: exiva outfit 124")
      DoEvents
    Else
      If (TibiaVersionLong >= 773) Then 'still valid for 7.83
        intNumber = NewValidateOutfitNumber(strNumber)
      Else
        intNumber = ValidateOutfitNumber(strNumber)
      End If
      If intNumber >= 0 Then
        If TibiaVersionLong >= 870 Then
         '0E 00 8E B2 48 C8 01 82 00 12 12 12 12 00 00 00
           b1 = LowByteOfLong(intNumber)
          b2 = HighByteOfLong(intNumber)
          sCheat = "8E " & SpaceID(myID(idConnection)) & " " & _
         GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(&H0) & " " & GoodHex(&H0) & " " & GoodHex(&H0) & " " & GoodHex(&H0) & " " & GoodHex(&H0) & " " & GoodHex(&H0) & " " & GoodHex(&H0)
        ElseIf TibiaVersionLong >= 773 Then
           b1 = LowByteOfLong(intNumber)
          b2 = HighByteOfLong(intNumber)
          sCheat = "8E " & SpaceID(myID(idConnection)) & " " & _
         GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(&H0) & " " & GoodHex(&H0) & " " & GoodHex(&H0) & " " & GoodHex(&H0) & " " & GoodHex(&H0)
        ElseIf TibiaVersionLong > 760 Then
          b1 = LowByteOfLong(intNumber)
          b2 = HighByteOfLong(intNumber)
          sCheat = "8E " & SpaceID(myID(idConnection)) & " " & _
         GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(&H0) & " " & GoodHex(&H0) & " " & GoodHex(&H0) & " " & GoodHex(&H0)
        Else
          sCheat = "8E " & SpaceID(myID(idConnection)) & " " & _
         GoodHex(CByte(intNumber)) & " " & GoodHex(&H0) & " " & GoodHex(&H0) & " " & GoodHex(&H0) & " " & GoodHex(&H0)
        End If
        SafeCastCheatString "SendOutfit1-2", idConnection, sCheat
'        inRes = GetCheatPacket(cPacket, sCheat)
'        frmMain.UnifiedSendToClientGame idConnection, cPacket
'        DoEvents
      Else
        aRes = SendLogSystemMessageToClient(idConnection, "Sorry, that is an invalid outfit for a player.")
        DoEvents
      End If
    End If
  End If
  SendOutfit = 0
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " got error at SendOutfit #"
  'frmMain.DoCloseActions idConnection
  'DoEvents
  SendOutfit = -1
End Function

Public Function SendOutfit2(idConnection As Integer, b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte, b5 As Byte)
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim inRes As Integer
  Dim aRes As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  '0A 00 8E 0C 4D 9C 00 82 00 4D 5E 72
  aRes = SendLogSystemMessageToClient(idConnection, _
   "BlackdProxy: cheat outfit selected ( " & GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(b3) & " " & GoodHex(b4) & " " & GoodHex(b5) & " ) Only you will see it")
  DoEvents
  SendingSpecialOutfit(idConnection) = False
  sCheat = "8E " & SpaceID(myID(idConnection)) & " " & _
   GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(b3) & " " & GoodHex(b4) & " " & GoodHex(b5)

 ' aRes = SendLogSystemMessageToClient(idConnection, sCheat)
  'DoEvents
  SafeCastCheatString "SendOutfit2", idConnection, sCheat
'  inRes = GetCheatPacket(cPacket, sCheat)
'  frmMain.UnifiedSendToClientGame idConnection, cPacket
'  DoEvents
  SendOutfit2 = 0
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at SendOutfit2 #"
  frmMain.DoCloseActions idConnection
  DoEvents
  SendOutfit2 = -1
End Function

Public Function SendOutfit3(idConnection As Integer, b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte, b5 As Byte) As Long
  Dim aRes As Long
  Dim myBpos As Long
  Dim pid As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  ' Change outfit by memory
  If TibiaVersionLong >= 1100 Then
    SendOutfit3 = -1 ' This function is not compatible with Tibia 11+
    Exit Function
  End If
  
  GetProcessIDs idConnection
  aRes = SendLogSystemMessageToClient(idConnection, _
   "BlackdProxy: cheat outfit selected ( " & GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(b3) & " " & GoodHex(b4) & " " & GoodHex(b5) & " ) Only you will see it")
  DoEvents
  SendingSpecialOutfit(idConnection) = False
  'change outfit in memory
  pid = ProcessID(idConnection)
  myBpos = MyBattleListPosition(idConnection)
  If (myBpos > -1) Then
    Memory_WriteByte adrOutfit + (myBpos * CharDist), b1, pid
    Memory_WriteByte adrOutfit + 4 + (myBpos * CharDist), b2, pid
    Memory_WriteByte adrOutfit + 8 + (myBpos * CharDist), b3, pid
    Memory_WriteByte adrOutfit + 12 + (myBpos * CharDist), b4, pid
    Memory_WriteByte adrOutfit + 16 + (myBpos * CharDist), b5, pid
  End If
  SendOutfit3 = 0
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at SendOutfit2 #"
  frmMain.DoCloseActions idConnection
  DoEvents
  SendOutfit3 = -1
End Function

Public Function SendOutfit4(idConnection As Integer, b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte)
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim inRes As Integer
  Dim aRes As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  ' tibia 7.61+
  '0B 00 8E 0C 4D 9C 00 82 00 4D 5E 72
  aRes = SendLogSystemMessageToClient(idConnection, _
   "BlackdProxy: cheat outfit selected ( " & GoodHex(b1) & " " & GoodHex(b2) & " " & _
    GoodHex(b3) & " " & GoodHex(b4) & " " & GoodHex(b5) & " " & GoodHex(b6) & _
    " ) Only you will see it")
  DoEvents
  SendingSpecialOutfit(idConnection) = False
  If (TibiaVersionLong >= 870) Then
   sCheat = "8E " & SpaceID(myID(idConnection)) & " " & _
   GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(b3) & " " & GoodHex(b4) & " " & _
   GoodHex(b5) & " " & GoodHex(b6) & " 00 00 00"
  ElseIf (TibiaVersionLong >= 773) Then
   sCheat = "8E " & SpaceID(myID(idConnection)) & " " & _
   GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(b3) & " " & GoodHex(b4) & " " & _
   GoodHex(b5) & " " & GoodHex(b6) & " 00"
  Else
   sCheat = "8E " & SpaceID(myID(idConnection)) & " " & _
   GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(b3) & " " & GoodHex(b4) & " " & _
   GoodHex(b5) & " " & GoodHex(b6)
  End If
 ' aRes = SendLogSystemMessageToClient(idConnection, sCheat)
  'DoEvents
  SafeCastCheatString "SendOutfit4", idConnection, sCheat
'  inRes = GetCheatPacket(cPacket, sCheat)
'  frmMain.UnifiedSendToClientGame idConnection, cPacket
'  DoEvents
  SendOutfit4 = 0
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at SendOutfit4 #"
  frmMain.DoCloseActions idConnection
  DoEvents
  SendOutfit4 = -1
End Function

Public Function SaveHand(idConnection As Integer, allowAmmo As Boolean, handID As Byte, preferedBP As Byte) As Long
  '&H5 for right hand
  '&H6 for left hand
  Dim i As Long
  Dim j As Long
  Dim cPacket() As Byte
  Dim aRes As Long
  Dim fRes As TypeSearchItemResult2
  Dim sCheat As String
  Dim b16 As Byte
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  If publicDebugMode = True Then
    aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Storing item. Status: mana = " & _
    CStr(myMana(idConnection)) & " Saved item = " & GoodHex(savedItem(idConnection).t1) & " " & GoodHex(savedItem(idConnection).t2))
    DoEvents
  End If
  If ((mySlot(idConnection, SLOT_AMMUNITION).t1 = &H0) And (mySlot(idConnection, SLOT_AMMUNITION).t2 = &H0) And (allowAmmo = True)) Then
  
  If publicDebugMode = True Then
    If handID = CByte(SLOT_RIGHTHAND) Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Moving item " & GoodHex(mySlot(idConnection, SLOT_RIGHTHAND).t1) & " " & GoodHex(mySlot(idConnection, SLOT_RIGHTHAND).t2) & " " & GoodHex(mySlot(idConnection, SLOT_RIGHTHAND).t3) & " to ammo from righthand")
    Else ' left hand
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Moving item " & GoodHex(mySlot(idConnection, SLOT_LEFTHAND).t1) & " " & GoodHex(mySlot(idConnection, SLOT_LEFTHAND).t2) & " " & GoodHex(mySlot(idConnection, SLOT_LEFTHAND).t3) & " to ammo from lefthand")
    End If
    DoEvents
  End If
  sCheat = "78 FF FF " & GoodHex(handID) & " 00 00 "
'  ReDim cPacket(16)
'  cPacket(0) = LowByteOfLong(15)
'  cPacket(1) = HighByteOfLong(15)
'  cPacket(2) = &H78
'  cPacket(3) = &HFF
'  cPacket(4) = &HFF
'  cPacket(5) = handID ' from right hand
'  cPacket(6) = &H0
'  cPacket(7) = &H0
  If handID = CByte(SLOT_RIGHTHAND) Then
    sCheat = sCheat & GoodHex(mySlot(idConnection, SLOT_RIGHTHAND).t1) & " " & GoodHex(mySlot(idConnection, SLOT_RIGHTHAND).t2)
  Else
    sCheat = sCheat & GoodHex(mySlot(idConnection, SLOT_LEFTHAND).t1) & " " & GoodHex(mySlot(idConnection, SLOT_LEFTHAND).t2)
  End If
  sCheat = sCheat & "00 FF FF 0A 00 00 "
'  cPacket(10) = &H0
'  cPacket(11) = &HFF
'  cPacket(12) = &HFF
'  cPacket(13) = &HA ' to ammo
'  cPacket(14) = &H0
'  cPacket(15) = &H0
  If handID = CByte(SLOT_RIGHTHAND) Then
    b16 = mySlot(idConnection, SLOT_RIGHTHAND).t3
  Else
    b16 = mySlot(idConnection, SLOT_LEFTHAND).t3
  End If
  'If ((TibiaVersionLong >= 860) And (cPacket(16) = &H0)) Then
  If (b16 = &H0) Then
    b16 = &H1
  End If
  sCheat = sCheat & GoodHex(b16)
  SafeCastCheatString "SaveHand1", idConnection, sCheat
  

  
  Else
  '5 peces a slot 3
  '0F 00 78 FF FF 06 00 00 BC 0D 00 FF FF 40 00 03 05
  If preferedBP = &HFF Then
    fRes = SearchFreeSlot(idConnection)
  Else
    fRes = SearchFreeSlotInContainer(idConnection, preferedBP)
    If fRes.foundCount = 0 Then
      fRes = SearchFreeSlot(idConnection)
    End If
  End If
  If fRes.foundCount = 1 Then
    If publicDebugMode = True Then
      If handID = CByte(SLOT_RIGHTHAND) Then
        aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Moving item " & GoodHex(mySlot(idConnection, SLOT_RIGHTHAND).t1) & " " & GoodHex(mySlot(idConnection, SLOT_RIGHTHAND).t2) & " " & GoodHex(mySlot(idConnection, SLOT_RIGHTHAND).t3) & " to [bpID " & fRes.bpID & " slotid " & fRes.slotID & "] from righthand")
      Else
        aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Moving item " & GoodHex(mySlot(idConnection, SLOT_LEFTHAND).t1) & " " & GoodHex(mySlot(idConnection, SLOT_LEFTHAND).t2) & " " & GoodHex(mySlot(idConnection, SLOT_LEFTHAND).t3) & " to [bpID " & fRes.bpID & " slotid " & fRes.slotID & "] from lefthand")
      End If
      DoEvents
    End If
    sCheat = "78 FF FF " & GoodHex(handID) & " 00 00 "
'    ReDim cPacket(16)
'    cPacket(0) = LowByteOfLong(15)
'    cPacket(1) = HighByteOfLong(15)
'    cPacket(2) = &H78
'    cPacket(3) = &HFF
'    cPacket(4) = &HFF
'    cPacket(5) = handID
'    cPacket(6) = &H0
'    cPacket(7) = &H0
    If handID = CByte(SLOT_RIGHTHAND) Then
'      cPacket(8) = mySlot(idConnection, SLOT_RIGHTHAND).t1
'      cPacket(9) = mySlot(idConnection, SLOT_RIGHTHAND).t2
      sCheat = sCheat & GoodHex(mySlot(idConnection, SLOT_RIGHTHAND).t1) & " " & GoodHex(mySlot(idConnection, SLOT_RIGHTHAND).t2)
    Else
'      cPacket(8) = mySlot(idConnection, SLOT_LEFTHAND).t1
'      cPacket(9) = mySlot(idConnection, SLOT_LEFTHAND).t2
      sCheat = sCheat & GoodHex(mySlot(idConnection, SLOT_LEFTHAND).t1) & " " & GoodHex(mySlot(idConnection, SLOT_LEFTHAND).t2)
    End If
    sCheat = sCheat & " 00 FF FF " & GoodHex(&H40 + fRes.bpID) & " 00 " & GoodHex(fRes.slotID) & " "
'    cPacket(10) = &H0
'    cPacket(11) = &HFF
'    cPacket(12) = &HFF
'    cPacket(13) = &H40 + fRes.bpID
'    cPacket(14) = &H0
'    cPacket(15) = fRes.slotID
    If handID = CByte(SLOT_RIGHTHAND) Then
      b16 = mySlot(idConnection, SLOT_RIGHTHAND).t3
    Else
      b16 = mySlot(idConnection, SLOT_LEFTHAND).t3
    End If
    If (b16 = &H0) Then
     b16 = &H1
    End If
    
    sCheat = sCheat & GoodHex(b16)
    SafeCastCheatString "SaveHand2", idConnection, sCheat
  

  Else
    aRes = SendLogSystemMessageToClient(idConnection, "Sorry, you need a free slot in a BACKPACK to use runemaker")
    DoEvents
    SaveHand = -1
    Exit Function
  End If
  
  End If
  SaveHand = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at SaveHand #"
  frmMain.DoCloseActions idConnection
  DoEvents
  SaveHand = -1
End Function

Public Function BooleanAsStr(bln As Boolean) As String
  If bln = True Then
    BooleanAsStr = "True"
  Else
   BooleanAsStr = "False"
  End If
End Function

Public Function MoveItemToEquip(idConnection As Integer, b1 As Byte, b2 As Byte, equip As Byte) As Long
  ' 0F 00 78 FF FF 40 00 06 0D 0C 06 FF FF 05 00 00 01
  ' 0F 00 78 FF FF 40 00 02 81 0D 02 FF FF 0A 00 00 01 ' nuevo
  Dim i As Long
  Dim j As Long
  Dim cPacket() As Byte
  Dim aRes As Long
  Dim bpID As Byte
  Dim slotID As Byte
  Dim resF As TypeSearchItemResult2
  Dim sCheat As String
  Dim b16 As Byte
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  resF = SearchItem(idConnection, b1, b2)
  
  If resF.foundCount = 0 Then
    'aRes = SendLogSystemMessageToClient(idconnection, "Sorry, can't find such item in any backpack.")
    'DoEvents
  Else
    sCheat = "78 FF FF " & GoodHex(&H40 + resF.bpID) & " 00 " & GoodHex(resF.slotID) & " " & _
    GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(resF.slotID) & " FF FF " & GoodHex(equip) & " 00 00 "
'    ReDim cPacket(16)
'    cPacket(0) = LowByteOfLong(15)
'    cPacket(1) = HighByteOfLong(15)
'    cPacket(2) = &H78
'    cPacket(3) = &HFF
'    cPacket(4) = &HFF
'    cPacket(5) = &H40 + resF.bpID
'    cPacket(6) = &H0
'    cPacket(7) = resF.slotID
'    cPacket(8) = b1
'    cPacket(9) = b2
'    cPacket(10) = resF.slotID
'    cPacket(11) = &HFF
'    cPacket(12) = &HFF
'    cPacket(13) = equip
'    cPacket(14) = &H0
'    cPacket(15) = &H0
    If TibiaVersionLong >= 860 Or TibiaVersionLong = 760 Then ' o 861?
    'fix for 7.6 OT servers
        If (resF.amount = 0) Then
          b16 = 1
        Else
          b16 = resF.amount ' amount
        End If
    Else
      b16 = resF.amount ' amount
    End If
    sCheat = sCheat & GoodHex(b16)
    SafeCastCheatString "MoveItemToEquip", idConnection, sCheat
  End If
   
  MoveItemToEquip = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at MoveItemToEquip #"
  frmMain.DoCloseActions idConnection
  DoEvents
  MoveItemToEquip = -1
End Function

Public Function StopMove(idConnection As Integer) As Long
  Dim i As Long
  Dim j As Long
  Dim sCheat As String
  #If FinalMode Then
  On Error GoTo errIgnoreIt
  #End If
'  ReDim cPacket(2)
'  cPacket(0) = &H1
'  cPacket(1) = &H0
'  cPacket(2) = &HBE
  sCheat = "BE"
  SafeCastCheatString "sCheat1", idConnection, sCheat
  StopMove = 0
  Exit Function
errIgnoreIt:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & _
   "error (" & Err.Number & ") detected at StopMove : " & Err.Description
  StopMove = -1
End Function

Public Function ChangePauseStatus(idConnection As Integer, newStatus As Boolean, exStatus As Boolean) As Long
  Dim aRes As Long
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  CheatsPaused(idConnection) = newStatus
  AllowUHpaused(idConnection) = exStatus
  If newStatus = True Then
    If (GameConnected(idConnection) = True) Then
      'stop attack
      aRes = MeleeAttack(idConnection, 0)
      DoEvents
      'pause move
      aRes = StopMove(idConnection)
      If AllowUHpaused(idConnection) = True Then
        aRes = GiveGMmessage(idConnection, "Automatic functions have been paused for this client, except auto rune heal - restore with exiva play", "BlackdProxy")
        DoEvents
      Else
        aRes = GiveGMmessage(idConnection, "Automatic functions have been paused for this client - restore with exiva play", "BlackdProxy")
        DoEvents
      End If
    End If
  Else
    AllowUHpaused(idConnection) = False
    DangerPK(idConnection) = False
    DangerGM(idConnection) = False
    LogoutTimeGM(idConnection) = 0
    moveRetry(idConnection) = 0
    ChangePlayTheDangerSound False
    RemoveSpamOrder idConnection, 1
    logoutAllowed(idConnection) = 0
    If (GameConnected(idConnection) = True) Then
      aRes = SendLogSystemMessageToClient(idConnection, "BlackdProxy: All automatic functions are enabled again")
      DoEvents
    End If
  End If
  ChangePauseStatus = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at ChangePauseStatus #"
  frmMain.DoCloseActions idConnection
  DoEvents
  ChangePauseStatus = -1
End Function

Public Function PlotPosition(idConnection As Integer, rawPosition As String)
  Dim aRes As Long
  Dim x As Long
  Dim y As Long
  Dim z As Long
  Dim xs As String
  Dim ys As String
  Dim zs As String
  Dim pos As Long
  Dim toEnd As Long
  Dim justAdded As String
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  pos = 1
  toEnd = Len(rawPosition)
  xs = ParseString(rawPosition, pos, toEnd, ",")
  x = CLng(xs)
  pos = pos + 1
  ys = ParseString(rawPosition, pos, toEnd, ",")
  y = CLng(ys)
  pos = pos + 1
  zs = ParseString(rawPosition, pos, toEnd, ",")
  z = CLng(zs)

  frmMapReader.WindowState = vbNormal
  frmMapReader.Show
  DoEvents
  justAdded = frmMapReader.AddMarkToBigMap(x, y, z)
  frmMapReader.SetCurrentCenter (justAdded)
  frmMapReader.ShowCenter
  DoEvents
  frmMapReader.timerBigMapUpdate.enabled = True
  aRes = SendLogSystemMessageToClient(idConnection, "Wrote mark " & justAdded & " at position: " & CStr(x) & "," & CStr(y) & "," & CStr(z))
  DoEvents
  PlotPosition = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " got error at PlotPosition #"
  DoEvents
  PlotPosition = -1
End Function

Public Function ChangeSpeed(idConnection As Integer, rawLevel As String)
  Dim aRes As Long
  Dim xs As String
  Dim speed As Long
  Dim pos As Long
  Dim toEnd As Long
  Dim justAdded As String
  Dim b1 As Byte
  Dim b2 As Byte
  Dim lastPos As Long
  Dim tibiaclient As Long
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  GetProcessIDs idConnection
  If rawLevel = "" Then
    tibiaclient = ProcessID(idConnection)
    lastPos = MyBattleListPosition(idConnection)
    If (lastPos > -1) Then
        b1 = Memory_ReadByte(adrNChar + (lastPos * CharDist) + SpeedDist, tibiaclient)
        b2 = Memory_ReadByte(adrNChar + (lastPos * CharDist) + SpeedDist + 1, tibiaclient)
        speed = GetTheLong(b1, b2)
        aRes = SendLogSystemMessageToClient(idConnection, "Your current internal speed is  " & CStr(speed) & Chr(10) & _
         "Change it with exiva speed X" & Chr(10) & _
         "For example, to set internal speed = 500: exiva speed 500")
        DoEvents
        ChangeSpeed = 0
        Exit Function
    Else
        aRes = GiveGMmessage(idConnection, "Unable to use this feature at this moment", "BlackdProxy")
        ChangeSpeed = -1
        Exit Function
    End If
  Else
    pos = 1
    toEnd = Len(rawLevel)
    xs = ParseString(rawLevel, pos, toEnd, ",")
    speed = CLng(xs)
    tibiaclient = ProcessID(idConnection)
    lastPos = MyBattleListPosition(idConnection)
    If (lastPos > -1) Then
        b1 = LowByteOfLong(speed)
        b2 = HighByteOfLong(speed)
        Memory_WriteByte (adrNChar + (lastPos * CharDist) + SpeedDist), b1, tibiaclient
        Memory_WriteByte (adrNChar + (lastPos * CharDist) + SpeedDist + 1), b2, tibiaclient
        aRes = GiveGMmessage(idConnection, "Speed changed. Now your internal speed is " & CStr(speed), "BlackdProxy")
        DoEvents
        ChangeSpeed = 0
    Else
        aRes = GiveGMmessage(idConnection, "Unable to use this feature at this moment", "BlackdProxy")
        DoEvents
        ChangeSpeed = 0
    End If
  End If
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " got error at ChangeSpeed #"
  DoEvents
  ChangeSpeed = -1
End Function

Public Function ExecuteInFocusedTibia(spellString As String) As Integer
  Dim idConnection As Integer
  Dim cPacket() As Byte
  Dim lonS As Long
  Dim totalL As Long
  Dim limL As Long
  Dim i As Long
  Dim j As Long
  Dim aRes As Long
  Dim pIDfocusedWindow As Long
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  ' get the windows with focus
  idConnection = 0 ' by default
  pIDfocusedWindow = GetForegroundWindow()
  GetProcessAllProcessIDs
  For i = 1 To MAXCLIENTS
    If GameConnected(i) = True Then
      If (ProcessID(i) = pIDfocusedWindow) Then
         idConnection = i
         Exit For
      End If
    End If
  Next i
  If idConnection = 0 Then 'if tibia is not focused
    ExecuteInFocusedTibia = 0
    Exit Function ' then end
  End If
  ExecuteInFocusedTibia = ExecuteInTibia(spellString, idConnection, True)
  Exit Function
errclose:
  ExecuteInFocusedTibia = -1
End Function


Public Function ExecuteInTibia(spellString As String, idConnection As Integer, cantBeIgnored As Boolean, Optional onlyCommands As Boolean = False) As Integer
  Dim cPacket() As Byte
  Dim lonS As Long
  Dim totalL As Long
  Dim limL As Long
  Dim i As Long
  Dim j As Long
  Dim aRes As Long
  Dim msgto As String
  Dim theRealMsg As String
  Dim sok As Boolean
  Dim thec As String
  Dim thecL As Long
  Dim theRealMsgL As Long
  Dim pos As Long
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  If frmHardcoreCheats.chkApplyCheats.value = 0 Then
    ExecuteInTibia = 0
    Exit Function ' fixed since 13.3
  End If
  If spellString = "" Then
    ExecuteInTibia = 0
    Exit Function
  End If
  lonS = Len(spellString)
  sok = False
  msgto = ""
  If Left(spellString, 1) = "*" Then
    For i = 2 To lonS
      thec = Mid$(spellString, i, 1)
      If (thec = "*") Then
        sok = True
        theRealMsgL = lonS - i
        Exit For
      Else
        msgto = msgto & thec
      End If
    Next i
  End If
  If (sok = True) And i > 2 Then
    ' SEND PRIVATE MESSAGE
    theRealMsg = Right(spellString, theRealMsgL)
    thecL = Len(msgto)
    limL = 7 + thecL + theRealMsgL
    ReDim cPacket(limL)
    totalL = 6 + lonS
    cPacket(0) = LowByteOfLong(totalL)
    cPacket(1) = HighByteOfLong(totalL)
    cPacket(2) = &H96
    If TibiaVersionLong >= 872 Then
        cPacket(3) = &H5
    ElseIf TibiaVersionLong >= 820 Then
        cPacket(3) = &H6
    Else
        cPacket(3) = &H4
    End If
    cPacket(4) = LowByteOfLong(thecL)
    cPacket(5) = HighByteOfLong(thecL)
    pos = 5
    For i = 1 To thecL
      pos = pos + 1
      cPacket(pos) = CByte(Asc(Mid(msgto, i, 1)))
    Next i
    pos = pos + 1
    cPacket(pos) = LowByteOfLong(theRealMsgL)
    pos = pos + 1
    cPacket(pos) = HighByteOfLong(theRealMsgL)
    For i = 1 To theRealMsgL
      pos = pos + 1
      cPacket(pos) = CByte(Asc(Mid(theRealMsg, i, 1)))
    Next i
    aRes = ApplyHardcoreCheats(cPacket, idConnection, True)
    If onlyCommands = True Then
        aRes = 1
    End If
    If aRes = 1 Then ' Hardcore cheats require skiping this packet
      DoEvents
      ExecuteInTibia = 0
    ElseIf aRes = 0 Then
      If ((GetTickCount() > nextAllowedmsg(idConnection)) Or (cantBeIgnored = True)) Then
        nextAllowedmsg(idConnection) = GetTickCount() + DELAYBETWEENAUTOMSG_ms
        If (GameConnected(idConnection) = True) And (frmMain.sckServerGame(idConnection).State = sckConnected) Then
          frmMain.UnifiedSendToServerGame idConnection, cPacket, True
        End If
        DoEvents
        aRes = SendLogSystemMessageToClient(idConnection, "Blackd Proxy sent msg to " & msgto & " : " & theRealMsg)
        DoEvents
      End If
      ExecuteInTibia = 0
    Else
      DoEvents
      ExecuteInTibia = -1
    End If
    Exit Function
  End If
  limL = 5 + lonS
  ReDim cPacket(limL)
  totalL = 4 + lonS
  cPacket(0) = LowByteOfLong(totalL)
  cPacket(1) = HighByteOfLong(totalL)
  cPacket(2) = &H96
  cPacket(3) = &H1
  cPacket(4) = LowByteOfLong(lonS)
  cPacket(5) = HighByteOfLong(lonS)
  j = 1
  For i = 6 To limL
    cPacket(i) = CByte(Asc(Mid(spellString, j, 1)))
    j = j + 1
  Next i
  aRes = ApplyHardcoreCheats(cPacket, idConnection, True)
  If onlyCommands = True Then
    aRes = 1
  End If
  If aRes = 1 Then ' Hardcore cheats require skiping this packet
    ExecuteInTibia = 0
    Exit Function
  ElseIf aRes = 0 Then
    If ((GetTickCount() > nextAllowedmsg(idConnection)) Or (cantBeIgnored = True)) Then
      nextAllowedmsg(idConnection) = GetTickCount() + DELAYBETWEENAUTOMSG_ms
      If (GameConnected(idConnection) = True) And (frmMain.sckServerGame(idConnection).State = sckConnected) Then
        frmMain.UnifiedSendToServerGame idConnection, cPacket, True
      End If
      DoEvents
    End If
    ExecuteInTibia = 0
  Else
    DoEvents
    ExecuteInTibia = -1
  End If
  Exit Function
errclose:
  ExecuteInTibia = -1
End Function


Public Function expReset(idConnection As Integer) As Long
  Dim aRes As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  myInitialExp(idConnection) = myExp(idConnection)
  myInitialTickCount(idConnection) = GetTickCount()
  aRes = SendLogSystemMessageToClient(idConnection, "Your exp/h has been reseted")
  DoEvents
  expReset = 0
  Exit Function
goterr:
  expReset = -1
End Function

Public Function ProcessKillOrder(idConnection As Integer, Target As String) As Long
  Dim aRes As Long
  Dim myS As Byte
  Dim lTarget As String
  Dim sCheat As String
  Dim cPacket() As Byte
  Dim x As Long
  Dim y As Long
  Dim s As Long
  Dim tileID As Long
  Dim tmpID As Double
  Dim lSquare As String
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  If (Target = "") Or (Target = " ") Then ' use last targeted
    Target = currTargetName(idConnection)
  ElseIf Left(Target, 1) = " " Then ' delete space
    Target = Right(Target, Len(Target) - 1)
  End If
  If (Target = GotKillOrderTargetName(idConnection)) Then
    GotKillOrderTargetName(idConnection) = ""
    GotKillOrderTargetID(idConnection) = 0
    GotKillOrder(idConnection) = False
    aRes = GiveGMmessage(idConnection, "Auto kill mode OFF" & GotKillOrderTargetName(idConnection), "Stoped")
    DoEvents
    ProcessKillOrder = 0
    Exit Function
  End If
  GotKillOrderTargetID(idConnection) = 0
  GotKillOrderTargetName(idConnection) = ""
  If (TibiaVersionLong < 760) Then
    myS = MyStackPos(idConnection)
  Else
    myS = FirstPersonStackPos(idConnection)
  End If
  ' search yourself
  If myS = &HFF Then
    aRes = GiveGMmessage(idConnection, "BlackdProxy core turned highly unstable. Cheats might fail.", "Error")
    DoEvents
    GotKillOrder(idConnection) = False
    Exit Function
  End If
  ' search the person
  lTarget = LCase(Target)
  For y = -6 To 7
    For x = -8 To 9
      For s = 0 To 10
        tmpID = Matrix(y, x, myZ(idConnection), idConnection).s(s).dblID
        If tmpID = 0 Then
          lSquare = ""
        Else
          lSquare = LCase(GetNameFromID(idConnection, tmpID))
        End If
        If lSquare = lTarget Then
          GotKillOrderTargetID(idConnection) = Matrix(y, x, myZ(idConnection), idConnection).s(s).dblID
          GoTo continue
        End If
      Next s
    Next x
  Next y
continue:
  If (GotKillOrderTargetID(idConnection) = 0) Then
    GotKillOrder(idConnection) = False
    aRes = GiveGMmessage(idConnection, "- " & Target & " -", "Sorry, target NOT FOUND")
    DoEvents
  Else
    GotKillOrder(idConnection) = True
    GotKillOrderTargetName(idConnection) = Target
    aRes = GiveGMmessage(idConnection, ">> " & GotKillOrderTargetName(idConnection) & " <<", "Search and destroy")
    DoEvents
    ThinkTheKill (idConnection)
    DoEvents
  End If
  ProcessKillOrder = 0
  Exit Function
goterr:
  ProcessKillOrder = -1
End Function

Public Function ProcessKillOrder2(idConnection As Integer, Target As String) As Long
  Dim aRes As Long
  Dim myS As Byte
  Dim lTarget As String
  Dim sCheat As String
  Dim cPacket() As Byte
  Dim x As Long
  Dim y As Long
  Dim s As Long
  Dim tileID As Long
  Dim tmpID As Double
  Dim lSquare As String
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  If (Target = "") Or (Target = " ") Then ' use last targeted
    Target = currTargetName(idConnection)
  ElseIf Left(Target, 1) = " " Then ' delete space
    Target = Right(Target, Len(Target) - 1)
  End If
  If (TibiaVersionLong < 760) Then
    myS = MyStackPos(idConnection)
  Else
    myS = FirstPersonStackPos(idConnection)
  End If
  ' search yourself
  If myS = &HFF Then
    aRes = GiveGMmessage(idConnection, "BlackdProxy core turned highly unstable. Cheats might fail.", "Error")
    DoEvents
    GotKillOrder(idConnection) = False
    Exit Function
  End If
  ' search the person
  lTarget = LCase(Target)
  For y = -6 To 7
    For x = -8 To 9
      For s = 0 To 10
        tmpID = Matrix(y, x, myZ(idConnection), idConnection).s(s).dblID
        If tmpID = 0 Then
          lSquare = ""
        Else
          lSquare = LCase(GetNameFromID(idConnection, tmpID))
        End If
        If lSquare = lTarget Then
          GotKillOrderTargetID(idConnection) = Matrix(y, x, myZ(idConnection), idConnection).s(s).dblID
          GoTo continue
        End If
      Next s
    Next x
  Next y
continue:
  If (GotKillOrderTargetID(idConnection) = 0) Then
  Target = currTargetName(idConnection)
    GotKillOrder(idConnection) = False
    'aRes = GiveGMmessage(idConnection, "- " & target & " -", "Sorry, target NOT FOUND")
    DoEvents
  Else
    GotKillOrder(idConnection) = True
    GotKillOrderTargetName(idConnection) = Target
    'RuneMakerOptions(idConnection).autoarme4 = True
    'aRes = GiveGMmessage(idConnection, "" & GotKillOrderTargetName(idConnection) & "", "Target")
    'DoEvents
    ThinkTheKill (idConnection)
    DoEvents
  End If
 ' ProcessKillOrder2 = 0
  Exit Function
goterr:
  ProcessKillOrder2 = -1
End Function

Public Function ThinkTheKill(idConnection As Integer) As Long
  Dim aRes As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  'aRes = SendLogSystemMessageToClient(idConnection, "Attacking " & GotKillOrderTargetName(idConnection) & " ...")
  'DoEvents
  aRes = MeleeAttack(idConnection, GotKillOrderTargetID(idConnection))
  DoEvents
  ThinkTheKill = 0
  Exit Function
goterr:
  ThinkTheKill = -1
End Function

Public Function UseFluid(idConnection As Integer, byteFluid As Byte) As Long
  Dim aRes As Long
  Dim item_tileID As Long
  Dim fRes As TypeSearchItemResult2
  Dim inRes As Integer

  Dim sCheat As String
  Dim myS As Byte
  #If FinalMode Then
  On Error GoTo goterr
  #End If
If (frmHardcoreCheats.chkTotalWaste.value = True) Then 'And (TibiaVersionLong >= 773)) Then
  GoTo justdoit
End If
  If TibiaVersionLong >= 970 Then
    fRes = SearchItemWithAmount(idConnection, _
   LowByteOfLong(tileID_flask), HighByteOfLong(tileID_flask), byteFluid) 'search flask with fluid
  Else
    fRes = SearchFirstItemWithExactAmmount(idConnection, _
   LowByteOfLong(tileID_flask), HighByteOfLong(tileID_flask), byteFluid) 'search flask with fluid
  End If
  If fRes.foundCount > 0 Then
    myS = FirstPersonStackPos(idConnection)
    If myS < &HFF Then
      sCheat = "83 FF FF " & GoodHex(&H40 + fRes.bpID) & " 00 " & _
       GoodHex(fRes.slotID) & " " & FiveChrLon(tileID_flask) & " " & GoodHex(fRes.slotID) & " " & _
       MyHexPosition(idConnection) & " 63 00 " & GoodHex(myS)
       
      SafeCastCheatString "UseFluid1", idConnection, sCheat
    Else
    

      If PlayTheDangerSound = False Then
        ChangePlayTheDangerSound True
        aRes = GiveGMmessage(idConnection, "Unable to cast fluid here. Try moving or reloging!", "BlackdProxy")
        DoEvents
      End If
    End If
  Else
justdoit:
      If ((frmHardcoreCheats.chkEnhancedCheats.value = True) Or (frmHardcoreCheats.chkTotalWaste.value = True)) Then ' And (TibiaVersionLong >= 773)) Then
         myS = FirstPersonStackPos(idConnection)
         If myS < &HFF Then
        ' NEW
        sCheat = "84 FF FF 00 00 00 " & _
         GoodHex(LowByteOfLong(tileID_flask)) & " " & _
         GoodHex(HighByteOfLong(tileID_flask)) & " " & _
         GoodHex(byteFluid) & " " & SpaceID(myID(idConnection))
        
        
        SafeCastCheatString "UseFluid2", idConnection, sCheat
        UseFluid = 0
        Exit Function
      Else
              If PlayTheDangerSound = False Then
        ChangePlayTheDangerSound True
        aRes = GiveGMmessage(idConnection, "Unable to cast fluid here. Try moving or reloging!", "BlackdProxy")
        DoEvents
      End If
        
        End If
     Else
       aRes = SendSystemMessageToClient(idConnection, "can't find fluids, open new bp of fluids!")
       DoEvents
       UseFluid = -1
       Exit Function
    End If
  End If
  UseFluid = 0
goterr:
  UseFluid = -1
End Function






Public Function UsePotion(idConnection As Integer, tileID_potion As Long) As Long
  Dim aRes As Long
  Dim item_tileID As Long
  Dim fRes As TypeSearchItemResult2
  Dim inRes As Integer
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim myS As Byte
  #If FinalMode Then
  On Error GoTo goterr
  #End If
If (frmHardcoreCheats.chkTotalWaste.value = True) Then 'And (TibiaVersionLong >= 773)) Then
  GoTo justdoit
End If
  If TibiaVersionLong >= 970 Then
    fRes = SearchItemWithAmount(idConnection, _
   LowByteOfLong(tileID_potion), HighByteOfLong(tileID_potion), 0)
  Else
    fRes = SearchFirstItemWithExactAmmount(idConnection, _
   LowByteOfLong(tileID_flask), HighByteOfLong(tileID_potion), 0)
  End If
  If fRes.foundCount > 0 Then
    myS = FirstPersonStackPos(idConnection)
    If myS < &HFF Then
      sCheat = "83 FF FF " & GoodHex(&H40 + fRes.bpID) & " 00 " & _
       GoodHex(fRes.slotID) & " " & FiveChrLon(tileID_potion) & " " & GoodHex(fRes.slotID) & " " & _
       MyHexPosition(idConnection) & " 63 00 " & GoodHex(myS)
      SafeCastCheatString "UsePotion1", idConnection, sCheat
    Else
    

      If PlayTheDangerSound = False Then
        ChangePlayTheDangerSound True
        aRes = GiveGMmessage(idConnection, "Unable to cast fluid here. Try moving or reloging!", "BlackdProxy")
        DoEvents
      End If
    End If
  Else
justdoit:
      If ((frmHardcoreCheats.chkEnhancedCheats.value = True) Or (frmHardcoreCheats.chkTotalWaste.value = True)) Then 'And (TibiaVersionLong >= 773)) Then
         myS = FirstPersonStackPos(idConnection)
         If myS < &HFF Then
        ' NEW
        sCheat = "84 FF FF 00 00 00 " & _
         GoodHex(LowByteOfLong(tileID_potion)) & " " & _
         GoodHex(HighByteOfLong(tileID_potion)) & " " & _
         GoodHex(0) & " " & SpaceID(myID(idConnection))
        
        
        SafeCastCheatString "UsePotion2", idConnection, sCheat
        
        UsePotion = 0
        Exit Function
      Else
              If PlayTheDangerSound = False Then
        ChangePlayTheDangerSound True
        aRes = GiveGMmessage(idConnection, "Unable to cast potion here. Try moving or reloging!", "BlackdProxy")
        DoEvents
      End If
        
        End If
     Else
       aRes = SendSystemMessageToClient(idConnection, "can't find potions, open new bp of potions!")
       DoEvents
       UsePotion = -1
       Exit Function
    End If
  End If
  UsePotion = 0
  Exit Function
goterr:
  UsePotion = -1
End Function

Public Sub StartReconnection(idConnection As Integer)
  ' reconnect character
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  Dim tmpPORT As Long
  Dim tmpHost As String
  Dim st As Long
  Dim st2 As Long
  Dim gtnext As Long
  Dim aRes As Long
  Dim j As Long
  Dim k As Long
  Dim okcontinue As Boolean
  Dim jumpit As Boolean
  If (ReconnectionStage(idConnection) <> 0) Then
    Exit Sub
  End If
  nextReconnectionRetry(idConnection) = RETRYDELAY + GetTickCount()
  reconnectionRetryCount(idConnection) = reconnectionRetryCount(idConnection) + 1
  GotPacketWarning(idConnection) = False
  sentWelcome(idConnection) = False
  okcontinue = True
  If reconnectionRetryCount(idConnection) = 1 Then
    frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Reconnecting character " & CharacterName(idConnection) & "..."
  End If
  ReconnectionStage(idConnection) = 1
  GameConnected(idConnection) = True
  sentFirstPacket(idConnection) = True
  If frmMain.sckClientGame(idConnection).State = sckClosed Then
    frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Client of character " & CharacterName(idConnection) & " got closed! I stop the reconnection."
    frmMain.DoCloseActions idConnection
  Else
    aRes = SendSystemMessageToClient(idConnection, "Try #" & CStr(reconnectionRetryCount(idConnection)) & " - Reconnection in progress, please wait...")
    If aRes = -1 Then
        frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Can't reconnect character " & CharacterName(idConnection) & ". Client is missing. I stop the reconnection."
        frmMain.DoCloseActions idConnection
    End If
  End If
  If TibiaVersionLong >= 841 Then
    NeedToIgnoreFirstGamePacket(idConnection) = True
    MustCheckFirstClientPacket(1) = True
  End If
  GameConnected(idConnection) = False
  sentFirstPacket(idConnection) = False
  sentWelcome(idConnection) = False
  DoEvents
  UHRetryCount(idConnection) = 0
  For j = 0 To HIGHEST_BP_ID
      Backpack(idConnection, j).open = False
      Backpack(idConnection, j).cap = 0
      Backpack(idConnection, j).used = 0
      Backpack(idConnection, j).name = ""
  Next j
  For k = 1 To EQUIPMENT_SLOTS
    mySlot(idConnection, k).t1 = &H0
    mySlot(idConnection, k).t2 = &H0
    mySlot(idConnection, k).t3 = &H0
  Next k
  ConnectionBuffer(idConnection).numbytes = 0
  tmpPORT = frmMain.sckServerGame(idConnection).RemotePort
  tmpHost = frmMain.sckServerGame(idConnection).RemoteHost
  frmMain.sckServerGame(idConnection).Close
  onDepotPhase(idConnection) = 0
  frmMain.sckServerGame(idConnection).RemotePort = tmpPORT
  frmMain.sckServerGame(idConnection).RemoteHost = tmpHost
  frmMain.sckServerGame(idConnection).Connect
  DoEvents
  Exit Sub
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Reconnection function crashed. Canceling the process. Error: " & Err.Description
  frmMain.DoCloseActions idConnection
End Sub

Public Function openBP(idConnection As Integer) As Long
  #If FinalMode Then
    On Error GoTo goterr
  #End If
  Dim i As Byte
  Dim j As Byte
  Dim ec As Long
  Dim aRes As Long
  Dim lastOpen As Byte
  Dim firstAv As Byte
  Dim tileID As Long
  Dim sCheat As String
  Dim cPacket() As Byte
  Dim iRes As Integer
  Dim lastSlot As Byte
  Dim slotpos As Byte
  Dim tileID2 As Long
  Dim b1 As Byte
  Dim b2 As Byte
  lastOpen = &HFF
  For ec = HIGHEST_BP_ID To 0 Step -1
    If Backpack(idConnection, ec).open = True Then
      lastOpen = CByte(ec)
      Exit For
    End If
  Next ec
  If lastOpen = &HFF Then ' try to open main bp
    tileID = GetTheLong(mySlot(idConnection, SLOT_BACKPACK).t1, mySlot(idConnection, SLOT_BACKPACK).t2)
    If DatTiles(tileID).iscontainer = False Then
      aRes = SendLogSystemMessageToClient(idConnection, "Can't open any backpack!")
      DoEvents
      openBP = -1 ' can't open bp
      Exit Function
    Else
      ' 0A 00 82 FF FF 03 00 00 26 0B 00 00
      sCheat = "82 FF FF 03 00 00 " & GoodHex(mySlot(idConnection, SLOT_BACKPACK).t1) & " " & _
       GoodHex(mySlot(idConnection, 3).t2) & " 00 00"
      SafeCastCheatString "openBP1", idConnection, sCheat
      
      openBP = 0
      Exit Function
    End If
  Else ' try to open other bp
    firstAv = &HFF
    For j = 0 To HIGHEST_BP_ID
      If Backpack(idConnection, j).open = False Then
        firstAv = j
        Exit For
      End If
    Next j
    ' 0A 00 82 FF FF 41 00 03 26 0B 03 02
    ' 0A 00 82 FF FF 42 00 01 25 0B 01 03
    If firstAv = &HFF Then
      aRes = SendLogSystemMessageToClient(idConnection, "Can't open more backpacks!")
      DoEvents
      openBP = -1
      Exit Function
    Else
      lastSlot = Backpack(idConnection, lastOpen).cap - 1
      slotpos = &HFF
      For i = 0 To lastSlot
        b1 = Backpack(idConnection, lastOpen).item(i).t1
        b2 = Backpack(idConnection, lastOpen).item(i).t2
        tileID2 = GetTheLong(b1, b2)
        If DatTiles(tileID2).iscontainer = True Then
          slotpos = i
          Exit For
        End If
      Next i
      If slotpos = &HFF Then
        aRes = SendLogSystemMessageToClient(idConnection, "Can't find more backpacks.")
        DoEvents
        openBP = -1
        Exit Function
      Else
        sCheat = "82 FF FF " & GoodHex(&H40 + lastOpen) & " 00 " & _
         GoodHex(slotpos) & " " & GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(slotpos) & _
         " " & GoodHex(firstAv)
        SafeCastCheatString "openBP2", idConnection, sCheat
        openBP = 0
        Exit Function
      End If
    End If
  End If
  openBP = 0
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "openbp function failed : " & Err.Description
openBP = -1
End Function

Public Function GetRandomLineOf(strFileName As String) As String
  Dim fso As Scripting.FileSystemObject
  Dim fn As Integer
  Dim strLine As String
  Dim Filename As String
  Dim i As Long
  Dim p As Long
  Dim seguir As Boolean
  Dim completed As Boolean
  Dim aRes As Long
  Dim res As String
  Dim rcounter As Long
  Dim theL As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  Set fso = New Scripting.FileSystemObject
  Filename = App.Path & "\randline\" & strFileName
  res = ""
  If fso.FileExists(Filename) = True Then
    fn = FreeFile
    rcounter = 0
    Open Filename For Input As #fn
      While Not EOF(fn)
        Line Input #fn, strLine
        rcounter = rcounter + 1
      Wend
      Close #fn
      Open Filename For Input As #fn
      i = 0
      theL = randomNumberBetween(1, rcounter)
      While Not EOF(fn)
        Line Input #fn, strLine
        i = i + 1
        If i = theL Then
          res = strLine
        End If
      Wend
    Close #fn
    GetRandomLineOf = res
  Else 'file doesn't exist
    GetRandomLineOf = "FILE DO NOT EXIST"
  End If
  Exit Function
goterr:
  GetRandomLineOf = "UNKNOWN ERROR"
End Function

Public Sub TurnMe(idConnection As Integer, direction As Byte)
  Dim sCheat As String
  Dim iRes As Integer
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  sCheat = GoodHex(&H6F + direction)
  SafeCastCheatString "TurnMe1", idConnection, sCheat
  Exit Sub
goterr:
  sCheat = ""
End Sub

Public Function stringIn2(str As String) As String
  Dim s As String
  Dim l As Long
  Dim i As Long
  Dim tmp1 As String
  Dim tmp2 As Long
  Dim tmp3 As String
  If str <> "" Then
  l = Len(str)
  s = ""
  For i = 1 To l
    tmp1 = Mid(str, i, 1)
    tmp2 = CLng(AscB(tmp1))
    tmp3 = CStr(tmp2)
    If tmp2 < 100 Then
      tmp3 = "0" & tmp3
    End If
    If tmp2 < 10 Then
      tmp3 = "0" & tmp3
    End If
    s = s & tmp3
  Next i
  stringIn2 = s
  Else
    s = CStr(randomNumberBetween(0, 9)) & CStr(randomNumberBetween(0, 9)) & CStr(randomNumberBetween(0, 9))
    If randomNumberBetween(0, 1) = 0 Then
    s = s & CStr(randomNumberBetween(0, 9))
    End If
    If randomNumberBetween(0, 1) = 0 Then
    s = s & CStr(randomNumberBetween(0, 9))
    End If
    If randomNumberBetween(0, 1) = 0 Then
    s = s & CStr(randomNumberBetween(0, 9))
    End If
    If randomNumberBetween(0, 1) = 0 Then
    s = s & CStr(randomNumberBetween(0, 9))
    End If
    If randomNumberBetween(0, 1) = 0 Then
    s = s & CStr(randomNumberBetween(0, 9))
    End If
    s = s & "!!!*"
    stringIn2 = s
  End If
End Function

Public Function encriptionSumChr(ByVal chr1 As String, ByVal chr2 As String, Optional dosum As Boolean = True) As String
    Dim asc1 As Long
    Dim asc2 As Long
    Dim newchr3 As String
    Dim res As Long
    Dim i As Long
    asc1 = CLng(AscB(chr1))
    asc2 = CLng(AscB(chr2))
    If dosum = True Then
        res = asc1
        For i = 1 To asc2
            res = res + 1
            If res > 127 Then
                res = 0
            End If
        Next i
    Else
        res = asc1
        For i = 1 To asc2
            res = res - 1
            If res < 0 Then
                res = 127
            End If
        Next i
    End If
    newchr3 = Chr$(res)
    encriptionSumChr = newchr3
End Function





Public Function equipmentInfo(idConnection As Integer, lngMode As Long, strPos As String) As String
  On Error GoTo goterr
  Dim ucaseStrPos As String
  Dim idSlot As Long
  Dim strRes As String
  Dim tileID As Long
  Dim amval As Byte
  If ((idConnection < 1) Or (idConnection > MAXCLIENTS)) Then
    equipmentInfo = ""
    Exit Function
  End If
  ucaseStrPos = UCase(strPos)
  Select Case ucaseStrPos
  Case "01"
    idSlot = 1
  Case "02"
    idSlot = 2
  Case "03"
    idSlot = 3
  Case "04"
    idSlot = 4
  Case "05"
    idSlot = 5
  Case "06"
    idSlot = 6
  Case "07"
    idSlot = 7
  Case "08"
    idSlot = 8
  Case "09"
    idSlot = 9
  Case "0A"
    idSlot = 10
  Case "0B"
    idSlot = 11
  Case Else
    idSlot = 1
  End Select
  tileID = GetTheLong(mySlot(idConnection, idSlot).t1, mySlot(idConnection, idSlot).t2)
  Select Case lngMode
  Case 1
    strRes = GoodHex(mySlot(idConnection, idSlot).t1) & " " & GoodHex(mySlot(idConnection, idSlot).t2)
  Case 2
    If DatTiles(tileID).haveExtraByte = True Then
      strRes = GoodHex(mySlot(idConnection, idSlot).t3)
    Else
      strRes = ""
    End If
  Case 3
    If DatTiles(tileID).haveExtraByte = True Then
      strRes = GoodHex(mySlot(idConnection, idSlot).t3)
    Else
      strRes = "01"
    End If
  Case Else
    If DatTiles(tileID).haveExtraByte = True Then
      strRes = CStr(CLng(mySlot(idConnection, idSlot).t3))
    Else
      strRes = "1"
    End If
  End Select
  equipmentInfo = strRes
  Exit Function
goterr:
  equipmentInfo = "ERROR"
End Function

Public Function sendString(idConnection As Integer, str As String, toServer As Boolean, safeMode As Boolean) As Long
  On Error GoTo goterr
  Dim strSending As String
  Dim ub As Long
  Dim lopa As Long
  Dim aRes As Long
  Dim tmpNumA As Long
  Dim cheatpacket() As Byte
  strSending = parseVars(idConnection, str)
  If safeMode = True Then
    strSending = "00 00 " & strSending
  End If
  If GetCheatPacket(cheatpacket, strSending) = -1 Then
    aRes = SendLogSystemMessageToClient(idConnection, "Invalid packet format, blocked to avoid crash")
    DoEvents
    sendString = -1
    Exit Function
  End If
  If safeMode = True Then
    ub = UBound(cheatpacket)
    cheatpacket(0) = LowByteOfLong(ub - 1)
    cheatpacket(1) = HighByteOfLong(ub - 1)
  Else
    ub = UBound(cheatpacket)
  End If
  If ub < 1 Then
    aRes = SendLogSystemMessageToClient(idConnection, "Too short packet, blocked to avoid crash")
    DoEvents
    sendString = -1
    Exit Function
  End If
'  lopa = GetTheLong(cheatpacket(0), cheatpacket(1))
'  If (lopa <> (ub - 1)) Then
'    aRes = SendLogSystemMessageToClient(idConnection, "Packet header doesn't match with packet length, blocked to avoid crash")
'    DoEvents
'    Exit Function ' fixed in 9.38
'  End If
  
  
    ' fixed in 24.3
    
    tmpNumA = 0
    
    'i Think the new code may still fail when the packet to send is Len()>65535 bytes...
    'but i dont really care enought to test it..
    While (tmpNumA < ub + 1)
    If (ub < tmpNumA + 2) Then  'ERROR
        aRes = SendLogSystemMessageToClient(idConnection, "1Packet header doesn't match with packet length, blocked to avoid crash")
        DoEvents
        Exit Function '
    End If
    tmpNumA = tmpNumA + 2 + GetTheLong(cheatpacket(tmpNumA), cheatpacket(tmpNumA + 1))
    If (ub + 1 < tmpNumA) Then 'ERROR'
        aRes = SendLogSystemMessageToClient(idConnection, "2Packet header doesn't match with packet length, blocked to avoid crash")
        DoEvents
        Exit Function '
    End If
    Wend

  If (GameConnected(idConnection) = True) Then
    If toServer = False Then
      ' send the packet to client
      'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " > SENDING :" & frmMain.showAsStr2(cheatpacket, 2)
      frmMain.UnifiedSendToClientGame idConnection, cheatpacket
    Else
      ' send the packet to server
      'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " > SENDING :" & frmMain.showAsStr2(cheatpacket, 2)
      frmMain.UnifiedSendToServerGame idConnection, cheatpacket, True
    End If
    DoEvents
    sendString = 0
    Exit Function
  Else
    sendString = -1
    Exit Function
  End If
  Exit Function
goterr:
  sendString = -1
End Function

Public Function IDofName(idConnection As Integer, strName As String, lngOption As Long) As String
  '0 = name
  '1 = my id
  '2 = last attacked
  On Error GoTo goterr
  Dim lTarget As String
  Dim i As Long
  Dim lim As Long
  Dim TheIDisFound As Boolean
  Dim idsOnMemory
  Dim currItem As Double
  Dim currName As String
  Dim tmpID As Double
  Dim x, y, z As Long
  Dim Zstage As Long
  Dim SS As Byte
  Dim bestID As Double
  Dim bestDist As Long
  Dim currDist As Long
  Dim currZ As Long
  If ((idConnection < 1) Or (idConnection > MAXCLIENTS)) Then
    IDofName = "00 00 00 00"
    Exit Function
  End If
  Select Case lngOption
  Case 0
    lTarget = LCase(strName)
    ' search the creature on map, first on current floor
    bestDist = 10000
    bestID = 0
    Zstage = 0
    currZ = myZ(idConnection)
    Do
      If (Zstage = 0) Then
        z = currZ
        Zstage = 1
      ElseIf (Zstage = 1) Then
        If bestID <> 0 Then
          IDofName = SpaceID(bestID)
          Exit Function
        End If
        z = 0
        Zstage = 2
      Else
        z = z + 1
      End If
      For x = -8 To 9
        For y = -6 To 7
          For SS = 0 To 10
            currItem = Matrix(y, x, z, idConnection).s(SS).dblID
            If currItem <> 0 Then
              currName = LCase(NameOfID(idConnection).item(currItem))
              If currName = lTarget Then
                currDist = Abs(x) + Abs(y) + Abs(z - currZ)
                If currDist < bestDist Then
                  bestID = currItem
                End If
              End If
            End If
          Next SS
        Next y
      Next x
    Loop Until (z = 15)
    If (bestID <> 0) Then
      IDofName = SpaceID(bestID)
      Exit Function
    End If
    ' search on memory list
    TheIDisFound = False
    idsOnMemory = NameOfID(idConnection).Keys
    lim = NameOfID(idConnection).Count - 1
    For i = 0 To lim
      currItem = CDbl(idsOnMemory(i))
      currName = LCase(NameOfID(idConnection).item(currItem))
      If currName = lTarget Then
        tmpID = currItem
        TheIDisFound = True
        Exit For
      End If
    Next i
    If TheIDisFound = True Then
      IDofName = SpaceID(tmpID)
      Exit Function
    Else
      IDofName = "00 00 00 00"
      Exit Function
    End If
  Case 1
    IDofName = SpaceID(myID(idConnection))
    Exit Function
  Case Else
    IDofName = SpaceID(currTargetID(idConnection))
    Exit Function
  End Select
goterr:
  IDofName = "00 00 00 00"
  Exit Function
End Function

Public Function storeVar(idConnection As Integer, strRawStore As String)
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  Dim aRes As Long
  Dim strVariable As String
  Dim strValue As String
  Dim strEqualPos As String
  Dim strEspecial As String
  Dim strRealVar As String
  If idConnection < 1 Then
    storeVar = -1
    Exit Function
  End If
  strEqualPos = InStr(1, strRawStore, "=")
  If strEqualPos = 0 Then
    storeVar = -1
    Exit Function
  End If
  strEspecial = "(" & CStr(idConnection) & ")"
  strVariable = Trim$(LCase(Left$(strRawStore, strEqualPos - 1)))
  If (Len(strVariable) > 2) Then
    If Mid$(strVariable, 2, 1) = "_" Then
      strEspecial = ""
    End If
  End If
  strValue = parseVars(idConnection, Trim$(Right$(strRawStore, Len(strRawStore) - strEqualPos)))
  strRealVar = strVariable & strEspecial
  AddUserVar strRealVar, strValue
  If publicDebugMode = True Then
    If strEspecial = "" Then
      aRes = SendLogSystemMessageToClient(idConnection, "GLOBAL Variable $" & strVariable & "$ is now :" & strValue)
    Else
      aRes = SendLogSystemMessageToClient(idConnection, "Local Variable at client #" & CStr(idConnection) & " $" & strVariable & "$ is now :" & strValue)
    End If
    DoEvents
  End If
  storeVar = 0
  Exit Function
goterr:
  storeVar = -1
End Function

Public Function safeLong(strString As String) As Long
  On Error GoTo goterr
  Dim lngRes As Long
  lngRes = CLng(strString)
  safeLong = lngRes
  Exit Function
goterr:
  safeLong = 0
End Function

Public Function safeDouble(strString As String) As Double
  On Error GoTo goterr
  Dim dblRes As Double
  dblRes = CDbl(strString)
  safeDouble = dblRes
  Exit Function
goterr:
  safeDouble = 0
End Function

Public Function NumericValueOfExpresion(strExp As String) As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  Dim strFilteredExp As String
  Dim lngSymbolPosition As String
  Dim strSymbol As String
  Dim strOp1 As String
  Dim strOp2 As String
  Dim lngOp1 As Long
  Dim lngOp2 As Long
  Dim strRes As Long
  strFilteredExp = Replace(strExp, " ", "")
  lngSymbolPosition = InStr(1, strFilteredExp, "+")
  If lngSymbolPosition = 0 Then
    lngSymbolPosition = InStr(1, strFilteredExp, "-")
  End If
  If lngSymbolPosition = 0 Then
    lngSymbolPosition = InStr(1, strFilteredExp, "*")
  End If
  If lngSymbolPosition = 0 Then
    lngSymbolPosition = InStr(1, strFilteredExp, "/")
  End If
  If lngSymbolPosition = 0 Then
    NumericValueOfExpresion = 0
    Exit Function
  End If
  strSymbol = Mid$(strFilteredExp, lngSymbolPosition, 1)
  strOp1 = Left$(strFilteredExp, lngSymbolPosition - 1)
  strOp2 = Right$(strFilteredExp, Len(strFilteredExp) - lngSymbolPosition)
  lngOp1 = safeLong(strOp1)
  lngOp2 = safeLong(strOp2)
  strRes = 0
  Select Case strSymbol
  Case "+"
    strRes = lngOp1 + lngOp2
  Case "-"
    strRes = lngOp1 - lngOp2
  Case "*"
    strRes = lngOp1 * lngOp2
  Case "/"
    If lngOp2 <> 0 Then
       strRes = lngOp1 / lngOp2
    End If
  End Select
  NumericValueOfExpresion = strRes
  Exit Function
goterr:
  NumericValueOfExpresion = 0
End Function

Public Function HexConverter(strExpr As String, intOpType As Long) As String
  On Error GoTo goterr
  Dim strFilteredExpr As String
  Dim strRes As String
  strFilteredExpr = UCase(Trim$(strExpr))
  strRes = ""
  Select Case intOpType
  Case 1 'numbertohex1
    strRes = GoodHex(CLng(strFilteredExpr))
  Case 2 'numbertohex2
    strRes = FiveChrLon(CLng(strFilteredExpr))
  Case 3 'hex1tonumber
   strRes = CStr(GetTheByteFromTwoChr(strFilteredExpr))
  Case 4 'hex2tonumber
    strRes = CStr(GetTheLongFromFiveChr(strFilteredExpr))
  End Select
  HexConverter = strRes
  Exit Function
goterr:
  Select Case intOpType
  Case 1 'numbertohex1
    HexConverter = "00"
    Exit Function
  Case 2 'numbertohex2
    HexConverter = "00 00"
    Exit Function
  Case 3 'hex1tonumber
    HexConverter = 0
    Exit Function
  Case 4 'hex2tonumber
    HexConverter = 0
    Exit Function
  End Select
  HexConverter = 0
End Function


Public Function NameOfHexID(idConnection As Integer, strName As String) As String
On Error GoTo goterr
  Dim dblID As Double
  Dim strRes As String
  Dim strFiltered As String
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim b4 As Byte
  Dim s1 As String
  Dim s2 As String
  Dim s3 As String
  Dim s4 As String

  strFiltered = Trim$(strName)
  If Len(strFiltered) <> 11 Then
    NameOfHexID = ""
    Exit Function
  End If
  s1 = Mid$(strFiltered, 1, 2)
  s2 = Mid$(strFiltered, 4, 2)
  s3 = Mid$(strFiltered, 7, 2)
  s4 = Mid$(strFiltered, 10, 2)
  b1 = GetTheByteFromTwoChr(s1)
  b2 = GetTheByteFromTwoChr(s2)
  b3 = GetTheByteFromTwoChr(s3)
  b4 = GetTheByteFromTwoChr(s4)
  dblID = FourBytesDouble(b1, b2, b3, b4)
 
  strRes = GetNameFromID(idConnection, dblID)
  NameOfHexID = strRes
  Exit Function
goterr:
  NameOfHexID = ""
End Function

Public Function getScreenshotname() As String
  Dim strbase As String
  strbase = ""
  If lngNextScreenshotNumber < 10 Then
    strbase = "00000" & CStr(lngNextScreenshotNumber)
  ElseIf lngNextScreenshotNumber < 100 Then
    strbase = "0000" & CStr(lngNextScreenshotNumber)
  ElseIf lngNextScreenshotNumber < 1000 Then
    strbase = "000" & CStr(lngNextScreenshotNumber)
  ElseIf lngNextScreenshotNumber < 10000 Then
    strbase = "00" & CStr(lngNextScreenshotNumber)
  ElseIf lngNextScreenshotNumber < 100000 Then
    strbase = "0" & CStr(lngNextScreenshotNumber)
  Else
    strbase = CStr(lngNextScreenshotNumber)
  End If
  lngNextScreenshotNumber = lngNextScreenshotNumber + 1
  getScreenshotname = "screenshots\ss" & strbase & ".bmp"
End Function

Public Function DoAutocombo(idConnection As Integer, strChannel As String) As Long
  #If FinalMode = 1 Then
  On Error GoTo goterr
  #End If
  Dim useChannel As String
  Dim aRes As Long
  Dim strToSend As String
  Dim strTargetName As String
  If strChannel = "" Then
    useChannel = parseVars(idConnection, "$lastrecchannelid$")
    If useChannel = "05 00" Then
        aRes = SendLogSystemMessageToClient(idConnection, "First you need to talk in a channel ...")
        DoEvents
        DoAutocombo = -1
        Exit Function
    End If
  ElseIf GetTheLongFromFiveChr(Trim$(strChannel)) = -1 Then
    useChannel = parseVars(idConnection, "$lastrecchannelid$")
    If useChannel = "05 00" Then
        aRes = SendLogSystemMessageToClient(idConnection, "First you need to talk in a channel ...")
        DoEvents
        DoAutocombo = -1
        Exit Function
    End If
  Else
    useChannel = Trim$(strChannel)
  End If
  ' 18 00 96 05 04 00 12 00 2E 2E 2E 2E 2E 2E 2E 2E 2E 2E 2E 2E 2E 2E 2E 2E 2E 2E
  strTargetName = NameOfHexID(idConnection, IDofName(idConnection, "", 2))
  
  If LCase(strTargetName) = LCase(CharacterName(idConnection)) Then
    aRes = SendLogSystemMessageToClient(idConnection, "Last target = " & strTargetName & " (yourself), avoiding this autocombo ...")
    DoEvents
    DoAutocombo = -1
    Exit Function
  End If
  If strTargetName = "" Then
    aRes = SendLogSystemMessageToClient(idConnection, "You have no last target! first click someone.")
    DoEvents
    DoAutocombo = -1
    Exit Function
  End If
  If TibiaVersionLong < 820 Then
     strToSend = "96 05 "
  Else
     strToSend = "96 07 "
  End If
  strToSend = strToSend & useChannel & " " & Hexarize2(frmHardcoreCheats.txtOrder.Text & " " & strTargetName)
  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "> " & strToSend
  aRes = sendString(idConnection, strToSend, True, True)
  DoAutocombo = 0
  Exit Function
goterr:
  LogOnFile "errors.txt", "Unexpected error at DoAutocombo : (" & CStr(Err.Number) & ") : " & Err.Description
  DoAutocombo = -1
  Exit Function
End Function

Public Function CountOnFloor(idConnection As Integer, strFloor As String, _
 Optional blnCountPKs As Boolean = False, Optional blnCountGMs As Boolean = False, _
 Optional blnCountMeleeTargets As Boolean = False) As Long
  Dim lngFloor As Long
  Dim lngCount As Long
  Dim lngRealFloor As Long
  Dim x As Long
  Dim y As Long
  Dim z As Long
  Dim s As Long
  Dim tileID As Long
  Dim nameofgivenID As String
  lngFloor = CLng(Trim$(strFloor))
  z = myZ(idConnection) + lngFloor
  If ((z > 15) Or (z < 0)) Then
    CountOnFloor = 0
    Exit Function
  End If
  lngCount = 0
  For x = -8 To 9
    For y = -6 To 7
        For s = 1 To 10
          tileID = GetTheLong(Matrix(y, x, z, idConnection).s(s).t1, Matrix(y, x, z, idConnection).s(s).t2)
          If tileID = 97 Then
           nameofgivenID = GetNameFromID(idConnection, Matrix(y, x, z, idConnection).s(s).dblID)
            If ((isMelee(idConnection, nameofgivenID) = True) Or (isHmm(idConnection, nameofgivenID) = True)) And (frmWarbot.IsAutoHealFriend(LCase(nameofgivenID)) = False) Then
              If (blnCountMeleeTargets = True And nameofgivenID <> CharacterName(idConnection)) Then
                lngCount = lngCount + 1
              End If
            Else
            If (frmWarbot.IsAutoHealFriend(LCase(nameofgivenID)) = False) Then
              If (nameofgivenID <> CharacterName(idConnection)) Then
                If (IsGM(nameofgivenID) = True) Then
                  If (blnCountGMs = True) Then
                    lngCount = lngCount + 1
                  End If
                Else
                  If (blnCountPKs = True) Then
                    lngCount = lngCount + 1
                  End If
                End If
              End If
            End If
            End If
          ElseIf tileID = 0 Then
            Exit For
          End If
        Next s
    Next y
  Next x
  
  CountOnFloor = lngCount
  Exit Function
goterr:
  CountOnFloor = 0
End Function

Public Function GetStatusBit(idConnection As Integer, strBit As String) As String
  On Error GoTo goterr
  Dim lngBit As Long
  lngBit = CLng(Trim$(strBit))
  If lngBit > 0 And lngBit < 16 Then
    GetStatusBit = Mid$(StatusBits(idConnection), lngBit, 1)
  Else
    GetStatusBit = "?"
  End If
  Exit Function
goterr:
  GetStatusBit = "?"
End Function

Public Function CountTheItemsForUser(idConnection As Integer, strTile As String) As Long
    #If FinalMode = 1 Then
        On Error GoTo goterr
    #End If
    Dim res As Long
    Dim lngLen As Long
    Dim tmpLong As Long
    Dim lngTotal As Long
    Dim am As Byte
    Dim b1 As Byte
    Dim b2 As Byte
    Dim i As Long
    Dim amSlot As Byte
    lngLen = Len(strTile)
    lngTotal = 0
    If lngLen = 5 Then
        tmpLong = GetTheLongFromFiveChr(strTile)
        b1 = LowByteOfLong(tmpLong)
        b2 = HighByteOfLong(tmpLong)
        lngTotal = SearchAmmount(idConnection, b1, b2)
    ElseIf lngLen = 8 Then
        tmpLong = GetTheLongFromFiveChr(Left$(strTile, 5))
        b1 = LowByteOfLong(tmpLong)
        b2 = HighByteOfLong(tmpLong)
        am = GetTheByteFromTwoChr(Right$(strTile, 2))
        lngTotal = SearchExactAmmount(idConnection, b1, b2, am)
    End If
    CountTheItemsForUser = lngTotal
    Exit Function
goterr:
    CountTheItemsForUser = 0
End Function


Public Function SafeGiveIDinfo(idConnection As Integer, strName As String, whatinfo As Integer) As Long
'On Error GoTo goterr
  Dim dblID As Double
  Dim strRes As String
  Dim strFiltered As String
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim b4 As Byte
  Dim s1 As String
  Dim s2 As String
  Dim s3 As String
  Dim s4 As String

  strFiltered = Trim$(strName)
  If Len(strFiltered) <> 11 Then
    SafeGiveIDinfo = 0
    Exit Function
  End If
  s1 = Mid$(strFiltered, 1, 2)
  s2 = Mid$(strFiltered, 4, 2)
  s3 = Mid$(strFiltered, 7, 2)
  s4 = Mid$(strFiltered, 10, 2)
  b1 = GetTheByteFromTwoChr(s1)
  b2 = GetTheByteFromTwoChr(s2)
  b3 = GetTheByteFromTwoChr(s3)
  b4 = GetTheByteFromTwoChr(s4)
  dblID = FourBytesDouble(b1, b2, b3, b4)
  Select Case whatinfo
  Case 1
    SafeGiveIDinfo = CLng(GetHPFromID(idConnection, dblID))
  Case 2
    SafeGiveIDinfo = CLng(GetDirectionFromID(idConnection, dblID))
  Case Else
    SafeGiveIDinfo = 0
  End Select
  Exit Function
goterr:
  SafeGiveIDinfo = 0
End Function

Public Function SayInTrade(ByRef idConnection As Integer, ByVal strMsg As String) As Long
   ' 09 00 96 04 05 00 74 72 61 64 65
   ' only for tibia 8.2 +
   
   
   ' Tibia 8.72:
   ' 09 00 96 0B 05 00 74 72 61 64 65
   
    Dim cPacket() As Byte
    Dim sCheat As String
    Dim tmplng As Long
    Dim b1 As Byte
    Dim b2 As Byte
    Dim b3 As Byte
    Dim b4 As Byte
    Dim inRes As Integer
    #If FinalMode Then
    On Error GoTo goterr
    #End If
    If (doingTrade(idConnection) = False) Then
        SayInTrade = 0
        Exit Function
    End If
    tmplng = Len(strMsg) + 4
    b1 = LowByteOfLong(tmplng)
    b2 = HighByteOfLong(tmplng)
    If TibiaVersionLong >= 1036 Then
      ' 09 00 96 0C 05 00 61 61 61 61 61
      sCheat = GoodHex(b1) & " " & GoodHex(b2) & " 96 0C " & Hexarize2(strMsg)
    ElseIf TibiaVersionLong >= 872 Then
      sCheat = GoodHex(b1) & " " & GoodHex(b2) & " 96 0B " & Hexarize2(strMsg)
    Else
      sCheat = GoodHex(b1) & " " & GoodHex(b2) & " 96 04 " & Hexarize2(strMsg)
    End If
    
    
    inRes = GetCheatPacket(cPacket, sCheat)
    ' debug.Print frmMain.showAsStr(cPacket, True)
    frmMain.UnifiedSendToServerGame idConnection, cPacket, True
    DoEvents
    SayInTrade = 0
    Exit Function
goterr:
    SayInTrade = -1
End Function


Public Function SellInTrade(ByRef idConnection As Integer, ByVal strMsg As String) As Long
   ' 05 00 7B CA 0C 00 04
   ' only for tibia 8.2 +
    Dim cPacket() As Byte
    Dim sCheat As String
    Dim tmplng As Long
    Dim inRes As Integer
    Dim b1 As Byte
    Dim b2 As Byte
    Dim strAm As String
    Dim aRes As Long
    Dim tradeAm As Byte
    Dim remainingAm As Long
    Dim nextgtc As Long
    Dim gtc As Long
    
    #If FinalMode Then
    On Error GoTo goterr
    #End If
    If (doingTrade2(idConnection) = False) Then
        aRes = SendLogSystemMessageToClient(idConnection, "Blackd Proxy is unable to sell that: trade window is not opened")
        DoEvents
        SellInTrade = 0
        Exit Function
    End If
    If Len(strMsg) < 7 Then
        aRes = SendLogSystemMessageToClient(idConnection, "Blackd Proxy is unable to sell that: bad format in parameters (too short)")
        DoEvents
        SellInTrade = 0
        Exit Function
    End If
    If Mid(strMsg, 6, 1) <> ":" Then
        aRes = SendLogSystemMessageToClient(idConnection, "Blackd Proxy is unable to sell that: bad format in parameters (missing separator)")
        DoEvents
        SellInTrade = 0
        Exit Function
    End If
    strAm = Right$(strMsg, Len(strMsg) - 6)
    If IsNumeric(strAm) = False Then
        aRes = SendLogSystemMessageToClient(idConnection, "Blackd Proxy is unable to sell that: bad format in parameters (bad amount)")
        DoEvents
        SellInTrade = 0
        Exit Function
    End If

    
    If CLng(strAm) <= 0 Then
        aRes = SendLogSystemMessageToClient(idConnection, "Blackd Proxy is unable to sell that: amount should be positive")
        SellInTrade = 0
        Exit Function
    End If
    
    ClientExecutingLongCommand(idConnection) = True
    
    remainingAm = CLng(strAm)
    
    nextgtc = 0
    Do
        gtc = GetTickCount()
        If (gtc >= nextgtc) Then
            nextgtc = gtc + randomNumberBetween(400, 700)
            If ClientExecutingLongCommand(idConnection) = False Then
                ' stop
                remainingAm = 0
            End If
            If remainingAm > 0 Then
                If remainingAm >= 100 Then
                    tradeAm = CByte(CLng(100))
                    remainingAm = remainingAm - 100
                Else
                    tradeAm = CByte(CLng(remainingAm))
                    remainingAm = 0
                End If
                If publicDebugMode = True Then
                    aRes = SendLogSystemMessageToClient(idConnection, "Long Command - Now buying " & CStr(CLng(tradeAm)) & " units. Remaining = " & CStr(remainingAm))
                    DoEvents
                End If
                tmplng = Len(strMsg) + 4
                b1 = FromHexToDec(Mid(strMsg, 1, 1)) * 16 + FromHexToDec(Mid(strMsg, 2, 1))
                b2 = FromHexToDec(Mid(strMsg, 4, 1)) * 16 + FromHexToDec(Mid(strMsg, 5, 1))
                If TibiaVersionLong >= 871 Then
                    sCheat = "07 00 7B " & GoodHex(b1) & " " & GoodHex(b2) & " 00 " & GoodHex(tradeAm) & " 00" ' sell equipped
                ElseIf TibiaVersionLong >= 830 Then
                    sCheat = "07 00 7B " & GoodHex(b1) & " " & GoodHex(b2) & " 00 " & GoodHex(tradeAm) & " 01 01"
                Else
                    sCheat = "05 00 7B " & GoodHex(b1) & " " & GoodHex(b2) & " 00 " & GoodHex(tradeAm)
                End If
                inRes = GetCheatPacket(cPacket, sCheat)
                frmMain.UnifiedSendToServerGame idConnection, cPacket, True
                DoEvents
            End If
        Else
            DoEvents
        End If
    Loop Until remainingAm <= 0
    
    ClientExecutingLongCommand(idConnection) = False
    
    
    
    
    
    
    SellInTrade = 0
    Exit Function
goterr:
    aRes = SendLogSystemMessageToClient(idConnection, "Blackd Proxy is unable to sell that: unknown error, vb err code " & CStr(Err.Number))
    DoEvents
    SellInTrade = -1
End Function

Public Function BuyInTrade(ByRef idConnection As Integer, ByVal strMsg As String) As Long
   ' 05 00 7A CA 0C 00 04
   ' only for tibia 8.2 +
    Dim cPacket() As Byte
    Dim sCheat As String
    Dim tmplng As Long
    Dim inRes As Integer
    Dim b1 As Byte
    Dim b2 As Byte
    Dim strAm As String
    Dim aRes As Long
    Dim tradeAm As Byte
    Dim remainingAm As Long
    Dim gtc As Long
    Dim nextgtc As Long
    #If FinalMode Then
    On Error GoTo goterr
    #End If
    If (doingTrade2(idConnection) = False) Then
        aRes = SendLogSystemMessageToClient(idConnection, "Blackd Proxy is unable to buy that: trade window is not opened")
        DoEvents
        BuyInTrade = 0
        Exit Function
    End If
    If Len(strMsg) < 7 Then
        aRes = SendLogSystemMessageToClient(idConnection, "Blackd Proxy is unable to buy that: bad format in parameters (too short)")
        DoEvents
        BuyInTrade = 0
        Exit Function
    End If
    If Mid(strMsg, 6, 1) <> ":" Then
        aRes = SendLogSystemMessageToClient(idConnection, "Blackd Proxy is unable to buy that: bad format in parameters (missing separator)")
        DoEvents
        BuyInTrade = 0
        Exit Function
    End If
    strAm = Right$(strMsg, Len(strMsg) - 6)
    If IsNumeric(strAm) = False Then
        aRes = SendLogSystemMessageToClient(idConnection, "Blackd Proxy is unable to buy that: bad format in parameters (bad amount)")
        DoEvents
        BuyInTrade = 0
        Exit Function
    End If
    
    If CLng(strAm) <= 0 Then
        aRes = SendLogSystemMessageToClient(idConnection, "Blackd Proxy is unable to buy that: amount should be positive")
        BuyInTrade = 0
        Exit Function
    End If
    
    ClientExecutingLongCommand(idConnection) = True
    
    remainingAm = CLng(strAm)
    nextgtc = 0
    Do
        gtc = GetTickCount()
        If (gtc >= nextgtc) Then
            nextgtc = gtc + randomNumberBetween(400, 700)
            If ClientExecutingLongCommand(idConnection) = False Then
                ' stop
                remainingAm = 0
            End If
            If remainingAm > 0 Then
                If remainingAm >= 100 Then
                    tradeAm = CByte(CLng(100))
                    remainingAm = remainingAm - 100
                Else
                    tradeAm = CByte(CLng(remainingAm))
                    remainingAm = 0
                End If
                If publicDebugMode = True Then
                    aRes = SendLogSystemMessageToClient(idConnection, "Long Command - Now buying " & CStr(CLng(tradeAm)) & " units. Remaining = " & CStr(remainingAm))
                    DoEvents
                End If
                tmplng = Len(strMsg) + 4
                b1 = FromHexToDec(Mid(strMsg, 1, 1)) * 16 + FromHexToDec(Mid(strMsg, 2, 1))
                b2 = FromHexToDec(Mid(strMsg, 4, 1)) * 16 + FromHexToDec(Mid(strMsg, 5, 1))
                If TibiaVersionLong >= 830 Then
                    sCheat = "07 00 7A " & GoodHex(b1) & " " & GoodHex(b2) & " 00 " & GoodHex(tradeAm) & " 00 00"
                Else
                    sCheat = "05 00 7A " & GoodHex(b1) & " " & GoodHex(b2) & " 00 " & GoodHex(tradeAm)
                End If
                inRes = GetCheatPacket(cPacket, sCheat)
                frmMain.UnifiedSendToServerGame idConnection, cPacket, True
                DoEvents
            End If
        Else
            DoEvents
        End If
    Loop Until remainingAm <= 0
    
    ClientExecutingLongCommand(idConnection) = False
    BuyInTrade = 0
    Exit Function
goterr:
    aRes = SendLogSystemMessageToClient(idConnection, "Blackd Proxy is unable to buy that: unknown error, vb err code " & CStr(Err.Number))
    DoEvents
    BuyInTrade = -1
End Function

Private Function ProcessVarRandom(strRight As String) As Long
    On Error GoTo goterr
    Dim pos As Long
    Dim lpart As String
    Dim rpart As String
    Dim lon1 As Long
    Dim lon2 As Long
    pos = InStr(1, strRight, ">")
    lpart = Left$(strRight, pos - 1)
    rpart = Right$(strRight, Len(strRight) - pos)
    lon1 = CLng(lpart)
    lon2 = CLng(rpart)
    ProcessVarRandom = randomNumberBetween(lon1, lon2)
    Exit Function
goterr:
    ProcessVarRandom = 0
End Function

Public Sub ChaotizeNextMaxAttackTime(ByVal idConnection As Integer)
    On Error GoTo goterr
    Dim dblBase As Double
    Dim dblPercent As Double
    Dim dblMaxIncrease As Double
    Dim dblMaxDecrease As Double
    Dim lngChaos As Long
    dblBase = CDbl(maxAttackTime(idConnection))
    dblPercent = 10
    dblMaxIncrease = dblBase * ((100 + dblPercent) / 100)
    dblMaxDecrease = dblBase * ((100 - dblPercent) / 100)
    lngChaos = randomNumberBetween(CLng(Round(dblMaxDecrease)), CLng(Round(dblMaxIncrease)))
    maxAttackTimeCHAOS(idConnection) = lngChaos
    Exit Sub
goterr:
    maxAttackTimeCHAOS(idConnection) = maxAttackTime(idConnection)
End Sub

Public Function IngameCheck2(ByVal idConnection As Integer, ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
  '09 00 8C 6F 7E 06 7E 07 67 00 00
  Dim SS As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim xdif As Long
  Dim ydif As Long
  Dim tileID As Long
  Dim aRes As Long
  Dim sCheat As String
  Dim cPacket() As Byte
  Dim inRes As Integer
  Dim SOPT As Byte
  Dim currentPriority As Long
  Dim bestP As Long
  'Dim strname As String
  'strname = ""
  lastIngameCheck(idConnection) = ""
  lastIngameCheckTileID(idConnection) = "00 00"
  SS = &HB
  xdif = x - myX(idConnection)
  ydif = y - myY(idConnection)
  bestP = 10000
  tileID = 0
  For SOPT = 0 To 10
    b1 = Matrix(ydif, xdif, z, idConnection).s(SOPT).t1
    b2 = Matrix(ydif, xdif, z, idConnection).s(SOPT).t2
    tileID = GetTheLong(b1, b2)
    If SOPT = 0 Then
      currentPriority = 1000
    Else
      If ((tileID >= 97) And (tileID <= 99)) Then
        currentPriority = 1
        ' strname = GetNameFromID(idConnection, Matrix(ydif, xdif, z, idConnection).s(SOPT).dblID)
        SS = SOPT
        'Debug.Print GoodHex(SOPT) & ": " & GoodHex(b1) & " " & GoodHex(b2) & " priority " & CStr(currentPriority) & " - " & strname
        Exit For
      ElseIf tileID = 0 Then
        currentPriority = 0
      Else
        currentPriority = 2 + DatTiles(tileID).stackPriority
      End If
    End If
    If ((tileID = 0) Or (currentPriority = 0)) Then
      Exit For
    Else
      If currentPriority <= bestP Then
         bestP = currentPriority
         SS = SOPT
      End If
    End If
    'Debug.Print GoodHex(SOPT) & ": " & GoodHex(b1) & " " & GoodHex(b2) & " priority " & CStr(currentPriority) & " - " & strname
  Next SOPT
  b1 = Matrix(ydif, xdif, z, idConnection).s(SS).t1
  b2 = Matrix(ydif, xdif, z, idConnection).s(SS).t2
  tileID = GetTheLong(b1, b2)
  
  If tileID = 97 Then
    b1 = &H63
    b2 = &H0
    tileID = 99
  End If
  lastIngameCheckTileID(idConnection) = FiveChrLon(tileID)
  sCheat = "09 00 8C " & FiveChrLon(x) & " " & FiveChrLon(y) & " " & GoodHex(CByte(z)) & " " & GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(SS)
    

  inRes = GetCheatPacket(cPacket, sCheat)
  frmMain.UnifiedSendToServerGame idConnection, cPacket, True
  DoEvents
  IngameCheck2 = 0
End Function
Public Function IngameCheck(ByVal idConnection As Integer, ByVal strXYZ As String) As Long
    #If FinalMode = 1 Then
        On Error GoTo goterr
    #End If
    Dim x As String
    Dim y As String
    Dim z As String
    Dim aRes As Long
    Dim pos1 As Long
    Dim pos2 As Long
    Dim lonX As Long
    Dim lonY As Long
    Dim lonZ As Long
    lastIngameCheck(idConnection) = ""
    lastIngameCheckTileID(idConnection) = "00 00"
    If strXYZ = "" Then
        aRes = GiveGMmessage(idConnection, "Expected format: exiva check x,y,z", "Exiva check error")
        DoEvents
        IngameCheck = -1
        Exit Function
    End If
    strXYZ = Trim$(strXYZ)
    pos1 = InStr(1, strXYZ, ".")
    If pos1 > 0 Then
        aRes = GiveGMmessage(idConnection, "Illegal symbols in parameters", "Exiva check error")
        DoEvents
        IngameCheck = -1
        Exit Function
    End If
    pos1 = InStr(1, strXYZ, "-")
    If pos1 > 0 Then
        aRes = GiveGMmessage(idConnection, "Illegal symbols in parameters", "Exiva check error")
        DoEvents
        IngameCheck = -1
        Exit Function
    End If
    pos1 = InStr(1, strXYZ, ",")
    If pos1 <= 0 Then
        aRes = GiveGMmessage(idConnection, "Expected format: exiva check x,y,z", "Exiva check error")
        DoEvents
        IngameCheck = -1
        Exit Function
    End If
    pos2 = InStr(pos1 + 1, strXYZ, ",")
    If pos2 <= 0 Then
        aRes = GiveGMmessage(idConnection, "Expected format: exiva check x,y,z", "Exiva check error")
        DoEvents
        IngameCheck = -1
        Exit Function
    End If
    x = Trim$(Left$(strXYZ, pos1 - 1))
    y = Trim$(Mid$(strXYZ, pos1 + 1, pos2 - pos1 - 1))
    z = Trim$(Right$(strXYZ, Len(strXYZ) - pos2))
    If (IsNumeric(x) = False) Then
        aRes = GiveGMmessage(idConnection, "x was not numeric!", "Exiva check error")
        DoEvents
        IngameCheck = -1
        Exit Function
    End If
    If (IsNumeric(y) = False) Then
        aRes = GiveGMmessage(idConnection, "y was not numeric!", "Exiva check error")
        DoEvents
        IngameCheck = -1
        Exit Function
    End If
    If (IsNumeric(z) = False) Then
        aRes = GiveGMmessage(idConnection, "z was not numeric!", "Exiva check error")
        DoEvents
        IngameCheck = -1
        Exit Function
    End If
    lonX = CLng(x)
    lonY = CLng(y)
    lonZ = CLng(z)
    'ares = GiveGMmessage(idConnection, "x='" & CStr(lonX) & "' y='" & CStr(lonY) & "' z='" & CStr(lonZ) & "'", "Exiva check error")
    'DoEvents
    ' -6 To 7, -8 To 9
    If (lonX > myX(idConnection) + 8) Or _
     (lonX < myX(idConnection) - 7) Or _
     (lonY > myY(idConnection) + 6) Or _
     (lonY < myY(idConnection) - 5) Or _
     (lonZ <> myZ(idConnection)) Then
        aRes = GiveGMmessage(idConnection, "The requested x,y,z were out of range. Check cancelled.", "Exiva check error")
        DoEvents
        IngameCheck = -1
        Exit Function
    End If
    IngameCheck = IngameCheck2(idConnection, x, y, z)
    Exit Function
goterr:
    aRes = GiveGMmessage(idConnection, "Unexpected error " & CStr(Err.Number), "Exiva check error")
    DoEvents
    IngameCheck = -1
    Exit Function
End Function

Public Function TibiaDatExists() As Boolean
    On Error GoTo goterr
  Dim tibiadathere As String
  Dim fs As FileSystemObject
  Set fs = New Scripting.FileSystemObject
'  If configPath = "" Then
'    tibiadathere = App.path & "\tibia.dat"
'  Else
'    tibiadathere = App.path & "\" & configPath & "\tibia.dat"
'  End If
  tibiadathere = TibiaExePathWITHTIBIADAT
  
  If fs.FileExists(tibiadathere) = False Then
    TibiaDatExists = False
    Exit Function
  End If
  TibiaDatExists = True
  Exit Function
goterr:
  DBGtileError = "Error number = " & CStr(Err.Number) & vbCrLf & " ; Error description = " & Err.Description & " ; Path = " & tibiadathere
  TibiaDatExists = False
End Function


Public Function UseItemWithAmount(idConnection As Integer, strTile As String) As Long
'HHBCODE
'basically a copy of CountTheItemsForUser
'will use the item and return true (if item with exact ammount found), or false.
'strTile: "DB 0F,100" == use DB 0F with 100 ammount//len9
'strTile: "DB 0F,10" == use DB 0F with 10 ammount//len8
'strTile: "DB 0F,1" == use DB 0F with 1 ammount//len7

'strTile: "DB 0F FF,100" == use DB 0F with 100 ammount//len12
'strTile: "DB 0F FF,10" == use DB 0F with 10 ammount//len11
'strTile: "DB 0F FF,1" == use DB 0F with 1 ammount//len10

    #If FinalMode = 1 Then
        On Error GoTo goterr
    #End If
    Dim res As Long
    Dim tmpLong As Long
    Dim Success As Boolean
    Dim am As Byte
    Dim b1 As Byte
    Dim b2 As Byte
    Dim i As Long
    Dim amSlot As Byte
    'lngLen = Len(strTile)
    Success = False

    'im trying to find a way to extract the comma,number...
    Dim TmpAmmountString As String
    Dim Ammount As Long
    Dim mint As Long
    
    TmpAmmountString = Mid$(strTile, (InStr(strTile, ",") + 1))
    Ammount = CLng(TmpAmmountString)
    'now im trying to remove the ,X[X[X]] from strTile..
    mint = Len(strTile)
    mint = mint - (Len(TmpAmmountString) + 1)
    strTile = Left(strTile, mint)
        tmpLong = GetTheLongFromFiveChr(strTile)
        b1 = LowByteOfLong(tmpLong)
        b2 = HighByteOfLong(tmpLong)
        
        Success = UseItemWithAmountX(idConnection, b1, b2, CByte(Ammount))
    If Success = True Then
    UseItemWithAmount = 0
    Else
    UseItemWithAmount = -1
    End If
    Exit Function
goterr:
    UseItemWithAmount = -1
End Function


Public Function GetSizeOfFile(ByVal strFilePath As String) As Long
    On Error GoTo goterr
    Dim fn As Integer
    Dim res As Long
    fn = FreeFile
    Open strFilePath For Binary As fn
    res = LOF(fn)
    Close fn
    GetSizeOfFile = res
    Exit Function
goterr:
    GetSizeOfFile = -1
End Function


Public Function GetDATEOfFile(ByVal strFilePath As String) As Date
    On Error GoTo goterr
    Dim strLine As String
    strLine = "Dim FileTS As Date"
    Dim FileTS As Date
    strLine = "FileTS = FileDateTime(""" & strFilePath & """)"
    FileTS = FileDateTime(strFilePath)
    strLine = "GetDATEOfFile = FileTS"
    GetDATEOfFile = FileTS
    Exit Function
goterr:
    dateErrDescription = "Error " & Err.Number & " at GetDATEOfFile. Here:" & vbCrLf & _
     strLine & vbCrLf & vbCrLf & "Error description: " & Err.Description
    GetDATEOfFile = MyErrorDate
End Function


Public Function getBlackdINI(ByRef par1 As String, ByRef par2 As String, _
 ByRef par3 As String, ByRef par4 As String, ByRef par5 As Long, ByVal par6 As String, Optional forcePath As Boolean = False) As Long
  Dim tmpNum As Long 'D2
  If forcePath = True Then
    tmpNum = GetPrivateProfileString(par1, par2, par3, par4, par5, par6)
    getBlackdINI = tmpNum
    Exit Function
  End If
  If ((par1 = "MemoryAddresses") Or (par1 = "tileIDs") Or (par2 = "configPath")) Then
  If (Right$(par6, 10) = "config.ini") Then '2Dis this config.ini ?
  par6 = Left$(par6, (Len(par6) - 10)) & "config.override.ini" 'now its config.override.ini
  tmpNum = GetPrivateProfileString(par1, par2, par3, par4, par5, par6)
  par6 = Left$(par6, (Len(par6) - 19)) & "config.ini" 'fixing it to original config.ini again here. its ByRef btw!
  If (tmpNum > 0) Then 'if above 0, assume success! return, exit function
  getBlackdINI = tmpNum
  Exit Function
  End If
  'else, continue check the config.ini,
  End If
    getBlackdINI = GetPrivateProfileString(par1, par2, par3, par4, par5, par6)
  Else
  tmpNum = GetPrivateProfileString(par1, par2, par3, par4, par5, App.Path & "\settings.override.ini")
  If (tmpNum > 0) Then 'if above 0, assume success! return, exit function
  getBlackdINI = tmpNum
  Exit Function
  End If
    getBlackdINI = GetPrivateProfileString(par1, par2, par3, par4, par5, App.Path & "\settings.ini")
  End If
End Function

Public Function setBlackdINI(ByRef par1 As String, ByRef par2 As String, _
 ByRef par3 As String, ByRef par4 As String)
   If ((par1 = "MemoryAddresses") Or (par1 = "tileIDs") Or (par2 = "configPath")) Then
    setBlackdINI = WritePrivateProfileString(par1, par2, par3, par4)
  Else
    setBlackdINI = WritePrivateProfileString(par1, par2, par3, App.Path & "\settings.ini")
  End If
End Function

Public Sub SaveConfigWizard(ByVal blnValue As Boolean)
  On Error GoTo goterr
  Dim res As Boolean
  Dim strInfo As String
  Dim i As Long
  'Dim here As String
  res = True
  'here = App.path & "\" & configPath & "\config.ini"
  If blnValue = True Then
    strInfo = "1"
  Else
    strInfo = "0"
  End If
  i = setBlackdINI("Proxy", "ShowConfigWizard", strInfo, "")
  
  strInfo = ProxyVersion
  i = setBlackdINI("Proxy", "LastTimeDisplayedConfig", strInfo, "")
  
  Exit Sub
goterr:
  res = False
End Sub

Public Function myMainConfigINIPath() As String
    Dim res As String
    res = App.Path
    If ((Right$(res, 1) = "\") Or (Right$(res, 1) = "/")) Then
       myMainConfigINIPath = res & "config.ini"
    Else
       myMainConfigINIPath = res & "\config.ini"
    End If
End Function

Public Sub SaveFirstScreenConfig()
  On Error GoTo goterr
  Dim res As Boolean
  Dim strInfo As String
  Dim here As String
  Dim i As Long
  here = App.Path & "\" & OVERWRITE_CONFIGPATH & "\config.ini"
  res = True
  SaveConfigWizard OVERWRITE_SHOWAGAIN
  
  strInfo = OVERWRITE_CLIENT_PATH
  i = setBlackdINI("MemoryAddresses", "TibiaExePath", strInfo, here)
  
  strInfo = OVERWRITE_MAPS_PATH
  i = setBlackdINI("Proxy", "TibiaPath", strInfo, here)
  
  strInfo = OVERWRITE_CONFIGPATH
  i = WritePrivateProfileString("Proxy", "configPath", strInfo, myMainConfigINIPath)
  
  If (OVERWRITE_OT_MODE = True) Then
    strInfo = "3"
  Else
    strInfo = "1"
  End If
  i = setBlackdINI("Proxy", "ForwardOption", strInfo, here)
  
  strInfo = OVERWRITE_OT_IP
  i = setBlackdINI("Proxy", "ForwardGameTo", strInfo, here)
  
  strInfo = OVERWRITE_OT_PORT
  i = setBlackdINI("Proxy", "txtServerLoginP", strInfo, here)
  

  
Exit Sub
goterr:
  res = False
End Sub



Public Function CheckVariableCondition(par As String) As String
    On Error GoTo goterr
    Dim res As String
    Dim parts() As String
    Dim blnRes As Boolean
    Dim op As String
    Dim trop As String
    res = "0"
    parts = Split(par, ",")
    If UBound(parts) <> 2 Then
        GoTo goterr
    End If
    trop = Trim$(parts(1))
    op = Mid$(trop, 2, Len(trop) - 2)

    blnRes = ProcessRawCondition(Trim$(parts(0)), op, Trim$(parts(2)))
    If blnRes = True Then
        res = "0"
    Else
        res = "1"
    End If
    CheckVariableCondition = res
    Exit Function
goterr:
    CheckVariableCondition = "error"
End Function

Public Function CheckIsTrue(par As String) As String
    On Error GoTo goterr
    Dim res As String
    Dim parts() As String
    Dim lastPart As Long
    Dim i As Long
    res = "0"
    If par = "" Then
        GoTo goterr
    End If
    parts = Split(par, ",")
    lastPart = UBound(parts)
    For i = 0 To lastPart
        If Not (Trim$(parts(i)) = "0") Then
            res = "1"
            Exit For
        End If
    Next i
    CheckIsTrue = res
    Exit Function
goterr:
    CheckIsTrue = "error"
End Function

Public Function FixLastMSG(idConnection As Integer) As String
    Dim res As String
    On Error GoTo goterr
    
    res = var_lastmsg(idConnection)
    res = Replace(res, "{", "")
    res = Replace(res, "}", "")
    FixLastMSG = res
    Exit Function
goterr:
    FixLastMSG = "error"
End Function

Public Function GetTibiaTileSTR(ByVal b1 As Byte, ByVal b2 As Byte, ByVal b3 As Byte, ByVal b4 As Byte) As TibiaTileStr
    Dim res As TibiaTileStr
    Dim tileID As Long
    tileID = GetTheLong(b1, b2)
    If TibiaVersionLong >= 990 Then ' new tibia tiles
        res.str = GoodHex(b1) & " " & GoodHex(b2) & " FF"
        res.num = 3
        If DatTiles(tileID).haveExtraByte = True Then
            res.str = res.str & " " & GoodHex(b3)
            res.num = res.num + 1
        End If
        If DatTiles(tileID).haveExtraByte2 = True Then
            res.str = res.str & " " & GoodHex(b4)
            res.num = res.num + 1
        End If
    Else 'old tibia
        res.str = GoodHex(b1) & " " & GoodHex(b2)
        res.num = 2
        If DatTiles(tileID).haveExtraByte = True Then
            res.str = res.str & " " & GoodHex(b3)
            res.num = res.num + 1
        End If
        If DatTiles(tileID).haveExtraByte2 = True Then
            res.str = res.str & " " & GoodHex(b4)
            res.num = res.num + 1
        End If
    End If
End Function


Public Function UseItemOnName(idConnection As Integer, ByVal strTile As String) As Long 'HHBCODE 'use first found item with ID on first found creature with name....
'will use the item and return 0 (if item and name found), or -1.

    #If FinalMode = 1 Then
        On Error GoTo goterr
    #End If
    Dim ItemID As Long
    Dim skipBytes As Integer
    Dim ItemIDHexStr As String
    Dim LowByte As Byte
    Dim HighByte As Byte
    Dim Success As Boolean
    ItemIDHexStr = UCase(Left(strTile, 5))
    ItemID = GetTheLongFromFiveChr(ItemIDHexStr)
    LowByte = LowByteOfLong(ItemID)
    HighByte = HighByteOfLong(ItemID)
    strTile = Mid(strTile, 7) '7?? with the comma, i thought it'd be 6?? ...
    Dim name As String

    While Left(strTile, 1) = " "
    strTile = Mid(strTile, 1) 'ignore spaces..
    Wend


    name = strTile
    Success = SendAimbot(name, idConnection, LowByte, HighByte)
    If Success = True Then
    UseItemOnName = 0
    Else
    UseItemOnName = -1
    End If
    Exit Function
goterr:
    UseItemOnName = -1
End Function



Public Function IsIDE() As Boolean '
        gISIDE = False
        'This line is only executed if running in the IDE and then returns True
        Debug.Assert CheckIDE
        If gISIDE Then
          IsIDE = True
        Else
          IsIDE = False
        End If
End Function

Private Function CheckIDE() As Boolean ' this is a helper function for Public Function IsIDE()
        gISIDE = True 'set global flag
        CheckIDE = True
End Function


Function ExtractUrl(ByVal strUrl As String) As url
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    
    Dim retURL As url
    
    '1 look for a scheme it ends with ://
    intPos1 = InStr(strUrl, "://")
    
    If intPos1 > 0 Then
        retURL.Scheme = Mid(strUrl, 1, intPos1 - 1)
        strUrl = Mid(strUrl, intPos1 + 3)
    End If
        
    '2 look for a port
    intPos1 = InStr(strUrl, ":")
    intPos2 = InStr(strUrl, "/")
    
    If intPos1 > 0 And intPos1 < intPos2 Then
        ' a port is specified
        retURL.Host = Mid(strUrl, 1, intPos1 - 1)
        
        If (IsNumeric(Mid(strUrl, intPos1 + 1, intPos2 - intPos1 - 1))) Then
                retURL.port = CInt(Mid(strUrl, intPos1 + 1, intPos2 - intPos1 - 1))
        End If
    ElseIf intPos2 > 0 Then
        retURL.Host = Mid(strUrl, 1, intPos2 - 1)
    Else
        retURL.Host = strUrl
        retURL.uri = "/"
        
        ExtractUrl = retURL
        Exit Function
    End If
    
    strUrl = Mid(strUrl, intPos2)
    
    ' find a question mark ?
    intPos1 = InStr(strUrl, "?")
    
    If intPos1 > 0 Then
        retURL.uri = Mid(strUrl, 1, intPos1 - 1)
        retURL.Query = Mid(strUrl, intPos1 + 1)
    Else
        retURL.uri = strUrl
    End If
    
    ExtractUrl = retURL
End Function
Public Sub openErrorsTXTfolder()
On Error GoTo goterr
    If (shouldOpenErrorsTXTfolder = True) Then
    shouldOpenErrorsTXTfolder = False
    Shell "explorer " & GetMyAppDataFolder()
    End If
Exit Sub
goterr:

End Sub
Public Sub ResetConEventLogs()
   conEventLog = "Connection events:"
End Sub

Public Function FillIntWithZeroes(ByVal inte As Integer, ByVal digits As Long) As String
    FillIntWithZeroes = Right("0000" & CStr(inte), digits)
End Function
Public Function GetBetterTimestamp() As String
    Dim sAns As String
    Dim typTime As SYSTEMTIME
    On Error Resume Next
    GetSystemTime typTime
    sAns = "[" & _
    FillIntWithZeroes(typTime.wDay, 2) & "/" & _
    FillIntWithZeroes(typTime.wMonth, 2) & "/" & _
    FillIntWithZeroes(typTime.wYear, 4) & " " & _
    FillIntWithZeroes(typTime.wHour, 2) & ":" & _
    FillIntWithZeroes(typTime.wMinute, 2) & ":" & _
    FillIntWithZeroes(typTime.wSecond, 2) & "." & _
    FillIntWithZeroes(typTime.wMilliseconds, 3) & "]"
    GetBetterTimestamp = sAns
End Function
Public Sub LogConEvent(ByRef strDebug As String)
  Dim ts As String
  ts = GetBetterTimestamp()
  Debug.Print ts & " " & strDebug
  conEventLog = conEventLog & vbCrLf & ts & " " & strDebug
End Sub
