Attribute VB_Name = "modMap"
#Const FinalMode = 1
#Const MapDebug = 0
Option Explicit


Public Const cte_automapfolder = "c:\Documents and Settings\USER\Application Data\Tibia\Automap"



' The chaos in the variable order is forced
' to cause chaos in the final memory,
' so it is harder to crack anything in the final .exe
' Note that important variables are around all the code, not in the same zone.

Public Const LimitedLeader = "-"
'Public Const LimitedLeader = "Someone" ' Used to compile with restricted leader of sync SD attack

Private Const OptCte1 = 31416 'size of 1 map floor
Private Const OptCte2 = 1716 'size of 1 map line
Private Const OptCte3 = 132 'size of 1 map square

'Public Const defaultSelectedTibiaFolder As String = "TibiaPreview"
Public Const defaultSelectedTibiaFolder As String = "Tibia"
'Public Const highestTibiaVersionLong As Long = 1033 ' highest known Tibia version (long)
'Public Const TibiaVersionDefaultString As String = "10.33" ' highest known Tibia version (string)
'Public Const TibiaVersionForceString As String = "10.33" ' set this version by default (string)

Public highestTibiaVersionLong As Long   ' highest known Tibia version (long)
Public TibiaVersionDefaultString As String    ' highest known Tibia version (string)
Public TibiaVersionForceString As String   ' set this version by default (string)

Public addConfigPaths As String ' list of new config paths here
Public addConfigVersions As String ' relative versions
Public addConfigVersionsLongs As String 'relative version longs

Public Const ProxyVersion = "36.7" ' Proxy version ' string version
Public Const myNumericVersion = 36700 ' numeric version
Public Const myAuthProtocol = 2 ' authetication protocol
Public Const TrialVersion = False ' true=trial version

' HERE YOU CAN LIMIT BLACKDPROXY TO ONLY RUN IN ONE COMPUTER
Public Const DoMAC_check = False ' change this to TRUE to limit it
' Enter the MAC of the allowed computer here
' Example for MAC address 00:80:C8:2C:44:89
Public Const Trial_MAC1 = &H0 ' 1st byte of MAC address
Public Const Trial_MAC2 = &H80 ' 2nd byte of MAC address
Public Const Trial_MAC3 = &HC8 ' 3rd byte of MAC address
Public Const Trial_MAC4 = &H2C ' 4th byte of MAC address
Public Const Trial_MAC5 = &H44 ' 5th byte of MAC address
Public Const Trial_MAC6 = &H89 ' 6th byte of MAC address





' some colors for map
Public Const ColourField = &H80&
Public Const ColourPath = &HFFFFC0
Public Const ColourNothing = &H111111


' some more colors for map
Public Const ColourGround = &HC0FFC0
Public Const ColourPlayer = &H80C0FF



' some more colors for map
Public Const ColourUnknown = &H8000000F
Public Const ColourWithInfo = &HC0FFC0
Public Const ColourWithMe = &H8080FF
Public Const ColourSelected = &HC0FFFF
Public Const ColourSomething = &H808080
Public Const ColourSomething2 = &H80&
Public Const ColourSomething3 = &HFFC0FF
Public Const ColourSomething4 = &HFF80FF



' some more colors for map
Public Const ColourSomething5 = &HFF00FF
Public Const ColourSomething6 = &HC000C0
Public Const ColourSomething7 = &HFFC0C0
Public Const ColourDown = &H800080
Public Const ColourUp = &HFF00FF
Public Const ColourBlockMoveable = &HC000&
Public Const ColourWater = &HFF8080
Public Const ColourFish = &HFFFF00


' some types
Public Type TypeSpecialRes
  str As String
  bln As Boolean
  bestX As Long
  bestY As Long
  bestMelee As Boolean
  bestHMM As Boolean
End Type
Public Type TypeMatrixPosition ' 1 stack position of any square of map , might not be valid
  valid As Boolean ' is valid?
  X As Long
  y As Long
  z As Long
  s As Long ' stack
End Type
Public Type TypeTileInfo ' 1 stack positon of a square of map inside limits
  t1 As Byte ' t1
  t2 As Byte ' and t2 builds the tile ID
  t3 As Byte ' t3 is the amount if avaiable, else store 0 / (store FF instead 0 since tibia 9.9)
  t4 As Byte ' new byte since Tibia 9.9 , but struct TypeTileInfo will still have same size
  dblID As Double ' id of creature if avaiable, else store 0
End Type
Public Type TypePlayerInfo ' info of creature
  pos As Long
  stackpos As Long
  newID As Double
End Type
Public Type TypeLearnResult ' result of learning about a packet
  fail As Boolean ' fail in protocol
  pos As Long ' position after processing packet
  skipThis As Boolean ' recommend skiping this
  firstMapDone As Boolean ' first map have been read
  gotHPupdate As Boolean ' got info about Heal update
  gotManaupdate As Boolean ' got info about mana update
  gotSoulupdate As Boolean ' got info about soulpoints update
  'gotBlankRune As Boolean ' got info about new blank rune in hand
  gotNewCorpse As Boolean ' got info about new corpse in screen
End Type
Public Type TypeSearchItemResult2 ' result of searching item in backpacks
  foundcount As Long ' total items found matching the search
  bpID As Byte ' bestChoose: ID of container
  slotID As Byte ' bestCHoose: slot inside that container
  b1 As Byte ' b1
  b2 As Byte ' and b2 hold the tileID of bestChoose if we are searching class (ex:food)
  amount As Byte ' amount if item bestChoose is stackable
  b4 As Byte ' new byte since tibia 9.9
End Type
Public Type TypeStackTileInfo '1 entire square of map info
  s(0 To 10) As TypeTileInfo
End Type
' API for fast moves of memory in blocks
Private Declare Sub RtlMoveMemory Lib "Kernel32" ( _
    lpDest As Any, _
    lpSource As Any, _
    ByVal ByValcbCopy As Long)
    
' some vars we are going to use mainly in this module:

Public forcedDebugChain As Boolean

Public gmStart As String
Public gmStart2 As String

Public oldmessage_H0 As Byte
Public oldmessage_H1 As Byte
Public oldmessage_H2 As Byte
Public oldmessage_H3 As Byte
Public oldmessage_H4 As Byte
Public oldmessage_H5 As Byte
Public oldmessage_H6 As Byte
Public oldmessage_H7 As Byte
Public oldmessage_H8 As Byte
Public oldmessage_H9 As Byte
Public oldmessage_HA As Byte
Public oldmessage_HB As Byte
Public oldmessage_HC As Byte
Public oldmessage_HD As Byte
Public oldmessage_HE As Byte
Public oldmessage_HF As Byte
Public oldmessage_H10 As Byte
Public oldmessage_H11 As Byte
Public oldmessage_H12 As Byte
Public oldmessage_H13 As Byte
Public oldmessage_H14 As Byte
Public oldmessage_H15 As Byte
Public newmessage_H8 As Byte

Public newchatmessage_HA As Byte
Public newchatmessage_H9 As Byte
Public verynewchatmessage_HB As Byte

Public lastIngameCheck() As String
Public lastIngameCheckTileID() As String
Public lastPossibleOrder As Integer
Public UsedSpamOrders As Integer
Public IDstring() As String
Public myID() As Double
Public myX() As Long
Public myY() As Long
Public myZ() As Long
' for autolooter:
Public myLastCorpseX() As Long
Public myLastCorpseY() As Long
Public myLastCorpseZ() As Long
Public myLastCorpseS() As Long
Public myLastCorpseTileID() As Long
Public lootWaiting() As Boolean
Public receivedLogin() As Boolean
' for cavebot anti trap system:
Public lastFloorChangeX() As Long
Public lastFloorChangeY() As Long
Public lastFloorChangeZ() As Long
' important stats:
Public myNewStat() As Long
Public myHP() As Long
Public myMaxHP() As Long
Public myMaxMana() As Long
Public myLevel() As Long
Public myExp() As Long
Public myInitialExp() As Long
Public myInitialTickCount() As Long
Public myMagLevel() As Long
Public myMana() As Long
Public myCap() As Long
Public myStamina() As Long
Public mySoulpoints() As Long
Public mySlot() As TypeItem ' all slots
Public savedItem() As TypeItem ' saved item before runemaking
Public CharacterName() As String
' for trial system:
Public sentFirstPacket() As String
Public sentWelcome() As Boolean
Public HighestConnectionID As Integer
' some other vars:
Public LogoutReason() As String
Public mapIDselected As Integer
Public mapFloorSelected As Long
Public cavebotIDselected As Integer

Public Matrix() As TypeStackTileInfo ' THE MAP MATRIX , it will have 5 dimensions!

Public NameOfID() As scripting.Dictionary  ' A dictionary ID (double) -> name (string)

Public HPOfID() As scripting.Dictionary ' A dictionary ID (double) -> HP% (byte)

Public DirectionOfID() As scripting.Dictionary ' A dictionary ID (double) -> direction (byte)




Public GotPacketWarning() As Boolean ' for safe mode

' some more vars:
Public doingTrade() As Boolean
Public doingTrade2() As Boolean
Public lastPing() As Long
Public SendToServer() As Byte
Public AfterLoginLogoutReason() As String
Public pushTarget() As Double
Public pushDelay() As Integer
Public DangerGM() As Boolean
Public DangerPK() As Boolean
Public DangerPlayer() As Boolean
Public DangerGMname() As String
Public DangerPKname() As String
Public DangerPlayerName() As String
Public LogoutTimeGM() As Long
Public GMname() As String
Public LoginMsgCount() As Long

Public lastHPchange() As Long

Public StatusBits() As String



Public SpamAutoHeal() As Boolean ' client require autoheal
Public SpamAutoMana() As Boolean
Public SpamAutoPush() As Boolean ' client require autopushing someone

Public GotTrialLock As Boolean
Public blank1 As Byte
Public blank2 As Byte
Public lastLockReason As String


Public tmpStack As TypeStackTileInfo ' for map update speed optimization

' sound vars:
Public PlayTheDangerSound As Boolean
Public PlayMsgSound As Boolean
Public PlayMsgSound2 As Boolean

' some more trial vars:
Public TrialMode As Integer  '1=some days / 2 = month
Public TrialLimit_Day As Long ' days from 1 January 2005 that marks the limit

' for tests - when you see desync
Public givenUFO As Boolean ' used to don't give UFO message more than once in same session

' for anti crack protection:
Public trialSafety2 As Integer
Public trialSafety1 As Integer


Public trialSafety4 As Integer
Public trialSafety300 As Integer
Public UHRetryCount() As Long

Public runemakerMana1() As Long




Public Function ByteToBitstring(b As Byte) As String
 ' convert a byte into a string of 1's and 0's
 ' by Blackd . www.blackdtools.com
 Dim testbit As Integer
 Dim testb As Integer
 Dim i As Integer
 Dim res As String
 Dim modr As Integer
 testb = CInt(b)
 res = ""
 For i = 1 To 8
   testbit = 2 ^ i
   modr = testb Mod testbit
   If modr = 0 Then
     res = "0" & res
   Else
     res = "1" & res
   End If
   testb = testb - modr
 Next i
 ByteToBitstring = res
End Function

Public Function IsGM(str As String) As Boolean
  Dim gmpart As String
  Dim lCaseStr As String
  Dim specialGuido As Boolean
  If str <> "" Then
    lCaseStr = LCase(str)
    gmpart = Left(lCaseStr, 3)
    specialGuido = False
    If isSpecialGMname(str) = True Then
      specialGuido = True
    End If
    If (gmpart = gmStart) Or (gmpart = gmStart2) Or (specialGuido = True) Then
      IsGM = True
    Else
      IsGM = False
    End If
  Else
    IsGM = False
  End If
End Function

Public Sub CheckIfGM(idConnection As Integer, ByRef str As String, zpos As Long, Optional forceGm As Boolean = False)
  ' check if a name is gm
  Dim gmpart As String
  Dim aRes As Long
  Dim i As Integer
  Dim secL As Long
  Dim gotGM As Boolean
  Dim condEnter As Boolean
  Dim specialGuido As Boolean
  Dim lCaseStr As String
  lCaseStr = LCase(str)
  gotGM = False
  condEnter = False
  If (cavebotEnabled(idConnection) = True) Or (RuneMakerOptions(idConnection).activated = True) Or (RuneMakerOptions(idConnection).autoEat = True) Then
    If DangerGM(idConnection) = False Then
        If forceGm = False Then
            gotGM = IsGM(lCaseStr)
        Else
            gotGM = True
        End If
        If (gotGM = True) Then
            DangerGM(idConnection) = True
            secL = CLng(Int((180 * Rnd) + 60))
            GMname(idConnection) = str
            DangerGMname(idConnection) = str
            If CheatsPaused(idConnection) = True Then
                logoutAllowed(idConnection) = 1200000 + GetTickCount() ' allowed for 20 min
                aRes = SendLogSystemMessageToClient(idConnection, "GM detected ( " & GMname(idConnection) & " ). Automatic cheats are already paused")
            Else
                logoutAllowed(idConnection) = 1200000 + GetTickCount() ' allowed for 20 min
                LogoutTimeGM(idConnection) = GetTickCount() + (secL * 1000)
                aRes = SendLogSystemMessageToClient(idConnection, "GM detected ( " & GMname(idConnection) & " ). Some cheats are now disabled. Closing in " & secL & " seconds. Cancel with Exiva cancel")
            End If
            DoEvents
            If frmRunemaker.ChkDangerSound.Value = 1 Then
                ChangePlayTheDangerSound True
            End If
        End If
    End If
    If (gotGM = False) Then
       If zpos = myZ(idConnection) Then
         If (cavebotOnPLAYERpause(idConnection) = True) Or (RuneMakerOptions(idConnection).msgSound2 = True) Then
           If (isMelee(idConnection, str) = False) And (isHmm(idConnection, str) = False) And (frmRunemaker.IsFriend(LCase(str)) = False) And (CharacterName(idConnection) <> str) And (str <> "") Then
             DangerPlayerName(idConnection) = str
             DangerPlayer(idConnection) = True
            End If
         End If
       End If
    End If
  End If
  
End Sub
Public Sub ResetSpamOrders()
  ' reset all orders that require spam
  Dim i As Integer
  For i = 1 To MAXCLIENTS
    SpamAutoHeal(i) = False
    SpamAutoPush(i) = False
    SpamAutoFastHeal(i) = False
    SpamAutoMana(i) = False
  Next i
  UsedSpamOrders = 0
End Sub
Public Function GetSpamOrderPosition(idConnection As Integer, order As Integer) As Integer
  ' return 1 if order exist
  ' return 0 if order doesnt exist
  Dim res As Integer
  Select Case order
  Case 1
    If SpamAutoHeal(idConnection) = True Then
      res = 1
    Else
      res = 0
    End If
  Case 2
    If SpamAutoPush(idConnection) = True Then
      res = 1
    Else
      res = 0
    End If
  Case 3
    If SpamAutoFastHeal(idConnection) = True Then
      res = 1
    Else
      res = 0
    End If
 Case 4
    If SpamAutoMana(idConnection) = True Then
      res = 1
    Else
      res = 0
    End If
 Case Else
    res = 0
  End Select
  GetSpamOrderPosition = res
End Function
Public Sub AddSpamOrder(idConnection As Integer, order As Integer)
  ' add spam order for certain idConnection
  If GetSpamOrderPosition(idConnection, order) = 0 Then
    UsedSpamOrders = UsedSpamOrders + 1
    Select Case order
    Case 1
      SpamAutoHeal(idConnection) = True
    Case 2
      SpamAutoPush(idConnection) = True
    Case 3
      SpamAutoFastHeal(idConnection) = True
      nextFastHeal(idConnection) = GetTickCount()
    Case 4
      SpamAutoMana(idConnection) = True

    End Select
  End If
  If frmHardcoreCheats.timerSpam.enabled = False Then
    frmHardcoreCheats.timerSpam.enabled = True ' enable spam timer
  End If
End Sub
Public Sub RemoveSpamOrder(idConnection As Integer, order As Integer)
  ' remove spam order for certain idConnection
  Dim pos As Integer
  Dim i As Integer
  pos = GetSpamOrderPosition(idConnection, order)
  If pos > 0 Then
    UsedSpamOrders = UsedSpamOrders - 1
    Select Case order
    Case 1
      SpamAutoHeal(idConnection) = False
    Case 2
      SpamAutoPush(idConnection) = False
    Case 3
      SpamAutoFastHeal(idConnection) = False
    Case 4
      SpamAutoMana(idConnection) = False
    End Select
  End If
  If UsedSpamOrders = 0 Then
    frmHardcoreCheats.timerSpam.enabled = False 'disable timer to save CPU
  End If
End Sub
Public Sub RemoveAllClientSpamOrders(idConnection As Integer)
  ' remove all spam orders of that client
  RemoveSpamOrder idConnection, 1
  RemoveSpamOrder idConnection, 2
  RemoveSpamOrder idConnection, 3
  RemoveSpamOrder idConnection, 4
  RemoveSpamOrder idConnection, 5
  RemoveSpamOrder idConnection, 6
End Sub

Public Function SearchItem(idConnection As Integer, t1 As Byte, t2 As Byte) As TypeSearchItemResult2
  ' search item in any open container
  ' return last found
  ' and count
  Dim res As TypeSearchItemResult2
  Dim i As Long
  Dim j As Long
  Dim limitJ As Long
  res.foundcount = 0
  res.bpID = &HFF
  res.slotID = &HFF
  res.b1 = t1
  res.b2 = t2
  res.amount = 0
  res.b4 = 0
  For i = 0 To HIGHEST_BP_ID
    If (Backpack(idConnection, i).open = True) Then
    limitJ = (Backpack(idConnection, i).used) - 1
    For j = 0 To limitJ
      If Backpack(idConnection, i).item(j).t1 = t1 And _
       Backpack(idConnection, i).item(j).t2 = t2 Then
        res.foundcount = res.foundcount + 1
        res.bpID = CByte(i)
        res.slotID = CByte(j)
        res.amount = Backpack(idConnection, i).item(j).t3
        res.b4 = Backpack(idConnection, i).item(j).t4
      End If
    Next j
    End If
  Next i
  SearchItem = res
End Function


Public Function SearchItemGoodLoot(idConnection As Integer) As TypeSearchItemResult2
  ' search item in any open container
  ' return last found
  ' and count
  Dim res As TypeSearchItemResult2
  Dim i As Long
  Dim j As Long
  Dim limitJ As Long
  Dim tileID As Long
  res.foundcount = 0
  res.bpID = &HFF
  res.slotID = &HFF
  res.b1 = &HFF
  res.b2 = &HFF
  res.amount = 0
  res.b4 = 0
  For i = 0 To HIGHEST_BP_ID
    If (Backpack(idConnection, i).open = True) Then
    limitJ = (Backpack(idConnection, i).used) - 1
    For j = 0 To limitJ
      tileID = GetTheLong(Backpack(idConnection, i).item(j).t1, Backpack(idConnection, i).item(j).t2)
      If (IsGoodLoot(idConnection, tileID) = True) Then
        res.foundcount = res.foundcount + 1
        res.bpID = CByte(i)
        res.slotID = CByte(j)
        res.b1 = Backpack(idConnection, i).item(j).t1
        res.b2 = Backpack(idConnection, i).item(j).t2
        res.amount = Backpack(idConnection, i).item(j).t3
        res.b4 = Backpack(idConnection, i).item(j).t4
      End If
    Next j
    End If
  Next i
  SearchItemGoodLoot = res
End Function

Public Function SearchFirstItem(idConnection As Integer, t1 As Byte, t2 As Byte) As TypeSearchItemResult2
  ' search item in any open container
  ' return first found
  Dim res As TypeSearchItemResult2
  Dim i As Long
  Dim j As Long
  Dim limitJ As Long
  res.foundcount = 0
  res.bpID = &HFF
  res.slotID = &HFF
  res.b1 = t1
  res.b2 = t2
  res.amount = 0
  res.b4 = 0
  For i = 0 To HIGHEST_BP_ID
    If (Backpack(idConnection, i).open = True) Then
    limitJ = (Backpack(idConnection, i).used) - 1
    For j = 0 To limitJ
      If Backpack(idConnection, i).item(j).t1 = t1 And _
       Backpack(idConnection, i).item(j).t2 = t2 Then
        res.foundcount = 1
        res.bpID = CByte(i)
        res.slotID = CByte(j)
        res.amount = Backpack(idConnection, i).item(j).t3
        SearchFirstItem = res
        res.b4 = Backpack(idConnection, i).item(j).t4
        Exit Function
      End If
    Next j
    End If
  Next i
  SearchFirstItem = res
End Function

Public Function SearchSubContainer(idConnection As Integer, t1 As Byte, t2 As Byte, containerBPid As Byte) As TypeSearchItemResult2
  ' search subcontainer in any open container
  ' return last found
  ' and count
  Dim res As TypeSearchItemResult2
  Dim i As Long
  Dim j As Long
  Dim limitJ As Long
  Dim tileID As Long
  res.foundcount = 0
  res.bpID = &HFF
  res.slotID = &HFF
  res.b1 = t1
  res.b2 = t2
  res.amount = 0
  res.b4 = 0
  i = CLng(containerBPid)
    limitJ = (Backpack(idConnection, i).used) - 1
    For j = 0 To limitJ
      tileID = GetTheLong(Backpack(idConnection, i).item(j).t1, Backpack(idConnection, i).item(j).t2)
      If DatTiles(tileID).iscontainer = True Then
        res.foundcount = res.foundcount + 1
        res.bpID = CByte(i)
        res.slotID = CByte(j)
        res.b1 = Backpack(idConnection, i).item(j).t1
        res.b2 = Backpack(idConnection, i).item(j).t2
        res.amount = Backpack(idConnection, i).item(j).t3
        res.b4 = Backpack(idConnection, i).item(j).t4
        SearchSubContainer = res
        Exit Function
      End If
    Next j
  SearchSubContainer = res
End Function
Public Function SearchItemWithBPException(idConnection As Integer, t1 As Byte, t2 As Byte, noValidBP As Byte) As TypeSearchItemResult2
  ' search item in any open container, except given container ID
  ' return last found
  ' and count
  Dim res As TypeSearchItemResult2
  Dim i As Long
  Dim j As Long
  Dim limitJ As Long
  res.foundcount = 0
  res.bpID = &HFF
  res.slotID = &HFF
  res.amount = 0
  res.b4 = 0
  For i = 0 To HIGHEST_BP_ID
    If (Backpack(idConnection, i).open = True) Then
    If i <> noValidBP Then
    limitJ = (Backpack(idConnection, i).used) - 1
    For j = 0 To limitJ
      If Backpack(idConnection, i).item(j).t1 = t1 And _
       Backpack(idConnection, i).item(j).t2 = t2 Then
        res.foundcount = res.foundcount + 1
        res.bpID = CByte(i)
        res.slotID = CByte(j)
        res.amount = Backpack(idConnection, i).item(j).t3
        res.b4 = Backpack(idConnection, i).item(j).t4
      End If
    Next j
    End If
    End If
  Next i
  res.b1 = t1
  res.b2 = t2
  SearchItemWithBPException = res
End Function

Public Function SearchItemWithBPExceptionGoodLoot(idConnection As Integer, noValidBP As Byte) As TypeSearchItemResult2
  ' search item in any open container, except given container ID
  ' Only good loot items are valid
  Dim res As TypeSearchItemResult2
  Dim i As Long
  Dim j As Long
  Dim limitJ As Long
  Dim tileID As Long
  res.foundcount = 0
  res.bpID = &HFF
  res.slotID = &HFF
  res.amount = 0
  res.b4 = 0
  For i = 0 To HIGHEST_BP_ID
    If (Backpack(idConnection, i).open = True) Then
    If i <> noValidBP Then
    limitJ = (Backpack(idConnection, i).used) - 1
    For j = 0 To limitJ
      tileID = GetTheLong(Backpack(idConnection, i).item(j).t1, Backpack(idConnection, i).item(j).t2)
      If IsGoodLoot(idConnection, tileID) = True Then
        res.b1 = Backpack(idConnection, i).item(j).t1
        res.b2 = Backpack(idConnection, i).item(j).t2
        res.foundcount = res.foundcount + 1
        res.bpID = CByte(i)
        res.slotID = CByte(j)
        res.amount = Backpack(idConnection, i).item(j).t3
        res.b4 = Backpack(idConnection, i).item(j).t4
      End If
    Next j
    End If
    End If
  Next i
  SearchItemWithBPExceptionGoodLoot = res
End Function

Public Function SearchAmmount(idConnection As Integer, t1 As Byte, t2 As Byte) As Long
  ' search item in any open container
  ' return last found
  ' and count
  ' include ammount
  
  Dim i As Long
  Dim j As Long
  Dim limitJ As Long
  Dim rCount As Long
  rCount = 0
  For i = 1 To EQUIPMENT_SLOTS
    If mySlot(idConnection, i).t1 = t1 And mySlot(idConnection, i).t2 = t2 Then
      If mySlot(idConnection, i).t3 = 0 Then
        rCount = rCount + 1
      Else
        rCount = rCount + CLng(mySlot(idConnection, i).t3)
      End If
    End If
  Next i
  For i = 0 To HIGHEST_BP_ID
    If (Backpack(idConnection, i).open = True) Then
    limitJ = (Backpack(idConnection, i).used) - 1
    For j = 0 To limitJ
      If Backpack(idConnection, i).item(j).t1 = t1 And _
       Backpack(idConnection, i).item(j).t2 = t2 Then
        If Backpack(idConnection, i).item(j).t3 = 0 Then
          rCount = rCount + 1
        Else
          rCount = rCount + CLng(Backpack(idConnection, i).item(j).t3)
        End If
      End If
    Next j
    End If
  Next i
  SearchAmmount = rCount
End Function


Public Function SearchExactAmmount(idConnection As Integer, t1 As Byte, t2 As Byte, am As Byte) As Long
  ' search item in any open container
  ' return last found
  ' and count
  ' include ammount

  Dim i As Long
  Dim j As Long
  Dim limitJ As Long
  Dim rCount As Long
  Dim amSlot As Byte
  rCount = 0
  For i = 1 To EQUIPMENT_SLOTS
        amSlot = mySlot(idConnection, i).t3
        If ((mySlot(idConnection, i).t1 = t1) And (mySlot(idConnection, i).t2 = t2) And (amSlot = am)) Then
             rCount = rCount + 1
        End If
  Next i
  For i = 0 To HIGHEST_BP_ID
    If (Backpack(idConnection, i).open = True) Then
    limitJ = (Backpack(idConnection, i).used) - 1
    For j = 0 To limitJ
      If Backpack(idConnection, i).item(j).t1 = t1 And _
       Backpack(idConnection, i).item(j).t2 = t2 Then
        If Backpack(idConnection, i).item(j).t3 = am Then
          rCount = rCount + 1
        End If
      End If
    Next j
    End If
  Next i
  SearchExactAmmount = rCount
End Function

Public Function SearchItemWithAmount(idConnection As Integer, t1 As Byte, t2 As Byte, am As Byte) As TypeSearchItemResult2
  ' search item in any open container , should have at least certain ammount
  ' return last found
  ' and count
  ' include ammount
  Dim res As TypeSearchItemResult2
  Dim i As Long
  Dim j As Long
  Dim limitJ As Long
  res.foundcount = 0
  res.bpID = &HFF
  res.slotID = &HFF
  res.amount = 0
  res.b4 = 0
  For i = 0 To HIGHEST_BP_ID
    If (Backpack(idConnection, i).open = True) Then
    limitJ = (Backpack(idConnection, i).used) - 1
    For j = 0 To limitJ
      If Backpack(idConnection, i).item(j).t1 = t1 And _
       Backpack(idConnection, i).item(j).t2 = t2 And _
        Backpack(idConnection, i).item(j).t3 >= am Then
        res.foundcount = res.foundcount + 1
        res.bpID = CByte(i)
        res.slotID = CByte(j)
        res.amount = Backpack(idConnection, i).item(j).t3
        res.b4 = Backpack(idConnection, i).item(j).t4
      End If
    Next j
    End If
  Next i
  res.b1 = t1
  res.b2 = t2
  SearchItemWithAmount = res
End Function
' ..
'SearchFirstItemWithExactAmmount
Public Function SearchFirstItemWithExactAmmount(idConnection As Integer, t1 As Byte, t2 As Byte, am As Byte) As TypeSearchItemResult2
  ' search item in any open container , should have at least certain ammount
  ' return last found
  ' and count
  ' include ammount
  Dim res As TypeSearchItemResult2
  Dim i As Long
  Dim j As Long
  Dim limitJ As Long
  res.foundcount = 0
  res.bpID = &HFF
  res.slotID = &HFF
  res.amount = 0
  res.b4 = 0
  For i = 0 To HIGHEST_BP_ID
    If (Backpack(idConnection, i).open = True) Then
    limitJ = (Backpack(idConnection, i).used) - 1
    For j = 0 To limitJ
      If Backpack(idConnection, i).item(j).t1 = t1 And _
       Backpack(idConnection, i).item(j).t2 = t2 And _
       Backpack(idConnection, i).item(j).t3 = am Then
        res.foundcount = res.foundcount + 1
        res.b1 = t1
        res.b2 = t2
        res.bpID = CByte(i)
        res.slotID = CByte(j)
        res.amount = am
        res.b4 = Backpack(idConnection, i).item(j).t4
        SearchFirstItemWithExactAmmount = res ' return first
        Exit Function
      End If
    Next j
    End If
  Next i
  res.b1 = t1
  res.b2 = t2
  res.amount = 7
  SearchFirstItemWithExactAmmount = res
End Function

Public Function ValidAsBP(strName As String) As Boolean
    On Error GoTo notvali
    Dim lcName As String
    Dim fPart As String
    lcName = LCase(strName)
    fPart = Right$(lcName, 8)
    If fPart = "backpack" Then
        ValidAsBP = True
    Else
        ValidAsBP = False
    End If
    Exit Function
notvali:
    ValidAsBP = False
End Function

Public Function SearchFreeSlot(idConnection As Integer) As TypeSearchItemResult2
  ' search free slot in any open container
  ' return first found
  Dim res As TypeSearchItemResult2
  Dim i As Long
  Dim j As Long
  Dim limitJ As Long
  Dim tmpb1 As Byte
  Dim tmpb2 As Byte
  res.foundcount = 0
  res.bpID = &HFF
  res.slotID = &HFF
  res.amount = 0
  res.b4 = 0
  For i = 0 To HIGHEST_BP_ID
    If (Backpack(idConnection, i).open = True) Then
    If ValidAsBP(Backpack(idConnection, i).name) = True Then
    limitJ = (Backpack(idConnection, i).cap) - 1
    For j = 0 To limitJ
    tmpb1 = Backpack(idConnection, i).item(j).t1
    tmpb2 = Backpack(idConnection, i).item(j).t2
      If ((tmpb1 = 0) And (tmpb2 = 0)) Then
        res.foundcount = res.foundcount + 1
        res.bpID = CByte(i)
        res.slotID = CByte(j)
        res.amount = Backpack(idConnection, i).item(j).t3
        res.b4 = Backpack(idConnection, i).item(j).t4
        SearchFreeSlot = res
        Exit Function
      End If
    Next j
    End If
    End If
  Next i
  res.b1 = 0
  res.b2 = 0
  SearchFreeSlot = res
End Function
Public Function SearchFreeSlotInContainer(idConnection As Integer, i As Byte) As TypeSearchItemResult2
  ' search free slot in a given container
  ' return first found
  Dim res As TypeSearchItemResult2
  Dim j As Long
  Dim limitJ As Long
  res.foundcount = 0
  res.bpID = &HFF
  res.slotID = &HFF
  res.amount = 0
  res.b4 = 0
    If ValidAsBP(Backpack(idConnection, i).name) Then
    limitJ = (Backpack(idConnection, i).cap) - 1
    For j = 0 To limitJ
      If Backpack(idConnection, i).item(j).t1 = 0 And _
       Backpack(idConnection, i).item(j).t2 = 0 Then
        res.foundcount = res.foundcount + 1
        res.bpID = CByte(i)
        res.slotID = CByte(j)
        res.amount = Backpack(idConnection, i).item(j).t3
        res.b4 = Backpack(idConnection, i).item(j).t4
        SearchFreeSlotInContainer = res
        Exit Function
      End If
    Next j
    End If
  res.b1 = 0
  res.b2 = 0
  SearchFreeSlotInContainer = res
End Function

Public Function SearchItemInBP(idConnection As Integer, t1 As Byte, t2 As Byte, bpID As Byte) As TypeSearchItemResult2
  ' search item in given container
  ' return last found
  ' and count
  Dim res As TypeSearchItemResult2
  Dim i As Long
  Dim j As Long
  Dim limitJ As Long
  res.foundcount = 0
  res.bpID = &HFF
  res.slotID = &HFF
  res.amount = 0
  res.b4 = 0
  limitJ = (Backpack(idConnection, i).used) - 1
  For j = 0 To limitJ
    If Backpack(idConnection, bpID).item(j).t1 = t1 And _
     Backpack(idConnection, bpID).item(j).t2 = t2 Then
      res.foundcount = res.foundcount + 1
      res.bpID = CByte(i)
      res.slotID = CByte(j)
      res.amount = Backpack(idConnection, bpID).item(j).t3
      res.b4 = Backpack(idConnection, bpID).item(j).t4
    End If
  Next j
  res.b1 = t1
  res.b2 = t2
  SearchItemInBP = res
End Function

Public Function SearchItemDestination(idConnection As Integer, t1 As Byte, t2 As Byte, novalidBpID As Byte) As TypeSearchItemResult2
  ' search destination in bps except novalidBpID
  ' (for loot operation)
  SearchItemDestination = SearchItemDestinationForLoot(idConnection, t1, t2, novalidBpID)
End Function

Public Function SearchItemDestinationForLoot(idConnection As Integer, t1 As Byte, t2 As Byte, novalidBpID As Byte) As TypeSearchItemResult2
  ' search destination in bps except novalidBpID
  ' (for loot operation)
  Dim res As TypeSearchItemResult2
  Dim i As Long
  Dim j As Long
  Dim limitJ As Long
  Dim tileID As Long
  Dim isStackable As Boolean
  tileID = GetTheLong(t1, t2)
  isStackable = DatTiles(tileID).stackable
  res.foundcount = 0
  res.bpID = &HFF
  res.slotID = &HFF
  res.amount = 0
  res.b4 = 0
  For i = 0 To HIGHEST_BP_ID
    If (Backpack(idConnection, i).open = True) Then
    If i <> CLng(novalidBpID) Then
      If (isStackable = True) Then
      limitJ = (Backpack(idConnection, i).used) - 1
      For j = 0 To limitJ
        If Backpack(idConnection, i).item(j).t1 = t1 And _
         Backpack(idConnection, i).item(j).t2 = t2 And _
         Backpack(idConnection, i).item(j).t3 < &H64 Then
          res.foundcount = res.foundcount + 1
          res.bpID = CByte(i)
          res.slotID = CByte(j)
          res.amount = Backpack(idConnection, i).item(j).t3
          res.b4 = Backpack(idConnection, i).item(j).t4
          SearchItemDestinationForLoot = res
          Exit Function
        End If
      Next j
      End If
      If Backpack(idConnection, i).used < Backpack(idConnection, i).cap Then
          res.foundcount = res.foundcount + 1
          res.bpID = CByte(i)
          res.slotID = CByte(Backpack(idConnection, i).used)
          res.amount = 0
          res.b4 = 0
          SearchItemDestinationForLoot = res
          Exit Function
      End If
    End If
    End If
  Next i
  res.b1 = t1
  res.b2 = t2
  SearchItemDestinationForLoot = res
End Function
Public Function SearchItemDestinationInDepot(idConnection As Integer, t1 As Byte, t2 As Byte, depotBpID As Byte) As TypeSearchItemResult2
  ' search destination in bp depotBpID
  ' (for Depot deploy operation)
  Dim res As TypeSearchItemResult2
  Dim i As Long
  Dim j As Long
  Dim limitJ As Long
  Dim isStackable As Boolean
  Dim tileID As Long
  tileID = GetTheLong(t1, t2)
  isStackable = DatTiles(tileID).stackable
  res.foundcount = 0
  res.bpID = &HFF
  res.slotID = &HFF
  res.amount = 0
  res.b4 = 0
  i = CLng(depotBpID)
  If (Backpack(idConnection, i).open = True) Then
    If (isStackable = True) Then
      limitJ = (Backpack(idConnection, i).used) - 1
      For j = 0 To limitJ
        If Backpack(idConnection, i).item(j).t1 = t1 And _
         Backpack(idConnection, i).item(j).t2 = t2 And _
         Backpack(idConnection, i).item(j).t3 < &H64 Then
          res.foundcount = res.foundcount + 1
          res.bpID = CByte(i)
          res.slotID = CByte(j)
          res.amount = Backpack(idConnection, i).item(j).t3
          res.b4 = Backpack(idConnection, i).item(j).t4
          SearchItemDestinationInDepot = res
          Exit Function
        End If
      Next j
      End If
      If Backpack(idConnection, i).used < Backpack(idConnection, i).cap Then
          res.foundcount = res.foundcount + 1
          res.bpID = CByte(i)
          res.slotID = CByte(Backpack(idConnection, i).used)
          res.amount = 0
          res.b4 = 0
          SearchItemDestinationInDepot = res
          Exit Function
      End If
    End If
  res.b1 = t1
  res.b2 = t2
  SearchItemDestinationInDepot = res
End Function

Public Function SearchFood(idConnection As Integer) As TypeSearchItemResult2
  ' search food in any open container
  ' return last found
  ' and count
  Dim res As TypeSearchItemResult2
  Dim i As Long
  Dim j As Long
  Dim tileID As Long
  Dim limitJ As Long
  res.foundcount = 0
  res.bpID = &HFF
  res.slotID = &HFF
  res.b1 = 0
  res.b2 = 0
  res.amount = 0
  res.b4 = 0
  For i = 0 To HIGHEST_BP_ID
    limitJ = (Backpack(idConnection, i).used) - 1
    For j = 0 To limitJ
      tileID = GetTheLong(Backpack(idConnection, i).item(j).t1, Backpack(idConnection, i).item(j).t2)
      If DatTiles(tileID).isFood = True Then
        res.foundcount = res.foundcount + 1
        res.bpID = CByte(i)
        res.slotID = CByte(j)
        res.b1 = Backpack(idConnection, i).item(j).t1
        res.b2 = Backpack(idConnection, i).item(j).t2
        res.amount = Backpack(idConnection, i).item(j).t3
        res.b4 = Backpack(idConnection, i).item(j).t4
      End If
    Next j
  Next i
  SearchFood = res
End Function

' < The dictionary tibia IDs -> name string >
Public Sub AddIDname(idConnection As Integer, tibiaID As Double, mobName As String)
  ' add item to dictionary
  Dim res As Boolean
  NameOfID(idConnection).item(tibiaID) = mobName
End Sub
Public Sub RemoveID(idConnection As Integer, tibiaID As Double)
  ' remove item from dictionary
  Dim res As Boolean
  If NameOfID(idConnection).Exists(tibiaID) = True Then
    NameOfID(idConnection).Remove (tibiaID)
  End If
End Sub
Public Function GetNameFromID(idConnection As Integer, tibiaID As Double) As String
  ' get the name from an ID
  Dim aRes As Long
  Dim res As Boolean
  If NameOfID(idConnection).Exists(tibiaID) = True Then
    GetNameFromID = NameOfID(idConnection).item(tibiaID)
  Else
    'If givenUFO = False Then
    '  aRes = GiveGMmessage(idConnection, "UFO detected in truemap! (creature with no name), with ID : " & CStr(tibiaID), "BlackdProxy")
    '  DoEvents
    '  givenUFO = True
    'End If
    GetNameFromID = ""
  End If
End Function


' < The dictionary tibia IDs -> HP >
Public Sub AddID_HP(idConnection As Integer, tibiaID As Double, HP As Byte)
  ' add item to dictionary
  Dim res As Boolean
  HPOfID(idConnection).item(tibiaID) = HP
End Sub
Public Sub RemoveID_HP(idConnection As Integer, tibiaID As Double)
  ' remove item from dictionary
  Dim res As Boolean
  If HPOfID(idConnection).Exists(tibiaID) = True Then
    HPOfID(idConnection).Remove (tibiaID)
  End If
End Sub
Public Function GetHPFromID(idConnection As Integer, tibiaID As Double) As Byte
  ' get the name from an ID
  Dim aRes As Long
  Dim res As Boolean
  If HPOfID(idConnection).Exists(tibiaID) = True Then
    GetHPFromID = HPOfID(idConnection).item(tibiaID)
  Else
    'If givenUFO = False Then
    '  aRes = GiveGMmessage(idConnection, "UFO detected in truemap! (creature with no name), with ID : " & CStr(tibiaID), "BlackdProxy")
    '  DoEvents
    '  givenUFO = True
    'End If
    GetHPFromID = &H64 'default = full hp
  End If
End Function





' < The dictionary tibia IDs -> direction >
Public Sub AddID_Direction(idConnection As Integer, tibiaID As Double, direction As Byte)
  ' add item to dictionary
  Dim res As Boolean
  DirectionOfID(idConnection).item(tibiaID) = direction
End Sub
Public Sub RemoveID_Direction(idConnection As Integer, tibiaID As Double)
  ' remove item from dictionary
  Dim res As Boolean
  If DirectionOfID(idConnection).Exists(tibiaID) = True Then
    DirectionOfID(idConnection).Remove (tibiaID)
  End If
End Sub
Public Function GetDirectionFromID(idConnection As Integer, tibiaID As Double) As Byte
  ' get the name from an ID
  Dim aRes As Long
  Dim res As Boolean
  If DirectionOfID(idConnection).Exists(tibiaID) = True Then
    GetDirectionFromID = DirectionOfID(idConnection).item(tibiaID)
  Else
    'If givenUFO = False Then
    '  aRes = GiveGMmessage(idConnection, "UFO detected in truemap! (creature with no name), with ID : " & CStr(tibiaID), "BlackdProxy")
    '  DoEvents
    '  givenUFO = True
    'End If
    GetDirectionFromID = &H0  'default
  End If
End Function




Public Sub ShowPositionChange(Index As Integer)
  ' do any required update of position and map
  If Index = mapIDselected Then
    If TrialVersion = True Then
      If sentWelcome(mapIDselected) = False Or GotPacketWarning(mapIDselected) = True Then
        Exit Sub
      End If
    End If
    'update map
    If frmHardcoreCheats.chkAutoUpdateMap.Value = True Then
      If mapIDselected = Index Then
        If frmHardcoreCheats.chkLockOnMyFloor.Value = 1 Then
          mapFloorSelected = myZ(Index)
        End If
        frmTrueMap.SetButtonColours
        frmTrueMap.DrawFloor
      End If
    Else
     If mapIDselected = Index Then
        If mapFloorSelected <> myZ(Index) Then
          frmTrueMap.SetButtonColours
        End If
      End If
    End If
  End If
End Sub
Public Function GetTheMobileInfo(idConnection As Integer, ByRef packet() As Byte, firstPos As Long) As TypePlayerInfo
  ' there is a lot of info in this subpacket, however I only will take the important info
  Dim resF As TypePlayerInfo
  Dim lon As Long
  Dim i As Long
  Dim strangeID As Double
  Dim newID As Double
  Dim outfitType As String
  Dim tileID As Long
  Dim name As String
  Dim originalPos As Long
  Dim debugStr As String
  ' tibia 10.36 full: 00 00 AD 01 FF 61 00 00 00 00 00 7A 01 00 40 02 04 00 4E 61 6A 69 64 00 81 00 39 71 5F 71 00 00 00 00 00 32 00 00 00 00 02 02 FF 00 00 01 00 FF
  ' tibia 10.36 mobile info=61 00 00 00 00 00 7A 01 00 40 02 04 00 4E 61 6A 69 64 00 81 00 39 71 5F 71 00 00 00 00 00 32 00 00 00 00 02 02 FF 00 00 01 00 FF
  originalPos = firstPos
  
  strangeID = FourBytesDouble(packet(firstPos + 2), packet(firstPos + 3), packet(firstPos + 4), packet(firstPos + 5))
  newID = FourBytesDouble(packet(firstPos + 6), packet(firstPos + 7), packet(firstPos + 8), packet(firstPos + 9))
  resF.newID = newID
  
  ' remove "old" ID , why? don't know, but that is what server want
  RemoveID idConnection, strangeID
  RemoveID_HP idConnection, strangeID
  RemoveID_Direction idConnection, strangeID
  resF.pos = firstPos + 10
  If TibiaVersionLong >= 872 Then
    resF.pos = resF.pos + 1 'unknown new byte
   ' 8.71: 61 00 00 00 00 00 EF B9 ?? 42 01 0E 00 xxx 64 02 90 00 72 84 52 84 01 00 00 06 41 F4 01 00 00 00 01 00 FF
   ' 8.72: 61 00 00 00 00 00 A8 02 00 40 02 08 00 xxx 64 03 39 00 00 00 00 00 00 00 00 00 00 64 00 00 00 00 01 00 FF
   ' 9.9 : 61 00 00 00 00 00 01 F3 01 40 01 03 00 xxx 64 01 14 01 00 00 00 00 00 00 00 00 00 3E 00 00 00 00 01 FF 00 00 01 00 FF
   ' 9.9 : 61 00 00 00 00 00 01 F3 01 40 01 03 00 xxx 64 03 14 01 00 00 00 00 00 00 00 00 00 3E 00 00 00 00 01 FF 00 00 01 00 FF
         ' 61 00 00 00 00 00 01 F3 01 40 01 03 00 xxx 64 03 14 01 00 00 00 00 00 00 00 00 00 3E 00 00 00 00 01 FF 00 00 01
         ' 61 00 00 00 00 00 3B 02 00 40 02 06 00 xxx 64 02 81 00 61 4D 57 73 00 00 00 00 00 30 00 00 00 00 02 FF 00 00 01
   ' 10.36'61 00 00 00 00 00 7A 01 00 40 02 04 00 xxx 64 00 81 00 39 71 5F 71 00 00 00 00 00 32 00 00 00 00 02 02 FF 00 00 01 00 FF
  End If
  lon = GetTheLong(packet(resF.pos), packet(1 + resF.pos))
  resF.pos = resF.pos + 2
  
  ' get the name of the mobile
  name = ""
  For i = resF.pos To -1 + lon + resF.pos
    name = name & Chr(packet(i))
  Next i
  
  ' add new ID
  AddIDname idConnection, newID, name
  resF.pos = resF.pos + lon
  
  ' give position after having read mobile info
  resF.stackpos = CLng(packet(resF.pos + 1))
  AddID_HP idConnection, newID, packet(resF.pos)
  'Debug.Print "direction4=" & GoodHex(packet(resF.pos + 1)) & " " & GoodHex(packet(resF.pos)) & " " & GoodHex(packet(resF.pos - 1)) & " " & GoodHex(packet(resF.pos - 2)) & " " & GoodHex(packet(resF.pos - 3))
  AddID_Direction idConnection, newID, packet(resF.pos + 1)
  
  resF.pos = resF.pos + 2 ' skip hp / direction?
  If TibiaVersionLong <= 760 Then
    outfitType = CLng(packet(resF.pos))
  Else
    outfitType = GetTheLong(packet(resF.pos), packet(resF.pos + 1))
    resF.pos = resF.pos + 1
  End If
   ' the change in 853
  
  ' now the outfit
  If outfitType = &H0 Then ' thing outfit
    If (packet(resF.pos + 1) = &H0) And (packet(resF.pos + 2) = &H0) Then
      'unhide invis beings
      If (resF.newID <> myID(idConnection)) And (frmHardcoreCheats.chkReveal.Value = 1) Then
        packet(resF.pos + 1) = LowByteOfLong(tileID_Oracle)
        packet(resF.pos + 2) = HighByteOfLong(tileID_Oracle)
      End If
    End If
    resF.pos = resF.pos + 3
  Else
    resF.pos = resF.pos + 5
    If TibiaVersionLong >= 773 Then
      resF.pos = resF.pos + 1 ' new strange thing
    End If
  End If

  resF.pos = resF.pos + 6 ' skip light,speed,etc
  
  If TibiaVersionLong >= 853 Then ' 1
    resF.pos = resF.pos + 1 ' skip all at once
  End If
  If TibiaVersionLong >= 854 Then ' 1
    resF.pos = resF.pos + 1 ' skip one more
  End If

  If TibiaVersionLong >= 870 Then
    resF.pos = resF.pos + 2 ' xxx2 skip 2 more
  End If
  
  If TibiaVersionLong >= 990 Then
    resF.pos = resF.pos + 4 '  skip 4 more
  End If
  
    If TibiaVersionLong >= 1036 Then
    resF.pos = resF.pos + 1 '  skip 1 more
  End If
'  For i = originalPos To resF.pos - 1
 '   debugStr = debugStr & " " & GoodHex(packet(i))
 ' Next i
 ' Debug.Print "[" & Trim$(debugStr) & "]"
  
  GetTheMobileInfo = resF
End Function
Public Function MinV(v1 As Long, v2 As Long) As Long
  ' min between 2 numbers
  If v1 < v2 Then
    MinV = v1
  Else
    MinV = v2
  End If
End Function
Public Function ReadMap(idConnection As Integer, ByRef packet() As Byte, firstByte As Long) As Long
  ' Read first map
  ' firstByte of first map packet should be &H64
  Dim pos As Long ' packet position
  Dim idTile As Long
  Dim count As Long
  Dim resF As TypePlayerInfo
  Dim strRes As String 'temp
  Dim zstep As Long
  Dim startz As Long 'we will get info from this floor ...
  Dim endz As Long ' ... to this floor
  Dim z As Long 'my z
  Dim Nfloors As Long ' number of floors we will get in the packet
  Dim expectedPositions As Long ' expected map positions info (including skiped positions)
  Dim posX As Long
  Dim posY As Long
  Dim posZ As Long
  Dim nx As Long
  Dim ny As Long
  Dim nz As Long
  Dim skipcount As Long
  Dim tmpdebugstrange As Long
  Dim resT As Long
  Dim resSkip As Long
  #If FinalMode = 1 Then
  On Error GoTo badbug
  #End If
  pos = firstByte
 ' Debug.Print frmMain.showAsStr3(packet, True, pos, pos + 100)
 
  ' set my x,y,z
    myX(idConnection) = GetTheLong(packet(pos + 1), packet(pos + 2))
    myY(idConnection) = GetTheLong(packet(pos + 3), packet(pos + 4))
    myZ(idConnection) = CLng(packet(pos + 5))
    ' skip position bytes
    pos = pos + 6

  
  ' work with a local z
  z = myZ(idConnection)
  ' two cases: you are underground (z>7) or not
  If (z > 7) Then
    startz = z - 2
     endz = MinV(15, z + 2) ' there is a special case on the most deep. This deal with that
    zstep = 1 ' floors will be given from lower to higher
  Else
    startz = 7
    endz = 0
    zstep = -1 '  floors will be given from higher to lower
  End If

  ' evaluate expected map positions to be read in the first packet
  Nfloors = Abs(startz - endz) + 1
  'expectedPositions = Nfloors * 252


  ' init counters
  count = 0
  skipcount = 0
  ' ENTER THE MATRIX!!
   #If MapDebug = 1 Then
     OverwriteOnFileSimple "mapdebug.txt", "Trying to read map. Expecting to read " & CStr(Nfloors * 252) & " positions"
   #End If
  For nz = startz To endz Step zstep
    For nx = -8 To 9
      For ny = -6 To 7

        If skipcount = 0 Then
          If packet(pos + 1) = &HFF Then 'first we could have a skipper (&H?? &HFF)
            skipcount = skipcount + packet(pos)
            #If MapDebug = 1 Then
              AddwriteOnFileSimple "mapdebug.txt", "[SKIPER: " & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1)) & "]"
            #End If
            pos = pos + 2

            'Matrix(ny, nx, nz, idconnection).s(0).t1 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).t2 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).t3 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).dblID = 0
            RtlMoveMemory Matrix(ny, nx, nz, idConnection), tmpStack, OptCte3
            count = count + 1
          Else 'else we have info about ground tile
'            If count = expectedPositions Then
'              Debug.Print "hey"
'            End If
            'If count < expectedPositions Then
                resT = ReadSinglePosition(idConnection, nx, ny, nz, packet, pos)
                If resT = -1 Then
                  GoTo badbug
                Else
                  pos = resT
                End If
                resSkip = packet(pos)
                If resSkip > 0 Then
                    skipcount = skipcount + resSkip
                    #If MapDebug = 1 Then
                      AddwriteOnFileSimple "mapdebug.txt", "[SKIPING " & skipcount & "]"
                    #End If
                End If
                pos = pos + 2
                count = count + 1
           ' End If
          End If
        Else
          ' skip a map position (no info)
          ' the TrueMap module will read a ground tile &H00 &H00 as "no info" -> colour black
          count = count + 1


          skipcount = skipcount - 1
          
          #If MapDebug = 1 Then
            AddwriteOnFileSimple "mapdebug.txt", "[POSITION " & CStr(count) & " SKIPED, REAMINING SKIPS=" & CStr(skipcount) & "]"
          #End If
          'Matrix(ny, nx, nz, idconnection).s(0).t1 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).t2 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).t3 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).dblID = 0
          RtlMoveMemory Matrix(ny, nx, nz, idConnection), tmpStack, OptCte3
        End If
      Next ny
    Next nx
  Next nz
  
  #If MapDebug = 1 Then
    AddwriteOnFileSimple "mapdebug.txt", "READ COMPLETED SUCESSFULLY!"
  #End If
  ' update map for this idConnection
  ' if we want debug...
  ' LogOnFile "res.txt", "p(pos)=" & GoodHex(packet(pos)) & " (expected 83) ; count = " & count & " (expected " & expectedPositions & ")"
  ' return packet position after reading all the map
  ReadMap = pos
  Exit Function
badbug:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "BAD READ IN MAP IN THIS POSITION: " & pos
  LogOnFile "errors.txt", "BAD READ IN MAP IN THIS POSITION: " & pos
  ReadMap = pos
End Function

Public Function ReadNewFloors(idConnection As Integer, ByRef packet() As Byte, firstByte As Long, startz As Long, endz As Long, zstep As Long) As Long
  ' Read map update when changing floors
  ' firstByte of first map packet should be &H64
  Dim pos As Long ' packet position
  Dim idTile As Long
  Dim count As Long
  Dim resF As TypePlayerInfo
  Dim strRes As String 'temp
  Dim z As Long 'my z
  Dim Nfloors As Long ' number of floors we will get in the packet

  Dim posX As Long
  Dim posY As Long
  Dim posZ As Long
  Dim nx As Long
  Dim ny As Long
  Dim nz As Long
  Dim skipcount As Long
  Dim resT As Long
  #If FinalMode Then
  On Error GoTo badbug
  #End If
  pos = firstByte + 1

  ' init counters
  count = 0
  skipcount = 0
  ' ENTER THE MATRIX!!
  For nz = startz To endz Step zstep
    For nx = -8 To 9
      For ny = -6 To 7

        If skipcount = 0 Then
          If packet(pos + 1) = &HFF Then 'first we could have a skipper (&H?? &HFF)
            skipcount = skipcount + packet(pos)
            'Debug.Print ">>" & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1)) & "<<"
            pos = pos + 2

            'Matrix(ny, nx, nz, idconnection).s(0).t1 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).t2 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).t3 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).dblID = 0
            RtlMoveMemory Matrix(ny, nx, nz, idConnection), tmpStack, OptCte3
            count = count + 1
          Else 'else we have info about ground tile
'            If count = expectedPositions Then
'              Debug.Print "hey"
'            End If
            'If count < expectedPositions Then
                resT = ReadSinglePosition(idConnection, nx, ny, nz, packet, pos)
                If resT = -1 Then
                  GoTo badbug
                Else
                  pos = resT
                End If

                skipcount = skipcount + packet(pos)
                pos = pos + 2
                count = count + 1
           ' End If
          End If
        Else
          ' skip a map position (no info)
          ' the TrueMap module will read a ground tile &H00 &H00 as "no info" -> colour black
          count = count + 1
          skipcount = skipcount - 1
          'Matrix(ny, nx, nz, idconnection).s(0).t1 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).t2 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).t3 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).dblID = 0
          RtlMoveMemory Matrix(ny, nx, nz, idConnection), tmpStack, OptCte3
        End If
      Next ny
    Next nx
  Next nz
  ReadNewFloors = pos
  Exit Function
badbug:
  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "BAD READ: " & frmMain.showAsStr2(packet, 0) & " LAST GOOD POS: " & pos
  ReadNewFloors = 10000
End Function


Public Function GetMatrixPosition(idConnection As Integer, X As Long, y As Long, z As Long, s As Long) As TypeMatrixPosition
  ' get relative positions to our matrix
  Dim res As TypeMatrixPosition
  res.X = X - myX(idConnection) + z - myZ(idConnection)
  res.y = y - myY(idConnection) + z - myZ(idConnection)
  res.z = z
  res.s = s
  If (res.X >= -8) And (res.X <= 9) And (res.y >= -6) And (res.y <= 7) _
   And (res.z >= 0) And (res.z <= 15) And (res.s >= 0) And (res.s <= 10) Then
    ' valid matrix position
    res.valid = True
  Else
    res.valid = False
  End If
  GetMatrixPosition = res
End Function

Public Function RemoveThingFromStack(idConnection As Integer, X As Long, y As Long, z As Long, s As Long) As Long
  ' remove something from stack
  Dim i As Long
  Dim res As Long

  #If FinalMode Then
  On Error GoTo gotError
  #End If
  res = -1
  If ((X > 9) Or (X < -8) Or (y > 7) Or (y < -6) Or (z < 0) Or (z > 15)) Then
    RemoveThingFromStack = res
    Exit Function
  End If

  For i = s To 9
    Matrix(y, X, z, idConnection).s(i).t1 = Matrix(y, X, z, idConnection).s(i + 1).t1
    Matrix(y, X, z, idConnection).s(i).t2 = Matrix(y, X, z, idConnection).s(i + 1).t2
    Matrix(y, X, z, idConnection).s(i).t3 = Matrix(y, X, z, idConnection).s(i + 1).t3
    Matrix(y, X, z, idConnection).s(i).t4 = Matrix(y, X, z, idConnection).s(i + 1).t4
    Matrix(y, X, z, idConnection).s(i).dblID = Matrix(y, X, z, idConnection).s(i + 1).dblID
  Next i
  Matrix(y, X, z, idConnection).s(10).t1 = &H0
  Matrix(y, X, z, idConnection).s(10).t2 = &H0
  Matrix(y, X, z, idConnection).s(10).t3 = &H0
  Matrix(y, X, z, idConnection).s(10).t4 = &H0
  Matrix(y, X, z, idConnection).s(10).dblID = 0
  res = 0
gotError:
  RemoveThingFromStack = res
End Function
Public Function AddThingToStack(idConnection As Integer, X As Long, y As Long, z As Long, _
 t1 As Byte, t2 As Byte, t3 As Byte, dblID As Double, Optional t4 As Byte = &H0) As Long
  ' this add anything to the stack
  ' + will return stackPos that received
  Dim i As Long
  Dim j As Long
  Dim res As Long
  Dim tileID As Long
  Dim tempTileID As Long
  Dim newItemPriority As Long
  Dim currentPosPriority As Long
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  res = -1
  tileID = GetTheLong(t1, t2)
  If tempTileID > highestDatTile Then
    'Debug.Print "AddThingToStack failed!"
    AddThingToStack = -1
  End If
  newItemPriority = DatTiles(tileID).stackPriority
  i = 0
    If TibiaVersionLong >= 870 Then
        Do
          i = i + 1
          tempTileID = GetTheLong(Matrix(y, X, z, idConnection).s(i).t1, Matrix(y, X, z, idConnection).s(i).t2)
          currentPosPriority = DatTiles(tempTileID).stackPriority
          If i = 10 Then
            Exit Do
          End If
        Loop Until (newItemPriority > currentPosPriority) ' strictly > since tibia 8.7
    Else
        Do
          i = i + 1
          tempTileID = GetTheLong(Matrix(y, X, z, idConnection).s(i).t1, Matrix(y, X, z, idConnection).s(i).t2)
          currentPosPriority = DatTiles(tempTileID).stackPriority
          If i = 10 Then
            Exit Do
          End If
        Loop Until (newItemPriority >= currentPosPriority)
  End If
  For j = 10 To i Step -1
    Matrix(y, X, z, idConnection).s(j).t1 = Matrix(y, X, z, idConnection).s(j - 1).t1
    Matrix(y, X, z, idConnection).s(j).t2 = Matrix(y, X, z, idConnection).s(j - 1).t2
    Matrix(y, X, z, idConnection).s(j).t3 = Matrix(y, X, z, idConnection).s(j - 1).t3
    Matrix(y, X, z, idConnection).s(j).t4 = Matrix(y, X, z, idConnection).s(j - 1).t4
    Matrix(y, X, z, idConnection).s(j).dblID = Matrix(y, X, z, idConnection).s(j - 1).dblID
  Next j
  Matrix(y, X, z, idConnection).s(i).t1 = t1
  Matrix(y, X, z, idConnection).s(i).t2 = t2
  Matrix(y, X, z, idConnection).s(i).t3 = t3
  Matrix(y, X, z, idConnection).s(i).t4 = t4
  Matrix(y, X, z, idConnection).s(i).dblID = dblID
  ' MODIFIED at for 8.4 TO FIX STACK OF PERSONS BUG
  Matrix(y, X, z, idConnection).s(10).t1 = &H0
  Matrix(y, X, z, idConnection).s(10).t2 = &H0
  Matrix(y, X, z, idConnection).s(10).t3 = &H0
  Matrix(y, X, z, idConnection).s(10).t4 = &H0
  Matrix(y, X, z, idConnection).s(10).dblID = 0
  res = CLng(i)
gotError:
  AddThingToStack = res
End Function


Public Sub EvalMyMove(idConnection As Integer, increaseX As Long, increaseY As Long, increaseZ As Long)
  ' move true map
  ' a side of the map will be a blank line for some mseconds
  ' until the updated info fills it
  Dim X As Long
  Dim y As Long
  Dim z As Long
  Dim s As Long
  Dim nx As Long
  Dim ny As Long
  Dim nz As Long
  Dim x1 As Long
  Dim x2 As Long
  Dim x3 As Long
  Dim y1 As Long
  Dim y2 As Long
  Dim y3 As Long
  Dim z1 As Long
  Dim z2 As Long
  Dim z3 As Long
  z1 = 0
  z2 = 15
  z3 = 1
  If increaseZ <> 0 Then
    increaseX = increaseX + increaseZ
    increaseY = increaseY + increaseZ
    increaseZ = 0
    Exit Sub
  End If
  If increaseX > 0 Then
    x1 = -8
    x2 = 9
    x3 = 1
  Else
    x1 = 9
    x2 = -8
    x3 = -1
  End If
  If increaseY > 0 Then
    y1 = -6
    y2 = 7
    y3 = 1
  Else
    y1 = 7
    y2 = -6
    y3 = -1
  End If

  'OptCte1 = 238 * LenB(tmpStack)
  'OptCte2 = 13 * LenB(tmpStack)
  'OptCte3 = LenB(tmpStack)
  If increaseX = 1 Then
    For z = 0 To 15
      RtlMoveMemory Matrix(-6, -8, z, idConnection), Matrix(-6, -7, z, idConnection), OptCte1
      ' draw black line
      For y = -6 To 7
        RtlMoveMemory Matrix(y, 9, z, idConnection), tmpStack, OptCte3
      Next y
    Next z
  ElseIf increaseX = -1 Then
    For z = 0 To 15
      RtlMoveMemory Matrix(-6, -7, z, idConnection), Matrix(-6, -8, z, idConnection), OptCte1
      ' draw black line
      For y = -6 To 7
        RtlMoveMemory Matrix(y, -8, z, idConnection), tmpStack, OptCte3
      Next y
    Next z
  End If
  If increaseY = 1 Then
    For z = 0 To 15
      For X = -8 To 9
        RtlMoveMemory Matrix(-6, X, z, idConnection), Matrix(-5, X, z, idConnection), OptCte2
      Next X
      ' draw black line
      For X = -8 To 9
        RtlMoveMemory Matrix(7, X, z, idConnection), tmpStack, OptCte3
      Next X
    Next z
  ElseIf increaseY = -1 Then
    For z = 0 To 15
      For X = -8 To 9
        RtlMoveMemory Matrix(-5, X, z, idConnection), Matrix(-6, X, z, idConnection), OptCte2
      Next X
      ' draw black line
      For X = -8 To 9
        RtlMoveMemory Matrix(-6, X, z, idConnection), tmpStack, OptCte3
      Next X
    Next z
  End If
  
End Sub
Public Function UpdateRightSide(idConnection As Integer, packet() As Byte, startPos As Long) As Long
  ' update right side of the map
  Dim pos As Long
  Dim nx As Long
  Dim ny As Long
  Dim nz As Long
  Dim skipcount As Long
  Dim count As Long
  Dim idTile As Long
  Dim resF As TypePlayerInfo
  Dim nameofgivenID As String
  Dim z As Long
  Dim startz As Long
  Dim endz As Long
  Dim zstep As Long
  Dim resT As Long
  Dim tmpdebugstrange As Long
  Dim Nfloors As Long
  count = 0
  skipcount = 0
  pos = startPos + 1 ' skip type byte
  nx = 9
  ny = myY(idConnection)
  z = myZ(idConnection)
  ' two cases: you are underground (z>7) or not
  If (z > 7) Then
    startz = z - 2
    endz = MinV(15, z + 2) ' there is a special case on the most deep. This deal with that
    zstep = 1 ' floors will be given from lower to higher
  Else
    startz = 7
    endz = 0
    zstep = -1 '  floors will be given from higher to lower
  End If
  Nfloors = Abs(startz - endz) + 1
  
   #If MapDebug = 1 Then
     OverwriteOnFileSimple "mapdebug.txt", "Trying to read right update. Expecting to read " & CStr(Nfloors * 14) & " positions"
   #End If
  
  
  For nz = startz To endz Step zstep
    For ny = -6 To 7

 
        If skipcount = 0 Then
          If packet(pos + 1) = &HFF Then 'first we could have a skipper (&H?? &HFF)
            skipcount = skipcount + packet(pos)
            'Debug.Print ">>" & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1)) & "<<"
            pos = pos + 2

            'Matrix(ny, nx, nz, idconnection).s(0).t1 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).t2 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).t3 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).dblID = 0
            RtlMoveMemory Matrix(ny, nx, nz, idConnection), tmpStack, OptCte3
            count = count + 1
          Else 'else we have info about ground tile
'            If count = expectedPositions Then
'              Debug.Print "hey"
'            End If
            'If count < expectedPositions Then
                resT = ReadSinglePosition(idConnection, nx, ny, nz, packet, pos)
                If resT = -1 Then
                  GoTo badbug
                Else
                  pos = resT
                End If

                skipcount = skipcount + packet(pos)
                pos = pos + 2
                count = count + 1
            'End If
          End If
        Else
          ' skip a map position (no info)
          ' the TrueMap module will read a ground tile &H00 &H00 as "no info" -> colour black
          count = count + 1
          skipcount = skipcount - 1
          'Matrix(ny, nx, nz, idconnection).s(0).t1 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).t2 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).t3 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).dblID = 0
          RtlMoveMemory Matrix(ny, nx, nz, idConnection), tmpStack, OptCte3
        End If
    Next ny
  Next nz
  #If MapDebug = 1 Then
    AddwriteOnFileSimple "mapdebug.txt", "SOUTH UPDATE COMPLETED SUCESSFULLY!"
  #End If
  
  UpdateRightSide = pos
  Exit Function
badbug:
  UpdateRightSide = 10000
End Function
Public Function UpdateLeftSide(idConnection As Integer, packet() As Byte, startPos As Long) As Long
  ' update left side of the map
  Dim pos As Long
  Dim nx As Long
  Dim ny As Long
  Dim nz As Long
  Dim skipcount As Long
  Dim count As Long
  Dim idTile As Long
  Dim resF As TypePlayerInfo
  Dim nameofgivenID As String
  Dim z As Long
  Dim startz As Long
  Dim endz As Long
  Dim zstep As Long
  Dim resT As Long
  Dim tmpdebugstrange As Long
  count = 0
  skipcount = 0
  pos = startPos + 1 ' skip type byte
  nx = -8
  ny = myY(idConnection)
  z = myZ(idConnection)
  ' two cases: you are underground (z>7) or not
  If (z > 7) Then
    startz = z - 2
    endz = MinV(15, z + 2) ' there is a special case on the most deep. This deal with that
    zstep = 1 ' floors will be given from lower to higher
  Else
    startz = 7
    endz = 0
    zstep = -1 '  floors will be given from higher to lower
  End If
  For nz = startz To endz Step zstep
    For ny = -6 To 7

        If skipcount = 0 Then
          If packet(pos + 1) = &HFF Then 'first we could have a skipper (&H?? &HFF)
            skipcount = skipcount + packet(pos)
            'Debug.Print ">>" & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1)) & "<<"
            pos = pos + 2

            'Matrix(ny, nx, nz, idconnection).s(0).t1 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).t2 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).t3 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).dblID = 0
            RtlMoveMemory Matrix(ny, nx, nz, idConnection), tmpStack, OptCte3
            count = count + 1
          Else 'else we have info about ground tile
'            If count = expectedPositions Then
'              Debug.Print "hey"
'            End If
           ' If count < expectedPositions Then
                resT = ReadSinglePosition(idConnection, nx, ny, nz, packet, pos)
                If resT = -1 Then
                  GoTo badbug
                Else
                  pos = resT
                End If

                skipcount = skipcount + packet(pos)
                pos = pos + 2
                count = count + 1
            'End If
          End If
        Else
          ' skip a map position (no info)
          ' the TrueMap module will read a ground tile &H00 &H00 as "no info" -> colour black
          count = count + 1
          skipcount = skipcount - 1
          'Matrix(ny, nx, nz, idconnection).s(0).t1 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).t2 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).t3 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).dblID = 0
          RtlMoveMemory Matrix(ny, nx, nz, idConnection), tmpStack, OptCte3
        End If
    Next ny
  Next nz
  UpdateLeftSide = pos
  Exit Function
badbug:
  UpdateLeftSide = 10000
End Function
Public Function UpdateNorthSide(idConnection As Integer, packet() As Byte, startPos As Long) As Long
  ' update north side of the map
  Dim pos As Long
  Dim nx As Long
  Dim ny As Long
  Dim nz As Long
  Dim stackpos As Long
  Dim skipcount As Long
  Dim count As Long
  Dim idTile As Long
  Dim resF As TypePlayerInfo
  Dim nameofgivenID As String
  Dim z As Long
  Dim startz As Long
  Dim endz As Long
  Dim zstep As Long
  Dim resT As Long
  Dim tmpdebugstrange As Long
  count = 0
  skipcount = 0
  pos = startPos + 1 ' skip type byte
  nx = myX(idConnection)
  ny = -6
  z = myZ(idConnection)
  ' two cases: you are underground (z>7) or not
  If (z > 7) Then
    startz = z - 2
    endz = MinV(15, z + 2) ' there is a special case on the most deep. This deal with that
    zstep = 1 ' floors will be given from lower to higher
  Else
    startz = 7
    endz = 0
    zstep = -1 '  floors will be given from higher to lower
  End If
  For nz = startz To endz Step zstep
    For nx = -8 To 9

        If skipcount = 0 Then
          If packet(pos + 1) = &HFF Then 'first we could have a skipper (&H?? &HFF)
            skipcount = skipcount + packet(pos)
            'Debug.Print ">>" & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1)) & "<<"
            pos = pos + 2

            'Matrix(ny, nx, nz, idconnection).s(0).t1 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).t2 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).t3 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).dblID = 0
            RtlMoveMemory Matrix(ny, nx, nz, idConnection), tmpStack, OptCte3
            count = count + 1
          Else 'else we have info about ground tile
'            If count = expectedPositions Then
'              Debug.Print "hey"
'            End If
            ' If count < expectedPositions Then
                resT = ReadSinglePosition(idConnection, nx, ny, nz, packet, pos)
                If resT = -1 Then
                  GoTo badbug
                Else
                  pos = resT
                End If

                skipcount = skipcount + packet(pos)
                pos = pos + 2
                count = count + 1
          '  End If
          End If
        Else
          ' skip a map position (no info)
          ' the TrueMap module will read a ground tile &H00 &H00 as "no info" -> colour black
          count = count + 1
          skipcount = skipcount - 1
          'Matrix(ny, nx, nz, idconnection).s(0).t1 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).t2 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).t3 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).dblID = 0
          RtlMoveMemory Matrix(ny, nx, nz, idConnection), tmpStack, OptCte3
        End If
    Next nx
  Next nz
  UpdateNorthSide = pos
  Exit Function
badbug:
  UpdateNorthSide = 10000
End Function
Public Function UpdateSouthSide(idConnection As Integer, packet() As Byte, startPos As Long) As Long
  ' update south side of the map
  Dim pos As Long
  Dim nx As Long
  Dim ny As Long
  Dim nz As Long
  Dim skipcount As Long
  Dim count As Long
  Dim idTile As Long
  Dim resF As TypePlayerInfo
  Dim nameofgivenID As String
  Dim z As Long
  Dim startz As Long
  Dim endz As Long
  Dim zstep As Long
  Dim resT As Long
  Dim debugstrange As Long
  Dim Nfloors As Long
  count = 0
  skipcount = 0
  pos = startPos + 1 ' skip type byte
  nx = myX(idConnection)
  ny = 7
  z = myZ(idConnection)
  ' two cases: you are underground (z>7) or not
  If (z > 7) Then
    startz = z - 2
    endz = MinV(15, z + 2) ' there is a special case on the most deep. This deal with that
    zstep = 1 ' floors will be given from lower to higher
  Else
    startz = 7
    endz = 0
    zstep = -1 '  floors will be given from higher to lower
  End If
  Nfloors = Abs(startz - endz) + 1
  
   #If MapDebug = 1 Then
     OverwriteOnFileSimple "mapdebug.txt", "Trying to read south update. Expecting to read " & CStr(Nfloors * 18) & " positions"
   #End If
  
  For nz = startz To endz Step zstep
    For nx = -8 To 9

        If skipcount = 0 Then
          If packet(pos + 1) = &HFF Then 'first we could have a skipper (&H?? &HFF)
            skipcount = skipcount + packet(pos)
            'Debug.Print ">>" & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1)) & "<<"
            pos = pos + 2

            'Matrix(ny, nx, nz, idconnection).s(0).t1 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).t2 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).t3 = &H0
            'Matrix(ny, nx, nz, idconnection).s(0).dblID = 0
            RtlMoveMemory Matrix(ny, nx, nz, idConnection), tmpStack, OptCte3
            count = count + 1
          Else 'else we have info about ground tile
'            If count = expectedPositions Then
'              Debug.Print "hey"
'            End If
'            If count < expectedPositions Then
                resT = ReadSinglePosition(idConnection, nx, ny, nz, packet, pos)
                If resT = -1 Then
                  GoTo badbug
                Else
                  pos = resT
                End If

                skipcount = skipcount + packet(pos)
                pos = pos + 2
                count = count + 1
'            End If
          End If
        Else
          ' skip a map position (no info)
          ' the TrueMap module will read a ground tile &H00 &H00 as "no info" -> colour black

          count = count + 1

          skipcount = skipcount - 1
          'Matrix(ny, nx, nz, idconnection).s(0).t1 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).t2 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).t3 = &H0
          'Matrix(ny, nx, nz, idconnection).s(0).dblID = 0
          RtlMoveMemory Matrix(ny, nx, nz, idConnection), tmpStack, OptCte3
        End If
    Next nx
  Next nz
  #If MapDebug = 1 Then
    AddwriteOnFileSimple "mapdebug.txt", "SOUTH UPDATE COMPLETED SUCESSFULLY!"
  #End If
  UpdateSouthSide = pos
  Exit Function
badbug:
  UpdateSouthSide = 10000
End Function


Public Function ReadSinglePositionOld(idConnection As Integer, nx As Long, ny As Long, nz As Long, ByRef packet() As Byte, pos As Long) As Long
 ' update a single square of the map
  Dim stackpos As Long
  Dim idTile As Long
  Dim resF As TypePlayerInfo
  Dim nameofgivenID As String
  Dim tempID As Double
  Dim tmpName As String
  Dim res As Long
  Dim outfit As Long
  Dim spaux As Long
  Dim tmpdebugstrange As Long
  Dim DEBUGpos As Long
  #If FinalMode Then
  On Error GoTo gotfail
  #End If
  res = -1
  stackpos = 0
  DEBUGpos = pos
  'If pos = 5107 Then
  'Debug.Print "warning at " & CStr(pos)
  'End If

  Matrix(ny, nx, nz, idConnection).s(0).t1 = packet(pos)
  Matrix(ny, nx, nz, idConnection).s(0).t2 = packet(pos + 1)
  Matrix(ny, nx, nz, idConnection).s(0).t3 = &H0
  Matrix(ny, nx, nz, idConnection).s(0).dblID = 0
  idTile = GetTheLong(packet(pos), packet(pos + 1))
  'skip 2 or 3 bytes depending of tile info
  ' updated since Blackd Proxy 16.3
  If (idTile = &H61) Or (idTile = &H62) Or (idTile = &H63) Then
    stackpos = 0
  Else
    'skip 2 or 3 bytes depending of tile info
    stackpos = 1
    If (idTile > highestDatTile) Then
      frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "ERROR WHILE READING MAP UPDATE ( TileID beyond limits : " & CStr(idTile) & ") last good pos=" & pos & " packet=" & frmMain.showAsStr2(packet, 0)
      ReadSinglePositionOld = res
      Exit Function
    End If
    If DatTiles(idTile).haveExtraByte = True Then
      pos = pos + 3 ' probably all ground tiles are only 2 bytes ... but who knows
    Else
      pos = pos + 2
    End If
  End If
  While packet(pos + 1) <> &HFF 'read until there is no more info for this position
    idTile = GetTheLong(packet(pos), packet(pos + 1))
   'Debug.Print CStr(stackpos) & ": " & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1))
    If idTile = &H61 Then
      'mobile info ; mobile = { player, monster, npc }
      'If packet(pos + 6) = &H1F Then
      'Debug.Print "STOP"
      'End If
      resF = GetTheMobileInfo(idConnection, packet, pos)
      If (resF.pos) - pos > 255 Then
        GoTo gotfail
      Else
        pos = resF.pos
      End If

      ' adding this mobile to the names string for this position of the matrix
      ' ReDim Matrix(-6 To 7, -8 To 9, 0 To 15, 1 To MAXCLIENTS
      If ((ny >= -6) And (ny <= 7) And (nx >= -8) And (nx <= 9) And (nz >= 0) And (nz <= 15)) Then
        Matrix(ny, nx, nz, idConnection).s(stackpos).t1 = &H61
        Matrix(ny, nx, nz, idConnection).s(stackpos).t2 = &H0
        Matrix(ny, nx, nz, idConnection).s(stackpos).t3 = &H0
        Matrix(ny, nx, nz, idConnection).s(stackpos).dblID = resF.newID
      End If
      stackpos = stackpos + 1
      ' an optional autologout if danger feature ...
      If frmHardcoreCheats.chkLogoutIfDanger.Value = 1 And sentFirstPacket(idConnection) = False Then
        If resF.newID <> myID(idConnection) Then
          ' proxy will tell the reason (names) of the logout in this case
          If LogoutReason(idConnection) = "" Then
            LogoutReason(idConnection) = GetNameFromID(idConnection, resF.newID)
          Else
            LogoutReason(idConnection) = LogoutReason(idConnection) & " , " & GetNameFromID(idConnection, resF.newID)
          End If
        End If
      End If
      ' AFTERLOGIN LOGOUT
      tmpName = GetNameFromID(idConnection, resF.newID)
      CheckIfGM idConnection, tmpName, nz
      If RuneMakerOptions(idConnection).autoLogoutAnyFloor = True Then
        If frmRunemaker.IsFriend(LCase(tmpName)) = False And tmpName <> CharacterName(idConnection) Then
          AfterLoginLogoutReason(idConnection) = tmpName
        End If
      ElseIf RuneMakerOptions(idConnection).autoLogoutCurrentFloor = True Then
        If nz = myZ(idConnection) Then
          If frmRunemaker.IsFriend(LCase(tmpName)) = False And tmpName <> CharacterName(idConnection) Then
            AfterLoginLogoutReason(idConnection) = tmpName
          End If
        End If
      End If
    ElseIf idTile = &H62 Then
      ' we already knew his ID + include some info
      tempID = FourBytesDouble(packet(pos + 2), packet(pos + 3), packet(pos + 4), packet(pos + 5))
      AddID_HP idConnection, tempID, packet(pos + 6) 'update hp
      nameofgivenID = GetNameFromID(idConnection, tempID)
      ' ReDim Matrix(-6 To 7, -8 To 9, 0 To 15, 1 To MAXCLIENTS
      If ((ny >= -6) And (ny <= 7) And (nx >= -8) And (nx <= 9) And (nz >= 0) And (nz <= 15)) Then
        Matrix(ny, nx, nz, idConnection).s(stackpos).t1 = &H61
        Matrix(ny, nx, nz, idConnection).s(stackpos).t2 = &H0
        Matrix(ny, nx, nz, idConnection).s(stackpos).t3 = &H0
        Matrix(ny, nx, nz, idConnection).s(stackpos).dblID = tempID
      End If
      ' AFTERLOGIN LOGOUT
      CheckIfGM idConnection, nameofgivenID, nz
      If RuneMakerOptions(idConnection).autoLogoutAnyFloor = True Then
        If frmRunemaker.IsFriend(LCase(nameofgivenID)) = False And nameofgivenID <> CharacterName(idConnection) Then
          AfterLoginLogoutReason(idConnection) = nameofgivenID
        End If
      ElseIf RuneMakerOptions(idConnection).autoLogoutCurrentFloor = True Then
        If nz = myZ(idConnection) Then
          If frmRunemaker.IsFriend(LCase(nameofgivenID)) = False And nameofgivenID <> CharacterName(idConnection) Then
            AfterLoginLogoutReason(idConnection) = nameofgivenID
          End If
        End If
      End If
      stackpos = stackpos + 1
      
      'eval outfit
      If TibiaVersionLong <= 760 Then
        outfit = CLng(packet(pos + 8))
      Else
        outfit = GetTheLong(packet(pos + 8), packet(pos + 9))
        pos = pos + 1
      End If
      
      If outfit = 0 Then
        If (packet(pos + 9) = &H0) And (packet(pos + 10) = &H0) Then
          If (nameofgivenID <> CharacterName(idConnection)) And (frmHardcoreCheats.chkReveal.Value = 1) Then
            packet(pos + 9) = LowByteOfLong(tileID_Oracle)
            packet(pos + 10) = HighByteOfLong(tileID_Oracle)
          End If
        End If
        pos = pos + 17
      Else
        pos = pos + 19
        If TibiaVersionLong >= 773 Then
          pos = pos + 1
        End If
      End If
      If TibiaVersionLong >= 853 Then '2
        pos = pos + 1
      End If
        If TibiaVersionLong >= 870 Then ' 1
           pos = pos + 2 ' fixed since 18.5
        End If
      If TibiaVersionLong >= 990 Then ' new 4 bytes
          pos = pos + 4
      End If
      If TibiaVersionLong >= 1036 Then ' new 1 byte
        pos = pos + 1
      End If
      'Debug.Print "direction5=" & GoodHex(packet(pos - 1)) & " " & GoodHex(packet(pos - 2)) & " " & GoodHex(packet(pos - 3)) & " " & GoodHex(packet(pos - 4)) & " " & GoodHex(packet(pos - 5))
      AddID_Direction idConnection, tempID, packet(pos - 1) 'update direction

    ElseIf idTile = &H63 Then
      ' new mobile, we already knew his ID
      tempID = FourBytesDouble(packet(pos + 2), packet(pos + 3), packet(pos + 4), packet(pos + 5))
      nameofgivenID = GetNameFromID(idConnection, tempID)
      Matrix(ny, nx, nz, idConnection).s(stackpos).t1 = &H61
      Matrix(ny, nx, nz, idConnection).s(stackpos).t2 = &H0
      Matrix(ny, nx, nz, idConnection).s(stackpos).t3 = &H0
      Matrix(ny, nx, nz, idConnection).s(stackpos).dblID = tempID
      ' AFTERLOGIN LOGOUT
      CheckIfGM idConnection, nameofgivenID, nz
      If RuneMakerOptions(idConnection).autoLogoutAnyFloor = True Then
        If frmRunemaker.IsFriend(LCase(nameofgivenID)) = False And nameofgivenID <> CharacterName(idConnection) Then
          AfterLoginLogoutReason(idConnection) = nameofgivenID
        End If
      ElseIf RuneMakerOptions(idConnection).autoLogoutCurrentFloor = True Then
        If nz = myZ(idConnection) Then
          If frmRunemaker.IsFriend(LCase(nameofgivenID)) = False And nameofgivenID <> CharacterName(idConnection) Then
            AfterLoginLogoutReason(idConnection) = nameofgivenID
          End If
        End If
      End If
      stackpos = stackpos + 1
      pos = pos + 7
      'Debug.Print "direction6=" & GoodHex(packet(pos - 1)) & " " & GoodHex(packet(pos - 2)) & " " & GoodHex(packet(pos - 3)) & " " & GoodHex(packet(pos - 4)) & " " & GoodHex(packet(pos - 5))
      AddID_Direction idConnection, tempID, packet(pos - 1) 'update direction
    Else
      'normal info
      'skip 2 or 3 bytes depending of tile info
      If DatTiles(idTile).haveExtraByte = True Then
            Matrix(ny, nx, nz, idConnection).s(stackpos).t1 = packet(pos)
            Matrix(ny, nx, nz, idConnection).s(stackpos).t2 = packet(pos + 1)
            Matrix(ny, nx, nz, idConnection).s(stackpos).t3 = packet(pos + 2)
            Matrix(ny, nx, nz, idConnection).s(stackpos).dblID = 0

        pos = pos + 3
      Else
     
            Matrix(ny, nx, nz, idConnection).s(stackpos).t1 = packet(pos)
            Matrix(ny, nx, nz, idConnection).s(stackpos).t2 = packet(pos + 1)
            Matrix(ny, nx, nz, idConnection).s(stackpos).t3 = &H0
            Matrix(ny, nx, nz, idConnection).s(stackpos).dblID = 0
     
        pos = pos + 2
      End If
      stackpos = stackpos + 1
    End If
  Wend

  ' we are at the end of the info for 1 position
  ' &H00 &H00 will mark the end for us
  For spaux = stackpos To 10
  Matrix(ny, nx, nz, idConnection).s(spaux).t1 = &H0
  Matrix(ny, nx, nz, idConnection).s(spaux).t2 = &H0
  Matrix(ny, nx, nz, idConnection).s(spaux).t3 = &H0
  Matrix(ny, nx, nz, idConnection).s(spaux).dblID = 0
  Next spaux
  res = pos
  
  'tmpName = ""
  'For spaux = DEBUGpos To (pos - 1)
  '  tmpName = tmpName & " " & GoodHex(packet(spaux))
  'Next spaux
  'Debug.Print tmpName
  ReadSinglePositionOld = res
  Exit Function
gotfail:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "ERROR WHILE READING MAP UPDATE (" & Err.Description & ")last good pos=" & pos & " packet=" & frmMain.showAsStr2(packet, 0)
  ReadSinglePositionOld = res
End Function




Public Function ReadSinglePosition(idConnection As Integer, nx As Long, _
 ny As Long, nz As Long, ByRef packet() As Byte, pos As Long) As Long
 ' update a single square of the map
  Dim stackpos As Long
  Dim idTile As Long
  Dim resF As TypePlayerInfo
  Dim nameofgivenID As String
  Dim tempID As Double
  Dim tmpName As String
  Dim res As Long
  Dim outfit As Long
  Dim spaux As Long
  Dim DEBUGpos As Long
  Dim CompleteDebugS As String
  Dim end_of_s As Boolean
  Dim dothedebug As Boolean
  #If FinalMode = 1 Then
  On Error GoTo gotfail
  #End If
  #If FinalMode = 0 Then
  CompleteDebugS = "STACK:"
  #End If
  res = -1
'  If TibiaVersion < 860 Then
'    ReadSinglePosition = ReadSinglePositionOld(idConnection, nx, ny, nz, packet, pos)
'  Exit Function
'  End If
  stackpos = 0
  DEBUGpos = pos
  dothedebug = True
  'If pos = 5107 Then
  'Debug.Print "warning at " & CStr(pos)
  'End If
  end_of_s = False
  Do
    If (packet(pos) = 0) And (packet(pos + 1) = 0) Then
        #If FinalMode = 0 Then
        CompleteDebugS = CompleteDebugS & " [00 00]"
        #End If
        If TibiaVersionLong < 872 Then
          LogOnFile "errors.txt", "WARNING: Unexpected tile 00 00 in version " & CStr(TibiaVersionLong) & " : " & frmMain.showAsStr(packet, True)
          ' errors.txt
        End If
        pos = pos + 2
    ElseIf packet(pos + 1) = &HFF Then
        #If FinalMode = 0 Then
        CompleteDebugS = CompleteDebugS & " [" & GoodHex(packet(pos)) & " FF]"
        #End If
        end_of_s = True
    Else
        idTile = GetTheLong(packet(pos), packet(pos + 1))
        If stackpos <= 10 Then
            Matrix(ny, nx, nz, idConnection).s(stackpos).t1 = packet(pos)
            Matrix(ny, nx, nz, idConnection).s(stackpos).t2 = packet(pos + 1)
            Matrix(ny, nx, nz, idConnection).s(stackpos).t3 = &H0
            Matrix(ny, nx, nz, idConnection).s(stackpos).dblID = 0
        Else
          dothedebug = True
        End If
        If (idTile > highestDatTile) Then
          #If FinalMode = 0 Then
            Debug.Print CompleteDebugS
            Debug.Print "ERROR: [" & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1)) & "]"
          #End If
          frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "ERROR WHILE READING MAP UPDATE ( TileID beyond limits : " & CStr(idTile) & ") last good pos=" & pos & " packet=" & frmMain.showAsStr2(packet, 0)
          ReadSinglePosition = res
          Exit Function
        End If
    
        
        Select Case idTile
        Case &H61
          
          'mobile info ; mobile = { player, monster, npc }
          'If packet(pos + 6) = &H1F Then
          'Debug.Print "STOP"
          'End If
          resF = GetTheMobileInfo(idConnection, packet, pos)
          If (resF.pos) - pos > 255 Then
          
            #If FinalMode = 0 Then
              CompleteDebugS = CompleteDebugS & " [overread at type 61]"
              Debug.Print CompleteDebugS
            #End If
            GoTo gotfail
          Else
            pos = resF.pos
          End If
    
          ' adding this mobile to the names string for this position of the matrix
          ' ReDim Matrix(-6 To 7, -8 To 9, 0 To 15, 1 To MAXCLIENTS
          If ((ny >= -6) And (ny <= 7) And (nx >= -8) And (nx <= 9) And (nz >= 0) And (nz <= 15) And (stackpos <= 10)) Then
            Matrix(ny, nx, nz, idConnection).s(stackpos).t1 = &H61
            Matrix(ny, nx, nz, idConnection).s(stackpos).t2 = &H0
            Matrix(ny, nx, nz, idConnection).s(stackpos).t3 = &H0
            Matrix(ny, nx, nz, idConnection).s(stackpos).t4 = &H0
            Matrix(ny, nx, nz, idConnection).s(stackpos).dblID = resF.newID
          End If
          #If FinalMode = 0 Then
          CompleteDebugS = CompleteDebugS & " [MOB61:" & GetNameFromID(idConnection, resF.newID) & "]"
          #End If

          
          'stackpos = stackpos + 1
          ' an optional autologout if danger feature ...
          If frmHardcoreCheats.chkLogoutIfDanger.Value = 1 And sentFirstPacket(idConnection) = False Then
            If resF.newID <> myID(idConnection) Then
              ' proxy will tell the reason (names) of the logout in this case
              If LogoutReason(idConnection) = "" Then
                LogoutReason(idConnection) = GetNameFromID(idConnection, resF.newID)
              Else
                LogoutReason(idConnection) = LogoutReason(idConnection) & " , " & GetNameFromID(idConnection, resF.newID)
              End If
            End If
          End If
          ' AFTERLOGIN LOGOUT
          tmpName = GetNameFromID(idConnection, resF.newID)
          CheckIfGM idConnection, tmpName, nz
          If RuneMakerOptions(idConnection).autoLogoutAnyFloor = True Then
            If frmRunemaker.IsFriend(LCase(tmpName)) = False And tmpName <> CharacterName(idConnection) Then
              AfterLoginLogoutReason(idConnection) = tmpName
            End If
          ElseIf RuneMakerOptions(idConnection).autoLogoutCurrentFloor = True Then
            If nz = myZ(idConnection) Then
              If frmRunemaker.IsFriend(LCase(tmpName)) = False And tmpName <> CharacterName(idConnection) Then
                AfterLoginLogoutReason(idConnection) = tmpName
              End If
            End If
          End If
        Case &H62
 
          ' we already knew his ID + include some info
          tempID = FourBytesDouble(packet(pos + 2), packet(pos + 3), packet(pos + 4), packet(pos + 5))
          AddID_HP idConnection, tempID, packet(pos + 6) 'update hp
          nameofgivenID = GetNameFromID(idConnection, tempID)
          ' ReDim Matrix(-6 To 7, -8 To 9, 0 To 15, 1 To MAXCLIENTS
          If ((ny >= -6) And (ny <= 7) And (nx >= -8) And (nx <= 9) And (nz >= 0) And (nz <= 15) And (stackpos <= 10)) Then
            Matrix(ny, nx, nz, idConnection).s(stackpos).t1 = &H61
            Matrix(ny, nx, nz, idConnection).s(stackpos).t2 = &H0
            Matrix(ny, nx, nz, idConnection).s(stackpos).t3 = &H0
            Matrix(ny, nx, nz, idConnection).s(stackpos).t4 = &H0
            Matrix(ny, nx, nz, idConnection).s(stackpos).dblID = tempID
          End If
          #If FinalMode = 0 Then
          CompleteDebugS = CompleteDebugS & " [MOB62:" & GetNameFromID(idConnection, tempID) & "]"
          #End If
          ' AFTERLOGIN LOGOUT
          CheckIfGM idConnection, nameofgivenID, nz
          If RuneMakerOptions(idConnection).autoLogoutAnyFloor = True Then
            If frmRunemaker.IsFriend(LCase(nameofgivenID)) = False And nameofgivenID <> CharacterName(idConnection) Then
              AfterLoginLogoutReason(idConnection) = nameofgivenID
            End If
          ElseIf RuneMakerOptions(idConnection).autoLogoutCurrentFloor = True Then
            If nz = myZ(idConnection) Then
              If frmRunemaker.IsFriend(LCase(nameofgivenID)) = False And nameofgivenID <> CharacterName(idConnection) Then
                AfterLoginLogoutReason(idConnection) = nameofgivenID
              End If
            End If
          End If
          'stackpos = stackpos + 1
          
          'eval outfit
          If TibiaVersionLong <= 760 Then
            outfit = CLng(packet(pos + 8))
          Else
            outfit = GetTheLong(packet(pos + 8), packet(pos + 9))
            pos = pos + 1
          End If
          
          If outfit = 0 Then
            If (packet(pos + 9) = &H0) And (packet(pos + 10) = &H0) Then
              If (nameofgivenID <> CharacterName(idConnection)) And (frmHardcoreCheats.chkReveal.Value = 1) Then
                packet(pos + 9) = LowByteOfLong(tileID_Oracle)
                packet(pos + 10) = HighByteOfLong(tileID_Oracle)
              End If
            End If
            pos = pos + 17
          Else
            pos = pos + 19
            If TibiaVersionLong >= 773 Then
              pos = pos + 1
            End If
          End If
          If TibiaVersionLong >= 853 Then '2
            pos = pos + 1
          End If
            If TibiaVersionLong >= 870 Then ' 1
               pos = pos + 2 ' fixed since 18.5
            End If
          If TibiaVersionLong >= 990 Then ' new 4 bytes
            pos = pos + 4
          End If
          If TibiaVersionLong >= 1036 Then ' new 1 byte
          pos = pos + 1
        End If
          'Debug.Print "direction5=" & GoodHex(packet(pos - 1)) & " " & GoodHex(packet(pos - 2)) & " " & GoodHex(packet(pos - 3)) & " " & GoodHex(packet(pos - 4)) & " " & GoodHex(packet(pos - 5))
          AddID_Direction idConnection, tempID, packet(pos - 1) 'update direction
    
        Case &H63
          CompleteDebugS = CompleteDebugS & " [MOB63]"
          ' new mobile, we already knew his ID
          tempID = FourBytesDouble(packet(pos + 2), packet(pos + 3), packet(pos + 4), packet(pos + 5))
          nameofgivenID = GetNameFromID(idConnection, tempID)
          If ((ny >= -6) And (ny <= 7) And (nx >= -8) And (nx <= 9) And (nz >= 0) And (nz <= 15) And (stackpos <= 10)) Then
          Matrix(ny, nx, nz, idConnection).s(stackpos).t1 = &H61
          Matrix(ny, nx, nz, idConnection).s(stackpos).t2 = &H0
          Matrix(ny, nx, nz, idConnection).s(stackpos).t3 = &H0
          Matrix(ny, nx, nz, idConnection).s(stackpos).t4 = &H0
          Matrix(ny, nx, nz, idConnection).s(stackpos).dblID = tempID
          End If
          #If FinalMode = 0 Then
          CompleteDebugS = CompleteDebugS & " [MOB63:" & GetNameFromID(idConnection, tempID) & "]"
          #End If
          ' AFTERLOGIN LOGOUT
          CheckIfGM idConnection, nameofgivenID, nz
          If RuneMakerOptions(idConnection).autoLogoutAnyFloor = True Then
            If frmRunemaker.IsFriend(LCase(nameofgivenID)) = False And nameofgivenID <> CharacterName(idConnection) Then
              AfterLoginLogoutReason(idConnection) = nameofgivenID
            End If
          ElseIf RuneMakerOptions(idConnection).autoLogoutCurrentFloor = True Then
            If nz = myZ(idConnection) Then
              If frmRunemaker.IsFriend(LCase(nameofgivenID)) = False And nameofgivenID <> CharacterName(idConnection) Then
                AfterLoginLogoutReason(idConnection) = nameofgivenID
              End If
            End If
          End If
          'stackpos = stackpos + 1
          pos = pos + 7
          'Debug.Print "direction6=" & GoodHex(packet(pos - 1)) & " " & GoodHex(packet(pos - 2)) & " " & GoodHex(packet(pos - 3)) & " " & GoodHex(packet(pos - 4)) & " " & GoodHex(packet(pos - 5))
          AddID_Direction idConnection, tempID, packet(pos - 1) 'update direction
          
          If TibiaVersionLong >= 950 Then
           pos = pos + 1 ' new strange byte since Tibia 9.5
          End If
        Case Else
          'normal info
          If TibiaVersionLong >= 990 Then
              'skip 3, 4 or 5 bytes depending of tile info
              If DatTiles(idTile).haveExtraByte = True Then
                Matrix(ny, nx, nz, idConnection).s(stackpos).t1 = packet(pos)
                Matrix(ny, nx, nz, idConnection).s(stackpos).t2 = packet(pos + 1)
                Matrix(ny, nx, nz, idConnection).s(stackpos).t3 = packet(pos + 3)
                Matrix(ny, nx, nz, idConnection).s(stackpos).dblID = 0
                If DatTiles(idTile).haveExtraByte2 = True Then
                    Matrix(ny, nx, nz, idConnection).s(stackpos).t4 = packet(pos + 4)
                    pos = pos + 5
                    #If FinalMode = 0 Then
                    CompleteDebugS = CompleteDebugS & " [" & GoodHex(packet(pos - 5)) & " " & GoodHex(packet(pos - 4)) & " " & GoodHex(packet(pos - 3)) & " " & GoodHex(packet(pos - 2)) & " " & GoodHex(packet(pos - 1)) & "]"
                    #End If
                Else
                    Matrix(ny, nx, nz, idConnection).s(stackpos).t4 = &H0
                    pos = pos + 4
                    #If FinalMode = 0 Then
                    CompleteDebugS = CompleteDebugS & " [" & GoodHex(packet(pos - 4)) & " " & GoodHex(packet(pos - 3)) & " " & GoodHex(packet(pos - 2)) & " " & GoodHex(packet(pos - 1)) & "]"
                    #End If
                End If
              Else
                Matrix(ny, nx, nz, idConnection).s(stackpos).t1 = packet(pos)
                Matrix(ny, nx, nz, idConnection).s(stackpos).t2 = packet(pos + 1)
                Matrix(ny, nx, nz, idConnection).s(stackpos).t3 = &H0
                Matrix(ny, nx, nz, idConnection).s(stackpos).dblID = 0
                If DatTiles(idTile).haveExtraByte2 = True Then
                  Matrix(ny, nx, nz, idConnection).s(stackpos).t4 = packet(pos + 3)
                  pos = pos + 4
                    #If FinalMode = 0 Then
                    CompleteDebugS = CompleteDebugS & " [" & GoodHex(packet(pos - 4)) & " " & GoodHex(packet(pos - 3)) & " " & GoodHex(packet(pos - 2)) & " " & GoodHex(packet(pos - 1)) & "]"
                    #End If
                Else
                    Matrix(ny, nx, nz, idConnection).s(stackpos).t4 = &H0
                    pos = pos + 3
                    #If FinalMode = 0 Then
                    CompleteDebugS = CompleteDebugS & " [" & GoodHex(packet(pos - 3)) & " " & GoodHex(packet(pos - 2)) & " " & GoodHex(packet(pos - 1)) & "]"
                    #End If
                End If
              End If
          
          Else ' older tibia
              'skip 2, 3 or 4 bytes depending of tile info
              If DatTiles(idTile).haveExtraByte = True Then
    
                Matrix(ny, nx, nz, idConnection).s(stackpos).t1 = packet(pos)
                Matrix(ny, nx, nz, idConnection).s(stackpos).t2 = packet(pos + 1)
                Matrix(ny, nx, nz, idConnection).s(stackpos).t3 = packet(pos + 2)
                Matrix(ny, nx, nz, idConnection).s(stackpos).t4 = &H0
                Matrix(ny, nx, nz, idConnection).s(stackpos).dblID = 0
                pos = pos + 3
                If DatTiles(idTile).haveExtraByte2 = True Then
                  pos = pos + 1 ' skip byte 00
                    #If FinalMode = 0 Then
                    CompleteDebugS = CompleteDebugS & " [" & GoodHex(packet(pos - 4)) & " " & GoodHex(packet(pos - 3)) & " " & GoodHex(packet(pos - 2)) & " " & GoodHex(packet(pos - 1)) & "]"
                    #End If
                Else
                    #If FinalMode = 0 Then
                    CompleteDebugS = CompleteDebugS & " [" & GoodHex(packet(pos - 3)) & " " & GoodHex(packet(pos - 2)) & " " & GoodHex(packet(pos - 1)) & "]"
                    #End If
                End If
              Else
                Matrix(ny, nx, nz, idConnection).s(stackpos).t1 = packet(pos)
                Matrix(ny, nx, nz, idConnection).s(stackpos).t2 = packet(pos + 1)
                Matrix(ny, nx, nz, idConnection).s(stackpos).t3 = &H0
                Matrix(ny, nx, nz, idConnection).s(stackpos).t4 = &H0
                Matrix(ny, nx, nz, idConnection).s(stackpos).dblID = 0
                pos = pos + 2
                If DatTiles(idTile).haveExtraByte2 = True Then
                  pos = pos + 1 ' skip byte 00
                    #If FinalMode = 0 Then
                    CompleteDebugS = CompleteDebugS & " [" & GoodHex(packet(pos - 3)) & " " & GoodHex(packet(pos - 2)) & " " & GoodHex(packet(pos - 1)) & "]"
                    #End If
                Else
                    #If FinalMode = 0 Then
                    CompleteDebugS = CompleteDebugS & " [" & GoodHex(packet(pos - 2)) & " " & GoodHex(packet(pos - 1)) & "]"
                    #End If
                End If
              End If
          End If
        End Select
        stackpos = stackpos + 1
    End If
  Loop Until end_of_s = True ' XX FF
  For spaux = stackpos To 10
  Matrix(ny, nx, nz, idConnection).s(spaux).t1 = &H0
  Matrix(ny, nx, nz, idConnection).s(spaux).t2 = &H0
  Matrix(ny, nx, nz, idConnection).s(spaux).t3 = &H0
  Matrix(ny, nx, nz, idConnection).s(spaux).t4 = &H0
  Matrix(ny, nx, nz, idConnection).s(spaux).dblID = 0
  Next spaux
 ' #If FinalMode = 0 Then
 ' If dothedebug = True Then
 '   Debug.Print CompleteDebugS
 ' End If
 ' #End If
  
  #If MapDebug = 1 Then
    AddwriteOnFileSimple "mapdebug.txt", CompleteDebugS
  #End If
  res = pos
  'Debug.Print "res=" & CStr(res)
  'If res = -1 Then
  ' Debug.Print "ERROR HERE"
  'End If
  ReadSinglePosition = res
  Exit Function
gotfail:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "ERROR WHILE READING MAP UPDATE (" & Err.Description & ")last good pos=" & pos & " packet=" & frmMain.showAsStr2(packet, 0)
  ReadSinglePosition = res
End Function
Public Function LearnFromPacket(ByRef packet() As Byte, pos As Long, idConnection As Integer) As TypeLearnResult
  ' read the info contained in a filtered packet
  Dim pType As Byte
  Dim lastpType As Byte
  Dim subType As Byte
  Dim initX As Long
  Dim initY As Long
  Dim initZ As Long
  Dim initS As Long
  Dim destX As Long
  Dim destY As Long
  Dim destZ As Long
  Dim matrixP As TypeMatrixPosition
  Dim matrixP2 As TypeMatrixPosition
  Dim mobName As String
  Dim resF As TypePlayerInfo
  Dim nameofgivenID As String
  Dim expectMore As Boolean
  Dim lonN As Long
  Dim finalAfterPos As Long
  Dim debug1 As String
  Dim debug2 As String
  Dim debugLon1 As Long
  Dim debugChain As String
  Dim showDebug As Boolean
  Dim gotStackP As Long
  Dim fRes As Long
  Dim myres As TypeLearnResult
  Dim itemCount As Long
  Dim tileID As Long
  Dim mobID As Double
  Dim floorDone As Boolean
  Dim numC As Long
  Dim resT As Long
  Dim debugFile As String
  Dim tempID As Double
  Dim startz As Long
  Dim endz As Long
  Dim zstep As Long
  Dim gotMapUpdate As Boolean
  Dim addSpecialCase As Boolean
  Dim tempb1 As Byte
  Dim tempb2 As Byte
  Dim tempb3 As Byte
  Dim tempb4 As Byte
  Dim templ1 As Long
  Dim templ2 As Long
  Dim outfitType As Long
  Dim strDebug As String
  Dim lonCap As Long
  Dim lonNumItems As Long
  Dim aRes As Long
  Dim lonO As Long
  Dim msg As String
  Dim rightpart As String
  Dim temps1 As String
  Dim corpseX As Long
  Dim corpseY As Long
  Dim corpseZ As Long
  Dim corpseS As Long
  Dim corpseTileID As Long
  Dim gotCorpsePop As Boolean
  Dim ignoreCorpsePop As Boolean
  Dim OTsType As Long
  Dim tmpID As Long
  Dim blnDebug1 As Boolean
  Dim blnTmp As Boolean

  Dim itsMe As Boolean
  Dim tmpStr As String
  Dim lastGoodPos As Long
  Dim debugChainType As String
  Dim oldHP As Long
  Dim debugReasons As String
  #If FinalMode Then
  On Error GoTo fatalError
  #End If
  debugChainType = ""
  debugReasons = ""
  ignoreCorpsePop = True
  gotCorpsePop = False
  addSpecialCase = False
  debugFile = "track.txt"
  OTsType = 0
  myres.fail = False
  myres.skipThis = False
  myres.firstMapDone = False
  myres.pos = 0
  myres.gotHPupdate = False
  myres.gotManaupdate = False
  myres.gotSoulupdate = False
 ' myres.gotBlankRune = False
  myres.gotNewCorpse = False
  showDebug = False
  debugChain = ""
  finalAfterPos = UBound(packet) + 1
  lastpType = &H0
  gotMapUpdate = False
  ' Loop until we read all the subpacket info contained inside.
  ' Example of correct debug chain of types received at login :
  ' 10.35: 0F 64 6A 83 78 78 78 78 78 82 8D 9F A2 92 B4 93 92 90 1E A0 A1
  
  ' 10.38: 0F 64 6A 83 78 78 78 78 78 78 78 82 8D 9F A2 92 D2 B4 9E 93 92 90 AC A0 A1
  Do
    lastGoodPos = pos
    mobName = ""
    expectMore = True
    pType = packet(pos)
    debugChain = debugChain & " " & GoodHex(pType)
    'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "! EVAL : " & GoodHex(pType)
    Select Case pType ' type of subpacket
    Case &HA
      ' my ID update ?
      If TibiaVersionLong >= 980 Then
        ' no extra info now. Just a popup of yourself (blue circle)
        pos = pos + 1
      Else
        If myID(idConnection) <> 0 Then
          aRes = SendSystemMessageToClient(idConnection, "Your ID have changed !")
          DoEvents
        End If
        IDstring(idConnection) = GoodHex(packet(pos + 1)) & GoodHex(packet(pos + 2)) & GoodHex(packet(pos + 3)) & GoodHex(packet(pos + 4))
        myID(idConnection) = FourBytesDouble(packet(pos + 1), packet(pos + 2), packet(pos + 3), packet(pos + 4))
        pos = pos + 5
      End If
    Case &HB
      ' special gm powers array
      If TibiaVersionLong >= 850 Then
        pos = pos + 21 ' fixed in 15.6
      ElseIf TibiaVersionLong >= 811 Then
        pos = pos + 24
      Else
        pos = pos + 33
      End If
    Case &HF
      ' finish Pending Status - Tibia 9.8
      CheatsPaused(idConnection) = False
      pos = pos + 1
    Case &H14
      ' server error
      If ReconnectionStage(idConnection) > 0 Then
        myres.skipThis = True
      End If
      expectMore = False
    Case &H15
      ' premium account activated message
      lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
      pos = pos + 3 + lonN ' fixed since 8.42
    Case &H16
      ' on login queue
      expectMore = False
    Case &H17
      ' new since Tibia 9.8 - new pending state
   
      CheatsPaused(idConnection) = True
      IDstring(idConnection) = GoodHex(packet(pos + 1)) & GoodHex(packet(pos + 2)) & GoodHex(packet(pos + 3)) & GoodHex(packet(pos + 4))
      myID(idConnection) = FourBytesDouble(packet(pos + 1), packet(pos + 2), packet(pos + 3), packet(pos + 4))
      If TibiaVersionLong >= 1080 Then
      ' tibia 10.80+
      ' 17 FA D5 7B 02 32 00 03 0F 15 0D 80 03 A9 FC 03 80 03 7E D5 B6 7F 00 01 01 24 00 68 74 74 70 3A 2F 2F 73 74 61 74 69 63 2E 74 69 62 69 61 2E 63 6F 6D 2F 69 6D 61 67 65 73 2F 73 74 6F 72 65 19 00 0A
        pos = pos + 25
        lonN = GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 5 + lonN
      ElseIf TibiaVersionLong >= 1058 Then
      ' tibia 10.58+
      ' 17 FA D5 7B 02 32 00 03 0F 15 0D 80 03 A9 FC 03 80 03 7E D5 B6 7F 00 01 01 0A
         pos = pos + 26
      Else
       pos = pos + 24
      End If
    Case &H1D ' tibia 9.5
      ' server ping ??
      If publicDebugMode = True Then
        aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Server sent ping 1D")
        DoEvents
      End If
      pos = pos + 1
    Case &H1E
      ' server ping (confirmed)
      If publicDebugMode = True Then
        aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Server sent ping 1E")
        DoEvents
      End If
      pos = pos + 1
    Case &H28
      'you are dead
       pos = pos + 1
       If TibiaVersionLong >= 870 Then
        pos = pos + 1 '  1 new byte since Tibia 8.7
       End If
       If TibiaVersionLong >= 1076 Then ' maybe before 10.76 too - unconfirmed
         pos = pos + 1 '  1 new byte since Tibia 10.76
       End If
       If (cavebotEnabled(idConnection) = True) Or (RuneMakerOptions(idConnection).activated = True) Or (RuneMakerOptions(idConnection).autoEat = True) Then
         AfterLoginLogoutReason(idConnection) = "YOU DIED!" & vbLf & "Auto reconnection canceled to avoid potential disasters."
       End If
    Case &H2E
      ' Unknown command
      ' UNDER INVESTIGATION - AUTOBAN
      nameofgivenID = "[WARNING!] Request to search cheats? Please report this log with timestamp to daniel@blackdtools.com .Strange packet received: " & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1))
      LogOnFile "errors.txt", nameofgivenID
      aRes = SendLogSystemMessageToClient(idConnection, nameofgivenID)
      DoEvents
      pos = pos + 2
    Case &H32
      ' level of gm powers
      ' 00 00 = none
      ' 00 01 = tutor
      ' ?? ?? = gm
      pos = pos + 3
    Case &H33
      ' unknown, but fixed lenght
      ' UNDER INVESTIGATION - AUTOBAN
      nameofgivenID = "[WARNING!] Request to search cheats? Please report this log with timestamp to daniel@blackdtools.com .Strange packet received: " & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1))
      LogOnFile "errors.txt", nameofgivenID
      aRes = SendLogSystemMessageToClient(idConnection, nameofgivenID)
      DoEvents
      pos = pos + 2
    Case &H59
      ' unknown, but fixed lenght
      ' UNDER INVESTIGATION - AUTOBAN
      nameofgivenID = "[WARNING!] Request to search cheats? Please report this log with timestamp to daniel@blackdtools.com .Strange packet received: " & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1))
      LogOnFile "errors.txt", nameofgivenID
      aRes = SendLogSystemMessageToClient(idConnection, nameofgivenID)
      DoEvents
      pos = pos + 2
    Case &H64
      ' all the map (first packet)
      pos = ReadMap(idConnection, packet, pos)
      If sentFirstPacket(idConnection) = False Then
        myres.firstMapDone = True
        If frmHardcoreCheats.chkLockOnMyFloor.Value = 1 Then
          mapFloorSelected = myZ(mapIDselected)
        End If
      End If
      lastFloorChangeX(idConnection) = myX(idConnection)
      lastFloorChangeY(idConnection) = myY(idConnection)
      lastFloorChangeZ(idConnection) = myZ(idConnection)
      gotMapUpdate = True
    Case &H65
      ' north update
      myY(idConnection) = myY(idConnection) - 1
      EvalMyMove idConnection, 0, -1, 0
      pos = UpdateNorthSide(idConnection, packet, pos)
      gotMapUpdate = True
    Case &H66
     ' right update
      myX(idConnection) = myX(idConnection) + 1
      EvalMyMove idConnection, 1, 0, 0
      pos = UpdateRightSide(idConnection, packet, pos)
      gotMapUpdate = True
    Case &H67
     ' south update
      myY(idConnection) = myY(idConnection) + 1
      EvalMyMove idConnection, 0, 1, 0
      pos = UpdateSouthSide(idConnection, packet, pos)
      gotMapUpdate = True
    Case &H68
     ' left update
      myX(idConnection) = myX(idConnection) - 1
      EvalMyMove idConnection, -1, 0, 0
      pos = UpdateLeftSide(idConnection, packet, pos)
      gotMapUpdate = True
    Case &H69
      ' full update of a single position
      initX = GetTheLong(packet(pos + 1), packet(pos + 2))
      initY = GetTheLong(packet(pos + 3), packet(pos + 4))
      initZ = CLng(packet(pos + 5))
      matrixP = GetMatrixPosition(idConnection, initX, initY, initZ, 0)
      pos = pos + 6
      If packet(pos + 1) = &HFF Then 'first we could have a skipper (&H01 &HFF)
        pos = pos + 2
        Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(0).t1 = 0
        Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(0).t2 = 0
        Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(0).t3 = 0
        Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(0).dblID = 0
      Else 'else we have info about map position
         resT = ReadSinglePosition(idConnection, matrixP.X, matrixP.y, matrixP.z, packet, pos)
         If resT = -1 Then
           GoTo fatalError
         Else
           pos = resT
         End If

         ' we have packet(pos)=number and packet(pos+1)=&HFF at this point
         pos = pos + 2
      End If
    Case &H6A
      ' something get in screen
      ' tibia ????  0D 00 6A 1F 00 1F 00 07 83 1F 00 1F 00 07 0B
      ' tibia 9.5   0F 00 6A 5E 7E E0 7D 09 FF 63 00 FB F5 00 40 02 01
      ' tibia 9.5         6A 58 7E E2 7D 07 FF 61 00 00 00 00 00 F7 E1 5C 02 00 07 00 50 72 69 7A 6F 6B 61 64 02 89 00 72 4C 4C 55 00 00 00 00 00 00 01 00 00 00 00

      'inix=31
      'iniy=31
      'iniz=7
      'tileid=
      initX = GetTheLong(packet(pos + 1), packet(pos + 2))
      initY = GetTheLong(packet(pos + 3), packet(pos + 4))
      initZ = CLng(packet(pos + 5))
      initS = 0
      matrixP = GetMatrixPosition(idConnection, initX, initY, initZ, initS)
      pos = pos + 6
      templ2 = 255
      If (TibiaVersionLong >= 841) Then
         templ2 = CLng(packet(pos))
         pos = pos + 1 ' 1 new byte, &HFF , &H02 , level??
      End If

      tileID = GetTheLong(packet(pos), packet(pos + 1))
      

         
      Select Case tileID
      Case &H63
        ' new mobile, we already knew his ID
        tempID = FourBytesDouble(packet(pos + 2), packet(pos + 3), packet(pos + 4), packet(pos + 5))
        ' AFTERLOGIN LOGOUT
        nameofgivenID = GetNameFromID(idConnection, tempID)
        
       ' frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "New char1 (" & nameofgivenID & ") with level " & CStr(templ2)
        
        CheckIfGM idConnection, nameofgivenID, initZ
        If RuneMakerOptions(idConnection).autoLogoutAnyFloor = True Then
          If nameofgivenID <> "" And frmRunemaker.IsFriend(LCase(nameofgivenID)) = False And nameofgivenID <> CharacterName(idConnection) Then
            AfterLoginLogoutReason(idConnection) = nameofgivenID
          End If
        ElseIf RuneMakerOptions(idConnection).autoLogoutCurrentFloor = True Then
          If nameofgivenID <> "" And matrixP.z = myZ(idConnection) Then
            If frmRunemaker.IsFriend(LCase(nameofgivenID)) = False And nameofgivenID <> CharacterName(idConnection) Then
              AfterLoginLogoutReason(idConnection) = nameofgivenID
            End If
          End If
        End If
        gotStackP = AddThingToStack(idConnection, matrixP.X, matrixP.y, matrixP.z, &H61, &H0, &H0, tempID)
        pos = pos + 7
        'Debug.Print "direction7=" & GoodHex(packet(pos - 1)) & " " & GoodHex(packet(pos - 2)) & " " & GoodHex(packet(pos - 3)) & " " & GoodHex(packet(pos - 4)) & " " & GoodHex(packet(pos - 5))
        AddID_Direction idConnection, tempID, packet(pos - 1) 'update direction
        
      If (TibiaVersionLong >= 950) Then
         pos = pos + 1 ' 1 new byte since Tibia 9.5
      End If
      Case &H62
  ' tibia 9.9 : 6A 2A 7D 9E 7D 0B FF 62 00 0D FA C1 02 4F 01 80 00 4E 45 3A 4C 00 00 00 00 00 70 00 00 00 00 FF 00 00 00
  ' tibia 10.36:6A C4 81 C3 7E 07 FF 62 00 3A A7 3B 02 62 00 86 00 09 00 72 5D 03 00 00 06 1D EE 00 00 00 00 00 FF 00 00 00

        tempID = FourBytesDouble(packet(pos + 2), packet(pos + 3), packet(pos + 4), packet(pos + 5))
        AddID_HP idConnection, tempID, packet(pos + 6) 'update hp
        ' AFTERLOGIN LOGOUT
        nameofgivenID = GetNameFromID(idConnection, tempID)
        
       ' frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "New char1 (" & nameofgivenID & ") with level " & CStr(templ2)
                
        CheckIfGM idConnection, nameofgivenID, initZ
        If RuneMakerOptions(idConnection).autoLogoutAnyFloor = True Then
          If nameofgivenID <> "" And frmRunemaker.IsFriend(LCase(nameofgivenID)) = False And nameofgivenID <> CharacterName(idConnection) Then
            AfterLoginLogoutReason(idConnection) = nameofgivenID
          End If
        ElseIf RuneMakerOptions(idConnection).autoLogoutCurrentFloor = True Then
          If nameofgivenID <> "" And matrixP.z = myZ(idConnection) Then
            If frmRunemaker.IsFriend(LCase(nameofgivenID)) = False And nameofgivenID <> CharacterName(idConnection) Then
              AfterLoginLogoutReason(idConnection) = nameofgivenID
            End If
          End If
        End If
        gotStackP = AddThingToStack(idConnection, matrixP.X, matrixP.y, matrixP.z, &H61, &H0, &H0, tempID)
        
        ' outfit ...
        If TibiaVersionLong <= 760 Then
          templ1 = CLng(packet(pos + 8))
        Else
          templ1 = GetTheLong(packet(pos + 8), packet(pos + 9))
          pos = pos + 1
        End If
        'XXYY
        
        If templ1 = 0 Then
          If (packet(pos + 9) = &H0) And (packet(pos + 10) = &H0) Then
            If (tempID <> myID(idConnection)) And (frmHardcoreCheats.chkReveal.Value = 1) Then
              packet(pos + 9) = LowByteOfLong(tileID_Oracle)
              packet(pos + 10) = HighByteOfLong(tileID_Oracle)
            End If
          End If
          pos = pos + 17
        Else
          pos = pos + 19
          If TibiaVersionLong >= 773 Then
            pos = pos + 1
          End If
        End If
        If TibiaVersionLong >= 853 Then ' 3
          pos = pos + 1
        End If
  'If TibiaVersionLong >= 854 Then ' 1
  '  pos = pos + 1 ' skip one more
 ' End If
        'Debug.Print "direction1=" & GoodHex(packet(pos - 1)) & " " & GoodHex(packet(pos - 2)) & " " & GoodHex(packet(pos - 3)) & " " & GoodHex(packet(pos - 4)) & " " & GoodHex(packet(pos - 5))
        
        If TibiaVersionLong >= 870 Then ' xxx1
          pos = pos + 2
        End If
        If TibiaVersionLong >= 990 Then ' new 4 bytes
          pos = pos + 4
        End If
        If TibiaVersionLong >= 1036 Then ' new 1 byte
          pos = pos + 1
        End If
        AddID_Direction idConnection, tempID, packet(pos - 1) 'update dir
      Case &H61
        ' new character info
        resF = GetTheMobileInfo(idConnection, packet, pos)
        ' AFTERLOGIN LOGOUT
       nameofgivenID = GetNameFromID(idConnection, resF.newID)

        CheckIfGM idConnection, nameofgivenID, initZ
        
     'If templ2 < 255 Then
      '  tempb3 = 0
       ' frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "new char (" & nameofgivenID & ") with level " & CStr(templ2) & ", tileid " & CStr(tileID)
     ' End If
        
        
        If RuneMakerOptions(idConnection).autoLogoutAnyFloor = True Then
          If nameofgivenID <> "" And frmRunemaker.IsFriend(LCase(nameofgivenID)) = False And nameofgivenID <> CharacterName(idConnection) Then
            AfterLoginLogoutReason(idConnection) = nameofgivenID
          End If
        ElseIf RuneMakerOptions(idConnection).autoLogoutCurrentFloor = True Then
          If nameofgivenID <> "" And matrixP.z = myZ(idConnection) Then
            If frmRunemaker.IsFriend(LCase(nameofgivenID)) = False And nameofgivenID <> CharacterName(idConnection) Then
              AfterLoginLogoutReason(idConnection) = nameofgivenID
            End If
          End If
        End If
        pos = resF.pos
        'If TibiaVersionLong >= 870 Then
        '    pos = pos + 2 ' 2 strange bytes 00 01
        'End If
        gotStackP = AddThingToStack(idConnection, matrixP.X, matrixP.y, matrixP.z, &H61, &H0, &H0, resF.newID)
      Case Else
        If TibiaVersionLong >= 990 Then
            If DatTiles(tileID).haveExtraByte = True Then
              If DatTiles(tileID).haveExtraByte2 = True Then
                gotStackP = AddThingToStack(idConnection, matrixP.X, matrixP.y, matrixP.z, packet(pos), packet(pos + 1), packet(pos + 3), &H0, packet(pos + 4))
                pos = pos + 5
              Else
                gotStackP = AddThingToStack(idConnection, matrixP.X, matrixP.y, matrixP.z, packet(pos), packet(pos + 1), packet(pos + 3), &H0)
                pos = pos + 4
              End If
            Else
              
              If DatTiles(tileID).haveExtraByte2 = True Then
                gotStackP = AddThingToStack(idConnection, matrixP.X, matrixP.y, matrixP.z, packet(pos), packet(pos + 1), &H0, &H0, packet(pos + 3))
                pos = pos + 4
              Else
                gotStackP = AddThingToStack(idConnection, matrixP.X, matrixP.y, matrixP.z, packet(pos), packet(pos + 1), &H0, &H0)
                pos = pos + 3
              End If
            End If
        Else
            If DatTiles(tileID).haveExtraByte = True Then
              gotStackP = AddThingToStack(idConnection, matrixP.X, matrixP.y, matrixP.z, packet(pos), packet(pos + 1), packet(pos + 2), 0)
              pos = pos + 3
              If DatTiles(tileID).haveExtraByte2 = True Then
                pos = pos + 1 ' skip new strange byte 00
              End If
            Else
              gotStackP = AddThingToStack(idConnection, matrixP.X, matrixP.y, matrixP.z, packet(pos), packet(pos + 1), &H0, 0)
              pos = pos + 2
              If DatTiles(tileID).haveExtraByte2 = True Then
                pos = pos + 1 ' skip new strange byte 00
              End If
            End If
        End If
          ' for cavebot - autolooter
        If ((autoLoot(idConnection) = True) And (DatTiles(tileID).iscontainer = True)) Then
            ' check if it is something that died near and in the same floor
            If (initX < myX(idConnection) + AllowedLootDistance(idConnection)) And (initX > myX(idConnection) - AllowedLootDistance(idConnection)) And _
             (initY < myY(idConnection) + AllowedLootDistance(idConnection)) And (initY > myY(idConnection) - AllowedLootDistance(idConnection)) And _
             (initZ = myZ(idConnection)) Then
              ignoreCorpsePop = StairsExistsAt(initX, initY, initZ, idConnection)
'              If ignoreCorpsePop = True Then
'                aRes = GiveGMmessage(idConnection, "(dbg1) Stair exists at " & CStr(initX) & "," & CStr(initY) & "," & CStr(initZ), "Debug")
'                DoEvents
'              End If
            Else
              ignoreCorpsePop = True
            End If
            corpseX = initX
            corpseY = initY
            corpseZ = initZ
            corpseS = gotStackP
            corpseTileID = tileID
            ' B
            If autoLoot(idConnection) = True Then
                If gotCorpsePop = True Then
                    ' required for ot servers, change since blackd prox 15.3
                      matrixP = GetMatrixPosition(idConnection, initX, initY, initZ, gotStackP)
                      If ignoreCorpsePop = False Then ' new since blackd 18.9
                        myLastCorpseX(idConnection) = corpseX
                        myLastCorpseY(idConnection) = corpseY
                        myLastCorpseZ(idConnection) = corpseZ
                        myLastCorpseS(idConnection) = corpseS
                        myLastCorpseTileID(idConnection) = tileID
                        If OldLootMode(idConnection) = True Then
                          myres.gotNewCorpse = True
                        Else
                          templ1 = AddLootPoint(idConnection, corpseX, corpseY, corpseZ)
                          'aRes = GiveGMmessage(idConnection, PrintLootStats(idConnection), "BlackdProxy")
                        End If
                      End If
                      If GotKillOrderTargetID(idConnection) <> 0 Then
                        If (Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).dblID = GotKillOrderTargetID(idConnection)) Then
                          GotKillOrder(idConnection) = False
                          aRes = GiveGMmessage(idConnection, ">> " & GotKillOrderTargetName(idConnection) & " <<", "Target is DEAD")
                          DoEvents
                          GotKillOrderTargetID(idConnection) = 0
                          GotKillOrderTargetName(idConnection) = ""
                        End If
                      End If
                Else
                    gotCorpsePop = True
                End If
            End If
        End If
      End Select
      gotMapUpdate = True
    Case &H6B
      ' tibia 9.5 1B 00 6B 56 7E E4 7D 07 01 63 00 48 01 00 40 01 01
      ' tibia 9.9       6B 43 7D C0 7D 07 02 47 0B FF 05
      ' something turns / updates
      initX = GetTheLong(packet(pos + 1), packet(pos + 2))
      initY = GetTheLong(packet(pos + 3), packet(pos + 4))
      initZ = CLng(packet(pos + 5))
      initS = CLng(packet(pos + 6))
      matrixP = GetMatrixPosition(idConnection, initX, initY, initZ, initS)
      pos = pos + 7
      tileID = GetTheLong(packet(pos), packet(pos + 1))
      Select Case tileID
      Case &H63
        ' update a stack position
        tempID = FourBytesDouble(packet(pos + 2), packet(pos + 3), packet(pos + 4), packet(pos + 5))
        If matrixP.valid = True Then
          Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t1 = &H61
          Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t2 = &H0
          Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t3 = &H0
          Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t4 = &H0
          Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).dblID = tempID
        End If
        pos = pos + 7
        'Debug.Print "direction2=" & GoodHex(packet(pos - 1)) & " " & GoodHex(packet(pos - 2)) & " " & GoodHex(packet(pos - 3)) & " " & GoodHex(packet(pos - 4)) & " " & GoodHex(packet(pos - 5))
        AddID_Direction idConnection, tempID, packet(pos - 1) 'update direction
        
      If (TibiaVersionLong >= 950) Then
         'templ2 = CLng(packet(pos))
         pos = pos + 1 ' 1 new byte since Tibia 9.5
      End If
      
      Case &H62
        frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "WARNING (62) - please report to blackd :" & frmMain.showAsStr2(packet, 0)
        tempID = FourBytesDouble(packet(pos + 2), packet(pos + 3), packet(pos + 4), packet(pos + 5))
        AddID_HP idConnection, tempID, packet(pos + 6) 'update hp
        ' outfit
        If TibiaVersionLong <= 760 Then
          templ1 = packet(pos + 8)
        Else
          templ1 = GetTheLong(packet(pos + 8), packet(pos + 9))
          pos = pos + 1
        End If
        If templ1 = 0 Then
          If (packet(pos + 9) = &H0) And (packet(pos + 10) = &H0) Then
            If (tempID <> myID(idConnection)) And (frmHardcoreCheats.chkReveal.Value = 1) Then
              packet(pos + 9) = LowByteOfLong(tileID_Oracle)
              packet(pos + 10) = HighByteOfLong(tileID_Oracle)
            End If
          End If
          pos = pos + 17
        Else
          pos = pos + 19
          If TibiaVersionLong >= 773 Then
            pos = pos + 1
          End If
        End If
        If TibiaVersionLong >= 990 Then ' new 4 bytes
          pos = pos + 4
        End If
        If TibiaVersionLong >= 1036 Then ' new 1 byte
          pos = pos + 1
        End If
        'Debug.Print "direction3=" & GoodHex(packet(pos - 1)) & " " & GoodHex(packet(pos - 2)) & " " & GoodHex(packet(pos - 3)) & " " & GoodHex(packet(pos - 4)) & " " & GoodHex(packet(pos - 5))
        AddID_Direction idConnection, tempID, packet(pos - 1) 'update direction
      Case &H61
        frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "WARNING (61) - plz report to blackd :" & frmMain.showAsStr2(packet, 0)
        resF = GetTheMobileInfo(idConnection, packet, pos)
        pos = resF.pos
      Case Else
        If (TibiaVersionLong >= 990) Then
            If (DatTiles(tileID).haveExtraByte = True) Then
              If (DatTiles(tileID).haveExtraByte2 = True) Then
                If matrixP.valid = True Then
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t1 = packet(pos)
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t2 = packet(pos + 1)
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t3 = packet(pos + 3)
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t4 = packet(pos + 4)
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).dblID = 0
                End If
                pos = pos + 5
              Else
                If matrixP.valid = True Then
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t1 = packet(pos)
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t2 = packet(pos + 1)
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t3 = packet(pos + 3)
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t4 = &H0
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).dblID = 0
                End If
                pos = pos + 4
              End If
            Else
              If (DatTiles(tileID).haveExtraByte2 = True) Then
                If matrixP.valid = True Then
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t1 = packet(pos)
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t2 = packet(pos + 1)
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t3 = &H0
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t4 = packet(pos + 3)
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).dblID = 0
                End If
                pos = pos + 4
              Else
                If matrixP.valid = True Then
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t1 = packet(pos)
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t2 = packet(pos + 1)
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t3 = &H0
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t4 = &H0
                  Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).dblID = 0
                End If
                pos = pos + 3
              End If
            End If
        
        Else
        
            If DatTiles(tileID).haveExtraByte Then
              If matrixP.valid = True Then
                Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t1 = packet(pos)
                Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t2 = packet(pos + 1)
                Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t3 = packet(pos + 2)
                Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).dblID = 0
              End If
              pos = pos + 3
            Else
              If matrixP.valid = True Then
                Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t1 = packet(pos)
                Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t2 = packet(pos + 1)
                Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t3 = &H0
                Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).dblID = 0
              End If
              pos = pos + 2
            End If
            If DatTiles(tileID).haveExtraByte2 Then
              pos = pos + 1 ' skip new byte 00
            End If
        End If
      End Select
      gotMapUpdate = True
    Case &H6C
      ' removes something from stack
      initX = GetTheLong(packet(pos + 1), packet(pos + 2))
      initY = GetTheLong(packet(pos + 3), packet(pos + 4))
      initZ = CLng(packet(pos + 5))
      initS = CLng(packet(pos + 6))
      
      matrixP = GetMatrixPosition(idConnection, initX, initY, initZ, initS)
      ' for cavebot - autolooter
      ' A
      If autoLoot(idConnection) = True Then
        If ((gotCorpsePop = True) And (myres.gotNewCorpse = False)) Then
          If ((initX = corpseX) And (initY = corpseY) And (initZ = corpseZ)) Then
          
            ignoreCorpsePop = StairsExistsAt(initX, initY, initZ, idConnection)
'            If ignoreCorpsePop = True Then
'                aRes = GiveGMmessage(idConnection, "(dbg1) Stair exists at " & CStr(initX) & "," & CStr(initY) & "," & CStr(initZ), "Debug")
'                DoEvents
'            End If
            If ignoreCorpsePop = False Then
                myLastCorpseX(idConnection) = corpseX
                myLastCorpseY(idConnection) = corpseY
                myLastCorpseZ(idConnection) = corpseZ
                myLastCorpseS(idConnection) = corpseS
                myLastCorpseTileID(idConnection) = tileID
                If OldLootMode(idConnection) = True Then
                  myres.gotNewCorpse = True
                Else
                  templ1 = AddLootPoint(idConnection, corpseX, corpseY, corpseZ)
                  'aRes = GiveGMmessage(idConnection, PrintLootStats(idConnection), "BlackdProxy")
                End If
            End If
            If GotKillOrderTargetID(idConnection) <> 0 Then
                If (Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).dblID = GotKillOrderTargetID(idConnection)) Then
                  GotKillOrder(idConnection) = False
                  aRes = GiveGMmessage(idConnection, ">> " & GotKillOrderTargetName(idConnection) & " <<", "Target is DEAD")
                  DoEvents
                  GotKillOrderTargetID(idConnection) = 0
                  GotKillOrderTargetName(idConnection) = ""
                End If
            End If
          End If
        End If
      End If
      
      If gotCorpsePop = False Then
        gotCorpsePop = True ' for ot servers
      End If
      fRes = RemoveThingFromStack(idConnection, matrixP.X, matrixP.y, matrixP.z, matrixP.s)
      If fRes = -1 Then
        showDebug = True
        debugReasons = debugReasons & vbCrLf & " [ FAIL to remove mobile from " & _
         matrixP.X & "," & matrixP.y & "," & matrixP.z & "," & matrixP.s & " ] "
      End If
      pos = pos + 7
      gotMapUpdate = True
    Case &H6D
      ' something moves
      initX = GetTheLong(packet(pos + 1), packet(pos + 2))
      initY = GetTheLong(packet(pos + 3), packet(pos + 4))
      initZ = CLng(packet(pos + 5))
      initS = CLng(packet(pos + 6))
      matrixP = GetMatrixPosition(idConnection, initX, initY, initZ, initS)
      If matrixP.valid = True Then
        tempb1 = Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t1
        tempb2 = Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t2
        tempb3 = Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t3
        tempb4 = Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).t4
        tempID = Matrix(matrixP.y, matrixP.X, matrixP.z, idConnection).s(matrixP.s).dblID
      ElseIf ((packet(pos + 1) = &HFF) And (packet(pos + 2) = &HFF)) Then ' id from nowhere ' blackd 19.4
        tempb1 = &H61
        tempb2 = &H0
        tempb3 = &H0
        tempb4 = &H0
        tempID = FourBytesDouble(packet(pos + 3), packet(pos + 4), packet(pos + 5), packet(pos + 6))
      Else ' ??
        tempb1 = &H0
        tempb2 = &H0
        tempb3 = &H0
        tempb4 = &H0
        tempID = 0
      End If
      destX = GetTheLong(packet(pos + 7), packet(pos + 8))
      destY = GetTheLong(packet(pos + 9), packet(pos + 10))
      destZ = CLng(packet(pos + 11))
      If matrixP.valid = True Then
        If (destX - initX) = 0 Then
            If destY > initY Then
                AddID_Direction idConnection, tempID, &H2 'update direction
            ElseIf destY < initY Then
                AddID_Direction idConnection, tempID, &H0 'update direction
            End If
        ElseIf destX > initX Then
            AddID_Direction idConnection, tempID, &H1 'update direction
        ElseIf destX < initX Then
            AddID_Direction idConnection, tempID, &H3 'update direction
        End If
      End If
      matrixP2 = GetMatrixPosition(idConnection, destX, destY, destZ, 1)
      If matrixP.valid = True Then
        fRes = RemoveThingFromStack(idConnection, matrixP.X, matrixP.y, matrixP.z, matrixP.s)
        If fRes = -1 Then
          'it was out of matrix so ignore and give error
          debugReasons = debugReasons & vbCrLf & " [ ERROR at RemoveThingFromStack while trying to move from " & _
           matrixP.X & "," & matrixP.y & "," & matrixP.z & "," & matrixP.s & " ] "
          myres.fail = True
          showDebug = True
        End If
      End If
      ' AFTERLOGIN LOGOUT
      If tempID = 0 Then
        mobName = ""
      Else
        mobName = GetNameFromID(idConnection, tempID)
        CheckIfGM idConnection, mobName, destZ
      End If
      If RuneMakerOptions(idConnection).autoLogoutAnyFloor = True Then
        If mobName <> "" And frmRunemaker.IsFriend(LCase(mobName)) = False And mobName <> CharacterName(idConnection) Then
          AfterLoginLogoutReason(idConnection) = mobName
        End If
      ElseIf RuneMakerOptions(idConnection).autoLogoutCurrentFloor = True Then
        If mobName <> "" And matrixP2.z = myZ(idConnection) Then
          If frmRunemaker.IsFriend(LCase(mobName)) = False And mobName <> CharacterName(idConnection) Then
            AfterLoginLogoutReason(idConnection) = mobName
          End If
        End If
      End If
      ' insert part
      If matrixP2.valid = True Then
         gotStackP = AddThingToStack(idConnection, matrixP2.X, matrixP2.y, matrixP2.z, tempb1, tempb2, tempb3, tempID, tempb4)
      Else
        ' should not happen
      End If
      pos = pos + 12
      gotMapUpdate = True
    Case &H6E
      ' open / update container
      debugChainType = " ---- packet 6E DEBUG: "
      ' DEBUG=21 00 6E 00 26 0B 08 00 62 61 63 6B 70 61 63 6B 32 00 06 26 0B 26 0B DB 0B 27 D7 0B 1E FE 0D 0A F9 0D 07
      ' DEBUG=35 00 6E 00 B5 0C 13 00 62 61 63 6B 70 61 63 6B 20 6F 66 20 68 6F 6C 64 69 6E 67 32 00 0B 07 0C 16 0D FE 0D 0A 26 0B B1 0C DB 0B 19 E3 0B 2A 81 0D F1 0B 38 0B 26 0B
      
      
' problems872 [03/05/2011 00:47:56 using version 20.0 , with config.ini v872 ] FATAL ERROR [False ] DEBUG CHAIN [  6E SP: ---- packet 6E DEBUG: 1,2,3,4,5,6, ] CLIENT [ 1 ] (Number:9 Description:Subscript out of range Source: Notepad) in packet : ( hex ) 28 00
   ' debug872= 6E 02 CA 1F 00 1B 00 74 68 65 20 72 65 6D 61 69 6E 73 20 6F 66 20 61 6E 20 65 6C 65 6D 65 6E 74 61 6C 0E 00 01 D7 0B 41  Last good position=2 x=32390;y=32768;z=3
   ' debug872= 6E 02 CA 1F 00 1B 00 74 68 65 20 72 65 6D 61 69 6E 73 20 6F 66 20 61 6E 20 65 6C 65 6D 65 6E 74 61 6C 0E 00 03 F9 02 07 00 D7 0B 15 D7 0C 02  Last good position=2 x=32409;y=32736;z=3
   ' debug872= 6E 02 CA 1F 00 1B 00 74 68 65 20 72 65 6D 61 69 6E 73 20 6F 66 20 61 6E 20 65 6C 65 6D 65 6E 74 61 6C 0E 00 01 D7 0B 25  Last good position=2 x=32398;y=32726;z=3
      ' debug872= 6E 00 4C 17 08 00 64 65 61 64 20 72 61 74 05 00 03 17 0E 01 D7 0B 01 A4 0D 02 00

      somethingChangedInBps(idConnection) = True
      debugChainType = debugChainType & "1" & ","
      pauseStacking(idConnection) = 0
      debugChainType = debugChainType & "2" & ","
      tempID = CLng(packet(pos + 1)) 'ID
      debugChainType = debugChainType & "3" & ","
      tileID = GetTheLong(packet(pos + 2), packet(pos + 3))
      If tileID > highestDatTile Then
        debugChain = debugChain & " <tile ID not found:" & FiveChrLon(tileID) & "> "
        
      End If
      debugChainType = debugChainType & "4" & ","
      If TibiaVersionLong >= 990 Then
        ' 5 - 7 bytes
        pos = pos + 5
        If (DatTiles(tileID).haveExtraByte = True) Then
          pos = pos + 1
        End If
        If (DatTiles(tileID).haveExtraByte2 = True) Then
          pos = pos + 1
        End If
      Else
        ' 4 - 6 bytes
        If DatTiles(tileID).haveExtraByte = True Then
          pos = pos + 5
        Else
          pos = pos + 4
        End If
        If DatTiles(tileID).haveExtraByte2 = True Then
          pos = pos + 1
        End If
      End If
      debugChainType = debugChainType & "5" & ","
      lonN = GetTheLong(packet(pos), packet(pos + 1))
      pos = pos + 2
      mobName = ""
      debugChainType = debugChainType & "6" & ","
      For itemCount = 1 To lonN
        mobName = mobName & Chr(packet(pos))
        pos = pos + 1
      Next itemCount
      debugChainType = debugChainType & "open:" & mobName & ","
      If mobName = "locker" Then
        lastDepotBPID(idConnection) = CByte(tempID)
      End If
      If mobName = "depot chest" Then
        doneDepotChestOpen(idConnection) = True
      End If
      debugChainType = debugChainType & "7" & ","
      If ((cavebotEnabled(idConnection) = True) And (CheatsPaused(idConnection) = False)) Then
        If Len(mobName) > 5 Then
            If Left$(mobName, 5) = "dead " Then
                requestLootBp(idConnection) = CByte(tempID)
                lootTimeExpire(idConnection) = GetTickCount() + 5000
            End If
            If Len(mobName) > 6 Then
                If Left$(mobName, 6) = "slayn " Then
                    requestLootBp(idConnection) = CByte(tempID)
                    lootTimeExpire(idConnection) = GetTickCount() + 5000
                End If
            End If
        End If
      End If
      lonCap = CLng(packet(pos))
      debugChainType = debugChainType & "cap:" & CStr(lonCap) & ","
      'next byte is action (new or update)
      pos = pos + 2 ' skip container cap
      debugChainType = debugChainType & "8" & ","
      
      
      
      If ((TibiaVersionLong >= 991) Or ((TibiaVersionLong >= 984) And (TibiaVersionLong < 990))) Then
        lonN = GetTheLong(packet(pos), packet(pos + 1))
        debugChainType = debugChainType & "new1:" & CStr(lonN) & ","
        pos = pos + 2 ' skip page number
        lonN = GetTheLong(packet(pos), packet(pos + 1))  ' how many items it have inside
        ' (and 2 unknown new bytes, usually 00 00 )
        debugChainType = debugChainType & "new2:" & CStr(lonN) & ","
      '  Debug.Print debugChainType
        pos = pos + 4
      End If
      
      
      lonN = CLng(packet(pos)) ' how many items it have inside
      pos = pos + 1
    
      
      
      debugChainType = debugChainType & "bpID:" & CStr(tempID) & ","
      debugChainType = debugChainType & "HIGHEST_BP_ID:" & CStr(HIGHEST_BP_ID) & ","
      Backpack(idConnection, tempID).open = True
      Backpack(idConnection, tempID).name = mobName
      Backpack(idConnection, tempID).cap = lonCap
      If lonCap > (HIGHEST_ITEM_BPSLOT + 1) Then
        LogOnFile "errors.txt", "WARNING: detected cotainer with outstanding cap (" & CStr(lonCap) & ")"
      End If
      If lonN > lonCap Then
        debugChainType = debugChainType & "--BAD CAP DETECTED, using fix--,"
        Backpack(idConnection, tempID).used = lonCap
      Else
        Backpack(idConnection, tempID).used = lonN
      End If
      debugChainType = debugChainType & "9" & ","
      debugChainType = debugChainType & "totalItems:" & CStr(lonN) & ","
      For itemCount = 0 To (lonN - 1)
        tileID = GetTheLong(packet(pos), packet(pos + 1))
        debugChainType = debugChainType & "item:" & FiveChrLon(tileID) & ","
        If tileID > highestDatTile Then
          debugChain = debugChain & " <tile ID not found:" & FiveChrLon(tileID) & "> "
        
        End If
        If TibiaVersionLong >= 990 Then
          If (DatTiles(tileID).haveExtraByte = True) Then
              If (DatTiles(tileID).haveExtraByte2 = True) Then
                    Backpack(idConnection, tempID).item(itemCount).t1 = packet(pos)
                    Backpack(idConnection, tempID).item(itemCount).t2 = packet(pos + 1)
                    Backpack(idConnection, tempID).item(itemCount).t3 = packet(pos + 3)
                    Backpack(idConnection, tempID).item(itemCount).t4 = packet(pos + 4)
                    pos = pos + 5
              Else
                    Backpack(idConnection, tempID).item(itemCount).t1 = packet(pos)
                    Backpack(idConnection, tempID).item(itemCount).t2 = packet(pos + 1)
                    Backpack(idConnection, tempID).item(itemCount).t3 = packet(pos + 3)
                    Backpack(idConnection, tempID).item(itemCount).t4 = &H0
                    pos = pos + 4
              End If
          Else
              If (DatTiles(tileID).haveExtraByte2 = True) Then
                    Backpack(idConnection, tempID).item(itemCount).t1 = packet(pos)
                    Backpack(idConnection, tempID).item(itemCount).t2 = packet(pos + 1)
                    Backpack(idConnection, tempID).item(itemCount).t3 = &H0
                    Backpack(idConnection, tempID).item(itemCount).t4 = packet(pos + 3)
                    pos = pos + 4
              Else
                    Backpack(idConnection, tempID).item(itemCount).t1 = packet(pos)
                    Backpack(idConnection, tempID).item(itemCount).t2 = packet(pos + 1)
                    Backpack(idConnection, tempID).item(itemCount).t3 = &H0
                    Backpack(idConnection, tempID).item(itemCount).t4 = &H0
                    pos = pos + 3
              End If
          End If
        Else
            If DatTiles(tileID).haveExtraByte = True Then
                If itemCount <= lonCap - 1 Then
                    Backpack(idConnection, tempID).item(itemCount).t1 = packet(pos)
                    Backpack(idConnection, tempID).item(itemCount).t2 = packet(pos + 1)
                    Backpack(idConnection, tempID).item(itemCount).t3 = packet(pos + 2)
                    Backpack(idConnection, tempID).item(itemCount).t4 = &H0
                End If
                If DatTiles(tileID).haveExtraByte2 = True Then
                   ' Debug.Print "[" & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1)) & " " & GoodHex(packet(pos + 2)) & " " & GoodHex(packet(pos + 3)) & "]"
                    pos = pos + 1 ' skip new extra byte &H0
                Else
                   ' Debug.Print "[" & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1)) & " " & GoodHex(packet(pos + 2)) & "]"
                End If
                pos = pos + 3
            Else
                If itemCount <= lonCap - 1 Then
                    Backpack(idConnection, tempID).item(itemCount).t1 = packet(pos)
                    Backpack(idConnection, tempID).item(itemCount).t2 = packet(pos + 1)
                    Backpack(idConnection, tempID).item(itemCount).t3 = &H0
                    Backpack(idConnection, tempID).item(itemCount).t4 = &H0
                End If
                If DatTiles(tileID).haveExtraByte2 = True Then
                   ' Debug.Print "[" & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1)) & " " & GoodHex(packet(pos + 2)) & "]"
                    pos = pos + 1 ' skip new extra byte &H0
                Else
                   'Debug.Print "[" & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1)) & "]"
                End If
                pos = pos + 2
            End If

        End If
      Next itemCount
      debugChainType = debugChainType & "10" & ","
      For itemCount = lonN To (lonCap - 1) ' adds safety - since Blackd Proxy 8.44
          Backpack(idConnection, tempID).item(itemCount).t1 = &H0
          Backpack(idConnection, tempID).item(itemCount).t2 = &H0
          Backpack(idConnection, tempID).item(itemCount).t3 = &H0
          Backpack(idConnection, tempID).item(itemCount).t4 = &H0
      Next itemCount
      debugChainType = debugChainType & "OK !!"
      ' SendMessageToClient idConnection, "You opened/updated the backpack " & CStr(tempID) & " - " & mobName & " ( cap " & lonN & "/" & lonCap & ")", "GM BlackdProxy"
    Case &H6F
      ' close container
      somethingChangedInBps(idConnection) = True
      pauseStacking(idConnection) = 0
      tempID = CLng(packet(pos + 1)) 'ID
      If (tempID <= HIGHEST_BP_ID) Then
        Backpack(idConnection, tempID).open = False
        Backpack(idConnection, tempID).name = ""
        Backpack(idConnection, tempID).cap = 0
        Backpack(idConnection, tempID).used = 0
      Else
        Debug.Print "BUG: bad parsing of 6F packet. Errors probably happened before this point."
      End If
      ' SendMessageToClient idConnection, "You closed a backpack ID : " & CStr(lonN), "GM BlackdProxy"
      pos = pos + 2
    Case &H70
      ' add item to container ID
      somethingChangedInBps(idConnection) = True
      pauseStacking(idConnection) = 0
      templ1 = CLng(packet(pos + 1))
      If ((TibiaVersionLong >= 991) Or ((TibiaVersionLong >= 984) And (TibiaVersionLong < 990))) Then
        pos = pos + 2 ' skip 2 new strange bytes
      End If
      tileID = GetTheLong(packet(pos + 2), packet(pos + 3))
      If (TibiaVersionLong >= 990) Then
        If (DatTiles(tileID).haveExtraByte = True) Then
          If (DatTiles(tileID).haveExtraByte2 = True) Then
            frmBackpacks.AddItem idConnection, templ1, packet(pos + 2), packet(pos + 3), packet(pos + 5), packet(pos + 6)
            pos = pos + 7
          Else
            frmBackpacks.AddItem idConnection, templ1, packet(pos + 2), packet(pos + 3), packet(pos + 5), &H0
            pos = pos + 6
          End If
        Else
          If (DatTiles(tileID).haveExtraByte2 = True) Then
            frmBackpacks.AddItem idConnection, templ1, packet(pos + 2), packet(pos + 3), &H0, packet(pos + 5)
            pos = pos + 6
          Else
            frmBackpacks.AddItem idConnection, templ1, packet(pos + 2), packet(pos + 3), &H0, &H0
            pos = pos + 5
          End If
        End If
      Else
        If DatTiles(tileID).haveExtraByte = True Then
          frmBackpacks.AddItem idConnection, templ1, packet(pos + 2), packet(pos + 3), packet(pos + 4)
          pos = pos + 5
        Else
          frmBackpacks.AddItem idConnection, templ1, packet(pos + 2), packet(pos + 3), &H0
          pos = pos + 4
        End If
        If DatTiles(tileID).haveExtraByte2 = True Then
          pos = pos + 1 ' skip new extra byte 00
        End If
      End If
    Case &H71
      ' something transform in bp
      'Debug.Print frmMain.showAsStr3(packet, True, pos, pos + 400)
      somethingChangedInBps(idConnection) = True
      pauseStacking(idConnection) = 0
      ' transform container ID (1) , slot (2) to other item
      templ1 = packet(pos + 1)
      If ((TibiaVersionLong >= 991) Or ((TibiaVersionLong >= 984) And (TibiaVersionLong < 990))) Then
        templ2 = GetTheLong(packet(pos + 2), packet(pos + 3))
        pos = pos + 1 ' 1 byte bigger
      Else
        templ2 = packet(pos + 2)
      End If
      tileID = GetTheLong(packet(pos + 3), packet(pos + 4))
      If TibiaVersionLong >= 990 Then
        If (DatTiles(tileID).haveExtraByte = True) Then
          If (DatTiles(tileID).haveExtraByte2 = True) Then
            frmBackpacks.UpdateItem idConnection, templ1, templ2, packet(pos + 3), packet(pos + 4), packet(pos + 6), packet(pos + 7)
            pos = pos + 8
          Else
            frmBackpacks.UpdateItem idConnection, templ1, templ2, packet(pos + 3), packet(pos + 4), packet(pos + 6), &H0
            pos = pos + 7
          End If
        Else
          If (DatTiles(tileID).haveExtraByte2 = True) Then
            frmBackpacks.UpdateItem idConnection, templ1, templ2, packet(pos + 3), packet(pos + 4), &H0, packet(pos + 6)
            pos = pos + 7
          Else
            frmBackpacks.UpdateItem idConnection, templ1, templ2, packet(pos + 3), packet(pos + 4), &H0, &H0
            pos = pos + 6
          End If
        End If
      Else
        If DatTiles(tileID).haveExtraByte = True Then
          frmBackpacks.UpdateItem idConnection, templ1, templ2, packet(pos + 3), packet(pos + 4), packet(pos + 5)
          pos = pos + 6
        Else
          frmBackpacks.UpdateItem idConnection, templ1, templ2, packet(pos + 3), packet(pos + 4), &H0
          pos = pos + 5
        End If
        If DatTiles(tileID).haveExtraByte2 = True Then
           pos = pos + 1
        End If
      End If
    Case &H72
      ' remove item from container ID
      somethingChangedInBps(idConnection) = True
      pauseStacking(idConnection) = 0
      templ1 = packet(pos + 1)
      If ((TibiaVersionLong >= 991) Or ((TibiaVersionLong >= 984) And (TibiaVersionLong < 990))) Then
        ' slot id
        templ2 = GetTheLong(packet(pos + 2), packet(pos + 3))
        ' and 2 extra bytes , usually 00 00
        pos = pos + 6
      Else
        templ2 = packet(pos + 2)
        pos = pos + 3
      End If
      frmBackpacks.RemoveItem idConnection, templ1, templ2
     
    Case &H78
      ' inventory slot get something
      tileID = GetTheLong(packet(pos + 2), packet(pos + 3))
      If ((tileID < 0) Or (tileID > highestDatTile)) Then
        myres.fail = True
        LogOnFile "errors.txt", "Unexpected tileID at ptype 78 ( " & tileID & " ) considering highestDatTile = " & CStr(highestDatTile)
        aRes = GiveGMmessage(idConnection, "Unexpected tileID at ptype 78 ( " & tileID & " ) please report blackd", "BlackdProxy")
        DoEvents
        pos = pos + 10000
      Else
        If TibiaVersionLong >= 990 Then
            If DatTiles(tileID).haveExtraByte = True Then
              If DatTiles(tileID).haveExtraByte2 = True Then
                mySlot(idConnection, packet(pos + 1)).t1 = packet(pos + 2)
                mySlot(idConnection, packet(pos + 1)).t2 = packet(pos + 3)
                mySlot(idConnection, packet(pos + 1)).t3 = packet(pos + 5)
                mySlot(idConnection, packet(pos + 1)).t4 = packet(pos + 6)
                pos = pos + 7
              Else
                mySlot(idConnection, packet(pos + 1)).t1 = packet(pos + 2)
                mySlot(idConnection, packet(pos + 1)).t2 = packet(pos + 3)
                mySlot(idConnection, packet(pos + 1)).t3 = packet(pos + 5)
                mySlot(idConnection, packet(pos + 1)).t4 = &H0
                pos = pos + 6
              End If
            Else
              If DatTiles(tileID).haveExtraByte2 = True Then
                mySlot(idConnection, packet(pos + 1)).t1 = packet(pos + 2)
                mySlot(idConnection, packet(pos + 1)).t2 = packet(pos + 3)
                mySlot(idConnection, packet(pos + 1)).t3 = &H0
                mySlot(idConnection, packet(pos + 1)).t4 = packet(pos + 5)
                pos = pos + 6
              Else
                mySlot(idConnection, packet(pos + 1)).t1 = packet(pos + 2)
                mySlot(idConnection, packet(pos + 1)).t2 = packet(pos + 3)
                mySlot(idConnection, packet(pos + 1)).t3 = &H0
                mySlot(idConnection, packet(pos + 1)).t4 = &H0
                pos = pos + 5
              End If
            End If
        Else
            If DatTiles(tileID).haveExtraByte = True Then
              mySlot(idConnection, packet(pos + 1)).t1 = packet(pos + 2)
              mySlot(idConnection, packet(pos + 1)).t2 = packet(pos + 3)
              mySlot(idConnection, packet(pos + 1)).t3 = packet(pos + 4)
              mySlot(idConnection, packet(pos + 1)).t4 = &H0
              pos = pos + 5
            Else
              mySlot(idConnection, packet(pos + 1)).t1 = packet(pos + 2)
              mySlot(idConnection, packet(pos + 1)).t2 = packet(pos + 3)
              mySlot(idConnection, packet(pos + 1)).t3 = &H0
              mySlot(idConnection, packet(pos + 1)).t4 = &H0
              pos = pos + 4
            End If

            If DatTiles(tileID).haveExtraByte2 = True Then
              pos = pos + 1
            End If
        
        End If
      End If
    Case &H79
      ' remove inventory slot
      mySlot(idConnection, packet(pos + 1)).t1 = &H0
      mySlot(idConnection, packet(pos + 1)).t2 = &H0
      mySlot(idConnection, packet(pos + 1)).t3 = &H0
      pos = pos + 2
    Case &H7D
      'trade part 1
      lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
      pos = pos + 3 + lonN
      lonN = CLng(packet(pos))
      pos = pos + 1
      For itemCount = 1 To lonN
      
      
      
      
    
        tileID = GetTheLong(packet(pos), packet(pos + 1))
        Debug.Print "debug> " & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1))
        
        If TibiaVersionLong >= 990 Then
            If DatTiles(tileID).haveExtraByte = True Then
              If DatTiles(tileID).haveExtraByte2 = True Then
                pos = pos + 5
              Else
                pos = pos + 4
              End If
            Else
              If DatTiles(tileID).haveExtraByte2 = True Then
                pos = pos + 4
              Else
                pos = pos + 3
              End If
            End If
        Else
          ' older tibia
            If DatTiles(tileID).haveExtraByte = True Then
              pos = pos + 3
            Else
              pos = pos + 2
            End If
            If DatTiles(tileID).haveExtraByte2 = True Then
             pos = pos + 1 ' skip new extra byte 00
            End If
        End If
      Next itemCount
    Case &H7E
      'trade part 2
      lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
      pos = pos + 3 + lonN
      lonN = CLng(packet(pos))
      pos = pos + 1
      For itemCount = 1 To lonN
        tileID = GetTheLong(packet(pos), packet(pos + 1))
        If TibiaVersionLong >= 990 Then
            If DatTiles(tileID).haveExtraByte = True Then
              If DatTiles(tileID).haveExtraByte2 = True Then
                pos = pos + 5
              Else
                pos = pos + 4
              End If
            Else
              If DatTiles(tileID).haveExtraByte2 = True Then
                pos = pos + 4
              Else
                pos = pos + 3
              End If
            End If
        Else
        ' older tibia
            If DatTiles(tileID).haveExtraByte = True Then
              pos = pos + 3
            Else
              pos = pos + 2
            End If
            If DatTiles(tileID).haveExtraByte2 = True Then
             pos = pos + 1 ' skip new extra byte 00
            End If
        End If
      Next itemCount
    Case &H7F
      ' unknown , but fixed lenght
      pos = pos + 1
    Case &H82
      ' light update
      pos = pos + 3
    Case &H83
      ' teleport
       pos = pos + 7
    Case &H84
      ' animated text (floating EXP number)
      lonN = GetTheLong(packet(pos + 7), packet(pos + 8))
      pos = pos + 9 + lonN
    Case &H85
      'show proyectile
      pos = pos + 12
    Case &H86
      ' getting direct attack
      tempID = FourBytesDouble(packet(pos + 1), packet(pos + 2), packet(pos + 3), packet(pos + 4))
      nameofgivenID = GetNameFromID(idConnection, tempID)
     ' aRes = SendSystemMessageToClient(idConnection, "You are being attacked by " & nameOfGivenID)
      'DangerPK
      If nameofgivenID = "" Then
        ' attacked by nothing? ... ignore
      Else
        If SelfDefenseID(idConnection) = 0 Then
          SelfDefenseID(idConnection) = tempID
        End If
        If DangerPK(idConnection) = False Then
          If (cavebotEnabled(idConnection) = True) Then
            If (PKwarnings(idConnection) = True) Then ' new since Blackd Proxy 23.8
                templ1 = GetTickCount()
                If templ1 > (CavebotTimeStart(idConnection) + 5000) Then ' new since Blackd Proxy 24.0
                    templ1 = exeLine(idConnection)
                    temps1 = GetStringFromIDLine(idConnection, templ1)
                    If Len(temps1) > 2 Then
                        If (isMelee(idConnection, nameofgivenID) = False) And (isHmm(idConnection, nameofgivenID) = False) And (frmRunemaker.IsFriend(LCase(nameofgivenID)) = False) Then
                            DangerPKname(idConnection) = nameofgivenID
                            DangerPK(idConnection) = True
                        End If
                    End If
                End If
            End If
          ElseIf RuneMakerOptions(idConnection).activated = True Then
            If frmRunemaker.IsFriend(LCase(nameofgivenID)) = False Then
              DangerPKname(idConnection) = nameofgivenID
              DangerPK(idConnection) = True
            End If
          End If
        End If
      End If
      pos = pos + 6
    Case &H87 ' new since Tibia 8.7 , "number of trappers" , ids of persons that are trapping you
        ' 87 02 C6 11 52 01 AE AA DF 01
        ' 87 01 CF 89 57 02
        ' 87 00
        ' 87 04 D6 C5 4E 02 17 58 28 02 20 36 B7 01 CB C9 1F 02
        templ1 = CLng(packet(pos + 1)) ' number of trappers
        pos = pos + 2 + 4 * templ1 ' skip ids
    Case &H8C
      ' hp update
      tempID = FourBytesDouble(packet(pos + 1), packet(pos + 2), packet(pos + 3), packet(pos + 4))

      If CheatsPaused(idConnection) = False Then
        
        templ1 = CLng(packet(pos + 5))
        AddID_HP idConnection, tempID, packet(pos + 5)
        If TrainerOptions(idConnection).misc_stoplowhp = 1 Then
          If tempID = currTargetID(idConnection) Then
            If (templ1 < TrainerOptions(idConnection).stoplowhpHP) Then
              'If (lastAttackedID(idconnection) <> 0) Then
               aRes = MeleeAttack(idConnection, 0, True)
               lastAttackedID(idConnection) = 0
               'End If
            'Else
             ' If cavebotEnabled(idConnection) = False Then
             '   aRes = MeleeAttack(idConnection, tempID)
             '   lastAttackedID(idConnection) = tempID
             ' End If
            End If
          End If
        End If
      End If
      pos = pos + 6
    Case &H8D
      ' Light update
      If (frmHardcoreCheats.chkLight.Value = 1) Then
        ' keep cheat light - NEW since 25.8
        tmpStr = GoodHex(packet(pos + 1)) & GoodHex(packet(pos + 2)) & GoodHex(packet(pos + 3)) & GoodHex(packet(pos + 4))
        If (tmpStr = IDstring(idConnection)) Then
          packet(pos + 5) = CByte("&H" & LightIntesityHex)
          packet(pos + 6) = CByte("&H" & nextLight(idConnection))
        End If
      End If
      pos = pos + 7
    Case &H8E
      ' SOMEONE CHANGES OUTFIT (confirmed)
      ' 00 DA 07 = oracle outfit
                  ' 8E B8 E7 25 02 80 00 00 00 00 00 00 00 00
      ' tibia 8.7 : 8E 05 39 ED 01 80 00 72 64 72 63 00 00 00
      '             8E 7B 23 00 40 26 01 00 00 00 00 00 00 00
      tempID = FourBytesDouble(packet(pos + 1), packet(pos + 2), packet(pos + 3), packet(pos + 4))
      pos = pos + 5
     
      outfitType = CLng(packet(pos))
      
      If TibiaVersionLong >= 870 Then
        outfitType = GetTheLong(packet(pos), packet(pos + 1))
        'Debug.Print "outfit:" & outfitType & "; " & GoodHex(packet(pos)) & " ; " & GoodHex(packet(pos + 1))
        pos = pos + 1
      ElseIf TibiaVersionLong <= 760 Then
        outfitType = CLng(packet(pos))
      Else ' Tibia 7.64 and above
        outfitType = GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 1
      End If
      
      ' now skip enough bytes for the outfit
      If (outfitType = 0) Then ' thing outfit
        If (packet(pos + 1) = &H0) And (packet(pos + 2) = &H0) And (tempID <> myID(idConnection)) And (frmHardcoreCheats.chkReveal.Value = 1) Then
          nameofgivenID = GetNameFromID(idConnection, tempID)
          packet(pos + 1) = LowByteOfLong(tileID_Oracle)
          packet(pos + 2) = HighByteOfLong(tileID_Oracle)
          'aRes = SendSystemMessageToClient(idConnection, nameOfGivenID & " cant hide from you =)")
        End If
        pos = pos + 3 'if tile was 00 00 -> invisible will be removed
      Else
        pos = pos + 5
        If TibiaVersionLong >= 773 Then
          pos = pos + 1 ' new strange outfit tag
        End If
      End If
      If TibiaVersionLong >= 870 Then
        pos = pos + 2 ' xx3 new strange outfit tag
      End If
    Case &H8F
      ' someone (4 bytes ID) change speed (2 bytes)
      If TibiaVersionLong < 1059 Then
        pos = pos + 7
      Else
        ' It probably adds an effect (2 bytes) since Tibia 10.59
        ' 8F 3E 1D D0 02 9F 01 B3 01
        pos = pos + 9
      End If
    Case &H90
     ' being attacked? something updates
      pos = pos + 6
    Case &H91
      ' party invite
      pos = pos + 6
    Case &H92
      ' unknown. mounting in something? New since Tibia 8.7
      If (packet(pos + 5) = &H1) Then
        ' new detection of being attacked
        tempID = FourBytesDouble(packet(pos + 1), packet(pos + 2), packet(pos + 3), packet(pos + 4))
        nameofgivenID = GetNameFromID(idConnection, tempID)
        
        
        If nameofgivenID = "" Then
        
        Else
        If SelfDefenseID(idConnection) = 0 Then
          SelfDefenseID(idConnection) = tempID
        End If
        If DangerPK(idConnection) = False Then
          If (cavebotEnabled(idConnection) = True) Then
            If (PKwarnings(idConnection) = True) Then ' new since Blackd Proxy 23.8
                templ1 = GetTickCount()
                If templ1 > (CavebotTimeStart(idConnection) + 5000) Then ' new since Blackd Proxy 24.0
                    templ1 = exeLine(idConnection)
                    temps1 = GetStringFromIDLine(idConnection, templ1)
                    If Len(temps1) > 2 Then
                        If (isMelee(idConnection, nameofgivenID) = False) And (isHmm(idConnection, nameofgivenID) = False) And (frmRunemaker.IsFriend(LCase(nameofgivenID)) = False) Then
                            DangerPKname(idConnection) = nameofgivenID
                            DangerPK(idConnection) = True
                        End If
                    End If
                End If
            End If
          ElseIf RuneMakerOptions(idConnection).activated = True Then
            If frmRunemaker.IsFriend(LCase(nameofgivenID)) = False Then
              DangerPKname(idConnection) = nameofgivenID
              DangerPK(idConnection) = True
            End If
          End If
        End If
        
        
        'Debug.Print nameofgivenID & " is attacking you"
      End If
      End If
      pos = pos + 6
      'Debug.Print "strange thing happened: 92"
    Case &H93
    

      If TibiaVersionLong >= 1035 Then
            ' tibia 10.35: fixed length??
      ' 93 AE B6 85 02 02 FF 92 AE B6 85 02 00
      ' util? = 93 AE B6 85 02 02 FF
         pos = pos + 7
      Else
        ' special effects at creature id. New since Tibia 9.9
        ' 93 01 F1 F3 C0 02 02 FF
        ' 93 02 F3 37 9E 02 02 FF 07 F5 02 40 02 FF
        '    ?? -----id.--- ?? ??
        ' Debug.Print "NEW: " & frmMain.showAsStr3(packet, True, pos, pos + 7)
        pos = pos + 2 + (CLng(packet(pos + 1)) * 6)
      End If
    Case &H94
      ' more effects - happens at summoned monks, for example
      ' 94 13 44 A7 02 02 00
      pos = pos + 7
    Case &H96
      ' open book
      ' tibia 10.01 : 96 2A FF 74 02 B1 0D FF CF 07 00 00 00 00 00 00
      If TibiaVersionLong >= 992 Then
        lonN = GetTheLong(packet(pos + 10), packet(pos + 11))
        pos = pos + 12 + lonN
        lonN = GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 2 + lonN
        lonN = GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 2 + lonN
      ElseIf TibiaVersionLong > 781 Then
        lonN = GetTheLong(packet(pos + 9), packet(pos + 10))
        pos = pos + 11 + lonN
        lonN = GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 2 + lonN
        lonN = GetTheLong(packet(pos), packet(pos + 1)) 'new : edition date
        pos = pos + 2 + lonN
      ElseIf TibiaVersionLong >= 760 Then
        lonN = GetTheLong(packet(pos + 9), packet(pos + 10))
        pos = pos + 11 + lonN
        lonN = GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 2 + lonN
      Else
        lonN = GetTheLong(packet(pos + 9), packet(pos + 10))
        pos = pos + 11 + lonN
      End If
    Case &H97
      ' guildhall book
      lonN = GetTheLong(packet(pos + 6), packet(pos + 7))
      pos = pos + 8 + lonN
    Case &H9C
      ' new since tibia 10.55
      ' Somehow related with hotkeys settings?
      ' 9C 01 00
     ' lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
      'Debug.Print "char selected hotkey set #" & CStr(lonN)
      pos = pos + 3
      
    Case &H9D
      ' UNDER INVESTIGATION - MAYBE IT IS THE AUTOBAN / BAN WAVE PACKET
      ' tibia 10.55: 9D 02 00 00 00
      ' tibia 10.56: 9D 01 00 00 00
      
      ' Parser fixed in Blackd Proxy 33.4
      'nameofgivenID = "[WARNING!] Request to search cheats? Please report this log with timestamp to daniel@blackdtools.com .Strange packet received: " & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1)) & " " & GoodHex(packet(pos + 2))
      'LogOnFile "errors.txt", nameofgivenID
      'aRes = SendLogSystemMessageToClient(idConnection, nameofgivenID)
     ' DoEvents
      If TibiaVersionLong >= 1055 Then
        pos = pos + 5
      Else
        pos = pos + 3
      End If
      
    Case &H9E
      ' 9E 05 00 01 02 03 04 00
      ' 9E 05 0B 00 01 02 03 01
      ' 9E 01 0E 01
      lonN = CLng(packet(pos + 1))
      ' new since Tibia 10.38
      ' "Premium features" window
      pos = pos + 3 + lonN
    Case &H9F
      ' 9F 00 01 01 00 0A
      ' 9.5   9F 00 04 2B 00 01 02 04 05 06 07 08 09 0A 0B 0C 0E 11 14 19 1A 1B 1D 1E 1F 20 26 27 2A 2C 32 4B 4C 4D 4E 51 53 54 58 59 5B 5E 70 71 72 79 92 94
      ' 10.38 9F 00 64 6E 80 4E 01 01 00 0A
      ' 10.55 9F 00 A0 49 26 4A 00 00 00
      ' unknown,  since Tibia 9.5
      If TibiaVersionLong < 1038 Then
        lonN = GetTheLong(packet(pos + 3), packet(pos + 4))
        pos = pos + 5 + lonN
      Else
        lonN = GetTheLong(packet(pos + 7), packet(pos + 8))
        pos = pos + 9 + lonN
      End If
    Case &HA0
      ' full stats update : hp,mana,exp,etc
      ' necesita ser revisado
      oldHP = myHP(idConnection)
      If (TibiaVersionLong >= 1054) Then
'      If ((TibiaVersionLong >= 1053) And (tibiaclassname = "TibiaClientPreview")) Then
        ' tibia 10.53 preview
        ' A0 39 00 96 00 F8 7A 00 00 40 9C 00 00 5C 00 00 00 00 00 00 00 01 00 5C 04 0F 27 00 80 05 00 05 00 00 00 00 64 D8 09 6E 00 00 00 D0 02
          '  01 02 03 04 05 06 07 08 09 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37 38
        
        
        
        
        ' A1 0A 00 0A 00 00 0C 00 0C 00 34 0A 00 0A 00 00 0A 00 0A 00 00 0A 00 0A 00 00 0A 00 0A 00 00 0A 00 0A 00 00
        
        

        lonN = GetTheLong(packet(pos + 1), packet(pos + 2)) ' current hp
        If lonN <> myHP(idConnection) Then
          myres.gotHPupdate = True
          myHP(idConnection) = lonN
        End If
        myMaxHP(idConnection) = GetTheLong(packet(pos + 3), packet(pos + 4)) ' max hp
        lonN = GetTheLong(packet(pos + 24), packet(pos + 25)) 'PLAYER_MANA
        If lonN <> myMana(idConnection) Then
          myres.gotManaupdate = True
          myMana(idConnection) = lonN
        End If
        myCap(idConnection) = FourBytesLong(packet(pos + 5), packet(pos + 6), packet(pos + 7), packet(pos + 8)) ' cap x 100
        myExp(idConnection) = FourBytesLong(packet(pos + 13), packet(pos + 14), packet(pos + 15), packet(pos + 16))
       
        lonN = GetTheLong(packet(pos + 21), packet(pos + 22)) ' PLAYER_LEVEL
        If lonN > myLevel(idConnection) Then
          If sentWelcome(idConnection) = True Then
            If frmHardcoreCheats.chkAutoGratz.Value = 1 Then
              SendLogSystemMessageToClient idConnection, "BlackdProxy: Gratz!"
              DoEvents
            End If
          End If
        End If
        myLevel(idConnection) = lonN
        
        lonN = GetTheLong(packet(pos + 29), packet(pos + 30)) ' PLAYER_MANA
        If lonN <> myMana(idConnection) Then
          myres.gotManaupdate = True
          myMana(idConnection) = lonN
        End If
        
          myMaxMana(idConnection) = GetTheLong(packet(pos + 31), packet(pos + 32))
          myMagLevel(idConnection) = CLng(packet(pos + 33))
    
          lonN = CLng(packet(pos + 34)) ' PLAYER_MAGIC_LEVEL_PER
          myNewStat(idConnection) = lonN
          
          lonN = CLng(packet(pos + 36)) ' soulpoints
          If lonN <> mySoulpoints(idConnection) Then
            myres.gotSoulupdate = True
          End If
          mySoulpoints(idConnection) = lonN
    
        
          myStamina(idConnection) = GetTheLong(packet(pos + 37), packet(pos + 38))


        
        pos = pos + 45
      ElseIf TibiaVersionLong >= 872 Then
'    WriteByte(0xA0);
'    WriteUInt16(*PLAYER_HP); '1,2
'    WriteUInt16(*PLAYER_HP_MAX);  '3,4
'    WriteUInt32(*PLAYER_CAP);  '5,6,7,8
'    WriteUInt32(0); // unknown - flash client '9,10,11,12
'    WriteUInt64(*PLAYER_EXP); '13,14,15,16,17,18,19,20
'    WriteUInt16(*PLAYER_LEVEL); ,21,22
'    WriteByte(*PLAYER_LEVEL_PER); ,23
'    WriteUInt16(*PLAYER_MANA); ,24,25
'    WriteUInt16(*PLAYER_MANA_MAX); '26,27
'    WriteByte(*PLAYER_MAGIC_LEVEL); '28
'    WriteByte(0); // unknown - flash client ' 29
'    WriteByte(*PLAYER_MAGIC_LEVEL_PER); '30
'    WriteByte(*PLAYER_SOUL); '31
'    WriteUInt16(*PLAYER_STAMINA); '32,33
'    WriteUInt16(0);// unknown - flash client '34,35
'    WriteUInt16(0);// unknown - flash client '36,37
        'Debug.Print frmMain.showAsStr3(packet, True, pos, pos + 37)
        ' ( hex ) A0 9E 00 AF 00 18 6A 00 00 C8 AF 00 00 0B 07 00 00 00 00 00 00 06 00 1B 19 00 19 00 00 00 00 64 D8 09 E6 00 00 00
        '            01 02 03 04 05 06 07 08 09 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37
        
     ' tibia 10.53 preview
                 'A0 39 00 96 00 F8 7A 00 00 40 9C 00 00 5C 00 00 00 00 00 00 00 01 00 5C 04 0F 27 00 80 05 00 05 00 00 00 00 64 D8 09 6E 00 00 00 D0 02 A1 0A 00 0A 00 00 0C 00 0C 00 34 0A 00 0A 00 00 0A 00 0A 00 00 0A 00 0A 00 00 0A 00 0A 00 00 0A 00 0A 00 00

        
        
        
        
        
        
        lonN = GetTheLong(packet(pos + 1), packet(pos + 2)) ' current hp
        If lonN <> myHP(idConnection) Then
          myres.gotHPupdate = True
          myHP(idConnection) = lonN
        End If
        myMaxHP(idConnection) = GetTheLong(packet(pos + 3), packet(pos + 4)) ' max hp
        lonN = GetTheLong(packet(pos + 24), packet(pos + 25)) 'PLAYER_MANA
        If lonN <> myMana(idConnection) Then
          myres.gotManaupdate = True
          myMana(idConnection) = lonN
        End If
        myCap(idConnection) = FourBytesLong(packet(pos + 5), packet(pos + 6), packet(pos + 7), packet(pos + 8)) ' cap x 100
        myExp(idConnection) = FourBytesLong(packet(pos + 13), packet(pos + 14), packet(pos + 15), packet(pos + 16))
       
        lonN = GetTheLong(packet(pos + 21), packet(pos + 22)) ' PLAYER_LEVEL
        If lonN > myLevel(idConnection) Then
          If sentWelcome(idConnection) = True Then
            If frmHardcoreCheats.chkAutoGratz.Value = 1 Then
              SendLogSystemMessageToClient idConnection, "BlackdProxy: Gratz!"
              DoEvents
            End If
          End If
        End If
        myLevel(idConnection) = lonN
        
        lonN = GetTheLong(packet(pos + 24), packet(pos + 25)) ' PLAYER_MANA
        If lonN <> myMana(idConnection) Then
          myres.gotManaupdate = True
          myMana(idConnection) = lonN
        End If
        
          myMaxMana(idConnection) = GetTheLong(packet(pos + 26), packet(pos + 27))
          myMagLevel(idConnection) = CLng(packet(pos + 28))
    
          lonN = CLng(packet(pos + 30)) ' PLAYER_MAGIC_LEVEL_PER
          myNewStat(idConnection) = lonN
          
          lonN = CLng(packet(pos + 31)) ' soulpoints
          If lonN <> mySoulpoints(idConnection) Then
            myres.gotSoulupdate = True
        If TibiaVersionLong <= 772 Then
          mySoulpoints(idConnection) = 100
        Else
          mySoulpoints(idConnection) = lonN
        End If
          End If
          myStamina(idConnection) = GetTheLong(packet(pos + 32), packet(pos + 33))
          pos = pos + 38 ' new
          
          If TibiaVersionLong >= 960 Then
            ' new stat: Offline training time: full = 720
            ' I will just ignore it.
            ' lonN = GetTheLong(packet(pos), packet(pos + 1))
            pos = pos + 2
          End If

      ElseIf TibiaVersionLong >= 870 Then
        lonN = GetTheLong(packet(pos + 1), packet(pos + 2)) ' current hp
        If lonN <> myHP(idConnection) Then
          myres.gotHPupdate = True
          myHP(idConnection) = lonN
        End If
        myMaxHP(idConnection) = GetTheLong(packet(pos + 3), packet(pos + 4)) ' max hp
        lonN = GetTheLong(packet(pos + 16), packet(pos + 17))
        If lonN <> myMana(idConnection) Then
          myres.gotManaupdate = True
          myMana(idConnection) = lonN
        End If
        myCap(idConnection) = FourBytesLong(packet(pos + 5), packet(pos + 6), packet(pos + 7), packet(pos + 8)) ' cap x 100
        myExp(idConnection) = FourBytesLong(packet(pos + 9), packet(pos + 10), packet(pos + 11), packet(pos + 12))
        ' new stats 8.7: 13,14,15,16 IGNORED
        lonN = GetTheLong(packet(pos + 17), packet(pos + 18)) ' level
        If lonN > myLevel(idConnection) Then
          If sentWelcome(idConnection) = True Then
            If frmHardcoreCheats.chkAutoGratz.Value = 1 Then
              SendLogSystemMessageToClient idConnection, "BlackdProxy: Gratz!"
              DoEvents
            End If
          End If
        End If
        myLevel(idConnection) = lonN
        
        ' stat 19 ? IGNORED
        'Debug.Print packet(pos + 19)
        lonN = GetTheLong(packet(pos + 20), packet(pos + 21)) ' mana
        If lonN <> myMana(idConnection) Then
          myres.gotManaupdate = True
          myMana(idConnection) = lonN
        End If
        
          myMaxMana(idConnection) = GetTheLong(packet(pos + 22), packet(pos + 23))
          myMagLevel(idConnection) = CLng(packet(pos + 24)) ' tibia 7.6
    
          lonN = CLng(packet(pos + 25)) ' new stat
          myNewStat(idConnection) = lonN
          
          lonN = CLng(packet(pos + 26)) ' soulpoints
          If lonN <> mySoulpoints(idConnection) Then
            myres.gotSoulupdate = True
        If TibiaVersionLong <= 772 Then
          mySoulpoints(idConnection) = 100
        Else
          mySoulpoints(idConnection) = lonN
        End If
          End If
          myStamina(idConnection) = GetTheLong(packet(pos + 27), packet(pos + 28))
          pos = pos + 29
      ElseIf TibiaVersionLong >= 830 Then
      
'
          myCap(idConnection) = FourBytesLong(packet(pos + 5), packet(pos + 6), packet(pos + 7), packet(pos + 8)) ' cap x 100
           
          lonN = GetTheLong(packet(pos + 13), packet(pos + 14))
          If lonN > myLevel(idConnection) Then
            If sentWelcome(idConnection) = True Then
              If frmHardcoreCheats.chkAutoGratz.Value = 1 Then
                SendLogSystemMessageToClient idConnection, "BlackdProxy: Gratz!"
                DoEvents
              End If
            End If
          End If
          myLevel(idConnection) = lonN
          
          myExp(idConnection) = FourBytesLong(packet(pos + 9), packet(pos + 10), packet(pos + 11), packet(pos + 12))

          lonN = GetTheLong(packet(pos + 1), packet(pos + 2)) ' current hp OK
          If lonN <> myHP(idConnection) Then
            myres.gotHPupdate = True
            myHP(idConnection) = lonN
          End If
          myMaxHP(idConnection) = GetTheLong(packet(pos + 3), packet(pos + 4)) ' max hp OK
          lonN = GetTheLong(packet(pos + 16), packet(pos + 17)) ' mana
          If lonN <> myMana(idConnection) Then
            myres.gotManaupdate = True
            myMana(idConnection) = lonN
          End If
  
      
          myMaxMana(idConnection) = GetTheLong(packet(pos + 18), packet(pos + 19))
          myMagLevel(idConnection) = CLng(packet(pos + 20)) ' tibia 7.6
    
          lonN = CLng(packet(pos + 21)) ' new stat
          myNewStat(idConnection) = lonN
          
          lonN = CLng(packet(pos + 22)) ' tibia 7.6
          If lonN <> mySoulpoints(idConnection) Then
            myres.gotSoulupdate = True
        If TibiaVersionLong <= 772 Then
          mySoulpoints(idConnection) = 100
        Else
          mySoulpoints(idConnection) = lonN
        End If
          End If
          pos = pos + 23
          myStamina(idConnection) = GetTheLong(packet(pos), packet(pos + 1))
          pos = pos + 2

          
          

      
      ElseIf TibiaVersionLong >= 760 Then
      
      myCap(idConnection) = GetTheLong(packet(pos + 5), packet(pos + 6))
       
      If CLng(packet(pos + 11)) > myLevel(idConnection) Then
        If sentWelcome(idConnection) = True Then
          If frmHardcoreCheats.chkAutoGratz.Value = 1 Then
            SendLogSystemMessageToClient idConnection, "BlackdProxy: Gratz!"
            DoEvents
          End If
        End If
      End If
      myExp(idConnection) = FourBytesLong(packet(pos + 7), packet(pos + 8), packet(pos + 9), packet(pos + 10))
      myLevel(idConnection) = GetTheLong(packet(pos + 11), packet(pos + 12))
      lonN = GetTheLong(packet(pos + 1), packet(pos + 2)) ' current hp OK
      If lonN <> myHP(idConnection) Then
        myres.gotHPupdate = True
        myHP(idConnection) = lonN
      End If
      myMaxHP(idConnection) = GetTheLong(packet(pos + 3), packet(pos + 4)) ' max hp OK
      lonN = GetTheLong(packet(pos + 14), packet(pos + 15))
      If lonN <> myMana(idConnection) Then
        myres.gotManaupdate = True
        myMana(idConnection) = lonN
      End If
      ' 13 = % to mag lv up
      ' 16 + 17 = max Mana - Ignored
      myMaxMana(idConnection) = GetTheLong(packet(pos + 16), packet(pos + 17))
      myMagLevel(idConnection) = CLng(packet(pos + 18)) ' tibia 7.6

      lonN = CLng(packet(pos + 19)) ' new stat
      myNewStat(idConnection) = lonN
      
      lonN = CLng(packet(pos + 20)) ' tibia 7.6
      If lonN <> mySoulpoints(idConnection) Then
        myres.gotSoulupdate = True
        If TibiaVersionLong <= 772 Then
          mySoulpoints(idConnection) = 100
        Else
          mySoulpoints(idConnection) = lonN
        End If
      End If
      pos = pos + 21
        If TibiaVersionLong >= 773 Then ' NEW STAT : STAMINA
          myStamina(idConnection) = GetTheLong(packet(pos), packet(pos + 1))
         pos = pos + 2
        End If
      Else ' Tibia 7.5 and under
      myCap(idConnection) = GetTheLong(packet(pos + 5), packet(pos + 6))
      
      If CLng(packet(pos + 11)) > myLevel(idConnection) Then
        If sentWelcome(idConnection) = True Then
          If frmHardcoreCheats.chkAutoGratz.Value = 1 Then
            SendLogSystemMessageToClient idConnection, "BlackdProxy: Gratz!"
            DoEvents
          End If
        End If
      End If
      
      myExp(idConnection) = FourBytesLong(packet(pos + 7), packet(pos + 8), packet(pos + 9), packet(pos + 10))
      myLevel(idConnection) = CLng(packet(pos + 11))
      lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
      If lonN <> myHP(idConnection) Then
        myres.gotHPupdate = True
        myHP(idConnection) = lonN
      End If
      myMaxHP(idConnection) = GetTheLong(packet(pos + 3), packet(pos + 4))
      lonN = GetTheLong(packet(pos + 13), packet(pos + 14))
      If lonN <> myMana(idConnection) Then
        myres.gotManaupdate = True
        myMana(idConnection) = lonN
      End If
      myMaxMana(idConnection) = GetTheLong(packet(pos + 15), packet(pos + 16))
      myMagLevel(idConnection) = CLng(packet(pos + 17))
      lonN = CLng(packet(pos + 19))
      If lonN <> mySoulpoints(idConnection) Then
        myres.gotSoulupdate = True
        If TibiaVersionLong <= 772 Then
          mySoulpoints(idConnection) = 100
        Else
          mySoulpoints(idConnection) = lonN
        End If
      End If
      pos = pos + 20
      
      End If
      
      If oldHP <> cte_initHP Then
        lastHPchange(idConnection) = myHP(idConnection) - oldHP
        If (0 - lastHPchange(idConnection)) >= maxHit(idConnection) Then
           ' IGNORE CURRENT TARGET
           If cavebotEnabled(idConnection) = True Then
           
              If lastAttackedID(idConnection) <> 0 Then
                 If publicDebugMode = True Then
                     aRes = SendLogSystemMessageToClient(idConnection, "Creature ID #" & CStr(lastAttackedID(idConnection)) & _
                       " ( " & GetNameFromID(idConnection, lastAttackedID(idConnection)) & _
                       " ) will be ignored (because too much attack:" & CStr(0 - lastHPchange(idConnection)) & ")")
                 End If
                  aRes = AddIgnoredcreature(idConnection, lastAttackedID(idConnection))
                  aRes = MeleeAttack(idConnection, 0, True)
                  lastAttackedID(idConnection) = 0
              End If
           End If
        End If
      End If
    Case &HA1
      ' my skills
      pos = pos + 15
      If TibiaVersionLong >= 872 Then
'    WriteByte(0xA1);
'    WriteByte(*PLAYER_FIST);
'    WriteByte(*PLAYER_FIST_PER);
'    WriteByte(0); // unknown - flash client
'    WriteByte(*PLAYER_CLUB);
'    WriteByte(*PLAYER_CLUB_PER);
'    WriteByte(0); // unknown - flash client
'    WriteByte(*PLAYER_SWORD);
'    WriteByte(*PLAYER_SWORD_PER);
'    WriteByte(0); // unknown - flash client
'    WriteByte(*PLAYER_AXE);
'    WriteByte(*PLAYER_AXE_PER);
'    WriteByte(0); // unknown - flash client
'    WriteByte(*PLAYER_DIST);
'    WriteByte(*PLAYER_DIST_PER);
'    WriteByte(0); // unknown - flash client
'    WriteByte(*PLAYER_SHIELD);
'    WriteByte(*PLAYER_SHIELD_PER);
'    WriteByte(0); // unknown - flash client
'    WriteByte(*PLAYER_FISH);
'    WriteByte(*PLAYER_FISH_PER);
'    WriteByte(0); // unknown - flash client

        pos = pos + 7
      End If
      
      
    If TibiaVersionLong >= 1035 Then
' new skills ??
        pos = pos + 14
      End If

    'Case &H20
      ' unknown, 1 byte, usually 20 0D
    '  pos = pos + 2
    Case &HA2
      ' add 1 status (lock,skull,etc)
      ' A2 00 40 - pz
      ' A2 00 00 - nothing
      'Debug.Print frmMain.showAsStr2(packet, True)
      If TibiaVersionLong >= 773 Then
        StatusBits(idConnection) = ByteToBitstring(packet(pos + 1)) & ByteToBitstring(packet(pos + 2))
        'aRes = SendLogSystemMessageToClient(idConnection, "Status changed> " & StatusBits(idConnection))
        'Debug.Print StatusBits(idConnection)
        pos = pos + 3 ' NEW STRANGE THING FOR STATUS (POISONED, PZLOCKED, ETC)
      Else
        StatusBits(idConnection) = ByteToBitstring(packet(pos + 1)) & "00000000"
        'aRes = SendLogSystemMessageToClient(idConnection, "Status changed> " & StatusBits(idConnection))
        pos = pos + 2
      End If
    Case &HA3
      ' stop attack !
      pos = pos + 1
      If TibiaVersionLong >= 860 Then
          'Debug.Print frmMain.showAsStr3(packet, True, pos, pos + 3)
          templ1 = FourBytesLong(packet(pos), packet(pos + 1), packet(pos + 2), packet(pos + 3))
          If packet(pos + 3) < &HFF Then
           'Debug.Print "N=" & templ1
            FixRightNumberOfClicks idConnection, templ1
          Else
            LogOnFile "errors.txt", "WARNING: Dangerous bytes in packet A3: " & _
            GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1)) & " " & GoodHex(packet(pos + 2)) & " " & GoodHex(packet(pos + 3)) & _
            " Forced pause of cheats. DEBUG: " & RightNumberOfClicksJUSTREAD(idConnection)
           aRes = GiveGMmessage(idConnection, "RELOG ASAP. Don't activate the cavebot again until you relog. Else you risk for a sure ban. And please send errors.txt to daniel@blackdtools.com", "Blackd Proxy")
           CheatsPaused(idConnection) = True
           ChangePlayTheDangerSound True
           DoEvents
          End If
          pos = pos + 4 ' current number of clicks so far

        'Debug.Print CLng(packet(pos - 4))
      End If
    Case &HA4
      ' new, related to spell cooldown
      pos = pos + 6
    Case &HA5
      ' new, related to spell cooldown
      pos = pos + 6
    Case &HA6
      ' new since Tibia 9.5
      ' used life potion: A6 E8 03 00 00
      
      ' E8 03 = special effect, tile of 2 bytes.
      pos = pos + 3
      ' now it seems to be always 00 00, possible future string here?
      lonN = GetTheLong(packet(pos), packet(pos + 1))
      pos = pos + 2 + lonN
      
    Case &HA7
      ' new since Tibia 9.9 - pvp modes?
      ' example: A7 01 01 01 00
      pos = pos + 5

    Case &HAA
 
 
 

       ' chat - eval to skip enough bytes
      If TibiaVersionLong > 760 Then
        pos = pos + 4 'skip 4 strange bytes (always 00 00 00 00 )
      End If
      lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
      pos = pos + 3

      nameofgivenID = ""
      For itemCount = 1 To lonN
        nameofgivenID = nameofgivenID & Chr(packet(pos))
        pos = pos + 1
      Next itemCount
      If (nameofgivenID = CharacterName(idConnection)) Then
        itsMe = True
      Else
        itsMe = False
      End If
      blnDebug1 = False 'for debug

      If TibiaVersionLong >= 773 Then
        ' NEW : level of the person who is talking
        'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & _
         "TALKING > " & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1))
        templ1 = GetTheLong(packet(pos), packet(pos + 1))
        If templ1 = 0 Then ' npc is talking
            doingTrade(idConnection) = True
        End If
        pos = pos + 2 'skip level
      End If
      subType = CLng(packet(pos))
      Select Case subType
'      Case newchatmessage_H22
'        ' new since Tibia 8.72 ' eat food
'        ' AA 00 00 00 00 05 00 61 20 63 61 74 00 00 22 42 7D BB 7D 07 05 00 4D 65 6F 77 21
'        pos = pos + 6 ' skip subtype and x,y,z
'        lonN = GetTheLong(packet(pos), packet(pos + 1))
'        pos = pos + 2
'        tmpStr = ""
'        For itemCount = 1 To lonN
'          tmpStr = tmpStr & Chr(packet(pos))
'          pos = pos + 1
'        Next itemCount
'        'Debug.Print "msg: " & tmpStr
'        If (itsMe = False) Then
'          var_lastsender(idConnection) = nameofgivenID
'          var_lastmsg(idConnection) = tmpStr
'          ProcessEventMsg idConnection, subType
'          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
'            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
'          End If
'        End If
'        If (GotPacketWarning(idConnection) = False) And (RuneMakerOptions(idConnection).msgSound = True) And _
'         (itsMe = False) And (CheatsPaused(idConnection) = False) Then
'          PlayMsgSound = True
'        End If
      Case newchatmessage_H9
        ' new since Tibia 8.72 ' use hotkey

        pos = pos + 6 ' skip subtype and x,y,z
        lonN = GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 2
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        'Debug.Print "msg: " & tmpStr
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
          End If
        End If
        If (GotPacketWarning(idConnection) = False) And (RuneMakerOptions(idConnection).msgSound = True) And _
         (itsMe = False) And (CheatsPaused(idConnection) = False) Then
          PlayMsgSound = True
        End If
        
        Case verynewchatmessage_HB
        ' new since Tibia 10.36 . rookgard npcs?
'AA 00 00 00 00 06 00 41 6C 20 44 65 65 00 00 0A 3F 7D B3 7D 07 CD 00 48 65 6C 6C 6F 2C 20 68 65 6C 6C 6F 2C 20 41 72 61 6E 74 61 20 41 6E 74 65 79 21 20 50 6C 65 61 73 65 20 63 6F 6D 65 20 69 6E 2C
'20 6C 6F 6F 6B 2C 20 61 6E 64 20 62 75 79 21 20 49 27 6D 20 61 20 73 70 65 63 69 61 6C 69 73 74 20 66 6F 72 20 61 6C 6C 20 73 6F 72 74 73 20 6F 66 20 7B 74 6F 6F 6C 73 7D 2E 20 4A 75 73 74 20 61 73 6B 20 6D 65 20 66 6F 72 20 61 20 7B 74 72 61 64 65 7D 20 74 6F 20 73 65 65 20 6D 79 20 6F 66 66 65 72 73 21 20 59 6F 75 20 63 61 6E 20 61 6C 73 6F 20 61 73 6B 20 6D 65 20 66 6F 72 20 67 65 6E 65 72 61 6C 20 7B 68 69 6E 74 73 7D 20 61 62 6F 75 74 20 74 68 65 20 67 61 6D 65 2E 20 2E 2E 2E AA 00 00 00 00 06 00 41 6C 20 44 65 65 00 00 0B 35 00 59 6F 75 20 63 61 6E 20 61 6C 73 6F 20 61 73 6B 20 6D 65 20 61 62 6F 75 74 20 65 61 63 68 20 7B 63 69 74 69 7A 65 6E 7D 20 6F 66 20 74 68 65 20 69 73 6C 65 2E 93 9B 30 01 40 01 00 92 9B 30 01 40 01 8C B2 48 C8 01 3B 83 3D 7D B3 7D 07 01 6A 3D 7D B3 7D 07 01 49 0B FF 05 B4 17 3D 7D B3
'7D 07 03 00 00 00 B4 00 00 00 00 00 2F 00 59 6F 75 20 6C 6F 73 65 20 33 20 68 69 74 70 6F 69 6E 74 73 20 64 75 65 20 74 6F 20 61 6E 20 61 74 74 61 63 6B 20 62 79 20 61 20 72 61 74 2E 8C B2 48 C8 01 3B 6B 3D 7D B2 7D 07 01 63 00 9B 30 01 40 02 01 A0 6B 00 B4 00 9A 4C 00 00 B0 B3 00 00 A4 0D 00 00 00 00 00 00 07 00 37 23 00 23 00 00 00 00 64 D8 09 74 00 00 00 D0 02
        pos = pos + 1
        lonN = GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 2
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        'Debug.Print "msg: " & tmpStr
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
          End If
        End If
        If (GotPacketWarning(idConnection) = False) And (RuneMakerOptions(idConnection).msgSound = True) And _
         (itsMe = False) And (CheatsPaused(idConnection) = False) Then
          PlayMsgSound = True
        End If
        
        
      Case newchatmessage_HA
        ' new since Tibia 8.72 ' eat food
        ' AA 00 00 00 00 06 00 57 69 6C 6C 69 65 00 00 0A 45 7D CC 7D 07 31 00 48 69 68 6F 20 41 72 61 6E 74 61 20 41 6E 74 65 79 2E
        ' 20 49 20 68 6F 70 65 20 79 6F 75 27 72 65 20 68 65 72 65 20 74 6F 20 7B 74 72 61 64 65 7D 2E
        pos = pos + 6 ' skip subtype and x,y,z
        lonN = GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 2
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        'Debug.Print "msg: " & tmpStr
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
          End If
        End If
        If (GotPacketWarning(idConnection) = False) And (RuneMakerOptions(idConnection).msgSound = True) And _
         (itsMe = False) And (CheatsPaused(idConnection) = False) Then
          PlayMsgSound = True
        End If

      Case oldmessage_H14
        ' unknown
        lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
        pos = pos + 3
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
        End If
        If (GotPacketWarning(idConnection) = False) And (RuneMakerOptions(idConnection).msgSound = True) And _
         (itsMe = False) And (CheatsPaused(idConnection) = False) Then
          PlayMsgSound = True
        End If
      Case oldmessage_H15
        ' say
        lonN = GetTheLong(packet(pos + 6), packet(pos + 7))
        pos = pos + 8
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
          End If
        End If
        If (GotPacketWarning(idConnection) = False) And (RuneMakerOptions(idConnection).msgSound = True) And _
         (itsMe = False) And (CheatsPaused(idConnection) = False) Then
          PlayMsgSound = True
        End If
      Case oldmessage_H0
        ' unknown
        lonN = GetTheLong(packet(pos + 6), packet(pos + 7))
        pos = pos + 8
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
          End If
        End If
        If (GotPacketWarning(idConnection) = False) And (RuneMakerOptions(idConnection).msgSound = True) And _
         (itsMe = False) And (CheatsPaused(idConnection) = False) Then
          PlayMsgSound = True
        End If
      Case oldmessage_H1
        ' say
        lonN = GetTheLong(packet(pos + 6), packet(pos + 7))
        pos = pos + 8
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
          End If
        End If
        If (GotPacketWarning(idConnection) = False) And (RuneMakerOptions(idConnection).msgSound = True) And _
         (itsMe = False) And (CheatsPaused(idConnection) = False) Then
          PlayMsgSound = True
        End If
      Case oldmessage_H2
        ' whisper
        lonN = GetTheLong(packet(pos + 6), packet(pos + 7))
        pos = pos + 8
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
          End If
        End If
        If (GotPacketWarning(idConnection) = False) And (RuneMakerOptions(idConnection).msgSound = True) And _
         (itsMe = False) And (CheatsPaused(idConnection) = False) Then
          PlayMsgSound = True
        End If
      Case oldmessage_H3
        ' yell
        lonN = GetTheLong(packet(pos + 6), packet(pos + 7))
        pos = pos + 8
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
          End If
        End If
        If (GotPacketWarning(idConnection) = False) And (RuneMakerOptions(idConnection).msgSound = True) And _
         (itsMe = False) And (CheatsPaused(idConnection) = False) Then
          PlayMsgSound = True
        End If
      Case oldmessage_H4
        ' tell
        lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
        pos = pos + 3
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
          End If
        End If
        If (GotPacketWarning(idConnection) = False) And (RuneMakerOptions(idConnection).msgSound = True) And _
         (itsMe = False) And (CheatsPaused(idConnection) = False) Then
          PlayMsgSound = True
        End If
      Case oldmessage_H5
        ' channel
        lonN = GetTheLong(packet(pos + 3), packet(pos + 4))
        If blnDebug1 = False Then
          tempb1 = packet(pos + 1)
          tempb2 = packet(pos + 2)
          lastRecChannelID(idConnection) = GoodHex(tempb1) & " " & GoodHex(tempb2)
          'here tempb1 and tempb2 contains the id of the channel
          msg = ""
          For itemCount = 1 To lonN
             msg = msg & Chr(packet(pos + 4 + itemCount))
          Next itemCount
        
         If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = msg
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
          End If
        End If
          
          ' here msg will contain the text said in the channel
          If (tempb1 = getSpamChannelB1(idConnection)) And _
           (tempb2 = getSpamChannelB2(idConnection)) Then
            If Left(msg, 1) = "@" Then
              frmMapReader.AddPlayerToBigMap nameofgivenID, Right(msg, Len(msg) - 1)
            End If
          End If
        
          If (frmHardcoreCheats.chkAcceptSDorder.Value = 1) And _
           (sentWelcome(idConnection) = True) And _
           (GotPacketWarning(idConnection) = False) Then
          lonO = Len(frmHardcoreCheats.txtOrder)
          If (lonN > lonO) Then
            If (Left(msg, lonO) = frmHardcoreCheats.txtOrder) And _
             ((frmHardcoreCheats.txtRemoteLeader.Text = "") Or _
              (LCase(frmHardcoreCheats.txtRemoteLeader.Text) = LCase(nameofgivenID))) Then
             
              rightpart = Right(msg, (Len(msg) - lonO - 1))
              aRes = SendLogSystemMessageToClient(idConnection, "Received remote order. Casting " & _
               frmHardcoreCheats.cmbOrderType.Text & " on " & rightpart)
              DoEvents
              Select Case frmHardcoreCheats.cmbOrderType.ListIndex
              Case 0
                aRes = SendAimbot(rightpart, idConnection, LowByteOfLong(tileID_SD), HighByteOfLong(tileID_SD))
                If frmHardcoreCheats.chkColorEffects.Value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              Case 1
                aRes = SendAimbot(rightpart, idConnection, LowByteOfLong(tileID_HMM), HighByteOfLong(tileID_HMM))
                If frmHardcoreCheats.chkColorEffects.Value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              Case 2
                aRes = SendAimbot(rightpart, idConnection, LowByteOfLong(tileID_Explosion), HighByteOfLong(tileID_Explosion))
                If frmHardcoreCheats.chkColorEffects.Value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              Case 3
                aRes = SendAimbot(rightpart, idConnection, LowByteOfLong(tileID_IH), HighByteOfLong(tileID_IH))
                If frmHardcoreCheats.chkColorEffects.Value = 1 Then
                  nextLight(idConnection) = "04"
                  enLight idConnection
                End If
              Case 4
                aRes = SendAimbot(rightpart, idConnection, LowByteOfLong(tileID_UH), HighByteOfLong(tileID_UH))
                If frmHardcoreCheats.chkColorEffects.Value = 1 Then
                  nextLight(idConnection) = "04"
                  enLight idConnection
                End If
              Case 5
                aRes = SendMobAimbot(rightpart, idConnection, LowByteOfLong(tileID_SD), HighByteOfLong(tileID_SD))
                If frmHardcoreCheats.chkColorEffects.Value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              Case 6
                aRes = SendMobAimbot(rightpart, idConnection, LowByteOfLong(tileID_HMM), HighByteOfLong(tileID_HMM))
                If frmHardcoreCheats.chkColorEffects.Value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              Case 7
                aRes = SendMobAimbot(rightpart, idConnection, LowByteOfLong(tileID_Explosion), HighByteOfLong(tileID_Explosion))
                If frmHardcoreCheats.chkColorEffects.Value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              Case 8
                aRes = SendMobAimbot(rightpart, idConnection, LowByteOfLong(tileID_IH), HighByteOfLong(tileID_IH))
                If frmHardcoreCheats.chkColorEffects.Value = 1 Then
                  nextLight(idConnection) = "04"
                  enLight idConnection
                End If
              Case 9
                aRes = SendMobAimbot(rightpart, idConnection, LowByteOfLong(tileID_UH), HighByteOfLong(tileID_UH))
                If frmHardcoreCheats.chkColorEffects.Value = 1 Then
                  nextLight(idConnection) = "04"
                  enLight idConnection
                End If
              Case 10 'type A
                aRes = ExecuteInTibia(rightpart, idConnection, True)
                
              Case 11 'type B
                aRes = SendMobAimbot(rightpart, idConnection, LowByteOfLong(tileID_fireball), HighByteOfLong(tileID_fireball))
                If frmHardcoreCheats.chkColorEffects.Value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              Case 12 'type C
                aRes = SendMobAimbot(rightpart, idConnection, LowByteOfLong(tileID_stalagmite), HighByteOfLong(tileID_stalagmite))
                If frmHardcoreCheats.chkColorEffects.Value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
              Case 13 'type D
                aRes = SendMobAimbot(rightpart, idConnection, LowByteOfLong(tileID_icicle), HighByteOfLong(tileID_icicle))
                If frmHardcoreCheats.chkColorEffects.Value = 1 Then
                  nextLight(idConnection) = "FD"
                  enLight idConnection
                End If
                
              End Select
              DoEvents
            End If
          End If
          End If
        End If

        pos = pos + 5 + lonN
      Case oldmessage_H6
        ' to channel by counsellor
        ' strange case 7.6 : 29 00 AA 03 00 41 63 65 06 00 00 00 00 1C 00 4D 61 6D 20 42 4F 74 74 65 72 61 20 21 21 21 20 77 79 73 6F 6B 69 20 6C 76 20 21 21
        'lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
        'If lonN = 0 Then ' strange case
            pos = pos + 4
            lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
        'End If
        pos = pos + 3
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
          End If
        End If
      Case oldmessage_H7
        ' counsellor private message / OTserver broadcast
        lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
        pos = pos + 3
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
'        If (itsMe = False) Then
'          var_lastsender(idConnection) = nameofgivenID
'          var_lastmsg(idConnection) = tmpStr
'          ProcessEventMsg idConnection, subType
'          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
'            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
'          End If
'        End If
      Case newmessage_H8
        ' party loot ' new since Tibia 8.4
        lonN = GetTheLong(packet(pos + 3), packet(pos + 4))
        pos = pos + 5
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        'aRes = GiveGMmessage(idConnection, tmpStr, "Debugmsg")
      Case oldmessage_H9
        ' to channel by gm (red)
        lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
        pos = pos + 3
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection), True
          End If
        End If
      Case oldmessage_HA
        ' gm private message
        lonN = GetTheLong(packet(pos + 3), packet(pos + 4))
        pos = pos + 5
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection), True
          End If
        End If
      Case oldmessage_HB 'strangeGM
        lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
        pos = pos + 3
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection), True
          End If
        End If
      Case oldmessage_HC
        ' to channel by tutor (orange)
        lonN = GetTheLong(packet(pos + 3), packet(pos + 4))
        pos = pos + 5
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
          End If
        End If
      'case &HC 'unknown0C
      'case &HD 'unknown0D
      Case oldmessage_HE
       ' AA 06 00 45 72 72 6F 72 72 0E 08 00 33 00 59 6F 75 20 6C 6F 73 65 20 31 34 38 20 68 69 74 70 6F 69 6E 74 73 20 64 75 65 20 74 6F 20 61 6E 20 61 74 74 61 63 6B 20 62 79 20 61 20 64 65 6D 6F 6E 2E
        ' Error message ?
        lonN = GetTheLong(packet(pos + 3), packet(pos + 4))
        pos = pos + 5
        tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
          End If
        End If
       
      'case &HF 'unknown0F
      Case oldmessage_H10
        ' monster say
        lonN = GetTheLong(packet(pos + 6), packet(pos + 7))
        pos = pos + 8
         tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
        End If
      Case oldmessage_H11
        If TibiaVersionLong > 860 Then
            ' monster yell 2 'ok in Tibia 8.70
            lonN = GetTheLong(packet(pos + 6), packet(pos + 7))
            pos = pos + 8
            tmpStr = ""
            For itemCount = 1 To lonN
              tmpStr = tmpStr & Chr(packet(pos))
              pos = pos + 1
            Next itemCount
            If (itsMe = False) Then
              var_lastsender(idConnection) = nameofgivenID
              var_lastmsg(idConnection) = tmpStr
              ProcessEventMsg idConnection, subType
            End If
        Else
            ' monster yell
            lonN = GetTheLong(packet(pos + 6), packet(pos + 7))
            pos = pos + 8
            tmpStr = ""
            For itemCount = 1 To lonN
              tmpStr = tmpStr & Chr(packet(pos))
              pos = pos + 1
            Next itemCount
            If (itsMe = False) Then
              var_lastsender(idConnection) = nameofgivenID
              var_lastmsg(idConnection) = tmpStr
              ProcessEventMsg idConnection, subType
            End If
        End If
      Case &H10
        ' new monster say ?? ' detected since blackd proxy 12.8
        lonN = GetTheLong(packet(pos + 3), packet(pos + 4))
        pos = pos + 5
         tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
          End If
        End If
      Case &H11
        ' someone joins the guild
        ' fixed since Blackd Proxy 15.3
        ' 37 00 AA 00 00 00 00 00 00 63 01 11 00 00 29 00 44 69 65 65 65 20 68 61 73 20 69 6E 76 69 74 65 64 20 53 69 72 20 42 75 72 65 6E 20 74 6F 20 74 68 65 20 67 75 69 6C 64 2E
        lonN = GetTheLong(packet(pos + 3), packet(pos + 4))
        pos = pos + 5
         tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
          End If
        End If
      Case &H25
        ' 1B 00 AA 00 00 00 00 06 00 61 20 77 79 72 6D 00 00 25 8D 7E F8 7F 03 04 00 47 52 52 52
        ' new since Tibia 10.55
              lonN = GetTheLong(packet(pos + 6), packet(pos + 7))
        pos = pos + 8
         tmpStr = ""
        For itemCount = 1 To lonN
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        If (itsMe = False) Then
          var_lastsender(idConnection) = nameofgivenID
          var_lastmsg(idConnection) = tmpStr
          ProcessEventMsg idConnection, subType
          If (frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1) Then
            CheckIfGM idConnection, nameofgivenID, myZ(idConnection)
          End If
        End If
      Case Else
         debugChain = debugChain & " at message subtype -> [" & GoodHex(subType) & "] "
        pos = pos + 10000
      End Select
 
 
 
 
 
 
 
    


    Case &HAB
       ' list of channels
       numC = CLng(packet(pos + 1))
       pos = pos + 2
       For itemCount = 1 To numC
         pos = pos + 2 ' skip chat ID
         lonN = GetTheLong(packet(pos), packet(pos + 1))
         pos = pos + 2 + lonN
       Next itemCount
    Case &HAC
       ' add channel
       ' pre 872: 14 00 AC 05 00 0B 00 41 64 76 65 72 74 69 73 69 6E 67 00 00 00 00
       
       
      ' post
' AC 00 00
' 13 00 49 6C 6C 75 73 69 6F 6E 73 20 47 61 74 68 65 72 69 6E 67
' 05 00
' 10 00 4C 61 64 79 20 4D 61 72 79 20 57 61 6C 6B 65 72
' 0B 00 52 6F 78 69 65 27 44 65 6D 6F 6E
' 08 00 49 6E 66 61 6E 74 72 79
' 09 00 4D 69 6E 75 6E 69 6E 68 61
' 0F 00 46 65 6E 69 72 20 64 65 20 41 72 69 6F 74 6F
' 01 00
' 10 00 43 61 68 68 7A 69 6E 68 61 20 6F 66 20 55 6E 69

      ' Debug.Print ">> " & frmMain.showAsStr3(packet, True, pos, 10000)
       lonN = GetTheLong(packet(pos + 3), packet(pos + 4))
        pos = pos + 5 + lonN
       If TibiaVersionLong >= 872 Then
            numC = GetTheLong(packet(pos), packet(pos + 1))
            pos = pos + 2
            For itemCount = 1 To numC
                lonN = GetTheLong(packet(pos), packet(pos + 1))
                pos = pos + 2 + lonN
            Next itemCount
            numC = GetTheLong(packet(pos), packet(pos + 1))
            pos = pos + 2
            For itemCount = 1 To numC
                lonN = GetTheLong(packet(pos), packet(pos + 1))
                pos = pos + 2 + lonN
            Next itemCount
       End If
       
    Case &HAD
      ' remove name from vip list
      lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
      pos = pos + 3 + lonN
    Case &HA9
      ' ?
      ' UNDER INVESTIGATION - AUTOBAN
      nameofgivenID = "[WARNING!] Request to search cheats? Please report this log with timestamp to daniel@blackdtools.com .Strange packet received: " & GoodHex(packet(pos)) & " " & GoodHex(packet(pos + 1))
      LogOnFile "errors.txt", nameofgivenID
      aRes = SendLogSystemMessageToClient(idConnection, nameofgivenID)
      DoEvents
      pos = pos + 2
    Case &HAF
       ' reported name??
       lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
        pos = pos + 3 + lonN
    Case &HB1
      ' close report order
       pos = pos + 1
    Case &HB2
      ' open private channel / guild channel
      If TibiaVersionLong >= 1070 Then
          ' channel title
          lonN = GetTheLong(packet(pos + 3), packet(pos + 4))
          pos = pos + 5 + lonN
          ' number of players at channel
          lonN = GetTheLong(packet(pos), packet(pos + 1))
          pos = pos + 2
          ' list of players at channel
          For itemCount = 1 To lonN
             templ1 = GetTheLong(packet(pos), packet(pos + 1))
             pos = pos + 2 + templ1
          Next itemCount
          ' number of special hidden? players
          lonN = GetTheLong(packet(pos), packet(pos + 1))
          pos = pos + 2
          ' somehow special hidden? players, usually 00 00, else we have a list below:
          For itemCount = 1 To lonN
             templ1 = GetTheLong(packet(pos), packet(pos + 1))
             pos = pos + 2 + templ1
          Next itemCount
      Else
        lonN = GetTheLong(packet(pos + 3), packet(pos + 4))
        pos = pos + 5 + lonN
        If (TibiaVersionLong >= 944) Then
          pos = pos + 4
        End If
      End If
    Case &HB3
      ' close private channel

        pos = pos + 3

    Case &HB4
      ' system message (msgtype,lenght and message)
      If (TibiaVersionLong >= 1055) Then
            tempb1 = packet(pos + 1)
        Select Case tempb1
        Case &H16
          ' 26 00 B4 16 22 00 59 6F 75 20 73 65 65 20 61 20 73 74 61 6D 70 65 64 20 70 61 72 63 65 6C 20 28 56 6F 6C 3A 31 30 29 2E
               

            pos = pos + 2
            lonN = GetTheLong(packet(pos), packet(pos + 1)) ' 36 00
            pos = pos + 2
            mobName = ""
            For itemCount = 1 To lonN
                mobName = mobName & Chr(packet(pos))
                pos = pos + 1
            Next itemCount
        Case &H17, &H18, &H1B
                ' B4 17 C5 7E E1 7D 07 1A 00 00 00 1E 00 00 00 00 00 36 00 41 20 70 6F 69 73 6F 6E 20 73 70 69 64 65 72 20 6C 6F 73 65 73 20 32 36 20 68 69 74 70 6F 69 6E 74 73 20 64 75 65 20 74 6F 20 79 6F 75 72 20 61 74 74 61 63 6B 2E
                ' B4 18 AD 7E 5B 7D 0F 19 00 00 00 B4 00 00 00 00 00 3B 00 59 6F 75 20 6C 6F 73 65 20 32 35 20 68 69 74 70 6F 69 6E 74 73 20 64 75 65 20 74 6F 20 61 6E 20 61 74 74 61 63 6B 20 62 79 20 61 20 6D 69 6E 6F 74 61 75 72 20 67 75 61 72 64 2E
               
                pos = pos + 2
                pos = pos + 5 ' x,y,z
                pos = pos + 4 ' ?
                mobName = ""
                pos = pos + 6
               
                lonN = GetTheLong(packet(pos), packet(pos + 1)) ' 36 00
                pos = pos + 2
                mobName = ""
                For itemCount = 1 To lonN
                    mobName = mobName & Chr(packet(pos))
                    pos = pos + 1
                Next itemCount
                
                
        Case &H19, &H1A, &H1C, &H1D
              ' B4 1D AE 7E 5C 7D 0F 32 00 00 00 D7 2D 00 41 20 6D 69 6E 6F 74 61 75 72 20 67 75 61 72 64 20 67 61 69 6E 65 64 20 35 30 20 65 78 70 65 72 69 65 6E 63 65 20 70 6F 69 6E 74 73 2E
            
              
              'Debug.Print frmMain.showAsStr3(packet, True, pos, pos + 10000)
                pos = pos + 2
                pos = pos + 5 ' x,y,z
                pos = pos + 4 ' ?
              lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
              pos = pos + 3
              mobName = ""
        
              For itemCount = 0 To lonN - 1
                mobName = mobName & Chr(packet(pos))
                pos = pos + 1
              Next itemCount
     
        Case &H6, &H21, &H22
                   ' B4 21 37 27 26 00 47 75 69 6C 64 20 6D 65 73 73 61 67 65 3A 20 48 61 69 6C 20 54 69 62 69 61 6E 6F 73 20 56 69 63 69 61 64 6F 73 21
                   ' B4 22 30 4E 56 00 41 6C 67 61 74 61 20 53 74 6F 72 20 68 61 73 20 62 65 65 6E 20 69 6E 76 69 74 65 64 2E 20 4F 70 65 6E 20 74 68 65 20 70 61 72 74 79 20 63 68 61 6E 6E 65 6C 20 74 6F 20 63 6F 6D 6D 75 6E 69 63 61 74 65 20 77 69 74 68 20 79 6F 75 72 20 6D 65 6D 62 65 72 73 2E 91 FA D5 7B 02 04 90 FA D5 7B 02 00 94 FA D5 7B 02 02 00 91 0A F7 DC 02 02 94 0A F7 DC 02 00 00 91 FA D5 7B 02 04 94 FA D5 7B 02 02 00
                  
                    'Debug.Print frmMain.showAsStr3(packet, True, pos, 1000000)
                    pos = pos + 4
                    lonN = GetTheLong(packet(pos), packet(pos + 1))
                    pos = pos + 2
                    mobName = ""
                    For itemCount = 1 To lonN
                        mobName = mobName & Chr(packet(pos))
                        pos = pos + 1
                    Next itemCount
        Case Else
                         'Debug.Print frmMain.showAsStr3(packet, True, pos, 1000000)
                    pos = pos + 2
                    lonN = GetTheLong(packet(pos), packet(pos + 1))
                    pos = pos + 2
                    mobName = ""
                    For itemCount = 1 To lonN
                        mobName = mobName & Chr(packet(pos))
                        pos = pos + 1
                    Next itemCount
              
        End Select

        If (tempb1 = &H15) Then
          lastIngameCheck(idConnection) = mobName
        End If
        If (tempb1 = &H14) Then
           If mobName = "You are not the owner." Then
             lootTimeExpire(idConnection) = 0 ' 5 seconds loot timer forced to expire
           End If
        End If
      ElseIf (TibiaVersionLong >= 1036) Then
         
         
         
               tempb1 = packet(pos + 1)
        If ((tempb1 >= &H16) And (tempb1 <= &H1C)) Then
            ' B4 15 37 7D A5 7D 09 03 00 00 00
            'Debug.Print frmMain.showAsStr3(packet, True, pos, 1000000)
            pos = pos + 2
            pos = pos + 5 ' x,y,z
            pos = pos + 4 ' ?
            mobName = ""
            ' B4 09 00 00 00 C6 30 00 41 20 63 ....
            Select Case tempb1
            Case &H16, &H17, &H1A
                ' B4 17 C5 7E E1 7D 07 1A 00 00 00 1E 00 00 00 00 00 36 00 41 20 70 6F 69 73 6F 6E 20 73 70 69 64 65 72 20 6C 6F 73 65 73 20 32 36 20 68 69 74 70 6F 69 6E 74 73 20 64 75 65 20 74 6F 20 79 6F 75 72 20 61 74 74 61 63 6B 2E
                ' B4 1A C4 7E E2 7D 07 24 00 00 00 D7 20 00 59 6F 75 20 67 61 69 6E 65 64 20 33 36 20 65 78 70 65 72 69 65 6E 63 65 20 70 6F 69 6E 74 73 2E

                tempb2 = packet(pos)
          
                  pos = pos + 6
              
                lonN = GetTheLong(packet(pos), packet(pos + 1)) ' 36 00
                pos = pos + 2
                mobName = ""
                For itemCount = 1 To lonN
                    mobName = mobName & Chr(packet(pos))
                    pos = pos + 1
                Next itemCount
            Case &H18, &H19, &H1B, &H1C
              ' B4 17 D0 7E 4B 7D 07 4E 00 00 00 5F 25 00 59 6F 75 20 68 65 61 6C 65 64 20 79 6F 75 72 73 65 6C 66 20 66 6F 72 20 37 38 20 68 69 74 70 6F 69 6E 74 73 2E 83 D0 7E 4B 7D 07 0D A4 02 E8 03 00 00 A5 02 E8 03 00 00 AA 00 00 00 00 0B 00 42 6C 61 63 6B 79 20 4A 61 6B 65 28 00 09 D0 7E 4B 7D 07 0A 00 65 78 75 72 61 20 67 72 61 6E A0 F9 01 F9 01 3A 93 00 00 98 B1 01 00 8A 53 0E 00 00 00 00 00 28 00 1C BD 01 03 02 0C 0C 57 64 C4 09 2A 01 9B 00
              'Debug.Print frmMain.showAsStr3(packet, True, pos, pos + 10000)
              lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
              pos = pos + 3
              mobName = ""
        
              For itemCount = 0 To lonN - 1
                mobName = mobName & Chr(packet(pos))
                pos = pos + 1
              Next itemCount
            End Select
            
        Else
                If (tempb1 = &H6) Then
                    'Debug.Print frmMain.showAsStr3(packet, True, pos, 1000000)
                    pos = pos + 4
                    lonN = GetTheLong(packet(pos), packet(pos + 1))
                    pos = pos + 2
                    mobName = ""
                    For itemCount = 1 To lonN
                        mobName = mobName & Chr(packet(pos))
                        pos = pos + 1
                    Next itemCount
                Else
                    'Debug.Print frmMain.showAsStr3(packet, True, pos, 1000000)
                    pos = pos + 2
                    lonN = GetTheLong(packet(pos), packet(pos + 1))
                    pos = pos + 2
                    mobName = ""
                    For itemCount = 1 To lonN
                        mobName = mobName & Chr(packet(pos))
                        pos = pos + 1
                    Next itemCount
                End If
        End If
        If (tempb1 = &H15) Then
          lastIngameCheck(idConnection) = mobName
        End If
        If (tempb1 = &H14) Then
           If mobName = "You are not the owner." Then
             lootTimeExpire(idConnection) = 0 ' 5 seconds loot timer forced to expire
           End If
        End If
         
         
         
         
         
         
         
         
         
         
         
         
         
      ElseIf TibiaVersionLong >= 872 Then
        tempb1 = packet(pos + 1)
        If ((tempb1 >= &H15) And (tempb1 <= &H1A)) Or (tempb1 = &H1B) Then
            ' B4 15 37 7D A5 7D 09 03 00 00 00
            'Debug.Print frmMain.showAsStr3(packet, True, pos, 1000000)
            pos = pos + 2
            pos = pos + 5 ' x,y,z
            pos = pos + 4 ' ?
            mobName = ""
            ' B4 09 00 00 00 C6 30 00 41 20 63 ....
            Select Case tempb1
            Case &H15, &H16, &H19
                pos = pos + 2
                pos = pos + 4 ' id1?
                lonN = GetTheLong(packet(pos), packet(pos + 1))
                pos = pos + 2
                mobName = ""
                For itemCount = 1 To lonN
                    mobName = mobName & Chr(packet(pos))
                    pos = pos + 1
                Next itemCount
            Case &H17, &H18, &H1A, &H1B
              ' B4 17 D0 7E 4B 7D 07 4E 00 00 00 5F 25 00 59 6F 75 20 68 65 61 6C 65 64 20 79 6F 75 72 73 65 6C 66 20 66 6F 72 20 37 38 20 68 69 74 70 6F 69 6E 74 73 2E 83 D0 7E 4B 7D 07 0D A4 02 E8 03 00 00 A5 02 E8 03 00 00 AA 00 00 00 00 0B 00 42 6C 61 63 6B 79 20 4A 61 6B 65 28 00 09 D0 7E 4B 7D 07 0A 00 65 78 75 72 61 20 67 72 61 6E A0 F9 01 F9 01 3A 93 00 00 98 B1 01 00 8A 53 0E 00 00 00 00 00 28 00 1C BD 01 03 02 0C 0C 57 64 C4 09 2A 01 9B 00
              'Debug.Print frmMain.showAsStr3(packet, True, pos, pos + 10000)
              lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
              pos = pos + 3
              mobName = ""
        
              For itemCount = 0 To lonN - 1
                mobName = mobName & Chr(packet(pos))
                pos = pos + 1
              Next itemCount
            End Select
            
        Else
            If TibiaVersionLong >= 931 Then
                If (tempb1 = &H6) Then
                    'Debug.Print frmMain.showAsStr3(packet, True, pos, 1000000)
                    pos = pos + 4
                    lonN = GetTheLong(packet(pos), packet(pos + 1))
                    pos = pos + 2
                    mobName = ""
                    For itemCount = 1 To lonN
                        mobName = mobName & Chr(packet(pos))
                        pos = pos + 1
                    Next itemCount
                Else
                    'Debug.Print frmMain.showAsStr3(packet, True, pos, 1000000)
                    pos = pos + 2
                    lonN = GetTheLong(packet(pos), packet(pos + 1))
                    pos = pos + 2
                    mobName = ""
                    For itemCount = 1 To lonN
                        mobName = mobName & Chr(packet(pos))
                        pos = pos + 1
                    Next itemCount
                End If
            Else
                'Debug.Print frmMain.showAsStr3(packet, True, pos, 1000000)
                pos = pos + 2
                lonN = GetTheLong(packet(pos), packet(pos + 1))
                pos = pos + 2
                mobName = ""
                For itemCount = 1 To lonN
                    mobName = mobName & Chr(packet(pos))
                    pos = pos + 1
                Next itemCount
            End If
        End If
        If (tempb1 = &H14) Then
          lastIngameCheck(idConnection) = mobName
        End If
        If (tempb1 = &H13) Then
           If mobName = "You are not the owner." Then
             lootTimeExpire(idConnection) = 0 ' 5 seconds loot timer forced to expire
           End If
        End If
        'Debug.Print "B4 " & GoodHex(tempb1) & ": " & mobName
      Else
        'old
         ' system message (msgtype,lenght and message)
        tempb1 = packet(pos + 1)
        lonN = GetTheLong(packet(pos + 2), packet(pos + 3))
        pos = pos + 4
        mobName = ""
        For itemCount = 1 To lonN
          mobName = mobName & Chr(packet(pos))
          pos = pos + 1
        Next itemCount
        
        
        
        If ((TibiaVersionLong >= 860) And (TibiaVersionLong < 872) And (tempb1 = &H13)) Then
          lastIngameCheck(idConnection) = mobName
          'Debug.Print lastIngameCheck(idConnection)
        End If
        
        If ((TibiaVersionLong >= 840) And (TibiaVersionLong < 872) And (tempb1 = &H1A)) Then
           If mobName = "You are not the owner." Then
             lootTimeExpire(idConnection) = 0 ' 5 seconds loot timer forced to expire
           End If
        ElseIf (TibiaVersionLong >= 820) And (TibiaVersionLong < 872) And (tempb1 = &H19) Then
           If mobName = "You are not the owner." Then
             lootTimeExpire(idConnection) = 0 ' 5 seconds loot timer forced to expire
           End If
        ElseIf (TibiaVersionLong < 820) And (tempb1 = &H17) Then
           If mobName = "You are not the owner." Then
             lootTimeExpire(idConnection) = 0 ' 5 seconds loot timer forced to expire
           End If
        End If
        
        
        
      End If
      

      
      templ2 = 0
      If TibiaVersionLong >= 1055 Then
        If tempb1 = &H12 Then
            templ2 = 1
        End If
      ElseIf TibiaVersionLong >= 872 Then
        If tempb1 = &H11 Then
            templ2 = 1
        End If
      ElseIf TibiaVersionLong >= 841 Then
        If tempb1 = &H15 Then
            templ2 = 1
        End If
      ElseIf TibiaVersionLong >= 820 Then
        If tempb1 = &H14 Then
            templ2 = 1
        End If
      Else
        If tempb1 = &H12 Then
            templ2 = 1
        End If
      End If
      
      'Debug.Print GoodHex(tempb1) & ": message =" & mobName
      
      If (templ2 = 1) Then
               
        
        If ((TrainerOptions(idConnection).misc_dance_14min = 1) And _
            (CheatsPaused(idConnection) = False)) Then

          If mobName = serverLogoutMessage Then
           ' Debug.Print "DANCE OK"
            aRes = randomNumberBetween(0, 3)
            tmpStr = "exiva turn" & CStr(aRes)
            templ1 = GetTickCount() + randomNumberBetween(200, 800)
            AddSchedule idConnection, tmpStr, templ1
            aRes = aRes + 1
            tmpStr = "exiva turn" & CStr(aRes Mod 4)
            templ1 = templ1 + randomNumberBetween(400, 800)
            AddSchedule idConnection, tmpStr, templ1
            aRes = aRes + 1
            tmpStr = "exiva turn" & CStr(aRes Mod 4)
            templ1 = templ1 + randomNumberBetween(400, 800)
            AddSchedule idConnection, tmpStr, templ1
            aRes = aRes + 1
            tmpStr = "exiva turn" & CStr(aRes Mod 4)
            templ1 = templ1 + randomNumberBetween(400, 800)
            AddSchedule idConnection, tmpStr, templ1
          End If
        End If
      End If
      If ((TibiaVersionLong >= 820) And (tempb1 = &H16)) Or ((TibiaVersionLong < 820) And (tempb1 = &H14)) Then
        'BEING ATTACKED???
        If (sentFirstPacket(idConnection) = False) Then
        If TrialVersion = True Then
          If Len(mobName) > 7 Then
            If mobName = "You are permanently ignoring players." Then
              frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "[Debug] You are permanently ignoring players." & mobName
            ElseIf OTsType = 1 Then
              mobName = Mid(mobName, 32, 2) & ". " & Mid(mobName, 28, 3) & " " & Mid(mobName, 44, 4)
              receivedLogin(idConnection) = True
              blnTmp = CompareTibiaDate(mobName)
              LoginMsgCount(idConnection) = LoginMsgCount(idConnection) + 1
              If (blnTmp = False) Or (LoginMsgCount(idConnection) > 1) Then
                GotTrialLock = True
                lastLockReason = " Failed Trial check #1"
              End If
            ElseIf Left(mobName, 7) = "Welcome" Then
              OTsType = 1
            ElseIf Len(mobName) > 40 Then
              mobName = Mid(mobName, 27, 12)
              receivedLogin(idConnection) = True
              blnTmp = CompareTibiaDate(mobName)
              LoginMsgCount(idConnection) = LoginMsgCount(idConnection) + 1
              If (blnTmp = False) Or (LoginMsgCount(idConnection) > 1) Then
                GotTrialLock = True
                lastLockReason = " Failed Trial check #2"
              End If
            End If
            mobName = "...................................................................................."
            rightpart = "........"
          End If
        Else 'trialversion=false
          If Not ((trialSafety1 = 1) And (trialSafety2 = 2) And (trialSafety300 = 300)) Then
            End
          End If
        End If
        End If
      End If
    Case &HB5
      'cancel autowalk order
      pos = pos + 2
    Case &HB6
      ' unknown. Wall tile after it? new since tibia 8.71
      ' B6 DC 05
      pos = pos + 3
    Case &HB7
      ' new since tibia 10.51 preview
      
'B7
 ' 15 FF 00 FF
 ' 00 00 77 01 FF B4 15 FF 00 FF
 ' 00 00 7D 01 FF 00 FF
 ' 00 00 62 01 FF 00 FF
 ' 00 00 62 01 FF 00 FF
 ' 00 00 80 01 FF 00 FF
 ' 00 00 7C 01 FF 00 FF
 ' 00 00 65 00 FF 00 FF
 ' 00 00 65 00 FF 00 FF
 ' 00 00 65 00 FF 00 FF
 ' 00 00 65 00 FF 00 FF
 ' 00 00 65 00 FF 00 FF
 ' 00 00 59 16 FF 00 FF
 ' 00 00 79 01 FF 00 FF
 ' 00 00 7D 01 FF 00 FF
 ' 00 00 61 01 FF 42 07 FF 00 FF
 ' 00 00 62 01 FF 00 FF 00 00 63 01 FF 00 FF
 ' 00 00 5F 01 FF 50 0B FF 06 00 FF
 ' 00 00 63 01 FF 00 FF
 ' 00 00 80 01 FF 00 FF
 ' 00 00 65 00 FF
 ' 00 FF
 ' 00 00 79 01 FF B7 15 FF 00 FF
 ' 00 00 77 01 FF 00 FF
 ' 00 00 77 01 FF B4 15 FF 00 FF
 ' 00 00 77 01 FF B4 15 FF 00 FF
 ' 00 00 77 01 FF C8 00 FF
 ' 00 FF
 ' 00 00 7D 01 FF 8C 0A FF 00 FF
 ' 00 00 62 01 FF 00 FF
 ' 00 00 63 01 FF 00 FF
 ' 00 00 62 01 FF 00 FF
 ' 00 00 62 01 FF 00 FF
 ' 00 00 60 01 FF 00 FF
 ' 00 00 62 01 FF 00 FF
 ' 00 00 62 01 FF 00 FF
 ' 00 00 65 00 FF
 ' 00 FF 00 00 75 01 FF B8 00 FF 00
 ' FF 00 00 62 01 FF 4F 0B FF 06 8B 0A FF 00 FF 00 00 62 01 FF 00 FF 00 00 62 01 FF 00 FF 00 00 62 01 FF 00 FF 00 00 62 01 FF 00 FF 00 00 5F 01 FF 00 FF 00 00 62 01 FF 00 FF
 ' ... 00 00 60 01 FF 00 FF
 ' 00 00 62 01 FF 00 FF
 ' 00 00 62 01 FF 00 FF
 ' 00 00 62 01 FF 00 FF
 ' 00 00 63 01 FF 00 FF
 ' 00 00 65 00 FF 00 FF
 ' 00 00 75 01 FF B6 00 FF 00 FF
 ' 00 00 62 01 FF 00 FF
 ' 00 00 7F 01 FF 00 FF 00 00 78 01 FF 00 FF 00 00 7E 01 FF 00 FF 00 00 62 01 FF A1 10 FF 00 FF 00 00 63 01 FF 00 FF 00 00 63 01 FF 00 FF 00 00 62 01 FF 00 FF 00 00 61 01 FF 61 00 00 00 00 00 B2 16 00 40 01 06 00 73 70 69 64 65 72 64 03 1E 00 00 00 00 00 00 00 00 00 00 4C 00 00 00 00 01 00 FF 00 00 01 A1 10 FF 00 FF 00 00 62 01 FF 00 FF 00 00 62 01 FF F3 06 FF 00 FF 00 00 62 01 FF 00 FF 82 66 D7 9F 00 A0 49 26 4A 00 00 00 78 03 25 0B FF 78 04 E9 0D FF 78 06 C6 0C FF 78 0A 68 0B FF 78 0B D7 3E FF A2 80 00 83 3D 7D AC 7D 08 0B 9E 05 00 01 02 03 04 00 B7 00 03 00 05 00 0A 00 A0 39 00 96 00 F8 7A 00 00 40 9C 00 00 5C 00 00 00 00 00 00 00 01 00 5C 04 0F 27 00 80 05 00 05 00 00 00 00 64 D8 09 6E 00 00 00 D0 02 A1 0A 00 0A 00 00 0C 00 0C 00 34 0A 00 0A 00 00 0A 00 0A 00 00 0A 00 0A 00 00 0A 00 0A 00 00 0A 00 0A 00 00

        pos = pos + 8
     
    Case &HB8
      ' new since tibia 10.51 preview
      ' B8 00 B7 00 03 00 05 00 0A 00 A0 94 00 96 00 F1 0E 00 00 40 9C 00 00 00 00 00 00 00 00 00 00 01 00 00 05
      pos = pos + 2
    
    Case &HBE
     ' up -1 floor
     ' eval how many new floors should we get
     myX(idConnection) = myX(idConnection) + 1
     myY(idConnection) = myY(idConnection) + 1
     myZ(idConnection) = myZ(idConnection) - 1
     If (myZ(idConnection) < 7) Then
       ' we don't get any new info
       pos = pos + 1
     ElseIf myZ(idConnection) = 7 Then
       startz = 5
       endz = 0
       zstep = -1
       pos = ReadNewFloors(idConnection, packet, pos, startz, endz, zstep)
     Else
       startz = myZ(idConnection) - 2
       endz = myZ(idConnection) - 2
       zstep = 1
       pos = ReadNewFloors(idConnection, packet, pos, startz, endz, zstep)
     End If
     EvalMyMove idConnection, 0, 0, -1
     
     gotMapUpdate = True
     
      lastFloorChangeX(idConnection) = myX(idConnection)
      lastFloorChangeY(idConnection) = myY(idConnection)
      lastFloorChangeZ(idConnection) = myZ(idConnection)
    Case &HBF
      ' down +1 floor
      ' eval how many new floors should we get
      myX(idConnection) = myX(idConnection) - 1
      myY(idConnection) = myY(idConnection) - 1
      myZ(idConnection) = myZ(idConnection) + 1
      If (myZ(idConnection) <= 7) Then
        ' we don't get any new info
        pos = pos + 1
      ElseIf myZ(idConnection) = 8 Then
        ' 3 floors
        startz = 8
        endz = 10
        zstep = 1
        ' you were removed in a packet type 6C before, but the new floors info will contain you again,
        ' so don't worry!
        pos = ReadNewFloors(idConnection, packet, pos, startz, endz, zstep)
      ElseIf myZ(idConnection) > 13 Then
        ' we don't get any new info
        pos = pos + 1
      Else
        ' 1 floor
        startz = myZ(idConnection) + 2
        endz = myZ(idConnection) + 2
        zstep = 1
        pos = ReadNewFloors(idConnection, packet, pos, startz, endz, zstep)
      End If
      EvalMyMove idConnection, 0, 0, 1
      gotMapUpdate = True
      
      lastFloorChangeX(idConnection) = myX(idConnection)
      lastFloorChangeY(idConnection) = myY(idConnection)
      lastFloorChangeZ(idConnection) = myZ(idConnection)
    Case &HC8
      ' available outfits
      ' tibia 8.7 : C8 81 00 72 72 72 72 00 00 00 04 80 00 07 00 43 69 74 69 7A 65 6E 00 81 00 06 00 48 75 6E 74 65 72 00 82 00 04 00 4D 61 67 65 00 83 00 06 00 4B 6E 69 67 68 74 00 00
      ' tibia 8.7 : C8 84 00 4C 4F 58 52 00 72 01 0B 80 00 07 00 43 69 74 69 7A 65 6E 04 81 00 06 00 48 75 6E 74 65 72 04 82 00 04 00 4D 61 67 65 04 83 00 06 00 4B 6E 69 67 68 74 04 84 00 08 00 4E 6F 62 6C 65 6D 61 6E 04 85 00 08 00 53 75 6D 6D 6F 6E 65 72 04 86 00 07 00 57 61 72 72 69 6F 72 04 8F 00 09 00 42 61 72 62 61 72 69 61 6E 04 90 00 05 00 44 72 75 69 64 04 91 00 06 00 57 69 7A 61 72 64 04 92 00 08 00 4F 72 69 65 6E 74 61 6C 04 01 72 01 08 00 57 61 72 20 42 65 61 72
      ' fake 8.7  : C8 84 00 4C 4F 58 52 00 72 01 0B 80 00 07 00 43 69 74 69 7A 65 6E 04 81 00 06 00 48 75 6E 74 65 72 04 82 00 04 00 4D 61 67 65 04 83 00 06 00 4B 6E 69 67 68 74 04 84 00 08 00 4E 6F 62 6C 65 6D 61 6E 04 85 00 08 00 53 75 6D 6D 6F 6E 65 72 04 86 00 07 00 57 61 72 72 69 6F 72 04 8F 00 09 00 42 61 72 62 61 72 69 61 6E 04 90 00 05 00 44 72 75 69 64 04 91 00 06 00 57 69 7A 61 72 64 04 92 00 08 00 4F 72 69 65 6E 74 61 6C 04 02 72 01 08 00 57 61 72 20 42 65 61 72 72 01 08 00 57 61 72 20 42 65 61 73
      If TibiaVersionLong >= 870 Then
        ' mounts
        pos = pos + 10 ' skip current outfit
        lonN = CLng(packet(pos)) ' ammount of avaiable outfits
        'Debug.Print "total=" & CStr(lonN)
        pos = pos + 1
        For itemCount = 1 To lonN
          pos = pos + 2
          templ1 = GetTheLong(packet(pos), packet(pos + 1))
          'templ1 = ReadStringAndReturnLen(packet, pos)
          pos = pos + 2 + templ1 + 1
        Next itemCount
        lonN = CLng(packet(pos)) ' ammount of avaiable MOUNT outfits
        pos = pos + 1
        For itemCount = 1 To lonN
          pos = pos + 2
          templ1 = GetTheLong(packet(pos), packet(pos + 1))
          'templ1 = ReadStringAndReturnLen(packet, pos)
          pos = pos + 2 + templ1
        Next itemCount
      ElseIf TibiaVersionLong >= 790 Then
        ' now includes a name for each outfit:
        pos = pos + 8 ' skip current outfit
        lonN = CLng(packet(pos)) ' ammount of avaiable outfits
        pos = pos + 1
        For itemCount = 1 To lonN
          pos = pos + 2
          templ1 = GetTheLong(packet(pos), packet(pos + 1))
          pos = pos + 2 + templ1 + 1
        Next itemCount
      ElseIf TibiaVersionLong >= 773 Then
        ' 2A 00 C8 8B 00 43 72 72 72 00 0B 88 00 04 89 00 04 8A 00 04 8B 00 04 8C 00 04 8D 00 04 8E 00 04 93 00 04 94 00 04 95 00 04 96 00 04
        pos = pos + 8 ' skip current outfit
        lonN = CLng(packet(pos)) ' ammount of avaiable outfits
        pos = pos + 1 + (3 * lonN)
      ElseIf TibiaVersionLong <= 760 Then
        pos = pos + 8
      Else
        ' 0B 00 C8 82 00 72 5E 72 72 80 00 83 00
        pos = pos + 11
      End If

    Case &HD2
      ' update vip list item
      mobID = FourBytesDouble(packet(pos + 1), packet(pos + 2), packet(pos + 3), packet(pos + 4))
      lonN = GetTheLong(packet(pos + 5), packet(pos + 6))
      pos = pos + 7
      mobName = ""
      For itemCount = 0 To lonN - 1
        mobName = mobName & Chr(packet(pos))
        pos = pos + 1
      Next itemCount
      AddIDname idConnection, mobID, mobName
      If TibiaVersionLong >= 962 Then ' migrated vip list
        lonN = GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 2 + lonN
        pos = pos + 4 ' skip vip symbol (4 bytes)
        pos = pos + 1 ' skip notify at login (1 byte)
      End If
      lonN = CLng(packet(pos)) '00=offline ; 01 = online
      pos = pos + 1
    Case &HD3
      ' something about vip list
      ' D3 28 2A 8E 02 01
      ' D3 DC 76 B1 02 01
      pos = pos + 5
      If TibiaVersionLong >= 980 Then
        pos = pos + 1
      End If
    Case &HD4
      ' vip list update
      ' D4 09 00 00 00
      ' (at least in a ot server 7.6)
      pos = pos + 5
    Case &HDC
      ' hint , tibia 8.21+
      ' DC 01
      tempb1 = packet(pos + 1)
'      Select Case tempb1
'      Case &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF, &H10, &H11, &H12, &H13, &H14, &H15, &H16
'        ' remove this subpacket to avoid the popup
        ' new types since Tibia 9.81
      RemoveBytesFromTibiaPacket packet, pos, 2, finalAfterPos
'      Case Else
'        aRes = GiveGMmessage(idConnection, "GM? Unknown popup with code " & GoodHex(tempb1) & " All cheats have been paused. Reactivate with exiva play", "Blackd Proxy")
'        DoEvents
'        CheatsPaused(idConnection) = True
'        DangerGMname(idConnection) = "Strange popup"
'        ChangePlayTheDangerSound True
'        pos = pos + 2
'      End Select
    Case &HDD
      ' DD A9 81 28 7C 03 09 06 00 43 61 72 70 65 74
      ' DD 01 7D 16 7E 07 10 0E 00 54 6F 20 74 68 65 20 56 69 6C 6C 61 67 65
      ' mark in map , tibia 8.21+
      pos = pos + 7
      lonN = GetTheLong(packet(pos), packet(pos + 1))
      pos = pos + 2
      mobName = ""
      For itemCount = 0 To lonN - 1
        mobName = mobName & Chr(packet(pos))
        pos = pos + 1
      Next itemCount
    Case &HF0
      ' only in Tibia 7.82 +
      ' quest log
      '          AMMNT TYPE? STRLN QUEST NAME ------------------------------------------------ STATUSBYTE?
      ' 1C 00 F0 01 00 08 00 14 00 54 68 65 20 50 6F 73 74 6D 61 6E 20 4D 69 73 73 69 6F 6E 73 00
      lonN = GetTheLong(packet(pos + 1), packet(pos + 2))
      pos = pos + 3
      For itemCount = 1 To lonN
        pos = pos + 5 + GetTheLong(packet(pos + 2), packet(pos + 3))
      Next itemCount
    Case &HF1
      ' only in Tibia 7.82 +
      ' quest log part 2
     ' 6D 00 F1 08 00 01 1B 00 4D 69 73 73 69 6F 6E 20 30 31 3A 20 54 68 65 20 53 68 69 70 20 52 6F 75 74 65 73
     '                   4A 00 59 6F 75 72 20 63 75 72 72 65 6E 74 20 74 61 73 6B 20 69 73 20 74 6F 20 74 72 61 76 65 6C 20 77 69 74 68 20 43 61 70 74 61 69 6E 20 42 6C 75 65 62 65 61 72 20 66 72 6F 6D 20 54 68 61 69 73 20 74 6F 20 43 61 72 6C 69 6E 2E
      lonN = CLng(packet(pos + 3)) ' ammount of mission entries?
      pos = pos + 4
      For itemCount = 1 To lonN
        pos = pos + 2 + GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 2 + GetTheLong(packet(pos), packet(pos + 1))
      Next itemCount
    Case &H7A
      ' only in Tibia 8.2+
      ' new trade system - open trade
     ' Debug.Print ">>" & frmMain.showAsStr3(packet, True, pos, pos + 200)
     If TibiaVersionLong >= 940 Then
' 7A
' 08 00 42 65 6E 6A 61 6D 69 6E
' 03 00
' B3 0D 00
' 05 00 6C 61 62 65 6C
' 0A 00 00 00 01 00 00 00 00 00 00 00
' B1 0D 00
' 06 00 6C 65 74 74 65 72
' 32 00 00 00 08 00 00 00 00 00 00 00
' AF 0D 00
' 06 00 70 61 72 63 65 6C
' 08 07 00 00 0F 00 00 00 00 00 00 00
        pos = pos + 1
        templ1 = GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 2 + templ1
        lonN = GetTheLong(packet(pos), packet(pos + 1)) ' ammount of items in the trade list
        pos = pos + 2
        For itemCount = 1 To lonN
           pos = pos + 3
           templ1 = GetTheLong(packet(pos), packet(pos + 1))
           pos = pos + 2 + templ1
           pos = pos + 12
        Next itemCount
     Else
      pos = pos + 1
      If TibiaVersionLong >= 872 Then ' name of the trader npc
        templ1 = GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 2 + templ1
      End If
      lonN = CLng(packet(pos)) ' ammount of items in the trade list
      pos = pos + 1
      For itemCount = 1 To lonN
        pos = pos + 3 ' skip item id (who cares)
        templ1 = GetTheLong(packet(pos), packet(pos + 1))
        pos = pos + 2
        tmpStr = ""
        For templ2 = 1 To templ1
          tmpStr = tmpStr & Chr(packet(pos))
          pos = pos + 1
        Next templ2
        If TibiaVersionLong >= 830 Then
            templ1 = FourBytesLong(packet(pos), packet(pos + 1), packet(pos + 2), packet(pos + 3)) ' cap
            pos = pos + 4
        End If
        templ1 = FourBytesLong(packet(pos), packet(pos + 1), packet(pos + 2), packet(pos + 3)) ' sell price
        pos = pos + 4
        templ1 = FourBytesLong(packet(pos), packet(pos + 1), packet(pos + 2), packet(pos + 3)) ' buy price
        pos = pos + 4
      Next itemCount
     End If
      ' common part
      doingTrade2(idConnection) = True
    Case &H7B
      ' only in Tibia 8.2+
      ' new trade system - total money?, after 7A
      'Debug.Print ">>" & frmMain.showAsStr3(packet, True, pos, pos + 40)
      
      ' 7B E8 03 00 00 00 00 00 00 00
      ' 7B 32 00 00 00 00 00 00 00 02 BB 0B 01 81 0D 01
      pos = pos + 1
      
      templ1 = FourBytesLong(packet(pos), packet(pos + 1), packet(pos + 2), packet(pos + 3))
      pos = pos + 4
      doingTrade2(idConnection) = True
      If TibiaVersionLong >= 980 Then
        pos = pos + 4 ' strange new thing
        templ1 = CLng(packet(pos))
        pos = pos + 1 ' ammount of listed items
        pos = pos + 3 * templ1
      ElseIf TibiaVersionLong >= 830 Then
        templ1 = CLng(packet(pos)) ' ammount of listed items
        pos = pos + 1
        If TibiaVersionLong >= 900 Then
          pos = pos + (3 * templ1)
        ElseIf TibiaVersionLong >= 872 Then
          pos = pos + (5 * templ1)
        Else
          pos = pos + (3 * templ1)
        End If

      End If
    Case &H7C
      ' only in Tibia 8.2+
      ' new trade system - close trade
      pos = pos + 1
      doingTrade2(idConnection) = False
    Case &HF2
      ' Tibia 8.7 +
      ' report statement result?
      If TibiaVersionLong >= 1080 Then
        ' Tibia 10.80+
        ' F2 00
        pos = pos + 2
      Else
        templ1 = GetTheLong(packet(pos + 1), packet(pos + 2))
        pos = pos + 3 + templ1
      End If
    Case &HF3
      ' Tibia 8.72
      ' unknown
      ' F3 00 00 09 00 4D 69 6E 75 6E 69 6E 68 61 02
      ' F3 00 00 09 00 4D 69 6E 75 6E 69 6E 68 61 00
      templ1 = GetTheLong(packet(pos + 3), packet(pos + 4))
      pos = pos + 6 + templ1
    Case &HF5
      ' new since Tibia 10.76
      ' List of equipable items found in your char
      ' F5
      ' 11 00 - item count: 17, then items are listed below:
      ' 01 00 00 01 00
      ' 02 00 00 01 00
      ' 03 00 00 01 00
      ' 04 00 00 01 00
      ' 05 00 00 01 00
      ' 06 00 00 01 00
      ' 07 00 00 01 00
      ' 08 00 00 01 00
      ' 09 00 00 01 00
      ' 0A 00 00 01 00
      ' 0B 00 00 01 00
      ' 25 0B 00 01 00
      ' D7 0B 00 05 00
      ' 51 0D 00 01 00
      ' E0 0D 00 01 00
      ' 5E 1E 00 01 00
      ' D7 3E 00 01 00
      templ1 = GetTheLong(packet(pos + 1), packet(pos + 2))
      pos = pos + 3 + (templ1 * 5)
      
    Case &HF6 ' TIBIA 9.4 - OPENING AUCTION HOUSE
      '            F6 00 00 00 00 01 00 00 00
      '            F6 10 27 00 00 01 00 06 00 0A 01 01 00 1F 0D 01 00 2C 0D 01 00 2F 0D 01 00 51 0D 01 00 5E 1E 01 00
      ' tibia 9.5: F6 1C 25 00 00 64 00 00
      pos = pos + 5 ' skip money at bank
      If TibiaVersionLong >= 980 Then
        pos = pos + 1 ' skip first number (unknown)
        pos = pos + 4 ' skip second  number (unknown)
      ElseIf TibiaVersionLong >= 950 Then
        pos = pos + 1 ' skip first number (unknown)
      Else
        pos = pos + 2 ' skip first number (unknown)
      End If
      templ1 = GetTheLong(packet(pos), packet(pos + 1))
      pos = pos + 2 ' skip number of items in depot
      ' skip item info
      pos = pos + (4 * templ1)
    Case &HF7 ' TIBIA 9.4 - CLOSING AUCTION HOUSE
      'F7
      pos = pos + 1
    Case &HF8 ' TIBIA 9.4 - Auction House - Details
' F8
' A7 2D
' 02 00 31 32
' 00 00
' 00 00
' 00 00
' 00 00
' 00 00
' 08 00 66 69 72 65 20 2B 35 25
' 03 00 31 30 30
' 00 00
' 11 00 64 72 75 69 64 73 2C 20 73 6F 72 63 65 72 65 72 73
' 00 00
' 0E 00 6D 61 67 69 63 20 6C 65 76 65 6C 20 2B 32
' 00 00
' 00 00
' 08 00 34 35 2E 30 30 20 6F 7A
' 01
' 00 00 00 00
' 00 00 00 00
' 00 00 00 00
' FF FF FF FF
' 01
' 01 00 00 00
' C0 27 09 00
' C0 27 09 00
' C0 27 09 00
        pos = pos + 3
        For templ1 = 1 To 15
            lonN = GetTheLong(packet(pos), packet(pos + 1))
            pos = pos + 2 + lonN
        Next templ1
        lonN = CLng(packet(pos)) 'info blocks
        ' skip first block of info
        pos = pos + 1 + (16 * lonN)
        
        lonN = CLng(packet(pos)) 'info blocks
        ' skip second block of info
        pos = pos + 1 + (16 * lonN)
      
    Case &HF9 ' TIBIA 9.4 - Auction House - Sell offers / Buy offers
      ' 2 ways of parsing depending the first bytes (00 00 / FF FF or FE FF)
      ' F9 FE FF 01 00 00 00 EC 04 29 50 00 00 EF 0D 01 00 01 00 00 00 00 00 00 00
      If ((packet(pos + 1) = &HFE) And (packet(pos + 2) = &HFF)) Then
        ' detected in tibia 9.54
        ' PARSER A
      ' F9
      ' FE FF
      ' 00 00 00 00
      
      ' 02 00 00 00
      
      ' 3B 45 06 50
      ' 00 00 A7 2D
      ' 01 00
      ' 10 EB 09 00
      
      ' 70 46 06 50
      ' 00 00 57 1F
      ' 01 00
      ' A0 D9 08 00
        pos = pos + 3
        templ1 = CLng(FourBytesDouble(packet(pos), packet(pos + 1), packet(pos + 2), packet(pos + 3)))
        pos = pos + 4 + (14 * templ1)
        templ2 = CLng(FourBytesDouble(packet(pos), packet(pos + 1), packet(pos + 2), packet(pos + 3)))
        pos = pos + 4 + (14 * templ2)
      ElseIf ((packet(pos + 1) = &HFF) And (packet(pos + 2) = &HFF)) Then
      ' PARSER A
' F9
' FF FF

' 01 00 00 00

' 79 FA 11 4F
' 00 00 37 0D
' 01 00
' 90 01 00 00
' 03

' 02 00 00 00

' 0F FA 11 4F
' 00 00 1F 0D
' 01 00
' C8 00 00 00
' 01

' 86 FD 12 4F
' 00 00 37 0D
' 01 00
' F4 01 00 00
' 03
pos = pos + 3
templ1 = CLng(FourBytesDouble(packet(pos), packet(pos + 1), packet(pos + 2), packet(pos + 3)))
pos = pos + 4 + (15 * templ1)
templ2 = CLng(FourBytesDouble(packet(pos), packet(pos + 1), packet(pos + 2), packet(pos + 3)))
pos = pos + 4 + (15 * templ2)
      
      Else
        ' PARSER B
          pos = pos + 3
          templ1 = CLng(FourBytesDouble(packet(pos), packet(pos + 1), packet(pos + 2), packet(pos + 3)))
          pos = pos + 4
          For templ2 = 1 To templ1
    ' example of item:
    '77 FC 10 4F
    '00 00
    '45 00 '69 items
    '64 00 00 00 ' for 100 gold
    '0C 00 44 61 72 6B 20 56 65 6E 74 72 75 6D ' Dark Ventrum
             pos = pos + 12
             lonN = GetTheLong(packet(pos), packet(pos + 1))
             pos = pos + 2 + lonN
          Next templ2
          templ1 = CLng(FourBytesDouble(packet(pos), packet(pos + 1), packet(pos + 2), packet(pos + 3)))
          pos = pos + 4
          For templ2 = 1 To templ1
    ' example of item:
    '77 FC 10 4F
    '00 00
    '45 00 '69 items
    '64 00 00 00 ' for 100 gold
    '0C 00 44 61 72 6B 20 56 65 6E 74 72 75 6D ' Dark Ventrum
             pos = pos + 12
             lonN = GetTheLong(packet(pos), packet(pos + 1))
             pos = pos + 2 + lonN
          Next templ2
      End If
    Case &HFA
      ' activate premium scroll (Tibia 9.54 +)
      ' FA 01 00 00 00 17 00 41 63 74 69 76 61 74 65 20 50 72 65 6D 69 75 6D 20 53 63 72 6F 6C 6C 5C 00 41 72 65 20 79 6F 75 20 73 75 72 65 20 74 68 61 74 20 79 6F 75 20 77 61 6E 74 20 74 6F 20 61 63 74 69 76 61 74 65 20 74 68 69 73 20 70 72 65 6D 69 75 6D 20 73 63 72 6F 6C 6C 3F 20 54 68 69 73 20 6F 70 65 72 61 74 69 6F 6E 20 63 61 6E 6E 6F 74 20 62 65 20 75 6E 64 6F 6E 65 2E 02 04 00 4F 6B 61 79 00 06 00 43 61 6E 63 65 6C 01 FF FF
      ' FA 01 00 00 00 17 00 41 63 74 69 76 61 74 65 20 50 72 65 6D 69 75 6D 20 53 63 72 6F 6C 6C 5C 00 41 72 65 20 79 6F 75 20 73 75 72 65 20 74 68 61 74 20 79 6F 75 20 77 61 6E 74 20 74 6F 20 61 63 74 69 76 61 74 65 20 74 68 69 73 20 70 72 65 6D 69 75 6D 20 73 63 72 6F 6C 6C 3F 20 54 68 69 73 20 6F 70 65 72 61 74 69 6F 6E 20 63 61 6E 6E 6F 74 20 62 65 20 75 6E 64 6F 6E 65 2E 02 04 00 4F 6B 61 79 00 06 00 43 61 6E 63 65 6C 01 FF FF
      ' FA 01 00 00 00 17 00 41 63 74 69 76 61 74 65 20 50 72 65 6D 69 75 6D 20 53 63 72 6F 6C 6C 5C 00 41 72 65 20 79 6F 75 20 73 75 72 65 20 74 68 61 74 20 79 6F 75 20 77 61 6E 74 20 74 6F 20 61 63 74 69 76 61 74 65 20 74 68 69 73 20 70 72 65 6D 69 75 6D 20 73 63 72 6F 6C 6C 3F 20 54 68 69 73 20 6F 70 65 72 61 74 69 6F 6E 20 63 61 6E 6E 6F 74 20 62 65 20 75 6E 64 6F 6E 65 2E 02 04 00 4F 6B 61 79 00 06 00 43 61 6E 63 65 6C 01 FF FF
      
      ' new beds training 9.7 :
      ' FA 01 00 00 00 0E 00 43 68 6F 6F 73 65 20 61 20 53 6B 69 6C 6C 16 00 50 6C 65 61 73 65 20 63 68 6F 6F 73 65 20 61 20 73 6B 69 6C 6C 3A 02 04 00 4F 6B 61 79 00 06 00 43 61 6E 63 65 6C 01 05 1C 00 53 77 6F 72 64 20 46 69 67 68 74 69 6E 67 20 61 6E 64 20 53 68 69 65 6C 64 69 6E 67 01 1A 00 41 78 65 20 46 69 67 68 74 69 6E 67 20 61 6E 64 20 53 68 69 65 6C 64 69 6E 67 02 1B 00 43 6C 75 62 20 46 69 67 68 74 69 6E 67 20 61 6E 64 20 53 68 69 65 6C 64 69 6E 67 03 1F 00 44 69 73 74 61 6E 63 65 20 46 69 67 68 74 69 6E 67 20 61 6E 64 20 53 68 69 65 6C 64 69 6E 67 04 19 00 4D 61 67 69 63 20 4C 65 76 65 6C 20 61 6E 64 20 53 68 69 65 6C 64 69 6E 67 05 01 00 00
      
      ' FA 01 00 00 00
      ' 0E 00 43 68 6F 6F 73 65 20 61 20 53 6B 69 6C 6C
      
      ' 16 00 50 6C 65 61 73 65 20 63 68 6F 6F 73 65 20 61 20 73 6B 69 6C 6C 3A
      
      ' 02
      ' 04 00 4F 6B 61 79
      ' 00
      ' 06 00 43 61 6E 63 65 6C
      ' 01
      ' 05
      ' 1C 00 53 77 6F 72 64 20 46 69 67 68 74 69 6E 67 20 61 6E 64 20 53 68 69 65 6C 64 69 6E 67
      ' 01
      ' 1A 00 41 78 65 20 46 69 67 68 74 69 6E 67 20 61 6E 64 20 53 68 69 65 6C 64 69 6E 67
      ' 02
      ' 1B 00 43 6C 75 62 20 46 69 67 68 74 69 6E 67 20 61 6E 64 20 53 68 69 65 6C 64 69 6E 67
      ' 03
      ' 1F 00 44 69 73 74 61 6E 63 65 20 46 69 67 68 74 69 6E 67 20 61 6E 64 20 53 68 69 65 6C 64 69 6E 67
      ' 04
      ' 19 00 4D 61 67 69 63 20 4C 65 76 65 6C 20 61 6E 64 20 53 68 69 65 6C 64 69 6E 67
      ' 05
      ' 01 00 00
      
      
      
      
      ' forcedDebugChain = False
      ' FA
      pos = pos + 1
      ' 01 00 00 00
      pos = pos + 4
      ' 17 00 41 63 74 69 76 61 74 65 20 50 72 65 6D 69 75 6D 20 53 63 72 6F 6C 6C
      lonN = GetTheLong(packet(pos), packet(pos + 1))
      pos = pos + 2 + lonN
      ' 5C 00 41 72 65 20 79 6F 75 20 73 75 72 65 20 74 68 61 74 20 79 6F 75 20 77 61 6E 74 20 74 6F 20 61 63 74 69 76 61 74 65 20 74 68 69 73 20 70 72 65 6D 69 75 6D 20 73 63 72 6F 6C 6C 3F 20 54 68 69 73 20 6F 70 65 72 61 74 69 6F 6E 20 63 61 6E 6E 6F 74 20 62 65 20 75 6E 64 6F 6E 65 2E
      lonN = GetTheLong(packet(pos), packet(pos + 1))
      pos = pos + 2 + lonN
      ' 02
      pos = pos + 1
      ' 04 00 4F 6B 61 79
      lonN = GetTheLong(packet(pos), packet(pos + 1))
      pos = pos + 2 + lonN
      ' 00
      pos = pos + 1
      ' 06 00 43 61 6E 63 65 6C
      lonN = GetTheLong(packet(pos), packet(pos + 1))
      pos = pos + 2 + lonN
      ' 01
      pos = pos + 1
      ' FF FF
      If (packet(pos) = &HFF) Then
        pos = pos + 2
      Else
        ' new offline training in beds, since 9.7
        templ1 = CLng(packet(pos))
        pos = pos + 1
        For itemCount = 1 To templ1
            templ2 = GetTheLong(packet(pos), packet(pos + 1))
            pos = pos + 2 + templ2 + 1
        Next itemCount
        templ2 = GetTheLong(packet(pos), packet(pos + 1))
        ' 01 00 00
        pos = pos + 2 + templ2
      End If
    Case Else
      ' should not happen, unless protocol get updated
       ' LogOnFile "errors.txt", "WARNING IN PACKET" & frmMain.showAsStr2(packet, 0) & vbCrLf & "UNKNOWN PTYPE : " & GoodHex(pType)
       debugReasons = debugReasons & vbCrLf & " [ UNKNOWN PTYPE : " & GoodHex(pType) & " ] "
       showDebug = True
       expectMore = False
         myres.fail = True
    End Select
    If pos = finalAfterPos Then
      expectMore = False
    ElseIf pos > finalAfterPos Then
      debugLon1 = pos - finalAfterPos
      debugReasons = debugReasons & vbCrLf & " [ BAD EVAL IN PTYPE : " & GoodHex(pType) & " ; overread of +" & CStr(debugLon1) & _
      " bytes. Last good position=" & CStr(lastGoodPos) & _
      " x=" & CStr(myX(idConnection)) & _
      ";y=" & CStr(myY(idConnection)) & _
      ";z=" & CStr(myZ(idConnection)) & " ] "
      showDebug = True
        myres.fail = True
      expectMore = False
    End If
    lastpType = pType
  Loop Until expectMore = False
  'showDebug = True ' uncomment for forced debug
  If (forcedDebugChain = True) Then
    showDebug = True
    forcedDebugChain = False
    debugReasons = debugReasons & vbCrLf & " [THE FORMAT OF THE PACKET IS CORRECT] "
  End If
  If showDebug = True Then
    LogOnFile "errors.txt", "ERROR:" & vbCrLf & ">>>> 1. RECEIVED PACKET:" & vbCrLf & frmMain.showAsStr2(packet, 0) & vbCrLf & vbCrLf & ">>>> 2. LAST TOTAL CHAIN:" & debugChain & vbCrLf & vbCrLf & ">>>> 3. DEBUG REASONS:" & debugReasons & vbCrLf
  End If
  myres.pos = pos
  If gotMapUpdate = True And GotTrialLock = False Then
    If sentFirstPacket(idConnection) = False Then
      frmTrueMap.SetButtonColours
      frmTrueMap.DrawFloor
    End If
    ShowPositionChange idConnection
  End If
  ' should we ignore the corpse pop? (not being near or being in other floor)
  If (myres.gotNewCorpse = True) And (ignoreCorpsePop = True) Then
    myres.gotNewCorpse = False
  End If
  LearnFromPacket = myres
  Exit Function
fatalError:
  myres.pos = pos
  myres.fail = True
  LogOnFile "errors.txt", "FATAL ERROR [" & showDebug & " ] DEBUG CHAIN [ " & debugChain & " SP:" & debugChainType & _
   " ] CLIENT [ " & CStr(idConnection) & " ] (Number:" & Err.Number & _
   " Description:" & Err.Description & " Source: " & Err.Source & ") in packet : " & _
   frmMain.showAsStr2(packet, 0) & _
   " Last good position=" & CStr(lastGoodPos) & _
      " x=" & CStr(myX(idConnection)) & _
      ";y=" & CStr(myY(idConnection)) & _
      ";z=" & CStr(myZ(idConnection))
  LearnFromPacket = myres
End Function

Public Function LearnFromServer(ByRef packet() As Byte, idConnection As Integer) As Integer
  ' process info returned by LearnFromPacket
  Dim res As Integer
  Dim aRes As Long
  Dim fRes As Long
  Dim lastB As Long
  Dim pos As Long
  Dim afterMapPos As Long
  Dim stopA As Long
  Dim learnResult As TypeLearnResult
 ' Dim resS As TypeSearchItemResult2
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim percent As Long
  Dim inRes As Integer
  Dim blnTmp As Boolean
  Dim lootDone As Boolean
  Dim gtCount As Long
  Dim posSetting As Long
  Dim act As String
  #If FinalMode Then
  On Error GoTo exitE
  #End If
  res = 0
  gtCount = GetTickCount()
  If GameConnected(idConnection) = True Then
    stopA = False
    pos = 2
    If UBound(packet) < 2 Then
       res = 1 ' skip this packet
       GoTo exitL
    End If
    ' learn about moves
    If GotPacketWarning(idConnection) = False Then
      LogoutReason(idConnection) = ""
      learnResult = LearnFromPacket(packet, pos, idConnection)
      
      If learnResult.skipThis = True Then
        res = 1 ' skip this packet
        GoTo exitL
      End If
      If learnResult.fail = True Then
        GotPacketWarning(idConnection) = True
        frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Unexpected packet received. A detailed report can be found in ERRORS.TXT" & vbCrLf & "PLEASE send me that file to daniel@blackdtools.com so I can fix it!"
        aRes = GiveGMmessage(idConnection, "Proxy got critical error! Safe mode ON . Alarm on. Report log details (errors.txt) to daniel@blackdtools.com . Relog to reactivate cheats", "Blackd")
        ChangePlayTheDangerSound True
        DoEvents
        GoTo exitL
      End If
      ' autoloot
      
      If (learnResult.gotNewCorpse = True) And (CheatsPaused(idConnection) = False) And (ReconnectionStage(idConnection) = 0) Then
        lootDone = False
        blnTmp = False
        If (autoLoot(idConnection) = True) And (GotPacketWarning(idConnection) = False) And (sentWelcome(idConnection) = True) Then
          If (DangerGM(idConnection) = False) And (gtCount > lootTimeExpire(idConnection)) Then
            If (friendlyMode(idConnection) > 0) Then
              blnTmp = PotentialDanger(idConnection)
            Else
              blnTmp = False
            End If
            If blnTmp = False Then
              aRes = OpenCorpse(idConnection)
              DoEvents
              lootDone = True
            End If
          End If
        End If
        If lootDone = False Then
          If publicDebugMode = True Then
            If (DangerGM(idConnection) = True) Then
              aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Is not safe to loot that corpse (GM alarm is up)")
              DoEvents
            Else
              aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Is not safe to loot that corpse (PotentialDanger returned " & BooleanAsStr(blnTmp) & ")")
              DoEvents
            End If
          End If
        End If
      End If
 
      If frmHardcoreCheats.chkLogoutIfDanger.Value = 1 And GotTrialLock = False Then
        If LogoutReason(idConnection) <> "" Then
          res = 3 'ignore packet
          aRes = GiveServerError("Logged out because found : " & LogoutReason(idConnection) & " in your screen", idConnection)
          DoEvents
          GoTo exitL
        End If
      End If
      If sentFirstPacket(idConnection) = True And sentWelcome(idConnection) = False Then
        GetProcessAllProcessIDs ' get new relations of process IDs
        If CharacterName(idConnection) = "" Then
          'MsgBox "debug"
        Else
        frmMapReader.AddListItem CharacterName(idConnection) ' add player to big map
        frmMapReader.SetCurrentCenter CharacterName(idConnection)
        End If

        If ReconnectionStage(idConnection) > 0 Then
            frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Reconnection of " & CharacterName(idConnection) & " completed!"
            If frmEvents.chkReconnectionAlarm.Value = 1 Then
              ChangePlayTheDangerSound True
            End If
            aRes = GiveGMmessage(idConnection, "Your client had to be reconnected", "Warning")
            DoEvents
            GotPacketWarning(idConnection) = False
            ReconnectionStage(idConnection) = 3 ' now open backpacks
        Else
        myInitialExp(idConnection) = myExp(idConnection)
        myInitialTickCount(idConnection) = gtCount
        If TrialVersion = True Then
          UpdateCompDate
          If receivedLogin(idConnection) = False Then
            aRes = GiveGMmessage(idConnection, "Sorry, this kind of login is not allowed in trial version. CHEATS WILL NOT WORK WITH THIS CLIENT!", "Blackd")
            GotTrialLock = False
            GotPacketWarning(idConnection) = True
            DoEvents
          ElseIf GotTrialLock = False Then
            aRes = GiveGMmessage(idConnection, "Welcome to your trial of Blackd Proxy v" & ProxyVersion & " ! Cheats ENABLED. Official site: www.blackdtools.com", "Blackd")
            DoEvents
          Else
            aRes = GiveGMmessage(idConnection, "Your trial version have expired. CHEATS WILL NOT WORK WITH THIS CLIENT. Buy the program if you liked it! . Official site: www.blackdtools.com", "Blackd")
            GotTrialLock = False
            GotPacketWarning(idConnection) = True
            DoEvents
          End If
        Else
          If LimitedLeader = "-" Then
            aRes = GiveGMmessage(idConnection, "Welcome to Blackd Proxy v" & ProxyVersion & " ! Official site: www.blackdtools.com", "Blackd")
            DoEvents
          Else
            aRes = GiveGMmessage(idConnection, "Welcome! This is a limited version of Blackd Proxy v" & ProxyVersion & " - Cheats are limited to allow sync attack orders from me", LimitedLeader)
            DoEvents
          End If
        End If
        End If
        sentWelcome(idConnection) = True
    

      End If
      If learnResult.firstMapDone = True And sentFirstPacket(idConnection) = False Then
        res = 2
      End If
      
      ' AUTO MANA RECHARGE
      If learnResult.gotManaupdate Then
        If GotCustomMANAsettings(idConnection) = True Then
            ' nothing
        Else ' apply global settings
            If RuneMakerOptions(idConnection).ManaFluid = True Then
              If CheatsPaused(idConnection) = False Then
                If RuneMakerOptions(idConnection).LowMana > myMana(idConnection) Then

                    AddSpamOrder idConnection, 4 'add auto mana
          
              
                End If
              End If
              If RuneMakerOptions(idConnection).LowMana <= myMana(idConnection) Then
                RemoveSpamOrder idConnection, 4 'remove auto mana
              End If
            End If
        End If
      End If
      ' AUTO HP RECHARGE
      If (learnResult.gotHPupdate = True) Then
        percent = 100 * ((myHP(idConnection) / myMaxHP(idConnection)))
        If GotCustomHPsettings(idConnection) = True Then
            ' nothing
        Else ' apply global settings
            If (percent < GLOBAL_RUNEHEAL_HP) And _
             (frmHardcoreCheats.chkAutoHeal.Value = 1) And _
             (sentFirstPacket(idConnection) = True) Then
              If ((CheatsPaused(idConnection) = False) Or (AllowUHpaused(idConnection) = True)) Then
                AddSpamOrder idConnection, 1 'add auto UH
              End If
            ElseIf (percent >= GLOBAL_RUNEHEAL_HP) And _
             (sentFirstPacket(idConnection) = True) Then
              RemoveSpamOrder idConnection, 1 'remove  auto UH
              UHRetryCount(idConnection) = 0
            End If
            If (percent < frmHardcoreCheats.scrollHP2.Value) And _
             (frmHardcoreCheats.chkAutoVita.Value = 1) And _
             (sentFirstPacket(idConnection) = True) Then
              If CheatsPaused(idConnection) = False Then
                If (myMana(idConnection) >= CLng(frmHardcoreCheats.txtExuraVitaMana.Text)) Then
                  aRes = CastSpell(idConnection, frmHardcoreCheats.txtExuraVita.Text)
                  DoEvents
                End If
              End If
            End If
        
        End If
      End If
      ' LOGOUT-RUNEMAKER
      If (AfterLoginLogoutReason(idConnection) <> "") Then 'And (CheatsPaused(idConnection) = False) Then
        If frmRunemaker.ChkDangerSound.Value = 1 Then
          ChangePlayTheDangerSound True
        End If
        If DangerGM(idConnection) = True Then
          'if it is a gm, then cancel autologouts
          frmRunemaker.DisableAll idConnection
          AfterLoginLogoutReason(idConnection) = ""
        Else
        frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & CharacterName(idConnection) & " did runemaker logout - logged out because : " & AfterLoginLogoutReason(idConnection)
        
        
        ReconnectionStage(idConnection) = 0
        aRes = GiveServerError("Blackd Proxy logout - logged out because : " & AfterLoginLogoutReason(idConnection), idConnection)
        DoEvents
        
        sCheat = "01 00 14"
        inRes = GetCheatPacket(cPacket, sCheat)
        frmMain.UnifiedSendToServerGame idConnection, cPacket, True
        DoEvents

        ' dont relog if reconnection is enabled
        If ReconnectionStage(idConnection) > 0 Then ' blackd proxy 19.1
          ReconnectionStage(idConnection) = 10
        End If
        frmMain.DoCloseActions idConnection
        ReconnectionStage(idConnection) = 0
        logoutAllowed(idConnection) = 0
        DoEvents
        res = 3
        End If
      End If
    End If
  End If
exitL:
  LearnFromServer = res
  Exit Function
exitE:
  LogOnFile "errors.txt", "Unexpected error - loging packet: " & frmMain.showAsStr2(packet, 0)
  sentFirstPacket(idConnection) = True
  LearnFromServer = res
End Function

Public Function PotentialDanger(idConnection As Integer) As Boolean
  ' decides if there is potential danger in looting
  ' (if any player is near)
  Dim X As Integer
  Dim y As Integer
  Dim z As Integer
  Dim s As Integer
  Dim tileID As Integer
  Dim nameofgivenID As String
  If frmCavebot.chkLootProtection.Value = 1 Then
        PotentialDanger = False
        Exit Function
  End If
  z = myZ(idConnection)
  For X = -5 To 6
    For y = -4 To 5
        For s = 1 To 10
          tileID = GetTheLong(Matrix(y, X, z, idConnection).s(s).t1, Matrix(y, X, z, idConnection).s(s).t2)
          If tileID = 97 Then
           nameofgivenID = GetNameFromID(idConnection, Matrix(y, X, z, idConnection).s(s).dblID)
            If (isMelee(idConnection, nameofgivenID) = False) And (isHmm(idConnection, nameofgivenID) = False) And (frmRunemaker.IsFriend(LCase(nameofgivenID)) = False) Then
              If nameofgivenID <> CharacterName(idConnection) Then
                PotentialDanger = True
                Exit Function
              End If
            End If
          ElseIf tileID = 0 Then
            Exit For
          End If
        Next s
    Next y
  Next X
  PotentialDanger = False
End Function

Public Function PlayerOnScreen(idConnection As Integer) As String
  ' decides if there is potential danger in looting
  ' (if any player is near)
  Dim X As Integer
  Dim y As Integer
  Dim z As Integer
  Dim s As Integer
  Dim tileID As Integer
  Dim nameofgivenID As String
  z = myZ(idConnection)
  For X = -8 To 9
    For y = -6 To 7
        For s = 1 To 10
          tileID = GetTheLong(Matrix(y, X, z, idConnection).s(s).t1, Matrix(y, X, z, idConnection).s(s).t2)
          If tileID = 97 Then
           nameofgivenID = GetNameFromID(idConnection, Matrix(y, X, z, idConnection).s(s).dblID)
            If (isMelee(idConnection, nameofgivenID) = False) And (isHmm(idConnection, nameofgivenID) = False) And (frmRunemaker.IsFriend(LCase(nameofgivenID)) = False) Then
              If (nameofgivenID <> CharacterName(idConnection)) And (nameofgivenID <> "") Then
                PlayerOnScreen = nameofgivenID
                Exit Function
              End If
            End If
          ElseIf tileID = 0 Then
            Exit For
          End If
        Next s
    Next y
  Next X
  PlayerOnScreen = ""
End Function

Public Function PlayerOnScreen2(idConnection As Integer) As TypeSpecialRes
  ' decides if there is potential danger in looting
  ' (if any player is near)
  Dim X As Integer
  Dim y As Integer
  Dim z As Integer
  Dim s As Integer
  Dim tileID As Integer
  Dim nameofgivenID As String
  Dim res As TypeSpecialRes
  Dim theID As Double
  res.str = ""
  res.bln = False
  res.bestX = 0
  res.bestY = 0
  res.bestHMM = False
  res.bestMelee = False
  z = myZ(idConnection)
  For X = -6 To 7
    For y = -5 To 6
        For s = 1 To 10
          tileID = GetTheLong(Matrix(y, X, z, idConnection).s(s).t1, Matrix(y, X, z, idConnection).s(s).t2)
          If tileID = 97 Then
           theID = Matrix(y, X, z, idConnection).s(s).dblID
           nameofgivenID = GetNameFromID(idConnection, theID)
           If theID = SelfDefenseID(idConnection) Then
             res.bln = True
             res.bestX = X
             res.bestY = y
             If isMelee(idConnection, nameofgivenID) = True Then
               res.bestMelee = True
             End If
             If isHmm(idConnection, nameofgivenID) = True Then
               res.bestHMM = True
             End If
           End If
            If (isMelee(idConnection, nameofgivenID) = False) And (isHmm(idConnection, nameofgivenID) = False) And (frmRunemaker.IsFriend(LCase(nameofgivenID)) = False) Then
              If ((nameofgivenID <> CharacterName(idConnection)) And (nameofgivenID <> "")) Then
                res.str = nameofgivenID
              End If
            End If
          ElseIf tileID = 0 Then
            Exit For
          End If
        Next s
    Next y
  Next X
  PlayerOnScreen2 = res
End Function

Private Function ReadStringAndReturnLen(ByRef packet() As Byte, ByVal pos As Long) As Long
    Dim str As String
    Dim tl As Long
    Dim i As Long
    str = ""
    tl = GetTheLong(packet(pos), packet(pos + 1))
    For i = 0 To tl - 1
        str = str & Chr(packet(pos + 2 + i))
    Next i
    Debug.Print "one>" & str & "<"
    ReadStringAndReturnLen = tl
End Function

Public Sub CheckStackThing()
    Dim mmyid As Double
    mmyid = myID(1)
    Debug.Print "myid=" & mmyid
    Debug.Print "mysid=" & GetNameFromID(1, mmyid)
End Sub

Public Function StairsExistsAt(ByVal initX As Long, ByVal initY As Long, ByVal initZ As Long, ByVal idConnection As Integer) As Boolean
    Dim res As Boolean
    Dim relX As Long
    Dim relY As Long
    Dim SS As Long
    Dim tileID As Long
    res = False
    relX = initX - myX(idConnection)
    relY = initY - myY(idConnection)
    If ((relX < -7) Or (relX > 8) Or (relY < -5) Or (relY > 6)) Then
        StairsExistsAt = False
        Exit Function
    End If
    '   ReDim Matrix(-6 To 7, -8 To 9, 0 To 15, 1 To MAXCLIENTS) ' y, x, z, idConnection
    For SS = 0 To 10
        tileID = GetTheLong(Matrix(relY, relX, initZ, idConnection).s(SS).t1, Matrix(relY, relX, initZ, idConnection).s(SS).t2)
        If (DatTiles(tileID).requireRightClick = True) Then
            res = True
            Exit For
        End If
    Next SS
    StairsExistsAt = res
End Function

Public Function UseItemWithAmountX(idConnection As Integer, t1 As Byte, t2 As Byte, Ammount As Byte) As Boolean
'HHBCODE
'Maintenance note: this function is mostly just a copy of SearchAmmount, using EatFood
'searches for a single item with X ammount, and uses the first it find, then exits  function.
'(good for the common OT 100 Gold/100 platinum/100cc/ functions)

  'Dim res As TypeSearchItemResult
  Dim i As Long
  Dim j As Long
  Dim limitJ As Long
  'Dim rCount As Long
  'rCount = 0
  For i = 1 To EQUIPMENT_SLOTS
    If mySlot(idConnection, i).t1 = t1 And mySlot(idConnection, i).t2 = t2 Then
        If mySlot(idConnection, i).t3 = Ammount Then
         EatFood idConnection, t1, t2, CByte(i), CByte(j), False
         UseItemWithAmountX = True
         Exit Function
        End If
    End If
  Next i

  For i = 0 To HIGHEST_BP_ID
    If (Backpack(idConnection, i).open = True) Then
    limitJ = (Backpack(idConnection, i).used) - 1
    For j = 0 To limitJ
      If Backpack(idConnection, i).item(j).t1 = t1 And _
       Backpack(idConnection, i).item(j).t2 = t2 Then
        
        
        If Backpack(idConnection, i).item(j).t3 = Ammount Then
         EatFood idConnection, t1, t2, CByte(i), CByte(j), False
         UseItemWithAmountX = True
         Exit Function
        End If
      End If
      
      
    Next j
    End If
  Next i
  UseItemWithAmountX = False
  'SearchAmmount = rCount
End Function

Public Sub RemoveBytesFromTibiaPacket(ByRef packet() As Byte, ByVal pos As Long, ByVal bytesToRemove As Long, ByRef finalAfterPos)
Dim uboundBEFORE As Long
Dim i As Long
Dim uboundAFTER As Long
Dim headerBEFORE As Long
Dim headerAFTER As Long
uboundBEFORE = UBound(packet)
uboundAFTER = uboundBEFORE - bytesToRemove
If pos < 2 Then
    MsgBox "Bad use of function RemoveBytesFromPacket. Can't edit header", vbCritical + vbOKOnly, "Internal error of Blackd Proxy"
    End
End If
If uboundAFTER < 0 Then
    MsgBox "Bad use of function RemoveBytesFromPacket. Resulting packet would be null", vbCritical + vbOKOnly, "Internal error of Blackd Proxy"
    End
End If
' change header
headerBEFORE = GetTheLong(packet(0), packet(1))
headerAFTER = headerBEFORE - bytesToRemove
packet(0) = LowByteOfLong(headerAFTER)
packet(1) = HighByteOfLong(headerAFTER)
' move bytes
For i = pos To uboundAFTER
packet(i) = packet(i + bytesToRemove)
Next i
ReDim Preserve packet(uboundAFTER) ' resize packet
finalAfterPos = UBound(packet) + 1 ' update finalAfterPos
End Sub



