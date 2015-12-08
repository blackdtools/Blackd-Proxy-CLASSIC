Attribute VB_Name = "modCavebot"
#Const FinalMode = 1
#Const withdebugreposition = False
Option Explicit
Private Const maxSETUSEITEMDist As Long = 2
Public Const cte_RepositionDelay = 700
Public Const StackingWaitTurns = 7
'scripts
Public Type TypePosibleSETUSEITEM
  lngX As Long
  lngY As Long
  byteS As Byte
  tileb1 As Byte
  tileb2 As Byte
  strItem As String
End Type

Public Type TypeSpellKill
    spell As String
    dist As Long
End Type
Public Type TypeChangeFloorResult
  result As Byte
  X As Long
  y As Long
  z As Long
End Type
Public Type TypePathMatrix
  walkable(-8 To 9, -6 To 7) As Boolean
End Type
Public Type TypeIDmap
  isSafe(-9 To 10, -7 To 8) As Boolean
  isMelee(-9 To 10, -7 To 8) As Boolean
  isHmm(-9 To 10, -7 To 8) As Boolean
  dblID(-9 To 10, -7 To 8) As Double
End Type
Public ClientExecutingLongCommand() As Boolean
Public TibiaExePath As String
Public TibiaExePathWITHTIBIADAT As String
'Public MagebotPath As String
'Public MagebotExe As String
Public AllowedLootDistance() As Long
Public debugPIDs() As String
Public CteMoveDelay As Long
Public prevAttackState() As Boolean
Public TurnsWithRedSquareZero() As Long
Public lastAttackedIDstatus() As Double
Public CavebotTimeWithSameTarget() As Long
Public CavebotTimeStart() As Long
Public maxAttackTime() As Long
Public maxAttackTimeCHAOS() As Long
Public maxHit() As Long
Public previousAttackedID() As Double
Public cavebotLenght() As Long
Public cavebotOnDanger() As Long
Public cavebotOnGMclose() As Boolean
Public cavebotOnGMpause() As Boolean
Public cavebotOnPLAYERpause() As Boolean
Public cavebotOnTrapGiveAlarm() As Boolean
Public cavebotCurrentTargetPriority() As Long
Public cavebotScript() As scripting.Dictionary
Public cavebotMelees() As scripting.Dictionary
Public cavebotAvoid() As scripting.Dictionary
Public cavebotExorivis() As scripting.Dictionary
Public cavebotHMMs() As scripting.Dictionary
Public shotTypeDict() As scripting.Dictionary
Public exoriTypeDict() As scripting.Dictionary
Public cavebotGoodLoot() As scripting.Dictionary
Public killPriorities() As scripting.Dictionary
Public SpellKills_SpellName() As scripting.Dictionary
Public SpellKills_Dist() As scripting.Dictionary

Public DictSETUSEITEM() As scripting.Dictionary
Public DictSETUSEITEM_used() As Boolean
Public SETUSEITEM_lastX() As Long
Public SETUSEITEM_lastY() As Long

Public cavebotEnabled() As Boolean
Public EnableMaxAttackTime() As Boolean

Public setFollowTarget() As Boolean
Public lastAttackedID() As Double
Public exeLine() As Long
Public ProcessID() As Long
Public fishCounter() As Long
Public waitCounter() As Long
Public moveRetry() As Long
Public lastX() As Long
Public lastY() As Long
Public lastZ() As Long
Public lastDestX() As Long
Public lastDestY() As Long
Public lastDestZ() As Long
Public autoLoot() As Boolean
Public requestLootBp() As Byte
Public lootTimeExpire() As Long
Public currTargetName() As String
Public currTargetID() As Double
Public friendlyMode() As Byte
Public ignoreNext() As Long
Public lastFloorTrap() As Long
Public onDepotPhase() As Long
Public CavebotChaoticMode() As Long
Public cancelAllMove() As Long
Public depotX() As Long
Public depotY() As Long
Public depotZ() As Long
Public depotS() As Byte
Public depotTileID() As Long
Public AllowRepositionAtStart() As Long
Public AllowRepositionAtTrap() As Long
Public doneDepotChestOpen() As Boolean
Public lastDepotBPID() As Byte
Public somethingChangedInBps() As Boolean
Public nextForcedDepotDeployRetry() As Long
Public CheatsPaused() As Boolean
Public avoidC As Boolean
Public TimeToGiveTrapAlarm As Long
Public DelayAttacks() As Long
Public AvoidReAttacks() As Boolean
Public CavebotHaveSpecials() As Boolean
Public CavebotLastSpecialMove() As Long
Public specialGMnames As scripting.Dictionary
' MEMORY ADDRESSES

Public adrXgo As Long ' goto this x
Public adrYgo As Long ' goto this y
Public adrZgo As Long ' goto this z
Public adrGo As Long  ' start goto process of first battlelist item
Public adrOutfit As Long  ' first outfit byte of first battlelist item
Public adrNChar As Long ' updated - first ID in battlelist
Public CharDist As Long 'not changed
Public adrNum As Long  'updated  - yourID - now fixed
Public adrConnected As Long  ' 0 if not connected / else it is connected
Public adrPointerToInternalFPSminusH5D As Long ' pointer to an address near the internal value for FPS (inversely relative to FPS) , add +&H5D and you are there
Public adrInternalFPS As Long ' only for Tibia 7.6
Public adrNumberOfAttackClicks As Long ' number of clicks so far
' some other vars
Public SelfDefenseID() As Double
Public pauseStacking() As Integer
Public SetVeryFriendly_NOATTACKTIMER_ms As Long

Public MAX_LOCKWAIT As Long

Public EXORIVIS_COST As Long
Public EXORIVIS_SPELL As String

Public EXORIMORT_COST As Long
Public EXORIMORT_SPELL As String

Public SpellKillHPlimit() As Long
Public SpellKillMaxHPlimit() As Long
'...



' Sid=client id
' newExeLine= new line you want to set.
' RelativeToOldLine = should the new line number be relative to the old line (+1?), or absolute?
' (OPTIONAL) updateLst , set it to false if you don't want to update frmCavebot.lstScript.ListIndex
Public Sub updateExeLine(ByVal Sid As Long, ByVal newExeLine As Long, ByVal RelativeToOldLine As Boolean, Optional updateLst As Boolean = True)
    If (RelativeToOldLine = True) Then
        exeLine(Sid) = exeLine(Sid) + newExeLine
    Else
        exeLine(Sid) = newExeLine
    End If
    If (updateLst = True) Then
        If (modMap.cavebotIDselected = Sid) Then
            Dim eLine As Long
            eLine = exeLine(Sid)
            If frmCavebot.lstScript.ListCount > eLine Then
                #If FinalMode = 0 Then
                Debug.Print "Executing line " & eLine
                #End If
                frmCavebot.lstScript.ListIndex = eLine
            Else
                #If FinalMode = 0 Then
                Debug.Print "Trying to execute a line beyond script limits: " & eLine
                #End If
            End If
        End If
    End If
End Sub



Public Sub AddKillPriority(idConnection As Integer, str As String, lngPriority As Long)
  ' add item to dictionary
  Dim res As Boolean
  killPriorities(idConnection).item(LCase(str)) = lngPriority
End Sub

Public Sub RemoveAllKillPriorities(idConnection As Integer)
  ' remove item from dictionary
  Dim res As Boolean
  killPriorities(idConnection).RemoveAll
End Sub

Public Function getKillPriority(idConnection As Integer, ByRef str As String) As Long
  ' get the name from an ID
  Dim res As Boolean
  If killPriorities(idConnection).Exists(LCase(str)) = True Then
    getKillPriority = killPriorities(idConnection).item(LCase(str))
  Else
    getKillPriority = 0
  End If
End Function


Public Sub AddSpellKill(idConnection As Integer, mobName As String, spellName As String, dist As Long)
  ' add item to dictionary
  SpellKills_SpellName(idConnection).item(LCase(mobName)) = spellName
  SpellKills_Dist(idConnection).item(LCase(mobName)) = dist
End Sub

Public Sub RemoveAllSpellKills(idConnection As Integer)
  ' remove item from dictionary
  SpellKills_SpellName(idConnection).RemoveAll
  SpellKills_Dist(idConnection).RemoveAll
End Sub

Public Function getSpellKill(idConnection As Integer, ByRef mobName As String) As TypeSpellKill
  ' get the name from an ID
  Dim res As TypeSpellKill
  If SpellKills_SpellName(idConnection).Exists(LCase(mobName)) = True Then
    res.spell = SpellKills_SpellName(idConnection).item(LCase(mobName))
    res.dist = SpellKills_Dist(idConnection).item(LCase(mobName))
    getSpellKill = res
  Else
    res.spell = ""
    res.dist = -1
    getSpellKill = res
  End If
End Function


Public Sub SafeLoadSpecialGMnames(ByVal filename As String)
    On Error GoTo goterr
  Dim fso As scripting.FileSystemObject
  Dim fn As Integer
  Dim strLine As String
  Set fso = New scripting.FileSystemObject
    RemoveAllSpecialGMname

    If fso.FileExists(filename) = True Then
      fn = FreeFile
      Open filename For Input As #fn
      While Not EOF(fn)
        Line Input #fn, strLine
        If strLine <> "" Then
            If isSpecialGMname(LCase(strLine)) = False Then
             AddSpecialGMname LCase(strLine)
            End If
        End If
      Wend
      Close #fn
    End If
    Exit Sub
goterr:
    Exit Sub
End Sub

Public Sub LoadSpecialGMnames()
  Dim strMyPath As String
  Dim strFile As String
  Dim strAll As String
  strFile = "specialgm\names.txt"
  strMyPath = App.path
  If Right$(strMyPath, 1) <> "\" Then
    strMyPath = strMyPath & "\"
  End If
  If BlackdFileExistCheck(strMyPath & strFile) = True Then
    strAll = strMyPath & strFile
    SafeLoadSpecialGMnames strAll
  End If
End Sub


Public Sub AddSpecialGMname(str As String)
  ' add item to dictionary
  Dim res As Boolean
  specialGMnames.item(LCase(str)) = True
End Sub
Public Sub RemoveAllSpecialGMname()
  ' remove item from dictionary
  Dim res As Boolean
  specialGMnames.RemoveAll
End Sub
Public Function isSpecialGMname(ByRef str As String) As Boolean
  ' get the name from an ID
  Dim res As Boolean
  If specialGMnames.Exists(LCase(str)) = True Then
    isSpecialGMname = True
  Else
    isSpecialGMname = False
  End If
End Function



Public Sub AddMelee(idConnection As Integer, str As String)
  ' add item to dictionary
  Dim res As Boolean
  cavebotMelees(idConnection).item(LCase(str)) = True
End Sub
Public Sub RemoveAllMelee(idConnection As Integer)
  ' remove item from dictionary
  Dim res As Boolean
  cavebotMelees(idConnection).RemoveAll
End Sub
Public Function isMelee(idConnection As Integer, ByRef str As String) As Boolean
  ' get the name from an ID
  Dim res As Boolean
  If cavebotMelees(idConnection).Exists(LCase(str)) = True Then
    isMelee = True
  Else
    isMelee = False
  End If
End Function


Public Function AddIgnoredcreature(idConnection As Integer, dblID As Double) As Long
  ' add item to dictionary
  If isIgnoredcreature(idConnection, dblID) = True Then
    AddIgnoredcreature = -1
  Else
    IgnoredCreatures(idConnection).item(dblID) = True
    AddIgnoredcreature = 0
  End If
End Function
Public Sub RemoveAllIgnoredcreature(idConnection As Integer)
  ' remove item from dictionary
  IgnoredCreatures(idConnection).RemoveAll
End Sub
Public Function isIgnoredcreature(idConnection As Integer, ByRef dblID As Double) As Boolean
  ' get the name from an ID
  Dim res As Boolean
  If IgnoredCreatures(idConnection).Exists(dblID) = True Then
    isIgnoredcreature = True
  Else
    isIgnoredcreature = False
  End If
End Function



Public Sub AddExorivis(idConnection As Integer, str As String)
  ' add item to dictionary
  Dim res As Boolean
  cavebotExorivis(idConnection).item(LCase(str)) = True
End Sub
Public Sub RemoveAllExorivis(idConnection As Integer)
  ' remove item from dictionary
  Dim res As Boolean
  cavebotExorivis(idConnection).RemoveAll
End Sub
Public Function isExorivis(idConnection As Integer, ByRef str As String) As Boolean
  ' get the name from an ID
  Dim res As Boolean
  If cavebotExorivis(idConnection).Exists(LCase(str)) = True Then
    isExorivis = True
  Else
    isExorivis = False
  End If
End Function



Public Sub AddAvoid(idConnection As Integer, str As String)
  ' add item to dictionary
  Dim res As Boolean
  cavebotAvoid(idConnection).item(LCase(str)) = True
End Sub
Public Sub RemoveAllAvoid(idConnection As Integer)
  ' remove item from dictionary
  Dim res As Boolean
  cavebotAvoid(idConnection).RemoveAll
End Sub
Public Function isAvoid(idConnection As Integer, ByRef str As String) As Boolean
  ' get the name from an ID
  Dim res As Boolean
  If cavebotAvoid(idConnection).Exists(LCase(str)) = True Then
    isAvoid = True
  Else
    isAvoid = False
  End If
End Function


Public Sub AddHMM(idConnection As Integer, str As String)
  ' add item to dictionary
  Dim res As Boolean
  cavebotHMMs(idConnection).item(LCase(str)) = True
End Sub
Public Sub RemoveAllHMM(idConnection As Integer)
  ' remove item from dictionary
  Dim res As Boolean
  cavebotHMMs(idConnection).RemoveAll
End Sub
Public Function isHmm(idConnection As Integer, ByRef str As String) As Boolean
  ' get the name from an ID
  Dim res As Boolean
  If cavebotHMMs(idConnection).Exists(LCase(str)) = True Then
    isHmm = True
  Else
    isHmm = False
  End If
End Function


Public Sub AddSETUSEITEM(idConnection As Integer, str As String, strval As String)
  ' add item to dictionary
  Dim res As Boolean
  DictSETUSEITEM(idConnection).item(UCase(str)) = strval
  DictSETUSEITEM_used(idConnection) = True
End Sub
Public Sub RemoveAllSETUSEITEM(idConnection As Integer)
  ' remove item from dictionary
  Dim res As Boolean
  DictSETUSEITEM(idConnection).RemoveAll
  DictSETUSEITEM_used(idConnection) = False
  SETUSEITEM_lastX(idConnection) = 0
  SETUSEITEM_lastY(idConnection) = 0
End Sub
Public Function getSETUSEITEM(idConnection As Integer, ByRef str As String) As String
  ' get the name from an ID
  Dim res As String
  If DictSETUSEITEM(idConnection).Exists(UCase(str)) = True Then
    res = DictSETUSEITEM(idConnection).item(UCase(str))
    If res = "00 00" Then
        getSETUSEITEM = ""
    Else
        getSETUSEITEM = res
    End If
  Else
    getSETUSEITEM = ""
  End If
End Function


Public Sub AddShotType(idConnection As Integer, str As String, shottype As Long)
  ' add item to dictionary
  Dim res As Boolean
  shotTypeDict(idConnection).item(LCase(str)) = shottype
End Sub
Public Sub RemoveAllShotType(idConnection As Integer)
  ' remove item from dictionary
  Dim res As Boolean
  shotTypeDict(idConnection).RemoveAll
End Sub
Public Function getShotType(idConnection As Integer, ByRef str As String) As Long
  ' get the name from an ID
  Dim res As Boolean
  If shotTypeDict(idConnection).Exists(LCase(str)) = True Then
    getShotType = shotTypeDict(idConnection).item(LCase(str))
  Else
    getShotType = tileID_HMM
  End If
End Function



Public Sub AddExoriType(idConnection As Integer, str As String, shottype As Long)
  ' add item to dictionary
  Dim res As Boolean
  exoriTypeDict(idConnection).item(LCase(str)) = shottype
End Sub
Public Sub RemoveAllExoriType(idConnection As Integer)
  ' remove item from dictionary
  Dim res As Boolean
  exoriTypeDict(idConnection).RemoveAll
End Sub
Public Function getExoriType(idConnection As Integer, ByRef str As String) As Long
  ' get the name from an ID
  Dim res As Boolean
  If exoriTypeDict(idConnection).Exists(LCase(str)) = True Then
    getExoriType = exoriTypeDict(idConnection).item(str)
  Else
    getExoriType = 1 'exori vis
  End If
End Function

Public Sub AddGoodLoot(idConnection As Integer, str As Long)
  ' add item to dictionary
  Dim res As Boolean
  cavebotGoodLoot(idConnection).item(str) = True
End Sub
Public Sub RemoveGoodLoot(idConnection As Integer, str As Long)
  ' add item to dictionary
  Dim res As Boolean
  cavebotGoodLoot(idConnection).item(str) = False
End Sub
Public Sub RemoveAllGoodLoot(idConnection As Integer)
  ' remove item from dictionary
  Dim res As Boolean
  cavebotGoodLoot(idConnection).RemoveAll
End Sub
Public Function IsGoodLoot(idConnection As Integer, ByRef str As Long) As Boolean
  ' get the name from an ID
  Dim res As Boolean
  If LootAll(idConnection) = True Then
    If cavebotGoodLoot(idConnection).Exists(str) = True Then
      IsGoodLoot = Not cavebotGoodLoot(idConnection).item(str)
    Else
      IsGoodLoot = True
    End If
  Else
    If cavebotGoodLoot(idConnection).Exists(str) = True Then
      IsGoodLoot = cavebotGoodLoot(idConnection).item(str)
    Else
      IsGoodLoot = False
    End If
  End If
End Function




Public Sub AddIDLine(idConnection As Integer, ByRef lineID As Long, ByRef strLine As String)
  ' add item to dictionary
  Dim res As Boolean
  cavebotScript(idConnection).item(lineID + 1) = strLine
End Sub
Public Sub RemoveIDLine(idConnection As Integer, ByRef lineID As Long)
  ' remove item from dictionary
  Dim res As Boolean
  If cavebotScript(idConnection).Exists(lineID + 1) = True Then
    cavebotScript(idConnection).Remove (lineID + 1)
  End If
End Sub
Public Function GetStringFromIDLine(idConnection As Integer, ByRef lineID As Long) As String
  ' get the name from an ID
  On Error GoTo goterr
  Dim res As Boolean
  If cavebotScript(idConnection).Exists(lineID + 1) = True Then
    GetStringFromIDLine = cavebotScript(idConnection).item(lineID + 1)
  Else
    GetStringFromIDLine = "?"
  End If
  Exit Function
goterr:
  LogOnFile "errors.txt", "Error atGetStringFromIDLine (" & _
   CStr(idConnection) & ", " & CStr(lineID) & ") , Err number : " & CStr(Err.Number) & _
   " ; Err description : " & Err.Description
  GetStringFromIDLine = "?"
End Function

Public Function DoOneStack(idConnection As Integer) As Long
  Dim res1 As TypeSearchItemResult2
  Dim res2 As TypeSearchItemResult2
  Dim i As Long
  Dim j As Long
  Dim limitJ As Long
  Dim blnStop As Boolean
  Dim tileID As Long
  Dim amount1 As Byte
  Dim amount2 As Byte
  Dim nextLimit As Long
 ' Dim cPacket(16) As Byte
  Dim sCheat As String
 ' SendLogSystemMessageToClient idConnection, "Stack order received"
  If (pauseStacking(idConnection) > 0) Then
    pauseStacking(idConnection) = pauseStacking(idConnection) - 1
    DoOneStack = 1
    Exit Function
  Else
    pauseStacking(idConnection) = StackingWaitTurns
  End If
  
  nextLimit = -1
nextIter:
  res1.foundcount = 0
  res1.bpID = &HFF
  res1.slotID = &HFF
  res1.b4 = 0
  ' search first item badly stacked
  blnStop = False
  For i = 0 To HIGHEST_BP_ID
    limitJ = (Backpack(idConnection, i).used) - 1
    For j = 0 To limitJ
      tileID = GetTheLong(Backpack(idConnection, i).item(j).t1, Backpack(idConnection, i).item(j).t2)
      If DatTiles(tileID).stackable = True And _
       Backpack(idConnection, i).item(j).t3 < 100 Then
        If ((i * 100) + j) > nextLimit Then
          res1.foundcount = 1
          res1.bpID = CByte(i)
          res1.slotID = CByte(j)
          res1.b1 = Backpack(idConnection, i).item(j).t1
          res1.b2 = Backpack(idConnection, i).item(j).t2
          amount1 = Backpack(idConnection, i).item(j).t3
          res1.b4 = Backpack(idConnection, i).item(j).t4
          nextLimit = (i * 100) + j 'next iteration won't be the same
          blnStop = True
          Exit For
        End If
      End If
    Next j
    If blnStop = True Then
      Exit For
    End If
  Next i
  If res1.foundcount = 0 Then
   ' SendLogSystemMessageToClient idConnection, "All is now stacked to the max"
    DoEvents
    DoOneStack = 0
    Exit Function
  End If
  
  res2.foundcount = 0
  res2.bpID = &HFF
  res2.slotID = &HFF
  res2.b4 = 0
  ' search second item badly stacked
  blnStop = False
  For i = 0 To HIGHEST_BP_ID
    limitJ = (Backpack(idConnection, i).used) - 1
    For j = 0 To limitJ
      tileID = GetTheLong(Backpack(idConnection, i).item(j).t1, Backpack(idConnection, i).item(j).t2)
      If DatTiles(tileID).stackable = True And _
       Backpack(idConnection, i).item(j).t3 < 100 Then
         If Not (CByte(i) = res1.bpID And _
          CByte(j) = res1.slotID) Then
           If res1.b1 = Backpack(idConnection, i).item(j).t1 And res1.b2 = Backpack(idConnection, i).item(j).t2 Then
             res2.foundcount = 1
             res2.bpID = CByte(i)
             res2.slotID = CByte(j)
             amount2 = Backpack(idConnection, i).item(j).t3
             res2.b4 = Backpack(idConnection, i).item(j).t4
             blnStop = True
             Exit For
           End If
         End If
      End If
    Next j
    If blnStop = True Then
      Exit For
    End If
  Next i
  If res2.foundcount = 0 Then
    GoTo nextIter
  Else
    '0F 00 78 FF FF 40 00 01 BC 0D 01 FF FF 40 00 03 06
   ' SendLogSystemMessageToClient idConnection, "Stacking tileID " & GoodHex(res1.b1) & " " & GoodHex(res1.b2) & " [bpID " & GoodHex(res1.bpID) & " slotID " & _
     GoodHex(res1.slotID) & " (x" & amount1 & ")] with [bpID " & GoodHex(res2.bpID) & " slotID " & _
     GoodHex(res2.slotID) & " (x" & amount2 & ")]"
    DoEvents
'    cPacket(0) = &HF
'    cPacket(1) = &H0
'    cPacket(2) = &H78
'    cPacket(3) = &HFF
'    cPacket(4) = &HFF
'    cPacket(5) = &H40 + res1.bpID
'    cPacket(6) = &H0
'    cPacket(7) = res1.slotID
'    cPacket(8) = res1.b1
'    cPacket(9) = res1.b2
'    cPacket(10) = res1.slotID
'    cPacket(11) = &HFF
'    cPacket(12) = &HFF
'    cPacket(13) = &H40 + res2.bpID
'    cPacket(14) = &H0
'    cPacket(15) = res2.slotID
'    cPacket(16) = amount1
    'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & frmMain.showAsStr2(cPacket, True)
'    frmMain.UnifiedSendToServerGame idConnection, cPacket, True
    
    sCheat = "78 FF FF " & GoodHex(&H40 + res1.bpID) & " 00 " & GoodHex(res1.slotID) & " " & _
     GoodHex(res1.b1) & " " & GoodHex(res1.b2) & " "
     If (TibiaVersionLong = 760) Then
     sCheat = sCheat & GoodHex(res1.slotID) & " "
     End If
     sCheat = sCheat & "FF FF " & GoodHex(&H40 + res2.bpID) & " 00 " & _
     GoodHex(res2.slotID) & " " & GoodHex(amount1)


     
    SafeCastCheatString "DoOneStack1", idConnection, sCheat
    
    'DoEvents
    DoOneStack = 1
  End If
End Function

Public Sub GetProcessAllProcessIDs()
  Dim i As Integer
  'Exit Sub ' SENSELESS FUNCTION NOW
  For i = 1 To MAXCLIENTS
    GetProcessIDs i
  Next i
End Sub
Public Sub GetProcessIDs(Sid As Integer)
  Dim i As Integer
  Dim compareID As String
  Dim tibiaclient As Long
  'Dim hWndDesktop As Long
  Dim IsConnected As Long
  Dim IsConnectedByte As Byte
  Dim pidMatch As String
  Dim conMatch As String
  Dim usedWaitX As Boolean
  If myID(Sid) = 0 Then
    Exit Sub
  End If
  usedWaitX = False
  debugPIDs(Sid) = ""
anotherTry:
  pidMatch = "NO"
  conMatch = "NO"
  ProcessID(Sid) = -1
  'hWndDesktop = GetDesktopWindow()
  tibiaclient = 0
  i = 0
  Do
    i = i + 1
    If i > MAXCLIENTS Then
        Exit Do
    Else
        tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
        If tibiaclient = 0 Then
            Exit Do
        Else
            debugPIDs(Sid) = debugPIDs(Sid) & "tibiaclient = FindWindowEx(0, " & CStr(tibiaclient) & ", " & tibiaclassname & ", vbNullString)" & vbCrLf
            debugPIDs(Sid) = debugPIDs(Sid) & "Result: tibiaclient = " & CStr(tibiaclient) & vbCrLf
            compareID = CDbl(Memory_ReadLong(adrNum, tibiaclient))
            If compareID = myID(Sid) Then
                pidMatch = "YES"
                debugPIDs(Sid) = debugPIDs(Sid) & "Found player ID: " & CStr(compareID) & " (MATCH)" & vbCrLf
                If TibiaVersionLong <= 760 Then
                    IsConnected = Memory_ReadLong(adrConnected, tibiaclient)
                Else
                    IsConnectedByte = Memory_ReadByte(adrConnected, tibiaclient)
                    IsConnected = CLng(IsConnectedByte)
                End If
                If IsConnected <> 0 Then
                    conMatch = "YES"
                    ProcessID(Sid) = tibiaclient
                    debugPIDs(Sid) = debugPIDs(Sid) & "REQUEST #" & Sid & " OK" & vbCrLf
                    'Debug.Print debugPIDs(Sid)
                    Exit Sub ' ID found and it is connected -> end search
                Else
                    debugPIDs(Sid) = debugPIDs(Sid) & "Connected: " & "NO" & vbCrLf
                End If
            Else
                    debugPIDs(Sid) = debugPIDs(Sid) & "Found player ID: " & CStr(compareID) & " (NO MATCH)" & vbCrLf
            End If
        End If
    End If
  Loop
    debugPIDs(Sid) = debugPIDs(Sid) & "REQUEST #" & Sid & " FAILED!" & vbCrLf
    debugPIDs(Sid) = debugPIDs(Sid) & "(Expected player ID: " & CStr(myID(Sid)) & ")" & vbCrLf
    If usedWaitX = False Then
        usedWaitX = True
        wait (200) ' wait a bit and try again
        GoTo anotherTry
    End If
    'Debug.Print debugPIDs(Sid)
End Sub

Public Function ParseString(ByRef entireLine As String, ByRef frompos As Long, toEnd As Long, ByRef limitChar As String) As String
  Dim pos As Long
  Dim newChar As String
  Dim res As String
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  pos = frompos
  res = ""
  Do
    If pos > toEnd Then
      Exit Do
    Else
      newChar = Mid(entireLine, pos, 1)
      If newChar = limitChar Then
        Exit Do
      Else
        res = res & newChar
        pos = pos + 1
      End If
    End If
  Loop
  frompos = pos
  ParseString = res
  Exit Function
goterr:
   frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error during ParseString. Number: " & Err.Number & " Description: " & Err.Description & " Source: " & Err.Source
   ParseString = ""
End Function

Private Sub SkipBlanks(ByRef entireLine As String, ByRef frompos As Long, toEnd As Long)
  ' skip spaces and ,
  Dim pos As Long
  Dim newChar As String
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  pos = frompos
  Do
    If pos > toEnd Then
      Exit Do
    Else
      newChar = Mid(entireLine, pos, 1)
      If newChar <> " " And newChar <> "," Then
        Exit Do
      Else
        pos = pos + 1
      End If
    End If
  Loop
  frompos = pos
  Exit Sub
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error during SkipBlanks. Number: " & Err.Number & " Description: " & Err.Description & " Source: " & Err.Source

End Sub

Public Function MyBattleListPosition(Sid) As Long
  Dim c1 As Long
  Dim id As Double
  Dim res As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  res = -1
  For c1 = 0 To LAST_BATTLELISTPOS
    id = CDbl(Memory_ReadLong(adrNChar + (CharDist * c1), ProcessID(Sid)))
    If myID(Sid) = id Then
      res = c1
      Exit For
    End If
  Next c1
  MyBattleListPosition = res
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error during MyBattleListPosition. Number: " & Err.Number & " Description: " & Err.Description & " Source: " & Err.Source
  MyBattleListPosition = -1
End Function

Public Sub PerformMove(Sid As Integer, parx As Long, pary As Long, parz As Long)
' adrXgo = &H49D070 ' goto this x
' adrYgo = &H49D06C ' goto this y
' adrGo = &H49D0DC ' start goto process
  Dim b1 As Byte
  Dim b2 As Byte
  Dim pid As Long
  Dim aRes As Long
  Dim myBpos As Long
  Dim xinc As Long
  Dim yinc As Long
  Dim sCheat As String
  Dim cPacket() As Byte
  Dim inRes As Integer
  Dim gotDestChange As Boolean
  Dim queue As String
  Dim strDebug As String
  Dim cfRes As TypeChangeFloorResult
  Dim X As Long
  Dim y As Long
  Dim z As Long
  Dim shouldBeExact As Boolean
  Dim status As Integer
  Dim completed As Boolean
  Dim iterac As Integer
  Dim awesomeStatus As Integer
  Dim tmpByte As Byte
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  status = 1
  completed = False
  iterac = 0
  awesomeStatus = 0
  Do
  Select Case status
  
Case 1
  ' initial state
  strDebug = "01"
  X = parx
  y = pary
  z = parz
  shouldBeExact = False
  pid = ProcessID(Sid)
  If pid >= 0 Then
    status = 2
  Else
    frmCavebot.lblInfo.Caption = "Error. Your pID=" & pid & " (!)"
    completed = True
  End If
Case 2
  ' there is any active exception?
  strDebug = strDebug & " > 02"
  If cancelAllMove(Sid) > GetTickCount() Then ' uh'ing ?
    status = 32
  ElseIf makingRune(Sid) = True Then ' making a rune ?
    status = 32
  ElseIf GetTickCount() < lootTimeExpire(Sid) Then ' looting ?
    status = 32
  Else
  ' ok to move
    status = 3
  End If
Case 3
  '  should reset moveretry? ?
  strDebug = strDebug & " > 03"
  If (X = lastDestX(Sid)) And (y = lastDestY(Sid)) And _
   (z = lastDestZ(Sid)) Then
    ' no
    lastAttackedIDstatus(Sid) = lastAttackedID(Sid)
    status = 4
  ElseIf (lastAttackedID(Sid) <> 0) And (lastAttackedID(Sid) = lastAttackedIDstatus(Sid)) Then
    'special no
    lastDestX(Sid) = X
    lastDestY(Sid) = y
    lastDestZ(Sid) = z
    lastAttackedIDstatus(Sid) = lastAttackedID(Sid)
    status = 4
  Else
    ' yes
    lastAttackedIDstatus(Sid) = lastAttackedID(Sid)
    status = 7
  End If
Case 4
  ' process same destination
  strDebug = strDebug & " > 04"
  moveRetry(Sid) = moveRetry(Sid) + frmCavebot.TimerScript.Interval
  status = 5
Case 5
  ' standing in same point ?
  strDebug = strDebug & " > 05"
  If (myX(Sid) = lastX(Sid)) And (myY(Sid) = lastY(Sid)) And _
   (myZ(Sid) = lastZ(Sid)) Then
    ' yes
    status = 6
  Else
    ' no
    lastX(Sid) = myX(Sid)
    lastY(Sid) = myY(Sid)
    lastZ(Sid) = myZ(Sid)
    ' no, but attacking
    If lastAttackedID(Sid) <> 0 Then
      status = 6
    Else
      ignoreNext(Sid) = CteMoveDelay + GetTickCount()
      status = 32
    End If
  End If
Case 6
  ' should wait because recent move order?
  strDebug = strDebug & " > 06"
  If ignoreNext(Sid) > GetTickCount() Then
    ' yes
    status = 32
  Else
   ' no
    status = 8
  End If
Case 7
  ' process destination change
  strDebug = strDebug & " > 07"
  lastDestX(Sid) = X
  lastDestY(Sid) = y
  lastDestZ(Sid) = z
  moveRetry(Sid) = 0
  ignoreNext(Sid) = GetTickCount() - 1
  status = 32
Case 8
  ' destination is the same
  strDebug = strDebug & " > 08"
  xinc = X - myX(Sid)
  yinc = y - myY(Sid)
  If z <> myZ(Sid) Then
    ' must change floor
    status = 12
  ElseIf (moveRetry(Sid) > 10000) And (z = myZ(Sid)) And (onDepotPhase(Sid) = 2) Then
    ' must choose other depot
    status = 10
  ElseIf moveRetry(Sid) > TimeToGiveTrapAlarm Then
    ' must give trapalarm
    status = 11
  Else
    ' process move
    status = 9
  End If
Case 9
  ' attacking or not attacking?
  strDebug = strDebug & " > 09"
  If (lastAttackedID(Sid) = 0) Then
    status = 27
  Else
    status = 21
  End If
Case 10
  ' choose other depot
  strDebug = strDebug & " > 10 : Choosing other depot"
  onDepotPhase(Sid) = 0 'changed from 1 to 0 in 8.74
  If exeLine(Sid) > 0 Then
    'exeLine(Sid) = exeLine(Sid) - 1
    updateExeLine Sid, -1, True
  End If
  moveRetry(Sid) = 0
  status = 32
Case 11
  ' Trap alarm
  strDebug = strDebug & " > 11 : Trap alarm - Trying a reposition"
  If cavebotOnTrapGiveAlarm(Sid) = True Then
    If frmRunemaker.ChkDangerSound.Value = 1 Then
      If PlayTheDangerSound = False Then
        aRes = GiveGMmessage(Sid, "WARNING : YOU ARE TRAPPED !", "BlackdProxy")
        DoEvents
        aRes = SendLogSystemMessageToClient(Sid, "BlackdProxy: To deactivate alarm do Exiva cancel")
        DoEvents
      End If
      ChangePlayTheDangerSound True
    Else
      aRes = SendSystemMessageToClient(Sid, "WARNING : YOU ARE TRAPPED !")
      DoEvents
    End If
  End If
  moveRetry(Sid) = 0
  If AllowRepositionAtTrap(Sid) = 1 Then
    RepositionScriptAtTrap Sid
    DoRandomMove Sid
  End If
  status = 100
Case 12
  ' change floor
  strDebug = strDebug & " > 12"
  If z < myZ(Sid) Then
    cfRes = PerformMoveUp(Sid, X, y, z)
  Else
    cfRes = PerformMoveDown(Sid, X, y, z)
  End If
  'myres.result=0 req_wait
  'myres.result=1 req_move
  'myres.result=2 req_click
  'myres.result=3 req_shovel
  'myres.result=4 req_rope
  'myres.result=5 req_random_move
  'myres.result>&H60 req_force_move
  Select Case cfRes.result
  Case &H0
    status = 32
  Case &H1
    status = 13
  Case &H2
    status = 14
  Case &H3
    status = 15
  Case &H4
    status = 16
  Case &H5
    status = 33
  Case Else
    status = 34
  End Select
Case 13
  strDebug = strDebug & " > 13"
  X = cfRes.X
  y = cfRes.y
  z = cfRes.z
  shouldBeExact = True
  status = 8
Case 14
  If ((Abs(cfRes.X - myX(Sid)) > 1) Or (Abs(cfRes.y - myY(Sid)))) > 1 Then
    strDebug = strDebug & " > 14 : Right Click required move"
    X = cfRes.X
    y = cfRes.y
    z = cfRes.z
    shouldBeExact = False
    status = 9
  Else
    strDebug = strDebug & " > 14 : Doing right click"
    PerformUseItem Sid, cfRes.X, cfRes.y, cfRes.z
    ignoreNext(Sid) = GetTickCount() + CteMoveDelay
    status = 100
  End If
Case 15
  If ((Abs(cfRes.X - myX(Sid)) > 1) Or (Abs(cfRes.y - myY(Sid)))) > 1 Then
    strDebug = strDebug & " > 14 : Shovel required move"
    X = cfRes.X
    y = cfRes.y
    z = cfRes.z
    shouldBeExact = False
    status = 9
  Else
    strDebug = strDebug & " > 15"
    aRes = PerformUseMyItem(Sid, LowByteOfLong(tileID_Shovel), HighByteOfLong(tileID_Shovel), cfRes.X, cfRes.y, cfRes.z, True, True)
    If aRes = 0 Then
      status = 18
    Else
        aRes = PerformUseMyItem(Sid, LowByteOfLong(tileID_LightShovel), HighByteOfLong(tileID_LightShovel), cfRes.X, cfRes.y, cfRes.z, , True)
        If aRes = 0 Then
          status = 18
        Else
          status = 17
        End If
    End If
  
  End If
Case 16
  If ((Abs(cfRes.X - myX(Sid)) > 1) Or (Abs(cfRes.y - myY(Sid)))) > 1 Then
    strDebug = strDebug & " > 16 : Rope required move"
    X = cfRes.X
    y = cfRes.y
    z = cfRes.z
    shouldBeExact = False
    status = 9
  Else
    strDebug = strDebug & " > 16"
    aRes = PerformUseMyItem(Sid, LowByteOfLong(tileID_Rope), HighByteOfLong(tileID_Rope), cfRes.X, cfRes.y, cfRes.z)
    If aRes = 0 Then
      status = 20
    Else
      aRes = PerformUseMyItem(Sid, LowByteOfLong(tileID_LightRope), HighByteOfLong(tileID_LightRope), cfRes.X, cfRes.y, cfRes.z)
      If aRes = 0 Then
        status = 20
      Else
        status = 19
      End If
    End If
  End If
Case 17
  ' Trap alarm
  strDebug = strDebug & " > 17 : Trap alarm - No shovel"
  If frmRunemaker.ChkDangerSound.Value = 1 Then
    If PlayTheDangerSound = False Then
      aRes = GiveGMmessage(Sid, "WARNING : YOU ARE TRAPPED ! (No shovel)", "BlackdProxy")
      DoEvents
      aRes = SendLogSystemMessageToClient(Sid, "BlackdProxy: To deactivate alarm do Exiva cancel")
      DoEvents
    End If
    ChangePlayTheDangerSound True
  Else
    aRes = SendSystemMessageToClient(Sid, "WARNING : YOU ARE TRAPPED ! (No shovel)")
    DoEvents
  End If
  moveRetry(Sid) = 0
  status = 100
Case 18
  strDebug = strDebug & " > 18 : Using shovel"
  ignoreNext(Sid) = GetTickCount() + CteMoveDelay
  status = 100
Case 19
  ' Trap alarm
  strDebug = strDebug & " > 19 : Trap alarm - No rope"
  If frmRunemaker.ChkDangerSound.Value = 1 Then
    If PlayTheDangerSound = False Then
      aRes = GiveGMmessage(Sid, "WARNING : YOU ARE TRAPPED ! (No rope)", "BlackdProxy")
      DoEvents
      aRes = SendLogSystemMessageToClient(Sid, "BlackdProxy: To deactivate alarm do Exiva cancel")
      DoEvents
    End If
    ChangePlayTheDangerSound True
  Else
    aRes = SendSystemMessageToClient(Sid, "WARNING : YOU ARE TRAPPED ! (No rope)")
    DoEvents
  End If
  moveRetry(Sid) = 0
  status = 100
Case 20
  strDebug = strDebug & " > 20 : Using rope"
  ignoreNext(Sid) = GetTickCount() + CteMoveDelay
  status = 100
Case 21
  ' attacking
  strDebug = strDebug & " > 21"
  If (Abs(xinc) < 2) And (Abs(yinc) < 2) Then
    moveRetry(Sid) = 0
    status = 22
  ElseIf moveRetry(Sid) < 1500 Then
    status = 24
  Else
    status = 23
  End If
Case 22
  ' close to the monster we are attacking
  strDebug = strDebug & " > 22"
  moveRetry(Sid) = 0
  status = 32
Case 23
  ' a* short move 1
  strDebug = strDebug & " > 23"
  If (xinc < 10) And (xinc > -9) And (yinc < 8) And (yinc > -7) Then
    aRes = FindBestPath(Sid, xinc, yinc, False)
    '#if can't find short best path, pick other target
    If aRes <> 0 Then
      status = 26
    Else
      status = 25
    End If
  Else
    '#if target is too far, pick other target
    status = 26
  End If
Case 24
  ' click fast move
  strDebug = strDebug & " > 24 : Doing Fast move"
  DoUnifiedClickMove Sid, X, y, z

  status = 100
Case 25
  ' A* Short move1
  strDebug = strDebug & " > 25 : Doing A*Short move"
  moveRetry(Sid) = 0
  'ignoreNext(sid) = GetTickCount() + CteMoveDelay
  status = 100
Case 26
  ' change target
  strDebug = strDebug & " > 26 : Target rejected"
  lastAttackedID(Sid) = 0
  status = 100
Case 27
  ' not attacking
  strDebug = strDebug & " > 27"
  xinc = X - myX(Sid)
  yinc = y - myY(Sid)
  If ((Abs(xinc) < 2) And (Abs(yinc) < 2)) Then
    status = 31
  ElseIf ((xinc < 10) And (xinc > -9) And (yinc < 8) And (yinc > -7)) Then
    status = 30
  ElseIf (moveRetry(Sid) < 5000) Then
    status = 24
  Else
    status = 28
  End If
Case 28
    AstarBig Sid, myX(Sid), myY(Sid), X, y, myZ(Sid), False
    If ((RequiredMoveBuffer(Sid) = "") Or (RequiredMoveBuffer(Sid) = "X")) Then
      If publicDebugMode = True Then
        aRes = SendLogSystemMessageToClient(Sid, "[Debug] Big map failed to move to " & X & "," & y & "," & z)
        DoEvents
      End If
      strDebug = strDebug & " > 28"
      status = 11
    Else
      OptimizeBuffer Sid
      ExecuteBuffer Sid
      If publicDebugMode = True Then
        aRes = SendLogSystemMessageToClient(Sid, "[Debug] Processing big map far distance move to " & X & "," & y & "," & z)
        DoEvents
      End If
      strDebug = strDebug & " > 28"
      status = 29
    End If
Case 29
  ' A* Long Move completed
  strDebug = strDebug & " > 29 : Doing A* Long Move"
  ignoreNext(Sid) = GetTickCount() + (CteMoveDelay * 2)
  ' moveRetry(sid) = 5000
  status = 100
Case 30
  ' a* short move 2
  strDebug = strDebug & " > 30"
  If (xinc < 10) And (xinc > -9) And (yinc < 8) And (yinc > -7) Then
    If shouldBeExact = True Then
      aRes = FindBestPath(Sid, xinc, yinc, False)
    Else
      aRes = FindBestPathV2(Sid, xinc, yinc, False)
    End If
    '#if can't find short best path, think alternative plan
    If aRes <> 0 Then
      If onDepotPhase(Sid) > 0 Then
        status = 10
      ElseIf lastAttackedID(Sid) <> 0 Then
        status = 24
      Else ' if not attacking, try a long path
        status = 28
      End If
    Else
      status = 35
    End If
  Else
    'try click move
    status = 24
  End If
Case 31
  ' very near move
  strDebug = strDebug & " > 31 : Doing very near move"
  tmpByte = &H0
  If xinc = -1 Then
    If yinc = -1 Then
      tmpByte = &H6D
    ElseIf yinc = 1 Then
      tmpByte = &H6C
    Else
      tmpByte = &H68
    End If
  ElseIf xinc = 1 Then
    If yinc = -1 Then
      tmpByte = &H6A
    ElseIf yinc = 1 Then
      tmpByte = &H6B
    Else
      tmpByte = &H66
    End If
  Else
    If yinc = -1 Then
      tmpByte = &H65
    ElseIf yinc = 1 Then
      tmpByte = &H67
    End If
  End If
  If tmpByte = &H0 Then
    DoManualMove Sid, tmpByte
    strDebug = strDebug & " > 31 : Waiting for floor change"
  Else
      DoManualMove Sid, tmpByte
    strDebug = strDebug & " > 31 : Doing very near move (" & GoodHex(tmpByte) & ")"
    moveRetry(Sid) = 0
    ignoreNext(Sid) = GetTickCount() + CteMoveDelay
  End If

  status = 100
Case 32
  ' end state
  strDebug = strDebug & " > 32 : Waiting..."
  If extremeDebugMode = False Then
  
    completed = True ' comment to log this too
  End If
  status = 100
Case 33
  strDebug = strDebug & " > 33 : Random move"
  DoRandomMove Sid
  ignoreNext(Sid) = GetTickCount() + CteMoveDelay
  status = 100
Case 34
  strDebug = strDebug & " > 34 : Force step"
  DoManualMove Sid, cfRes.result
  ignoreNext(Sid) = GetTickCount() + CteMoveDelay
  status = 100
Case 35
  ' A* Short move2
  strDebug = strDebug & " > 35 : Doing A*Short move v2"
  moveRetry(Sid) = 0
  ignoreNext(Sid) = GetTickCount() + CteMoveDelay
  status = 100
Case 100
  If publicDebugMode = True Then
    aRes = SendLogSystemMessageToClient(Sid, "[Debug] Status :" & strDebug)
    DoEvents
  End If
  completed = True
Case Else
  awesomeStatus = status
End Select
iterac = iterac + 1
If (iterac > 20) Then ' if there is no result after 20 iterations, then something is failing
  completed = True ' this avoid computer lock at least
  ' report and log the error
  If awesomeStatus = 0 Then
    LogOnFile "errors.txt", "Infinite loop detected (" & CStr(iterac) & " iterations at client " & CStr(Sid) & ") Trace : " & strDebug
    aRes = SendLogSystemMessageToClient(Sid, "Critical error on cavebot AI: Infinite loop detected. Details in errors.txt . Please report to daniel@blackdtools.com")
    DoEvents
  Else
    LogOnFile "errors.txt", "Status out of logic detected (status = " & CStr(awesomeStatus) & " at client " & CStr(Sid) & ") Trace : " & strDebug
    aRes = SendLogSystemMessageToClient(Sid, "Critical error on cavebot AI: Status out of logic detected (" & CStr(awesomeStatus) & ") Details in errors.txt . Please report to daniel@blackdtools.com")
    DoEvents
  End If
End If
Loop Until (completed = True)
  Exit Sub
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error during PerformMove. Number: " & Err.Number & " Description: " & Err.Description & " Source: " & Err.Source
End Sub

Public Function ProcessRawCondition(ByVal part1 As String, ByVal opstr As String, ByVal part2 As String) As Boolean
  On Error GoTo goterr
  Dim res As Boolean
  res = False
  Select Case opstr
    Case "number="
      If CDbl(part1) = CDbl(part2) Then
        res = True
      End If
    Case "number<="
      If CDbl(part1) <= CDbl(part2) Then
        res = True
      End If
    Case "number>="
      If CDbl(part1) >= CDbl(part2) Then
        res = True
      End If
    Case "number<>"
      If CDbl(part1) <> CDbl(part2) Then
        res = True
      End If
    Case "number<"
      If CDbl(part1) < CDbl(part2) Then
        res = True
      End If
    Case "number>"
      If CDbl(part1) > CDbl(part2) Then
        res = True
      End If
    Case "string="
      If part1 = part2 Then
        res = True
      End If
    Case "string<>"
      If part1 <> part2 Then
        res = True
      End If
    Case Else
      res = False
  End Select
  ProcessRawCondition = res
  Exit Function
goterr:
  ProcessRawCondition = False
End Function
Public Function ProcessCondition(Sid As Integer, currLine As String, pos As Long, lenCurrLine As Long, Optional justReturnLine As Boolean = False) As Long
  Dim part1 As String
  Dim opstr As String
  Dim part2 As String
  Dim actionLine As String
  Dim res As Boolean
  Dim aRes As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  res = False
  part1 = ParseString(currLine, pos, lenCurrLine, "(")
  pos = pos + 1
  part1 = ParseString(currLine, pos, lenCurrLine, "#")
  pos = pos + 1
  opstr = ParseString(currLine, pos, lenCurrLine, "#")
  pos = pos + 1
  part2 = ParseString(currLine, pos, lenCurrLine, ")")
  pos = pos + 1
  SkipBlanks currLine, pos, lenCurrLine
  actionLine = ParseString(currLine, pos, lenCurrLine, " ")
  pos = pos + 1
  actionLine = ParseString(currLine, pos, lenCurrLine, " ")
  res = ProcessRawCondition(part1, opstr, part2)
  If res = True Then
     If justReturnLine = True Then
        ProcessCondition = actionLine
        Exit Function
     Else
        If IsNumeric(actionLine) = False Then
            aRes = GiveGMmessage(Sid, "Cavebot syntax error: expecting to jump to a line number. You typed something can NOT be converted to a number: " & actionLine, "Blackd Proxy")
            DoEvents
            frmCavebot.TurnCavebotState Sid, False
        Else
           ' exeLine(Sid) = actionLine
            updateExeLine Sid, actionLine, False
            If publicDebugMode = True Then
              aRes = SendLogSystemMessageToClient(Sid, "Condition (" & part1 & " " & opstr & " " & part2 & ") = TRUE")
              DoEvents
            End If
        End If
    End If
  Else
   If justReturnLine = True Then
           ProcessCondition = -1
        Exit Function
   Else
   ' exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(Sid, "Condition (" & part1 & " " & opstr & " " & part2 & ") = FALSE")
      DoEvents
    End If
    End If
  End If
  ProcessCondition = 0
  Exit Function
goterr:
  ProcessCondition = -1
End Function
Public Sub ProcessScriptLine(Sid As Integer)
  Dim currLineNumber As Long
  Dim currLine As String
  Dim lenCurrLine As Long
  Dim mainCommand As String
  Dim pos As Long
  Dim param1 As String
  Dim param2 As String
  Dim param3 As String
  Dim val1 As Long
  Dim val2 As Long
  Dim val3 As Long
  Dim tileID As Long
  Dim paramTile As String
  Dim b1 As Byte
  Dim b2 As Byte
  Dim aRes As Long
  Dim mytime As Long
  Dim am As Long
  Dim fastM As Boolean
  Dim attackRes As Long
  Dim spellInfo As TypeSpellKill
  Dim mobName As String
  Dim continueP As Boolean
  Dim tmpHP As Long
  Dim stringParts() As String
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  fastM = False
  mytime = GetTickCount()
  
        If DoingNewLoot(Sid) = True Then
            If mytime > DoingNewLootMAXGTC(Sid) Then
                DoingNewLoot(Sid) = False
            End If
        End If
        
  ' process events
  If ((DangerGM(Sid) = True) Or (DangerPK(Sid) = True)) And cavebotOnDanger(Sid) <> -1 Then
   ' exeLine(Sid) = cavebotOnDanger(Sid)
    updateExeLine Sid, cavebotOnDanger(Sid), False
    cavebotOnDanger(Sid) = -1
  End If
  If CheatsPaused(Sid) = True Then
    Exit Sub
  End If
  If (DangerGM(Sid) = True) Then
    If (cavebotOnGMclose(Sid) = True) Then
      frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "#Client " & Sid & " (" & CharacterName(Sid) & ") closed by script#"
      GiveServerError "Client closed - condition onGMcloseConnection was activated: " & DangerGMname(Sid), Sid
      DoEvents
      ReconnectionStage(Sid) = 10
      frmMain.DoCloseActions Sid
      ReconnectionStage(Sid) = 0
      DoEvents
    ElseIf cavebotOnGMpause(Sid) = True Then
      logoutAllowed(Sid) = 1200000 + GetTickCount() ' logout allowed in the next 20 minutes
      aRes = ChangePauseStatus(Sid, True, False)
    End If
  ElseIf (DangerPlayer(Sid) = True) Then
   If cavebotOnPLAYERpause(Sid) = True Then
      aRes = SendLogSystemMessageToClient(Sid, "Cheats paused (except autoheal) because player: " & DangerPlayerName(Sid))
      DoEvents
      ChangePlayTheDangerSound True
      aRes = ChangePauseStatus(Sid, True, True)
      DangerPlayer(Sid) = False
    End If
  End If

  'avoid moves while uhing
  If cancelAllMove(Sid) > GetTickCount() Then
    Exit Sub
  End If
  
  'avoid conflict with runemaker
  If makingRune(Sid) = True Then
    lootTimeExpire(Sid) = 0
    Exit Sub
  End If
  
  'avoid loot / attack during dropItemsPhase
    If onDepotPhase(Sid) = 0 Then
        If lootTimeExpire(Sid) > mytime Then
            If Not (requestLootBp(Sid) = &HFF) Then
                aRes = LootGoodItems(Sid)
                DoEvents
                Exit Sub
            End If
        End If
        
        
        
        
        
        
    
        If ((DoingNewLoot(Sid) = False) And (lastAttackedID(Sid) = 0)) Then
        
            aRes = ChooseBestLoot(Sid)
            If aRes > -1 Then
                DoingNewLoot(Sid) = True
                DoingNewLootX(Sid) = Looter(Sid).points(aRes).X
                DoingNewLootY(Sid) = Looter(Sid).points(aRes).y
                DoingNewLootZ(Sid) = Looter(Sid).points(aRes).z
                Looter(Sid).points(aRes).X = 0
                Looter(Sid).points(aRes).y = 0
                Looter(Sid).points(aRes).z = 0
                Looter(Sid).points(aRes).addedTime = 0
                Looter(Sid).points(aRes).expireGtc = 0
                DoingNewLootMAXGTC(Sid) = mytime + MAXTIMETOREACHCORPSE(Sid)
            End If
        End If

        
    
        
        If DoingNewLoot(Sid) = False Then
            ' process attack configured creatures
            If (((mytime - CavebotTimeWithSameTarget(Sid)) > maxAttackTimeCHAOS(Sid)) And (EnableMaxAttackTime(Sid) = True)) Then
                ChaotizeNextMaxAttackTime Sid
                If TrainerOptions(Sid).misc_stoplowhp = 0 Then
                    ' IGNORE CURRENT TARGET
                    If lastAttackedID(Sid) <> 0 Then
                        If publicDebugMode = True Then
                            aRes = SendLogSystemMessageToClient(Sid, "Creature ID #" & CStr(lastAttackedID(Sid)) & _
                            " ( " & GetNameFromID(Sid, lastAttackedID(Sid)) & " ) will be ignored (because too much time)")
                        End If
                        aRes = AddIgnoredcreature(Sid, lastAttackedID(Sid))
                        aRes = MeleeAttack(Sid, 0, True)
                        lastAttackedID(Sid) = 0
                    End If
                Else
                    ' "Stop attacking target until regen" is ENABLED ...
                    ' RESET THE TIMER AND CONTINUE WITH OLD PROCEDURE
                    CavebotTimeWithSameTarget(Sid) = mytime
                    If publicDebugMode = True Then
                        aRes = SendLogSystemMessageToClient(Sid, "The attack will continue because " & _
                        "-Stop attacking target until regen- is enabled," & _
                        " else the current target would have been ignored")
                    End If
                End If
            End If
            attackRes = ProcessAttacks(Sid)
            If (attackRes = 1) And (prevAttackState(Sid) = False) Then
                moveRetry(Sid) = 0
            End If
        
        End If ' doingnewloot(sid)=false
        
        If lastAttackedID(Sid) <> 0 Then
            If lastAttackedID(Sid) <> previousAttackedID(Sid) Then
                previousAttackedID(Sid) = lastAttackedID(Sid)
                CavebotTimeWithSameTarget(Sid) = mytime
            End If
            
        
        
            mobName = GetNameFromID(Sid, lastAttackedID(Sid))
            spellInfo = getSpellKill(Sid, mobName)
            If spellInfo.dist <> -1 Then
                If spellInfo.dist >= DistBetweenMeAndID(Sid, lastAttackedID(Sid)) Then
                    tmpHP = GetHPFromID(Sid, lastAttackedID(Sid))
                    If (tmpHP >= SpellKillHPlimit(Sid)) And (tmpHP <= SpellKillMaxHPlimit(Sid)) Then ' changed since 26.8
                        aRes = ExecuteInTibia(spellInfo.spell, Sid, True)
                        'ares = SendLogSystemMessageToClient(sid, "Attacking to <" & mobName & _
                         "> with spell <" & spellInfo.spell & "> if dist <=" & CStr(spellInfo.dist))
                        DoEvents
                    End If
                End If
            End If
        Else
            previousAttackedID(Sid) = 0
            CavebotTimeWithSameTarget(Sid) = mytime ' ------------------------
        End If
        
        If attackRes = 1 Then
            prevAttackState(Sid) = True
            Exit Sub
        Else
            prevAttackState(Sid) = False
        End If
    Else
        prevAttackState(Sid) = False
        lootTimeExpire(Sid) = 0
        requestLootBp(Sid) = &HFF
        lastAttackedID(Sid) = 0
    End If

    If (DictSETUSEITEM_used(Sid) = True) Then
        If CheckSETUSEITEM(Sid) = True Then
            Exit Sub
        End If
    End If

fastSet:
  currLineNumber = exeLine(Sid)
  ' FIXED!
  'Debug.Print "Cavebot ID selected = " & cavebotIDselected & " Currently executing: " & Sid
  'If (cavebotIDselected = Sid) Then ' Only display current line being executed if it is our selected char
   ' frmCavebot.lstScript.ListIndex = currLineNumber  ' ListIndex starts at 0, currLineNumber starts at 0
  'End If
  If DoingNewLoot(Sid) = True Then
    currLine = "move " & CStr(DoingNewLootX(Sid)) & "," & _
     CStr(DoingNewLootY(Sid)) & "," & _
     CStr(DoingNewLootZ(Sid))
  
  Else
    'work faster with local var
    currLine = GetStringFromIDLine(Sid, currLineNumber)
    currLine = parseVars(Sid, currLine)
  End If
  
 ' SendLogSystemMessageToClient sid, "executing " & CStr(currLineNumber) & " : " & currLine
  'DoEvents
  lenCurrLine = Len(currLine)
  pos = 1
  If ((currLine = "") Or (currLine = "?")) Then
    Exit Sub
  Else
    If (Left(currLine, 1) = "#") Or (Left(currLine, 1) = ":") Then
      'exeLine(Sid) = exeLine(Sid) + 1
      updateExeLine Sid, 1, True
      GoTo fastSet
    End If
  End If
  mainCommand = LCase(ParseString(currLine, pos, lenCurrLine, " "))
  'SendLogSystemMessageToClient sID, "Executing line " & CStr(currLineNumber) & " : " & currLine & " (detected '" & mainCommand & "')"
  'DoEvents
  SkipBlanks currLine, pos, lenCurrLine
  If waitCounter(Sid) < mytime Then ' don't execute command yet if we are waiting
  Select Case mainCommand
  Case "move"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    val1 = CLng(param1)
    SkipBlanks currLine, pos, lenCurrLine
    param2 = ParseString(currLine, pos, lenCurrLine, ",")
    val2 = CLng(param2)
    SkipBlanks currLine, pos, lenCurrLine
    param3 = ParseString(currLine, pos, lenCurrLine, ",")
    val3 = CLng(param3)
    
    
    continueP = True
    If DoingNewLoot(Sid) = False Then
            If AnythingLootable(Sid) = True Then
                continueP = False
            End If
    
    End If
    If continueP = True Then
    
        If (myX(Sid) > val1 - 2) And (myX(Sid) < val1 + 2) And _
           (myY(Sid) > val2 - 2) And (myY(Sid) < val2 + 2) And _
           (myZ(Sid) = val3) Then
          ' move completed
          If DoingNewLoot(Sid) = False Then
            'exeLine(Sid) = exeLine(Sid) + 1
            updateExeLine Sid, 1, True
          
          Else
            
            'aRes = GiveGMmessage(Sid, "Reached loot point", "Development")
            SmartLootCorpse Sid
            DoingNewLoot(Sid) = False
            DoEvents
          End If
        Else
          
          ' keep moving
          If ignoreNext(Sid) = -1 Then
            If DoingNewLoot(Sid) = True Then
                 DoingNewLoot(Sid) = False ' on trap, cancel order of moving to corpse
            End If
            ignoreNext(Sid) = 0
            If AllowRepositionAtStart(Sid) = 1 Then
                RepositionScriptAtTrap Sid, True 'initial reposition
            Else
                ignoreNext(Sid) = 0
            End If
          Else
            PerformMove Sid, val1, val2, val3
          End If
        End If
    End If

  Case "gotoscriptline"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    val1 = CLng(param1)
    'exeLine(Sid) = val1
    updateExeLine Sid, val1, False
   ' SendLogSystemMessageToClient sID, "Script jumped to line " & val1
    DoEvents
  Case "fishx"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    val1 = Round(CLng(param1) * (frmCavebot.TimerScript.Interval \ 100)) 'number of cast depends on timer speed
    fishCounter(Sid) = fishCounter(Sid) + 1
    If fishCounter(Sid) >= val1 Then
      ' fishing completed after this cast
      'exeLine(Sid) = exeLine(Sid) + 1
      updateExeLine Sid, 1, True
      ' reset counter
      fishCounter(Sid) = 0
    End If
    ' keep fishing
    aRes = CatchFish(Sid)
    DoEvents
  Case "waitx"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    val1 = CLng(param1)
    waitCounter(Sid) = GetTickCount() + (val1 * 1000)
   ' exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
  Case "closeconnection"
    frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "#Client " & Sid & " (" & CharacterName(Sid) & ")closed by script#"
    GiveServerError "Client closed because script executed closeConnection", Sid
    DoEvents
    frmMain.DoCloseActions Sid
  Case "stackitems"
    aRes = DoOneStack(Sid)
    If aRes = 0 Then
      ' stacking process completed
     ' exeLine(Sid) = exeLine(Sid) + 1
      updateExeLine Sid, 1, True
    End If
  Case "setpriority"
    usingPriorities(Sid) = True
    param1 = ParseString(currLine, pos, lenCurrLine, ":")
    SkipBlanks currLine, pos, lenCurrLine
    param2 = Trim$(Right$(currLine, Len(currLine) - pos))
    If publicDebugMode = True Then
        aRes = SendLogSystemMessageToClient(Sid, "Changed priority for " & LCase(param1) & " to " & CStr(safeLong(param2)))
        DoEvents
    End If
    AddKillPriority Sid, LCase(param1), safeLong(param2)
   ' exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "setmeleekill"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    AddMelee Sid, LCase(param1)
    'exeLine(Sid) = exeLine(Sid) + 1
   updateExeLine Sid, 1, True
    fastM = True
  Case "setmaxattacktimems"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    maxAttackTime(Sid) = safeLong(param1)
    ChaotizeNextMaxAttackTime Sid
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
    If EnableMaxAttackTime(Sid) = False Then
        aRes = SendLogSystemMessageToClient(Sid, "WARNING: setmaxattacktimems will have no effect unless you use SetBot EnableMaxAttackTime=1")
        DoEvents
    End If
  Case "setmaxhit"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    maxHit(Sid) = safeLong(param1)
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "setspellkill"
    param1 = Trim$(ParseString(currLine, pos, lenCurrLine, ","))
    SkipBlanks currLine, pos, lenCurrLine
    param2 = Trim$(ParseString(currLine, pos, lenCurrLine, ","))
    SkipBlanks currLine, pos, lenCurrLine
    param3 = Trim$(ParseString(currLine, pos, lenCurrLine, ","))
    val3 = CLng(param3)
    AddMelee Sid, LCase(param1)
    AddSpellKill Sid, LCase(param1), param2, val3
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "setexorivis"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    AddMelee Sid, LCase(param1)
    AddExorivis Sid, LCase(param1)
    AddExoriType Sid, LCase(param1), 1
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    CavebotHaveSpecials(Sid) = True
    fastM = True
  Case "setexorimort"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    AddMelee Sid, LCase(param1)
    AddExorivis Sid, LCase(param1)
    AddExoriType Sid, LCase(param1), 2
   ' exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    CavebotHaveSpecials(Sid) = True
    fastM = True
  Case "setavoidfront"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    AddAvoid Sid, LCase(param1)
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    CavebotHaveSpecials(Sid) = True
    fastM = True
  Case "sethmmkill"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    AddHMM Sid, LCase(param1)
    AddShotType Sid, LCase(param1), tileID_HMM
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "setsdkill"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    AddHMM Sid, LCase(param1)
    AddShotType Sid, LCase(param1), tileID_SD
   ' exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "resetkill"
    RemoveAllMelee Sid
    RemoveAllHMM Sid
    RemoveAllExorivis Sid
    RemoveAllAvoid Sid
    RemoveAllShotType Sid
    RemoveAllExoriType Sid
    CavebotHaveSpecials(Sid) = False
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "setbot"
    param1 = ParseString(currLine, pos, lenCurrLine, "=")
    SkipBlanks currLine, pos, lenCurrLine
    pos = pos + 1
    param2 = ParseString(currLine, pos, lenCurrLine, ",")
    val2 = CLng(param2)
     ' completed

        Select Case UCase(param1)
        Case "OLDLOOTMODE"
            If param2 = 0 Then
                OldLootMode(Sid) = False
            Else
                OldLootMode(Sid) = True
            End If
        Case "LOOTALL"
            If param2 = 0 Then
                LootAll(Sid) = False
            Else
                LootAll(Sid) = True
            End If
        Case "PKWARNINGS"
            If param2 = 0 Then
                PKwarnings(Sid) = False
            Else
                PKwarnings(Sid) = True
            End If
        Case "ENABLEMAXATTACKTIME"
            If param2 = 0 Then
                EnableMaxAttackTime(Sid) = False
            Else
                EnableMaxAttackTime(Sid) = True
            End If
        Case "MINDELAYTOLOOT"
            MINDELAYTOLOOT(Sid) = val2
        Case "MAXTIMEINLOOTQUEUE"
            MAXTIMEINLOOTQUEUE(Sid) = val2
        Case "MAXTIMETOREACHCORPSE"
            MAXTIMETOREACHCORPSE(Sid) = val2
        Case "AUTOEATFOOD"
            If val2 = 0 Then
                RuneMakerOptions(Sid).autoEat = False
            Else
                RuneMakerOptions(Sid).autoEat = True
            End If
        Case "SPELLKILLHPLIMIT"
           SpellKillHPlimit(Sid) = val2
        Case "SPELLKILLMAXHPLIMIT"
           SpellKillMaxHPlimit(Sid) = val2
        Case "ALLOWREPOSITIONATSTART"
            AllowRepositionAtStart(Sid) = val2
        Case "ALLOWREPOSITIONATTRAP"
            AllowRepositionAtTrap(Sid) = val2
        End Select
      
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "useitem"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    val1 = CLng(param1)
    SkipBlanks currLine, pos, lenCurrLine
    param2 = ParseString(currLine, pos, lenCurrLine, ",")
    val2 = CLng(param2)
    SkipBlanks currLine, pos, lenCurrLine
    param3 = ParseString(currLine, pos, lenCurrLine, ",")
    val3 = CLng(param3)
    PerformUseItem Sid, val1, val2, val3
     ' completed
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
  Case "iftrue"
    aRes = ProcessCondition(Sid, currLine, pos, lenCurrLine)
  Case "ifenoughitemsgoto"
    param1 = Right$(currLine, Len(currLine) - 18)
    stringParts = Split(param1, ",")
    If UBound(stringParts) = 2 Then
    
      param1 = Trim$(stringParts(0))
      param2 = Trim$(stringParts(1))
      param3 = Trim$(stringParts(2))
      val2 = CLng(param2)
      val3 = CLng(param3)
    'param1 = ParseString(currLine, pos, lenCurrLine, ",")

    'SkipBlanks currLine, pos, lenCurrLine
    'param2 = ParseString(currLine, pos, lenCurrLine, ",")
    ''val2 = CLng(param2)
    'SkipBlanks currLine, pos, lenCurrLine
    'param3 = ParseString(currLine, pos, lenCurrLine, ",")
    'val3 = CLng(param3)
  
    am = CountTheItemsForUser(Sid, param1) ' changed since 9.38
    If am >= val2 Then
     ' exeLine(Sid) = val3
      updateExeLine Sid, val3, False
      If publicDebugMode = True Then
        aRes = SendLogSystemMessageToClient(Sid, "Condition (" & am & " number>= " & val2 & ") = TRUE")
        DoEvents
      End If
    Else
     ' exeLine(Sid) = exeLine(Sid) + 1
      updateExeLine Sid, 1, True
      If publicDebugMode = True Then
        aRes = SendLogSystemMessageToClient(Sid, "Condition  (" & am & " number>= " & val2 & ") = FALSE")
        DoEvents
      End If
    End If
    
    End If
  Case "iffewitemsgoto"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    'val1 = GetTheLongFromFiveChr(param1)  not needed since in 9.38
    SkipBlanks currLine, pos, lenCurrLine
    param2 = ParseString(currLine, pos, lenCurrLine, ",")
    val2 = CLng(param2) ' value to be compared
    SkipBlanks currLine, pos, lenCurrLine
    param3 = ParseString(currLine, pos, lenCurrLine, ",")
    val3 = CLng(param3) ' line where it should jump
    'ammount of items with given tileID
    am = CountTheItemsForUser(Sid, param1) ' changed since 9.38
    ' compare now
    If am >= val2 Then ' false : continue with next line of the script
      'exeLine(Sid) = exeLine(Sid) + 1
      updateExeLine Sid, 1, True
      If publicDebugMode = True Then
        aRes = SendLogSystemMessageToClient(Sid, "Condition (" & am & " number< " & val2 & ") = FALSE")
        DoEvents
      End If
    Else ' true : jump to given line
     ' exeLine(Sid) = val3
      updateExeLine Sid, val3, False
      If publicDebugMode = True Then
        aRes = SendLogSystemMessageToClient(Sid, "Condition  (" & am & " number< " & val2 & ") = TRUE")
        DoEvents
      End If
    End If
  Case "setretryattacks"
    AvoidReAttacks(Sid) = False
   ' exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
  Case "setdontretryattacks"
    AvoidReAttacks(Sid) = True
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
  Case "saymessage"
    param1 = Trim$(currLine)
    param2 = Right$(param1, Len(param1) - 11)
    aRes = ExecuteInTibia(param2, Sid, True)
    DoEvents
    'CastSpell sid, param1
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
  Case "sayintrade"
    param1 = Trim$(currLine)
    param2 = Right$(param1, Len(param1) - 11)
    aRes = ExecuteInTibia("exiva sayt:" & param2, Sid, True)
    DoEvents
    'CastSpell sid, param1
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
  Case "fastexiva"
    param1 = Trim$(currLine)
    param2 = Right$(param1, Len(param1) - 10)
    aRes = ExecuteInTibia("exiva " & param2, Sid, True)
    DoEvents
    'CastSpell sid, param1
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    ' instantly jump to next
    GoTo fastSet
  Case "ondangergoto"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    val1 = CLng(param1)
    cavebotOnDanger(Sid) = val1
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "ongmcloseconnection"
    If ((TibiaVersionLong < 811) Or (Antibanmode = 0)) Then
        cavebotOnGMclose(Sid) = True
    Else ' else it is ignored
        aRes = SendLogSystemMessageToClient(Sid, "WARNING: ongmcloseconnection is being ignored since Tibia 8.11 (you would get banished other way) Please delete that line from your script")
        DoEvents
    End If
   ' exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "ontrapgivealarm"
    cavebotOnTrapGiveAlarm(Sid) = True
   ' exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "ongmpause"
    cavebotOnGMpause(Sid) = True
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "onplayerpause-"
    cavebotOnPLAYERpause(Sid) = True
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "setnofollow"
    setFollowTarget(Sid) = False
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "setfollow"
    setFollowTarget(Sid) = True
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "setlooton"
    autoLoot(Sid) = True
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "setlootoff"
    autoLoot(Sid) = False
   ' exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "setloot"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    AddGoodLoot Sid, GetTheLongFromFiveChr(param1)
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "setnoloot"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    RemoveGoodLoot Sid, GetTheLongFromFiveChr(param1)
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "setuseitem"
    param1 = Trim$(ParseString(currLine, pos, lenCurrLine, ":"))
    
    pos = pos + 1
    param2 = Trim$(ParseString(currLine, pos, lenCurrLine, ","))
    
    AddSETUSEITEM Sid, param2, param1
    SendLogSystemMessageToClient Sid, "Cavebot will now use item '" & param1 & "' on near items with id '" & param2 & "'"
    
   ' exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "setlootdistance"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    val1 = CLng(param1)
    AllowedLootDistance(Sid) = val1
   ' exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "resetloot"
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    RemoveAllGoodLoot Sid
    fastM = True
  Case "setchaoticmovesoff"
   ' exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    'CavebotChaoticMode(Sid) = 0
    SendLogSystemMessageToClient Sid, "Warning: chaoticmoves setting is now ignored."
    fastM = True
  Case "setchaoticmoveson"
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    'CavebotChaoticMode(Sid) = 1
    SendLogSystemMessageToClient Sid, "Warning: chaoticmoves setting is now ignored."
    fastM = True
  Case "setany"
    friendlyMode(Sid) = 0
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "setfriendly"
    friendlyMode(Sid) = 1
    'SendLogSystemMessageToClient sid, "Friendly mode activated"
    'DoEvents
   ' exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "setveryfriendly"
    friendlyMode(Sid) = 2
    'SendLogSystemMessageToClient sid, "Friendly mode activated"
    'DoEvents
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    fastM = True
  Case "putlootondepot"
    If onDepotPhase(Sid) = 0 Then
      ' first time we read this
      ' stop attacks
      onDepotPhase(Sid) = 1
    ElseIf onDepotPhase(Sid) = 1 Then
      ' second time, find path
      moveRetry(Sid) = 0
      aRes = GetNearestDepot(Sid)
      If aRes = 0 Then
        onDepotPhase(Sid) = 2 ' got depotx,depoty,depotz
      Else
        onDepotPhase(Sid) = 0 'end
        ' could not find depot
        'exeLine(Sid) = exeLine(Sid) + 1
        updateExeLine Sid, 1, True
      End If
    ElseIf onDepotPhase(Sid) = 2 Then
    If (myX(Sid) > (depotX(Sid) - 2)) And (myX(Sid) < (depotX(Sid) + 2)) And _
       (myY(Sid) > (depotY(Sid) - 2)) And (myY(Sid) < (depotY(Sid) + 2)) And _
       (myZ(Sid) = depotZ(Sid)) Then
      ' move completed
      onDepotPhase(Sid) = 3
      lastDepotBPID(Sid) = &HFF
    Else
      ' keep moving
      PerformMove Sid, depotX(Sid), depotY(Sid), depotZ(Sid)
    End If
    ElseIf onDepotPhase(Sid) = 3 Then
      ' open depot
      If lastDepotBPID(Sid) = &HFF Then
        OpenTheDepot Sid
        lastDepotBPID(Sid) = &HFE
      ElseIf lastDepotBPID(Sid) <> &HFE Then
        onDepotPhase(Sid) = 4
        doneDepotChestOpen(Sid) = False
        OpenDepotChest Sid
      End If
      
    ElseIf onDepotPhase(Sid) = 4 Then
      If doneDepotChestOpen(Sid) = True Then
        onDepotPhase(Sid) = 5
        somethingChangedInBps(Sid) = True
        nextForcedDepotDeployRetry(Sid) = GetTickCount() + 5000
      Else
        If timeToRetryOpenDepot(Sid) > GetTickCount() Then
            OpenDepotChest Sid
        End If
      End If
    ElseIf onDepotPhase(Sid) = 5 Then
      aRes = DropLoot(Sid)
      If aRes = -1 Then
        onDepotPhase(Sid) = 0 'end of depot deploy command
        'exeLine(Sid) = exeLine(Sid) + 1
        updateExeLine Sid, 1, True
      End If
      
    End If
  Case "droplootonground"
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    val1 = CLng(param1)
    SkipBlanks currLine, pos, lenCurrLine
    param2 = ParseString(currLine, pos, lenCurrLine, ",")
    val2 = CLng(param2)
    SkipBlanks currLine, pos, lenCurrLine
    param3 = ParseString(currLine, pos, lenCurrLine, ",")
    val3 = CLng(param3)
    If onDepotPhase(Sid) < 6 Then
      onDepotPhase(Sid) = 6 'dropping loot
    End If
    If (myX(Sid) = val1) And _
       (myY(Sid) = val2) And _
       (myZ(Sid) = val3) Then
      ' move completed
      ' now drop loot
      If onDepotPhase(Sid) = 6 Then
        nextForcedDepotDeployRetry(Sid) = GetTickCount()
        somethingChangedInBps(Sid) = True
        onDepotPhase(Sid) = 7
      Else
        aRes = DropLootToGround(Sid)
        If aRes = -1 Then
          onDepotPhase(Sid) = 0 'end of ground deploy command
         ' exeLine(Sid) = exeLine(Sid) + 1
          updateExeLine Sid, 1, True
        End If
      End If
    Else
      ' keep moving
      PerformMove Sid, val1, val2, val3
    End If
  Case Else
    'exeLine(Sid) = exeLine(Sid) + 1
    updateExeLine Sid, 1, True
    SendLogSystemMessageToClient Sid, "Unknown command at line " & currLineNumber & " : " & mainCommand
    DoEvents
  End Select
  End If
  If fastM = True Then
    fastM = False
    GoTo fastSet
  End If
  Exit Sub
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & Sid & " lost connection at ProcessScriptLine #"
  frmMain.DoCloseActions Sid
  DoEvents
End Sub

Public Sub FixRightNumberOfClicks(idConnection As Integer, ByVal clicks As Long)
    If ProcessID(idConnection) = -1 Then
        Exit Sub
    End If
    Memory_WriteLong adrNumberOfAttackClicks, clicks, ProcessID(idConnection)
End Sub
Public Function RightNumberOfClicksJUSTREAD(idConnection As Integer) As String
    Dim clicks As Long
    Dim fourBytesCRC(3) As Byte
    Dim res As String
    If TibiaVersionLong < 860 Then
        RightNumberOfClicksJUSTREAD = ""
        Exit Function
    End If
    If ProcessID(idConnection) = -1 Then
        RightNumberOfClicksJUSTREAD = ""
        Exit Function
    End If
    clicks = Memory_ReadLong(adrNumberOfAttackClicks, ProcessID(idConnection))
    longToBytes fourBytesCRC, clicks
    res = "ReadLong = " & GoodHex(fourBytesCRC(0)) & " " & GoodHex(fourBytesCRC(1)) & " " & _
    GoodHex(fourBytesCRC(2)) & " " & GoodHex(fourBytesCRC(3))
    
    fourBytesCRC(0) = Memory_ReadByte(adrNumberOfAttackClicks, ProcessID(idConnection))
    fourBytesCRC(1) = Memory_ReadByte(adrNumberOfAttackClicks + 1, ProcessID(idConnection))
    fourBytesCRC(2) = Memory_ReadByte(adrNumberOfAttackClicks + 2, ProcessID(idConnection))
    fourBytesCRC(3) = Memory_ReadByte(adrNumberOfAttackClicks + 3, ProcessID(idConnection))
    
    res = res & " : ReadByte = " & GoodHex(fourBytesCRC(0)) & " " & GoodHex(fourBytesCRC(1)) & " " & _
    GoodHex(fourBytesCRC(2)) & " " & GoodHex(fourBytesCRC(3))
    RightNumberOfClicksJUSTREAD = res
End Function

Public Function RightNumberOfClicks(idConnection As Integer) As String
    Dim clicks As Long
    Dim fourBytesCRC(3) As Byte
    Dim res As String
    If TibiaVersionLong < 860 Then
        RightNumberOfClicks = ""
        Exit Function
    End If
    If ProcessID(idConnection) = -1 Then
        RightNumberOfClicks = ""
        Exit Function
    End If
    clicks = Memory_ReadLong(adrNumberOfAttackClicks, ProcessID(idConnection))
    'Debug.Print "M=" & CStr(clicks)
    clicks = clicks + 1
    longToBytes fourBytesCRC, clicks
    res = " " & GoodHex(fourBytesCRC(0)) & " " & GoodHex(fourBytesCRC(1)) & " " & _
    GoodHex(fourBytesCRC(2)) & " " & GoodHex(fourBytesCRC(3))
    
    Memory_WriteLong adrNumberOfAttackClicks, clicks, ProcessID(idConnection)
    RightNumberOfClicks = res
End Function
Public Sub WriteRedSquare(ByVal idConnection As Integer, ByVal targetID As Long)
    On Error GoTo goterr
    If RedSquare <> &H0 Then
        Memory_WriteLong RedSquare, targetID, ProcessID(idConnection)
    End If
    Exit Sub
goterr:
    Exit Sub
End Sub
Public Function ReadRedSquare(ByVal idConnection As Integer) As Long
    On Error GoTo goterr
    Dim clicks As Long
    If RedSquare = &H0 Then
        ReadRedSquare = -2
    ElseIf ProcessID(idConnection) = -1 Then
        ReadRedSquare = -1
    Else
        clicks = Memory_ReadLong(RedSquare, ProcessID(idConnection))
        ReadRedSquare = clicks
    End If
    Exit Function
goterr:
    ReadRedSquare = -1
    Exit Function
End Function
Public Function MeleeAttack(idConnection As Integer, targetID As Double, Optional forceSend As Boolean = False) As Long
  ' attack ID in melee
  Dim aRes As Long
  Dim sCheat As String
  Dim cPacket() As Byte
  Dim inRes As Integer
  Dim safeRes As String
  Dim currentRedSquare As Long
  Dim tempID As Long
  Dim templ1 As Long
  
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  If (GameConnected(idConnection) = True) And (sentFirstPacket(idConnection) = True) Then
    '05 00 A1 7B 8A 02 40
    'DEBUGG
    'aRes = SendLogSystemMessageToClient(idconnection, "DEBUG > Attacking " & CStr(targetID))
    'DoEvents
    If isIgnoredcreature(idConnection, targetID) = True Then ' stop!
        currTargetID(idConnection) = 0
    End If
    currTargetID(idConnection) = targetID

    currentRedSquare = ReadRedSquare(idConnection) ' will always return -2 in old versions
    If currentRedSquare = 0 Then
        TurnsWithRedSquareZero(idConnection) = TurnsWithRedSquareZero(idConnection) + 1
    Else
        TurnsWithRedSquareZero(idConnection) = 0
    End If
    If (currentRedSquare <> 0) Or (TurnsWithRedSquareZero(idConnection) >= 3) Then
        If currentRedSquare <> -1 Then
            If (currTargetID(idConnection)) <> 0 Or (forceSend = True) Then
                If ((currentRedSquare <> currTargetID(idConnection)) Or (forceSend = 0)) Then ' new since Blackd Proxy 24.0
                
                
                    If TrainerOptions(idConnection).misc_stoplowhp = 1 Then
                        If targetID <> 0 Then
                            If tempID = currTargetID(idConnection) Then
                                If (templ1 < TrainerOptions(idConnection).stoplowhpHP) Then
                                    ' should not attack this target
                                    MeleeAttack = -1
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                    If TibiaVersionLong >= 860 Then
                        safeRes = RightNumberOfClicks(idConnection)
                        If safeRes = "" Then
                                ' error happened, unsafe to attack
                                MeleeAttack = -1
                                Exit Function
                        End If
                        sCheat = "09 00 A1 " & SpaceID(targetID) & safeRes
                    Else
                
                        sCheat = "05 00 A1 " & SpaceID(targetID)
                    End If
                
                
                    WriteRedSquare idConnection, currTargetID(idConnection) ' new since in tibia 8.62
                    inRes = GetCheatPacket(cPacket, sCheat)
                    frmMain.UnifiedSendToServerGame idConnection, cPacket, True
                    DoEvents
                End If
            End If
        End If
    End If
  End If
  MeleeAttack = 0
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Got error at MeleeAttack : " & Err.Description
  frmMain.DoCloseActions idConnection
  MeleeAttack = -1
End Function

Public Function CavebotRuneAttack(idConnection As Integer, targetID As Double, runeB1 As Byte, runeB2 As Byte) As Long
  Dim aRes As Long
  Dim lTarget As String
  Dim lSquare As String
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim thing As String
  Dim fRes As TypeSearchItemResult2
  Dim myS As Byte
  Dim X As Long
  Dim y As Long
  Dim s As Long
  Dim tileID As Long
  Dim tmpID As Double
  Dim inRes As Integer
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  tileID = GetTheLong(runeB1, runeB2)
  Select Case tileID
  Case tileID_SD
    thing = "SDs"
  Case tileID_HMM
    thing = "HMMs"
  Case tileID_Explosion
    thing = "Explosions"
  Case tileID_IH
    thing = "IHs"
  Case tileID_UH
    thing = "UHs"
  Case tileID_fireball
    thing = "Fireballs"
  Case tileID_stalagmite
    thing = "Stalagmites"
  Case tileID_icicle
    thing = "Icicles"
  Case Else
    thing = "runes"
  End Select
  
 ' aRes = SendMessageToClient(idConnection, "Casting on " & target & " ;)", "GM BlackdProxy")
  ' search the rune
  fRes = SearchItem(idConnection, runeB1, runeB2)  'search thing
  If fRes.foundcount = 0 Then
   ' !!! BUG
     aRes = SendSystemMessageToClient(idConnection, "can't find " & thing & ", open new bp of " & thing & "!")
     CavebotRuneAttack = 0
     Exit Function
  End If
  If (TibiaVersionLong < 760) Then
    myS = MyStackPos(idConnection)
  Else
    myS = FirstPersonStackPos(idConnection)
  End If
  ' search yourself
  If myS = &HFF Then
    aRes = SendSystemMessageToClient(idConnection, "Your map is out of sync, can't use " & thing & "!")
    CavebotRuneAttack = 0
    Exit Function
  End If
  
  If targetID = 0 Then ' hmm to nothing??
    'should not happen ...
    CavebotRuneAttack = -1
    Exit Function
  End If
  ' just cast to id
   sCheat = "84 FF FF " & GoodHex(&H40 + fRes.bpID) & " 00 " & _
               GoodHex(fRes.slotID) & " " & GoodHex(runeB1) & " " & GoodHex(runeB2) & " " & _
               GoodHex(fRes.slotID) & " " & SpaceID(targetID)
   SafeCastCheatString "CavebotRuneAttack", idConnection, sCheat
'   inRes = GetCheatPacket(cPacket, sCheat)
'   frmMain.UnifiedSendToServerGame idConnection, cPacket, True
'   DoEvents
   CavebotRuneAttack = 0
   Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at CavebotRuneAttack #"
  frmMain.DoCloseActions idConnection
  DoEvents
  CavebotRuneAttack = -1
End Function
Public Sub PerformUseItem(idConnection As Integer, X As Long, y As Long, z As Long)
  '0A 00 82 3C 7D AD 7D 08 89 07 01 00
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim xdif As Long
  Dim ydif As Long
  Dim b1 As Byte
  Dim b2 As Byte
  Dim SOPT As Byte
  Dim SS As Byte
  Dim inRes As Integer
  Dim tileID As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  xdif = X - myX(idConnection)
  ydif = y - myY(idConnection)
  If xdif < -7 Or xdif > 8 Or ydif < -5 Or ydif > 6 Then
    'out of range: first move there
    PerformMove idConnection, X, y, myZ(idConnection)
    Exit Sub
  End If
  SOPT = 1
  For SS = 1 To 10
    tileID = GetTheLong(Matrix(ydif, xdif, z, idConnection).s(SS).t1, Matrix(ydif, xdif, z, idConnection).s(SS).t2)
    If DatTiles(tileID).alwaysOnTop = True Then
      SOPT = SS
    Else
      Exit For
    End If
  Next SS
  b1 = Matrix(ydif, xdif, z, idConnection).s(SOPT).t1
  b2 = Matrix(ydif, xdif, z, idConnection).s(SOPT).t2
  sCheat = "0A 00 82 " & FiveChrLon(X) & " " & FiveChrLon(y) & " " & GoodHex(CByte(z)) & _
   " " & GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(SS) & " 00"
   ' debug
 'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & sCheat
  inRes = GetCheatPacket(cPacket, sCheat)
  waitCounter(idConnection) = GetTickCount() + 2000
  frmMain.UnifiedSendToServerGame idConnection, cPacket, True
  DoEvents
  Exit Sub
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Got error at PerformUseItem"
End Sub
Public Sub WriteNoSafeToAttack(ByRef idMap As TypeIDmap, X As Long, y As Long)
  If (X > -10) And (X < 11) And (y > -8) And (y < 9) Then
   idMap.isSafe(X, y) = False
  End If
End Sub

Public Function DoSpecialCavebot(idConnection As Integer, xt As Long, yt As Long, zt As Long, dblID As Double) As Long
    #If FinalMode Then
      On Error GoTo goterr
    #End If
    Dim myMap As TypeAstarMatrix
    Dim idMap As TypeIDmap
    Dim bestmove As Byte
    Dim cost1 As Long
    Dim cost2 As Long
    Dim X As Long
    Dim y As Long
    Dim doingAvoid As Boolean
    Dim doingExori As Boolean
    Dim chooseFirst As Boolean
    Dim nameofgivenID As String
    Dim opt1 As Byte
    Dim opt2 As Byte
    Dim res As Long
    Dim mydir As Byte
    Dim aRes As Long
    Dim spellToUse As String
    Dim costOfSpellToUse As Long
    opt1 = 0
    opt2 = 0
    nameofgivenID = GetNameFromID(idConnection, dblID)
    If getExoriType(idConnection, LCase(nameofgivenID)) = 2 Then
        spellToUse = EXORIMORT_SPELL
        costOfSpellToUse = EXORIMORT_COST
    Else
        spellToUse = EXORIVIS_SPELL
        costOfSpellToUse = EXORIVIS_COST
    End If
    doingAvoid = False
    doingExori = False
    bestmove = 0
    If isAvoid(idConnection, nameofgivenID) = True Then
        doingAvoid = True
    ElseIf isExorivis(idConnection, nameofgivenID) = True Then
        doingExori = True
    Else
        DoSpecialCavebot = 0
        Exit Function
    End If
    X = 0
    y = 0
    LoadCurrentFloorIntoMatrix idConnection, myMap, idMap, True, True
    If randomNumberBetween(0, 1) = 1 Then
        chooseFirst = True
    Else
        chooseFirst = False
    End If
    If doingAvoid = True Then
        If CavebotLastSpecialMove(idConnection) < GetTickCount() Then
            If ((xt = X + 1) Or (xt = X - 1) Or (xt = X)) And yt = y Then  '  move north or south
                cost1 = myMap.cost(X, y - 1)
                cost2 = myMap.cost(X, y + 1)
                If (cost1 < CostBlock) And _
                   ((cost1 < cost2) Or ((cost1 = cost2))) Then
                    opt1 = &H65 ' north
                End If
                If (cost2 < CostBlock) And _
                   ((cost2 < cost1) Or ((cost2 = cost1))) Then
                    opt2 = &H67 ' south
                End If
            ElseIf ((yt = y + 1) Or (yt = y - 1) Or (yt = y)) And xt = X Then '  move left or right
                cost1 = myMap.cost(X - 1, y)
                cost2 = myMap.cost(X + 1, y)
                If (cost1 < CostBlock) And _
                   ((cost1 < cost2) Or ((cost1 = cost2))) Then
                    opt1 = &H68 ' left
                End If
                If (cost2 < CostBlock) And _
                   ((cost2 < cost1) Or ((cost2 = cost1))) Then
                    opt2 = &H66 ' right
                End If
            End If
            If ((opt1 > 0) And (opt2 > 0)) Then
                If chooseFirst = True Then
                    DoManualMove idConnection, opt1
                    DoEvents
                    CavebotLastSpecialMove(idConnection) = GetTickCount() + cte_RepositionDelay
                    DoSpecialCavebot = 0
                Else
                    DoManualMove idConnection, opt2
                    DoEvents
                    CavebotLastSpecialMove(idConnection) = GetTickCount() + cte_RepositionDelay
                    DoSpecialCavebot = 0
                End If
            ElseIf opt1 > 0 Then
                DoManualMove idConnection, opt1
                DoEvents
                CavebotLastSpecialMove(idConnection) = GetTickCount() + cte_RepositionDelay
                DoSpecialCavebot = 0
            ElseIf opt2 > 0 Then
                DoManualMove idConnection, opt2
                DoEvents
                CavebotLastSpecialMove(idConnection) = GetTickCount() + cte_RepositionDelay
                DoSpecialCavebot = 0
            End If
            Exit Function
        End If
    Else ' doing exorivis
        res = 1
        If CavebotLastSpecialMove(idConnection) < GetTickCount() Then
            If ((xt = X + 1) And (yt = y - 1)) Then '  move north or right
                cost1 = myMap.cost(X, y - 1)
                cost2 = myMap.cost(X + 1, y)
                If (cost1 < CostBlock) And _
                   ((cost1 < cost2) Or ((cost1 = cost2))) Then
                    opt1 = &H65 ' north
                End If
                If (cost2 < CostBlock) And _
                   ((cost2 < cost1) Or ((cost2 = cost1))) Then
                    opt2 = &H66 ' right
                End If
            ElseIf ((xt = X + 1) And (yt = y + 1)) Then '  move south or right
                cost1 = myMap.cost(X, y + 1)
                cost2 = myMap.cost(X + 1, y)
                If (cost1 < CostBlock) And _
                   ((cost1 < cost2) Or ((cost1 = cost2))) Then
                    opt1 = &H67 ' south
                End If
                If (cost2 < CostBlock) And _
                   ((cost2 < cost1) Or ((cost2 = cost1))) Then
                    opt2 = &H66 ' right
                End If
            ElseIf ((xt = X - 1) And (yt = y + 1)) Then '  move south or left
                cost1 = myMap.cost(X, y + 1)
                cost2 = myMap.cost(X - 1, y)
                If (cost1 < CostBlock) And _
                   ((cost1 < cost2) Or ((cost1 = cost2))) Then
                    opt1 = &H67 ' south
                End If
                If (cost2 < CostBlock) And _
                   ((cost2 < cost1) Or ((cost2 = cost1))) Then
                    opt2 = &H68  ' left
                End If
            ElseIf ((xt = X - 1) And (yt = y - 1)) Or ((xt = X) And (yt = y)) Then   '  move north or left
                cost1 = myMap.cost(X, y - 1)
                cost2 = myMap.cost(X - 1, y)
                If (cost1 < CostBlock) And _
                   ((cost1 < cost2) Or ((cost1 = cost2))) Then
                    opt1 = &H65 ' north
                End If
                If (cost2 < CostBlock) And _
                   ((cost2 < cost1) Or ((cost2 = cost1))) Then
                    opt2 = &H68  ' left
                End If
            End If
            If ((opt1 > 0) And (opt2 > 0)) Then
                If chooseFirst = True Then
                    DoManualMove idConnection, opt1
                    DoEvents
                    CavebotLastSpecialMove(idConnection) = GetTickCount() + cte_RepositionDelay
                    res = 0
                Else
                    DoManualMove idConnection, opt2
                    DoEvents
                    CavebotLastSpecialMove(idConnection) = GetTickCount() + cte_RepositionDelay
                    res = 0
                End If
            ElseIf opt1 > 0 Then
                DoManualMove idConnection, opt1
                DoEvents
                CavebotLastSpecialMove(idConnection) = GetTickCount() + cte_RepositionDelay
                res = 0
            ElseIf opt2 > 0 Then
                DoManualMove idConnection, opt2
                DoEvents
                CavebotLastSpecialMove(idConnection) = GetTickCount() + cte_RepositionDelay
                res = 0
            End If

        End If
        ' 00 = north ; 01 = right ; 02 = south ; 03 = left
        mydir = GetDirectionFromID(idConnection, myID(idConnection))
        If (xt = X + 1) And (yt = y) Then
           If mydir <> 1 Then
                res = 0
                TurnMe idConnection, 1
           End If
        End If
        If (xt = X - 1) And (yt = y) Then
           If mydir <> 3 Then
                res = 0
                TurnMe idConnection, 3
           End If
        End If
        If (xt = X) And (yt = y + 1) Then
           If mydir <> 2 Then
                res = 0
                TurnMe idConnection, 2
           End If
        End If
        If (xt = X) And (yt = y - 1) Then
           If mydir <> 0 Then
                res = 0
                TurnMe idConnection, 0
           End If
        End If
        If (xt = X + 1) And (yt = y) And (mydir = 1) Then
            If Round((myHP(idConnection) / myMaxHP(idConnection)) * 100) >= frmCavebot.scrollExorivis.Value Then
                If myMana(idConnection) >= costOfSpellToUse Then
                    res = 0
                    aRes = CastSpell(idConnection, spellToUse)
                    DoEvents
                Else
                    If PlayTheDangerSound = False Then
                        aRes = GiveGMmessage(idConnection, "Low of mana: can't use " & spellToUse, "Blackd Proxy")
                        DoEvents
                        PlayTheDangerSound = True
                    End If
                End If
            End If
        End If
        If (xt = X - 1) And (yt = y) And (mydir = 3) Then
            If Round((myHP(idConnection) / myMaxHP(idConnection)) * 100) >= frmCavebot.scrollExorivis.Value Then
                If myMana(idConnection) >= costOfSpellToUse Then
                    res = 0
                    aRes = CastSpell(idConnection, spellToUse)
                    DoEvents
                Else
                    If PlayTheDangerSound = False Then
                        aRes = GiveGMmessage(idConnection, "Low of mana: can't use " & spellToUse, "Blackd Proxy")
                        DoEvents
                        PlayTheDangerSound = True
                    End If
                End If
            End If
        End If
        If (xt = X) And (yt = y + 1) And (mydir = 2) Then
            If Round((myHP(idConnection) / myMaxHP(idConnection)) * 100) >= frmCavebot.scrollExorivis.Value Then
                If myMana(idConnection) >= costOfSpellToUse Then
                    res = 0
                    aRes = CastSpell(idConnection, spellToUse)
                    DoEvents
                Else
                    If PlayTheDangerSound = False Then
                        aRes = GiveGMmessage(idConnection, "Low of mana: can't use " & spellToUse, "Blackd Proxy")
                        DoEvents
                        PlayTheDangerSound = True
                    End If
                End If
            End If
        End If
        If (xt = X) And (yt = y - 1) And (mydir = 0) Then
            If Round((myHP(idConnection) / myMaxHP(idConnection)) * 100) >= frmCavebot.scrollExorivis.Value Then
                If myMana(idConnection) >= costOfSpellToUse Then
                    res = 0
                    aRes = CastSpell(idConnection, spellToUse)
                    DoEvents
                Else
                    If PlayTheDangerSound = False Then
                        aRes = GiveGMmessage(idConnection, "Low of mana: can't use " & spellToUse, "Blackd Proxy")
                        DoEvents
                        PlayTheDangerSound = True
                    End If
                End If
            End If
        End If
        DoSpecialCavebot = res
        Exit Function
    End If
    DoSpecialCavebot = 0
    Exit Function
goterr:
    DoSpecialCavebot = 0
End Function
Public Sub LoadCurrentFloorIntoMatrix(idConnection As Integer, ByRef myMap As TypeAstarMatrix, ByRef idMap As TypeIDmap, skipThisPart As Boolean, ByRef cond2 As Boolean)
  Dim nameofgivenID As String
  Dim X As Long
  Dim y As Long
  Dim z As Long
  Dim s As Long
  Dim continue As Boolean
  Dim tileID As Long
  Dim tmpID As Double
  Dim aRes As Long
  Dim tmp1 As Boolean
  Dim tmp2 As Boolean
  ' delimiter our map by a wall
  For X = -9 To 10
    myMap.cost(X, -7) = CostBlock
    myMap.cost(X, 8) = CostBlock
  Next X
  For y = -7 To 8
    myMap.cost(-9, y) = CostBlock
    myMap.cost(10, y) = CostBlock
  Next y
  If skipThisPart = False Then
  ' init id info
  For X = -9 To 10
    For y = -7 To 8
      idMap.dblID(X, y) = 0
      idMap.isSafe(X, y) = True
      idMap.isHmm(X, y) = False
      idMap.isMelee(X, y) = False
    Next y
  Next X
  End If
  ' extract all valuable info from truemap
  z = myZ(idConnection)
  For y = -6 To 7
    For X = -8 To 9
      continue = True
      tileID = GetTheLong(Matrix(y, X, z, idConnection).s(0).t1, Matrix(y, X, z, idConnection).s(0).t2)
      If tileID = 0 Then
        myMap.cost(X, y) = CostBlock
        continue = False
      ElseIf DatTiles(tileID).blocking = True Then
        myMap.cost(X, y) = CostBlock
        continue = False
      ElseIf (DatTiles(tileID).floorChangeDOWN = True) Or (DatTiles(tileID).floorChangeUP = True) Then
        
        If (DatTiles(tileID).requireShovel = False) And (DatTiles(tileID).requireRope = False) And (DatTiles(tileID).requireRightClick = False) Then
          myMap.cost(X, y) = CostBlock
               If (myMap.cost(X - 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y)) < CostNearHandicap Then
                 myMap.cost(X - 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y + 1)) < CostNearHandicap Then
                 myMap.cost(X, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y)) < CostNearHandicap Then
                 myMap.cost(X + 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y - 1)) < CostNearHandicap Then
                 myMap.cost(X, y - 1) = CostNearHandicap
               End If
          continue = False
        End If
      End If
      If continue = True Then
        If (myMap.cost(X, y)) < CostWalkable Then
          myMap.cost(X, y) = CostWalkable
        End If
        For s = 1 To 10
          tileID = GetTheLong(Matrix(y, X, z, idConnection).s(s).t1, Matrix(y, X, z, idConnection).s(s).t2)
          If tileID = 0 Then
            Exit For
          ElseIf tileID = 97 Then 'person
            tmpID = Matrix(y, X, z, idConnection).s(s).dblID
            nameofgivenID = GetNameFromID(idConnection, tmpID)
            If nameofgivenID = "" Then
              ' detected mobile with no name! ?
              aRes = -1
            ElseIf nameofgivenID = CharacterName(idConnection) Then
              ' myself
              aRes = -2
            Else
              myMap.cost(X, y) = CostBlock
              tmp1 = False
              tmp2 = False
              If isHmm(idConnection, nameofgivenID) = True Then
                tmp1 = True
                idMap.isHmm(X, y) = True
                idMap.dblID(X, y) = tmpID
              End If
              If isMelee(idConnection, nameofgivenID) = True Then
                tmp2 = True
                idMap.isMelee(X, y) = True
                idMap.dblID(X, y) = tmpID
              End If
              If (tmp1 = False) And (tmp2 = False) And (skipThisPart = False) Then ' person -> write unsafe ratio of 2 around
                If frmRunemaker.IsFriend(LCase(nameofgivenID)) = False Then
                  'write no safe around in a ratio of 2

                  If ((friendlyMode(idConnection) > 0) And (cond2 = True)) Then
                  
                  WriteNoSafeToAttack idMap, X - 2, y - 2
                  WriteNoSafeToAttack idMap, X - 1, y - 2
                  WriteNoSafeToAttack idMap, X, y - 2
                  WriteNoSafeToAttack idMap, X + 1, y - 2
                  WriteNoSafeToAttack idMap, X + 2, y - 2
                  
                  WriteNoSafeToAttack idMap, X - 2, y - 1
                  WriteNoSafeToAttack idMap, X - 1, y - 1
                  WriteNoSafeToAttack idMap, X, y - 1
                  WriteNoSafeToAttack idMap, X + 1, y - 1
                  WriteNoSafeToAttack idMap, X + 2, y - 1
                  
                  WriteNoSafeToAttack idMap, X - 2, y
                  WriteNoSafeToAttack idMap, X - 1, y
                  WriteNoSafeToAttack idMap, X, y
                  WriteNoSafeToAttack idMap, X + 1, y
                  WriteNoSafeToAttack idMap, X + 2, y
                  
                  WriteNoSafeToAttack idMap, X - 2, y + 1
                  WriteNoSafeToAttack idMap, X - 1, y + 1
                  WriteNoSafeToAttack idMap, X, y + 1
                  WriteNoSafeToAttack idMap, X + 1, y + 1
                  WriteNoSafeToAttack idMap, X + 2, y + 1
                  
                  WriteNoSafeToAttack idMap, X - 2, y + 2
                  WriteNoSafeToAttack idMap, X - 1, y + 2
                  WriteNoSafeToAttack idMap, X, y + 2
                  WriteNoSafeToAttack idMap, X + 1, y + 2
                  WriteNoSafeToAttack idMap, X + 2, y + 2
                  
                  End If
                End If
              End If
            End If
          ElseIf DatTiles(tileID).isField Then
            If (myMap.cost(X, y) < CostHandicap) Then
              myMap.cost(X, y) = CostHandicap
            End If
               If (myMap.cost(X - 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y)) < CostNearHandicap Then
                 myMap.cost(X - 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y + 1)) < CostNearHandicap Then
                 myMap.cost(X, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y)) < CostNearHandicap Then
                 myMap.cost(X + 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y - 1)) < CostNearHandicap Then
                 myMap.cost(X, y - 1) = CostNearHandicap
               End If
          ElseIf DatTiles(tileID).blocking Then
            myMap.cost(X, y) = CostBlock
            Exit For
          ElseIf (DatTiles(tileID).floorChangeDOWN = True) Or (DatTiles(tileID).floorChangeUP = True) Then
            If (DatTiles(tileID).requireShovel = False) And (DatTiles(tileID).requireRope = False) And (DatTiles(tileID).requireRightClick = False) Then
               myMap.cost(X, y) = CostBlock
               If (myMap.cost(X - 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y)) < CostNearHandicap Then
                 myMap.cost(X - 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y + 1)) < CostNearHandicap Then
                 myMap.cost(X, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y)) < CostNearHandicap Then
                 myMap.cost(X + 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y - 1)) < CostNearHandicap Then
                 myMap.cost(X, y - 1) = CostNearHandicap
               End If
               
              Exit For
            End If
          End If
        Next s
      End If
    Next X
  Next y

  'force start to be walkable
  myMap.cost(0, 0) = CostWalkable
End Sub


Public Function ProcessAttacks(idConnection As Integer) As Long
' new with priorities
 Dim orders As String
  Dim startX As Long
  Dim startY As Long
  Dim X As Long
  Dim y As Long
  Dim z As Long
  Dim s As Long
  Dim aRes As Long
  Dim tileID As Long
  Dim myMap As TypeAstarMatrix
  Dim idMap As TypeIDmap
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim Px As Long
  Dim Py As Long
  Dim lOrders As Long
  Dim chCompare As String
  Dim inRes As Integer
  Dim continue As Boolean
  Dim nameofgivenID As String
  Dim tmpID As Double
  Dim tmp1 As Boolean
  Dim tmp2 As Boolean
  Dim bestID As Double
  Dim possibleID As Double
  Dim bestHMM As Boolean
  Dim bestX As Long
  Dim bestY As Long
  Dim bestMelee As Double
  Dim bestDist As Long
  Dim tmpDist As Long
  Dim okEval As Boolean
  Dim saveCost As Long
  Dim playerS As String
  Dim xlim1 As Long
  Dim xlim2 As Long
  Dim ylim1 As Long
  Dim ylim2 As Long
  Dim friRes As TypeSpecialRes
  Dim cond2 As Boolean
  Dim mapLoaded As Boolean
  Dim runetoUseAsHmm As Long
  Dim BestPriority As Long
  Dim foundCurrentID As Boolean
  Dim currentPriority As Long
  Dim strNameOfMob As String
  Dim okeval2 As Boolean
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  foundCurrentID = False
  mapLoaded = False
  z = myZ(idConnection)
  
  If frmCavebot.Option1.Value = True Then
    MAX_LOCKWAIT = 100000
    cond2 = True
  Else
    If (moveRetry(idConnection) < MAX_LOCKWAIT) Then
      cond2 = True
    Else
      cond2 = False
    End If
  End If
  
  If friendlyMode(idConnection) = 2 Then
    friRes = PlayerOnScreen2(idConnection)
    playerS = friRes.str
    
    If playerS <> "" Then
      DelayAttacks(idConnection) = GetTickCount() + SetVeryFriendly_NOATTACKTIMER_ms
      If friRes.bln = True Then
        If isIgnoredcreature(idConnection, SelfDefenseID(idConnection)) = True Then
           GoTo continueNormal
        End If
        If SelfDefenseID(idConnection) <> TrainerOptions(idConnection).idToAvoid Then
          If TrainerOptions(idConnection).misc_stoplowhp = 1 Then
            If GetHPFromID(idConnection, SelfDefenseID(idConnection)) < TrainerOptions(idConnection).stoplowhpHP Then
              GoTo continueNormal
            Else
              bestX = friRes.bestX
              bestY = friRes.bestY
              bestID = SelfDefenseID(idConnection)
              bestMelee = friRes.bestMelee
              bestHMM = friRes.bestHMM
                strNameOfMob = LCase(GetNameFromID(idConnection, bestID))
                currentPriority = getKillPriority(idConnection, strNameOfMob)
              GoTo foundOne
            End If
          Else
          bestX = friRes.bestX
          bestY = friRes.bestY
          bestID = SelfDefenseID(idConnection)
          bestMelee = friRes.bestMelee
          bestHMM = friRes.bestHMM
            strNameOfMob = LCase(GetNameFromID(idConnection, bestID))
            currentPriority = getKillPriority(idConnection, strNameOfMob)
          GoTo foundOne
          End If
        Else
          SelfDefenseID(idConnection) = 0
        End If
      End If
      If lastAttackedID(idConnection) <> 0 Then
        If moveRetry(idConnection) < MAX_LOCKWAIT Then
          lastAttackedID(idConnection) = 0
          cavebotCurrentTargetPriority(idConnection) = 0
          If publicDebugMode = True Then
            aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Stopping ALL attacks because very friendly mode. (player='" & playerS & "')")
            DoEvents
          End If
          aRes = MeleeAttack(idConnection, 0)
          DoEvents
          SelfDefenseID(idConnection) = 0
          ProcessAttacks = 0
          Exit Function
        End If
      End If
      If moveRetry(idConnection) < MAX_LOCKWAIT Then
        If lastAttackedID(idConnection) <> 0 Then
          lastAttackedID(idConnection) = 0
          cavebotCurrentTargetPriority(idConnection) = 0
          aRes = MeleeAttack(idConnection, 0)
          DoEvents
        End If
        SelfDefenseID(idConnection) = 0
        ProcessAttacks = 0
        Exit Function
      End If
    End If
  End If
continueNormal:
  
  startX = 0
  startY = 0
  If mapLoaded = False Then
    LoadCurrentFloorIntoMatrix idConnection, myMap, idMap, False, cond2
    mapLoaded = True
  End If
  ' Now check all IDs, choose last ID if found, else choose better one
  bestID = 0
  bestHMM = False
  bestMelee = False
  bestDist = 10000
  BestPriority = -2000000000
  bestX = 0
  bestY = 0
  
  If DelayAttacks(idConnection) <= GetTickCount() Then
   xlim1 = -6
   xlim2 = 7
   ylim1 = -5
   ylim2 = 6
  Else
   xlim1 = -1
   xlim2 = 1
   ylim1 = -1
   ylim2 = 1
  End If
  
  For X = xlim1 To xlim2
    For y = ylim1 To ylim2
      If (idMap.isMelee(X, y) = True) Or (idMap.isHmm(X, y) = True) Then
        ' ok to attack that?
        okEval = idMap.isSafe(X, y)
        okeval2 = (Not isIgnoredcreature(idConnection, idMap.dblID(X, y)))
        okEval = okEval And okeval2
        strNameOfMob = LCase(GetNameFromID(idConnection, idMap.dblID(X, y)))
        currentPriority = getKillPriority(idConnection, strNameOfMob)
        ' is latest attacked?
        If (idMap.dblID(X, y) = lastAttackedID(idConnection)) And (okeval2 = True) Then
          ' there is way?
          If idMap.dblID(X, y) <> TrainerOptions(idConnection).idToAvoid Then
            saveCost = myMap.cost(X, y)
            myMap.cost(X, y) = CostWalkable
            orders = Astar(0, 0, X, y, myMap)
            myMap.cost(X, y) = saveCost
            If orders <> "X" Then
              If TrainerOptions(idConnection).misc_stoplowhp = 1 Then
                If GetHPFromID(idConnection, idMap.dblID(X, y)) >= TrainerOptions(idConnection).stoplowhpHP Then
                  If (currentPriority >= BestPriority) Then
                    bestID = idMap.dblID(X, y)
                    bestHMM = idMap.isHmm(X, y)
                    bestMelee = idMap.isMelee(X, y)
                    bestDist = ManhattanDistance(X, y, 0, 0)
                    bestX = X
                    bestY = y
                    BestPriority = currentPriority
                    foundCurrentID = True
                  End If
                  If (usingPriorities(idConnection) = False) Then
                    GoTo foundOne
                  End If
                End If
              Else
                If (currentPriority >= BestPriority) Then
                    bestID = idMap.dblID(X, y)
                    bestHMM = idMap.isHmm(X, y)
                    bestMelee = idMap.isMelee(X, y)
                    bestDist = ManhattanDistance(X, y, 0, 0)
                    bestX = X
                    bestY = y
                    BestPriority = currentPriority
                    foundCurrentID = True
                End If
                If (usingPriorities(idConnection) = False) Then
                  GoTo foundOne
                End If
              End If
            End If
          End If
        End If
  
        If okEval = True Then
          If idMap.dblID(X, y) <> TrainerOptions(idConnection).idToAvoid Then
            tmpDist = ManhattanDistance(X, y, 0, 0)
            ' nearest than latest one?
            If ((foundCurrentID = False) And (currentPriority = BestPriority) And (tmpDist < bestDist)) Or _
               (currentPriority > BestPriority) Then
              ' there is way?
              saveCost = myMap.cost(X, y)
              myMap.cost(X, y) = CostWalkable
              orders = Astar(0, 0, X, y, myMap)
              myMap.cost(X, y) = saveCost
              If orders = "" Then
                ' dist 0

                If TrainerOptions(idConnection).misc_stoplowhp = 1 Then
                  If GetHPFromID(idConnection, idMap.dblID(X, y)) >= TrainerOptions(idConnection).stoplowhpHP Then
                    bestID = idMap.dblID(X, y)
                    bestHMM = idMap.isHmm(X, y)
                    bestMelee = idMap.isMelee(X, y)
                    bestDist = 0
                    bestX = X
                    bestY = y
                    BestPriority = currentPriority
                  End If
                Else
                  bestID = idMap.dblID(X, y)
                  bestHMM = idMap.isHmm(X, y)
                  bestMelee = idMap.isMelee(X, y)
                  bestDist = 0
                  bestX = X
                  bestY = y
                  BestPriority = currentPriority
                End If

                
                
              ElseIf orders = "X" Then
               ' no direct path found
              Else
                If TrainerOptions(idConnection).misc_stoplowhp = 1 Then
                  If GetHPFromID(idConnection, idMap.dblID(X, y)) >= TrainerOptions(idConnection).stoplowhpHP Then
                    bestID = idMap.dblID(X, y)
                    bestHMM = idMap.isHmm(X, y)
                    bestMelee = idMap.isMelee(X, y)
                    bestDist = tmpDist
                    bestX = X
                    bestY = y
                    BestPriority = currentPriority
                  End If
                Else
                  bestID = idMap.dblID(X, y)
                  bestHMM = idMap.isHmm(X, y)
                  bestMelee = idMap.isMelee(X, y)
                  bestDist = tmpDist
                  bestX = X
                  bestY = y
                  BestPriority = currentPriority
                End If
              End If
            End If 'tmpdist<bestdist
          End If 'avoid id
        End If ' okeval
      End If 'ismelee or ishmm
    Next y
  Next X
  If bestID = 0 Then
    If lastAttackedID(idConnection) <> 0 Then
      If (playerS <> "") Then
        If publicDebugMode = True Then
          aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Stoping attacks because very friendly mode. (player='" & playerS & "')")
          DoEvents
        End If
      End If
      aRes = MeleeAttack(idConnection, 0)
      DoEvents
    End If
    lastAttackedID(idConnection) = 0
    SelfDefenseID(idConnection) = 0
    If prevAttackState(idConnection) = True Then
      ' we was attacking something before:
      ' check if we advanced in the script while following monsters.
      ' Allow a max skip of 2 lines
      RepositionScript idConnection, exeLine(idConnection), exeLine(idConnection) + 2
    End If
    ProcessAttacks = 0
    Exit Function
  End If
foundOne:
    If bestID <> 0 Then
        cavebotCurrentTargetPriority(idConnection) = BestPriority
    Else
        cavebotCurrentTargetPriority(idConnection) = 0
    End If
  If publicDebugMode = True Then
    If lastAttackedID(idConnection) <> bestID Then

      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] New target: '" & strNameOfMob & "' at relative pos x=" & CStr(bestX) & " y=" & CStr(bestY) & " priority=" & CStr(BestPriority))
    End If
  End If
  
  ' new in 9.14
  If adrNumberOfAttackClicks = &H0 Then ' new since blackd proxy 24.0
    If AvoidReAttacks(idConnection) = True Then
      If lastAttackedID(idConnection) <> 0 Then
        If lastAttackedID(idConnection) = bestID Then
          If publicDebugMode = True Then
            aRes = SendLogSystemMessageToClient(idConnection, "Retry melee attack: canceled")
            DoEvents
          End If
          bestMelee = False
        End If
      End If
    End If
  End If
  
  nameofgivenID = ""
  lastAttackedID(idConnection) = bestID
  If bestID > 0 Then
    nameofgivenID = GetNameFromID(idConnection, bestID)
  End If
  If bestHMM = True Then
    runetoUseAsHmm = getShotType(idConnection, nameofgivenID)
    If Round((myHP(idConnection) / myMaxHP(idConnection)) * 100) >= frmCavebot.scrollExorivis.Value Then
        aRes = CavebotRuneAttack(idConnection, bestID, LowByteOfLong(runetoUseAsHmm), HighByteOfLong(runetoUseAsHmm))
    End If
  End If
  If ((CavebotHaveSpecials(idConnection) = True) And (bestID > 0)) Then
        If ((isAvoid(idConnection, nameofgivenID) = True) Or (isExorivis(idConnection, nameofgivenID) = True)) Then
            For X = -1 To 1
                For y = -1 To 1
                   For s = 0 To 10
                        tileID = GetTheLong(Matrix(y, X, z, idConnection).s(s).t1, Matrix(y, X, z, idConnection).s(s).t2)
                        If tileID = 0 Then
                            Exit For
                        Else
                            possibleID = Matrix(y, X, z, idConnection).s(s).dblID
                            If ((possibleID <> 0) And (possibleID = bestID)) Then
                                If DoSpecialCavebot(idConnection, X, y, myZ(idConnection), bestID) = 1 Then
                                    ProcessAttacks = 1
                                    Exit Function
                                End If
                            End If
                        End If
                   Next s
                Next y
            Next X
        End If
        
  End If
  If bestMelee = True Then
    aRes = MeleeAttack(idConnection, bestID)
  End If
  If setFollowTarget(idConnection) = True Then
    PerformMove idConnection, myX(idConnection) + bestX, myY(idConnection) + bestY, myZ(idConnection)
  End If
  ProcessAttacks = 1
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Got error at ProcessAttacks"
  SelfDefenseID(idConnection) = 0
  ProcessAttacks = 0
End Function







Public Sub ProcessFindDown(idConnection As Integer, X As Integer, y As Integer, ByRef pMatrix As TypePathMatrix, ByRef fResult As TypePathResult)
  Dim nameofgivenID As String
  Dim tileID As Long
  Dim tmpID As Double
  Dim s As Byte
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  pMatrix.walkable(X, y) = pMatrix.walkable(X, y) Or pMatrix.walkable(X + 1, y + 1) Or pMatrix.walkable(X, y + 1) _
   Or pMatrix.walkable(X + 1, y) Or pMatrix.walkable(X - 1, y - 1) Or pMatrix.walkable(X, y - 1) _
   Or pMatrix.walkable(X - 1, y) Or pMatrix.walkable(X - 1, y + 1) Or pMatrix.walkable(X + 1, y - 1)
  If pMatrix.walkable(X, y) = True And fResult.id = 0 Then
    For s = 0 To 10
      tileID = GetTheLong(Matrix(y, X, myZ(idConnection), idConnection).s(s).t1, Matrix(y, X, myZ(idConnection), idConnection).s(s).t2)
      tmpID = Matrix(y, X, myZ(idConnection), idConnection).s(s).dblID
      If tmpID <> 0 Then
        pMatrix.walkable(X, y) = False
        Exit Sub
      ElseIf tileID = 0 Then
        If s = 0 Then
          pMatrix.walkable(X, y) = False
        End If
        Exit Sub
      ElseIf DatTiles(tileID).floorChangeDOWN = True Then
        If fResult.tileID = 0 Then
          fResult.tileID = tileID
          fResult.X = myX(idConnection) + X
          fResult.y = myY(idConnection) + y
          If DatTiles(tileID).requireShovel = True Then
            fResult.requireShovel = True
          End If
           If DatTiles(tileID).requireRightClick = True Then
            fResult.requireRightClick = True
          End If
          Exit Sub
        End If
      ElseIf DatTiles(tileID).blocking = True Then
        pMatrix.walkable(X, y) = False
        Exit Sub
      End If
    Next s
  End If
  Exit Sub
goterr:
    frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Got error at ProcessFindDown"
End Sub



'Public Function ProcessAttacks(idConnection As Integer) As Long
' Dim orders As String
'  Dim startX As Long
'  Dim startY As Long
'  Dim X As Long
'  Dim Y As Long
'  Dim Z As Long
'  Dim s As Long
'  Dim ares As Long
'  Dim tileID As Long
'  Dim myMap As TypeAstarMatrix
'  Dim idMap As TypeIDmap
'  Dim cPacket() As Byte
'  Dim sCheat As String
'  Dim px As Long
'  Dim py As Long
'  Dim lOrders As Long
'  Dim chCompare As String
'  Dim inRes As Integer
'  Dim continue As Boolean
'  Dim nameofgivenID As String
'  Dim tmpID As Double
'  Dim tmp1 As Boolean
'  Dim tmp2 As Boolean
'  Dim bestID As Double
'  Dim possibleID As Double
'  Dim bestHMM As Boolean
'  Dim bestX As Long
'  Dim bestY As Long
'  Dim bestMelee As Double
'  Dim bestDist As Long
'  Dim tmpDist As Long
'  Dim okEval As Boolean
'  Dim saveCost As Long
'  Dim playerS As String
'  Dim xlim1 As Long
'  Dim xlim2 As Long
'  Dim ylim1 As Long
'  Dim ylim2 As Long
'  Dim friRes As TypeSpecialRes
'  Dim cond2 As Boolean
'  Dim mapLoaded As Boolean
'  Dim runetoUseAsHmm As Long
'  #If FinalMode Then
'  On Error GoTo goterr
'  #End If
'  mapLoaded = False
'  Z = myZ(idConnection)
'
'  If frmCavebot.Option1.Value = True Then
'    MAX_LOCKWAIT = 100000
'    cond2 = True
'  Else
'    If (moveRetry(idConnection) < MAX_LOCKWAIT) Then
'      cond2 = True
'    Else
'      cond2 = False
'    End If
'  End If
'
'  If friendlyMode(idConnection) = 2 Then
'    friRes = PlayerOnScreen2(idConnection)
'    playerS = friRes.str
'
'    If playerS <> "" Then
'      DelayAttacks(idConnection) = GetTickCount() + SetVeryFriendly_NOATTACKTIMER_ms
'      If friRes.bln = True Then
'        If SelfDefenseID(idConnection) <> TrainerOptions(idConnection).idToAvoid Then
'          If TrainerOptions(idConnection).misc_stoplowhp = 1 Then
'            If GetHPFromID(idConnection, SelfDefenseID(idConnection)) < TrainerOptions(idConnection).stoplowhpHP Then
'              GoTo continueNormal
'            Else
'              bestX = friRes.bestX
'              bestY = friRes.bestY
'              bestID = SelfDefenseID(idConnection)
'              bestMelee = friRes.bestMelee
'              bestHMM = friRes.bestHMM
'              GoTo foundOne
'            End If
'          Else
'          bestX = friRes.bestX
'          bestY = friRes.bestY
'          bestID = SelfDefenseID(idConnection)
'          bestMelee = friRes.bestMelee
'          bestHMM = friRes.bestHMM
'          GoTo foundOne
'          End If
'        Else
'          SelfDefenseID(idConnection) = 0
'        End If
'      End If
'      If lastAttackedID(idConnection) <> 0 Then
'        If moveRetry(idConnection) < MAX_LOCKWAIT Then
'          lastAttackedID(idConnection) = 0
'          If publicDebugMode = True Then
'            ares = SendLogSystemMessageToClient(idConnection, "[Debug] Stopping ALL attacks because very friendly mode. (player='" & playerS & "')")
'            DoEvents
'          End If
'          ares = MeleeAttack(idConnection, 0)
'          DoEvents
'          SelfDefenseID(idConnection) = 0
'          ProcessAttacks = 0
'          Exit Function
'        End If
'      End If
'      If moveRetry(idConnection) < MAX_LOCKWAIT Then
'        If lastAttackedID(idConnection) <> 0 Then
'          lastAttackedID(idConnection) = 0
'          ares = MeleeAttack(idConnection, 0)
'          DoEvents
'        End If
'        SelfDefenseID(idConnection) = 0
'        ProcessAttacks = 0
'        Exit Function
'      End If
'    End If
'  End If
'continueNormal:
'
'  startX = 0
'  startY = 0
'  If mapLoaded = False Then
'    LoadCurrentFloorIntoMatrix idConnection, myMap, idMap, False, cond2
'    mapLoaded = True
'  End If
'  ' Now check all IDs, choose last ID if found, else choose better one
'  bestID = 0
'  bestHMM = False
'  bestMelee = False
'  bestDist = 10000
'  bestX = 0
'  bestY = 0
'
'  If DelayAttacks(idConnection) <= GetTickCount() Then
'   xlim1 = -6
'   xlim2 = 7
'   ylim1 = -5
'   ylim2 = 6
'  Else
'   xlim1 = -1
'   xlim2 = 1
'   ylim1 = -1
'   ylim2 = 1
'  End If
'
'  For X = xlim1 To xlim2
'    For Y = ylim1 To ylim2
'      If (idMap.isMelee(X, Y) = True) Or (idMap.isHmm(X, Y) = True) Then
'        ' ok to attack that?
'        okEval = idMap.isSafe(X, Y)
'
'        ' is latest attacked?
'        If idMap.dblID(X, Y) = lastAttackedID(idConnection) Then
'          ' there is way?
'          If idMap.dblID(X, Y) <> TrainerOptions(idConnection).idToAvoid Then
'            saveCost = myMap.cost(X, Y)
'            myMap.cost(X, Y) = CostWalkable
'            orders = Astar(0, 0, X, Y, myMap)
'            myMap.cost(X, Y) = saveCost
'            If orders <> "X" Then
'              If TrainerOptions(idConnection).misc_stoplowhp = 1 Then
'                If GetHPFromID(idConnection, idMap.dblID(X, Y)) >= TrainerOptions(idConnection).stoplowhpHP Then
'                  bestID = idMap.dblID(X, Y)
'                  bestHMM = idMap.isHmm(X, Y)
'                  bestMelee = idMap.isMelee(X, Y)
'                  bestDist = ManhattanDistance(X, Y, 0, 0)
'                  bestX = X
'                  bestY = Y
'                  GoTo foundOne
'                End If
'              Else
'                bestID = idMap.dblID(X, Y)
'                bestHMM = idMap.isHmm(X, Y)
'                bestMelee = idMap.isMelee(X, Y)
'                bestDist = ManhattanDistance(X, Y, 0, 0)
'                bestX = X
'                bestY = Y
'                GoTo foundOne
'              End If
'            End If
'          End If
'        End If
'
'        If okEval = True Then
'          If idMap.dblID(X, Y) <> TrainerOptions(idConnection).idToAvoid Then
'            tmpDist = ManhattanDistance(X, Y, 0, 0)
'            ' nearest than latest one?
'            If tmpDist < bestDist Then
'              ' there is way?
'              saveCost = myMap.cost(X, Y)
'              myMap.cost(X, Y) = CostWalkable
'              orders = Astar(0, 0, X, Y, myMap)
'              myMap.cost(X, Y) = saveCost
'              If orders = "" Then
'                If TrainerOptions(idConnection).misc_stoplowhp = 1 Then
'                  If GetHPFromID(idConnection, idMap.dblID(X, Y)) >= TrainerOptions(idConnection).stoplowhpHP Then
'                    bestID = idMap.dblID(X, Y)
'                    bestHMM = idMap.isHmm(X, Y)
'                    bestMelee = idMap.isMelee(X, Y)
'                    bestDist = 0
'                    bestX = X
'                    bestY = Y
'                    GoTo foundOne
'                  End If
'                Else
'                  bestID = idMap.dblID(X, Y)
'                  bestHMM = idMap.isHmm(X, Y)
'                  bestMelee = idMap.isMelee(X, Y)
'                  bestDist = 0
'                  bestX = X
'                  bestY = Y
'                  GoTo foundOne
'                End If
'              ElseIf orders = "X" Then
'               ' no direct path found
'              Else
'                If TrainerOptions(idConnection).misc_stoplowhp = 1 Then
'                  If GetHPFromID(idConnection, idMap.dblID(X, Y)) >= TrainerOptions(idConnection).stoplowhpHP Then
'                    bestID = idMap.dblID(X, Y)
'                    bestHMM = idMap.isHmm(X, Y)
'                    bestMelee = idMap.isMelee(X, Y)
'                    bestDist = tmpDist
'                    bestX = X
'                    bestY = Y
'                  End If
'                Else
'                  bestID = idMap.dblID(X, Y)
'                  bestHMM = idMap.isHmm(X, Y)
'                  bestMelee = idMap.isMelee(X, Y)
'                  bestDist = tmpDist
'                  bestX = X
'                  bestY = Y
'                End If
'              End If
'            End If 'tmpdist<bestdist
'          End If 'avoid id
'        End If ' okeval
'      End If 'ismelee or ishmm
'    Next Y
'  Next X
'  If bestID = 0 Then
'    If lastAttackedID(idConnection) <> 0 Then
'      If (playerS <> "") Then
'        If publicDebugMode = True Then
'          ares = SendLogSystemMessageToClient(idConnection, "[Debug] Stoping attacks because very friendly mode. (player='" & playerS & "')")
'          DoEvents
'        End If
'      End If
'      ares = MeleeAttack(idConnection, 0)
'      DoEvents
'    End If
'    lastAttackedID(idConnection) = 0
'    SelfDefenseID(idConnection) = 0
'    If prevAttackState(idConnection) = True Then
'      ' we was attacking something before:
'      ' check if we advanced in the script while following monsters.
'      ' Allow a max skip of 2 lines
'      RepositionScript idConnection, exeLine(idConnection), exeLine(idConnection) + 2
'    End If
'    ProcessAttacks = 0
'    Exit Function
'  End If
'foundOne:
'  If publicDebugMode = True Then
'    If lastAttackedID(idConnection) <> bestID Then
'      ares = SendLogSystemMessageToClient(idConnection, "[Debug] Selected new target at relative pos x=" & CStr(bestX) & " y=" & CStr(bestY))
'    End If
'  End If
'
'  ' new in 9.14
'  If AvoidReAttacks(idConnection) = True Then
'    If lastAttackedID(idConnection) <> 0 Then
'      If lastAttackedID(idConnection) = bestID Then
'        If publicDebugMode = True Then
'          ares = SendLogSystemMessageToClient(idConnection, "Retry melee attack: canceled")
'          DoEvents
'        End If
'        bestMelee = False
'      End If
'    End If
'  End If
'
'  nameofgivenID = ""
'  lastAttackedID(idConnection) = bestID
'  If bestID > 0 Then
'    nameofgivenID = GetNameFromID(idConnection, bestID)
'  End If
'  If bestHMM = True Then
'    runetoUseAsHmm = getShotType(idConnection, nameofgivenID)
'    If Round((myHP(idConnection) / myMaxHP(idConnection)) * 100) >= frmCavebot.scrollExorivis.Value Then
'        ares = CavebotRuneAttack(idConnection, bestID, LowByteOfLong(runetoUseAsHmm), HighByteOfLong(runetoUseAsHmm))
'    End If
'  End If
'  If ((CavebotHaveSpecials(idConnection) = True) And (bestID > 0)) Then
'        If ((isAvoid(idConnection, nameofgivenID) = True) Or (isExorivis(idConnection, nameofgivenID) = True)) Then
'            For X = -1 To 1
'                For Y = -1 To 1
'                   For s = 0 To 10
'                        tileID = GetTheLong(Matrix(Y, X, Z, idConnection).s(s).t1, Matrix(Y, X, Z, idConnection).s(s).t2)
'                        If tileID = 0 Then
'                            Exit For
'                        Else
'                            possibleID = Matrix(Y, X, Z, idConnection).s(s).dblID
'                            If ((possibleID <> 0) And (possibleID = bestID)) Then
'                                If DoSpecialCavebot(idConnection, X, Y, myZ(idConnection), bestID) = 1 Then
'                                    ProcessAttacks = 1
'                                    Exit Function
'                                End If
'                            End If
'                        End If
'                   Next s
'                Next Y
'            Next X
'        End If
'
'  End If
'  If bestMelee = True Then
'    ares = MeleeAttack(idConnection, bestID)
'  End If
'  If setFollowTarget(idConnection) = True Then
'    PerformMove idConnection, myX(idConnection) + bestX, myY(idConnection) + bestY, myZ(idConnection)
'  End If
'  ProcessAttacks = 1
'  Exit Function
'goterr:
'  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Got error at ProcessAttacks"
'  SelfDefenseID(idConnection) = 0
'  ProcessAttacks = 0
'End Function
'Public Sub ProcessFindDown(idConnection As Integer, X As Integer, Y As Integer, ByRef pMatrix As TypePathMatrix, ByRef fResult As TypePathResult)
'  Dim nameofgivenID As String
'  Dim tileID As Long
'  Dim tmpID As Double
'  Dim s As Byte
'  #If FinalMode Then
'  On Error GoTo goterr
'  #End If
'  pMatrix.walkable(X, Y) = pMatrix.walkable(X, Y) Or pMatrix.walkable(X + 1, Y + 1) Or pMatrix.walkable(X, Y + 1) _
'   Or pMatrix.walkable(X + 1, Y) Or pMatrix.walkable(X - 1, Y - 1) Or pMatrix.walkable(X, Y - 1) _
'   Or pMatrix.walkable(X - 1, Y) Or pMatrix.walkable(X - 1, Y + 1) Or pMatrix.walkable(X + 1, Y - 1)
'  If pMatrix.walkable(X, Y) = True And fResult.id = 0 Then
'    For s = 0 To 10
'      tileID = GetTheLong(Matrix(Y, X, myZ(idConnection), idConnection).s(s).t1, Matrix(Y, X, myZ(idConnection), idConnection).s(s).t2)
'      tmpID = Matrix(Y, X, myZ(idConnection), idConnection).s(s).dblID
'      If tmpID <> 0 Then
'        pMatrix.walkable(X, Y) = False
'        Exit Sub
'      ElseIf tileID = 0 Then
'        If s = 0 Then
'          pMatrix.walkable(X, Y) = False
'        End If
'        Exit Sub
'      ElseIf DatTiles(tileID).floorChangeDOWN = True Then
'        If fResult.tileID = 0 Then
'          fResult.tileID = tileID
'          fResult.X = myX(idConnection) + X
'          fResult.Y = myY(idConnection) + Y
'          If DatTiles(tileID).requireShovel = True Then
'            fResult.requireShovel = True
'          End If
'           If DatTiles(tileID).requireRightClick = True Then
'            fResult.requireRightClick = True
'          End If
'          Exit Sub
'        End If
'      ElseIf DatTiles(tileID).blocking = True Then
'        pMatrix.walkable(X, Y) = False
'        Exit Sub
'      End If
'    Next s
'  End If
'  Exit Sub
'goterr:
'    frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Got error at ProcessFindDown"
'End Sub
Public Function PerformMoveDown(Sid As Integer, X As Long, y As Long, z As Long) As TypeChangeFloorResult
  Dim pMatrix As TypePathMatrix
  Dim fResult As TypePathResult
  Dim xt As Long
  Dim yt As Long
  Dim xdif As Long
  Dim ydif As Long
  Dim aRes As Long
  Dim myres As TypeChangeFloorResult
  'myres.result=0 req_wait
  'myres.result=1 req_move
  'myres.result=2 req_click
  'myres.result=3 req_shovel
  'myres.result=4 req_rope
  'myres.result=5 req_random_move
  'myres.result>&H60 req_force_move
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  xdif = X - myX(Sid)
  ydif = y - myY(Sid)
  aRes = CLng(Int((2 * Rnd) + 1)) 'randomize this if
  If ((xdif < -7) Or (xdif > 8) Or (ydif < -5) Or (ydif > 6)) And (aRes = 1) Then
    'out of range: first move near
    myres.X = X
    myres.y = y
    myres.z = myZ(Sid)
    myres.result = 1 ' move
    PerformMoveDown = myres
    Exit Function
  End If
  For xt = -8 To 9
    For yt = -6 To 7
      pMatrix.walkable(xt, yt) = False
    Next yt
  Next xt
  pMatrix.walkable(0, 0) = True
  fResult.id = 0
  fResult.tileID = 0
  fResult.melee = False
  fResult.hmm = False
  fResult.requireShovel = False
  fResult.requireRope = False
  fResult.requireRightClick = False
  ProcessFindDown Sid, 0, 0, pMatrix, fResult

  If fResult.tileID > 0 Then
    lastFloorTrap(Sid) = -1
    If fResult.requireRightClick = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 2 ' click
    ElseIf fResult.requireShovel = True Then
      myres.X = 0
      myres.y = 0
      myres.z = 0
      myres.result = 5 ' random move
    Else
      myres.X = 0
      myres.y = 0
      myres.z = 0
      myres.result = 0 ' do nothing
    End If
    PerformMoveDown = myres
    Exit Function
  End If

  pMatrix.walkable(0, 0) = True
  'process counterclock circle from top left corner
  ProcessFindDown Sid, -1, -1, pMatrix, fResult
  ProcessFindDown Sid, -1, 0, pMatrix, fResult
  ProcessFindDown Sid, -1, 1, pMatrix, fResult
  ProcessFindDown Sid, 0, 1, pMatrix, fResult
  ProcessFindDown Sid, 1, 1, pMatrix, fResult
  ProcessFindDown Sid, 1, 0, pMatrix, fResult
  ProcessFindDown Sid, 1, -1, pMatrix, fResult
  ProcessFindDown Sid, 0, -1, pMatrix, fResult

   If fResult.tileID > 0 Then
    lastFloorTrap(Sid) = -1
    If fResult.requireRightClick = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 2 ' click
    ElseIf fResult.requireShovel = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 3 ' shovel
    Else
      myres.X = 0
      myres.y = 0
      myres.z = 0
      myres.result = 6 ' force move
      If (fResult.X = myX(Sid)) And (fResult.y = myY(Sid) - 1) Then
        myres.result = &H65 'step north
      ElseIf (fResult.X = myX(Sid) + 1) And (fResult.y = myY(Sid)) Then
        myres.result = &H66 'step right
      ElseIf (fResult.X = myX(Sid)) And (fResult.y = myY(Sid) + 1) Then
        myres.result = &H67 'step south
      ElseIf (fResult.X = myX(Sid) - 1) And (fResult.y = myY(Sid)) Then
        myres.result = &H68 'step left
      ElseIf (fResult.X = myX(Sid) + 1) And (fResult.y = myY(Sid) - 1) Then
        myres.result = &H6A 'step north + right
      ElseIf (fResult.X = myX(Sid) + 1) And (fResult.y = myY(Sid) + 1) Then
        myres.result = &H6B 'step south + right
      ElseIf (fResult.X = myX(Sid) - 1) And (fResult.y = myY(Sid) + 1) Then
        myres.result = &H6C 'step south + left
      ElseIf (fResult.X = myX(Sid) - 1) And (fResult.y = myY(Sid) - 1) Then
        myres.result = &H6D 'step north + left
      End If
    End If
    PerformMoveDown = myres
    Exit Function
  End If

  'process counterclock circle from top left corner
  ProcessFindDown Sid, -2, -2, pMatrix, fResult
  ProcessFindDown Sid, -2, -1, pMatrix, fResult
  ProcessFindDown Sid, -2, 0, pMatrix, fResult
  ProcessFindDown Sid, -2, 1, pMatrix, fResult
  ProcessFindDown Sid, -2, 2, pMatrix, fResult
  ProcessFindDown Sid, -1, 2, pMatrix, fResult
  ProcessFindDown Sid, 0, 2, pMatrix, fResult
  ProcessFindDown Sid, 1, 2, pMatrix, fResult
  ProcessFindDown Sid, 2, 2, pMatrix, fResult
  ProcessFindDown Sid, 2, 1, pMatrix, fResult
  ProcessFindDown Sid, 2, 0, pMatrix, fResult
  ProcessFindDown Sid, 2, -1, pMatrix, fResult
  ProcessFindDown Sid, 2, -2, pMatrix, fResult
  ProcessFindDown Sid, 1, -2, pMatrix, fResult
  ProcessFindDown Sid, 0, -2, pMatrix, fResult
  ProcessFindDown Sid, -1, -2, pMatrix, fResult

  If fResult.tileID > 0 Then
    lastFloorTrap(Sid) = -1
    If fResult.requireRightClick = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 2 ' click
    ElseIf fResult.requireShovel = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 3 ' shovel
    Else
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 1 ' move
    End If
    PerformMoveDown = myres
    Exit Function
  End If

  'process counterclock circle from top left corner
  ProcessFindDown Sid, -3, -3, pMatrix, fResult
  ProcessFindDown Sid, -3, -2, pMatrix, fResult
  ProcessFindDown Sid, -3, -1, pMatrix, fResult
  ProcessFindDown Sid, -3, 0, pMatrix, fResult
  ProcessFindDown Sid, -3, 1, pMatrix, fResult
  ProcessFindDown Sid, -3, 2, pMatrix, fResult
  ProcessFindDown Sid, -3, 3, pMatrix, fResult
  ProcessFindDown Sid, -2, 3, pMatrix, fResult
  ProcessFindDown Sid, -1, 3, pMatrix, fResult
  ProcessFindDown Sid, 0, 3, pMatrix, fResult
  ProcessFindDown Sid, 1, 3, pMatrix, fResult
  ProcessFindDown Sid, 2, 3, pMatrix, fResult
  ProcessFindDown Sid, 3, 3, pMatrix, fResult
  ProcessFindDown Sid, 3, 2, pMatrix, fResult
  ProcessFindDown Sid, 3, 1, pMatrix, fResult
  ProcessFindDown Sid, 3, 0, pMatrix, fResult
  ProcessFindDown Sid, 3, -1, pMatrix, fResult
  ProcessFindDown Sid, 3, -2, pMatrix, fResult
  ProcessFindDown Sid, 3, -3, pMatrix, fResult
  ProcessFindDown Sid, 2, -3, pMatrix, fResult
  ProcessFindDown Sid, 1, -3, pMatrix, fResult
  ProcessFindDown Sid, 0, -3, pMatrix, fResult
  ProcessFindDown Sid, -1, -3, pMatrix, fResult
  ProcessFindDown Sid, -2, -3, pMatrix, fResult

  If fResult.tileID > 0 Then
    lastFloorTrap(Sid) = -1
    If fResult.requireRightClick = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 2 ' click
    ElseIf fResult.requireShovel = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 3 ' shovel
    Else
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 1 ' move
    End If
    PerformMoveDown = myres
    Exit Function
  End If

 'process counterclock circle from top left corner
  ProcessFindDown Sid, -4, -4, pMatrix, fResult
  ProcessFindDown Sid, -4, -3, pMatrix, fResult
  ProcessFindDown Sid, -4, -2, pMatrix, fResult
  ProcessFindDown Sid, -4, -1, pMatrix, fResult
  ProcessFindDown Sid, -4, 0, pMatrix, fResult
  ProcessFindDown Sid, -4, 1, pMatrix, fResult
  ProcessFindDown Sid, -4, 2, pMatrix, fResult
  ProcessFindDown Sid, -4, 3, pMatrix, fResult
  ProcessFindDown Sid, -4, 4, pMatrix, fResult
  ProcessFindDown Sid, -3, 4, pMatrix, fResult
  ProcessFindDown Sid, -2, 4, pMatrix, fResult
  ProcessFindDown Sid, -1, 4, pMatrix, fResult
  ProcessFindDown Sid, 0, 4, pMatrix, fResult
  ProcessFindDown Sid, 1, 4, pMatrix, fResult
  ProcessFindDown Sid, 2, 4, pMatrix, fResult
  ProcessFindDown Sid, 3, 4, pMatrix, fResult
  ProcessFindDown Sid, 4, 4, pMatrix, fResult
  ProcessFindDown Sid, 4, 3, pMatrix, fResult
  ProcessFindDown Sid, 4, 2, pMatrix, fResult
  ProcessFindDown Sid, 4, 1, pMatrix, fResult
  ProcessFindDown Sid, 4, 0, pMatrix, fResult
  ProcessFindDown Sid, 4, -1, pMatrix, fResult
  ProcessFindDown Sid, 4, -2, pMatrix, fResult
  ProcessFindDown Sid, 4, -3, pMatrix, fResult
  ProcessFindDown Sid, 4, -4, pMatrix, fResult
  ProcessFindDown Sid, 3, -4, pMatrix, fResult
  ProcessFindDown Sid, 2, -4, pMatrix, fResult
  ProcessFindDown Sid, 1, -4, pMatrix, fResult
  ProcessFindDown Sid, 0, -4, pMatrix, fResult
  ProcessFindDown Sid, -1, -4, pMatrix, fResult
  ProcessFindDown Sid, -2, -4, pMatrix, fResult
  ProcessFindDown Sid, -3, -4, pMatrix, fResult

  If fResult.tileID > 0 Then
    lastFloorTrap(Sid) = -1
    If fResult.requireRightClick = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 2 ' click
    ElseIf fResult.requireShovel = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 3 ' shovel
    Else
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 1 ' move
    End If
    PerformMoveDown = myres
    Exit Function
  End If

 'process counterclock circle from top left corner
  ProcessFindDown Sid, -5, -5, pMatrix, fResult
  ProcessFindDown Sid, -5, -4, pMatrix, fResult
  ProcessFindDown Sid, -5, -3, pMatrix, fResult
  ProcessFindDown Sid, -5, -2, pMatrix, fResult
  ProcessFindDown Sid, -5, -1, pMatrix, fResult
  ProcessFindDown Sid, -5, 0, pMatrix, fResult
  ProcessFindDown Sid, -5, 1, pMatrix, fResult
  ProcessFindDown Sid, -5, 2, pMatrix, fResult
  ProcessFindDown Sid, -5, 3, pMatrix, fResult
  ProcessFindDown Sid, -5, 4, pMatrix, fResult
  ProcessFindDown Sid, -5, 5, pMatrix, fResult
  ProcessFindDown Sid, -5, 5, pMatrix, fResult
  ProcessFindDown Sid, -4, 5, pMatrix, fResult
  ProcessFindDown Sid, -3, 5, pMatrix, fResult
  ProcessFindDown Sid, -2, 5, pMatrix, fResult
  ProcessFindDown Sid, -1, 5, pMatrix, fResult
  ProcessFindDown Sid, 0, 5, pMatrix, fResult
  ProcessFindDown Sid, 1, 5, pMatrix, fResult
  ProcessFindDown Sid, 2, 5, pMatrix, fResult
  ProcessFindDown Sid, 3, 5, pMatrix, fResult
  ProcessFindDown Sid, 4, 5, pMatrix, fResult
  ProcessFindDown Sid, 5, 5, pMatrix, fResult
  ProcessFindDown Sid, 5, 4, pMatrix, fResult
  ProcessFindDown Sid, 5, 3, pMatrix, fResult
  ProcessFindDown Sid, 5, 2, pMatrix, fResult
  ProcessFindDown Sid, 5, 1, pMatrix, fResult
  ProcessFindDown Sid, 5, 0, pMatrix, fResult
  ProcessFindDown Sid, 5, -1, pMatrix, fResult
  ProcessFindDown Sid, 5, -2, pMatrix, fResult
  ProcessFindDown Sid, 5, -3, pMatrix, fResult
  ProcessFindDown Sid, 5, -4, pMatrix, fResult
  ProcessFindDown Sid, 5, -5, pMatrix, fResult
  ProcessFindDown Sid, 4, -5, pMatrix, fResult
  ProcessFindDown Sid, 3, -5, pMatrix, fResult
  ProcessFindDown Sid, 2, -5, pMatrix, fResult
  ProcessFindDown Sid, 1, -5, pMatrix, fResult
  ProcessFindDown Sid, 0, -5, pMatrix, fResult
  ProcessFindDown Sid, -1, -5, pMatrix, fResult
  ProcessFindDown Sid, -2, -5, pMatrix, fResult
  ProcessFindDown Sid, -3, -5, pMatrix, fResult
  ProcessFindDown Sid, -4, -5, pMatrix, fResult

  If fResult.tileID > 0 Then
    lastFloorTrap(Sid) = -1
    If fResult.requireRightClick = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 2 ' click
    ElseIf fResult.requireShovel = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 3 ' shovel
    Else
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 1 ' move
    End If
    PerformMoveDown = myres
    Exit Function
  End If

 'process counterclock circle from top left corner
  ProcessFindDown Sid, -6, -5, pMatrix, fResult
  ProcessFindDown Sid, -6, -4, pMatrix, fResult
  ProcessFindDown Sid, -6, -3, pMatrix, fResult
  ProcessFindDown Sid, -6, -2, pMatrix, fResult
  ProcessFindDown Sid, -5, -1, pMatrix, fResult
  ProcessFindDown Sid, -6, 0, pMatrix, fResult
  ProcessFindDown Sid, -6, 1, pMatrix, fResult
  ProcessFindDown Sid, -6, 2, pMatrix, fResult
  ProcessFindDown Sid, -6, 3, pMatrix, fResult
  ProcessFindDown Sid, -6, 4, pMatrix, fResult
  ProcessFindDown Sid, -6, 5, pMatrix, fResult

  ProcessFindDown Sid, 6, 5, pMatrix, fResult
  ProcessFindDown Sid, 6, 4, pMatrix, fResult
  ProcessFindDown Sid, 6, 3, pMatrix, fResult
  ProcessFindDown Sid, 6, 2, pMatrix, fResult
  ProcessFindDown Sid, 6, 1, pMatrix, fResult
  ProcessFindDown Sid, 6, 0, pMatrix, fResult
  ProcessFindDown Sid, 6, -1, pMatrix, fResult
  ProcessFindDown Sid, 6, -2, pMatrix, fResult
  ProcessFindDown Sid, 6, -3, pMatrix, fResult
  ProcessFindDown Sid, 6, -4, pMatrix, fResult
  ProcessFindDown Sid, 6, -5, pMatrix, fResult

  If fResult.tileID > 0 Then
    lastFloorTrap(Sid) = -1
    If fResult.requireRightClick = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 2 ' click
    ElseIf fResult.requireShovel = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 3 ' shovel
    Else
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 1 ' move
    End If
    PerformMoveDown = myres
    Exit Function
  End If

 'process counterclock circle from top left corner
  ProcessFindDown Sid, -7, -5, pMatrix, fResult
  ProcessFindDown Sid, -7, -4, pMatrix, fResult
  ProcessFindDown Sid, -7, -3, pMatrix, fResult
  ProcessFindDown Sid, -7, -2, pMatrix, fResult
  ProcessFindDown Sid, -7, -1, pMatrix, fResult
  ProcessFindDown Sid, -7, 0, pMatrix, fResult
  ProcessFindDown Sid, -7, 1, pMatrix, fResult
  ProcessFindDown Sid, -7, 2, pMatrix, fResult
  ProcessFindDown Sid, -7, 3, pMatrix, fResult
  ProcessFindDown Sid, -7, 4, pMatrix, fResult
  ProcessFindDown Sid, -7, 5, pMatrix, fResult
  ProcessFindDown Sid, 7, 5, pMatrix, fResult
  ProcessFindDown Sid, 7, 4, pMatrix, fResult
  ProcessFindDown Sid, 7, 3, pMatrix, fResult
  ProcessFindDown Sid, 7, 2, pMatrix, fResult
  ProcessFindDown Sid, 7, 1, pMatrix, fResult
  ProcessFindDown Sid, 7, 0, pMatrix, fResult
  ProcessFindDown Sid, 7, -1, pMatrix, fResult
  ProcessFindDown Sid, 7, -2, pMatrix, fResult
  ProcessFindDown Sid, 7, -3, pMatrix, fResult
  ProcessFindDown Sid, 7, -4, pMatrix, fResult
  ProcessFindDown Sid, 7, -5, pMatrix, fResult

  If fResult.tileID > 0 Then
    lastFloorTrap(Sid) = -1
    If fResult.requireRightClick = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 2 ' click
    ElseIf fResult.requireShovel = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 3 ' shovel
    Else
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 1 ' move
    End If
    PerformMoveDown = myres
    Exit Function
  End If

  ' New method: move back to last floor change
  myres.X = lastFloorChangeX(Sid)
  myres.y = lastFloorChangeY(Sid)
  myres.z = lastFloorChangeZ(Sid)
  myres.result = 1 ' move
  PerformMoveDown = myres
  Exit Function
goterr:
  myres.X = 0
  myres.y = 0
  myres.z = 0
  myres.result = 0 ' error ... wait and hope better luck next call
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Got error at PerformMoveDown"
  PerformMoveDown = myres
End Function

Public Sub ProcessFindUp(idConnection As Integer, X As Integer, y As Integer, ByRef pMatrix As TypePathMatrix, ByRef fResult As TypePathResult)
  Dim nameofgivenID As String
  Dim tileID As Long
  Dim tmpID As Double
  Dim s As Byte
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  pMatrix.walkable(X, y) = pMatrix.walkable(X, y) Or pMatrix.walkable(X + 1, y + 1) Or pMatrix.walkable(X, y + 1) _
   Or pMatrix.walkable(X + 1, y) Or pMatrix.walkable(X - 1, y - 1) Or pMatrix.walkable(X, y - 1) _
   Or pMatrix.walkable(X - 1, y) Or pMatrix.walkable(X - 1, y + 1) Or pMatrix.walkable(X + 1, y - 1)
  If pMatrix.walkable(X, y) = True And fResult.id = 0 Then
    For s = 0 To 10
      tileID = GetTheLong(Matrix(y, X, myZ(idConnection), idConnection).s(s).t1, Matrix(y, X, myZ(idConnection), idConnection).s(s).t2)
      tmpID = Matrix(y, X, myZ(idConnection), idConnection).s(s).dblID
      If tmpID <> 0 Then
        pMatrix.walkable(X, y) = False
        Exit Sub
      ElseIf tileID = 0 Then
        If s = 0 Then
          pMatrix.walkable(X, y) = False
        End If
        Exit Sub
      ElseIf DatTiles(tileID).floorChangeUP = True Then
        If fResult.tileID = 0 Then
          fResult.tileID = tileID
          fResult.X = myX(idConnection) + X
          fResult.y = myY(idConnection) + y
          If DatTiles(tileID).requireRope = True Then
            fResult.requireRope = True
          End If
           If DatTiles(tileID).requireRightClick = True Then
            fResult.requireRightClick = True
          End If
          Exit Sub
        End If
      ElseIf DatTiles(tileID).blocking = True Then
        pMatrix.walkable(X, y) = False
        Exit Sub
      End If
    Next s
  End If
  Exit Sub
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Got error at ProcessFindUp"

End Sub

Public Function PerformMoveUp(Sid As Integer, X As Long, y As Long, z As Long) As TypeChangeFloorResult
  Dim pMatrix As TypePathMatrix
  Dim fResult As TypePathResult
  Dim xt As Long
  Dim yt As Long
  Dim xdif As Long
  Dim ydif As Long
  Dim aRes As Long
  Dim myres As TypeChangeFloorResult
  'myres.result=0 req_wait
  'myres.result=1 req_move
  'myres.result=2 req_click
  'myres.result=3 req_shovel
  'myres.result=4 req_rope
  'myres.result=5 req_random_move
  'myres.result>&H60 req_force_move
  
  'CLICK
  'PerformUseItem sid, fResult.x, fResult.y, myZ(sid)
  'USE ITEM
  'aRes = PerformUseMyItem(sid, LowByteOfLong(tileID_Rope), HighByteOfLong(tileID_Rope), fResult.x, fResult.y, myZ(sid))
  
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  xdif = X - myX(Sid)
  ydif = y - myY(Sid)
  aRes = CLng(Int((2 * Rnd) + 1)) 'randomize this if
  If ((xdif < -7) Or (xdif > 8) Or (ydif < -5) Or (ydif > 6)) And (aRes = 1) Then
    'out of range: first move near
    myres.X = X
    myres.y = y
    myres.z = myZ(Sid)
    myres.result = 1 ' move
    PerformMoveUp = myres
    Exit Function
  End If
  For xt = -8 To 9
    For yt = -6 To 7
      pMatrix.walkable(xt, yt) = False
    Next yt
  Next xt
  pMatrix.walkable(0, 0) = True
  fResult.id = 0
  fResult.tileID = 0
  fResult.melee = False
  fResult.hmm = False
  fResult.requireShovel = False
  fResult.requireRope = False
  fResult.requireRightClick = False
  ProcessFindUp Sid, 0, 0, pMatrix, fResult
  
  If fResult.tileID > 0 Then
    lastFloorTrap(Sid) = -1
    If fResult.requireRightClick = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 2 ' click
    ElseIf fResult.requireRope = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 4 ' rope
    Else
      ' do nothing
      myres.X = 0
      myres.y = 0
      myres.z = 0
      myres.result = 0 ' wait
    End If
    PerformMoveUp = myres
    Exit Function
  End If
  
  pMatrix.walkable(0, 0) = True
  'process counterclock circle from top left corner
  ProcessFindUp Sid, -1, -1, pMatrix, fResult
  ProcessFindUp Sid, -1, 0, pMatrix, fResult
  ProcessFindUp Sid, -1, 1, pMatrix, fResult
  ProcessFindUp Sid, 0, 1, pMatrix, fResult
  ProcessFindUp Sid, 1, 1, pMatrix, fResult
  ProcessFindUp Sid, 1, 0, pMatrix, fResult
  ProcessFindUp Sid, 1, -1, pMatrix, fResult
  ProcessFindUp Sid, 0, -1, pMatrix, fResult
  
  If fResult.tileID > 0 Then
    lastFloorTrap(Sid) = -1
    If fResult.requireRightClick = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 2 ' click
    ElseIf fResult.requireRope = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 4 ' rope
    Else
      myres.X = 0
      myres.y = 0
      myres.z = 0
      If (fResult.X = myX(Sid)) And (fResult.y = myY(Sid) - 1) Then
        myres.result = &H65 'step north
      ElseIf (fResult.X = myX(Sid) + 1) And (fResult.y = myY(Sid)) Then
        myres.result = &H66  'step right
      ElseIf (fResult.X = myX(Sid)) And (fResult.y = myY(Sid) + 1) Then
        myres.result = &H67  'step south
      ElseIf (fResult.X = myX(Sid) - 1) And (fResult.y = myY(Sid)) Then
        myres.result = &H68  'step left
      ElseIf (fResult.X = myX(Sid) + 1) And (fResult.y = myY(Sid) - 1) Then
        myres.result = &H6A  'step north + right
      ElseIf (fResult.X = myX(Sid) + 1) And (fResult.y = myY(Sid) + 1) Then
        myres.result = &H6B  'step south + right
      ElseIf (fResult.X = myX(Sid) - 1) And (fResult.y = myY(Sid) + 1) Then
        myres.result = &H6C  'step south + left
      ElseIf (fResult.X = myX(Sid) - 1) And (fResult.y = myY(Sid) - 1) Then
        myres.result = &H6D  'step north + left
      End If
    End If
    PerformMoveUp = myres
    Exit Function
  End If
  
  'process counterclock circle from top left corner
  ProcessFindUp Sid, -2, -2, pMatrix, fResult
  ProcessFindUp Sid, -2, -1, pMatrix, fResult
  ProcessFindUp Sid, -2, 0, pMatrix, fResult
  ProcessFindUp Sid, -2, 1, pMatrix, fResult
  ProcessFindUp Sid, -2, 2, pMatrix, fResult
  ProcessFindUp Sid, -1, 2, pMatrix, fResult
  ProcessFindUp Sid, 0, 2, pMatrix, fResult
  ProcessFindUp Sid, 1, 2, pMatrix, fResult
  ProcessFindUp Sid, 2, 2, pMatrix, fResult
  ProcessFindUp Sid, 2, 1, pMatrix, fResult
  ProcessFindUp Sid, 2, 0, pMatrix, fResult
  ProcessFindUp Sid, 2, -1, pMatrix, fResult
  ProcessFindUp Sid, 2, -2, pMatrix, fResult
  ProcessFindUp Sid, 1, -2, pMatrix, fResult
  ProcessFindUp Sid, 0, -2, pMatrix, fResult
  ProcessFindUp Sid, -1, -2, pMatrix, fResult
  
  If fResult.tileID > 0 Then
    lastFloorTrap(Sid) = -1
    If fResult.requireRightClick = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 2 ' click
    ElseIf fResult.requireRope = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 4 ' rope
    Else
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 1 ' move
    End If
    PerformMoveUp = myres
    Exit Function
  End If
  
  'process counterclock circle from top left corner
  ProcessFindUp Sid, -3, -3, pMatrix, fResult
  ProcessFindUp Sid, -3, -2, pMatrix, fResult
  ProcessFindUp Sid, -3, -1, pMatrix, fResult
  ProcessFindUp Sid, -3, 0, pMatrix, fResult
  ProcessFindUp Sid, -3, 1, pMatrix, fResult
  ProcessFindUp Sid, -3, 2, pMatrix, fResult
  ProcessFindUp Sid, -3, 3, pMatrix, fResult
  ProcessFindUp Sid, -2, 3, pMatrix, fResult
  ProcessFindUp Sid, -1, 3, pMatrix, fResult
  ProcessFindUp Sid, 0, 3, pMatrix, fResult
  ProcessFindUp Sid, 1, 3, pMatrix, fResult
  ProcessFindUp Sid, 2, 3, pMatrix, fResult
  ProcessFindUp Sid, 3, 3, pMatrix, fResult
  ProcessFindUp Sid, 3, 2, pMatrix, fResult
  ProcessFindUp Sid, 3, 1, pMatrix, fResult
  ProcessFindUp Sid, 3, 0, pMatrix, fResult
  ProcessFindUp Sid, 3, -1, pMatrix, fResult
  ProcessFindUp Sid, 3, -2, pMatrix, fResult
  ProcessFindUp Sid, 3, -3, pMatrix, fResult
  ProcessFindUp Sid, 2, -3, pMatrix, fResult
  ProcessFindUp Sid, 1, -3, pMatrix, fResult
  ProcessFindUp Sid, 0, -3, pMatrix, fResult
  ProcessFindUp Sid, -1, -3, pMatrix, fResult
  ProcessFindUp Sid, -2, -3, pMatrix, fResult
  
  If fResult.tileID > 0 Then
    lastFloorTrap(Sid) = -1
    If fResult.requireRightClick = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 2 ' click
    ElseIf fResult.requireRope = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 4 ' rope
    Else
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 1 ' move
    End If
    PerformMoveUp = myres
    Exit Function
  End If
  
 'process counterclock circle from top left corner
  ProcessFindUp Sid, -4, -4, pMatrix, fResult
  ProcessFindUp Sid, -4, -3, pMatrix, fResult
  ProcessFindUp Sid, -4, -2, pMatrix, fResult
  ProcessFindUp Sid, -4, -1, pMatrix, fResult
  ProcessFindUp Sid, -4, 0, pMatrix, fResult
  ProcessFindUp Sid, -4, 1, pMatrix, fResult
  ProcessFindUp Sid, -4, 2, pMatrix, fResult
  ProcessFindUp Sid, -4, 3, pMatrix, fResult
  ProcessFindUp Sid, -4, 4, pMatrix, fResult
  ProcessFindUp Sid, -3, 4, pMatrix, fResult
  ProcessFindUp Sid, -2, 4, pMatrix, fResult
  ProcessFindUp Sid, -1, 4, pMatrix, fResult
  ProcessFindUp Sid, 0, 4, pMatrix, fResult
  ProcessFindUp Sid, 1, 4, pMatrix, fResult
  ProcessFindUp Sid, 2, 4, pMatrix, fResult
  ProcessFindUp Sid, 3, 4, pMatrix, fResult
  ProcessFindUp Sid, 4, 4, pMatrix, fResult
  ProcessFindUp Sid, 4, 3, pMatrix, fResult
  ProcessFindUp Sid, 4, 2, pMatrix, fResult
  ProcessFindUp Sid, 4, 1, pMatrix, fResult
  ProcessFindUp Sid, 4, 0, pMatrix, fResult
  ProcessFindUp Sid, 4, -1, pMatrix, fResult
  ProcessFindUp Sid, 4, -2, pMatrix, fResult
  ProcessFindUp Sid, 4, -3, pMatrix, fResult
  ProcessFindUp Sid, 4, -4, pMatrix, fResult
  ProcessFindUp Sid, 3, -4, pMatrix, fResult
  ProcessFindUp Sid, 2, -4, pMatrix, fResult
  ProcessFindUp Sid, 1, -4, pMatrix, fResult
  ProcessFindUp Sid, 0, -4, pMatrix, fResult
  ProcessFindUp Sid, -1, -4, pMatrix, fResult
  ProcessFindUp Sid, -2, -4, pMatrix, fResult
  ProcessFindUp Sid, -3, -4, pMatrix, fResult
  
  If fResult.tileID > 0 Then
    lastFloorTrap(Sid) = -1
    If fResult.requireRightClick = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 2 ' click
    ElseIf fResult.requireRope = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 4 ' rope
    Else
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 1 ' move
    End If
    PerformMoveUp = myres
    Exit Function
  End If
  
 'process counterclock circle from top left corner
  ProcessFindUp Sid, -5, -5, pMatrix, fResult
  ProcessFindUp Sid, -5, -4, pMatrix, fResult
  ProcessFindUp Sid, -5, -3, pMatrix, fResult
  ProcessFindUp Sid, -5, -2, pMatrix, fResult
  ProcessFindUp Sid, -5, -1, pMatrix, fResult
  ProcessFindUp Sid, -5, 0, pMatrix, fResult
  ProcessFindUp Sid, -5, 1, pMatrix, fResult
  ProcessFindUp Sid, -5, 2, pMatrix, fResult
  ProcessFindUp Sid, -5, 3, pMatrix, fResult
  ProcessFindUp Sid, -5, 4, pMatrix, fResult
  ProcessFindUp Sid, -5, 5, pMatrix, fResult
  ProcessFindUp Sid, -5, 5, pMatrix, fResult
  ProcessFindUp Sid, -4, 5, pMatrix, fResult
  ProcessFindUp Sid, -3, 5, pMatrix, fResult
  ProcessFindUp Sid, -2, 5, pMatrix, fResult
  ProcessFindUp Sid, -1, 5, pMatrix, fResult
  ProcessFindUp Sid, 0, 5, pMatrix, fResult
  ProcessFindUp Sid, 1, 5, pMatrix, fResult
  ProcessFindUp Sid, 2, 5, pMatrix, fResult
  ProcessFindUp Sid, 3, 5, pMatrix, fResult
  ProcessFindUp Sid, 4, 5, pMatrix, fResult
  ProcessFindUp Sid, 5, 5, pMatrix, fResult
  ProcessFindUp Sid, 5, 4, pMatrix, fResult
  ProcessFindUp Sid, 5, 3, pMatrix, fResult
  ProcessFindUp Sid, 5, 2, pMatrix, fResult
  ProcessFindUp Sid, 5, 1, pMatrix, fResult
  ProcessFindUp Sid, 5, 0, pMatrix, fResult
  ProcessFindUp Sid, 5, -1, pMatrix, fResult
  ProcessFindUp Sid, 5, -2, pMatrix, fResult
  ProcessFindUp Sid, 5, -3, pMatrix, fResult
  ProcessFindUp Sid, 5, -4, pMatrix, fResult
  ProcessFindUp Sid, 5, -5, pMatrix, fResult
  ProcessFindUp Sid, 4, -5, pMatrix, fResult
  ProcessFindUp Sid, 3, -5, pMatrix, fResult
  ProcessFindUp Sid, 2, -5, pMatrix, fResult
  ProcessFindUp Sid, 1, -5, pMatrix, fResult
  ProcessFindUp Sid, 0, -5, pMatrix, fResult
  ProcessFindUp Sid, -1, -5, pMatrix, fResult
  ProcessFindUp Sid, -2, -5, pMatrix, fResult
  ProcessFindUp Sid, -3, -5, pMatrix, fResult
  ProcessFindUp Sid, -4, -5, pMatrix, fResult
  
  If fResult.tileID > 0 Then
    lastFloorTrap(Sid) = -1
    If fResult.requireRightClick = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 2 ' click
    ElseIf fResult.requireRope = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 4 ' rope
    Else
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 1 ' move
    End If
    PerformMoveUp = myres
    Exit Function
  End If
    
 'process counterclock circle from top left corner
  ProcessFindUp Sid, -6, -5, pMatrix, fResult
  ProcessFindUp Sid, -6, -4, pMatrix, fResult
  ProcessFindUp Sid, -6, -3, pMatrix, fResult
  ProcessFindUp Sid, -6, -2, pMatrix, fResult
  ProcessFindUp Sid, -5, -1, pMatrix, fResult
  ProcessFindUp Sid, -6, 0, pMatrix, fResult
  ProcessFindUp Sid, -6, 1, pMatrix, fResult
  ProcessFindUp Sid, -6, 2, pMatrix, fResult
  ProcessFindUp Sid, -6, 3, pMatrix, fResult
  ProcessFindUp Sid, -6, 4, pMatrix, fResult
  ProcessFindUp Sid, -6, 5, pMatrix, fResult
  ProcessFindUp Sid, 6, 5, pMatrix, fResult
  ProcessFindUp Sid, 6, 4, pMatrix, fResult
  ProcessFindUp Sid, 6, 3, pMatrix, fResult
  ProcessFindUp Sid, 6, 2, pMatrix, fResult
  ProcessFindUp Sid, 6, 1, pMatrix, fResult
  ProcessFindUp Sid, 6, 0, pMatrix, fResult
  ProcessFindUp Sid, 6, -1, pMatrix, fResult
  ProcessFindUp Sid, 6, -2, pMatrix, fResult
  ProcessFindUp Sid, 6, -3, pMatrix, fResult
  ProcessFindUp Sid, 6, -4, pMatrix, fResult
  ProcessFindUp Sid, 6, -5, pMatrix, fResult
  
  If fResult.tileID > 0 Then
    lastFloorTrap(Sid) = -1
    If fResult.requireRightClick = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 2 ' click
    ElseIf fResult.requireRope = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 4 ' rope
    Else
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 1 ' move
    End If
    PerformMoveUp = myres
    Exit Function
  End If
      
 'process counterclock circle from top left corner
  ProcessFindUp Sid, -7, -5, pMatrix, fResult
  ProcessFindUp Sid, -7, -4, pMatrix, fResult
  ProcessFindUp Sid, -7, -3, pMatrix, fResult
  ProcessFindUp Sid, -7, -2, pMatrix, fResult
  ProcessFindUp Sid, -7, -1, pMatrix, fResult
  ProcessFindUp Sid, -7, 0, pMatrix, fResult
  ProcessFindUp Sid, -7, 1, pMatrix, fResult
  ProcessFindUp Sid, -7, 2, pMatrix, fResult
  ProcessFindUp Sid, -7, 3, pMatrix, fResult
  ProcessFindUp Sid, -7, 4, pMatrix, fResult
  ProcessFindUp Sid, -7, 5, pMatrix, fResult
  ProcessFindUp Sid, 7, 5, pMatrix, fResult
  ProcessFindUp Sid, 7, 4, pMatrix, fResult
  ProcessFindUp Sid, 7, 3, pMatrix, fResult
  ProcessFindUp Sid, 7, 2, pMatrix, fResult
  ProcessFindUp Sid, 7, 1, pMatrix, fResult
  ProcessFindUp Sid, 7, 0, pMatrix, fResult
  ProcessFindUp Sid, 7, -1, pMatrix, fResult
  ProcessFindUp Sid, 7, -2, pMatrix, fResult
  ProcessFindUp Sid, 7, -3, pMatrix, fResult
  ProcessFindUp Sid, 7, -4, pMatrix, fResult
  ProcessFindUp Sid, 7, -5, pMatrix, fResult
  
  If fResult.tileID > 0 Then
    lastFloorTrap(Sid) = -1
    If fResult.requireRightClick = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 2 ' click
    ElseIf fResult.requireRope = True Then
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 4 ' rope
    Else
      myres.X = fResult.X
      myres.y = fResult.y
      myres.z = myZ(Sid)
      myres.result = 1 ' move
    End If
    PerformMoveUp = myres
    Exit Function
  End If

  ' New method: move back to last floor change
  myres.X = lastFloorChangeX(Sid)
  myres.y = lastFloorChangeY(Sid)
  myres.z = lastFloorChangeZ(Sid)
  myres.result = 1 ' move
  PerformMoveUp = myres
  Exit Function
goterr:
  myres.X = 0
  myres.y = 0
  myres.z = 0
  myres.result = 0 ' error ... wait and hope better luck next call
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Got error at PerformMoveUp : " & Err.Description
  PerformMoveUp = myres
End Function

Public Function UseItemHere(idConnection As Integer, b1 As Byte, b2 As Byte, X As Long, y As Long, z As Long, parS As Byte) As Long
    Dim aRes As Long
    Dim cPacket() As Byte
    Dim sCheat As String
    Dim fRes As TypeSearchItemResult2
    Dim tileSTR As String
    Dim relX As Long
    Dim relY As Long
    Dim inRes As Integer
    Dim tileID As Long
    Dim i As Long
    Dim stackS As Byte
    Dim continue As Boolean
    Dim ts As Byte
    Dim shouldBeVisible As Boolean
    shouldBeVisible = False
    #If FinalMode Then
        On Error GoTo errclose
    #End If
    'aRes = SendLogSystemMessageToClient(idConnection, "Trying to use item to change floor")
    ' DoEvents
    tileSTR = "00 00"
    stackS = 0
    
    relX = X - myX(idConnection)
    relY = y - myY(idConnection)
    If ((relX < -7) Or (relX > 8) Or (relY < -5) Or (relY > 6)) Then
        aRes = SendLogSystemMessageToClient(idConnection, "You are to far from destination")
        DoEvents
        UseItemHere = -1
        Exit Function
    End If
    continue = True

    ts = parS
    tileID = GetTheLong(Matrix(relY, relX, z, idConnection).s(ts).t1, Matrix(relY, relX, z, idConnection).s(ts).t2)
    stackS = ts
    tileSTR = FiveChrLon(tileID)
  
    If ((frmHardcoreCheats.chkTotalWaste.Value = True) And (TibiaVersionLong >= 773) And (shouldBeVisible = False)) Then
        GoTo justdoit
    End If
    If mySlot(idConnection, SLOT_AMMUNITION).t1 = b1 And _
        mySlot(idConnection, SLOT_AMMUNITION).t2 = b2 Then 'use from ammo
    ' 11 00 83 FF FF 0A 00 00 7D 0B 00 EB 7C 99 7D 0C 80 01 00
    sCheat = "83 FF FF 0A 00 00 " & GoodHex(b1) & " " & GoodHex(b2) & " 00 " & GetHexStrFromPosition(X, y, z) & " " & tileSTR & " " & GoodHex(stackS)
   ' SendLogSystemMessageToClient idConnection, sCheat
   ' DoEvents
   
    SafeCastCheatString "UseItemHere1", idConnection, sCheat
  Else ' use from bp
    ' 11 00 83 FF FF 40 00 00 7D 0B 00 EB 7C 99 7D 0C 80 01 00
    fRes.foundcount = 0
    fRes = SearchItem(idConnection, b1, b2)
    If fRes.foundcount > 0 Then
      sCheat = "83 FF FF " & GoodHex(&H40 + fRes.bpID) & " 00 " & _
       GoodHex(fRes.slotID) & " " & GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(fRes.slotID) & " " & GetHexStrFromPosition(X, y, z) & " " & tileSTR & " " & GoodHex(stackS)
       SafeCastCheatString "UseItemHere2", idConnection, sCheat
    Else
justdoit:
      If ((((frmHardcoreCheats.chkEnhancedCheats.Value = True) Or (frmHardcoreCheats.chkTotalWaste.Value = True)) And (TibiaVersionLong >= 773))) And (shouldBeVisible = False) Then
          ' NEW
         sCheat = "83 FF FF 00 00 00 " & GoodHex(b1) & " " & GoodHex(b2) & " 00 " & GetHexStrFromPosition(X, y, z) & " " & tileSTR & " " & GoodHex(stackS)

         SafeCastCheatString "UseItemHere3", idConnection, sCheat
         UseItemHere = 0
         Exit Function
      Else
      If publicDebugMode = True Then
        aRes = SendSystemMessageToClient(idConnection, "Required item was not found")
        DoEvents
      End If
      UseItemHere = -1
      Exit Function
      End If
    End If
  End If
  UseItemHere = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at PerformUseMyItem #"
  frmMain.DoCloseActions idConnection
  DoEvents
  UseItemHere = -1
End Function

Public Function PerformUseMyItem(idConnection As Integer, b1 As Byte, b2 As Byte, X As Long, y As Long, z As Long, Optional shouldBeVisible As Boolean = False, Optional shovelMode As Boolean = False) As Long
    Dim aRes As Long
    Dim cPacket() As Byte
    Dim sCheat As String
    Dim fRes As TypeSearchItemResult2
    Dim tileSTR As String
    Dim relX As Long
    Dim relY As Long
    Dim inRes As Integer
    Dim tileID As Long
    Dim i As Long
    Dim stackS As Byte
    Dim continue As Boolean
    Dim ts As Byte
    #If FinalMode Then
        On Error GoTo errclose
    #End If
    'aRes = SendLogSystemMessageToClient(idConnection, "Trying to use item to change floor")
    ' DoEvents
    tileSTR = "00 00"
    stackS = 0
    
    relX = X - myX(idConnection)
    relY = y - myY(idConnection)
    If ((relX < -7) Or (relX > 8) Or (relY < -5) Or (relY > 6)) Then
        aRes = SendLogSystemMessageToClient(idConnection, "You are to far from a floor changer")
        DoEvents
        PerformUseMyItem = -1
        Exit Function
    End If
    continue = True
    If TibiaVersionLong <= 750 Then
        ' inspect players
        For ts = 1 To 10
            tileID = GetTheLong(Matrix(relY, relX, z, idConnection).s(ts).t1, Matrix(relY, relX, z, idConnection).s(ts).t2)
            If tileID = 97 Then
                stackS = ts
                tileSTR = "63 00"
                continue = False
                Exit For
            End If
        Next ts
        ' inspect items
        If continue = True Then
            For ts = 1 To 10
                tileID = GetTheLong(Matrix(relY, relX, z, idConnection).s(ts).t1, Matrix(relY, relX, z, idConnection).s(ts).t2)
                If DatTiles(tileID).alwaysOnTop = False And tileID <> 0 Then
                    stackS = ts
                    tileSTR = FiveChrLon(tileID)
                    continue = False
                    Exit For
                End If
            Next ts
        End If
        ' inspect ontop
        If continue = True Then
            For ts = 1 To 10
                tileID = GetTheLong(Matrix(relY, relX, z, idConnection).s(ts).t1, Matrix(relY, relX, z, idConnection).s(ts).t2)
                If DatTiles(tileID).alwaysOnTop = True And DatTiles(tileID).multitype = False Then
                    stackS = ts
                    tileSTR = FiveChrLon(tileID)
                    continue = False
                    Exit For
                End If
            Next ts
        End If
        ' inspect ground
        If continue = True Then
            tileID = GetTheLong(Matrix(relY, relX, z, idConnection).s(0).t1, Matrix(relY, relX, z, idConnection).s(0).t2)
            tileSTR = FiveChrLon(tileID)
            stackS = 0
        End If
    Else ' Tibia 7.55+ : always click in ropespot at stack position zero
        tileID = GetTheLong(Matrix(relY, relX, z, idConnection).s(0).t1, Matrix(relY, relX, z, idConnection).s(0).t2)
        tileSTR = FiveChrLon(tileID)
        stackS = 0
        If shovelMode = True Then
            For ts = 1 To 10
                tileID = GetTheLong(Matrix(relY, relX, z, idConnection).s(ts).t1, Matrix(relY, relX, z, idConnection).s(ts).t2)
                If DatTiles(tileID).alwaysOnTop = True And DatTiles(tileID).multitype = False Then
                    stackS = ts
                    tileSTR = FiveChrLon(tileID)
                End If
            Next ts
        End If
    End If
    If ((frmHardcoreCheats.chkTotalWaste.Value = True) And (TibiaVersionLong >= 773) And (shouldBeVisible = False)) Then
        GoTo justdoit
    End If
    If mySlot(idConnection, SLOT_AMMUNITION).t1 = b1 And _
        mySlot(idConnection, SLOT_AMMUNITION).t2 = b2 Then 'use from ammo
    ' 11 00 83 FF FF 0A 00 00 7D 0B 00 EB 7C 99 7D 0C 80 01 00
    sCheat = "83 FF FF 0A 00 00 " & GoodHex(b1) & " " & GoodHex(b2) & " 00 " & GetHexStrFromPosition(X, y, z) & " " & tileSTR & " " & GoodHex(stackS)
     SafeCastCheatString "PerformUseMyItem1", idConnection, sCheat
  Else ' use from bp
    ' 11 00 83 FF FF 40 00 00 7D 0B 00 EB 7C 99 7D 0C 80 01 00
    fRes.foundcount = 0
    fRes = SearchItem(idConnection, b1, b2)
    If fRes.foundcount > 0 Then
      sCheat = "83 FF FF " & GoodHex(&H40 + fRes.bpID) & " 00 " & _
       GoodHex(fRes.slotID) & " " & GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(fRes.slotID) & " " & GetHexStrFromPosition(X, y, z) & " " & tileSTR & " " & GoodHex(stackS)
      SafeCastCheatString "PerformUseMyItem2", idConnection, sCheat
    Else
justdoit:
      If ((((frmHardcoreCheats.chkEnhancedCheats.Value = True) Or (frmHardcoreCheats.chkTotalWaste.Value = True)) And (TibiaVersionLong >= 773))) And (shouldBeVisible = False) Then
          ' NEW
         sCheat = "83 FF FF 00 00 00 " & GoodHex(b1) & " " & GoodHex(b2) & " 00 " & GetHexStrFromPosition(X, y, z) & " " & tileSTR & " " & GoodHex(stackS)

         SafeCastCheatString "PerformUseMyItem3", idConnection, sCheat
         PerformUseMyItem = 0
         Exit Function
      Else
      If publicDebugMode = True Then
        aRes = SendSystemMessageToClient(idConnection, "Required item was not found")
        DoEvents
      End If
      PerformUseMyItem = -1
      Exit Function
      End If
    End If
  End If
  PerformUseMyItem = 0
  Exit Function
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at PerformUseMyItem #"
  frmMain.DoCloseActions idConnection
  DoEvents
  PerformUseMyItem = -1
End Function



Public Function LootGoodItems(idConnection As Integer) As Long
  Dim res1 As TypeSearchItemResult2
  Dim res2 As TypeSearchItemResult2
  Dim aRes As Long
  Dim limitJ As Long
  Dim tileID As Long
  Dim sCheat As String
  Dim cPacket() As Byte
  Dim inRes As Integer
  Dim amount As Byte
  Dim subContTileID As Long
  Dim subPos As Byte
  Dim foodTileID As Long
  Dim foodPos As Byte
  Dim j As Long
  subContTileID = 0
  foodTileID = 0
  If Not (requestLootBp(idConnection) = &HFF) Then
  'there have been a request to loot items in a given container ID
  
    If Backpack(idConnection, requestLootBp(idConnection)).open = True Then
    ' the requested container ID is already open
      
      ' for faster speed:
      ' loot "good items" starting with the end of the container, finishing with his slot 0
      limitJ = (Backpack(idConnection, requestLootBp(idConnection)).used) - 1
      For j = limitJ To 0 Step -1
        tileID = GetTheLong(Backpack(idConnection, requestLootBp(idConnection)).item(j).t1, Backpack(idConnection, requestLootBp(idConnection)).item(j).t2)
        
        ' if the item in current slot is a subcontainer, then take note about it
        If DatTiles(tileID).iscontainer = True Then
          subContTileID = tileID
          subPos = CByte(j)
        End If
        
         ' if the item in current slot is food, then take note about it
        If DatTiles(tileID).isFood = True Then
          If IsGoodLoot(idConnection, tileID) = False Then
            foodTileID = tileID
            foodPos = CByte(j)
          End If
        End If
        
        ' if the item in current slot is "good loot", then loot it !
        If IsGoodLoot(idConnection, tileID) = True Then
          ' search where it will be moved
          res1 = SearchItemDestinationForLoot(idConnection, Backpack(idConnection, _
           requestLootBp(idConnection)).item(j).t1, Backpack(idConnection, _
           requestLootBp(idConnection)).item(j).t2, requestLootBp(idConnection))
          If res1.foundcount > 0 Then
            ' if any place was found in containers, move it there
            If Backpack(idConnection, requestLootBp(idConnection)).item(j).t3 = 0 Then
              amount = &H1
            Else
              amount = Backpack(idConnection, requestLootBp(idConnection)).item(j).t3
            End If
            sCheat = "78 FF FF " & GoodHex(&H40 + requestLootBp(idConnection)) & _
             " 00 " & GoodHex(CByte(j)) & " " & FiveChrLon(tileID) & " " & _
             GoodHex(CByte(j)) & " FF FF " & GoodHex(&H40 + res1.bpID) & " 00 " & GoodHex(res1.slotID) & _
             " " & GoodHex(amount)
            SafeCastCheatString "LootGoodItems1", idConnection, sCheat
            LootGoodItems = 0
            Exit Function
          Else
            ' if there are no free slots avaiable then stop the looting process
            GoTo dontLootMore
          End If
        End If
      Next j
      ' after looting "good loot" ...
      
      If subContTileID > 0 Then
        ' if a subcontainer was found, open it
        sCheat = "82 FF FF " & GoodHex(&H40 + requestLootBp(idConnection)) & " 00 " & GoodHex(subPos) & " " & FiveChrLon(subContTileID) & " " & GoodHex(subPos) & " " & GoodHex(requestLootBp(idConnection))
'        inRes = GetCheatPacket(cPacket, sCheat)
'        frmMain.UnifiedSendToServerGame idConnection, cPacket, True
'        DoEvents
        SafeCastCheatString "LootGoodItems2", idConnection, sCheat
        LootGoodItems = 0
        Exit Function
      ElseIf foodTileID > 0 Then
        ' if a food was found, eat it
        aRes = EatFood(idConnection, LowByteOfLong(foodTileID), HighByteOfLong(foodTileID), requestLootBp(idConnection), foodPos)
        DoEvents
      End If
dontLootMore:
      ' finish the whole looting process
      lootTimeExpire(idConnection) = 0 ' 5 seconds timer forced to expire
      ' close requested container (close corpse)
      sCheat = "87 " & GoodHex(requestLootBp(idConnection))
      SafeCastCheatString "LootGoodItems3", idConnection, sCheat
      LootGoodItems = 0
    End If
  End If
End Function

Public Function OpenCorpse(idConnection As Integer) As Long
  Dim aRes As Long
  Dim bpID As Byte
  Dim sCheat As String
  Dim cPacket() As Byte
  'aRes = SendLogSystemMessageToClient(idConnection, "Detected new corpse at " & _
   myLastCorpseX(idConnection) & " , " & myLastCorpseY(idConnection) & " , " & _
   myLastCorpseZ(idConnection) & " , " & myLastCorpseS(idConnection) & " : " & _
   FiveChrLon(myLastCorpseTileID(idConnection)))
 ' DoEvents
  '0A 00 82 37 7D A7 7D 09 59 0F 02 00
  bpID = frmBackpacks.GetFirstFreeBpID(idConnection)
  If (myLastCorpseZ(idConnection) <> myZ(idConnection)) Then
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Ignored looting corpse order because different floor :" & _
       CStr(myLastCorpseX(idConnection)) & "," & CStr(myLastCorpseY(idConnection)) & "," & CStr(myLastCorpseZ(idConnection)) & _
       " at bpID " & GoodHex(bpID))
      DoEvents
    End If
    OpenCorpse = 0
    Exit Function
  End If
  If Not (bpID = &HFF) Then
    requestLootBp(idConnection) = bpID
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Looting corpse at :" & _
       CStr(myLastCorpseX(idConnection)) & "," & CStr(myLastCorpseY(idConnection)) & "," & CStr(myLastCorpseZ(idConnection)) & _
       " at bpID " & GoodHex(bpID))
      DoEvents
    End If
    sCheat = "0A 00 82 " & FiveChrLon(myLastCorpseX(idConnection)) & " " & _
     FiveChrLon(myLastCorpseY(idConnection)) & " " & GoodHex(CByte(myLastCorpseZ(idConnection))) & _
     FiveChrLon(myLastCorpseTileID(idConnection)) & " " & GoodHex(CByte(myLastCorpseS(idConnection))) & " " & GoodHex(bpID) & " "
    GetCheatPacket cPacket, sCheat
    frmMain.UnifiedSendToServerGame idConnection, cPacket, True
    DoEvents
    ' give 5 seconds max to loot
    lootTimeExpire(idConnection) = GetTickCount() + 5000
  Else
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Can't loot it. Please open new containers!")
      DoEvents
    End If
  End If
  OpenCorpse = 0
End Function

Public Sub ExecuteBuffer(idConnection As Integer)
  Dim orders As String
  Dim startX As Long
  Dim startY As Long
  Dim X As Long
  Dim y As Long
  Dim z As Long
  Dim s As Long
  Dim aRes As Long
  Dim tileID As Long
  Dim myMap As TypeAstarMatrix
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim Px As Long
  Dim Py As Long
  Dim lOrders As Long
  Dim chCompare As String
  Dim inRes As Integer
  Dim continue As Boolean
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  If ReadyBuffer(idConnection) = True Then
  orders = RequiredMoveBuffer(idConnection)
  If orders = "" Then
    ' we are already in the goal
    Exit Sub
  ElseIf orders = "X" Then
    ' no direct path found
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Far far path didn't find any solution!")
      DoEvents
    End If
    RequiredMoveBuffer(idConnection) = ""
    Exit Sub
  Else
    sCheat = ""
    lOrders = Len(orders)
    For X = 1 To lOrders
      sCheat = sCheat & " 0" & Mid(orders, X, 1)
    Next X
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Done far far distance path successfully: " & sCheat)
      DoEvents
    End If
    sCheat = FiveChrLon(lOrders + 2) & " 64 " & GoodHex(CByte(lOrders)) & sCheat
    inRes = GetCheatPacket(cPacket, sCheat)
    frmMain.UnifiedSendToServerGame idConnection, cPacket, True
    DoEvents
    RequiredMoveBuffer(idConnection) = ""
     Exit Sub
  End If
  End If
  Exit Sub
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " Got connection lose while doing ExecuteBuffer"
  DoEvents
End Sub

Public Sub ReadTrueMap(idConnection As Integer, ByRef myMap As TypeAstarMatrix)
  Dim X As Long
  Dim y As Long
  Dim z As Long
  Dim s As Long
  Dim startX As Long
  Dim startY As Long
  Dim tileID As Long
  Dim continue As Boolean
  Dim aRes As Long
  Dim tmpID As Double
  Dim nameofgivenID As String
  startX = 0
  startY = 0
  ' delimiter our map by a wall
  For X = -9 To 10
    myMap.cost(X, -7) = CostBlock
    myMap.cost(X, 8) = CostBlock
  Next X
  For y = -7 To 8
    myMap.cost(-9, y) = CostBlock
    myMap.cost(10, y) = CostBlock
  Next y
  ' export truemap into astar map inside zone
  z = myZ(idConnection)
  For y = -6 To 7
    For X = -8 To 9
      continue = True
      tileID = GetTheLong(Matrix(y, X, z, idConnection).s(0).t1, Matrix(y, X, z, idConnection).s(0).t2)
      If tileID = 0 Then
        myMap.cost(X, y) = CostBlock
        continue = False
      ElseIf DatTiles(tileID).blocking = True Then
        myMap.cost(X, y) = CostBlock
        continue = False
      ElseIf (DatTiles(tileID).floorChangeDOWN = True) Or (DatTiles(tileID).floorChangeUP = True) Then
        If (DatTiles(tileID).requireShovel = False) And (DatTiles(tileID).requireRope = False) And (DatTiles(tileID).requireRightClick = False) Then
          myMap.cost(X, y) = CostBlock
               If (myMap.cost(X - 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y)) < CostNearHandicap Then
                 myMap.cost(X - 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y + 1)) < CostNearHandicap Then
                 myMap.cost(X, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y)) < CostNearHandicap Then
                 myMap.cost(X + 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y - 1)) < CostNearHandicap Then
                 myMap.cost(X, y - 1) = CostNearHandicap
               End If
          
          
          continue = False
        End If
      End If
      If continue = True Then
        If (myMap.cost(X, y)) < CostWalkable Then
          myMap.cost(X, y) = CostWalkable
        End If
        For s = 1 To 10
          tileID = GetTheLong(Matrix(y, X, z, idConnection).s(s).t1, Matrix(y, X, z, idConnection).s(s).t2)
          If tileID = 0 Then
            Exit For
          ElseIf tileID = 97 Then 'person
            tmpID = Matrix(y, X, z, idConnection).s(s).dblID
            nameofgivenID = GetNameFromID(idConnection, tmpID)
            If nameofgivenID = "" Then
              ' detected mobile with no name! ?
              aRes = -1
            ElseIf (nameofgivenID = CharacterName(idConnection)) Or (tmpID = lastAttackedID(idConnection)) Then
              ' myself
              aRes = -2
            Else
              myMap.cost(X, y) = CostBlock
              Exit For
            End If
          ElseIf DatTiles(tileID).isField Then
            If (myMap.cost(X, y)) < CostHandicap Then
              myMap.cost(X, y) = CostHandicap
            End If
               If (myMap.cost(X - 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y)) < CostNearHandicap Then
                 myMap.cost(X - 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y + 1)) < CostNearHandicap Then
                 myMap.cost(X, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y)) < CostNearHandicap Then
                 myMap.cost(X + 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y - 1)) < CostNearHandicap Then
                 myMap.cost(X, y - 1) = CostNearHandicap
               End If
          ElseIf DatTiles(tileID).blocking Then
            myMap.cost(X, y) = CostBlock
            Exit For
          ElseIf (DatTiles(tileID).floorChangeDOWN = True) Or (DatTiles(tileID).floorChangeUP = True) Then
            If (DatTiles(tileID).requireShovel = False) And (DatTiles(tileID).requireRope = False) And (DatTiles(tileID).requireRightClick = False) Then
               myMap.cost(X, y) = CostBlock
               If (myMap.cost(X - 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y)) < CostNearHandicap Then
                 myMap.cost(X - 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y + 1)) < CostNearHandicap Then
                 myMap.cost(X, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y)) < CostNearHandicap Then
                 myMap.cost(X + 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y - 1)) < CostNearHandicap Then
                 myMap.cost(X, y - 1) = CostNearHandicap
               End If
               
              Exit For
            End If
          End If
        Next s
      End If
    Next X
  Next y
  'force start to be walkable
  myMap.cost(startX, startY) = CostWalkable
End Sub

Public Function FindBestPath(ByVal idConnection As Integer, ByVal goalX As Long, ByVal goalY As Long, ByVal showPath As Boolean) As Long
  Dim orders As String
  Dim X As Long
  Dim y As Long
  Dim z As Long
  Dim s As Long
  Dim aRes As Long
  Dim tileID As Long
  Dim myMap As TypeAstarMatrix
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim Px As Long
  Dim Py As Long
  Dim lOrders As Long
  Dim chCompare As String
  Dim inRes As Integer
  Dim continue As Boolean
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  ChaotizeXYrel idConnection, goalX, goalY, myZ(idConnection)
    
  ReadTrueMap idConnection, myMap
  'force goal to be walkable
  myMap.cost(goalX, goalY) = CostWalkable
  
  orders = Astar(0, 0, goalX, goalY, myMap)
  If orders = "" Then
    ' we are already in the goal
    FindBestPath = 0
  ElseIf orders = "X" Then
    ' no direct path found
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] (1) Can't find direct short path to relative " & CStr(goalX) & "," & CStr(goalY))
      DoEvents
    End If
    FindBestPath = -1
  Else
    sCheat = ""
    lOrders = Len(orders)
    For X = 1 To lOrders
      sCheat = sCheat & " 0" & Mid(orders, X, 1)
    Next X
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Doing short distance path to " & goalX & "," & goalY & " : " & sCheat)
      DoEvents
    End If
    sCheat = FiveChrLon(lOrders + 2) & " 64 " & GoodHex(CByte(lOrders)) & sCheat
    inRes = GetCheatPacket(cPacket, sCheat)
    frmMain.UnifiedSendToServerGame idConnection, cPacket, True
    
    
    If showPath = True Then
      frmTrueMap.gridMap.Redraw = False
      Px = 8 'start draw path in 0,0 of our gridmap
      Py = 6
      For X = 1 To lOrders
        chCompare = Mid(orders, X, 1)
        frmTrueMap.gridMap.Col = Px
        frmTrueMap.gridMap.Row = Py
        frmTrueMap.gridMap.CellBackColor = &HFFFFC0
        Select Case chCompare
        Case StrMoveNorth
          Py = Py - 1
        Case StrMoveRight
          Px = Px + 1
        Case StrMoveSouth
          Py = Py + 1
        Case StrMoveLeft
          Px = Px - 1
        Case StrMoveNorthRight
          Px = Px + 1
          Py = Py - 1
        Case StrMoveSouthRight
          Px = Px + 1
          Py = Py + 1
        Case StrMoveSouthLeft
          Px = Px - 1
          Py = Py + 1
        Case StrMoveNorthLeft
          Px = Px - 1
          Py = Py - 1
        End Select
      Next X
      frmTrueMap.gridMap.Col = Px
      frmTrueMap.gridMap.Row = Py
      frmTrueMap.gridMap.CellBackColor = ColourPath
      frmTrueMap.gridMap.Redraw = True
    End If
    DoEvents
    If ((showPath = True) Or (publicDebugMode = True)) And (lOrders > 0) Then
      '07 00 83 57 7E E6 7D 07 0A
      Px = myX(idConnection) 'start draw path in our ingame position
      Py = myY(idConnection)
      sCheat = ""
      For X = 1 To lOrders
        chCompare = Mid(orders, X, 1)
        Select Case chCompare
        Case StrMoveNorth
          Py = Py - 1
        Case StrMoveRight
          Px = Px + 1
        Case StrMoveSouth
          Py = Py + 1
        Case StrMoveLeft
          Px = Px - 1
        Case StrMoveNorthRight
          Px = Px + 1
          Py = Py - 1
        Case StrMoveSouthRight
          Px = Px + 1
          Py = Py + 1
        Case StrMoveSouthLeft
          Px = Px - 1
          Py = Py + 1
        Case StrMoveNorthLeft
          Px = Px - 1
          Py = Py - 1
        End Select
        sCheat = sCheat & " 83 " & FiveChrLon(Px) & " " & FiveChrLon(Py) & " " & GoodHex(CByte(myZ(idConnection))) & " 0A"
      Next X
      sCheat = FiveChrLon(lOrders * 7) & sCheat
      'aRes = SendLogSystemMessageToClient(idConnection, "Drawing path to " & goalX & "," & goalY & " : " & sCheat)
      inRes = GetCheatPacket(cPacket, sCheat)
      frmMain.UnifiedSendToClientGame idConnection, cPacket
      DoEvents
    End If
    
    FindBestPath = 0
  End If
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " Got connection lose while doing FindBestPath"
  DoEvents
  FindBestPath = -1
  Exit Function
End Function

Public Function FindBestPathV2(ByVal idConnection As Integer, ByVal goalX As Long, ByVal goalY As Long, ByVal showPath As Boolean) As Long
  Dim orders As String
  Dim X As Long
  Dim y As Long
  Dim z As Long
  Dim s As Long
  Dim aRes As Long
  Dim tileID As Long
  Dim myMap As TypeAstarMatrix
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim Px As Long
  Dim Py As Long
  Dim lOrders As Long
  Dim chCompare As String
  Dim inRes As Integer
  Dim continue As Boolean
  Dim backupX As Long
  Dim backupY As Long
  Dim optimized As Boolean
  Dim tryingX As Long
  Dim tryingY As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  ChaotizeXYrel idConnection, goalX, goalY, myZ(idConnection)
  
  optimized = False
  ReadTrueMap idConnection, myMap
  'search a better goal
  backupX = goalX
  backupY = goalY
  
  tryingX = goalX
  tryingY = goalY
  If (tryingX > -9) And (tryingX < 10) And (tryingY > -7) And (tryingY < 8) Then
    If myMap.cost(tryingX, tryingY) = CostWalkable Then
      optimized = False
      GoTo continue
    End If
  End If
  tryingX = goalX - 1
  tryingY = goalY - 1
  If (tryingX > -9) And (tryingX < 10) And (tryingY > -7) And (tryingY < 8) Then
    If myMap.cost(tryingX, tryingY) = CostWalkable Then
      optimized = True
      GoTo continue
    End If
  End If
  tryingX = goalX - 1
  tryingY = goalY
  If (tryingX > -9) And (tryingX < 10) And (tryingY > -7) And (tryingY < 8) Then
    If myMap.cost(tryingX, tryingY) = CostWalkable Then
      optimized = True
      GoTo continue
    End If
  End If
  tryingX = goalX - 1
  tryingY = goalY + 1
  If (tryingX > -9) And (tryingX < 10) And (tryingY > -7) And (tryingY < 8) Then
    If myMap.cost(tryingX, tryingY) = CostWalkable Then
      optimized = True
      GoTo continue
    End If
  End If
  tryingX = goalX
  tryingY = goalY + 1
  If (tryingX > -9) And (tryingX < 10) And (tryingY > -7) And (tryingY < 8) Then
    If myMap.cost(tryingX, tryingY) = CostWalkable Then
      optimized = True
      GoTo continue
    End If
  End If
  tryingX = goalX + 1
  tryingY = goalY + 1
  If (tryingX > -9) And (tryingX < 10) And (tryingY > -7) And (tryingY < 8) Then
    If myMap.cost(tryingX, tryingY) = CostWalkable Then
      optimized = True
      GoTo continue
    End If
  End If
  tryingX = goalX + 1
  tryingY = goalY
  If (tryingX > -9) And (tryingX < 10) And (tryingY > -7) And (tryingY < 8) Then
    If myMap.cost(tryingX, tryingY) = CostWalkable Then
      optimized = True
      GoTo continue
    End If
  End If
  tryingX = goalX + 1
  tryingY = goalY - 1
  If (tryingX > -9) And (tryingX < 10) And (tryingY > -7) And (tryingY < 8) Then
    If myMap.cost(tryingX, tryingY) = CostWalkable Then
      optimized = True
      GoTo continue
    End If
  End If
  tryingX = goalX
  tryingY = goalY - 1
  If (tryingX > -9) And (tryingX < 10) And (tryingY > -7) And (tryingY < 8) Then
    If myMap.cost(tryingX, tryingY) = CostWalkable Then
      optimized = True
      GoTo continue
    End If
  End If
  tryingX = goalX
  tryingY = goalY
continue:
  goalX = tryingX
  goalY = tryingY
  myMap.cost(goalX, goalY) = CostWalkable
  If optimized = True Then
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Destination got a slight change to avoid danger : modified from [" & backupX & "," & backupY & "] to [" & goalX & "," & goalY & "]")
      DoEvents
    End If
  End If
  orders = Astar(0, 0, goalX, goalY, myMap)
  If orders = "" Then
    ' we are already in the goal
    FindBestPathV2 = 0
    Exit Function
  ElseIf orders = "X" Then
    ' no direct path found
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] (2) Can't find direct short path to relative " & CStr(goalX) & "," & CStr(goalY))
      DoEvents
    End If
    FindBestPathV2 = -1
    Exit Function
  Else
    sCheat = ""
    lOrders = Len(orders)
    For X = 1 To lOrders
      sCheat = sCheat & " 0" & Mid(orders, X, 1)
    Next X
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Doing short distance path to " & goalX & "," & goalY & " : " & sCheat)
      DoEvents
    End If
    sCheat = FiveChrLon(lOrders + 2) & " 64 " & GoodHex(CByte(lOrders)) & sCheat
    inRes = GetCheatPacket(cPacket, sCheat)
    frmMain.UnifiedSendToServerGame idConnection, cPacket, True
    
    
    If showPath = True Then
      frmTrueMap.gridMap.Redraw = False
      Px = 8 'start draw path in 0,0 of our gridmap
      Py = 6
      For X = 1 To lOrders
        chCompare = Mid(orders, X, 1)
        frmTrueMap.gridMap.Col = Px
        frmTrueMap.gridMap.Row = Py
        frmTrueMap.gridMap.CellBackColor = &HFFFFC0
        Select Case chCompare
        Case StrMoveNorth
          Py = Py - 1
        Case StrMoveRight
          Px = Px + 1
        Case StrMoveSouth
          Py = Py + 1
        Case StrMoveLeft
          Px = Px - 1
        Case StrMoveNorthRight
          Px = Px + 1
          Py = Py - 1
        Case StrMoveSouthRight
          Px = Px + 1
          Py = Py + 1
        Case StrMoveSouthLeft
          Px = Px - 1
          Py = Py + 1
        Case StrMoveNorthLeft
          Px = Px - 1
          Py = Py - 1
        End Select
      Next X
      frmTrueMap.gridMap.Col = Px
      frmTrueMap.gridMap.Row = Py
      frmTrueMap.gridMap.CellBackColor = ColourPath
      frmTrueMap.gridMap.Redraw = True
    End If
    DoEvents ' <- ok in any case
    If ((showPath = True) Or (publicDebugMode = True)) And (lOrders > 0) Then
      '07 00 83 57 7E E6 7D 07 0A
      Px = myX(idConnection) 'start draw path in our ingame position
      Py = myY(idConnection)
      sCheat = ""
      For X = 1 To lOrders
        chCompare = Mid(orders, X, 1)
        Select Case chCompare
        Case StrMoveNorth
          Py = Py - 1
        Case StrMoveRight
          Px = Px + 1
        Case StrMoveSouth
          Py = Py + 1
        Case StrMoveLeft
          Px = Px - 1
        Case StrMoveNorthRight
          Px = Px + 1
          Py = Py - 1
        Case StrMoveSouthRight
          Px = Px + 1
          Py = Py + 1
        Case StrMoveSouthLeft
          Px = Px - 1
          Py = Py + 1
        Case StrMoveNorthLeft
          Px = Px - 1
          Py = Py - 1
        End Select
        sCheat = sCheat & " 83 " & FiveChrLon(Px) & " " & FiveChrLon(Py) & " " & GoodHex(CByte(myZ(idConnection))) & " 0A"
      Next X
      sCheat = FiveChrLon(lOrders * 7) & sCheat
      'aRes = SendLogSystemMessageToClient(idConnection, "Drawing path to " & goalX & "," & goalY & " : " & sCheat)
      inRes = GetCheatPacket(cPacket, sCheat)
      frmMain.UnifiedSendToClientGame idConnection, cPacket
      DoEvents
    End If
    
    FindBestPathV2 = 0
    Exit Function
  End If
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " Got connection lose while doing FindBestPathV2"
  DoEvents
  FindBestPathV2 = -1
  Exit Function
End Function

Public Function ExistsPath(idConnection As Integer, goalX As Long, goalY As Long) As Boolean
  Dim orders As String
  Dim startX As Long
  Dim startY As Long
  Dim X As Long
  Dim y As Long
  Dim z As Long
  Dim s As Long
  Dim aRes As Long
  Dim tileID As Long
  Dim myMap As TypeAstarMatrix
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim Px As Long
  Dim Py As Long
  Dim lOrders As Long
  Dim chCompare As String
  Dim inRes As Integer
  Dim continue As Boolean
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  startX = 0
  startY = 0
  ' delimiter our map by a wall
  For X = -9 To 10
    myMap.cost(X, -7) = CostBlock
    myMap.cost(X, 8) = CostBlock
  Next X
  For y = -7 To 8
    myMap.cost(-9, y) = CostBlock
    myMap.cost(10, y) = CostBlock
  Next y
  ' export truemap into astar map inside zone
  z = myZ(idConnection)
  For y = -6 To 7
    For X = -8 To 9
      continue = True
      tileID = GetTheLong(Matrix(y, X, z, idConnection).s(0).t1, Matrix(y, X, z, idConnection).s(0).t2)
      If tileID = 0 Then
        myMap.cost(X, y) = CostBlock
        continue = False
      ElseIf DatTiles(tileID).blocking = True Then
        myMap.cost(X, y) = CostBlock
        continue = False
      ElseIf (DatTiles(tileID).floorChangeDOWN = True) Or (DatTiles(tileID).floorChangeUP = True) Then
        If (DatTiles(tileID).requireShovel = False) And (DatTiles(tileID).requireRope = False) And (DatTiles(tileID).requireRightClick = False) Then
          myMap.cost(X, y) = CostBlock
               If (myMap.cost(X - 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y)) < CostNearHandicap Then
                 myMap.cost(X - 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y + 1)) < CostNearHandicap Then
                 myMap.cost(X, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y)) < CostNearHandicap Then
                 myMap.cost(X + 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y - 1)) < CostNearHandicap Then
                 myMap.cost(X, y - 1) = CostNearHandicap
               End If
          
          
          continue = False
        End If
      End If
      If continue = True Then
        If (myMap.cost(X, y)) < CostWalkable Then
          myMap.cost(X, y) = CostWalkable
        End If
        For s = 1 To 10
          tileID = GetTheLong(Matrix(y, X, z, idConnection).s(s).t1, Matrix(y, X, z, idConnection).s(s).t2)
          If tileID = 0 Then
            Exit For
          ElseIf DatTiles(tileID).isField Then
            myMap.cost(X, y) = CostHandicap
               If (myMap.cost(X - 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y)) < CostNearHandicap Then
                 myMap.cost(X - 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y + 1)) < CostNearHandicap Then
                 myMap.cost(X, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y)) < CostNearHandicap Then
                 myMap.cost(X + 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y - 1)) < CostNearHandicap Then
                 myMap.cost(X, y - 1) = CostNearHandicap
               End If
          ElseIf DatTiles(tileID).blocking Then
            myMap.cost(X, y) = CostBlock
            Exit For
          ElseIf (DatTiles(tileID).floorChangeDOWN = True) Or (DatTiles(tileID).floorChangeUP = True) Then
            If (DatTiles(tileID).requireShovel = False) And (DatTiles(tileID).requireRope = False) And (DatTiles(tileID).requireRightClick = False) Then
               myMap.cost(X, y) = CostBlock
               If (myMap.cost(X - 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y)) < CostNearHandicap Then
                 myMap.cost(X - 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y + 1)) < CostNearHandicap Then
                 myMap.cost(X, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y)) < CostNearHandicap Then
                 myMap.cost(X + 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y - 1)) < CostNearHandicap Then
                 myMap.cost(X, y - 1) = CostNearHandicap
               End If
               
              Exit For
            End If
          End If
        Next s
      End If
    Next X
  Next y
  'force goal to be walkable
  myMap.cost(goalX, goalY) = CostWalkable
  'force start to be walkable
  myMap.cost(startX, startY) = CostWalkable
  
  orders = Astar(startX, startY, goalX, goalY, myMap)
  If orders = "" Then
    ' we are already in the goal
    ExistsPath = True
    Exit Function
  ElseIf orders = "X" Then
    ' no direct path found
    ExistsPath = False
    Exit Function
  Else
    ExistsPath = True
    Exit Function
  End If
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " Got connection lose while doing ExistsPath"
  DoEvents
  ExistsPath = -1
  Exit Function
End Function

Public Function FindAnyLongPath(idConnection As Integer, goalX As Long, goalY As Long) As Long
  Dim orders As String
  Dim startX As Long
  Dim startY As Long
  Dim X As Long
  Dim y As Long
  Dim z As Long
  Dim s As Long
  Dim aRes As Long
  Dim tileID As Long
  Dim myMap As TypeAstarMatrix
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim Px As Long
  Dim Py As Long
  Dim lOrders As Long
  Dim chCompare As String
  Dim inRes As Integer
  Dim continue As Boolean
  Dim bestGoal As TypePathResult
  Dim difx As Long
  Dim dify As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  startX = 0
  startY = 0
  ' delimiter our map by a wall
  For X = -9 To 10
    myMap.cost(X, -7) = CostBlock
    myMap.cost(X, 8) = CostBlock
  Next X
  For y = -7 To 8
    myMap.cost(-9, y) = CostBlock
    myMap.cost(10, y) = CostBlock
  Next y
  ' export truemap into astar map inside zone
  z = myZ(idConnection)
  For y = -6 To 7
    For X = -8 To 9
      continue = True
      tileID = GetTheLong(Matrix(y, X, z, idConnection).s(0).t1, Matrix(y, X, z, idConnection).s(0).t2)
      If tileID = 0 Then
        myMap.cost(X, y) = CostBlock
        continue = False
      ElseIf DatTiles(tileID).blocking = True Then
        myMap.cost(X, y) = CostBlock
        continue = False
      ElseIf (DatTiles(tileID).floorChangeDOWN = True) Or (DatTiles(tileID).floorChangeUP = True) Then
        If (DatTiles(tileID).requireShovel = False) And (DatTiles(tileID).requireRope = False) And (DatTiles(tileID).requireRightClick = False) Then
          myMap.cost(X, y) = CostBlock
               If (myMap.cost(X - 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y)) < CostNearHandicap Then
                 myMap.cost(X - 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y + 1)) < CostNearHandicap Then
                 myMap.cost(X, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y)) < CostNearHandicap Then
                 myMap.cost(X + 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y - 1)) < CostNearHandicap Then
                 myMap.cost(X, y - 1) = CostNearHandicap
               End If
          
          
          continue = False
        End If
      End If
      If continue = True Then
        If (myMap.cost(X, y)) < CostWalkable Then
          myMap.cost(X, y) = CostWalkable
        End If
        For s = 1 To 10
          tileID = GetTheLong(Matrix(y, X, z, idConnection).s(s).t1, Matrix(y, X, z, idConnection).s(s).t2)
          If tileID = 0 Then
            Exit For
          ElseIf DatTiles(tileID).isField Then
            myMap.cost(X, y) = CostHandicap
               If (myMap.cost(X - 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y)) < CostNearHandicap Then
                 myMap.cost(X - 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y + 1)) < CostNearHandicap Then
                 myMap.cost(X, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y)) < CostNearHandicap Then
                 myMap.cost(X + 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y - 1)) < CostNearHandicap Then
                 myMap.cost(X, y - 1) = CostNearHandicap
               End If
          ElseIf DatTiles(tileID).blocking Then
            myMap.cost(X, y) = CostBlock
            Exit For
          ElseIf (DatTiles(tileID).floorChangeDOWN = True) Or (DatTiles(tileID).floorChangeUP = True) Then
            If (DatTiles(tileID).requireShovel = False) And (DatTiles(tileID).requireRope = False) And (DatTiles(tileID).requireRightClick = False) Then
               myMap.cost(X, y) = CostBlock
               If (myMap.cost(X - 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y)) < CostNearHandicap Then
                 myMap.cost(X - 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y + 1)) < CostNearHandicap Then
                 myMap.cost(X, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y)) < CostNearHandicap Then
                 myMap.cost(X + 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y - 1)) < CostNearHandicap Then
                 myMap.cost(X, y - 1) = CostNearHandicap
               End If
               
              Exit For
            End If
          End If
        Next s
      End If
    Next X
  Next y
  'force start to be walkable
  myMap.cost(startX, startY) = CostWalkable
  ' find best local goal
  bestGoal = AstarGiveBestGoal(startX, startY, goalX, goalY, myMap)
  difx = Abs(bestGoal.X)
  dify = Abs(bestGoal.y)
  If (difx < 2) And (dify < 2) Then
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Can't find alternative long way path")
      DoEvents
    End If
    FindAnyLongPath = -1
    Exit Function
  Else
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Trying to find an alternative long distance path to " & goalX & "," & goalY & " .Translated to local move to " & bestGoal.X & "," & bestGoal.y)
      DoEvents
    End If
  End If
  orders = Astar(startX, startY, bestGoal.X, bestGoal.y, myMap)
  If orders = "" Then
    ' we are already in the goal
    FindAnyLongPath = -1
  ElseIf orders = "X" Then
    ' no direct path found
    FindAnyLongPath = -1
  Else
    sCheat = ""
    lOrders = Len(orders)
    For X = 1 To lOrders
      sCheat = sCheat & " 0" & Mid(orders, X, 1)
    Next X

    sCheat = FiveChrLon(lOrders + 2) & " 64 " & GoodHex(CByte(lOrders)) & sCheat
    inRes = GetCheatPacket(cPacket, sCheat)
    frmMain.UnifiedSendToServerGame idConnection, cPacket, True
    DoEvents
    FindAnyLongPath = 0
  End If
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " Got connection lose while doing FindAnyLongPath"
  DoEvents
  FindAnyLongPath = -1
  Exit Function
End Function

Public Function GetNearestDepot(idConnection As Integer) As Long
  Dim orders As String
  Dim startX As Long
  Dim startY As Long
  Dim X As Long
  Dim y As Long
  Dim z As Long
  Dim s As Byte
  Dim aRes As Long
  Dim tileID As Long
  Dim myMap As TypeAstarMatrix
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim Px As Long
  Dim Py As Long
  Dim lOrders As Long
  Dim chCompare As String
  Dim inRes As Integer
  Dim continue As Boolean
  Dim bestX As Long
  Dim bestY As Long
  Dim bestZ As Byte
  Dim bestS As Byte
  Dim tmpDist As Long
  Dim bestTileID As Long
  Dim bestDist As Long
  Dim sMinusOne As Byte
  Dim PrevTileID As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  startX = 0
  startY = 0
  ' delimiter our map by a wall
  For X = -9 To 10
    myMap.cost(X, -7) = CostBlock
    myMap.cost(X, 8) = CostBlock
  Next X
  For y = -7 To 8
    myMap.cost(-9, y) = CostBlock
    myMap.cost(10, y) = CostBlock
  Next y
  ' export truemap into astar map inside zone
  z = myZ(idConnection)
  For y = -6 To 7
    For X = -8 To 9
      continue = True
      tileID = GetTheLong(Matrix(y, X, z, idConnection).s(0).t1, Matrix(y, X, z, idConnection).s(0).t2)
      If tileID = 0 Then
        myMap.cost(X, y) = CostBlock
        continue = False
      ElseIf DatTiles(tileID).blocking = True Then
        myMap.cost(X, y) = CostBlock
        continue = False
      ElseIf (DatTiles(tileID).floorChangeDOWN = True) Or (DatTiles(tileID).floorChangeUP = True) Then
        If (DatTiles(tileID).requireShovel = False) And (DatTiles(tileID).requireRope = False) And (DatTiles(tileID).requireRightClick = False) Then
          myMap.cost(X, y) = CostBlock
               If (myMap.cost(X - 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y)) < CostNearHandicap Then
                 myMap.cost(X - 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y + 1)) < CostNearHandicap Then
                 myMap.cost(X, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y)) < CostNearHandicap Then
                 myMap.cost(X + 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y - 1)) < CostNearHandicap Then
                 myMap.cost(X, y - 1) = CostNearHandicap
               End If
          
          
          continue = False
        End If
      End If
      If continue = True Then
        If (myMap.cost(X, y)) < CostWalkable Then
          myMap.cost(X, y) = CostWalkable
        End If
        For s = 1 To 10
          tileID = GetTheLong(Matrix(y, X, z, idConnection).s(s).t1, Matrix(y, X, z, idConnection).s(s).t2)
          If tileID = 0 Then
            Exit For
          ElseIf DatTiles(tileID).isField Then
            myMap.cost(X, y) = CostHandicap
               If (myMap.cost(X - 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y)) < CostNearHandicap Then
                 myMap.cost(X - 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y + 1)) < CostNearHandicap Then
                 myMap.cost(X, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y)) < CostNearHandicap Then
                 myMap.cost(X + 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y - 1)) < CostNearHandicap Then
                 myMap.cost(X, y - 1) = CostNearHandicap
               End If
          ElseIf DatTiles(tileID).blocking Then
            myMap.cost(X, y) = CostBlock
            Exit For
          ElseIf (DatTiles(tileID).floorChangeDOWN = True) Or (DatTiles(tileID).floorChangeUP = True) Then
            If (DatTiles(tileID).requireShovel = False) And (DatTiles(tileID).requireRope = False) And (DatTiles(tileID).requireRightClick = False) Then
               myMap.cost(X, y) = CostBlock
               If (myMap.cost(X - 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y)) < CostNearHandicap Then
                 myMap.cost(X - 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X - 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X - 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y + 1)) < CostNearHandicap Then
                 myMap.cost(X, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y + 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y + 1) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y)) < CostNearHandicap Then
                 myMap.cost(X + 1, y) = CostNearHandicap
               End If
               If (myMap.cost(X + 1, y - 1)) < CostNearHandicap Then
                 myMap.cost(X + 1, y - 1) = CostNearHandicap
               End If
               If (myMap.cost(X, y - 1)) < CostNearHandicap Then
                 myMap.cost(X, y - 1) = CostNearHandicap
               End If
               
              Exit For
            End If
          End If
        Next s
      End If
    Next X
  Next y
  myMap.cost(0, 0) = CostWalkable
  bestDist = 10000
  For y = -6 To 7
    For X = -8 To 9
      For s = 1 To 10
        tileID = GetTheLong(Matrix(y, X, z, idConnection).s(s).t1, Matrix(y, X, z, idConnection).s(s).t2)
        If DatTiles(tileID).isDepot = True Then
          If (myMap.cost(X - 1, y - 1) = CostWalkable) Or _
             (myMap.cost(X - 1, y) = CostWalkable) Or _
             (myMap.cost(X - 1, y + 1) = CostWalkable) Or _
             (myMap.cost(X, y + 1) = CostWalkable) Or _
             (myMap.cost(X + 1, y + 1) = CostWalkable) Or _
             (myMap.cost(X + 1, y) = CostWalkable) Or _
             (myMap.cost(X + 1, y - 1) = CostWalkable) Or _
             (myMap.cost(X, y - 1) = CostWalkable) Then
             sMinusOne = s - 1
             PrevTileID = GetTheLong(Matrix(y, X, z, idConnection).s(sMinusOne).t1, Matrix(y, X, z, idConnection).s(sMinusOne).t2)
             ' don't try depots with something over them
             If DatTiles(PrevTileID).blocking = True Then
               tmpDist = ManhattanDistance(X, y, 0, 0)
               If tmpDist < bestDist Then
                 If ExistsPath(idConnection, X, y) = True Then
                   bestX = X
                   bestY = y
                   bestS = s
                   bestTileID = tileID
                   bestDist = tmpDist
                 End If
               End If
             End If
          End If
        ElseIf tileID = 0 Then
          Exit For
        End If
      Next s
    Next X
  Next y
  If tmpDist = 10000 Then
    aRes = SendLogSystemMessageToClient(idConnection, "Could not find any depot avaiable! Resuming with next command ...")
    DoEvents
    GetNearestDepot = -1
    Exit Function
  Else
    depotX(idConnection) = myX(idConnection) + bestX
    depotY(idConnection) = myY(idConnection) + bestY
    depotZ(idConnection) = z
    depotS(idConnection) = bestS
    depotTileID(idConnection) = bestTileID
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Found nearest depot at : x=" & depotX(idConnection) & " ; y=" & depotY(idConnection) & " ; z=" & z & " ; s=" & GoodHex(s))
      DoEvents
    End If
    GetNearestDepot = 0
    Exit Function
  End If
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " Got connection lose while doing GetNearestDepot"
  DoEvents
  GetNearestDepot = -1
  Exit Function
End Function
Public Sub OpenTheDepot(idConnection As Integer)
  Dim aRes As Long
  Dim bpID As Byte
  Dim sCheat As String
  Dim cPacket() As Byte
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  '0A 00 82 37 7D A7 7D 09 59 0F 02 00
  bpID = frmBackpacks.GetFirstFreeBpID(idConnection)
  If Not (bpID = &HFF) Then
    requestLootBp(idConnection) = bpID
    sCheat = "0A 00 82 " & FiveChrLon(depotX(idConnection)) & " " & _
     FiveChrLon(depotY(idConnection)) & " " & GoodHex(CByte(depotZ(idConnection))) & _
     FiveChrLon(depotTileID(idConnection)) & " " & GoodHex(CByte(depotS(idConnection))) & " " & GoodHex(bpID) & " "
    GetCheatPacket cPacket, sCheat
    frmMain.UnifiedSendToServerGame idConnection, cPacket, True
    DoEvents
  End If
  Exit Sub
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " Got connection lose while doing OpenTheDepot"
End Sub

Public Sub OpenDepotChest(idConnection As Integer)
  Dim res1 As TypeSearchItemResult2
  Dim inRes As Integer
  Dim sCheat As String
  Dim cPacket() As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  'open subcontainer
  '0A 00 82 FF FF 40 00 05 E7 0A 05 00
  b1 = LowByteOfLong(tileID_depotChest)
  b2 = HighByteOfLong(tileID_depotChest)
  timeToRetryOpenDepot(idConnection) = GetTickCount() + 5000
  res1 = SearchItem(idConnection, b1, b2)
  If res1.foundcount > 0 Then
    sCheat = "82 FF FF " & GoodHex(&H40 + lastDepotBPID(idConnection)) & " 00 " & GoodHex(res1.slotID) & " " & FiveChrLon(tileID_depotChest) & " " & GoodHex(res1.slotID) & " " & GoodHex(lastDepotBPID(idConnection))
    SafeCastCheatString "OpenDepotChest1", idConnection, sCheat
  End If
End Sub
Public Function DropLoot(idConnection As Integer) As Long
'  Dim goodItems
  Dim currItem As Long
  Dim res1 As TypeSearchItemResult2
  Dim res2 As TypeSearchItemResult2
  Dim res3 As TypeSearchItemResult2
  Dim i As Long
  Dim lim As Long
  Dim debugStr As String
  Dim aRes As Long
  Dim sourceB1 As Byte
  Dim sourceB2 As Byte
  Dim sourceTileID As Long
  Dim foundAsource As Boolean
  Dim foundAdestination As Boolean
  Dim tileID As Long
  Dim sCheat As String
  Dim cPacket() As Byte
  Dim inRes As Integer
  Dim amount As Byte
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  DropDelayerTurn(idConnection) = DropDelayerTurn(idConnection) + 1
  If (DropDelayerTurn(idConnection) < DropDelayerConst) Then
    'wait a bit
    DropLoot = 0
    Exit Function
  End If
  DropDelayerTurn(idConnection) = 0
   
  If GetTickCount() > nextForcedDepotDeployRetry(idConnection) Then
    somethingChangedInBps(idConnection) = True 'force retry
  End If
  If somethingChangedInBps(idConnection) = False Then
    DropLoot = 0
    Exit Function
  End If
  foundAsource = False
  
  'goodItems = cavebotGoodLoot(idConnection).Keys
  'lim = cavebotGoodLoot(idConnection).count - 1
  'find source
'  For i = 0 To lim
'    currItem = CLng(goodItems(i))
'    sourceB1 = LowByteOfLong(currItem)
'    sourceB2 = HighByteOfLong(currItem)
'    res1 = SearchItemWithBPException(idConnection, sourceB1, sourceB2, lastDepotBPID(idConnection))
'    If res1.foundCount > 0 Then
'      foundAsource = True
'      sourceTileID = currItem
'      Exit For
'    End If
'  Next
  
    'NEW since Blackd Proxy 23.8
    res1 = SearchItemWithBPExceptionGoodLoot(idConnection, lastDepotBPID(idConnection))
    If (res1.foundcount > 0) Then
        foundAsource = True
        sourceB1 = res1.b1
        sourceB2 = res1.b2
        sourceTileID = GetTheLong(sourceB1, sourceB2)
        currItem = sourceTileID
    End If
  
  If (foundAsource = False) Then
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Depot deploy completed successfully")
      DoEvents
    End If
    DropLoot = -1
    Exit Function
  End If
  
  'find destination
  foundAdestination = False
  res2 = SearchItemDestinationInDepot(idConnection, sourceB1, sourceB2, lastDepotBPID(idConnection))
  If res2.foundcount > 0 Then
    ' do unit drop
    foundAdestination = True
    nextForcedDepotDeployRetry(idConnection) = GetTickCount() + randomNumberBetween(4000, 6000)
    somethingChangedInBps(idConnection) = False
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Dropping x" & CLng(res1.amount) & " items type " & _
       FiveChrLon(sourceTileID) & " from bpID " & GoodHex(res1.bpID) & " slotID " & GoodHex(res1.slotID) & _
       " ; destination: bpID " & GoodHex(res2.bpID) & " slotID " & GoodHex(res2.slotID))
      DoEvents
    End If
    
                '0F 00 78 FF FF 41 00 00 BC 0D 00 FF FF 40 00 05 06
            If res1.amount = 0 Then
              amount = &H1
            Else
              amount = res1.amount
            End If
            sCheat = "78 FF FF " & GoodHex(&H40 + res1.bpID) & _
             " 00 " & GoodHex(res1.slotID) & " " & FiveChrLon(currItem) & " " & _
             GoodHex(res1.slotID) & " FF FF " & GoodHex(&H40 + res2.bpID) & " 00 " & GoodHex(res2.slotID) & _
             " " & GoodHex(amount)
            SafeCastCheatString "DropLoot1", idConnection, sCheat
            
    DropLoot = 0
    Exit Function
  Else
   ', if no free slot found , open container,
    res3 = SearchSubContainer(idConnection, sourceB1, sourceB2, lastDepotBPID(idConnection))
    If res3.foundcount > 0 Then
      foundAdestination = True
      nextForcedDepotDeployRetry(idConnection) = GetTickCount() + randomNumberBetween(4000, 6000)
      somethingChangedInBps(idConnection) = False
      If publicDebugMode = True Then
        aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Opening subcontainer in depot at bpID " & GoodHex(res3.bpID) & " slotID " & GoodHex(res3.slotID))
        DoEvents
      End If
      'open subcontainer
      '0A 00 82 FF FF 40 00 05 E7 0A 05 00
      tileID = GetTheLong(res3.b1, res3.b2)
      sCheat = "82 FF FF " & GoodHex(&H40 + lastDepotBPID(idConnection)) & " 00 " & GoodHex(res3.slotID) & " " & FiveChrLon(tileID) & " " & GoodHex(res3.slotID) & " " & GoodHex(lastDepotBPID(idConnection))
      SafeCastCheatString "DropLoot2", idConnection, sCheat
      DropLoot = 0
      Exit Function
    Else
    ' if no containers found -> force end
      foundAdestination = False
      aRes = SendLogSystemMessageToClient(idConnection, "No free space found in your depot!")
      DoEvents
      DropLoot = -1
      Exit Function
    End If
  End If
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " Got connection lose while doing DropLoot"
  DropLoot = -1
  Exit Function
End Function

Public Function DropLootToGround(idConnection As Integer) As Long
  'Dim goodItems
  Dim currItem As Long
  Dim res1 As TypeSearchItemResult2
  Dim res2 As TypeSearchItemResult2
  Dim res3 As TypeSearchItemResult2
  Dim i As Long
  Dim lim As Long
  Dim debugStr As String
  Dim aRes As Long
  Dim sourceB1 As Byte
  Dim sourceB2 As Byte
  Dim sourceTileID As Long
  Dim foundAsource As Boolean
  Dim foundAdestination As Boolean
  Dim tileID As Long
  Dim sCheat As String
  Dim cPacket() As Byte
  Dim inRes As Integer
  Dim amount As Byte
  Dim b3 As Byte
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  DropDelayerTurn(idConnection) = DropDelayerTurn(idConnection) + 1
  If (DropDelayerTurn(idConnection) < DropDelayerConst) Then
    'wait a bit
    DropLootToGround = 0
    Exit Function
  End If
  DropDelayerTurn(idConnection) = 0
  
  If GetTickCount() > nextForcedDepotDeployRetry(idConnection) Then
    somethingChangedInBps(idConnection) = True 'force retry
  End If
  If somethingChangedInBps(idConnection) = False Then
    DropLootToGround = 0
    Exit Function
  End If
  foundAsource = False
  
  'goodItems = cavebotGoodLoot(idConnection).Keys
  'lim = cavebotGoodLoot(idConnection).count - 1
  'find source
'  For i = 0 To lim
'    currItem = CLng(goodItems(i))
'    sourceB1 = LowByteOfLong(currItem)
'    sourceB2 = HighByteOfLong(currItem)
'    res1 = SearchItem(idConnection, sourceB1, sourceB2)
'    If res1.foundCount > 0 Then
'      foundAsource = True
'      sourceTileID = currItem
'      Exit For
'    End If
'  Next

    ' new since blackd proxy 23.8
    res1 = SearchItemGoodLoot(idConnection)
    If res1.foundcount > 0 Then
      foundAsource = True
      sourceB1 = res1.b1
      sourceB2 = res1.b2
      sourceTileID = GetTheLong(res1.b1, res1.b2)
      currItem = sourceTileID
    End If



  
  If (foundAsource = False) Then
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Ground deploy completed successfully")
      DoEvents
    End If
    DropLootToGround = -1
    Exit Function
  End If
 '0F 00 78 FF FF 41 00 06 9D 0B 06 4F 7F A6 7C 0B 02

  nextForcedDepotDeployRetry(idConnection) = GetTickCount() + randomNumberBetween(4000, 6000)
  somethingChangedInBps(idConnection) = False

  tileID = GetTheLong(sourceB1, sourceB2)
  
  b3 = res1.amount
  If (TibiaVersionLong >= 860) Or (TibiaVersionLong = 760) Then ' fix by divinity76
    If (b3 = &H0) Then
        b3 = &H1
    End If
  End If
  sCheat = "78 FF FF " & GoodHex(&H40 + res1.bpID) & " 00 " & GoodHex(res1.slotID) & " " & FiveChrLon(tileID) & " " & GoodHex(res1.slotID) & " " & FiveChrLon(myX(idConnection)) & " " & FiveChrLon(myY(idConnection)) & " " & GoodHex(CByte(myZ(idConnection))) & " " & GoodHex(b3)
  If publicDebugMode = True Then
    aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Dropping 1 item to ground: " & sCheat)
    DoEvents
  End If
  SafeCastCheatString "DropLootToGround1", idConnection, sCheat
  DropLootToGround = 0
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " Got connection lose while doing DropLoot"
  DropLootToGround = -1
  Exit Function
End Function
Public Function DropItemOnGroundBytes(ByVal idConnection As Integer, ByVal b1 As Byte, ByVal b2 As Byte, maxamount As Byte) As Long
    On Error GoTo goterr
    Dim foundAsource As Boolean
    Dim res1 As TypeSearchItemResult2
    Dim i As Long
    Dim sCheat As String
    Dim b3 As Byte
    Dim inRes As Integer
    Dim aRes As Long
    Dim tileID As Long
    Dim cPacket() As Byte
    foundAsource = False
    res1 = SearchItem(idConnection, b1, b2)
    If res1.foundcount <= 0 Then
        If publicDebugMode = True Then
          aRes = SendLogSystemMessageToClient(idConnection, "[Debug] DropItemOnGround failed (item not found in opened backpacks)")
          DoEvents
        End If
        DropItemOnGroundBytes = -1
        Exit Function
    End If
    b3 = res1.amount
    If (b3 > maxamount) Then
        b3 = maxamount
    End If
    If TibiaVersionLong >= 860 Then
      If (b3 = &H0) Then
          b3 = &H1
      End If
    End If
    tileID = GetTheLong(b1, b2)
    sCheat = "78 FF FF " & GoodHex(&H40 + res1.bpID) & " 00 " & GoodHex(res1.slotID) & " " & FiveChrLon(tileID) & " " & GoodHex(res1.slotID) & " " & FiveChrLon(myX(idConnection)) & " " & FiveChrLon(myY(idConnection)) & " " & GoodHex(CByte(myZ(idConnection))) & " " & GoodHex(b3)
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] DropItemOnGround OK")
      DoEvents
    End If
    SafeCastCheatString "DropItemOnGroundBytes1", idConnection, sCheat
    DropItemOnGroundBytes = 0
    Exit Function
goterr:
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] DropItemOnGround failed (code " & CStr(Err.Number) & ": " & Err.Description & ")")
      DoEvents
    End If
    DropItemOnGroundBytes = -1
    Exit Function
End Function
Public Function DropItemOnGround(ByVal idConnection As Integer, ByVal strTile As String) As Long
    On Error GoTo goterr
    Dim b1 As Byte
    Dim b2 As Byte
    Dim b3 As Byte
    Dim aRes As Long
    strTile = Trim$(strTile)
    If Len(strTile) = 8 Then
       b1 = CByte("&H" & Left$(strTile, 2))
       b2 = CByte("&H" & Mid(strTile, 4, 2))
       b3 = CByte("&H" & Right$(strTile, 2))
       DropItemOnGround = DropItemOnGroundBytes(idConnection, b1, b2, b3)
       Exit Function
    End If
goterr:
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] DropItemOnGround failed (bad format)")
      DoEvents
    End If
    DropItemOnGround = -1
    Exit Function
End Function

Public Function PrintDictionary(idConnection As Integer) As Long
  Dim goodItems
  Dim currItem As Double
'  Dim res1 As TypeSearchItemResult2
'  Dim res2 As TypeSearchItemResult2
'  Dim res3 As TypeSearchItemResult2
  Dim i As Long
  Dim lim As Long
  Dim debugStr As String
  Dim aRes As Long
  Dim sourceB1 As Byte
  Dim sourceB2 As Byte
  Dim sourceTileID As Long
  Dim foundAsource As Boolean
  Dim foundAdestination As Boolean
  Dim tileID As Long
  Dim sCheat As String
  Dim cPacket() As Byte
  Dim inRes As Integer
  Dim amount As Byte
  Dim showStr As String
  Dim showStr2 As String
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  aRes = SendLogSystemMessageToClient(idConnection, "Listing all creatures on memory:")
  DoEvents
  goodItems = NameOfID(idConnection).Keys
  lim = NameOfID(idConnection).count - 1
  'find source
  foundAsource = False
  For i = 0 To lim
    currItem = CLng(goodItems(i))
    showStr = SpaceID(currItem)
    showStr2 = GetNameFromID(idConnection, currItem)
    aRes = SendLogSystemMessageToClient(idConnection, "[" & CStr(i) & "] " & showStr2 & " (" & showStr & ") : " & GetHPFromID(idConnection, currItem) & "% hp")
    DoEvents
  Next
  PrintDictionary = 0
  Exit Function
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at PrintDictionary #"
  frmMain.DoCloseActions idConnection
  DoEvents
  PrintDictionary = -1
End Function

Public Sub RepositionScript(idConnection As Integer, firstLine As Long, lastLine As Long)
' This will reposition the script in the nearest point of current position after
' following a monster to dead
  Dim evLine As Long
  Dim currLineNumber As Long
  Dim currLine As String
  Dim lenCurrLine As Long
  Dim bestLine As Long
  Dim bestDist As Long
  Dim tmpDist As Long
  Dim pos As Long
  Dim mainCommand As String
  Dim param1 As String
  Dim param2 As String
  Dim param3 As String
  Dim val1 As Long
  Dim val2 As Long
  Dim val3 As Long
  Dim aRes As Long
  If firstLine >= cavebotLenght(idConnection) Then
    firstLine = exeLine(idConnection)
  End If
  If lastLine >= cavebotLenght(idConnection) Then
    lastLine = cavebotLenght(idConnection) - 1
  End If
  bestLine = firstLine
  bestDist = 10000

  For evLine = firstLine To lastLine
    currLineNumber = evLine
    currLine = GetStringFromIDLine(idConnection, currLineNumber)
    lenCurrLine = Len(currLine)
    pos = 1
    mainCommand = ParseString(currLine, pos, lenCurrLine, " ")
    SkipBlanks currLine, pos, lenCurrLine
    If mainCommand = "move" Then
      param1 = ParseString(currLine, pos, lenCurrLine, ",")
      val1 = CLng(param1)
      SkipBlanks currLine, pos, lenCurrLine
      param2 = ParseString(currLine, pos, lenCurrLine, ",")
      val2 = CLng(param2)
      SkipBlanks currLine, pos, lenCurrLine
      param3 = ParseString(currLine, pos, lenCurrLine, ",")
      val3 = CLng(param3)
      If val3 = myZ(idConnection) Then
        tmpDist = ManhattanDistance(myX(idConnection), myY(idConnection), val1, val2)
        If evLine = exeLine(idConnection) Then
          tmpDist = tmpDist - 1 ' give a small priority to keep in current script line
        End If
        If tmpDist < bestDist Then ' this point is closer to current position
          bestLine = evLine
          bestDist = tmpDist
        End If
      Else
        ' floor change: not valid jump
      End If
    Else
      Exit For ' not a move command: can't reposition the script. it would give unpredictable results
    End If
  Next evLine
  ' at the end of the for we have the nearest script line to our current position
  ' and we will jump there
  If exeLine(idConnection) <> bestLine Then
      If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Optimized execution : jump from " & exeLine(idConnection) & " to " & bestLine)
      DoEvents
    End If
    'exeLine(idConnection) = bestLine
    updateExeLine idConnection, bestLine, False
  End If
End Sub

Public Sub RepositionScriptAtTrap(idConnection As Integer, Optional withallpoints As Boolean = False)
' This will reposition the script in the nearest point of current position after
' following a monster to dead
  Dim evLine As Long
  Dim currLineNumber As Long
  Dim currLine As String
  Dim lenCurrLine As Long
  Dim bestLine As Long
  Dim bestDist As Long
  Dim tmpDist As Long
  Dim pos As Long
  Dim mainCommand As String
  Dim param1 As String
  Dim param2 As String
  Dim param3 As String
  Dim val1 As Long
  Dim val2 As Long
  Dim val3 As Long
  Dim aRes As Long
  Dim firstLine As Long
  Dim lastLine As Long
  Dim incX As Long
  Dim incY As Long
  Dim maxSteps As Long
  Dim Sid As Integer
  Dim am As Long
  Dim roundstage As Long

  ' changed in 9.6

  Sid = idConnection
  firstLine = 0
  lastLine = cavebotLenght(idConnection) - 1
  maxSteps = lastLine
  bestLine = exeLine(idConnection)
  bestDist = 10000
  evLine = 0
  roundstage = 1
anotherRound:
  maxSteps = lastLine
  #If withdebugreposition = True Then
  LogOnFile "what.txt", vbCrLf & "REPOSITION START. STAGE " & CStr(roundstage) & "From line " & bestLine
  #End If
  
  Do
    If (evLine > lastLine) Then
        Exit Do
    End If
    maxSteps = maxSteps - 1
 'For evLine = actualLine To lastLine
    currLineNumber = evLine
    currLine = GetStringFromIDLine(idConnection, currLineNumber)
    currLine = parseVars(Sid, currLine)
    #If withdebugreposition = True Then
    LogOnFile "what.txt", CStr(currLineNumber) & ":" & currLine
     #End If
    lenCurrLine = Len(currLine)
    pos = 1
    mainCommand = LCase(ParseString(currLine, pos, lenCurrLine, " "))
    SkipBlanks currLine, pos, lenCurrLine
    
    Select Case mainCommand
 
    Case "move"
        param1 = ParseString(currLine, pos, lenCurrLine, ",")
        val1 = CLng(param1)
        SkipBlanks currLine, pos, lenCurrLine
        param2 = ParseString(currLine, pos, lenCurrLine, ",")
        val2 = CLng(param2)
        SkipBlanks currLine, pos, lenCurrLine
        param3 = ParseString(currLine, pos, lenCurrLine, ",")
        val3 = CLng(param3)
        If val3 = myZ(idConnection) Then
          tmpDist = ManhattanDistance(myX(idConnection), myY(idConnection), val1, val2)
          If tmpDist < bestDist Then ' this point is closer to current position
            incX = myX(idConnection) - val1
            incY = myY(idConnection) - val2
            If (Abs(incX) < 8) And (Abs(incY) < 7) Then
              If ExistsPath(idConnection, incX, incY) Then
                bestLine = evLine
                bestDist = tmpDist
              End If
            Else
              bestLine = evLine
              bestDist = tmpDist
            End If
          End If
        Else
          ' floor change: not valid jump
        End If
         evLine = evLine + 1
    Case "gotoscriptline"
    If withallpoints = True Then
             evLine = evLine + 1
    Else
     param1 = ParseString(currLine, pos, lenCurrLine, ",")
     val1 = CLng(param1)
     evLine = val1
    End If
  Case "iftrue"
      If withallpoints = True Then
             evLine = evLine + 1
    Else
    aRes = ProcessCondition(Sid, currLine, pos, lenCurrLine, True)
    If aRes = -1 Then
        evLine = evLine + 1
    Else
        evLine = aRes
    End If
    End If
  Case "ifenoughitemsgoto"
        If withallpoints = True Then
             evLine = evLine + 1
    Else
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    'val1 = GetTheLongFromFiveChr(param1) ' not needed since in 9.38
    SkipBlanks currLine, pos, lenCurrLine
    param2 = ParseString(currLine, pos, lenCurrLine, ",")
    val2 = CLng(param2)
    SkipBlanks currLine, pos, lenCurrLine
    param3 = ParseString(currLine, pos, lenCurrLine, ",")
    val3 = CLng(param3)
     ' completed
    am = CountTheItemsForUser(Sid, param1) ' changed since 9.38
    If am >= val2 Then
      evLine = val3
    Else
      evLine = evLine + 1
    End If
    End If
  Case "iffewitemsgoto"
        If withallpoints = True Then
             evLine = evLine + 1
    Else
    param1 = ParseString(currLine, pos, lenCurrLine, ",")
    'val1 = GetTheLongFromFiveChr(param1)  not needed since in 9.38
    SkipBlanks currLine, pos, lenCurrLine
    param2 = ParseString(currLine, pos, lenCurrLine, ",")
    val2 = CLng(param2) ' value to be compared
    SkipBlanks currLine, pos, lenCurrLine
    param3 = ParseString(currLine, pos, lenCurrLine, ",")
    val3 = CLng(param3) ' line where it should jump
    'ammount of items with given tileID
    am = CountTheItemsForUser(Sid, param1) ' changed since 9.38
    ' compare now
    If am >= val2 Then ' false : continue with next line of the script
      evLine = evLine + 1

    Else ' true : jump to given line
      evLine = val3

    End If
    
    End If
    Case Else
              evLine = evLine + 1
    End Select
    If (evLine = 0) Then
        Exit Do
    End If
  Loop Until maxSteps <= 0
  If ((roundstage = 1) And (withallpoints = False)) Then
    evLine = exeLine(idConnection) ' consider possible moves from current point
    roundstage = 2
    GoTo anotherRound
  End If
  ' at the end of the for we have the nearest script line to our current position
  ' and we will jump there
  
    #If withdebugreposition = True Then
        If exeLine(idConnection) = bestLine Then
            LogOnFile "what.txt", "****************" & vbCrLf & "KEEP ON CURRENT LINE :" & bestLine & vbCrLf & "****************"
        End If
     #End If
     
  If exeLine(idConnection) <> bestLine Then
    If publicDebugMode = True Then
      aRes = SendLogSystemMessageToClient(idConnection, "[Debug] Semi traped : trying to solve with jump from " & exeLine(idConnection) & " to " & bestLine)
      DoEvents
    End If
    #If withdebugreposition = True Then
    LogOnFile "what.txt", "****************" & "Took decision of changing line to :" & bestLine & vbCrLf & "****************"
     #End If
    'exeLine(idConnection) = bestLine
    updateExeLine idConnection, bestLine, False
  End If

End Sub

Public Sub AddCavebotMove()
  If cavebotIDselected > 0 Then
    frmCavebot.AddScriptLine "move " & myX(cavebotIDselected) & "," & myY(cavebotIDselected) & "," & myZ(cavebotIDselected)
  End If
End Sub

Public Sub AddCavebotMovePoint(idConnection As Integer, X As Long, y As Long, z As Long)
  If cavebotIDselected > 0 Then
    frmCavebot.AddScriptLine "move " & X & "," & y & "," & z
  End If
End Sub

Public Function DistBetweenMeAndID(idConnection As Integer, Sid As Double) As Long
  Dim X As Integer
  Dim y As Integer
  Dim z As Integer
  Dim s As Integer
  Dim tileID As Integer
  Dim currentID As Double
  Dim nameofgivenID As String
  z = myZ(idConnection)
  For X = -8 To 9
    For y = -6 To 7
        For s = 1 To 10
          tileID = GetTheLong(Matrix(y, X, z, idConnection).s(s).t1, Matrix(y, X, z, idConnection).s(s).t2)
          If tileID = 97 Then
           currentID = Matrix(y, X, z, idConnection).s(s).dblID
            If currentID = Sid Then
                If Abs(X) > Abs(y) Then
                    DistBetweenMeAndID = Abs(X)
                Else
                    DistBetweenMeAndID = Abs(y)
                End If
                Exit Function
            End If
          ElseIf tileID = 0 Then
            Exit For
          End If
        Next s
    Next y
  Next X
  DistBetweenMeAndID = 1000
End Function

Public Sub DoUnifiedClickMove(idConnection As Integer, ByVal Px As Long, ByVal Py As Long, ByVal Pz As Long)
  Dim myBpos As Long
  Dim b1 As Byte
  Dim b2 As Byte
  Dim X As Long
  Dim y As Long
  Dim z As Long

  Dim aRes As Long
  Dim Sid As Long
  Dim pid As Long
  Sid = idConnection
  pid = ProcessID(Sid)
  X = Px
  y = Py
  z = Pz
  If z <> myZ(idConnection) Then
    Exit Sub ' Stop because that would be an illegal move
  End If
  If ((onDepotPhase(idConnection) = 2) Or (onDepotPhase(idConnection) = 6)) Then
    b1 = 0 ' do nothing
  Else
    ChaotizeXY idConnection, X, y, z
  End If
  If X < 0 Then
    Exit Sub ' Stop because that would be an illegal move
  End If
  If y < 0 Then
    Exit Sub ' Stop because that would be an illegal move
  End If
  If z < 0 Then
    Exit Sub ' Stop because that would be an illegal move
  End If
  
  
  myBpos = MyBattleListPosition(Sid)
  b1 = LowByteOfLong(X)
  b2 = HighByteOfLong(X)
  Memory_WriteByte adrXgo, b1, pid
  Memory_WriteByte adrXgo + 1, b2, pid
  b1 = LowByteOfLong(y)
  b2 = HighByteOfLong(y)
  Memory_WriteByte adrYgo, b1, pid
  Memory_WriteByte adrYgo + 1, b2, pid
  b1 = CByte(z)
  Memory_WriteByte adrZgo, b1, pid
  Memory_WriteByte adrGo + (myBpos * CharDist), 1, pid
End Sub

Public Sub ChaotizeXYrel(ByVal idConnection As Integer, ByRef Px As Long, ByRef Py As Long, ByVal Pz As Long)
    Dim X As Long
    Dim y As Long
    X = myX(idConnection) + Px
    y = myY(idConnection) + Py
    ChaotizeXY idConnection, X, y, Pz
    Px = X - myX(idConnection)
    Py = y - myY(idConnection)
End Sub
Public Sub ChaotizeXY(ByVal idConnection As Integer, ByRef Px As Long, ByRef Py As Long, ByVal Pz As Long)
  Dim tries As Long
  Dim res As TypeBMSquare
  Dim X As Long
  Dim y As Long
  Dim z As Long
  Dim aRes As Long
  X = Px
  y = Py
  z = Pz
  tries = 0
  If cavebotEnabled(idConnection) = True Then
    If CavebotChaoticMode(idConnection) = 1 Then
       res.color = 1
       res.walkable = False
       ' randomize, but avoid no walkable points
       While ((res.color <> &H0) And (res.walkable = False) And (tries < 4))
          X = Px + randomNumberBetween(-1, 1)
          y = Py + randomNumberBetween(-1, 1)
          tries = tries + 1
          GetBigMapSquare res, X, y, z
       Wend
       If tries = 4 Then ' use original waypoint
          X = Px
          y = Py
       End If
        If publicDebugMode = True Then
          aRes = SendLogSystemMessageToClient(idConnection, "Doing chaotic move. Original waypoint=" & CStr(Px) & "," & CStr(Py) & "," & CStr(Pz) & " ; Final= " & CStr(X) & "," & CStr(y) & "," & CStr(z))
          DoEvents
        End If
    End If
  End If
  Px = X
  Py = y
End Sub

Private Function MaxV(ByVal dFirst As Long, ByVal dSecond As Long) As Long
    If dFirst > dSecond Then
        MaxV = dFirst
    Else
        MaxV = dSecond
    End If
End Function

Public Function CheckSETUSEITEM(idConnection As Integer) As Boolean
    
    Dim gtc As Long
    Dim gtclimit As Long
    Dim posiblePoints() As TypePosibleSETUSEITEM
    Dim totalPoints As Long
    Dim b1 As Byte
    Dim b2 As Byte
    Dim tileSTR As String
    Dim useSTR As String
    Dim realX As Long
    Dim realY As Long
    Dim currentDist As Long
    Dim chosenPoint As Long
    Dim useb1 As Byte
    Dim useb2 As Byte
    Dim X As Long
    Dim y As Long
    Dim z As Long
    Dim s As Byte
    Dim aRes As Long
    z = myZ(idConnection)
    chosenPoint = 0
    For currentDist = 1 To maxSETUSEITEMDist
        totalPoints = 0
        ReDim posiblePoints(0)
        posiblePoints(0).lngX = 0
        posiblePoints(0).lngY = 0
        posiblePoints(0).byteS = 0
        posiblePoints(0).tileb1 = 0
        posiblePoints(0).tileb2 = 0
        posiblePoints(0).strItem = ""
        For X = -currentDist To currentDist
          For y = -currentDist To currentDist
            If MaxV(Abs(X), Abs(y)) = currentDist Then
              
              For s = 1 To 10
                b1 = Matrix(y, X, z, idConnection).s(s).t1
                b2 = Matrix(y, X, z, idConnection).s(s).t2
                tileSTR = GoodHex(b1) & " " & GoodHex(b2)
                useSTR = getSETUSEITEM(idConnection, tileSTR)
                If (Not (useSTR = "")) Then
                    'Debug.Print "Found usable target at distance " & CStr(currentDist) & ": " & CStr(X) & "," & CStr(Y)
                    realX = myX(idConnection) + X
                    realY = myY(idConnection) + y
                    If (Not ((realX = SETUSEITEM_lastX(idConnection)) And (realY = SETUSEITEM_lastY(idConnection)))) Then
                        totalPoints = totalPoints + 1
                        ReDim Preserve posiblePoints(totalPoints)
                         posiblePoints(totalPoints).lngX = realX
                         posiblePoints(totalPoints).lngY = realY
                         posiblePoints(totalPoints).byteS = s
                         posiblePoints(totalPoints).tileb1 = b1
                         posiblePoints(totalPoints).tileb2 = b2
                         posiblePoints(totalPoints).strItem = useSTR
                    End If
                End If
              Next s
            End If
          Next y
        Next X
        If totalPoints > 0 Then
            chosenPoint = randomNumberBetween(1, totalPoints)
            Exit For
        End If
    Next currentDist
    If chosenPoint = 0 Then
        CheckSETUSEITEM = False
        Exit Function
    End If
    
    If GameConnected(idConnection) = False Then
        CheckSETUSEITEM = False
        Exit Function
    End If
 

    ClientExecutingLongCommand(idConnection) = True
    
    SETUSEITEM_lastX(idConnection) = posiblePoints(chosenPoint).lngX
    SETUSEITEM_lastY(idConnection) = posiblePoints(chosenPoint).lngY
    
    ' use item
    useb1 = CByte("&H" & Left$(posiblePoints(chosenPoint).strItem, 2))
    useb2 = CByte("&H" & Right$(posiblePoints(chosenPoint).strItem, 2))
    'Debug.Print "Using item " & GoodHex(useb1) & " " & GoodHex(useb2) & " at " & CStr(posiblePoints(chosenPoint).lngX) & "," & CStr(posiblePoints(chosenPoint).lngY) & "," & CStr(Z) & " ..."
    aRes = UseItemHere(idConnection, useb1, useb2, posiblePoints(chosenPoint).lngX, posiblePoints(chosenPoint).lngY, z, posiblePoints(chosenPoint).byteS)
    
    gtc = GetTickCount()
    gtclimit = gtc + randomNumberBetween(1000, 2000)
    Do
        gtc = GetTickCount()
        DoEvents
        DoEvents
    Loop Until (gtc >= gtclimit)
    ClientExecutingLongCommand(idConnection) = False
    CheckSETUSEITEM = True
End Function
