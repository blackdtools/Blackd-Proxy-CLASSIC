Attribute VB_Name = "modWarbot"
#Const FinalMode = 1
Option Explicit

Public Const MAX_NAME_LENGHT = 30


      
Public NameDist As Long
Public OutfitDist As Long
Public allowRename As Boolean
Public lastClient As Long
Public lastValid(0 To 5) As String
'tray icon

Public OutfitOfName(0 To 5) As Scripting.Dictionary
Public OutfitOfChar(0 To 5) As Scripting.Dictionary

Public GLOBAL_FRIENDSLOWLIMIT_HP As Long
Public GLOBAL_MYSAFELIMIT_HP As Long
Public GLOBAL_AUTOFRIENDHEAL_MODE As Long
Public Type TypeMagebomb
  CharacterName As String
  AttackMode As String
  TargetToShot As String
  RetryTime As Long
  LoginVersion As Long
  LogFileName As String
  IPstring As String
  PORTnumber As Long
  connectionStatus As Long
  ConnectionTimeout As Long
  nextSendLogin As Long
  Key(15) As Byte
  loginPacket() As Byte
  attackPacket() As Byte
End Type
Public Magebombs() As TypeMagebomb
Public MagebombsLoaded As Long
Public MagebombLeader As Integer
Public MagebombStage As Long
Public SafeModeOutPacket(5) As Byte


Public FIRSTCONNECTIONTIMEOUT_ms As Long
Public SECONDCONNECTIONTIMEOUT_ms As Long

Public DebugingMagebomb As Boolean
Public MagebombStartTime As Long

Public EnemyList As Scripting.Dictionary



Public Sub GetOutfit(pid As Long)
  Dim aRes As Long
  Dim myBpos As Long
  Dim myID As Long
  Dim b As Byte
  Dim bPos As Long
  Dim tmpID As Long
  On Error GoTo gotErr
  myID = Memory_ReadLong(adrNum, pid)
  myBpos = -1
  For bPos = 0 To LAST_BATTLELISTPOS
    tmpID = Memory_ReadLong(adrNChar + (bPos * CharDist), pid)
    If tmpID = myID Then
      myBpos = bPos
      Exit For
    End If
  Next bPos
  If myBpos = -1 Then
    Exit Sub
  End If
  'read outfit from memory
  b = Memory_ReadByte(adrNChar + OutfitDist + (myBpos * CharDist), pid)
  frmWarbot.txtOutfit(0).Text = CStr(CLng(b))
  lastValid(0) = CStr(CLng(b))
  b = Memory_ReadByte(adrNChar + OutfitDist + 4 + (myBpos * CharDist), pid)
  frmWarbot.txtOutfit(1).Text = CStr(CLng(b))
  lastValid(1) = CStr(CLng(b))
  b = Memory_ReadByte(adrNChar + OutfitDist + 8 + (myBpos * CharDist), pid)
  frmWarbot.txtOutfit(2).Text = CStr(CLng(b))
  lastValid(2) = CStr(CLng(b))
  b = Memory_ReadByte(adrNChar + OutfitDist + 12 + (myBpos * CharDist), pid)
  frmWarbot.txtOutfit(3).Text = CStr(CLng(b))
  lastValid(3) = CStr(CLng(b))
  b = Memory_ReadByte(adrNChar + OutfitDist + 16 + (myBpos * CharDist), pid)
  frmWarbot.txtOutfit(4).Text = CStr(CLng(b))
  lastValid(4) = CStr(CLng(b))
  If TibiaVersionLong >= 773 Then
    b = Memory_ReadByte(adrNChar + OutfitDist + 20 + (myBpos * CharDist), pid)
    frmWarbot.txtOutfit(4).Text = CStr(CLng(b))
    lastValid(5) = CStr(CLng(b))
  Else
    lastValid(5) = "0"
  End If
  Exit Sub
gotErr:
End Sub

Public Sub ProcessAllBattleLists()
  Dim tibiaclient As Long
  'Dim hWndDesktop As Long
  Dim i As Integer
  Dim j As Integer
  Dim num As Integer
  Dim num2 As Integer
  Dim posM As Integer
  Dim writeStr As String
  Dim writeChr As String
  Dim myID As Long
  Dim tmpID As Long
  Dim bPos As Long
  Dim lastPos As Long
  'hWndDesktop = GetDesktopWindow()
  tibiaclient = 0
  Do
    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      Exit Do
    Else
      ProcessBattleList tibiaclient
    End If
  Loop
End Sub

Public Sub ProcessBattleList(tibiaclient As Long)
 Dim aRes As Long
  Dim myBpos As Long
  Dim myID As Long
  Dim b As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim b4 As Byte
  Dim b5 As Byte
  Dim b6 As Byte
  Dim n As Byte
  Dim thename As String
  Dim bPos As Long
  Dim tmpID As Long
  Dim i As Long
  Dim tmpPlace As Long
  'On Error GoTo gotErr



  For bPos = 0 To LAST_BATTLELISTPOS
  'read char name
    tmpPlace = (adrNChar + (bPos * CharDist) + NameDist)
    n = Memory_ReadByte((adrNChar + (bPos * CharDist) + NameDist), tibiaclient)
    If n > 0 Then
      thename = ""
      i = 0
      While ((n > &H0) And (i < 40))
         thename = thename & Chr(n)
         i = i + 1
         n = Memory_ReadByte((adrNChar + (bPos * CharDist) + NameDist + i), tibiaclient)
      Wend
      If CharIsListed(thename) = True Then
        b1 = GetOutfitByteFromChar(0, thename)
        b2 = GetOutfitByteFromChar(1, thename)
        b3 = GetOutfitByteFromChar(2, thename)
        b4 = GetOutfitByteFromChar(3, thename)
        b5 = GetOutfitByteFromChar(4, thename)
        b6 = GetOutfitByteFromChar(5, thename)
        'modify outfit
        tmpPlace = adrNChar + OutfitDist + (bPos * CharDist)
        Memory_WriteByte tmpPlace, b1, tibiaclient
        tmpPlace = tmpPlace + 4
        Memory_WriteByte tmpPlace, b2, tibiaclient
        tmpPlace = tmpPlace + 4
        Memory_WriteByte tmpPlace, b3, tibiaclient
        tmpPlace = tmpPlace + 4
        Memory_WriteByte tmpPlace, b4, tibiaclient
        tmpPlace = tmpPlace + 4
        Memory_WriteByte tmpPlace, b5, tibiaclient
        If TibiaVersionLong >= 773 Then
         tmpPlace = tmpPlace + 4
         Memory_WriteByte tmpPlace, b6, tibiaclient
        End If
      End If
    End If
  Next bPos
  Exit Sub
gotErr:
End Sub

Public Sub LoadFile(thename As String)
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  
  Dim fso As Scripting.FileSystemObject
  Dim fn As Integer
  Dim strLine As String
  Dim Filename As String
  Dim i As Long
  Dim seguir As Boolean
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim b4 As Byte
  Dim b5 As Byte
  Dim b6 As Byte
  Set fso = New Scripting.FileSystemObject
    Filename = App.Path & "\wargroups\" & Left$(thename, Len(thename) - 3) & "out"
    If fso.FileExists(Filename) = True Then
      fn = FreeFile
      Open Filename For Input As #fn
        Line Input #fn, strLine
        b1 = CByte(CLng(strLine))
        If CLng(b1) < firstValidOutfit Then
          b1 = CByte(firstValidOutfit)
        End If
        Line Input #fn, strLine
        b2 = CByte(CLng(strLine))
        Line Input #fn, strLine
        b3 = CByte(CLng(strLine))
        Line Input #fn, strLine
        b4 = CByte(CLng(strLine))
        Line Input #fn, strLine
        b5 = CByte(CLng(strLine))
        If EOF(fn) Then
          b6 = &H0
        Else
          Line Input #fn, strLine
          b6 = CByte(CLng(strLine))
        End If
      Close #fn
    Else
        b1 = CByte(firstValidOutfit)
        b2 = 0
        b3 = 0
        b4 = 0
        b5 = 0
        b6 = 0
    End If
  
       


  AddNameOutfit 0, thename, b1
  AddNameOutfit 1, thename, b2
  AddNameOutfit 2, thename, b3
  AddNameOutfit 3, thename, b4
  AddNameOutfit 4, thename, b5
  AddNameOutfit 5, thename, b6
  frmWarbot.lstGroups.AddItem thename
  Exit Sub
gotErr:
  frmWarbot.Caption = "Load ERROR (" & Err.Number & "):" & Err.Description
  gotDictErr = 2
End Sub

Public Function AddEnemy(strName As String) As Long
  ' add item to dictionary
  Dim lStrName As String
  lStrName = LCase(strName)
  If isEnemy(lStrName) = True Then
    AddEnemy = -1
  Else
    EnemyList.item(lStrName) = True
    AddEnemy = 0
  End If
End Function
Public Sub RemoveAllEnemies()
  ' remove item from dictionary
  EnemyList.RemoveAll
End Sub
Public Function isEnemy(ByRef strName As String) As Boolean
  Dim lStrName As String
  lStrName = LCase(strName)
  ' get the name from an ID
  If EnemyList.Exists(lStrName) = True Then
    isEnemy = True
  Else
    isEnemy = False
  End If
End Function

Public Sub LoadEnemies()
  #If FinalMode = 1 Then
  On Error GoTo gotErr
  #End If
  Dim fso As Scripting.FileSystemObject
  Dim fn As Integer
  Dim strLine As String
  Dim Filename As String
  Dim res As Long
  Set fso = New Scripting.FileSystemObject
    RemoveAllEnemies
    Filename = App.Path & "\wargroups\" & "enemies.txt"
    If fso.FileExists(Filename) = True Then
      fn = FreeFile
      Open Filename For Input As #fn
      While Not EOF(fn)
        Line Input #fn, strLine
        If strLine <> "" Then
          res = AddEnemy(LCase(strLine))
        End If
      Wend
      Close #fn
    End If
  Exit Sub
gotErr:
  RemoveAllEnemies
End Sub

Public Sub LoadWarbotFiles()
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  frmWarbot.lstGroups.Clear
  Dim found As Long
  Dim Filename As String
  Dim fs As Scripting.FileSystemObject
  Dim f As Scripting.Folder
  Dim f1 As Scripting.File
  Set fs = New Scripting.FileSystemObject
  found = 0
  If (fs.FolderExists(App.Path & "\wargroups") = False) Then
    LogOnFile "errors.txt", "PLEASE UNZIP ALL: This path was not found: " & App.Path & "\wargroups"
  Else
    Set f = fs.GetFolder(App.Path & "\wargroups")
    For Each f1 In f.Files
      If LCase(Right(f1.name, 3)) = "txt" Then
        LoadFile f1.name
        found = found + 1
      End If
    Next
  End If
  If found > 0 Then
    frmWarbot.lstGroups.ListIndex = 0
    Filename = frmWarbot.lstGroups.List(0)
    LoadGroupOutfit Filename
  End If
  ReLoadAllCharOutfits
  LoadEnemies
  Exit Sub
gotErr:
  LogOnFile "errors.txt", "ERROR WITH FILESYSTEM OBJECT at LoadWarbotFiles (err number=" & _
   CStr(Err.Number) & " ; desc=" & Err.Description & ")"
  gotDictErr = 3
End Sub


Public Sub LoadGroupOutfit(groupFileName As String)
  Dim b0 As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim b4 As Byte
  Dim b5 As Byte
  b0 = GetOutfitByteFromName(0, groupFileName)
  b1 = GetOutfitByteFromName(1, groupFileName)
  b2 = GetOutfitByteFromName(2, groupFileName)
  b3 = GetOutfitByteFromName(3, groupFileName)
  b4 = GetOutfitByteFromName(4, groupFileName)
  b5 = GetOutfitByteFromName(5, groupFileName)
  frmWarbot.txtOutfit(0) = CStr(CLng(b0))
  frmWarbot.txtOutfit(1) = CStr(CLng(b1))
  frmWarbot.txtOutfit(2) = CStr(CLng(b2))
  frmWarbot.txtOutfit(3) = CStr(CLng(b3))
  frmWarbot.txtOutfit(4) = CStr(CLng(b4))
  frmWarbot.txtOutfit(5) = CStr(CLng(b5))
End Sub






Public Sub AddNameOutfit(idByte As Integer, ByVal thenamepar As String, ByVal thebyte As Byte)
  ' add item to dictionary
  Dim res As Boolean
  Dim thename As String
  thename = LCase(thenamepar)
  OutfitOfName(idByte).item(thename) = thebyte
End Sub
Public Sub RemoveName(idByte As Integer, ByVal thenamepar As String)
  ' remove item from dictionary
  Dim res As Boolean
  Dim thename As String
  thename = LCase(thenamepar)
  If OutfitOfName(idByte).Exists(thename) = True Then
    OutfitOfName(idByte).Remove (thename)
  End If
End Sub
Public Function GetOutfitByteFromName(idByte As Integer, ByRef thenamepar As String) As Byte
  ' get the name from an ID
  Dim aRes As Long
  Dim res As Boolean
  Dim thename As String
  thename = LCase(thenamepar)
  If OutfitOfName(idByte).Exists(thename) = True Then
    GetOutfitByteFromName = OutfitOfName(idByte).item(thename)
  Else
    If (idByte = 0) Then
      GetOutfitByteFromName = CByte(firstValidOutfit)
    Else
      GetOutfitByteFromName = &H0
    End If
  End If
End Function

Public Sub AddCharOutfit(idByte As Integer, ByVal thenamepar As String, ByVal thebyte As Byte)
  ' add item to dictionary
  Dim res As Boolean
  Dim thename As String
  thename = LCase(thenamepar)
  OutfitOfChar(idByte).item(thename) = thebyte
End Sub
Public Sub RemoveChar(idByte As Integer, ByVal thenamepar As String)
  ' remove item from dictionary
  Dim res As Boolean
  Dim thename As String
  thename = LCase(thenamepar)
  If OutfitOfChar(idByte).Exists(thename) = True Then
    OutfitOfChar(idByte).Remove (thename)
  End If
End Sub
Public Function GetOutfitByteFromChar(idByte As Integer, ByRef thenamepar As String) As Byte
  ' get the name from an ID
  Dim aRes As Long
  Dim res As Boolean
  Dim thename As String
  thename = LCase(thenamepar)
  If OutfitOfChar(idByte).Exists(thename) = True Then
    GetOutfitByteFromChar = OutfitOfChar(idByte).item(thename)
  Else
    If idByte = 0 Then
      GetOutfitByteFromChar = CByte(firstValidOutfit)
    Else
      GetOutfitByteFromChar = &H0
    End If
  End If
End Function

Public Function CharIsListed(thenamepar As String) As Boolean
  ' get the name from an ID
  Dim aRes As Long
  Dim res As Boolean
  Dim thename As String
  thename = LCase(thenamepar)
  If OutfitOfChar(0).Exists(thename) = True Then
    CharIsListed = True
  Else
    CharIsListed = False
  End If
End Function

Public Sub SaveOutfit(Filename As String, b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte, b5 As Byte)
  Dim fn As Integer
  Dim strLine As String
  Dim i As Integer
  On Error GoTo justend
  fn = FreeFile
  Open App.Path & "\wargroups\" & Filename For Output As #fn
    Print #fn, CStr(CLng(b0))
    Print #fn, CStr(CLng(b1))
    Print #fn, CStr(CLng(b2))
    Print #fn, CStr(CLng(b3))
    Print #fn, CStr(CLng(b4))
    Print #fn, CStr(CLng(b5))
  Close #fn
  Filename = Left$(Filename, Len(Filename) - 3) & "txt"
  AddNameOutfit 0, Filename, b0
  AddNameOutfit 1, Filename, b1
  AddNameOutfit 2, Filename, b2
  AddNameOutfit 3, Filename, b3
  AddNameOutfit 4, Filename, b4
  AddNameOutfit 5, Filename, b5
  Exit Sub
justend:
  frmWarbot.Caption = "ERROR : Saveoutfit : " & Err.Description
End Sub

Public Sub ReLoadAllCharOutfits()
  On Error GoTo gotErr
  Dim i As Long
  Dim lastI As Long
  Dim fso As Scripting.FileSystemObject
  Dim fn As Integer
  Dim strLine As String
  Dim Filename As String
  Dim seguir As Boolean
  Dim b0 As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim b4 As Byte
  Dim b5 As Byte
  Dim groupFileName As String
  frmWarbot.lstAllNames.Clear
  lastI = (frmWarbot.lstGroups.ListCount) - 1
  For i = 0 To lastI
    groupFileName = frmWarbot.lstGroups.List(i)
    b0 = GetOutfitByteFromName(0, groupFileName)
    b1 = GetOutfitByteFromName(1, groupFileName)
    b2 = GetOutfitByteFromName(2, groupFileName)
    b3 = GetOutfitByteFromName(3, groupFileName)
    b4 = GetOutfitByteFromName(4, groupFileName)
    b5 = GetOutfitByteFromName(5, groupFileName)
    Set fso = New Scripting.FileSystemObject
    Filename = App.Path & "\wargroups\" & groupFileName
    fn = FreeFile
    Open Filename For Input As #fn
    While Not EOF(fn)
    Line Input #fn, strLine
    If Len(strLine) > 0 Then
      strLine = LCase(strLine)
      'debug.Print strLine
      AddCharOutfit 0, strLine, b0
      AddCharOutfit 1, strLine, b1
      AddCharOutfit 2, strLine, b2
      AddCharOutfit 3, strLine, b3
      AddCharOutfit 4, strLine, b4
      AddCharOutfit 5, strLine, b5
      frmWarbot.lstAllNames.AddItem strLine & "  :  " & CStr(CLng(b0)) & " " & CStr(CLng(b1)) & " " & CStr(CLng(b2)) & " " & CStr(CLng(b3)) & " " & CStr(CLng(b4)) & " " & CStr(CLng(b5))
    End If
    Wend
    Close #fn
  Next i
  Exit Sub
gotErr:
  frmWarbot.Caption = "ERROR LOADING LISTS"
  gotDictErr = 5
End Sub


Public Sub ChangeGLOBAL_FRIENDSLOWLIMIT_HP(newValue As Long)
  Dim i As Integer
  Dim aRes As Long
  Dim oldVal As Long
  oldVal = GLOBAL_FRIENDSLOWLIMIT_HP
  frmWarbot.label_scrollFriendsHP.Caption = CStr(newValue) & " %"
  GLOBAL_FRIENDSLOWLIMIT_HP = newValue
  If frmWarbot.scrollFriendsHP.value <> newValue Then
    frmWarbot.scrollFriendsHP.value = newValue
  End If
End Sub

Public Sub ChangeGLOBAL_MYSAFELIMIT_HP(newValue As Long)
  Dim i As Integer
  Dim aRes As Long
  Dim oldVal As Long
  oldVal = GLOBAL_MYSAFELIMIT_HP
  frmWarbot.label_scrollSafeToHealHP.Caption = CStr(newValue) & " %"
  GLOBAL_MYSAFELIMIT_HP = newValue
  If frmWarbot.scrollSafeToHealHP.value <> newValue Then
    frmWarbot.scrollSafeToHealHP.value = newValue
  End If
End Sub

Public Sub ProcessAllFriendHeals()
  #If FinalMode = 1 Then
    On Error GoTo gotErr
  #End If
  Dim i As Integer
  Dim mx As Long
  Dim my As Long
  Dim mz As Long
  Dim ms As Long
  Dim friendX As Long
  Dim friendY As Long
  Dim friendZ As Long
  Dim friendS As Long
  Dim myPercentHP As Long
  Dim friendID As Double
  Dim friendPercentHP As Long
  Dim friendName As String
  Dim isafriend As Boolean
  Dim aRes As Long
  For i = 1 To MAXCLIENTS ' process in all mcs
    If (CheatsPaused(i) = False) Then
    If (GameConnected(i) = True) Then ' (only in connected clients)
      myPercentHP = 100 * ((myHP(i) / myMaxHP(i)))
      If (myPercentHP >= GLOBAL_MYSAFELIMIT_HP) Then 'only if I am healthy enough to cast a heal
        mz = myZ(i)
        For mx = -7 To 7
          For my = -5 To 5
            For ms = 1 To 10
              friendID = Matrix(my, mx, mz, i).s(ms).dblID
              If (friendID <> 0) Then
                friendPercentHP = GetHPFromID(i, friendID)
                If ((friendPercentHP > 0) And (friendPercentHP <= GLOBAL_FRIENDSLOWLIMIT_HP)) Then
                  friendName = GetNameFromID(i, friendID)
                  If (friendName <> CharacterName(i)) Then
                  isafriend = frmWarbot.IsAutoHealFriend(LCase(friendName))
                  If (isafriend = True) Then
                    friendX = mx
                    friendY = my
                    friendZ = mz
                    friendS = ms
                    GoTo HealIt
                  End If
                  End If
                End If
              ElseIf ((Matrix(my, mx, mz, i).s(ms).t1 = 0) And _
                      (Matrix(my, mx, mz, i).s(ms).t2 = 0)) Then
                Exit For
              End If
            Next ms
          Next my
        Next mx
      End If
    End If
    GoTo continueIt
HealIt:
    friendName = GetNameFromID(i, friendID)
    aRes = AutoHealFriend(i, friendName, friendID, friendX, friendY, friendZ, friendS, GLOBAL_AUTOFRIENDHEAL_MODE)
    If (aRes = 0) Then
      aRes = SendLogSystemMessageToClient(i, "BlackdProxy: Autohealed " & friendName & " (at " & CStr(friendPercentHP) & " % hp)")
      DoEvents
    End If
continueIt:
    friendName = ""
    End If
  Next i
  Exit Sub
gotErr:
  LogOnFile "errors.txt", "ProcessAllFriendHeals failed with code " & CStr(Err.Number) & " : " & Err.Description
End Sub


Public Function AutoHealFriend(idConnection As Integer, friendName As String, friendID As Double, _
 friendX As Long, friendY As Long, friendZ As Long, friendS As Long, castMode As Long) As Long
  #If FinalMode = 1 Then
    On Error GoTo gotErr
  #End If
  Dim aRes As Long
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim fRes As TypeSearchItemResult2
  Dim myS As Byte
  Dim runeB1 As Byte
  Dim runeB2 As Byte
  Dim SpecialSource As Boolean
  Dim inRes As Integer
  If (castMode = 3) Then
    If TibiaVersionLong <= 760 Then
        AutoHealFriend = CastSpell(idConnection, "exura sio """ & friendName)
    Else
        AutoHealFriend = CastSpell(idConnection, "exura sio """ & friendName & """")
    End If
    Exit Function
  End If
  If (TibiaVersionLong <= 760) Then
    runeB1 = LowByteOfLong(tileID_UH)
    runeB2 = HighByteOfLong(tileID_UH)
    If (TibiaVersionLong < 760) Then
      myS = MyStackPos(idConnection)
    Else
      myS = FirstPersonStackPos(idConnection)
    End If
    ' search yourself
    If myS = &HFF Then
      aRes = SendLogSystemMessageToClient(idConnection, "Your map is out of sync, can't use auto friend heal!")
      AutoHealFriend = -1
      Exit Function
    End If
    fRes = SearchItem(idConnection, runeB1, runeB2)  'search thing
    If fRes.foundCount = 0 Then
      AutoHealFriend = -1
      Exit Function
    End If
    sCheat = "83 FF FF " & GoodHex(&H40 + fRes.bpID) & " 00 " & _
     GoodHex(fRes.slotID) & " " & GoodHex(runeB1) & " " & GoodHex(runeB2) & " " & _
     GoodHex(fRes.slotID) & " " & GetHexStrFromPosition(myX(idConnection) + friendX, _
     myY(idConnection) + friendY, myZ(idConnection)) & _
     " 63 00 " & GoodHex(CByte(friendS))
    SafeCastCheatString "AutoHealFriend1", idConnection, sCheat
    AutoHealFriend = 0
    Exit Function
  Else ' ************************************ new Tibia
    runeB1 = LowByteOfLong(tileID_UH)
    runeB2 = HighByteOfLong(tileID_UH)
    If castMode = 1 Then
      SpecialSource = False
      fRes = SearchItem(idConnection, runeB1, runeB2)  'search thing
      If fRes.foundCount = 0 Then
        aRes = SendSystemMessageToClient(idConnection, "Open UHs or I won't autoheal friends!")
        DoEvents
        AutoHealFriend = -1
        Exit Function
      End If
    Else 'castMode=2
      SpecialSource = True
    End If
    If (TibiaVersionLong < 760) Then
      myS = MyStackPos(idConnection)
    Else
      myS = FirstPersonStackPos(idConnection)
    End If
    ' search yourself
    If myS = &HFF Then
      aRes = SendLogSystemMessageToClient(idConnection, "Your map is out of sync, can't use auto friend heal!")
      AutoHealFriend = -1
      Exit Function
    End If
    If SpecialSource = False Then
      sCheat = "84 FF FF " & GoodHex(&H40 + fRes.bpID) & " 00 " & _
       GoodHex(fRes.slotID) & " " & GoodHex(runeB1) & " " & GoodHex(runeB2) & " " & _
       GoodHex(fRes.slotID) & " " & SpaceID(friendID)
      'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "DEBUG> " & sCheat
      SafeCastCheatString "AutoHealFriend2", idConnection, sCheat
      AutoHealFriend = 0
      Exit Function
    Else
      sCheat = "84 FF FF 00 00 00 " & GoodHex(runeB1) & " " & GoodHex(runeB2) & " 00 " & SpaceID(friendID)
      SafeCastCheatString "AutoHealFriend3", idConnection, sCheat
      AutoHealFriend = 0
      Exit Function
    End If
  End If
  AutoHealFriend = 0
  Exit Function
gotErr:
  AutoHealFriend = -1
End Function

Public Sub RecordLoginOnFile(CharacterName As String, IPstring As String, _
 PORTnumber As Long, ByRef Index As Integer)
 Dim a As Integer
 On Error GoTo ignoreit
 If RecordLogin = True Then
   Dim fn As Integer
   Dim loginfilename As String
   fn = FreeFile
 
   If CharacterName <> "" Then
     loginfilename = App.Path & "\magebomb\" & CharacterName & ".log"
     Open loginfilename For Output As #fn
     Print #fn, CharacterName
     'Print #fn, FiveChrLon(tileID_SD)
     Print #fn, CStr(TibiaVersionLong)
     Print #fn, IPstring
     Print #fn, CStr(PORTnumber)
     If TibiaVersionLong <= 760 Then
       Print #fn, "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00"
     Else
       Print #fn, GoodHex(packetKey(Index).Key(0)) & " " & _
        GoodHex(packetKey(Index).Key(1)) & " " & _
        GoodHex(packetKey(Index).Key(2)) & " " & _
        GoodHex(packetKey(Index).Key(3)) & " " & _
        GoodHex(packetKey(Index).Key(4)) & " " & _
        GoodHex(packetKey(Index).Key(5)) & " " & _
        GoodHex(packetKey(Index).Key(6)) & " " & _
        GoodHex(packetKey(Index).Key(7)) & " " & _
        GoodHex(packetKey(Index).Key(8)) & " " & _
        GoodHex(packetKey(Index).Key(9)) & " " & _
        GoodHex(packetKey(Index).Key(10)) & " " & _
        GoodHex(packetKey(Index).Key(11)) & " " & _
        GoodHex(packetKey(Index).Key(12)) & " " & _
        GoodHex(packetKey(Index).Key(13)) & " " & _
        GoodHex(packetKey(Index).Key(14)) & " " & _
        GoodHex(packetKey(Index).Key(15))
     End If
     Print #fn, CStr(UBound(ReconnectionPacket(Index).packet))
     Print #fn, frmMain.showAsStr2(ReconnectionPacket(Index).packet, 2)
   Close #fn
   frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & ">Successfully recorded: " & loginfilename
  End If
End If
Exit Sub
ignoreit:
  a = -1
End Sub

Public Function ExistMagebombCharInMemory(strCharname As String)
  On Error GoTo gotErr
  Dim i As Long
  Dim res As Boolean
  Dim limt As Long
  res = False
  limt = MagebombsLoaded - 1
  For i = 0 To limt
    If Magebombs(i).CharacterName = strCharname Then
      res = True
    End If
  Next i
  ExistMagebombCharInMemory = res
  Exit Function
gotErr:
  ExistMagebombCharInMemory = False
End Function

Public Sub AddToMagebombMemory(AddingLogFileName As String, AddingCharname As String, AddingVersion As Long, AddingIP As String, _
 AddingPort As Long, AddingRawKey As String, AddingUBoundRawLoginPacket As Long, AddingRawLoginPacket As String, _
 AddingMode As String, AddingTarget As String, AddingTime As Long)
  #If FinalMode = 1 Then
  On Error GoTo gotErr
  #End If
  Dim strTmp As String
  Dim inRes As Integer
  Dim keyPacket() As Byte
  Dim loginPacket() As Byte
  Dim i As Long
  ReDim Preserve Magebombs(MagebombsLoaded)
  Magebombs(MagebombsLoaded).AttackMode = AddingMode
  Magebombs(MagebombsLoaded).CharacterName = AddingCharname
  Magebombs(MagebombsLoaded).IPstring = AddingIP
  Magebombs(MagebombsLoaded).LoginVersion = AddingVersion
  Magebombs(MagebombsLoaded).LogFileName = AddingLogFileName
  Magebombs(MagebombsLoaded).PORTnumber = AddingPort
  Magebombs(MagebombsLoaded).RetryTime = AddingTime
  Magebombs(MagebombsLoaded).TargetToShot = AddingTarget
  Magebombs(MagebombsLoaded).connectionStatus = 0
  Magebombs(MagebombsLoaded).ConnectionTimeout = 0
  Magebombs(MagebombsLoaded).nextSendLogin = 0
  strTmp = AddingRawKey
  inRes = GetCheatPacket(keyPacket, strTmp)
  strTmp = AddingRawLoginPacket
  inRes = GetCheatPacket(loginPacket, strTmp)
  ReDim Magebombs(MagebombsLoaded).loginPacket(AddingUBoundRawLoginPacket)
  For i = 0 To 15
    Magebombs(MagebombsLoaded).Key(i) = keyPacket(i)
  Next i
  For i = 0 To AddingUBoundRawLoginPacket
    Magebombs(MagebombsLoaded).loginPacket(i) = loginPacket(i)
  Next i
  MagebombsLoaded = MagebombsLoaded + 1
  Exit Sub
gotErr:
  MagebombsLoaded = 0
  ReDim Magebombs(0)
End Sub

Public Function DeleteMagebombMemory(elementID As Long) As Long
  #If FinalMode = 1 Then
  On Error GoTo gotErr
  #End If
  Dim i As Long
  Dim j As Long
  Dim limM As Long
  Dim limN As Long
  If (elementID = 0) And (MagebombsLoaded < 2) Then
    MagebombsLoaded = 0
    ReDim Magebombs(0)
  ElseIf (elementID = MagebombsLoaded - 1) Then
    MagebombsLoaded = MagebombsLoaded - 1
    ReDim Preserve Magebombs(MagebombsLoaded - 1)
  Else
    limM = MagebombsLoaded - 2
    For i = elementID To limM
      Magebombs(i).AttackMode = Magebombs(i + 1).AttackMode
      Magebombs(i).CharacterName = Magebombs(i + 1).CharacterName
      Magebombs(i).IPstring = Magebombs(i + 1).IPstring
      For j = 0 To 15
        Magebombs(i).Key(j) = Magebombs(i + 1).Key(j)
      Next j
      limN = UBound(Magebombs(i + 1).loginPacket)
      ReDim Magebombs(i).loginPacket(limN)
      For j = 0 To limN
        Magebombs(i).loginPacket(j) = Magebombs(i + 1).loginPacket(j)
      Next j
      Magebombs(i).LoginVersion = Magebombs(i + 1).LoginVersion
      Magebombs(i).LogFileName = Magebombs(i + 1).LogFileName
      Magebombs(i).PORTnumber = Magebombs(i + 1).PORTnumber
      Magebombs(i).RetryTime = Magebombs(i + 1).RetryTime
      Magebombs(i).TargetToShot = Magebombs(i + 1).TargetToShot
      Magebombs(i).connectionStatus = Magebombs(i + 1).connectionStatus
      Magebombs(i).ConnectionTimeout = Magebombs(i + 1).ConnectionTimeout
    Next i
    MagebombsLoaded = MagebombsLoaded - 1
    ReDim Preserve Magebombs(MagebombsLoaded - 1)
  End If
  DeleteMagebombMemory = 0
  Exit Function
gotErr:
  MagebombsLoaded = 0
  ReDim Magebombs(0)
  DeleteMagebombMemory = -1
End Function

Public Sub DeleteAllMagebombMemory()
  MagebombsLoaded = 0
  ReDim Magebombs(0)
End Sub

Public Function ExecuteMagebomb(idConnection As Integer, givenTarget As String) As Long
  Dim aRes As Long
  Dim i As Long
  Dim limM As Long
  Dim mustRestart As Boolean
  Dim gotNewTargetError As Boolean
  Dim gtc As Long
  Dim bRes As Long
  Dim dRes As Long
  Dim defaultTarget As String
  limM = MagebombsLoaded - 1
  If Not (givenTarget = "") Then
     For i = 0 To limM
       Magebombs(i).TargetToShot = givenTarget
     Next i
     frmMagebomb.DisplayMagebombMemory
  End If
  If MagebombLeader > 0 Then
    If MagebombsLoaded > 0 Then
      mustRestart = False
      For i = 0 To limM
        If Magebombs(i).connectionStatus = 0 Then
          mustRestart = True
        End If
      Next i
      If mustRestart = False Then
        gtc = GetTickCount()
        gotNewTargetError = False
        For i = 0 To limM
          Magebombs(i).ConnectionTimeout = gtc + Magebombs(i).RetryTime
          Magebombs(i).nextSendLogin = 0
          bRes = BuildAttackPacket(i, Magebombs(i).TargetToShot)
          If bRes < 0 Then
            gotNewTargetError = True
          End If
        Next i
        If (gotNewTargetError = True) Then
          aRes = GiveGMmessage(idConnection, "Can't see the new target on screen (" & Magebombs(0).TargetToShot & ")", "Error")
          DoEvents
        Else
          aRes = SendLogSystemMessageToClient(idConnection, "Blackdproxy: Successfully changed magebomb target to " & Magebombs(0).TargetToShot)
          DoEvents
        If DebugingMagebomb = True Then
          dRes = SendLogSystemMessageToClient(MagebombLeader, CStr(GetTickCount() - MagebombStartTime) & " ms : Timers extended, magebomb restarted with new target: " & Magebombs(0).TargetToShot)
          DoEvents
        End If
        End If
        ExecuteMagebomb = 0
        Exit Function
      Else ' must restart
        frmMagebomb.armageddonTimer.enabled = False
        aRes = GiveGMmessage(idConnection, "Reloading magebomb with new target : " & Magebombs(0).TargetToShot, "Warning")
        DoEvents
      End If
    End If
  End If
  
    If MagebombsLoaded = 0 Then
      aRes = GiveGMmessage(idConnection, "The magebomb is not ready yet. First preload the characters", "Error")
      DoEvents
      ExecuteMagebomb = 0
      Exit Function
    End If
    ' all ok
    MagebombLeader = idConnection
    For i = 0 To limM
      Magebombs(0).connectionStatus = 0
      If (i > frmMagebomb.clientLess.UBound) Then
         Load frmMagebomb.clientLess(i)
      End If
      frmMagebomb.clientLess(i).Close
      DoEvents
      bRes = BuildAttackPacket(i, Magebombs(i).TargetToShot)
      If bRes = -1 Then
        MagebombLeader = 0
        aRes = GiveGMmessage(idConnection, "The magebomb won't shoot : " & Magebombs(i).TargetToShot & " is not in your track", "Error")
        DoEvents
        ExecuteMagebomb = 0
        Exit Function
      End If
      If bRes = -2 Then
        MagebombLeader = 0
        aRes = GiveGMmessage(idConnection, "The magebomb won't shoot : The magebomb is not compatible with this Tibia version", "Error")
        DoEvents
        ExecuteMagebomb = 0
        Exit Function
      End If
    Next i
    MagebombStage = 1
    MagebombStartTime = GetTickCount()
        If DebugingMagebomb = True Then
          dRes = SendLogSystemMessageToClient(MagebombLeader, "0 ms : Magebomb script started.")
          DoEvents
        End If
    frmMagebomb.armageddonTimer.enabled = True
    frmMagebomb.ProcessArmageddon
    ExecuteMagebomb = 0
    Exit Function
gotErr:
  LogOnFile "errors.txt", "ExecuteMagebomb() failed with code number: " & CStr(Err.Number) & " and description: " & Err.Description
  ExecuteMagebomb = -1
End Function

Public Function BuildAttackPacket(magebombID As Long, defaultTarget As String) As Long
  Dim idConnection As Integer
  Dim Target As String
  Dim runeB1 As Byte
  Dim runeB2 As Byte
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
  Dim z As Long
  Dim tileID As Long
  Dim tmpID As Double
  Dim inRes As Integer
  Dim isDamageRune As Boolean
  Dim percent As Long
  Dim limP As Long
  Dim i As Long
  Dim idsOnMemory
  Dim lim As Long
  Dim currItem As Double
  Dim currName As String
  Dim TheIDisFound As Boolean
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  idConnection = MagebombLeader
  If idConnection < 1 Then
    BuildAttackPacket = -1
    Exit Function
  End If
  Select Case Magebombs(magebombID).AttackMode
  Case "5"
    runeB1 = LowByteOfLong(tileID_SD)
    runeB2 = HighByteOfLong(tileID_SD)
  Case "6"
    runeB1 = LowByteOfLong(tileID_HMM)
    runeB2 = HighByteOfLong(tileID_HMM)
  Case "7"
    runeB1 = LowByteOfLong(tileID_Explosion)
    runeB2 = HighByteOfLong(tileID_Explosion)
  Case "8"
    runeB1 = LowByteOfLong(tileID_UH)
    runeB2 = HighByteOfLong(tileID_UH)
  Case "9"
    runeB1 = LowByteOfLong(tileID_IH)
    runeB2 = HighByteOfLong(tileID_IH)
  Case "B"
    runeB1 = LowByteOfLong(tileID_fireball)
    runeB2 = HighByteOfLong(tileID_fireball)
  Case "C"
    runeB1 = LowByteOfLong(tileID_stalagmite)
    runeB2 = HighByteOfLong(tileID_stalagmite)
  Case "D"
    runeB1 = LowByteOfLong(tileID_icicle)
    runeB2 = HighByteOfLong(tileID_icicle)
  Case Else
    BuildAttackPacket = -1
    Exit Function
  End Select
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
  'SpecialSource = True
  If defaultTarget = "" Then
    Target = Magebombs(magebombID).TargetToShot
  Else
    Target = defaultTarget
  End If
  If Magebombs(magebombID).LoginVersion <= 760 Then
    BuildAttackPacket = -2
    Exit Function
  End If
  ' search the creature
  lTarget = LCase(Target)
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
  
  If (TheIDisFound = False) Then
    BuildAttackPacket = -1
    Exit Function
  End If
                   '0D 00 84 FF FF 00 00 00 7E 0C 00 40 0A D3 00
          sCheat = "0D 00 84 FF FF 00 00 00 " & GoodHex(runeB1) & " " & GoodHex(runeB2) & " 00 " & SpaceID(tmpID)
          'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "> " & sCheat
          
          inRes = GetCheatPacket(cPacket, sCheat)
          'frmMain.UnifiedSendToServerGame idConnection, cPacket, True
          limP = UBound(cPacket)
          ReDim Magebombs(magebombID).attackPacket(limP)
          For i = 0 To limP
            Magebombs(magebombID).attackPacket(i) = cPacket(i)
          Next i
          BuildAttackPacket = 0
          Exit Function
errclose:
  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at SendMobAimbot #"
  'frmMain.DoCloseActions idConnection
  'DoEvents
  BuildAttackPacket = -1
End Function

Public Function TellBestEnemy(idConnection As Integer) As String
  #If FinalMode = 1 Then
  On Error GoTo gotErr
  #End If
  Dim res As String
  Dim y As Long
  Dim x As Long
  Dim s As Byte
  Dim z As Long
  Dim lSquare As String
  Dim bestEnemy As String
  Dim bestHP As Long
  Dim tmpID As Double
  Dim currHP As Long
  res = ""
  bestEnemy = "NOENEMYFOUND"
  bestHP = 101
  z = myZ(idConnection)
  If GameConnected(idConnection) = False Then
    TellBestEnemy = "NOTCONNECTED"
    Exit Function
  End If
  For y = -5 To 6
    For x = -7 To 8
      For s = 1 To 10
        tmpID = Matrix(y, x, z, idConnection).s(s).dblID
        If tmpID <> 0 Then
          lSquare = GetNameFromID(idConnection, tmpID)
          If isEnemy(lSquare) Then
            currHP = CLng(GetHPFromID(idConnection, tmpID))
            If currHP < bestHP Then
                bestEnemy = lSquare
                bestHP = currHP
            End If
          End If
        End If
      Next s
    Next x
  Next y
  res = bestEnemy

  TellBestEnemy = res
  Exit Function
gotErr:
  res = "ERROR " & CStr(Err.Number) & " : " & Err.Description
  TellBestEnemy = res
End Function

Public Function TellBestEnemyID(idConnection As Integer) As Double
  #If FinalMode = 1 Then
  On Error GoTo gotErr
  #End If
  Dim res As String
  Dim y As Long
  Dim x As Long
  Dim s As Byte
  Dim z As Long
  Dim lSquare As String
  Dim bestEnemyID As Double
  Dim bestHP As Long
  Dim tmpID As Double
  Dim currHP As Long
  res = ""
  bestEnemyID = 0
  bestHP = 101
  z = myZ(idConnection)
  If GameConnected(idConnection) = False Then
     TellBestEnemyID = 0
    Exit Function
  End If
  For y = -5 To 6
    For x = -7 To 8
      For s = 1 To 10
        tmpID = Matrix(y, x, z, idConnection).s(s).dblID
        If tmpID <> 0 Then
          lSquare = GetNameFromID(idConnection, tmpID)
          If isEnemy(lSquare) Then
            currHP = CLng(GetHPFromID(idConnection, tmpID))
            If currHP < bestHP Then
                bestEnemyID = tmpID
                bestHP = currHP
            End If
          End If
        End If
      Next s
    Next x
  Next y

  TellBestEnemyID = bestEnemyID
  Exit Function
gotErr:
 TellBestEnemyID = 0
End Function

Public Function TellBestEnemyHP(idConnection As Integer) As Long
  #If FinalMode = 1 Then
  On Error GoTo gotErr
  #End If
  Dim res As String
  Dim y As Long
  Dim x As Long
  Dim s As Byte
  Dim z As Long
  Dim lSquare As String
  Dim bestEnemyID As Double
  Dim bestHP As Long
  Dim tmpID As Double
  Dim currHP As Long
  res = ""
  bestEnemyID = 0
  bestHP = 101
  z = myZ(idConnection)
  If GameConnected(idConnection) = False Then
     TellBestEnemyHP = 0
    Exit Function
  End If
  For y = -5 To 6
    For x = -7 To 8
      For s = 1 To 10
        tmpID = Matrix(y, x, z, idConnection).s(s).dblID
        If tmpID <> 0 Then
          lSquare = GetNameFromID(idConnection, tmpID)
          If isEnemy(lSquare) Then
            currHP = CLng(GetHPFromID(idConnection, tmpID))
            If currHP < bestHP Then
                bestEnemyID = tmpID
                bestHP = currHP
            End If
          End If
        End If
      Next s
    Next x
  Next y
  If bestHP = 101 Then
    TellBestEnemyHP = 0
  Else
    TellBestEnemyHP = bestHP
  End If
  Exit Function
gotErr:
 TellBestEnemyHP = 0
End Function

Public Function DoubleToStr(address As Double) As String
    On Error GoTo gotErr
    Dim res As String
    Dim b1 As Byte
    Dim b2 As Byte
    Dim b3 As Byte
    Dim b4 As Byte
    b1 = Byte1ofDouble(address)
    b2 = Byte2ofDouble(address)
    b3 = Byte3ofDouble(address)
    b4 = Byte4ofDouble(address)
    res = GoodHex(b4) & " " & GoodHex(b3) & " " & GoodHex(b2) & " " & GoodHex(b1)
    DoubleToStr = res
    Exit Function
gotErr:
    DoubleToStr = "00 00 00 00"
End Function
Public Function Byte1ofDouble(address As Double) As Byte
  Dim h As Byte
  Dim resT As Double
  resT = address
  h = CByte(resT \ 16777216) ' high byte
  Byte1ofDouble = h
End Function

Public Function Byte2ofDouble(address As Double) As Byte
  Dim h As Byte
  Dim resT As Double
  Dim minusF As Double
  resT = address
  minusF = (Byte1ofDouble(address) * 16777216)
  resT = resT - minusF
  h = CByte(resT \ 65536)
  Byte2ofDouble = h
End Function

Public Function Byte3ofDouble(address As Double) As Byte
  Dim h As Byte
  Dim resT As Double
  Dim minusF As Double
  resT = address
  minusF = (Byte1ofDouble(address) * 16777216)
  resT = resT - minusF
  minusF = (Byte2ofDouble(address) * 65536)
  resT = resT - minusF
  h = CByte(resT \ 256)
  Byte3ofDouble = h
End Function

Public Function Byte4ofDouble(address As Double) As Byte
  Dim h As Byte
  Dim resT As Double
  Dim minusF As Double
  resT = address
  minusF = (Byte1ofDouble(address) * 16777216)
  resT = resT - minusF
  minusF = (Byte2ofDouble(address) * 65536)
  resT = resT - minusF
  minusF = (Byte3ofDouble(address) * 256)
  resT = resT - minusF
  h = CByte(resT)
  Byte4ofDouble = h
End Function
