Attribute VB_Name = "modCrackd"
#Const FinalMode = 1
Option Explicit

Public Type TypeTibiaKey
 Key(15) As Byte
End Type
' firstPacketByte is the first byte of the packet array
' crackd.dll functions expect to find the rest of the packet bytes after firstPacketByte
' firstKeyByte is the first byte of the 16bytes-key array
' crackd.dll functions expect to find the other 15 bytes of the key after firstKeyByte
' crackd.dll functions expect to receive a packet with size (8*n) + 2 bytes
'   (fill with random trash if required)
' crackd.dll functions expect to receive the packet size (in bytes)
'   in the first two bytes of the packet array


Public Declare Function EncipherTibiaProtected Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long

Public Declare Function DecipherTibiaProtected Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long
      
Public Declare Function EncipherTibiaProtectedSP Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long

Public Declare Function DecipherTibiaProtectedSP Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long

Public Declare Function GetTibiaCRC Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, ByVal uboundpacketMinus6 As Long) As Long
    
Public Declare Function BlackdForceWrite Lib _
    "crackd.dll" (ByVal address As Long, ByRef mybuffer As Byte, ByVal mybuffersize As Long, ByVal hwndClientWindow As Long) As Long
    

Public Declare Sub RtlMoveMemory Lib "kernel32" ( _
    lpDest As Any, _
    lpSource As Any, _
    ByVal ByValcbCopy As Long)
    
Public packetKey() As TypeTibiaKey
Public loginPacketKey() As TypeTibiaKey
Public gotFirstLoginPacket() As Boolean
Public UseCrackd As Boolean
Public adrConnectionKey As AddressPath

Public adrSelectedCharIndex As AddressPath
Public adrSelectedItem_height As AddressPath
Public adrSelectedCharName As AddressPath
Public adrServerList_CollectionStart As AddressPath
Public adrBattlelist_CollectionStart As AddressPath

Public adrSelectedCharName_afterCharList As AddressPath
Public adrSelectedServerURL_afterCharList As AddressPath
Public adrSelectedServerPORT_afterCharList As AddressPath
Public adrSelectedServerNAME_afterCharList As AddressPath

Public offSetSquare_ARGB_8bytes As Long
Public adrNewRedSquare As AddressPath
Public adrNewBlueSquare As AddressPath

Public adrLastPacket As Long
Public adrCharListPtr As Long
Public adrCharListPtrEND As Long
Public debugStrangeFail As String
Public MAXCHARACTERLEN As Long
Public manualDebugOrder As Long
Public GameServerDictionary As Scripting.Dictionary  ' A dictionary server (string) -> IP (string)
Public GameServerDictionaryDOMAIN As Scripting.Dictionary

Public Sub JustReadPID(idConnection As Integer)
 ' should be only used at login stage, in tibia 7.63+
  Dim tibiaclient As Long
  'Dim hWndDesktop As Long
  Dim status As Long
  Dim abyte As Byte
  Dim sucess As Long
  Dim i As Integer
  sucess = -3
  If (GameConnected(idConnection) = True) Then
    ' keys will be only read at login
    Exit Sub
  End If
  ProcessID(idConnection) = 0
  'hWndDesktop = GetDesktopWindow()
  tibiaclient = 0
  Do
    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      Exit Do
    Else
      status = Memory_ReadLong(adrConnected, tibiaclient)
      If (status <> 0) Then ' doing login
        sucess = 0
        ProcessID(idConnection) = tibiaclient
        Exit Do
      End If
    End If
  Loop
End Sub

Public Function readLoginTibiaKeyAtPID(idConnection As Integer, ProcessID As Long) As Long
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  Dim startAdr As Long
  Dim abyte As Byte
  Dim i As Integer
  If (ProcessID = -1) Then
    readLoginTibiaKeyAtPID = -1
    Exit Function
  Else
    startAdr = ReadCurrentAddress(ProcessID, adrConnectionKey, -1, False)
    If (startAdr = -1) Then
        readLoginTibiaKeyAtPID = -1
        Exit Function
    End If
    For i = 0 To 15
      abyte = Memory_ReadByte(startAdr + i, ProcessID)
      loginPacketKey(idConnection).Key(i) = abyte
    Next i
    readLoginTibiaKeyAtPID = 0
  End If
  Exit Function
gotErr:
  readLoginTibiaKeyAtPID = -1
End Function

Public Function readTibiaKeyAtPID(ByVal idConnection As Integer, ByVal ProcessID As Long) As Long
    Dim abyte As Byte
    Dim i As Integer
    Dim startAdr As Long
    Dim allzeroes As Boolean
    Dim t As Integer
    Dim errorMsg As String
    startAdr = ReadCurrentAddress(ProcessID, adrConnectionKey, -1, False)
    If (startAdr = -1) Then
        errorMsg = "Failed to obtain XTEA key!"
        Debug.Print errorMsg
        If cteDebugConEvents = True Then
          errorMsg = errorMsg & vbCrLf & conEventLog
        End If
        LogOnFile "errors.txt", errorMsg
        readTibiaKeyAtPID = -1
        Exit Function
    End If
    allzeroes = True
    For i = 0 To 15
        abyte = Memory_ReadByte(startAdr + i, ProcessID)
        packetKey(idConnection).Key(i) = abyte
        If Not (abyte = &H0) Then
            allzeroes = False
        End If
    Next i
    If (allzeroes) Then
        errorMsg = "Failed to obtain XTEA key! (address value is zero)"
        Debug.Print errorMsg
        If cteDebugConEvents = True Then
          errorMsg = errorMsg & vbCrLf & conEventLog
        End If
        LogOnFile "errors.txt", errorMsg
        readTibiaKeyAtPID = -1
    Else
        If cteDebugConEvents = True Then
           LogConEvent "Obtained XTEA key : " & frmMain.showAsStr(packetKey(idConnection).Key, True)
           OverwriteOnFile "connEventsLog.txt", conEventLog
           ResetConEventLogs
        End If
        readTibiaKeyAtPID = 0
    End If
End Function

Public Function CompareLastPacket(ByVal pid As Long, ByRef packet() As Byte) As Boolean
  Dim res As Boolean
  Dim i As Long
  Dim lngp As Long
  Dim b As Byte
  Dim errmessage As String
  On Error GoTo cantdoit
  If UBound(packet) < 1 Then
    CompareLastPacket = False
    Exit Function
  End If
  res = True
  lngp = GetTheLong(packet(0), packet(1)) + 1
  For i = 0 To lngp
    b = Memory_ReadByte(adrLastPacket + i, pid)
    If b <> packet(i) Then
      res = False
      Exit For
    End If
  Next i
  CompareLastPacket = res
  Exit Function
cantdoit:
  errmessage = "Function failure : CompareLastPacket failed : Error number " & CStr(Err.Number) & " : " & Err.Description
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & errmessage
  LogOnFile "errors.txt", errmessage
  CompareLastPacket = False
End Function

Public Function GetLastPacket(pid As Long, lngp As Long) As String
  Dim res As Boolean
  Dim i As Long
  Dim b As Byte
  Dim errmessage As String
  Dim packetR() As Byte
  On Error GoTo cantdoit
  ReDim packetR(lngp)
  For i = 0 To lngp
    b = Memory_ReadByte(adrLastPacket + i, pid)
    packetR(i) = b
  Next i
  GetLastPacket = frmMain.showAsStr2(packetR, 0)
  Exit Function
cantdoit:
  GetLastPacket = "ERROR"
End Function

Public Sub UpdateProcessIDbyLastPacket(ByVal idConnection As Integer, ByRef packet() As Byte, Optional strIP As String = "")
  Dim tibiaclient As Long
  'Dim hWndDesktop As Long
  Dim status As Byte
  Dim abyte As Byte
  Dim sucess As Long
  Dim i As Integer
  Dim errmessage As String
  On Error GoTo gotErr
  sucess = -2
  ProcessID(idConnection) = 0
  If AlternativeBinding <> 0 Then
    If strIP <> "" Then
      ProcessID(idConnection) = GetProcessIdFromIP(strIP)
      Exit Sub
    End If
  End If
  'hWndDesktop = GetDesktopWindow()
  tibiaclient = 0
  Do

    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      Exit Do
    Else
      
      If CompareLastPacket(tibiaclient, packet) = True Then
        ProcessID(idConnection) = tibiaclient
        sucess = 0
        Exit Do
      End If
    End If
  Loop
  If (sucess = 0) Then
    Exit Sub
  End If
  errmessage = "Warning on function UpdateProcessIDbyLastPacket : could not find match"
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & errmessage
  LogOnFile "errors.txt", errmessage
  Exit Sub
gotErr:
  errmessage = "Function failure : UpdateProcessIDbyLastPacket could not match idconnection<->pid : Error number " & CStr(Err.Number) & " : " & Err.Description
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & errmessage
  LogOnFile "errors.txt", errmessage
End Sub


Public Function GiveProcessIDbyLastPacket(ByRef packet() As Byte, Optional strIP As String = "", Optional FromIP As String = "?", Optional part As String = "LOGIN1") As Long
  Dim tibiaclient As Long
  Dim hWndDesktop As Long
  Dim status As Byte
  Dim abyte As Byte
  Dim sucess As Long
  Dim i As Integer
  Dim errmessage As String
  Dim res As Long
  Dim comparing1 As String
  Dim comparing2 As String
  Dim tcount As Long
  Dim trivialRes As Long
  Dim packetSizeForComparing As Long
  #If FinalMode = 1 Then
  On Error GoTo gotErr
  #End If
  tcount = 0
  
  If AlternativeBinding <> 0 Then
    If strIP <> "" Then
      GiveProcessIDbyLastPacket = GetProcessIdFromIP(strIP)
      Exit Function
    End If
  End If
  debugStrangeFail = ""
  
  debugStrangeFail = "WARNING on GiveProcessIDbyLastPacket . Doing a complete report:"
  res = 0
  sucess = -2
  packetSizeForComparing = GetTheLong(packet(0), packet(1)) ' fix since 11.7 : lets only compare first subpacket
  
  comparing1 = frmMain.showAsStr2(packet, 0, packetSizeForComparing + 1)
  
  'hWndDesktop = GetDesktopWindow()
  'debugStrangeFail = debugStrangeFail & vbCrLf & "GetDesktopWindow() returned " & CStr(hWndDesktop)
  debugStrangeFail = debugStrangeFail & vbCrLf & "Now trying to determine what client sent the packet that Blackd Proxy just received."
  debugStrangeFail = debugStrangeFail & vbCrLf & "BLACKDPROXY RECEIVED, from ip [" & FromIP & "] at " & part & " :" & comparing1
  tibiaclient = 0
  Do

    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      debugStrangeFail = debugStrangeFail & vbCrLf & "Found a total of " & CStr(tcount) & " Tibia client(s) opened"
      Exit Do
    Else
      trivialRes = tibiaclient
      tcount = tcount + 1
      comparing2 = GetLastPacket(tibiaclient, packetSizeForComparing + 1)
      debugStrangeFail = debugStrangeFail & vbCrLf & "CLIENT #" & CStr(tcount) & " HAVE SENT :" & comparing2
      If (comparing1 = comparing2) Then
        debugStrangeFail = debugStrangeFail & vbCrLf & " ...MATCH at pid " & CStr(tibiaclient)
        res = tibiaclient
        sucess = 0
        'Exit Do
      Else
        debugStrangeFail = debugStrangeFail & vbCrLf & " ...FAIL! at pid " & CStr(tibiaclient)
      End If
    End If
  Loop
  If (sucess = 0) Then
    debugStrangeFail = debugStrangeFail & vbCrLf & "Function worked fine."
  Else
    debugStrangeFail = debugStrangeFail & vbCrLf & "Function failed!"
    If tcount = 1 Then
        debugStrangeFail = debugStrangeFail & vbCrLf & "However, there is a trivial match since only 1 client was detected: " & CStr(trivialRes)
        sucess = 0
        res = trivialRes
    Else
        debugStrangeFail = debugStrangeFail & vbCrLf & "Please report to daniel@blackdtools.com"
    End If
  End If
  If (sucess = 0) Then
    GiveProcessIDbyLastPacket = res
    Exit Function
  End If
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & debugStrangeFail
  LogOnFile "errors.txt", debugStrangeFail
  GiveProcessIDbyLastPacket = 0
  Exit Function
gotErr:
  errmessage = "Function failure : GiveProcessIDbyLastPacket could not match idconnection<->pid : Error number " & CStr(Err.Number) & " : " & Err.Description
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & errmessage
  LogOnFile "errors.txt", errmessage
  GiveProcessIDbyLastPacket = 0
End Function

Public Sub AddGameServer(ByVal ServerName As String, ByVal serverIPport As String, Optional ByVal serverDOMAIN As String = "")
  On Error GoTo gotErr
  ' add item to dictionary
  Dim res As Boolean
  GameServerDictionary.item(ServerName) = serverIPport
  GameServerDictionaryDOMAIN.item(ServerName) = serverDOMAIN
  Exit Sub
gotErr:
  LogOnFile "errors.txt", "Get error at AddGameServer : " & Err.Description
End Sub

Public Function GetGameServerPort(ByVal ServerName As String) As Long
  Dim allthing As String
  Dim pos As Integer
  Dim res As Long
  Dim tmps As String
  allthing = GetIPandPortfromServerName(ServerName)
  pos = InStr(1, allthing, ":", vbTextCompare)
  If pos = 0 Then
    res = 0
  Else
    tmps = Right$(allthing, Len(allthing) - pos)
    If (TibiaVersionLong >= 1100) Then
        If (GetGameServerDOMAIN(ServerName, True) = "127.0.0.1") Then
            res = 7171
        Else
            res = CLng(tmps)
        End If
        GetGameServerPort = res
        Exit Function
    Else
        res = CLng(tmps)
    End If
    res = CLng(tmps)
  End If
  GetGameServerPort = res
End Function

Public Function GetGameServerDOMAIN(ByVal ServerName As String, Optional ByVal getHiddenValue As Boolean = False) As String
  On Error GoTo gotErr
  ' get the IPandport from server name
  Dim aRes As String
  Dim res As Boolean
  Dim strBuildIt As String
  Dim b(3) As Byte
  Dim i As Long
  Dim lastI As Long
  Dim strTmp As String
  Dim pos1 As Long
  Dim pos2 As Long
  Dim pos3 As Long
  Dim resS As String
  If GameServerDictionary.Exists(ServerName) = True Then
    resS = GameServerDictionaryDOMAIN.item(ServerName)
    If (TibiaVersionLong >= 1100) Then
        If resS = "127.0.0.1" Then
            If (getHiddenValue) Then
                GetGameServerDOMAIN = resS
            Else
                resS = LCase(ServerName) & "-lb.ciproxy.com"
                Debug.Print "WARNING: Had to use emergency translation: " & ServerName & "=" & resS
                GetGameServerDOMAIN = resS
            End If
        End If
        GetGameServerDOMAIN = resS
        Exit Function
    Else
        GetGameServerDOMAIN = resS
    End If
  Else
    GetGameServerDOMAIN = ""
  End If
  Exit Function
gotErr:
  LogOnFile "errors.txt", "Got error at GetGameServerDOMAIN (" & ServerName & " ): " & Err.Description
  GetGameServerDOMAIN = ""
End Function
Public Function GetIPandPortfromServerName(ByVal ServerName As String) As String
  On Error GoTo gotErr
  ' get the IPandport from server name
  Dim aRes As String
  Dim res As Boolean
  Dim strBuildIt As String
  Dim b(3) As Byte
  Dim i As Long
  Dim lastI As Long
  Dim strTmp As String
  Dim pos1 As Long
  Dim pos2 As Long
  Dim pos3 As Long
  If GameServerDictionary.Exists(ServerName) = True Then
    GetIPandPortfromServerName = GameServerDictionary.item(ServerName)
  Else
    strTmp = GetIPofTibiaServer(ServerName)
    lastI = Len(strTmp)
    ' search the 3 points of the IP
    pos1 = InStr(1, strTmp, ".")
    If pos1 > 0 Then
        pos2 = InStr(pos1 + 1, strTmp, ".")
    Else
        GetIPandPortfromServerName = ""
        Exit Function
    End If
    If pos2 > 0 Then
        pos3 = InStr(pos2 + 1, strTmp, ".")
    Else
        GetIPandPortfromServerName = ""
        Exit Function
    End If
    b(0) = CByte(CLng(Left$(strTmp, pos1 - 1)))
    b(1) = CByte(CLng(Mid$(strTmp, pos1 + 1, pos2 - pos1 - 1)))
    b(2) = CByte(CLng(Mid$(strTmp, pos2 + 1, pos3 - pos2 - 1)))
    b(3) = CByte(CLng(Right$(strTmp, lastI - pos3)))
    strBuildIt = fixThreeDigits(b(0)) & "." & fixThreeDigits(b(1)) & "." & _
     fixThreeDigits(b(2)) & "." & fixThreeDigits(b(3)) & ":7171"
    GetIPandPortfromServerName = strBuildIt
  End If
  Exit Function
gotErr:
  LogOnFile "errors.txt", "Got error at GetIPandPortfromServerName (" & ServerName & " ): " & Err.Description
  GetIPandPortfromServerName = ""
End Function


'Public Function GetProcessIdByAccount(strAccount As String) As Long
'    Dim res As Long
'    res = GetProcessIdFromAccount(strAccount)
'    If res <= 0 Then
'        res = -1
'    End If
'    GetProcessIdByAccount = res
'End Function

Public Function GetProcessIdByManualDebug() As Long
   Dim tibiaclient As Long
   Dim bc As Byte
   Dim c As Long
   c = 0
   tibiaclient = 0
   
   Do
    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      Exit Do
    Else
        Debug.Print "#" & CStr(c) & " : " & CStr(tibiaclient)
        If manualDebugOrder = c Then
            GetProcessIdByManualDebug = tibiaclient
            Exit Function
        End If
        c = c + 1
    End If
  Loop

  GetProcessIdByManualDebug = -1
End Function

Public Function GetProcessIdByAdrConnected() As Long
   If TibiaVersionLong >= 1100 Then
     GetProcessIdByAdrConnected = GetProcessIdByAdrConnected_TibiaQ()
     Exit Function
   End If
   Dim tibiaclient As Long
   Dim bc As Byte
   Dim foundCount As Long
   Dim lastfound As Long
   Dim totalclients As Long
   Dim cantbeother As Long
   Dim cantbeotherBYTE As Byte
   foundCount = 0
   totalclients = 0
   cantbeother = 0
   Do
    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      If foundCount = 1 Then
        GetProcessIdByAdrConnected = lastfound
      ElseIf foundCount = 0 Then
        If totalclients = 1 Then
            Debug.Print "Warning: only 1 tibiaclient, with connection status " & GoodHex(cantbeotherBYTE)
            GetProcessIdByAdrConnected = cantbeother
        Else
            GetProcessIdByAdrConnected = -1
        End If
      Else
        GetProcessIdByAdrConnected = -2
      End If
      Exit Function
    Else
        totalclients = totalclients + 1
        bc = Memory_ReadLong(adrConnected, tibiaclient)
        If TibiaVersionLong >= 980 Then
            If ((bc = &H5) Or (bc = &H6) Or (bc = &H8) Or (bc = &H9)) Then ' tibia 10.11 = &H)
                lastfound = tibiaclient
                foundCount = foundCount + 1
            End If
        Else
            If ((bc = &H5) Or (bc = &H6)) Then
                lastfound = tibiaclient
                foundCount = foundCount + 1
            End If
        End If
        If totalclients = 1 Then
            cantbeother = tibiaclient
            cantbeotherBYTE = bc
        End If
    End If
  Loop

  GetProcessIdByAdrConnected = -1
End Function
Public Function UpdateCharListFromMemory(idConnection As Integer, maxr As Integer) As Long
  Dim tibiaclient As Long
  Dim curradr As Long
  Dim readSoFar As Long
  Dim partialRead As Long
  Dim strName As String
  Dim strServerName As String
  Dim thestart As Long
  Dim b As Byte
  Dim servIP1 As Byte
  Dim servIP2 As Byte
  Dim servIP3 As Byte
  Dim servIP4 As Byte
  Dim servPort As Long
  Dim strIPandPort As String
  Dim currRead As Long
  Dim maxread As Long
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  If TibiaVersionLong >= 1011 Then
    UpdateCharListFromMemory = UpdateCharListFromMemory3(idConnection, maxr)
    Exit Function
  ElseIf TibiaVersionLong >= 971 Then
    UpdateCharListFromMemory = UpdateCharListFromMemory2(idConnection, maxr)
    Exit Function
  End If
  currRead = 0
  maxread = CLng(maxr) + 1
  tibiaclient = ProcessID(idConnection)
  ResetCharList2 idConnection
  thestart = Memory_ReadLong(adrCharListPtr, tibiaclient)
  curradr = thestart
  readSoFar = 0
continueIt:
  strName = ""
  partialRead = 0
  Do
    b = Memory_ReadByte(curradr, tibiaclient, True)
    If b = &H0 Then
      Exit Do
    Else
      strName = strName & Chr(b)
      readSoFar = readSoFar + 1
      partialRead = partialRead + 1
      curradr = curradr + 1
    End If
    If readSoFar > 10000 Then
      UpdateCharListFromMemory = -1
      Exit Function
    End If
    If partialRead = MAXCHARACTERLEN Then
      Exit Do
    End If
  Loop
  curradr = curradr + MAXCHARACTERLEN - partialRead
  strServerName = ""
  partialRead = 0
  Do
    b = Memory_ReadByte(curradr, tibiaclient, True)
    If b = &H0 Then
      Exit Do
    Else
      strServerName = strServerName & Chr(b)
      readSoFar = readSoFar + 1
      partialRead = partialRead + 1
      curradr = curradr + 1
    End If
    If readSoFar > 10000 Then
      UpdateCharListFromMemory = -1
      Exit Function
    End If
    If partialRead = MAXCHARACTERLEN Then
      Exit Do
    End If
  Loop
  If strServerName = "" Or (currRead >= maxread) Then
    If currRead >= maxread Then
      UpdateCharListFromMemory = 0
    Else
      LogOnFile "errors.txt", "Got error at UpdateCharListFromMemory : could not read all (" & CStr(currRead) & "/" & CStr(maxread) & ")"
      UpdateCharListFromMemory = -1
    End If
    Exit Function
  Else
    Debug.Print strName & " : " & strServerName
    curradr = curradr + 54 - partialRead
    strIPandPort = GetIPandPortfromServerName(strServerName)
    If strIPandPort = "" Then
      LogOnFile "errors.txt", "Got error at UpdateCharListFromMemory : can't get IP of server '" & strServerName & "'"
      UpdateCharListFromMemory = -1
      Exit Function
    End If
    servIP1 = CByte(CLng(Mid$(strIPandPort, 1, 3)))
    servIP2 = CByte(CLng(Mid$(strIPandPort, 5, 3)))
    servIP3 = CByte(CLng(Mid$(strIPandPort, 9, 3)))
    servIP4 = CByte(CLng(Mid$(strIPandPort, 13, 3)))
    servPort = CLng(Right$(strIPandPort, Len(strIPandPort) - 16))
    AddCharServer2 idConnection, strName, strServerName, servIP1, servIP2, servIP3, servIP4, servPort
    currRead = currRead + 1
  End If
  b = Memory_ReadByte(curradr, tibiaclient, True)
  If b = &H0 Then
    UpdateCharListFromMemory = 0
    Exit Function
  Else
    GoTo continueIt
  End If
  Exit Function
gotErr:
  LogOnFile "errors.txt", "Got error at UpdateCharListFromMemory : " & Err.Description
  UpdateCharListFromMemory = -1
End Function


Public Function UpdateCharListFromMemory3(idConnection As Integer, maxr As Integer, Optional ByVal tibiaclient As Long = -1) As Long
  Dim i As Long
  Dim address As Long
  Dim address_end As Long
  Dim nick As String
  Dim world As String
  Dim entry_type As Long
  Dim remoteadr As Long
  Dim namesize As Long
  Dim nametype As Long
  Dim dist_name_size As Long
  Dim dist_name_type As Long
  Dim dist_name As Long
  Dim dist_world As Long
  Dim selectedNick As String
  Dim charlist_dist As Long
  'Dim tibiaclient As Long
  Dim ClientSelectedCharId As Long
  Dim currentID As Long
  Dim rememberPort As Long
  Dim rememberDomain As String
  Dim listind As Integer
  'Dim idebug As Integer
  If TibiaVersionLong < 1050 Then
    dist_name_size = 20
    dist_name_type = 24
    dist_name = 4
    dist_world = 32
    charlist_dist = 120
  Else
    dist_name_size = 20
    dist_name_type = 24
    dist_name = 4
    dist_world = 28
    charlist_dist = 104
  End If
  If tibiaclient = -1 Then
  tibiaclient = ProcessID(idConnection)
  End If
  ClientSelectedCharId = ReadCurrentAddress(tibiaclient, adrSelectedCharIndex, -1, True)
  If ClientSelectedCharId = -1 Then
    UpdateCharListFromMemory3 = -1
    Exit Function
  End If
'tryagain:

 ' idebug = 0
  
  
  listind = 0
  address = Memory_ReadLong(adrCharListPtr, tibiaclient)
  address_end = Memory_ReadLong(adrCharListPtrEND, tibiaclient)
  currentID = 0
  i = address
  ResetCharList2 idConnection
  Do
    'Debug.Print "pointer " & CStr(listind) & "=" & CStr(i)
    entry_type = Memory_ReadLong(i, tibiaclient, True)
    namesize = Memory_ReadLong(i + dist_name_size, tibiaclient, True)
    If ((namesize = 0) Or (entry_type = 0)) Then
              'Debug.Print "fallo en dist = " & charlist_dist
        '  charlist_dist = charlist_dist + 4
         ' GoTo tryagain
      Exit Do
    Else
      nametype = Memory_ReadLong(i + dist_name_type, tibiaclient, True)
      If nametype = 15 Then
        nick = readMemoryString(tibiaclient, i + dist_name, , True)
      Else
        remoteadr = Memory_ReadLong(i + dist_name, tibiaclient, True)
        nick = readMemoryString(tibiaclient, remoteadr, , True)
      End If
      'idebug = idebug + 1
      'If idebug = 2 Then
       ' If nick = "Folfah Jajota" Then
       '   Debug.Print "exito en dist = " & charlist_dist
      '    Exit Function
      '  Else
       '   Debug.Print "fallo en dist = " & charlist_dist
       '   charlist_dist = charlist_dist + 4
      '    GoTo tryagain
       ' End If
       ' If charlist_dist > 300 Then
       '  Exit Function
       ' End If
     ' End If
 
      world = readMemoryString(tibiaclient, i + dist_world, , True)
     
      If (currentID = ClientSelectedCharId) Then
         selectedNick = nick
       ' Debug.Print "Nick=" & nick & " ; Server=" & world & " * SELECTED"
      Else
       ' Debug.Print "Nick=" & nick & " ; Server=" & world
      End If
      rememberPort = GetGameServerPort(world)
      rememberDomain = GetGameServerDOMAIN(world)
      AddCharServer2 idConnection, nick, world, "127", "0", "0", "1", rememberPort, rememberDomain
    End If
    i = i + charlist_dist
    listind = listind + 1
    currentID = currentID + 1
  Loop Until i >= address_end
  'Debug.Print "Read completed"
  UpdateCharListFromMemory3 = 0
End Function



Public Function UpdateCharListFromMemory2(idConnection As Integer, maxr As Integer) As Long
  Dim tibiaclient As Long
  Dim curradr As Long
  Dim readSoFar As Long
  Dim partialRead As Long
  Dim strName As String
  Dim strServerName As String
  Dim thestart As Long
  Dim b As Byte
  Dim servIP1 As Byte
  Dim servIP2 As Byte
  Dim servIP3 As Byte
  Dim servIP4 As Byte
  Dim servPort As Long
  Dim strIPandPort As String
  Dim currRead As Long
  Dim maxread As Long
  Dim MAXCHARACTERLEN2 As Long
  Dim charCount As Long
  Dim namesize As Long
  Dim nametype As Byte
  Dim i As Long
  Dim remoteAddress As Long
  Dim badd(3) As Byte
  
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  'Debug.Print "You should not use this function since Tibia 9.71"
  'UpdateCharListFromMemory2 = -1
  'Exit Function
  
  ' the new list does not display clear names but it should be enough for our small purpose...
  MAXCHARACTERLEN2 = 28
  currRead = 0
  maxread = CLng(maxr) + 1
  tibiaclient = ProcessID(idConnection)
  ResetCharList2 idConnection
  thestart = Memory_ReadLong(adrCharListPtr, tibiaclient)
  curradr = thestart
  readSoFar = 0
  charCount = 0
continueIt:
  strName = ""
  partialRead = 0
  curradr = curradr + 4 ' skip 4 strange bytes

  namesize = Memory_ReadByte(curradr + 16, tibiaclient, True)
  nametype = Memory_ReadByte(curradr + 20, tibiaclient, True)
  If nametype = &HF Then
    Do
      b = Memory_ReadByte(curradr, tibiaclient, True)
      If b = &H0 Then
        Exit Do
      Else
        strName = strName & Chr(b)
        readSoFar = readSoFar + 1
        partialRead = partialRead + 1
        curradr = curradr + 1
      End If
      If readSoFar > 10000 Then
        UpdateCharListFromMemory2 = -1
        Exit Function
      End If
      If partialRead = MAXCHARACTERLEN2 Then
        Exit Do
      End If
    Loop
  Else
'    badd(0) = Memory_ReadByte(curradr, tibiaclient, True)
'    badd(1) = Memory_ReadByte(curradr + 1, tibiaclient, True)
'    badd(2) = Memory_ReadByte(curradr + 2, tibiaclient, True)
'    badd(3) = Memory_ReadByte(curradr + 3, tibiaclient, True)
    remoteAddress = Memory_ReadLong(curradr, tibiaclient, True)
'    Debug.Print "character name stored in remote address = " & remoteAddress
'    Debug.Print GoodHex(badd(0)) & " " & GoodHex(badd(1)) & " " & GoodHex(badd(2)) & " " & GoodHex(badd(3))
    partialRead = 0
    strName = readMemoryString(tibiaclient, remoteAddress, CLng(namesize), True)
  End If
  charCount = charCount + 1
  'strname = "#" & CStr(charCount)
 ' Debug.Print "Got name type " & GoodHex(nametype) & " :" & strName
  curradr = curradr + MAXCHARACTERLEN2 - partialRead
  strServerName = ""
  partialRead = 0
  Do
    b = Memory_ReadByte(curradr, tibiaclient, True)
    If b = &H0 Then
      Exit Do
    Else
      strServerName = strServerName & Chr(b)
      readSoFar = readSoFar + 1
      partialRead = partialRead + 1
      curradr = curradr + 1
    End If
    If readSoFar > 10000 Then
      UpdateCharListFromMemory2 = -1
      Exit Function
    End If
    If partialRead = MAXCHARACTERLEN2 Then
      Exit Do
    End If
  Loop
  
  'Debug.Print "Got server:" & strServerName
  If strServerName = "" Or (currRead >= maxread) Then
    If currRead >= maxread Then
      UpdateCharListFromMemory2 = 0
    Else
      LogOnFile "errors.txt", "Got error at UpdateCharListFromMemory2 : could not read all (" & CStr(currRead) & "/" & CStr(maxread) & ")"
      UpdateCharListFromMemory2 = -1
    End If
    Exit Function
  Else
    'Debug.Print "size=" & CLng(namesize) & " type= " & GoodHex(nametype); " : " & strname & " : " & strServerName
    curradr = curradr + 40 - partialRead
    strIPandPort = GetIPandPortfromServerName(strServerName)
    If strIPandPort = "" Then
      LogOnFile "errors.txt", "Got error at UpdateCharListFromMemory2 : can't get IP of server '" & strServerName & "'"
      UpdateCharListFromMemory2 = -1
      Exit Function
    End If
    servIP1 = CByte(CLng(Mid$(strIPandPort, 1, 3)))
    servIP2 = CByte(CLng(Mid$(strIPandPort, 5, 3)))
    servIP3 = CByte(CLng(Mid$(strIPandPort, 9, 3)))
    servIP4 = CByte(CLng(Mid$(strIPandPort, 13, 3)))
    servPort = CLng(Right$(strIPandPort, Len(strIPandPort) - 16))
    AddCharServer2 idConnection, strName, strServerName, servIP1, servIP2, servIP3, servIP4, servPort
    currRead = currRead + 1
  End If
  'b = Memory_ReadByte(curradr + 4, tibiaclient, True)
  If currRead >= maxread Then
    UpdateCharListFromMemory2 = 0
    Exit Function
  Else
    GoTo continueIt
  End If
  Exit Function
gotErr:
  LogOnFile "errors.txt", "Got error at UpdateCharListFromMemory2 : " & Err.Description
  UpdateCharListFromMemory2 = -1
End Function





Public Function GetProcessIDfromCharList2(ByVal idConnection As Long) As Long
    '...
    ' PENDIENTE DE PROGRAMAR
   Dim strName As String
   Dim readSoFar As Long
   Dim partialRead As Long
   Dim curradr As Long
   Dim thestart As Long
   Dim strName2 As String
   Dim tibiaclient As Long
   Dim b As Byte
   If idConnection = 0 Then
    GetProcessIDfromCharList2 = -1
    Exit Function
   End If
   Do

    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      Exit Do
    Else
    
    
        thestart = Memory_ReadLong(adrCharListPtr, tibiaclient)
        curradr = thestart
        strName = ""
        partialRead = 0
        Do
            b = Memory_ReadByte(curradr, tibiaclient, True)
            If b = &H0 Then
                Exit Do
            Else
                strName = strName & Chr(b)
                readSoFar = readSoFar + 1
                partialRead = partialRead + 1
                curradr = curradr + 1
            End If
            If readSoFar > 10000 Then
                Exit Do
            Else
                If partialRead = MAXCHARACTERLEN Then
                    Exit Do
                End If
            End If
        Loop
        strName2 = CharacterList2(idConnection).item(0).CharacterName
        If strName = strName2 Then
            GetProcessIDfromCharList2 = tibiaclient
            Exit Function
        End If
    End If
  Loop
    
    
    GetProcessIDfromCharList2 = -1
End Function

Public Function fixThreeDigits(n As Byte) As String
  Dim res As String
  res = ""
  If (n < 100) Then
    res = "0"
  End If
  If (n < 10) Then
    res = res & "0"
  End If
  res = res & CStr(CInt(n))
  fixThreeDigits = res
End Function



