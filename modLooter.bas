Attribute VB_Name = "modLooter"
#Const FinalMode = 1
Option Explicit
Public Const MAXLOOTQUEUE As Long = 19
Public Type TypeLootPoint
    x As Long
    y As Long
    z As Long
    addedTime As Long
    expireGtc As Long
End Type
Public Type TypeLootQueue
    points(0 To 19) As TypeLootPoint
End Type
Public Looter() As TypeLootQueue
Public NextLootStart() As Long
Public MAXTIMEINLOOTQUEUE() As Long
Public MINDELAYTOLOOT() As Long
Public MAXTIMETOREACHCORPSE() As Long
Public OldLootMode() As Boolean
Public LootAll() As Boolean
Public PKwarnings() As Boolean
Public DoingNewLoot() As Boolean
Public DoingNewLootX() As Long
Public DoingNewLootY() As Long
Public DoingNewLootZ() As Long
Public DoingNewLootMAXGTC() As Long

Public Function ResetLooter(ByVal idConnection As Integer)
    Dim i As Long
    For i = 0 To MAXLOOTQUEUE
        Looter(idConnection).points(i).x = 0
        Looter(idConnection).points(i).y = 0
        Looter(idConnection).points(i).z = 0
        Looter(idConnection).points(i).expireGtc = 0
        Looter(idConnection).points(i).addedTime = 0
    Next
    NextLootStart(idConnection) = 0
    MAXTIMEINLOOTQUEUE(idConnection) = 60000
    If frmMain.TrueServer1.value = True Then
      MINDELAYTOLOOT(idConnection) = 0 ' was changed to 0 in 16.5
    Else
      MINDELAYTOLOOT(idConnection) = 1000 ' ot servers require 1second wait
    End If
    OldLootMode(idConnection) = True
    DoingNewLoot(idConnection) = False
    DoingNewLootX(idConnection) = 0
    DoingNewLootY(idConnection) = 0
    DoingNewLootZ(idConnection) = 0
    MAXTIMETOREACHCORPSE(idConnection) = 10000
    DoingNewLootMAXGTC(idConnection) = 0
End Function

Public Function AddLootPoint(ByVal idConnection As Integer, _
 ByVal x As Long, _
 ByVal y As Long, _
 ByVal z As Long)
 Dim currPoint As Long
 Dim tried As Long
 Dim i As Long
 Dim gtc
 If autoLoot(idConnection) = False Then
    AddLootPoint = -1
    Exit Function
 End If
 If AlreadyAddedInLootQueue(idConnection, x, y, z) = True Then
    AddLootPoint = -1
    Exit Function
 End If
 ' Debug.Print x & "," & y & "," & z
 gtc = GetTickCount()
 tried = 0
 currPoint = 0
 While Looter(idConnection).points(currPoint).expireGtc > gtc
    currPoint = currPoint + 1
    tried = tried + 1
    If currPoint > MAXLOOTQUEUE Then
        currPoint = 0
    End If
    If tried > (MAXLOOTQUEUE + 1) Then
        AddLootPoint = -1
        Exit Function
    End If
 Wend
 Looter(idConnection).points(currPoint).addedTime = gtc
 Looter(idConnection).points(currPoint).expireGtc = gtc + MAXTIMEINLOOTQUEUE(idConnection)
 Looter(idConnection).points(currPoint).x = x
 Looter(idConnection).points(currPoint).y = y
 Looter(idConnection).points(currPoint).z = z
 AddLootPoint = 0
End Function

Public Function ChooseBestLoot(ByVal idConnection As Integer) As Long
    Dim i As Long
    Dim gtc As Long
    Dim bestPoint As Long
    Dim bestDist As Long
    Dim currDist As Long
    
    bestPoint = -1
    bestDist = 20 'max dist
    gtc = GetTickCount()
    For i = 0 To MAXLOOTQUEUE
       If (Looter(idConnection).points(i).expireGtc > gtc) And _
          (Looter(idConnection).points(i).addedTime < _
           (gtc - MINDELAYTOLOOT(idConnection)) _
           ) Then
         currDist = (50 * Abs(Looter(idConnection).points(i).z - myZ(idConnection)))
         currDist = currDist + Abs(Looter(idConnection).points(i).x - myX(idConnection))
         currDist = currDist + Abs(Looter(idConnection).points(i).y - myY(idConnection))
         If currDist < bestDist Then
            bestDist = currDist
            bestPoint = i
         End If
       End If
    Next i
    ChooseBestLoot = bestPoint
End Function


Public Function AlreadyAddedInLootQueue(ByVal idConnection As Integer, ByVal Px, ByVal Py, ByVal Pz) As Boolean
    Dim i As Long
    Dim gtc As Long
    gtc = GetTickCount()
    For i = 0 To MAXLOOTQUEUE
         If ( _
           (Looter(idConnection).points(i).x = Px) And _
           (Looter(idConnection).points(i).y = Py) And _
           (Looter(idConnection).points(i).z = Pz) _
           ) Then
            AlreadyAddedInLootQueue = True
            Exit Function
       End If
    Next i
    AlreadyAddedInLootQueue = False
End Function

Public Function AnythingLootable(ByVal idConnection As Integer) As Boolean
    Dim i As Long
    Dim gtc As Long
    gtc = GetTickCount()
    For i = 0 To MAXLOOTQUEUE
       If (Looter(idConnection).points(i).expireGtc > gtc) And _
          (Looter(idConnection).points(i).addedTime < _
           (gtc - MINDELAYTOLOOT(idConnection)) _
           ) Then
         If Looter(idConnection).points(i).z = myZ(idConnection) Then
            AnythingLootable = True
            Exit Function
         End If
       End If
    Next i
    AnythingLootable = False
End Function

Public Function PrintLootPosition(ByVal idConnection As Integer, _
 lootPos As Long) As String
    Dim res As String
    res = CStr(Looter(idConnection).points(lootPos).x) & "," & _
     CStr(Looter(idConnection).points(lootPos).y) & "," & _
     CStr(Looter(idConnection).points(lootPos).z)
    PrintLootPosition = res
End Function

Public Function PrintLootStats(ByVal idConnection As Integer) As String
    Dim res As String
    Dim i As Long
    Dim valid As Long
    Dim total As Long
    Dim gtc As Long
    gtc = GetTickCount()
    valid = 0
    total = 0
    For i = 0 To MAXLOOTQUEUE
       If (Looter(idConnection).points(i).expireGtc > gtc) Then
        total = total + 1

         If Looter(idConnection).points(i).addedTime < gtc - MINDELAYTOLOOT(idConnection) Then
            valid = valid + 1
          End If
       End If

    Next i
    res = "Lootable things right now: " & CStr(valid) & _
     " ; Total on queue: " & CStr(total)
     PrintLootStats = res
End Function


Public Sub SmartLootCorpse(ByVal idConnection As Integer)
  '0A 00 82 3C 7D AD 7D 08 89 07 01 00
  '0A 00 82 0A 01 73 00 07 9A 10 02 01
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim xdif As Long
  Dim ydif As Long
  Dim b1 As Byte
  Dim b2 As Byte
  Dim SOPT As Byte
  Dim SS As Byte
  Dim lSS As Long
  Dim inRes As Integer
  Dim tileID As Long
  Dim firstAv As Byte
  Dim j As Long
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  If DoingNewLoot(idConnection) = False Then
    Exit Sub
  End If
  xdif = DoingNewLootX(idConnection) - myX(idConnection)
  ydif = DoingNewLootY(idConnection) - myY(idConnection)
  If DoingNewLootZ(idConnection) <> myZ(idConnection) Then
    Exit Sub
  End If
  
  If ((xdif < -7) Or (xdif > 8) Or (ydif < -5) Or (ydif > 6)) Then
    'out of range
    Exit Sub
  End If
  
  
    firstAv = &HFF
    For j = 0 To HIGHEST_BP_ID
      If Backpack(idConnection, j).open = False Then
        firstAv = CByte(j)
        Exit For
      End If
    Next j
    ' 0A 00 82 FF FF 41 00 03 26 0B 03 02
    ' 0A 00 82 FF FF 42 00 01 25 0B 01 03
    If firstAv = &HFF Then
      'aRes = SendLogSystemMessageToClient(idConnection, "Can't open more backpacks!")
      'DoEvents
     '  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Bad backpack"
      Exit Sub
    End If
  
  
  
  
  
  SOPT = &HFF
  For lSS = 1 To 10
    SS = CByte(lSS)
    tileID = GetTheLong(Matrix(ydif, xdif, myZ(idConnection), idConnection).s(SS).t1, Matrix(ydif, xdif, myZ(idConnection), idConnection).s(SS).t2)
    If DatTiles(tileID).alwaysOnTop = True Then
      SOPT = &HFF
    ElseIf DatTiles(tileID).iscontainer = True Then
      SOPT = SS
      Exit For
    Else
      Exit For
    End If
  Next lSS
  If SOPT = &HFF Then
    'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Bad SOPT"
    Exit Sub 'no good corpse
  End If
  b1 = Matrix(ydif, xdif, myZ(idConnection), idConnection).s(SOPT).t1
  b2 = Matrix(ydif, xdif, myZ(idConnection), idConnection).s(SOPT).t2
  sCheat = "0A 00 82 " & FiveChrLon(DoingNewLootX(idConnection)) & " " & FiveChrLon(DoingNewLootY(idConnection)) & " " & GoodHex(CByte(myZ(idConnection))) & _
   " " & GoodHex(b1) & " " & GoodHex(b2) & " " & GoodHex(SS) & " " & GoodHex(firstAv)
   ' debug
  'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & sCheat
  inRes = GetCheatPacket(cPacket, sCheat)
  waitCounter(idConnection) = GetTickCount() + 2000
  frmMain.UnifiedSendToServerGame idConnection, cPacket, True
  DoEvents
  Exit Sub
gotErr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Got error at PerformUseItem"
End Sub
