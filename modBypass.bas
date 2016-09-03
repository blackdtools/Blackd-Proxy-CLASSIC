Attribute VB_Name = "modBypass"
#Const FinalMode = 1
Option Explicit
Dim ServerIP() As String
Dim ServerName() As String
Dim LastServerID As Long
Public bypass_def1 As Long
Public bypass_def2 As String
Public bypass_def3 As String

Public AlternativeBinding As Long
'...

Public Sub LoadServerIps()
    On Error GoTo gotErr
    Dim fso As scripting.FileSystemObject
    Dim fn As Integer
    Dim strLine As String
    Dim filename As String
    Dim lngPosSpace As Long
    Dim readingLine As Long
    Dim strServer As String
    Dim strIP As String
    readingLine = 0
    LastServerID = -1
    frmAdvanced.cmbTibiaServers.Clear
    Set fso = New scripting.FileSystemObject
    filename = App.path & "\ips\ips.txt"
    If fso.FileExists(filename) = True Then
        fn = FreeFile
        Open filename For Input As #fn
        While Not EOF(fn)
            Line Input #fn, strLine
            readingLine = readingLine + 1
            If strLine <> "" Then
                lngPosSpace = InStr(1, strLine, " ")
                If lngPosSpace = 0 Then
                    MsgBox "Bad format detected on ips.txt , line " & CStr(readingLine), vbOKOnly + vbCritical, "Load error"
                    End
                End If
                LastServerID = LastServerID + 1
                strServer = Left$(strLine, lngPosSpace - 1)
                strIP = Right$(strLine, Len(strLine) - lngPosSpace)
                frmAdvanced.cmbTibiaServers.AddItem strServer
                ReDim Preserve ServerName(LastServerID)
                ReDim Preserve ServerIP(LastServerID)
                ServerName(LastServerID) = strServer
                ServerIP(LastServerID) = strIP
            End If
        Wend
        Close #fn
        frmAdvanced.cmbTibiaServers.ListIndex = 0
    Else
        MsgBox "ips.txt is missing", vbOKOnly + vbCritical, "Load error"
        End
    End If
    Exit Sub
gotErr:
    MsgBox "Can't load ips.txt" & vbCrLf & "Got error code " & CStr(Err.Number) & vbCrLf & _
    "Error description: " & Err.Description, vbOKOnly + vbCritical, "Load error"
End Sub



Public Function sendStringAtStageOne(idConnection As Integer, str As String) As Long
  On Error GoTo gotErr
  Dim strSending As String
  Dim ub As Long
  Dim lopa As Long
  Dim aRes As Long
  Dim cheatpacket() As Byte
  Dim toServer As String
  Dim safeMode As String
  safeMode = True
  toServer = False
  strSending = str
  If safeMode = True Then
    strSending = "00 00 " & strSending
  End If
  If GetCheatPacket(cheatpacket, strSending) = -1 Then

    sendStringAtStageOne = -1
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
    sendStringAtStageOne = -1
    Exit Function
  End If
  lopa = GetTheLong(cheatpacket(0), cheatpacket(1))
  If (lopa <> (ub - 1)) Then
    sendStringAtStageOne = -1
    Exit Function
  End If
  If (Connected(idConnection) = True) Then
    If toServer = False Then
      ' send the packet to client
      'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " > SENDING :" & frmMain.showAsStr2(cheatpacket, 2)
      frmMain.UnifiedSendToClient idConnection, cheatpacket
    Else
      ' send the packet to server
      'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & " > SENDING :" & frmMain.showAsStr2(cheatpacket, 2)
     ' frmMain.UnifiedSendToServerGame idConnection, cheatpacket, True
    End If
    DoEvents
    sendStringAtStageOne = 0
    Exit Function
  Else
    sendStringAtStageOne = -1
    Exit Function
  End If
  Exit Function
gotErr:
  sendStringAtStageOne = -1
End Function

Public Function GetIPofTibiaServer(ByVal strServerName As String) As String
    Dim i As Long
    For i = 0 To LastServerID
        If strServerName = ServerName(i) Then
            GetIPofTibiaServer = ServerIP(i)
            Exit Function
        End If
    Next i
    GetIPofTibiaServer = strServerName
End Function

Public Function BypassLoginServer(Index As Integer)
    Dim strCheat As String
    strCheat = "14 " & Hexarize2("317" & vbCrLf & "Welcome to Blackd Proxy!") & "64 01 "
    strCheat = strCheat & Hexarize2(frmAdvanced.txtLoginCharacter.Text) & "" ' char name
    strCheat = strCheat & Hexarize2(frmAdvanced.cmbTibiaServers.Text) & "" ' gameserver name
    strCheat = strCheat & "7F 00 00 01 " '127.0.0.1
    strCheat = strCheat & FiveChrLon(frmMain.txtClientGameP.Text) & " " ' blackd proxy gameserver port
    strCheat = strCheat & "E8 03" ' 1000 days of premium account mirage (not real)
    sendStringAtStageOne Index, strCheat
    DoEvents
End Function
