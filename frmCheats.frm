VERSION 5.00
Begin VB.Form frmCheats 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tools ..."
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6930
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmCheats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLevel 
      Height          =   285
      Left            =   5880
      TabIndex        =   32
      Text            =   "255"
      Top             =   0
      Width           =   615
   End
   Begin VB.CheckBox chkAutoHead 
      BackColor       =   &H00000000&
      Caption         =   "Auto - header"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   5880
      TabIndex        =   31
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdGmMessage 
      BackColor       =   &H00C0FFFF&
      Caption         =   "GM MSG"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Debug packet"
      Height          =   375
      Left            =   3480
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdSpecial 
      BackColor       =   &H0000FF00&
      Caption         =   "DEBUG"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdSys 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SYS MSG"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   360
      Width           =   975
   End
   Begin VB.CheckBox chkInspectTileID 
      BackColor       =   &H00000000&
      Caption         =   "Inspect tileIDs ingame"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   5520
      Width           =   2775
   End
   Begin VB.TextBox txtTile 
      Height          =   375
      Left            =   120
      MaxLength       =   5
      TabIndex        =   10
      Text            =   "64 00"
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdTileInfo 
      BackColor       =   &H0080C0FF&
      Caption         =   "Get tile info (this id)"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CLEAR RESULTS"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox txtPackets 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   3960
      Width           =   6615
   End
   Begin VB.CommandButton cmdToAscii 
      BackColor       =   &H0080C0FF&
      Caption         =   "TO ASCII..."
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdToHex 
      BackColor       =   &H0080C0FF&
      Caption         =   "TO HEX ..."
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdCountBytes 
      BackColor       =   &H0080C0FF&
      Caption         =   "COUNT HEX..."
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdOpenBoard 
      BackColor       =   &H0080C0FF&
      Caption         =   "OPEN BOARD"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Send to server"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   2520
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Send to client"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   2520
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox txtSendHexID 
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Text            =   "1"
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton cmdSendHex 
      BackColor       =   &H0080C0FF&
      Caption         =   "SEND HEX ..."
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtLogoutID 
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Text            =   "1"
      Top             =   1530
      Width           =   495
   End
   Begin VB.CommandButton cmdLogout 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SEND LOGOUT"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtFrom 
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Text            =   "GM Guido"
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtClientID 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Text            =   "1"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtSend 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Text            =   "Hello!"
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton cmdSendMsg 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SEND MSG"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "fake level:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   33
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblDebug 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4560
      TabIndex        =   25
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblHexTools 
      BackColor       =   &H00000000&
      Caption         =   "Hex tools:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   6840
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblSendHexID 
      BackColor       =   &H00000000&
      Caption         =   " ID :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   23
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblCheatTest 
      BackColor       =   &H00000000&
      Caption         =   "Send any array of bytes to client or server. Format example: EB AF 08 00 30 3E 0B"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Width           =   6375
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   6840
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6840
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblFromLogout 
      BackColor       =   &H00000000&
      Caption         =   "to this server ID :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   21
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblFakeMessage 
      BackColor       =   &H00000000&
      Caption         =   "Fake messages to self:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblFrom 
      BackColor       =   &H00000000&
      Caption         =   "from this character (no need it even exists)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   19
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblID 
      BackColor       =   &H00000000&
      Caption         =   "to client ID :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frmCheats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 0
Option Explicit

Private Sub cmdCountBytes_Click()
  Dim strinput As String
  Dim strByte As String
  Dim res As String
  Dim countB As Long
  Dim lb As Byte
  Dim hb As Byte
  countB = 0
  #If FinalMode Then
  On Error GoTo badF
  #End If
  ClosedBoard = False
  frmBigText.lblText = "Enter a hex chain. For example: AE 8C 04 45 0F" & vbCrLf & _
  "Note that you should enter the left zeros too!"
  frmBigText.Show
  DisableBoardButtons
  While ClosedBoard = False
    DoEvents
  Wend
  EnableBoardButtons
  strinput = frmBigText.txtBoard.Text
  If CanceledBoard = False Then
  res = ""
  While Len(strinput) > 0
    If Left(strinput, 1) = " " Or Left(strinput, 1) = vbCr Or Left(strinput, 1) = vbLf Then
      strinput = Right(strinput, Len(strinput) - 1)
    Else
      strByte = Left(strinput, 2)
      strinput = Right(strinput, Len(strinput) - 2)
      res = res & ConvToAscii(strByte)
      countB = countB + 1
    End If
  Wend
  lb = LowByteOfLong(countB)
  hb = HighByteOfLong(countB)
  txtPackets.Text = txtPackets.Text & vbCrLf & _
  "COUNT RESULT: Dec " & CStr(countB) & " Hex " & _
  GoodHex(lb) & " " & _
  GoodHex(hb)
  End If
  Exit Sub
badF:
  MsgBox "Bad format"
End Sub

Private Sub cmdGmMessage_Click()
  Dim aRes As Long
  If CInt(txtClientID.Text) > 0 Then
    aRes = GiveGMmessage(CInt(txtClientID.Text), txtSend.Text, txtFrom.Text)
    DoEvents
  End If
End Sub

Private Sub cmdLogout_Click()
  Dim cheatpacket() As Byte
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  If GetCheatPacket(cheatpacket, "01 00 14") = -1 Then
    MsgBox "Bad format"
    Exit Sub
  End If
  If GameConnected(CInt(txtClientID.Text)) = False Then
    MsgBox txtClientID.Text & " is not a valid ID"
    Exit Sub
  End If
  ' send the packet
  frmMain.UnifiedSendToServerGame CInt(txtLogoutID.Text), cheatpacket, True
  Exit Sub
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Got unexpected error at cmdLogout_Click"
End Sub



Private Sub cmdOpenBoard_Click()
  ClosedBoard = False
  frmBigText.lblText = "Open board - Write anything you like here"
  frmBigText.Show
  DisableBoardButtons
  While ClosedBoard = False
    DoEvents
  Wend
  EnableBoardButtons
End Sub

Private Sub cmdOpenTrueRadar_Click()

End Sub

Private Sub cmdSendHex_Click()
  Dim i As Integer
  Dim aRes As Long
  Dim withsafe As Boolean
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  If chkAutoHead.Value = 1 Then
    withsafe = True
  Else
    withsafe = False
  End If
  ClosedBoard = False
  frmBigText.lblText = "Enter a hex chain. For example: AE 8C 04 45 0F" & vbCrLf & _
  "Note that you should enter the left zeros too!"
  frmBigText.Show
  DisableBoardButtons
  While ClosedBoard = False
    DoEvents
  Wend
  EnableBoardButtons
  If CanceledBoard = False Then
  If GameConnected(CInt(txtSendHexID.Text)) = True Then
    If Option1.Value = True Then
      ' send the packet to client
      aRes = sendString(CInt(txtSendHexID.Text), frmBigText.txtBoard.Text, False, withsafe)
    Else
      ' send the packet to server
      aRes = sendString(CInt(txtSendHexID.Text), frmBigText.txtBoard.Text, True, withsafe)
    End If
  Else
    MsgBox txtSendHexID.Text & " is not a valid ID"
  End If
  End If
  Exit Sub
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Got unexpected error at cmdSendHex_Click"
End Sub

Private Sub cmdSendMsg_Click()
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
  Dim aRes As Long
  Dim chCad As String
  If GameConnected(CInt(txtClientID.Text)) = False Then
    MsgBox txtClientID.Text & " is not a valid ID"
    Exit Sub
  End If
  aRes = SendMessageToClient(CInt(txtClientID.Text), txtSend.Text, txtFrom.Text)
End Sub

Private Sub cmdSpecial_Click()
  ' send packet to main flow (for error.txt debuging)
  Dim cheatpacket() As Byte
  Dim i As Integer
  Dim iRes As Integer
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  ClosedBoard = False
  frmBigText.lblText = "Enter a hex chain. For example: AE 8C 04 45 0F" & vbCrLf & _
  "Note that you should enter the left zeros too!"
  frmBigText.Show
  DisableBoardButtons
  While ClosedBoard = False
    DoEvents
  Wend
  EnableBoardButtons
  If CanceledBoard = False Then
  If GetCheatPacket(cheatpacket, frmBigText.txtBoard.Text) = -1 Then
    MsgBox "Bad format"
    Exit Sub
  End If
  If GameConnected(CInt(txtClientID.Text)) = True Then
    forcedDebugChain = True
    iRes = LearnFromServer(cheatpacket, CInt(txtClientID.Text))
  Else
    MsgBox txtClientID.Text & " is not a valid ID"
  End If
  End If
  Exit Sub
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Got unexpected error at debug"
End Sub

Private Sub cmdSys_Click()
  If CInt(txtClientID.Text) > 0 Then
    SendSystemMessageToClient CInt(txtClientID.Text), txtSend.Text
  End If
End Sub

Private Sub cmdTileInfo_Click()
  Dim lon As Long
  Dim b1 As Byte
  Dim b2 As Byte
  Dim s1 As Byte
  Dim s2 As Byte
  Dim pos As Long
  Dim strRes As String
  #If FinalMode Then
  On Error GoTo exitS
  #End If
  s1 = FromHexToDec(Mid(txtTile.Text, 1, 1))
  s2 = FromHexToDec(Mid(txtTile.Text, 2, 1))
  b1 = (s1 * 16) + s2
  s1 = FromHexToDec(Mid(txtTile.Text, 4, 1))
  s2 = FromHexToDec(Mid(txtTile.Text, 5, 1))
  b2 = (s1 * 16) + s2
  txtPackets.Text = txtPackets.Text & vbCrLf & GetTileInfoString(b1, b2)
  Exit Sub
exitS:
  MsgBox "Bad format"
End Sub

Private Sub cmdToAscii_Click()
  Dim strinput As String
  Dim strByte As String
  Dim res As String
  #If FinalMode Then
  On Error GoTo badF
  #End If
  ClosedBoard = False
  frmBigText.lblText = "Enter a hex chain. For example: AE 8C 04 45 0F" & vbCrLf & _
  "Note that you should enter the left zeros too!"
  frmBigText.Show
  DisableBoardButtons
  While ClosedBoard = False
    DoEvents
  Wend
  EnableBoardButtons
  strinput = frmBigText.txtBoard.Text
  If CanceledBoard = False Then
  res = ""
  While Len(strinput) > 0
    If Left(strinput, 1) = " " Or Left(strinput, 1) = vbCr Or Left(strinput, 1) = vbLf Then
      strinput = Right(strinput, Len(strinput) - 1)
    Else
      strByte = Left(strinput, 2)
      strinput = Right(strinput, Len(strinput) - 2)
      res = res & ConvToAscii(strByte)
    End If
  Wend
  txtPackets.Text = txtPackets.Text & vbCrLf & "CONVERT RESULT: " & res
  End If
  Exit Sub
badF:
  MsgBox "Bad format"
End Sub

Private Sub cmdToHex_Click()
  Dim strin As String
  Dim strByte As String
  Dim res As String
  #If FinalMode Then
  On Error GoTo badF
  #End If
  ClosedBoard = False
  frmBigText.lblText = "Enter a ascii string. For example: hello!"
  frmBigText.Show
  DisableBoardButtons
  While ClosedBoard = False
    DoEvents
  Wend
  EnableBoardButtons
  strin = frmBigText.txtBoard.Text
  If CanceledBoard = False Then
  res = Hexarize(strin)
  txtPackets.Text = txtPackets.Text & vbCrLf & "CONVERT RESULT: " & res
  End If
  Exit Sub
badF:
  MsgBox "Bad format"
End Sub

Private Sub Command1_Click()
txtPackets.Text = ""
End Sub







Private Sub Command2_Click()
  Dim cPacket() As Byte
  Dim sCheat As String
  Dim iRes As Integer
  sCheat = frmBigText.txtBoard.Text
  If sCheat <> "" Then
    iRes = GetCheatPacket(cPacket, sCheat)
    iRes = LearnFromServer(cPacket, 1)
    frmMain.UnifiedSendToClientGame 1, cPacket
  End If
End Sub

Private Sub Form_Load()
#If FinalMode = 0 Then
Command2.Visible = True
cmdSpecial.Visible = True
#End If
    fakemessagesLevel = 255
    fakemessagesLevel1 = LowByteOfLong(255)
    fakemessagesLevel2 = HighByteOfLong(255)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub



Private Sub txtLevel_Change()
    On Error GoTo usedefaultlvl
    Dim lngLvl As Long
    lngLvl = CLng(txtLevel.Text)
    fakemessagesLevel = lngLvl
    fakemessagesLevel1 = LowByteOfLong(lngLvl)
    fakemessagesLevel2 = HighByteOfLong(lngLvl)
    Exit Sub
usedefaultlvl:
    fakemessagesLevel = 255
    fakemessagesLevel1 = LowByteOfLong(255)
    fakemessagesLevel2 = HighByteOfLong(255)
End Sub

Private Sub txtPackets_Change()
  txtPackets.SelStart = Len(txtPackets.Text)
End Sub
