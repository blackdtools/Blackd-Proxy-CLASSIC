VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "JwldButn2b.ocx"
Begin VB.Form frmBroadcast 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Broadcast"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmBroadcast.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin JwldButn2b.JeweledButton cmdPlay 
      Height          =   255
      Left            =   6120
      TabIndex        =   24
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Caption         =   "Play"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      BorderColor_Hover=   16761024
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdRead 
      Height          =   375
      Left            =   3360
      TabIndex        =   23
      Top             =   1920
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "> read >"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      BorderColor_Hover=   16761024
      BorderColor_Inner=   16777215
   End
   Begin VB.CheckBox chkMC 
      Caption         =   "Multiclient broadcast by turns in same server."
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   4560
      Width           =   3495
   End
   Begin VB.Timer timerBroadcast 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3360
      Top             =   0
   End
   Begin VB.ComboBox cmbCharacter 
      Height          =   315
      Left            =   1200
      TabIndex        =   19
      Text            =   "-"
      ToolTipText     =   "Select one of your connected characters"
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtBroadcastDelay2 
      Height          =   285
      Left            =   2400
      MaxLength       =   7
      TabIndex        =   10
      Text            =   "30000"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtBroadcastDelay1 
      Height          =   285
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   9
      Text            =   "20000"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtBroadcastThis 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   6975
   End
   Begin VB.ListBox lstList 
      Height          =   2790
      Left            =   4200
      TabIndex        =   5
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtRaw 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "http://www.tibia.com/community/?subtopic=worlds&world=Aldora"
      Top             =   7320
      Width           =   6975
   End
   Begin JwldButn2b.JeweledButton cmdStop 
      Height          =   255
      Left            =   5040
      TabIndex        =   25
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Caption         =   "Stop"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      BorderColor_Hover=   16761024
      BorderColor_Inner=   16777215
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   8160
      Width           =   7455
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "0%"
      Height          =   255
      Left            =   6480
      TabIndex        =   20
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "Spammer:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblCurrentBroadcast 
      BackColor       =   &H00FFC0C0&
      Caption         =   "-nobody"
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Currently broadcasting to:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "Example 2: hello, I don't care about your name. This is a test"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Width           =   6975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Example 1: hello my friend $broadcast$ , I am broadcasting this by private messages"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5640
      Width           =   6975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Step 7: Press play to start. Press stop if you want to stop it"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   6360
      Width           =   6375
   End
   Begin VB.Label Label6 
      Caption         =   "ms"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "to"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Message delay:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Message to send:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Step 3: press the button below:"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   3360
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Copy/paste the list of online players here. Don't worry about level and class. Paste all."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   6975
   End
   Begin VB.Label lblText 
      BackColor       =   &H00000000&
      Caption         =   "Copy/paste this in your browser. Change server at the end:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   7080
      Width           =   6615
   End
End
Attribute VB_Name = "frmBroadcast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Private currentBroadcastID As Long

Public Function getCurrentBID() As Long
    getCurrentBID = currentBroadcastID
End Function
Public Function nextValidBID() As Long
    Dim nothingElse As Long
    Dim testIt As Long
    Dim i As Long
    nothingElse = broadcastIDselected
    testIt = nothingElse + 1
    For i = testIt To MAXCLIENTS
        If GameConnected(i) = True Then
            nextValidBID = i
            Exit Function
        End If
    Next i
    For i = 1 To (nothingElse - 1)
        If GameConnected(i) = True Then
            nextValidBID = i
            Exit Function
        End If
    Next i
    nextValidBID = nothingElse
End Function

Private Function filterIt(ByVal strName As String) As String
    Dim lenOfStrname As String
    Dim i As Long
    Dim p As String
    Dim filteredN As String
    filteredN = ""
    lenOfStrname = Len(strName)
    i = 1
    Do
        p = Mid$(strName, i, 1)
        i = i + 1
        If IsNumeric(p) = False Then
            filteredN = filteredN & p
        End If
    Loop Until (IsNumeric(p) = True) Or (i > lenOfStrname)

    filteredN = Trim$(filteredN)
    filteredN = Replace(filteredN, Chr(9), "") ' for firefox explorer
    filterIt = filteredN
End Function
Private Function ReadRaw()
    Dim pos1 As Long
    Dim pos2 As Long
    Dim part As String
    Dim nothingElse As Boolean
    Dim posI As Long
    Dim lenText As Long
    Dim strRaw As String
    Dim a As String
    Dim strName As String
    nothingElse = False
    strRaw = txtRaw.Text & vbCrLf
    lenText = Len(strRaw)
    lstList.Clear
    pos1 = 1
    Do
      pos2 = InStr(pos1, strRaw, vbCr)
      If pos2 > 0 Then
        part = Mid$(strRaw, pos1, pos2 - pos1)
        pos1 = pos2 + 2
        strName = filterIt(part)
        If strName <> "" Then
          lstList.AddItem strName
        End If
      End If
    Loop Until pos2 <= 0
End Function





Private Sub cmbCharacter_Click()
  broadcastIDselected = cmbCharacter.ListIndex
End Sub

Private Sub UpdateTimer()
    Dim lng1 As Long
    Dim lng2 As Long
    If (IsNumeric(txtBroadcastDelay1.Text) = True) And _
     (IsNumeric(txtBroadcastDelay2.Text)) Then
        lng1 = CLng(txtBroadcastDelay1.Text)
        lng2 = CLng(txtBroadcastDelay2.Text)
    Else
        txtBroadcastDelay1.Text = "20000"
        txtBroadcastDelay2.Text = "30000"
        lng1 = 300
        lng2 = 2000
    End If
    If (lng1 < 1) Or (lng2 < 1) Or (lng1 > lng2) Then
        txtBroadcastDelay1.Text = "20000"
        txtBroadcastDelay2.Text = "30000"
        lng1 = 300
        lng2 = 2000
    End If
    timerBroadcast.Interval = randomNumberBetween(lng1, lng2)
End Sub
Private Sub cmdPlay_Click()
    If lstList.ListCount = 0 Then
        lblInfo.Caption = "Destination list is empty. Nothing to broadcast"
    ElseIf txtBroadcastThis.Text = "" Then
        lblInfo.Caption = "First type a text for the broadcast message!"
    ElseIf broadcastIDselected = 0 Then
        lblInfo.Caption = "First select who will send the broadcast messages!"
    ElseIf GameConnected(broadcastIDselected) = False Then
        lblInfo.Caption = "Broadcast cancelled. Selected tibia client is offline"
    Else
        currentBroadcastID = broadcastIDselected - 1
        timerBroadcast.Interval = 300
        timerBroadcast.enabled = True
        lblInfo.Caption = "Broadcast started. Number of players: " & CStr(lstList.ListCount)
    End If
End Sub

Private Sub cmdRead_Click()
ReadRaw
End Sub

Private Sub cmdStop_Click()
    currentBroadcastIndex = -1
    Me.timerBroadcast.enabled = False
    Me.lblPer = "0%"
    Me.lblCurrentBroadcast = "-nobody"
    lblInfo.Caption = "Stopped by user request"
End Sub


Private Sub Form_Load()

  LoadBroadcastChars
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub

Public Sub LoadBroadcastChars()
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  Dim i As Long
  Dim firstC As Long
  firstC = 0
  cmbCharacter.Clear
  cmbCharacter.AddItem "-", 0
  For i = 1 To MAXCLIENTS
    If GameConnected(i) = True Then
      If firstC = 0 Then
        firstC = i
      End If
      cmbCharacter.AddItem CharacterName(i), i
    Else
      cmbCharacter.AddItem "-" & CStr(i) & "- NOT CONNECTED", i
    End If
  Next i
  cmbCharacter.ListIndex = firstC
  cmbCharacter.Text = cmbCharacter.List(firstC)
  broadcastIDselected = firstC
  Exit Sub
gotErr:
 LogOnFile "errors.txt", "Error at LoadBroadcastChars(). Err number " & CStr(Err.Number) & " ; Err description " & Err.Description
End Sub




Private Sub timerBroadcast_Timer()
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
    Dim lngv As Long
    Dim dblPer As Double
    Dim lstC As Long
    Dim strD As String
    Dim privm As String
    Dim aRes As Long
    Dim currentBID As Long
    If broadcastIDselected = 0 Then
        timerBroadcast.enabled = False
        Exit Sub
    End If
    If chkMC.Value = 1 Then
        currentBID = nextValidBID()
        broadcastIDselected = currentBID
        Me.cmbCharacter.ListIndex = broadcastIDselected
    Else
        currentBID = broadcastIDselected
    End If
    If GameConnected(currentBID) = False Then
        timerBroadcast.enabled = False
        Exit Sub
    End If
    If CheatsPaused(currentBID) = True Then
        Exit Sub
    End If
    lngv = currentBroadcastIndex
    lngv = lngv + 1
    lstC = Me.lstList.ListCount
    If lngv >= lstC Then
        timerBroadcast.enabled = False
        currentBroadcastIndex = -1
        Me.lblCurrentBroadcast = "-nobody"
        Me.lblPer = "0%"
        lblInfo.Caption = "Broadcast finished. Reached end of list"
    Else
        dblPer = Round((CDbl(lngv + 1) / CDbl(lstC)) * 100, 2)
        Me.lblPer = CStr(dblPer) & "%"
        currentBroadcastIndex = lngv
        strD = Me.lstList.List(lngv)
        Me.lblCurrentBroadcast = strD
        privm = "*" & strD & "* " & parseVars(CInt(currentBID), Trim$(txtBroadcastThis.Text))
        aRes = ExecuteInTibia(privm, CInt(currentBID), True)
        UpdateTimer
        DoEvents
    End If
    Exit Sub
gotErr:
        timerBroadcast.enabled = False
        currentBroadcastIndex = -1
        Me.lblCurrentBroadcast = "-nobody"
        Me.lblPer = "0%"
        lblInfo.Caption = "Broadcast finished. Unexpected error " & CStr(Err.Number)
End Sub
