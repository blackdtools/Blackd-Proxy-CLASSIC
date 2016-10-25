VERSION 5.00
Begin VB.Form frmMapReader 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blackd Proxy Big Map"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   Icon            =   "frmMapReader.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   255
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   427
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timerBigMapUpdate 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6000
      Top             =   -120
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FFFF80&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1230
      Width           =   855
   End
   Begin VB.CommandButton cmdRedraw 
      BackColor       =   &H00FFFF80&
      Caption         =   "REDRAW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1035
      Width           =   855
   End
   Begin VB.CommandButton cmdCenter2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1800
      Width           =   450
   End
   Begin VB.ComboBox cmbCenter 
      Height          =   315
      Left            =   4080
      TabIndex        =   14
      Text            =   "-"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.ComboBox cmbLeftAction 
      Height          =   315
      Left            =   4080
      TabIndex        =   12
      Text            =   "Center map here"
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   495
   End
   Begin VB.CommandButton cmdFloorDown 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Floor DOWN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1230
      Width           =   1215
   End
   Begin VB.CommandButton cmdFloorUp 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Floor UP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1035
      Width           =   1215
   End
   Begin VB.CommandButton cmdEast 
      BackColor       =   &H00FF8080&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   315
      Width           =   495
   End
   Begin VB.CommandButton cmdWest 
      BackColor       =   &H00FF8080&
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   315
      Width           =   495
   End
   Begin VB.CommandButton cmdSouth 
      BackColor       =   &H00FF8080&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   510
      Width           =   495
   End
   Begin VB.CommandButton cmdNorth 
      BackColor       =   &H00FF8080&
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3870
      Left            =   0
      ScaleHeight     =   256
      ScaleMode       =   0  'User
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   0
      Width           =   3870
   End
   Begin VB.Label Label6 
      BackColor       =   &H00404040&
      Caption         =   "Right click action: Center map"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   19
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Caption         =   "Marks:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404040&
      Caption         =   "Selected center (C) :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404040&
      Caption         =   "Left click action:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblName 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3960
      TabIndex        =   10
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   "Change floor:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Move map:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMapReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit
Private DefaultMapFile As String
Private defX As Long
Private defY As Long
Private defZ As Long

Public Sub SetDefaultMapPosition(x As Long, y As Long, z As Long)
  defX = x
  defY = y
  defZ = z
End Sub

Public Sub SetCurrentCenter(str As String)
  Dim i As Integer
  Dim res As TypeBMnameInfo
  If cmbCenter.Text <> str Then
    cmbCenter.Text = str
  End If
  For i = 1 To MAXCLIENTS
    If (str = CharacterName(i)) And (GameConnected(i) = True) Then
      SetDefaultMapPosition myX(i), myY(i), myZ(i)
      Exit Sub
    End If
  Next i
  If ExistBigMapName(str) Then
    res = GetBigMapNameInfo(str)
    SetDefaultMapPosition res.x, res.y, res.z
    Exit Sub
  End If
  'else
  SetDefaultMapPosition 32097, 32219, 7
End Sub
Public Sub DrawDefaultMap()
  currMapX = defX
  currMapY = defY
  currMapZ = defZ
  DrawMap
End Sub

Public Sub ShowCenter()
  SetCurrentCenter cmbCenter.Text
  DrawDefaultMap
End Sub













Private Sub cmdCenter2_Click()
  ShowCenter
End Sub
Private Function isMCchar(str As String) As Boolean
  Dim i As Integer
  For i = 1 To MAXCLIENTS
    If (str = CharacterName(i)) And (GameConnected(i) = True) Then
      isMCchar = True
      Exit Function
    End If
  Next i
  isMCchar = False
  Exit Function
End Function

Private Sub cmdClear_Click()
Dim toRemove As String
Dim i As Long
Dim co As Long
anotherRound:
co = cmbCenter.ListCount - 1
For i = 0 To co
  toRemove = cmbCenter.List(i)
  If (isMCchar(toRemove) = False) And (toRemove <> "-") Then
    RemoveListItem toRemove
    RemoveBigMapName toRemove
    DoEvents
    GoTo anotherRound
  End If
Next i
cmbCenter.Text = cmbCenter.List(0)
RedrawAllMarks
End Sub

Private Sub cmdEast_Click()
  Dim eval As Long
  currMapX = currMapX + MAXX
  DrawMap
End Sub

Private Sub cmdFloorDown_Click()
  If currMapZ < 15 Then
    currMapZ = currMapZ + 1
    DrawMap
  End If
End Sub

Private Sub cmdFloorUp_Click()
  If currMapZ > 0 Then
    currMapZ = currMapZ - 1
    DrawMap
  End If
End Sub

Private Sub cmdNorth_Click()
  Dim eval As Long
  currMapY = currMapY - MAXY
  DrawMap
End Sub

Public Sub RedrawAllMarks()
  Dim i As Long
  Dim li As Long
  Dim res As TypeBMnameInfo
  Dim x1 As Long
  Dim y1 As Long
  Dim z1 As Long
  Dim x2 As Long
  Dim y2 As Long
  Dim z2 As Long
  Dim x3 As Long
  Dim y3 As Long
  Dim z3 As Long
  Dim incX As Long
  Dim incY As Long
  Dim incZ As Long
  Dim isC As Boolean
  Dim co As Long
  picMap.Refresh
  co = cmbCenter.ListCount - 1
  For li = 0 To co
  isC = False
  For i = 1 To MAXCLIENTS
    If (cmbCenter.List(li) = CharacterName(i)) And (GameConnected(i) = True) Then
      res = GetBigMapNameInfo(cmbCenter.List(li))
      x3 = res.x
      y3 = res.y
      z3 = res.z
      x1 = myX(i)
      y1 = myY(i)
      z1 = myZ(i)
      incX = 3 * (x1 - x3)
      incY = 3 * (y1 - y3)
      incZ = 3 * (z1 - z3)
      x2 = x1 + incX
      y2 = y1 + incY
      z2 = z1 + incZ
      isC = True
      AddBigMapName cmbCenter.List(li), myX(i), myY(i), myZ(i), vbBlue
    End If
  Next i
  If ExistBigMapName(cmbCenter.List(li)) Then
    res = GetBigMapNameInfo(cmbCenter.List(li))
      If cmbCenter.List(li) = cmbCenter.Text Then
        DrawXYZnMap res.x, res.y, res.z, vbYellow
        If isC = True Then
          DrawLine x1, y1, z1, x2, y2, z2, vbYellow
        End If
      Else
        DrawXYZnMap res.x, res.y, res.z, res.color
        If isC = True Then
          DrawLine x1, y1, z1, x2, y2, z2, res.color
        End If
      End If
  End If
  Next li
End Sub
Private Sub cmdRedraw_Click()
  RedrawAllMarks
End Sub

Private Sub cmdSouth_Click()
  Dim eval As Long
  currMapY = currMapY + MAXY
  DrawMap
End Sub

Private Sub cmdUpdate_Click()
  ShowCenter
End Sub


Private Sub cmdWest_Click()
  Dim eval As Long
  currMapX = currMapX - MAXX
  DrawMap
End Sub

Private Function PositionListItem(str As String) As Long
  Dim i As Long
  Dim co As Long
  Dim res As Long
  res = -1
  co = cmbCenter.ListCount - 1
  For i = 0 To co
    If cmbCenter.List(i) = str Then
       res = i
     Exit For
    End If
  Next i
  PositionListItem = res
End Function

Public Sub AddListItem(str As String)
  If PositionListItem(str) = -1 Then
    cmbCenter.AddItem str
  End If
End Sub

Public Sub RemoveListItem(str As String)
  Dim posI As Long
  posI = PositionListItem(str)
  If posI <> -1 Then
    cmbCenter.RemoveItem posI
  End If
End Sub

Private Sub Form_Load()
cmbLeftAction.Clear
cmbLeftAction.AddItem "Common move"
cmbLeftAction.AddItem "Common move (all MCS)"
cmbLeftAction.AddItem "Experimental move"
cmbLeftAction.AddItem "Add scriptpoint here"
cmbLeftAction.AddItem "Set mark here"
cmbLeftAction.AddItem "Nothing"
cmbLeftAction.Text = "Common move"
cmbCenter.Clear
cmbCenter.AddItem "-"
cmbCenter.Text = "-"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  timerBigMapUpdate.enabled = False
  Cancel = BlockUnload
End Sub

Private Sub Form_Resize()
  If (MapWantedOnTop = True) And (frmMapReader.WindowState <> vbMinimized) Then
    ToggleTopmost frmMapReader.hwnd, True
  End If
End Sub
Public Function AddMarkToBigMap(x As Long, y As Long, z As Long) As String
  Dim i As Long
  Dim markName As String
  i = 1
  While PositionListItem("#" & CStr(i)) <> -1
    i = i + 1
  Wend
  markName = "#" & CStr(i)
  AddListItem markName
  AddBigMapName markName, x, y, z, vbRed

  DrawXYZnMap x, y, z, vbRed
  AddMarkToBigMap = markName
End Function

Public Sub AddPlayerToBigMap(pname As String, rawPosition As String)
  Dim aRes As Long
  Dim x As Long
  Dim y As Long
  Dim z As Long
  Dim xs As String
  Dim ys As String
  Dim zs As String
  Dim pos As Long
  Dim toEnd As Long
  On Error GoTo errIg
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
  AddListItem pname
  AddBigMapName pname, x, y, z, vbBlue
  Exit Sub
errIg:
  'ignore error
End Sub

Private Sub picMap_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim incX As Long
  Dim incY As Long
  Dim currposX As Long
  Dim currposY As Long
  Dim firstO As String
  Dim idConnection As Integer
  Dim i As Long
  Dim leftToDraw As String
  Dim nPoints As Long
  Dim iniX As Long
  Dim iniY As Long
  Dim iniZ As Long
  Dim aRes As Long
  Dim myBpos As Long
  Dim b1 As Byte
  Dim b2 As Byte
  Dim pid As Long
  Dim desx As Long
  Dim desy As Long
  Dim desz As Long
  Dim co As Long
  Dim li As Long
  Dim isC As Boolean
  Dim tmp As String
  #If FinalMode Then
    On Error GoTo gotError
  #End If
  If (LoadingAmap = False) Then
    incX = (x \ 2) + MINX
    incY = (y \ 2) + MINY
    If Button = 2 Then 'right click
        currMapX = currMapX + incX
        currMapY = currMapY + incY
        DrawMap
    ElseIf Button = 1 Then 'left click
      Select Case cmbLeftAction.Text
      Case "Set mark here"
         tmp = AddMarkToBigMap(currMapX + incX, currMapY + incY, currMapZ)
      Case "Add scriptpoint here"
        idConnection = 0
        For i = 1 To MAXCLIENTS
          If cmbCenter.Text = CharacterName(i) Then
            idConnection = CInt(i)
          End If
        Next i
        If (idConnection > 0) Then
          DrawXYZPixel2 currMapX + incX, currMapY + incY, currMapZ, vbYellow
          AddCavebotMovePoint idConnection, currMapX + incX, currMapY + incY, currMapZ
          aRes = SendLogSystemMessageToClient(idConnection, "Added cavebot script line: move " & (currMapX + incX) & "," & (currMapY + incY) & "," & currMapZ)
          DoEvents
        End If
      Case "Common move"
        idConnection = 0
        For i = 1 To MAXCLIENTS
          If cmbCenter.Text = CharacterName(i) Then
            idConnection = CInt(i)
          End If
        Next i
        If (idConnection > 0) Then
          GetProcessIDs idConnection
          pid = ProcessID(idConnection)
          desx = currMapX + incX
          desy = currMapY + incY
          desz = currMapZ
          
'          myBpos = MyBattleListPosition(idConnection)
'          If myBpos > -1 Then
'            b1 = LowByteOfLong(desx)
'            b2 = HighByteOfLong(desx)
'            Memory_WriteByte adrXgo, b1, pid
'            Memory_WriteByte adrXgo + 1, b2, pid
'            b1 = LowByteOfLong(desy)
'            b2 = HighByteOfLong(desy)
'            Memory_WriteByte adrYgo, b1, pid
'            Memory_WriteByte adrYgo + 1, b2, pid
'            b1 = CByte(desz)
'            Memory_WriteByte adrZgo, b1, pid
'            Memory_WriteByte adrGo + (myBpos * CharDist), 1, pid 'move it!
'          End If
          SafeMemoryMoveXYZ idConnection, desx, desy, desz
        End If
      Case "Common move (all MCS)"
      
co = cmbCenter.ListCount - 1
  For li = 0 To co
  isC = False
  For i = 1 To MAXCLIENTS
    If (cmbCenter.List(li) = CharacterName(i)) And (GameConnected(i) = True) Then
      idConnection = i
      isC = True
    End If
    If isC = True Then
          GetProcessIDs idConnection
          pid = ProcessID(idConnection)
          desx = currMapX + incX
          desy = currMapY + incY
          desz = currMapZ
          
'          myBpos = MyBattleListPosition(idConnection)
'          b1 = LowByteOfLong(desx)
'          b2 = HighByteOfLong(desx)
'          Memory_WriteByte adrXgo, b1, pid
'          Memory_WriteByte adrXgo + 1, b2, pid
'          b1 = LowByteOfLong(desy)
'          b2 = HighByteOfLong(desy)
'          Memory_WriteByte adrYgo, b1, pid
'          Memory_WriteByte adrYgo + 1, b2, pid
'          b1 = CByte(desz)
'          Memory_WriteByte adrZgo, b1, pid
'          Memory_WriteByte adrGo + (myBpos * CharDist), 1, pid 'move it!
           SafeMemoryMoveXYZ idConnection, desx, desy, desz
    End If
  Next i
  Next li
      
      

          
      Case "Experimental move"
        idConnection = 0
        For i = 1 To MAXCLIENTS
          If cmbCenter.Text = CharacterName(i) Then
            idConnection = CInt(i)
          End If
        Next i
        If (idConnection > 0) Then
          iniX = myX(idConnection)
          iniY = myY(idConnection)
          iniZ = myZ(idConnection)
        If (ReadyBuffer(idConnection) = True) And (iniZ = currMapZ) Then
          frmMapReader.picMap.Refresh
          AstarBig idConnection, iniX, iniY, currMapX + incX, currMapY + incY, iniZ, True
          OptimizeBuffer idConnection
          leftToDraw = RequiredMoveBuffer(idConnection)
          ExecuteBuffer idConnection
          If (leftToDraw = "X") Or (leftToDraw = "") Then
            frmMapReader.lblName.Caption = "Could not find path"
          Else
            frmMapReader.lblName.Caption = "Path found"
          End If
          If (leftToDraw <> "") And (leftToDraw <> "X") Then
          currposX = iniX
          currposY = iniY
          
          DrawXYZPixel iniX, iniY, iniZ, vbRed 'here should go current z
          nPoints = Len(RequiredMoveBuffer(idConnection))
          For i = 1 To nPoints
            firstO = Left(leftToDraw, 1)
            leftToDraw = Right(leftToDraw, Len(leftToDraw) - 1)
            Select Case firstO
            Case StrMoveRight
              currposX = currposX + 1
            Case StrMoveNorthRight
              currposX = currposX + 1
              currposY = currposY - 1
            Case StrMoveNorth
              currposY = currposY - 1
            Case StrMoveNorthLeft
              currposX = currposX - 1
              currposY = currposY - 1
            Case StrMoveLeft
              currposX = currposX - 1
            Case StrMoveSouthLeft
              currposX = currposX - 1
              currposY = currposY + 1
            Case StrMoveSouth
              currposY = currposY + 1
            Case StrMoveSouthRight
              currposX = currposX + 1
              currposY = currposY + 1
            End Select
            DrawXYZPixel currposX, currposY, currMapZ, vbRed 'here should go current z
          Next i
          End If
        End If
        End If
      End Select
    End If
  End If
  Exit Sub
gotError:
  
End Sub

Private Sub timerBigMapUpdate_Timer()
  If LoadingAmap = False Then
    If (cmbLeftAction.Text <> "Experimental move") And (cmbLeftAction.Text <> "Add scriptpoint here") Then
      RedrawAllMarks
    End If
  End If
End Sub
