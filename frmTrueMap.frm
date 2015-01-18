VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTrueMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "True map"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3900
   Icon            =   "frmTrueMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   3900
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid gridMap 
      Height          =   2775
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   14
      Cols            =   18
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   0
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   15
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3000
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   14
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2800
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2600
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   12
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   11
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2200
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   10
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2000
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   9
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   8
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1600
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00C0FFFF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   7
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1400
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1000
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   800
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   400
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   200
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   300
   End
   Begin VB.TextBox txtSelected 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   2780
      Width           =   3570
   End
End
Attribute VB_Name = "frmTrueMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit
Public Function ColourPriority(Colour As ColorConstants) As Integer
  Select Case Colour
  Case ColourNothing
    ColourPriority = 0
  Case ColourGround ' ground
    ColourPriority = 1
  Case ColourWater 'water
    ColourPriority = 2
  Case ColourFish ' with fish
    ColourPriority = 3
  Case ColourBlockMoveable ' blocking , but moveable
    ColourPriority = 4
  Case ColourSomething 'blocking + not moveable
    ColourPriority = 5
  Case ColourField 'field
    ColourPriority = 6
  Case ColourDown ' ladder down
    ColourPriority = 7
  Case ColourUp ' ladder up
    ColourPriority = 8
  Case ColourPlayer
    ColourPriority = 50
  Case ColourWithMe
    ColourPriority = 99
  End Select
End Function
Public Sub DrawFloor()
'draw floor
  Dim posX As Long
  Dim posY As Long
  Dim PosS As Long
  Dim tileID As Long
  Dim saveCol As Long
  Dim saveRow As Long
  Dim i As Long
  Dim j As Long
  Dim relCol As Long
  Dim relRow As Long
  Dim gotMobiles As Boolean
  Dim tmpID As Double
  Dim tmpName As String
  saveCol = gridMap.Col
  saveRow = gridMap.Row
  gridMap.Redraw = False

  If mapIDselected > 0 Then
  frmHardcoreCheats.lblPosition = "x=" & myX(mapIDselected) & ", y=" & myY(mapIDselected) & ", z=" & myZ(mapIDselected)
  For posY = -6 To 7
    For posX = -8 To 9
      gotMobiles = False
      gridMap.Col = 8 + posX
      gridMap.Row = 6 + posY
      gridMap.CellBackColor = ColourNothing
      gridMap.TextMatrix(6 + posY, 8 + posX) = ""
      For PosS = 0 To 10
        tileID = GetTheLong(Matrix(posY, posX, mapFloorSelected, mapIDselected).s(PosS).t1, _
         Matrix(posY, posX, mapFloorSelected, mapIDselected).s(PosS).t2)
        If tileID = 0 Then
           Exit For
        ElseIf tileID = 97 Then
          gridMap.Col = 8 + posX
          gridMap.Row = 6 + posY
          gridMap.CellBackColor = ColourPlayer
          If gridMap.TextMatrix(6 + posY, 8 + posX) = "" Then
          tmpID = Matrix(posY, posX, mapFloorSelected, mapIDselected).s(PosS).dblID
          If tmpID = 0 Then
          gridMap.TextMatrix(6 + posY, 8 + posX) = ""
          Else
          gridMap.TextMatrix(6 + posY, 8 + posX) = GetNameFromID(mapIDselected, tmpID)
          End If
          Else
            tmpID = Matrix(posY, posX, mapFloorSelected, mapIDselected).s(PosS).dblID
            If tmpID = 0 Then
              tmpName = ""
            Else
              tmpName = GetNameFromID(mapIDselected, tmpID)
            End If
            gridMap.TextMatrix(6 + posY, 8 + posX) = gridMap.TextMatrix(6 + posY, 8 + posX) & " , " & tmpName
          End If
           gotMobiles = True
        ElseIf PosS = 0 Then
          If tileID <> &H0 Then
            gridMap.Col = 8 + posX
            gridMap.Row = 6 + posY
            gridMap.CellBackColor = ColourGround

            If ColourPriority(ColourUp) > ColourPriority(gridMap.CellBackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap.CellBackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap.CellBackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap.CellBackColor = ColourDown
              End If
            End If
            
            If ColourPriority(ColourWater) > ColourPriority(gridMap.CellBackColor) Then
              If DatTiles(tileID).isWater = True Then
                gridMap.CellBackColor = ColourWater
                If DatTiles(tileID).haveFish = True And _
                 ColourPriority(ColourFish) > ColourPriority(gridMap.CellBackColor) Then
                  gridMap.CellBackColor = ColourFish
                End If
              End If
            End If
          End If
          If DatTiles(tileID).isWater = False Then
            If ColourPriority(ColourSomething2) > ColourPriority(gridMap.CellBackColor) Then
              If DatTiles(tileID).blocking = True And DatTiles(tileID).notMoveable = False Then
                gridMap.CellBackColor = ColourSomething2
              End If
            End If
            If ColourPriority(ColourSomething) > ColourPriority(gridMap.CellBackColor) Then
              If DatTiles(tileID).blocking = True And DatTiles(tileID).notMoveable = True Then
                gridMap.CellBackColor = ColourSomething
              End If
            End If
          End If
            If ColourPriority(ColourField) > ColourPriority(gridMap.CellBackColor) Then
              If DatTiles(tileID).isField Then
                gridMap.CellBackColor = ColourField
              End If
            End If
        Else
          If ((tileID > 99) And (tileID <= highestDatTile)) Then
            If DatTiles(tileID).blocking = False Then
                gridMap.Col = 8 + posX
                gridMap.Row = 6 + posY
                If ColourPriority(ColourGround) > ColourPriority(gridMap.CellBackColor) Then
                  gridMap.CellBackColor = ColourGround
                End If
                If ColourPriority(ColourUp) > ColourPriority(gridMap.CellBackColor) Then
                  If DatTiles(tileID).floorChangeUP Then
                    gridMap.CellBackColor = ColourUp
                  End If
                End If
               If ColourPriority(ColourDown) > ColourPriority(gridMap.CellBackColor) Then
                 If DatTiles(tileID).floorChangeDOWN Then
                   gridMap.CellBackColor = ColourDown
                 End If
               End If
               If ColourPriority(ColourField) > ColourPriority(gridMap.CellBackColor) Then
              If DatTiles(tileID).isField Then
                gridMap.CellBackColor = ColourField
              End If
            End If
            Else
                gridMap.Col = 8 + posX
                gridMap.Row = 6 + posY
                If DatTiles(tileID).notMoveable = True Then
                  ' blocking and not moveable
                  If ColourPriority(ColourSomething) > ColourPriority(gridMap.CellBackColor) Then
                    gridMap.CellBackColor = ColourSomething
                  End If
                  
            If ColourPriority(ColourUp) > ColourPriority(gridMap.CellBackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap.CellBackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap.CellBackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap.CellBackColor = ColourDown
              End If
            End If
                
                Else ' blocking but moveable
                  If ColourPriority(ColourBlockMoveable) > ColourPriority(gridMap.CellBackColor) Then
                    gridMap.CellBackColor = ColourBlockMoveable
                  End If
                
                         If ColourPriority(ColourUp) > ColourPriority(gridMap.CellBackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap.CellBackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap.CellBackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap.CellBackColor = ColourDown
              End If
            End If
                
                
                End If
            End If
          ElseIf tileID = 0 Then
             ' end of stack
             Exit For
          End If
        End If
      Next PosS
    Next posX
  Next posY
  relCol = 8 + mapFloorSelected - myZ(mapIDselected)
  relRow = 6 + mapFloorSelected - myZ(mapIDselected)
  If relCol >= 0 And relCol <= 17 And relRow >= 0 And relRow <= 13 Then
    gridMap.Col = relCol
    gridMap.Row = relRow
    gridMap.CellBackColor = ColourWithMe
  End If
  Else
  For i = 0 To 13
    For j = 0 To 17
      gridMap.TextMatrix(i, j) = ""
      gridMap.Col = j
      gridMap.Row = i
      gridMap.CellBackColor = ColourNothing
    Next j
  Next i
  relCol = 8 + mapFloorSelected - 7
  relRow = 6 + mapFloorSelected - 7
  If relCol >= 0 And relCol <= 17 And relRow >= 0 And relRow <= 13 Then
    gridMap.Col = relCol
    gridMap.Row = relRow
    If ColourPriority(ColourWithMe) > ColourPriority(gridMap.CellBackColor) Then
      gridMap.CellBackColor = ColourWithMe
    End If
  End If
  End If


  gridMap.Col = saveCol
  gridMap.Row = saveRow
  gridMap.Redraw = True
End Sub
Public Sub SetButtonColours()
 Dim z As Long
 Dim i As Long
 If mapIDselected = 0 Then
   z = 7
 Else
   z = myZ(mapIDselected)
 End If
 If z <= 7 Then
   For i = 0 To 7
     If i = z Then
       cmdFloor(i).BackColor = ColourWithMe
     Else
       cmdFloor(i).BackColor = ColourWithInfo
     End If
   Next i
   For i = 8 To 15
   cmdFloor(i).BackColor = ColourUnknown
   Next i
 Else
   For i = 0 To 15
     cmdFloor(i).BackColor = ColourUnknown
   Next i
   For i = z - 2 To z + 2
     If i = z Then
       cmdFloor(i).BackColor = ColourWithMe
     ElseIf i <= 15 Then
       cmdFloor(i).BackColor = ColourWithInfo
     End If
   Next i
 End If
 cmdFloor(mapFloorSelected).BackColor = ColourSelected
End Sub

Public Sub LoadChars()
  Dim i As Long
  Dim firstC As Long
  firstC = 0
  frmHardcoreCheats.cmbCharacter.Clear
  frmHardcoreCheats.cmbCharacter.AddItem "-", 0
  For i = 1 To MAXCLIENTS
    If GameConnected(i) = True Then
      If firstC = 0 Then
        firstC = i
      End If
      frmHardcoreCheats.cmbCharacter.AddItem CharacterName(i), i
    Else
      frmHardcoreCheats.cmbCharacter.AddItem "-", i
    End If
  Next i
  frmHardcoreCheats.cmbCharacter.ListIndex = firstC
  frmHardcoreCheats.cmbCharacter.Text = frmHardcoreCheats.cmbCharacter.List(firstC)
  mapIDselected = firstC
End Sub
























Private Sub cmdFloor_Click(Index As Integer)
   If TrialVersion = True Then
      If mapIDselected > 0 Then
        If sentWelcome(mapIDselected) = False Or GotPacketWarning(mapIDselected) = True Then
          Exit Sub
        End If
      End If
    End If
  mapFloorSelected = Index
  SetButtonColours
  DrawFloor
End Sub





Private Sub Command1_Click()
      EvalMyMove mapIDselected, 0, 0, -1
      myZ(mapIDselected) = myZ(mapIDselected) - 1
      SetButtonColours
      DrawFloor
End Sub



Private Sub Form_Load()
  Dim i As Integer
  Dim j As Integer
  gridMap.Clear
  gridMap.Rows = 14
  gridMap.Cols = 18
  For i = 0 To 17
    gridMap.ColWidth(i) = 200
  Next i
  For i = 0 To 13
    For j = 0 To 17
      gridMap.TextMatrix(i, j) = ""
      gridMap.Col = j
      gridMap.Row = i
      gridMap.CellBackColor = ColourNothing
    Next j
  Next i
  LoadChars
  If mapIDselected = 0 Then
    mapFloorSelected = 7
  Else
    mapFloorSelected = myZ(mapIDselected)
    frmHardcoreCheats.lblPosition = "x=" & myX(mapIDselected) & ", y=" & myY(mapIDselected) & ", z=" & myZ(mapIDselected)
  End If
  SetButtonColours
  DrawFloor
End Sub


Private Sub Form_Resize()
  If (MapWantedOnTop = True) And (frmTrueMap.WindowState <> vbMinimized) Then
    ToggleTopmost frmTrueMap.hwnd, True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmHardcoreCheats.chkManualUpdate.Value = True
  frmHardcoreCheats.chkUpdateMs.Value = False
  frmHardcoreCheats.chkAutoUpdateMap.Value = False
  frmHardcoreCheats.timerAutoUpdater.enabled = False
  Me.Hide
  If BlockUnload = 0 Then
    ToggleTopmost frmTrueMap.hwnd, False
  End If
  Cancel = BlockUnload
End Sub

Private Sub gridMap_Click()
  Dim i As Long
  Dim name As String
  Dim id As Double
  Dim sCheat As String
  Dim partTwo As String
  Dim numI As Byte
  Dim lon As Long
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim cPacket() As Byte
  Dim tileID As Long
  Dim inRes As Integer
  Dim aRes As Long
  
     If TrialVersion = True Then
      If mapIDselected > 0 Then
        If sentWelcome(mapIDselected) = False Or GotPacketWarning(mapIDselected) = True Then
          Exit Sub
        End If
      End If
    End If
    
  If frmHardcoreCheats.ActionInspect.Value = True Then
    ' update info
    If gridMap.TextMatrix(gridMap.Row, gridMap.Col) <> "" Then
      txtSelected.Text = gridMap.TextMatrix(gridMap.Row, gridMap.Col)
    End If
    If mapIDselected > 0 Then
    frmHardcoreCheats.lblArraySelected.Caption = ""
    For i = 0 To 10
      frmHardcoreCheats.lblArraySelected.Caption = frmHardcoreCheats.lblArraySelected.Caption & _
       GoodHex(Matrix(-6 + gridMap.Row, -8 + gridMap.Col, mapFloorSelected, mapIDselected).s(i).t1) & " " & _
       GoodHex(Matrix(-6 + gridMap.Row, -8 + gridMap.Col, mapFloorSelected, mapIDselected).s(i).t2) & " " & _
       GoodHex(Matrix(-6 + gridMap.Row, -8 + gridMap.Col, mapFloorSelected, mapIDselected).s(i).t3)
      
      id = Matrix(-6 + gridMap.Row, -8 + gridMap.Col, mapFloorSelected, mapIDselected).s(i).dblID
      If id > 0 Then
        name = GetNameFromID(mapIDselected, id)
        frmHardcoreCheats.lblArraySelected.Caption = frmHardcoreCheats.lblArraySelected.Caption & ":" & name & _
        " (" & SpaceID(id) & ") ; "
      Else
        frmHardcoreCheats.lblArraySelected.Caption = frmHardcoreCheats.lblArraySelected.Caption & " ; "
      End If
    Next i
    aRes = GameInspect(mapIDselected, -8 + gridMap.Col, -6 + gridMap.Row, mapFloorSelected)
    End If
  ElseIf frmHardcoreCheats.ActionMove.Value = True Then
    If gridMap.TextMatrix(gridMap.Row, gridMap.Col) <> "" Then
      txtSelected.Text = gridMap.TextMatrix(gridMap.Row, gridMap.Col)
    End If
    If mapIDselected > 0 Then
    If GameConnected(mapIDselected) = True Then
    frmHardcoreCheats.lblArraySelected.Caption = ""
    
    sCheat = "6E 00 " & FiveChrLon(tileID_Backpack) & " 08 00 42 61 63 6B 70 61 63 6B 14 00 "
    lon = 16
    partTwo = ""
    numI = 0
    For i = 0 To 10
      name = ""
      id = Matrix(-6 + gridMap.Row, -8 + gridMap.Col, mapFloorSelected, mapIDselected).s(i).dblID
      frmHardcoreCheats.lblArraySelected.Caption = frmHardcoreCheats.lblArraySelected.Caption & _
       GoodHex(Matrix(-6 + gridMap.Row, -8 + gridMap.Col, mapFloorSelected, mapIDselected).s(i).t1) & " " & _
       GoodHex(Matrix(-6 + gridMap.Row, -8 + gridMap.Col, mapFloorSelected, mapIDselected).s(i).t2) & " " & _
       GoodHex(Matrix(-6 + gridMap.Row, -8 + gridMap.Col, mapFloorSelected, mapIDselected).s(i).t3)

      If id > 0 Then
        name = GetNameFromID(mapIDselected, id)
        frmHardcoreCheats.lblArraySelected.Caption = frmHardcoreCheats.lblArraySelected.Caption & ":" & name & _
        " (" & SpaceID(id) & ") ; "
      Else
        frmHardcoreCheats.lblArraySelected.Caption = frmHardcoreCheats.lblArraySelected.Caption & " ; "
      End If
      b1 = Matrix(-6 + gridMap.Row, -8 + gridMap.Col, mapFloorSelected, mapIDselected).s(i).t1
      b2 = Matrix(-6 + gridMap.Row, -8 + gridMap.Col, mapFloorSelected, mapIDselected).s(i).t2
      b3 = Matrix(-6 + gridMap.Row, -8 + gridMap.Col, mapFloorSelected, mapIDselected).s(i).t3
      tileID = GetTheLong(b1, b2)
      If tileID = 97 Then
        lon = lon + 2
        numI = numI + 1
        partTwo = partTwo & FiveChrLon(tileID_Oracle)
      ElseIf tileID <> 0 Then
        numI = numI + 1
        If DatTiles(tileID).haveExtraByte = True Then
          partTwo = partTwo & " " & FiveChrLon(tileID) & " " & GoodHex(b3)
          lon = lon + 3
        Else
          partTwo = partTwo & " " & FiveChrLon(tileID)
          lon = lon + 2
        End If
      End If
    Next i
    lon = lon + 1
    sCheat = FiveChrLon(lon) & " " & sCheat & GoodHex(numI) & partTwo
    'frmMain.txtPackets = frmMain.txtPackets.Text & vbCrLf & "test: " & sCheat
    inRes = GetCheatPacket(cPacket, sCheat)
    frmMain.UnifiedSendToClientGame mapIDselected, cPacket
    End If
    End If
  ElseIf frmHardcoreCheats.ActionPath.Value = True Then
    If mapIDselected > 0 Then
      If (GameConnected(mapIDselected) = True) And (GotPacketWarning(mapIDselected) = False) Then
        aRes = FindBestPath(mapIDselected, -8 + gridMap.Col, -6 + gridMap.Row, True)
      End If
    End If
  End If
End Sub







