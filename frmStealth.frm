VERSION 5.00
Begin VB.Form frmStealth 
   BackColor       =   &H00000000&
   Caption         =   "Stealth"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8445
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmStealth.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAvoidChat 
      BackColor       =   &H00000000&
      Caption         =   "Avoid chat here"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   0
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox chkStealthExp 
      BackColor       =   &H00000000&
      Caption         =   "Exp here"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   0
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox chkStealthMessages 
      BackColor       =   &H00000000&
      Caption         =   "Bot messages here"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   0
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chkStealthCommands 
      BackColor       =   &H00000000&
      Caption         =   "Commands here"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox txtCommand 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   7215
   End
   Begin VB.ComboBox cmbCharacter 
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Text            =   "-"
      Top             =   0
      Width           =   1695
   End
   Begin VB.TextBox txtBoard 
      BackColor       =   &H00404040&
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
      Height          =   1575
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   8415
   End
   Begin VB.Label lblCommand 
      BackColor       =   &H00000000&
      Caption         =   "Command >"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblChar 
      BackColor       =   &H00000000&
      Caption         =   "Char:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmStealth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Private Const WMVSCROLL As Long = &H115
Private Const SBBOTTOM As Long = 7
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Const TheLastCommand As Long = 5
Private LastCommand(0 To TheLastCommand) As String
Private CurrCommand As Long
Private previewCommand As Long

Private Sub Doresize()
  If frmStealth.WindowState <> vbMinimized Then
    If frmStealth.ScaleHeight < 2000 Then
      frmStealth.Height = 2000
    End If
    If frmStealth.ScaleWidth < 8100 Then
      frmStealth.Width = 8100
    End If
    txtBoard.Height = frmStealth.ScaleHeight - 800
    txtBoard.Width = frmStealth.ScaleWidth
    txtCommand.Top = txtBoard.Height + 465
    txtCommand.Width = txtBoard.Width - 1250
    Me.lblCommand.Top = txtCommand.Top
  End If
End Sub




Private Sub cmbCharacter_Click()
 stealthIDselected = cmbCharacter.ListIndex
 UpdateValues
End Sub

Private Sub Form_Load()
    Dim i As Long
    Doresize
    LoadStealthChars
    For i = 0 To TheLastCommand
        LastCommand(i) = ""
    Next i
    CurrCommand = TheLastCommand
    previewCommand = TheLastCommand
End Sub

Private Sub Form_Resize()
Doresize
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub

Public Sub UpdateValues()
    '...
    Dim theindex As Integer
    theindex = CInt(cmbCharacter.ListIndex)
    If theindex = 0 Then
        Me.txtBoard.Text = "Select a character so you can read all their bot messages." & vbCrLf & _
        "All commands typed here will be also executed for that character." & vbCrLf & _
        "Commands casted while no character is selected will be ignored."
        Me.Caption = "Stealth"
    Else
        If Len(stealthLog(theindex)) > 5000 Then
            stealthLog(theindex) = "--Log cleared in order to save memory--"
        End If
        Me.txtBoard.Text = stealthLog(theindex)
        
        ScrollToBottom
    End If
    stealthIDselected = theindex
End Sub
Public Sub LoadStealthChars()
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
      cmbCharacter.AddItem "-", i
    End If
  Next i
  cmbCharacter.ListIndex = firstC
  cmbCharacter.Text = cmbCharacter.List(firstC)
  stealthIDselected = firstC
  UpdateValues
End Sub

Public Sub ScrollToBottom()
   SendMessage txtBoard.hwnd, WMVSCROLL, SBBOTTOM, 0
End Sub



Private Sub txtCommand_KeyDown(KeyCode As Integer, Shift As Integer)

    If ((KeyCode = 38) And (Shift = 1)) Then ' shift + up
        txtCommand.Text = LastCommand(previewCommand)
        previewCommand = previewCommand - 1
        If previewCommand < 0 Then
             previewCommand = TheLastCommand
        End If
  
'        If Len(txtCommand.Text) > 0 Then
'        txtCommand.SelStart = Len(txtCommand.Text) - 1
'        txtCommand.SelLength = 0
'        End If
    ElseIf ((KeyCode = 40) And (Shift = 1)) Then ' shift + down
        txtCommand.Text = LastCommand(previewCommand)
        previewCommand = previewCommand + 1
        If previewCommand > TheLastCommand Then
             previewCommand = 0
        End If
   
'        If Len(txtCommand.Text) > 0 Then
'        txtCommand.SelStart = Len(txtCommand.Text) - 1
'        txtCommand.SelLength = 0
'        End If
    End If
End Sub

Private Sub txtCommand_KeyPress(KeyAscii As Integer)
    Dim strCommand As String
    Dim iRes As Integer
    If KeyAscii = 13 Then
        strCommand = Trim$(txtCommand.Text)
        If ((txtCommand.Text <> "")) Then
            LastCommand(CurrCommand) = strCommand
            CurrCommand = CurrCommand + 1
            If CurrCommand > TheLastCommand Then
                CurrCommand = 0
            End If
            If stealthIDselected > 0 Then
                If chkAvoidChat.value = 1 Then
                    iRes = ExecuteInTibia(strCommand, stealthIDselected, True, True)
                Else
                    iRes = ExecuteInTibia(strCommand, stealthIDselected, True)
                End If
            End If
            txtCommand.Text = ""
        End If
    Else
        previewCommand = CurrCommand
    End If
End Sub

