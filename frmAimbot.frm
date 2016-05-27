VERSION 5.00
Begin VB.Form frmAimbot 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aimbot"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAimbot 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.CheckBox chkUEcombo 
      Caption         =   "Active UE combo"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1755
   End
   Begin VB.CheckBox chkSDcombo 
      Caption         =   "Active SD combo"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1755
   End
   Begin VB.TextBox txtCombo 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Text            =   "exevo gran mas vis"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtLeader 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.ComboBox cmbCharacter 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Text            =   "-"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "LEADER MUST SAY TARGET NAME ON DEFAULT TO AIMBOT CAST SD "
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "UE spell :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Leader :"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblChar 
      Caption         =   "Aimbot :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmAimbot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit


Private Sub chkSDcombo_Click()

If lock_chkSDcombo = False Then
If aimbotIDselected > 0 Then
  If chkSDcombo.Value = 1 Then
    aimbotOptions(aimbotIDselected).chkSDcombo = True
  Else
    aimbotOptions(aimbotIDselected).chkSDcombo = False
  End If
End If
End If

End Sub

Private Sub chkUEcombo_Click()

If lock_chkUEcombo = False Then
If aimbotIDselected > 0 Then
  If chkUEcombo.Value = 1 Then
    aimbotOptions(aimbotIDselected).chkUEcombo = True
  Else
    aimbotOptions(aimbotIDselected).chkUEcombo = False
  End If
End If
End If

End Sub

Private Sub cmbCharacter_Click()

 aimbotIDselected = cmbCharacter.ListIndex
  If aimbotIDselected > 0 Then
      UpdateValues
  End If
  
End Sub

Private Sub Form_Load()

LoadAimbotChars

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub

Private Sub tmrAimbot_Timer()
Dim aRes As Long
Dim idConnection As Integer
Dim i As Integer
Dim lastm As String
Dim UEspell As String

For idConnection = 1 To MAXCLIENTS
    If (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) Then

    
    'sd combo  / leader say targets name
    If (aimbotOptions(idConnection).chkSDcombo = True) Then
        lastm = LCase(var_lastmsg(idConnection))
        If (var_lastsender(idConnection)) = aimbotOptions(idConnection).txtLeader Then
            aRes = ExecuteInTibia("exiva 0" & lastm, idConnection, True)
        End If
    End If
    
    
    'ue combo  / leader say UE
    If (aimbotOptions(idConnection).chkUEcombo = True) Then
      If LCase(var_lastmsg(idConnection)) = "exevo gran mas vis" Or _
         LCase(var_lastmsg(idConnection)) = "exevo gran mas flam" Or _
         LCase(var_lastmsg(idConnection)) = "exevo gran mas frigo" Or _
         LCase(var_lastmsg(idConnection)) = "exevo gran mas tera" Or _
         LCase(var_lastmsg(idConnection)) = "exevo gran mas pox" Then
        If (var_lastsender(idConnection) = aimbotOptions(idConnection).txtLeader) Then
            UEspell = aimbotOptions(idConnection).txtCombo
            aRes = ExecuteInTibia(UEspell, idConnection, True)
        End If
      End If
    End If
    
    'blank
    If (aimbotOptions(idConnection).chkUEcombo = True) Or (aimbotOptions(idConnection).chkSDcombo = True) Then
        var_lastmsg(idConnection) = ""
    End If


    End If
Next idConnection

End Sub

Private Sub txtCombo_change() 'Validate(Cancel As Boolean)

If aimbotIDselected > 0 Then
  aimbotOptions(aimbotIDselected).txtCombo = txtCombo.Text
End If

End Sub

Private Sub txtLeader_change() 'Validate(Cancel As Boolean)

If aimbotIDselected > 0 Then
  aimbotOptions(aimbotIDselected).txtLeader = txtLeader.Text
End If

End Sub

Public Sub LoadAimbotChars()
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
  aimbotIDselected = firstC
  UpdateValues
  
End Sub

Public Sub UpdateValues()
Dim i As Integer
Dim idConnection As Integer

If aimbotIDselected <= 0 Then
    If aimbotOptions_chkSDcombo_default = True Then
      chkSDcombo.Value = 1
    Else
      chkSDcombo.Value = 0
    End If
    If aimbotOptions_chkUEcombo_default = True Then
      chkUEcombo.Value = 1
    Else
      chkUEcombo.Value = 0
    End If
frmAimbot.txtLeader.Text = aimbotOptions_txtLeader_default
frmAimbot.txtCombo.Text = aimbotOptions_txtCombo_default
 Else
    If aimbotOptions(aimbotIDselected).chkSDcombo = True Then
      chkSDcombo.Value = 1
    Else
      chkSDcombo.Value = 0
    End If
    If aimbotOptions(aimbotIDselected).chkUEcombo = True Then
      chkUEcombo.Value = 1
    Else
      chkUEcombo.Value = 0
    End If
frmAimbot.txtLeader.Text = aimbotOptions(aimbotIDselected).txtLeader
frmAimbot.txtCombo.Text = aimbotOptions(aimbotIDselected).txtCombo
End If

End Sub
