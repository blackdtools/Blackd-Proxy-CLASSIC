VERSION 5.00
Begin VB.Form frmBigText 
   BackColor       =   &H00000000&
   Caption         =   "Text board"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8235
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmBigText.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Clear board"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox txtBoard 
      Height          =   3495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   720
      Width           =   8055
   End
   Begin VB.Label lblText 
      BackColor       =   &H00000000&
      Caption         =   "Text board"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "frmBigText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 0
Option Explicit
Private Sub cmdCancel_Click()
  CanceledBoard = True
  ClosedBoard = True
  frmBigText.Hide
End Sub

Private Sub cmdClear_Click()
  txtBoard.Text = ""
End Sub

Private Sub cmdOk_Click()
  CanceledBoard = False
  ClosedBoard = True
  frmBigText.Hide
End Sub

Private Sub Form_Resize()
  If frmBigText.WindowState <> vbMinimized Then
    If frmBigText.ScaleHeight < 3000 Then
      frmBigText.Height = 3000
    End If
    If frmBigText.ScaleWidth < 5800 Then
      frmBigText.Width = 5800
    End If
    txtBoard.Height = frmBigText.ScaleHeight - 1300
    txtBoard.Width = frmBigText.ScaleWidth - 200
    cmdClear.Top = frmBigText.ScaleHeight - 480
    cmdOk.Top = frmBigText.ScaleHeight - 480
    cmdCancel.Top = frmBigText.ScaleHeight - 480
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  CanceledBoard = True
  ClosedBoard = True
  Me.Hide
  Cancel = BlockUnload
End Sub
