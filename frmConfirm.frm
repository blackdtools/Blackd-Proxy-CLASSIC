VERSION 5.00
Begin VB.Form frmConfirm 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Are you sure?"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "No"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdYes 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Yes"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00000000&
      Caption         =   "Do you really want to close Blackd Proxy?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Private Sub cmdNo_Click()
    Me.Hide
End Sub

Private Sub cmdYes_Click()
    confirmedExit = True
    Unload frmMenu
    End
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub
