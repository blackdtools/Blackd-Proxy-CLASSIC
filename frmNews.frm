VERSION 5.00
Begin VB.Form frmNews 
   BackColor       =   &H00000000&
   Caption         =   "News"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8235
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmNews.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBoard 
      Height          =   3495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   8055
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblText 
      BackColor       =   &H00000000&
      Caption         =   "What is new in Blackd Proxy?"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "frmNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    Dim strMsg As String
    Me.lblText = "What is new?"
    
    strMsg = "Blackd Proxy 37.2" & vbCrLf & _
    "----------------------------" & vbCrLf & _
     " - Many bug fixes." & vbCrLf & _
     " - Added command to add everything as cavebot target: SetMeleeKill *" & vbCrLf & _
     " - Added Event options to capture SYSTEM messages and RAID messages in newest Tibia version"
    
    strMsg = strMsg & vbCrLf & "Blackd Proxy 37.1" & vbCrLf & _
    "----------------------------" & vbCrLf & _
     " - Hopefully fixed an user interface bug (buttons were displayed blank for some users)."
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 37.0" & vbCrLf & _
    "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.90"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 36.9" & vbCrLf & _
    "----------------------------" & vbCrLf & _
     " - Fixed a bug with a new packet related with Tibia coins at Tibia 10.82"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 36.8" & vbCrLf & _
    "----------------------------" & vbCrLf & _
     " - Fixed a bug with a new packet related with Tibia coins at Tibia 10.81"
      
    strMsg = strMsg & vbCrLf & "Blackd Proxy 36.7" & vbCrLf & _
    "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.81"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 36.6" & vbCrLf & _
    "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.80" & vbCrLf & _
     " - Small fixes for Tibia 7.6 OT servers."
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 36.5" & vbCrLf & _
    "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.79" & vbCrLf & _
     " - Small fix for a bug with Tibia 7.6 OT servers."
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 36.4" & vbCrLf & _
    "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.78"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 36.3" & vbCrLf & _
    "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.77"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 36.2" & vbCrLf & _
    "----------------------------" & vbCrLf & _
     " - Fixed another bug with OT servers of Tibia 10.76"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 36.1" & vbCrLf & _
    "----------------------------" & vbCrLf & _
     " - Fixed a bug with OT servers of Tibia 10.76"

    strMsg = strMsg & vbCrLf & "Blackd Proxy 36.0" & vbCrLf & _
    "----------------------------" & vbCrLf & _
     " - Fixed a bug at death event in Tibia 10.76" & vbCrLf & _
     " - Fixed a bug with the anti idle feature of our trainer"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 35.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed bug that happened when a cavebot script reached the end."
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 35.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed login bug in OT servers 10.74+"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 35.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated parser for new packets at Tibia 10.76"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 35.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.76"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 35.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.75" & vbCrLf & _
     " - Many small optimizations."
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 35.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a problem reading dlls while executing code directly from sources."
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 35.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.74" & vbCrLf & _
     " - Cavebot form will now display the current line being executed."
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 35.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug with mc caveboting. Sorry for taking so much time."
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 35.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.73" & vbCrLf & _
     " - Several minor bug fixes." & vbCrLf & _
     " - Proyect now available in GitHub!"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 35.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - New exiva command: exiva useitemonname:AA BB,Abcdef =>uses custom item AA BB on first creature with name Abcdef" & vbCrLf & _
     " - New exiva command: exiva useitemonname:AA BB =>uses custom item AA BB on last attacked creature"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 34.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.72"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 34.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Blackd Proxy will now allow forced loading in OT servers with bad configs, but this is only intended for programmers who want to do debugs there. Be aware that 99% cheats will fail in such cases!!"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 34.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.71"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 34.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug with guild channels at Tibia 10.70" & vbCrLf & _
     " - Fixed a bug with party invites at Tibia 10.70"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 34.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.70"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 34.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.64"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 34.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.63"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 34.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.62"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 34.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.61"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 34.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.60"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 33.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug with speed at Tibia 10.59"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 33.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.59"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 33.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.58"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 33.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug with Tibia 10.57"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 33.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.57"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 33.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Packet parser updated with same improvements than Blackd Proxy NG 1.3"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 33.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Packet parser updated with same improvements than Blackd Proxy NG 1.2"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 33.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed another bug in Tibia 10.55"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 33.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed another bug in Tibia 10.55"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 33.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug that happened with a new kind of message in Tibia 10.55"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 32.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug that happened when being attacked in Tibia 10.55"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 32.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug that happened while inspecting things in Tibia 10.55"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 32.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed some important bugs at Tibia 10.55"
     
    strMsg = strMsg & vbCrLf & "Blackd Proxy 32.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.55"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 32.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.54"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 32.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.53"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 32.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.52"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 32.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug that some users was experiencing when trying to load latest config."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 32.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in all Tibia versions up to 10.51"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 32.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.4"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 31.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.39"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 31.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a remaining bug with the parser of the new packet type 9E (premium features window) It happened when you had few days of premium left (new warning window ingame)"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 31.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug with the parser of the new packet type 9E (premium features window)"

    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 31.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.38" & vbCrLf & _
     " - Fixed a new bug with Rookgard npcs that happened since Tibia 10.36"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 31.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.37" & vbCrLf & _
     " - Fixed a new bug with depositers that happened since Tibia 10.36"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 31.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed some more bugs in Tibia 10.36"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 31.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug in Tibia 10.36"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 31.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.36"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 31.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.35"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 31.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.34" & vbCrLf & _
     " - Now you can manually setup new config folders through the variables at the root config.ini - No need to recompile Blackd Proxy for future updates except in mayor updates."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 30.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.33"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 30.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.32"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 30.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.31"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 30.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.3"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 30.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.22"

    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 30.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.21" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.21 Preview"
    
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 30.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a rare bug of .dat reader" & vbCrLf & _
     " - Fixed crash when starting a trade with players at Tibia 10.2"
    
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 29.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.2"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 29.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.12"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 29.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed weird that caused some characters not loading"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 29.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed login bug for Windows 7 and Tibia 10.11"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 29.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.11"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 29.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.1"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 29.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.02"
    
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 29.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed new discovered bug when reading letters." & vbCrLf & _
     " - Minimum changes to work in Tibia 10.01"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 29.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed new discovered  bugs and optimized the core a bit more."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 29.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Added automatic RSA/IP change for all OT servers up to 8.55" & vbCrLf & _
     " - Optimized Blackd Proxy core speed"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 28.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed bug with detection of attacking creature/player in Tibia 10.0"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 28.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed bug when summoning creatures in Tibia 10.0"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 28.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 10.0"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 28.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 9.86"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 28.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 9.85"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 28.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Minimum changes to work in Tibia 9.84"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 28.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - We should remember you that Blackd Proxy is still detectable. Please only use it in accounts that you don't care to get deleted." & vbCrLf & _
     " - Updated to work in Tibia 9.83"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 28.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Blackd Proxy is now able to fix DirectX problems at load" & vbCrLf & _
     " - Updated to work in Tibia 9.82"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 28.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Blackd Proxy should now work for some more persons" & _
     " - Cavebot should now recognize all the stairs of Yalahar"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 28.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Bot will now detect some other advanced cavebot commands as dangerous for real servers: setRetryAttacks, dropLootOnGround" & _
     " - Bot should not pause cheats with rookgard tutorial popups. It will also hide such popups." & _
     " - Fixed a bug of the trainer feature 'Stop attacking target until regen'"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 27.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - sayMessage/sayInTrade exiva >/buy/sell should now give risk warning too."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 27.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - The variable $lastmsg$ now return the last message without symbols { }" & vbCrLf & _
     " - Cavebot will now display a yellow warning message if you load a script containing actions with high risk of detection"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 27.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.81"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 27.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a detectable thing: using cavebot command exiva sell or exiva trade with amount 0 was being detected. Since this version Blackd Proxy will not send any packet if amount is 0 or negative."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 27.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a small problem with OT servers 7.6"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 27.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - New command to define a maximum hp percent limit of target to use setSpellKill (to avoid use in monsters not attacked yet) : setBot SpellKillMaxHPlimit=X" & vbCrLf & _
     " - Cavebot can now find and use some new types of rope spots." & vbCrLf & _
     " - Fixed a bug with the classic mode. Now bot should use the potions from backpacks when this mode is enabled (except if there is none left in open backpacks)"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 27.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug with the new variable $check:{$A$},#conditionOperator#,{$B$}$"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 27.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed another rare bug for Tibia 9.8" & vbCrLf & _
     " - Added new variable: " & vbCrLf & _
     "   $check:{$A$},#conditionOperator#,{$B$}$" & vbCrLf & _
     "     returns 0 if condition A op B is evaluated as true, else returns 1 " & vbCrLf & _
     "     Example: $check:{$myhppercent$},#number<=#,{50}$" & vbCrLf & vbCrLf & _
     " - Added new variable: " & vbCrLf & _
     "   $istrue:{$var1$},{$var2$}, ... , {$varN$}$" & vbCrLf & _
     "     returns 0 if all variables are 0 else returns 1 " & vbCrLf & _
     "     Example: $istrue:{$check:{$myhppercent$},#number>=#,{50}$},{$check:{$myhppercent$},#number<=#,{70}$},{$check:{$mymana$},#number>=#,{40}$}"

    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 27.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed some reported bugs for Tibia 9.8"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 27.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.8"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 26.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug in the loading of trainer settings. Now it should work correctly."

    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 26.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - New command to define mininum hp percent limit of target to use setSpellKill (to avoid waste of mana) : setBot SpellKillHPlimit=X" & vbCrLf & _
     " - exiva save now saves the Trainer options too." & vbCrLf & _
     " - exiva load now load the Trainer options too if such settings were saved." & vbCrLf & _
     " - New advanced cavebot command for your scripts: setNoLoot xx xx   It allows you to remove the item xx xx from the loot list."

    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 26.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug with the load/save of settings when using the max number of mc chars (MAXCLIENTS=5 by default)"
    
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 26.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - The command exiva sell:XX XX:N now allows you to sell big amounts correctly. It will do it by parts, max 100 at a time."

    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 26.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Optimized loading of the character list. Blackd Proxy should now work for more people." & vbCrLf & _
     " - Fixed a small problem with the command 'exiva <' Now strings should preserve their initial uppercase/lowercase status. It was considering everything as lowercase before this patch."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 26.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Improved the internal method to get the base address of Tibia. It might fix load problems for some people." & vbCrLf & _
     " - Optimized crackd.dll for faster connection"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 26.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Blackd Proxy now automatically changes RSA key for Tibia 7.72 client. Now you don't need an ip loader for such ot servers."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 26.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed another minor bug with the character names in the menus." & vbCrLf & _
     " - Fixed an old that bug that was probably causing Tibia to reject login in some cases."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 26.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a problem with names not displaying in menus."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 26.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a problem with multiclient bot in Tibia 9.71"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 25.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.71" & vbCrLf & _
     " - Solved a load bug that caused some systems to try reading wrong .dat file"
    
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 25.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Light hack should work all the time. Now the light will be absolutely permanent. You won't see any change when using light spells / light items." & vbCrLf & _
     " - Autoload settings per char (cavebot + runemaker + conds). You can now save such settings for a character with this command: exiva save" & vbCrLf & _
     " - You can force instant reload of saved settings with command exiva load" & vbCrLf & _
     " - You can also load the settings that you saved with other character with command exiva load:characterName" & vbCrLf & _
     " - The command exiva buy:XX XX:N now allows you to buy big amounts correctly. It will do it by parts, max 100 at a time."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 25.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - New button at cavebot: Load by Copy/Paste. It will allow you to load scripts directly from forum to cavebot." & vbCrLf & _
     " - New link at main menu: VIP Support page. Because we want to give a better service to people who support us (by buying some gold from time to time)"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 25.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a minor problem when using beds in Tibia 9.7"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 25.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.7"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 25.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a small bug with exiva fish . Now it should work again for everybody."
    
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 25.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.63"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 25.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed the parser for VIP lists in Tibia 9.62"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 25.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.62"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 25.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.61"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 24.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a problem with setSdKill, setHmmKill and setSpellKill"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 24.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a problem with market in Tibia 9.6" & vbCrLf & _
     " - Fixed a minor problem in cavebot" & vbCrLf & _
     " - Fixed a problem with runemaker in Tibia 7.6" & vbCrLf & _
     " - Added a new thing that you can check with scripts: $hex-currenttargetid$ It will give the current creature marked with the red square (the creature that you are actually attacking)"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 24.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a small bug that still happened when activating premium scrolls"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 24.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a problem with Tibia 9.6 caused by the new stat Offline Time"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 24.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed useitemwithamount so it can now attempt to use any kind of item again"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 24.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.6"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 24.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Improved exiva >> Now advanced users will be able to send multiple packets with a single command."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 24.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug that was stopping the cavebot."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 24.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Some additional improvements to avoid autodetection." & vbCrLf & _
     " - Added a command for advanced users: exiva >> will allow you to send bytes without autoheader. Only for advanced users!"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 24.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed another detectable thing in the cavebot attack function. Cavebot should be very hard to detect now. However we won't guarantee anything yet." & vbCrLf & _
     " - Fixed a rare problem in market" & vbCrLf & _
     " - Fixed a common problem with the trainer menu" & vbCrLf & _
     " - Activating a premium scroll should not crash Blackd Proxy now" & vbCrLf & _
     " - Now cavebot will not have a time limit to kill monsters, unless you include this to your script: SetBot EnableMaxAttackTime=1" & vbCrLf & _
     " - Now cavebot will resync with Tibia client memory to ensure that it is really attacking the correct creature" & vbCrLf & _
     " - Now cavebot will not trigger any pk alarm during the first 5 seconds, to avoid false alarms"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 23.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.54" & vbCrLf & _
     " - All commands are now able to handle the new equipment slot 0B" & vbCrLf & _
     " - Included a menu linking to our new free flash game Shoot Fruits. Try playing it while you cavebot or make runes!" & vbCrLf & _
     " - Blackd Proxy will now show config screen again in case you get a tibia.dat error." & vbCrLf & _
     " - Cavebot now support a new discovered kind of hole at tarantula caves , tile id 6A 03" & vbCrLf & _
     " - Internal optimization: now Blackd Proxy should work correctly in more environments"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 23.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Improved safety of cavebot commands dropLootOnGround and putLootOnDepot. Now slower and more random." & vbCrLf & _
     " - Now the commands dropLootOnGround and putLootOnDepot are compatible with the new loot mode" & vbCrLf & _
     " - New cavebot command: setBot PKwarnings=0" & vbCrLf & _
     "  (this will disable all events that happen when you get attacked by something not included in your scripts)" & vbCrLf & _
     " - Runemaker is now random at multiclient level. Now it is very hard to be detected, even with human eyes." & vbCrLf & _
     " - New exiva command: exiva drop XX XX AA , drops up to AA items of tile ID XX XX under you." & vbCrLf & _
     " - Example 1: exiva drop D7 0B 01 : drops 1 gold under you = anti push." & vbCrLf & _
     " - Example 2: exiva drop D7 0B 02 : drops 2 gold under you = improved anti push." & vbCrLf & _
     " - Example 3: exiva drop D7 0B FF : drops a full stack of gold under you."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 23.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.53" & vbCrLf & _
     " - Fixed another detectable thing in the cavebot looter." & vbCrLf & _
     " - New inverse loot mode: Enabling this mode the cavebot looter will loot everything, except the IDs added with setLoot. Usage: setBot LootAll=1"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 23.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - To avoid some problems, we now force Blackd Proxy to show first config menu again after updating Blackd Proxy"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 23.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.52" & vbCrLf & _
     " - Several minor bugs are now fixed"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 23.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - New config menu at first run: A fastest way to connect to OT servers" & vbCrLf & _
     " - UseRealTibiaDatInLatestTibiaVersion is now ignored. Blackd Proxy will now always read .dat from selected Tibia folder" & vbCrLf & _
     " - Blackd Proxy no longer need having Tibia.dat in config folders." & vbCrLf & _
     " - Added support for modified Tibia 7.4 when it is really Tibia 7.72" & vbCrLf & _
     " - Added multiclient support for the following tibia versions: modified 7.4, 7.6 and 7.72" & vbCrLf & _
     " - Improved looting function: now it should be faster and less detectable."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 23.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.51"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 23.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug when opening Auction House in Tibia 9.5"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 23.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a problem with the parser of the new packet type A6"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 23.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a problem with the parser of the new packet type 9F"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 22.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.5"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 22.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - You can now set required soulpoints=0 in runemaker. That will allow you to make runes in oldest OT servers."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 22.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - We fixed and improved the trainer antilogout feature." & vbCrLf & _
     " - You can now redefine the serverlogout message in config.ini at the new variable serverLogoutMessage"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 22.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.46"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 22.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.45" & vbCrLf & _
     " - Disabled cavebot chaotic mode."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 22.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Blackd Proxy will keep executing even if it is unable to read Tibia.dat from Tibia folder. Anyways it still will give a safety warning in that case."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 22.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Blackd Proxy now avoid doing unnecessary closes of Tibia clients at Tibia.dat updates." & vbCrLf & _
     " - It now allows you to keep some config values unmodified through updates if you store them in configxxx\config.override.ini or settings.override.ini" & vbCrLf & _
     " - It now avoids tutorial hints popups." & vbCrLf & _
     " - Fixed a DirectX problem in Windows Vista/Windows 7" & vbCrLf & _
     " - Fixed a bug with private messages"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 22.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.44 (new tibia.dat - ninja updates)" & vbCrLf & _
     " - Added a new antiban feature: tibia.dat will be updated at ninja updates"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 22.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.44"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 22.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Added function exiva useitemwithamount:XX XX,A (will use an item with tile id XX XX only if there is A amount of that item in opened containers)" & vbCrLf & _
     " - Added new internal variable $useitemwithamount:XX XX,A$ that will do the same than above whenever you parse it. Returns_ 0=success, -1=failed" & vbCrLf & _
     " - Fixed invalid attack at cavebot so cavebot should be now a bit safer, but maybe not 100% safe yet. Untested."

    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 21.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug at GetProcessIDs"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 21.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.43"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 21.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a small problem in the loading of Tibia maps" & _
     " - Fixed a small problem launching Tibia from Blackd Proxy"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 21.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.42"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 21.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.41" & vbCrLf & _
     " - Blackd Proxy will now display a warning at start when loading an old config, and it will let you choose between resume the loading or load with newest config"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 21.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Added parser of packet type F9 (market feature)" & vbCrLf & _
     " - Updated parser of packet type 7A (trades with npc)"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 21.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.4" & vbCrLf & _
     " - We have set Old Loot mode enabled by default"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 21.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed debug problem when receiving invite to premium private channel in Tibia 9.31"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 21.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed multiclient address for Tibia 9.31"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 21.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.31" & vbCrLf & _
     " - Chaotic moves are now off by default. If you are going to get a ban anyways then at least now you will get more gold before that."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 20.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.20"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 20.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - The config for OT servers 7.7 was sucesfully tested."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 20.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Now you can use Blackd Proxy in multiclient in Tibia 9.10" & vbCrLf & _
     " - Included an experimental config for Tibia 7.7 It is not tested yet so it might not work."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 20.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.10" & vbCrLf & _
     " - Blackd Proxy is now free and open source"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 20.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a rare error at Tibia 7.6"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 20.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed npc trades in Tibia 9.00"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 20.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 9.00" & vbCrLf & _
     " - Warning: this bot will only work in the standard standalone C++ Tibia client. We will have to make a completely new bot for the flash client."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 20.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 8.74"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 20.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 8.73" & vbCrLf & _
     " - Fixed a bug that happened while opening containers with visual animations" & vbCrLf & _
     " - Blackd Proxy will now give extended details in debug files"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 20.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed command Exiva SayT:xxx for Tibia 8.72" & vbCrLf & _
     " - Fixed cavebot command SayInTrade for Tibia 8.72" & vbCrLf & _
     " - Fixed command Exiva Exp for Tibia 8.72" & vbCrLf & _
     " - If your Blackd Proxy account expired then you will be asked for your blackd login and password in next run"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 19.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug with the packet type AC in Tibia 8.72"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 19.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug that happened while moving backpacks in Tibia 8.72" & vbCrLf & _
     " - Fixed a bug with new packet type F3 in Tibia 8.72" & vbCrLf & _
     " - We reworked the runemaker for Tibia 8.72. Now it will not try to move runes to hand. Old runemaker mode will still be available if you play older Tibia versions." & vbCrLf & _
     " - Fixed a bug that happened in OT servers 8.61+"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 19.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed another bug at login"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 19.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug at login"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 19.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 8.72"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 19.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a very rare bug that happened with big stack of persons in same square"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 19.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - fastExiva should now accept exiva check so you can use fastExiva check x,y,z" & vbCrLf & _
     " - saymessage and sayintrade can now include the comma symbol (,)" & vbCrLf & _
     " - New variable! $lastchecktileid$ will return last checked tileid" & vbCrLf & _
     " - fixed a slight path bug with scripts that repeat the exact waypoint in different lines of the script"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 19.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated the command exiva sell:XX XX:N Packet format needed an update to be sure that it won't be automatically detected as cheat" & vbCrLf & _
     " - Be aware that now exiva sell will also count with equipped items!"
    
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 19.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Cavebot is now able to use ramps at Port Hope" & vbCrLf & _
     " - Bot is now able to autoeat all kind of food" & vbCrLf & _
     " - Fixed client freeze after some kinds of logout" & vbCrLf & _
     " - Updated tibia url in our broadcaster menu." & vbCrLf & _
     " - New command to do game inspect in a square near you: exiva check x,y,z" & vbCrLf & _
     " - New variable to check last game inspect result: $lastcheckresult$"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 19.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed exiva screenshot. Now it should work for everybody" & vbCrLf & _
     " - Fixed cavebot command DropLootonGround"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 18.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug that would crash the bot in a new packet type" & vbCrLf & _
     " - Now the cavebot will avoid looting things in stairs" & vbCrLf & _
     " - Now the cavebot will do less attempts of looting corpses of other players" & vbCrLf & _
     " - Improved debug system so we will be able to fix errors much faster in the future"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 18.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated tibia.dat after latest ninja update from cipsoft."

    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 18.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 8.71" & vbCrLf & _
     " - Fixed a bug with a system broadcast messages"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 18.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug with the autorecharge of mana."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 18.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug with big stack of things/persons"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 18.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Being trapped should not crash Blackd Proxy now." & vbCrLf & _
     " - New! : Autorecharge priorities for HP and Mana module: now this bot will be smart and it will only use the best recharge depending your current level of hp and mana. Now you can define strong heals for low HP and light heals for high HP."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 18.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Reporting other players should not crash Blackd Proxy now." & vbCrLf & _
     " - Fixed a bug that happened after swapping positions"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 18.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed cavebot command setBot AllowRepositionAtStart=0" & vbCrLf & _
     " - Fixed cavebot command setBot AllowRepositionAtTrap=0" & vbCrLf & _
     " - Added cavebot command setBot autoeatfood=0 (turn OFF autoeater)" & vbCrLf & _
     " - Added cavebot command setBot autoeatfood=1 (turn ON autoeater)"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 18.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug that happened after your player died." & vbCrLf & _
     " - Fixed a bug that happened while being attacked by a pk."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 18.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 8.7" & vbCrLf & _
     " - Fixed a small problem when dropping loot with cavebot. Now cavebot should do exact moves while performing that special operation." & vbCrLf & _
     " - setLootOff should now skip all the looting stage of the cavebot." & vbCrLf & _
     " - $cavebottimewithsametarget$ should now correctly display how much time cavebot is taking with same target. " & vbCrLf & _
     " - The value of config.ini AllowRepositionAtStart is now ignored. AllowRepositionAtStart will be enabled by default." & vbCrLf & _
     " - New command: setBot AllowRepositionAtStart=0 will force cavebot to start from first line." & vbCrLf & _
     " - New command: setBot AllowRepositionAtTrap=0 will force cavebot to do absolutely nothing on traps." & vbCrLf & _
     " - We should remember that you still can force cavebot to do always exact moves if you include SetChaoticMovesOff. Then it will move fast and good like old times. However that will add an important extra risk of ban."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 17.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Updated to work in Tibia 8.62" & vbCrLf & _
     " - We made a small fix in the new loot mode to avoid double click on same corpse." & vbCrLf & _
     " - New antidetection feature: current target will be now correctly marked with the red square."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 17.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed the command exiva #XX XX YY (moves item XX XX to slot YY)" & vbCrLf & _
     " - Fixed runemaker. It was not able to move items before. Now it should work perfectly again."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 17.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     " - Fixed a bug with party loot messages" & vbCrLf & _
     " - Fixed new loot mode (oldLootMode=0)" & vbCrLf & _
     " - New loot mode is the default mode again. It should be safer. If you want to use old loot mode anyways then just set oldLootMode=1"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 17.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Fixed a critical bug that was disabling cheats at Tibia 8.61 when gm talked in channel" & vbCrLf & _
     "- Now by default setDontRetryAttacks is enabled: Cavebot won't spam attacks over same monster again and again. It will only attack it once (more human)."
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 17.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Fixed a critical bug that was disabling cheats at Tibia 8.61 when someone talked in help channel"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 17.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Fixed a bug for OT servers 8.6"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 17.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Updated to work in Tibia 8.61"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 17.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Updated Tibia.dat (ninja update of cipsoft)"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 17.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Updated to work in Tibia 8.6" & vbCrLf & _
     "- Improved security against automatic bans" & vbCrLf & _
     "- Old loot mode is again the default loot mode" & vbCrLf & _
     "- Be warned that Blackd Proxy might be still detected as cheat so only use it if you dont care about ban"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 17.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Updated to work in Tibia 8.57" & vbCrLf & _
     "- Be warned that Blackd Proxy is still detected as cheat so only use it if you dont care about ban"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 16.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Updated to work in Tibia 8.56" & vbCrLf & _
     "- Be warned that Blackd Proxy is still detected as cheat so only use it if you dont care about ban"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 16.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Updated to work in Tibia 8.55" & vbCrLf & _
     "- Be warned that Blackd Proxy is still detected as cheat so only use it if you dont care about ban"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 16.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- New features at tray icon: hide/show Tibia clients"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 16.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Updated tibia.dat" & vbCrLf & _
     "- Fixed a rare loading problem"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 16.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Now compatible with Tibia 8.54" & vbCrLf & _
     "- Fixed Broadcaster when multiclient option is disabled" & vbCrLf & _
     "- Cavebot will now loot nearer corpses first" & vbCrLf & _
     "- Cavebot will not retry looting after message You are not the owner." & vbCrLf & _
     "- exiva pause- now will not pause any hp/mana recharge." & _
     "- Blackd Proxy will now run as Tibia.exe by default." & _
     "- Blackd Proxy will rename patch.exe to patch.bck to prevent possible cheat detections from that exe."
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 16.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Fixed the multiclient feature of the Broadcaster" & vbCrLf & _
     "- Small improvements in login functions"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 16.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Reworked the stack system to allow player/creatures at ground level"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 16.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Fixed a problem with hpmana module"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 16.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Fixed an important problem with cavebot in Tibia 8.53" & vbCrLf & _
     "- Protected broadcaster against unexpected errors"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 16.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Now compatible with Tibia 8.53" & vbCrLf & _
     "- Broadcaster now compatible with Firefox" & vbCrLf & _
     "- Broadcaster now allow mc by turns"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 15.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Added limit randomizations to runemaker so now it won't make rune exactly after having enough mana" & vbCrLf & _
     "- Added limit randomizations to hpmana so now it won't recharge heal mana exactly at same level" & vbCrLf & _
     "- Added a broadcaster to send private messages to a list of players"
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 15.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Fixed an important bug that was crashing Blackd Proxy with Tibia 8.52"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 15.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Minimum update for Tibia 8.52"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 15.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Fixed a bug that only happened in players with powers, like tutors and gms" & vbCrLf & _
     "- Updated to latest Tibia.dat file 8.5 in blackd proxy folders"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 15.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Minimum update for Tibia 8.5"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 15.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Bug fixed: uploaded correct config842\config.ini"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 15.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Bug fixed: guild invite messages should not disable cheats now" & vbCrLf & _
     "- Bug fixed: LIGHT_AMOUNT and LIGHT_NOP fixed for Tibia 8.42 : now xray should work at full light again" & vbCrLf & _
     "- Bug fixed: Looter should work for some strange OT servers 7.6 in combination with setDontRetryAttacks" & vbCrLf & _
     "- Antiban feature added: Trainer module now can randomize actions to look more human" & vbCrLf & _
     "- Antiban feature added: New looting function. You still can use old (risky) looting with setBot OldLootMode=1" & vbCrLf & _
     "- You can customize some new internal looting variables with new cavebot command setBot, by default..." & vbCrLf & _
     "- MINDELAYTOLOOT=1000 ms (wait at least 1 second before looting)" & vbCrLf & _
     "- MAXTIMEINLOOTQUEUE=60000 ms (forget corpse if not looted after 60seconds)" & vbCrLf & _
     "- MAXTIMETOREACHCORPSE=10000 ms (cancel loot if reaching corpse takes more than 10 seconds)" & vbCrLf & _
     "- OldLootMode=0 (old loot mode is disabled, new mode is enabled by default)"
 
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 15.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Minimum update for Tibia 8.42"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 15.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Stealth window won't cast chat by default. It still can be reenabled to do that." & vbCrLf & _
     "- Warbot autohealer timer is now randomized." & vbCrLf & _
     "- Warbot autohealer should now work correctly on OT servers 7.6" & vbCrLf & _
     "- You now can add special gm names in subfolder \specialgm\names.txt Any name on this list will be considered a gm for all internal operations."
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 15.0" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Autoeat food problem solved. Runemaker should eat food again correctly."
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 14.9" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Manual reconnection problem solved." & vbCrLf & _
     "- We forced all cheats to be stoped until the connection is completely established" & vbCrLf & _
     "- We forced a minimum delay after manual reconnection to a gameserver." & vbCrLf & _
     "- exiva view requires reconnection so we disabled it. Use exiva xray instead." & vbCrLf & _
     "- Ports settings will be ignored since Tibia 8.41+ Now we decided to use random available ports instead. That way we avoid some strange errors at start."
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 14.8" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Basic support for multiclient. It will just work, however it might give problems with manual reconnection through control+L .We recommend to close desired clients instead doing logout."
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 14.7" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Minimum changes to support Tibia 8.41 . At this moment we have no time for more changes." & vbCrLf & _
     "- Invisible creatures  are completely invisible now and they can't be detected anymore. Avoid zones with invisible creatures!!" & vbCrLf & _
     "- Autorelog, magebomb and all automatic connection functions have been disabled. It is no longer possible since Tibia 8.41!"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 14.6" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- We fixed a bug that allowed sending spaces through stealth command. Now it is no longer possible." & vbCrLf & _
     "- Fixed a backpack problem for a very strange OT server 7.6"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 14.5" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- We added randomization to the conds module. It should not be detectable from now" & vbCrLf & _
     "- Our reports say that cheating with Blackd Proxy should be safe as long you don't change tibia title and you don't repeat same action with a fixed timer. Randomization in timers is important."
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 14.4" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Improved timers" & vbCrLf & _
     "- Fixed cavebot problems with yalahar stairs" & vbCrLf & _
     "- new command: floor change by tibia memory modification (very risky now, but requested by lot of people) :" & vbCrLf & _
     "exiva xray +1 : view 1 floor below" & vbCrLf & _
     "exiva xray -1 : view 1 floor above" & vbCrLf & _
     "exiva xray +2 : view 2 floors below" & vbCrLf & _
     "exiva xray -2 : view 2 floors above"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 14.3" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "- Fixed setChaoticMovesOFF"
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 14.2" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "Second pack of anti detection features!" & vbCrLf & _
     "- Fixed a problem with exiva close" & vbCrLf & _
     "- Fixed a small problem with Windows Vista" & vbCrLf & _
     "- Added chaos to cavebot waypoints! setChaoticMovesON will be enabled by default." & vbCrLf & _
     "- New cavebot command: setChaoticMovesOFF : disable randomization for waypoints" & vbCrLf & _
     "- New cavebot command: setChaoticMovesON : (re)enable randomization for waypoints" & vbCrLf & _
     "- Improved the pathing functions of cavebot. Now the movement should be more human."
     
    strMsg = strMsg & vbCrLf & vbCrLf & "Blackd Proxy 14.1" & vbCrLf & _
     "----------------------------" & vbCrLf & _
     "After the mass bans we have decided to focus in developing stealth features:" & vbCrLf & vbCrLf & _
     "- After first load Blackd Proxy will be removed from installed programs list." & vbCrLf & _
     "- You will get a random process name for your Blackd Proxy instance." & vbCrLf & _
     "- Exp in title will be displayed in the new stealth window" & vbCrLf & _
     "- Blackd Proxy system messages will be displayed in the new stealth window" & vbCrLf & _
     "- Blackd Proxy commands will be executed from the new stealth window" & vbCrLf & _
     "- HPMana module: timer is now completely randomized for each mc" & vbCrLf & _
     "- HPMana module: heal recharge and mana recharge won't happen at same moment. They will alternate turns if both things are required." & vbCrLf & _
     "- Runemaker module: a chaos timer will now randomize the delay between each step of runemaking." & vbCrLf & _
     "- Cavebot module: timer is now completely randomized for each mc" & vbCrLf & _
     "- Detection of suspicious scans from Cipsoft: some checks have been added all over the code in order to analyze any strange packet received from Cipsoft." & vbCrLf & _
     "- New variable: $randomnumber:A>B$ will get you a random number from A to B. This will be needed as paranoid people now might want to randomize some paths in cavebot scripts." & vbCrLf & _
     vbCrLf & "Please consider that nobody knows how the detection is done at this moment, so cheating is still a bit risky right now. Cheat at your own responsability!"
 Me.txtBoard.Text = strMsg

End Sub

Private Sub Form_Resize()
  If frmNews.WindowState <> vbMinimized Then
    If frmNews.ScaleHeight < 3000 Then
      frmNews.Height = 3000
    End If
    If frmNews.ScaleWidth < 5800 Then
      frmNews.Width = 5800
    End If
    txtBoard.Height = frmNews.ScaleHeight - 1300
    txtBoard.Width = frmNews.ScaleWidth - 200
    cmdOk.Top = frmNews.ScaleHeight - 480
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub
