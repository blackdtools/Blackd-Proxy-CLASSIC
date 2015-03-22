VERSION 5.00
Begin VB.Form frmBackpacks 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backpacks"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5100
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmBackpacks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSearchItems 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Search for this item"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtTileID 
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Text            =   "1A 0C"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdateItems 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Reload items"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdateBP 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Reload bp list"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox cmbBackpackID 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "0 - closed"
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblResult 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   4575
   End
   Begin VB.Label lblContains 
      BackColor       =   &H00000000&
      Caption         =   "Contains:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblItems 
      BackColor       =   &H00000000&
      Caption         =   "Not updated"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Label lblCap 
      BackColor       =   &H00000000&
      Caption         =   "Container cap: 0 / 0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "frmBackpacks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Public Function GetFirstFreeBpID(idConnection As Integer) As Byte
  ' Return the first container ID not used by the client of that idConnection
  Dim res As Byte
  Dim i As Byte
  res = &HFF
  For i = 1 To HIGHEST_BP_ID
    If Backpack(idConnection, i).open = False Then
      res = i
      Exit For
    End If
  Next i
  GetFirstFreeBpID = res
End Function
Public Sub UpdateBPlist()
  ' Update the list of open containers for selected client
  Dim j As Long
  cmbBackpackID.Clear
  If mapIDselected > 0 Then
    For j = 0 To HIGHEST_BP_ID
      If Backpack(mapIDselected, j).open = True Then
        cmbBackpackID.AddItem CStr(j) & " - " & Backpack(mapIDselected, j).name, j
      Else
        cmbBackpackID.AddItem CStr(j) & " - " & Backpack(mapIDselected, j).name, j
      End If
    Next j
    cmbBackpackID.Text = cmbBackpackID.List(bpIDselected)
  Else
    For j = 0 To HIGHEST_BP_ID
      cmbBackpackID.AddItem CStr(j) & " - ", j
    Next j
    cmbBackpackID.Text = cmbBackpackID.List(bpIDselected)
  End If
End Sub

Public Sub UpdateItemList()
  Dim i As Integer
  Dim tileID As Long
  Dim currentCap As Long
  Dim currentNumItems As Long
  Dim currentList As String
  Dim currentItem As String
  If mapIDselected > 0 Then
  currentCap = Backpack(mapIDselected, bpIDselected).cap
  currentNumItems = Backpack(mapIDselected, bpIDselected).used
  lblCap.Caption = "Container cap: " & CStr(currentNumItems) & " / " & CStr(currentCap)
  currentList = ""
  For i = 0 To currentNumItems - 1
    tileID = GetTheLong(Backpack(mapIDselected, bpIDselected).item(i).t1, _
     Backpack(mapIDselected, bpIDselected).item(i).t2)

    If DatTiles(tileID).haveExtraByte Then
      currentItem = GoodHex(Backpack(mapIDselected, bpIDselected).item(i).t1) & _
              " " & GoodHex(Backpack(mapIDselected, bpIDselected).item(i).t2) & _
              "x" & CStr(CLng(Backpack(mapIDselected, bpIDselected).item(i).t3)) & " ; "
    Else
      currentItem = GoodHex(Backpack(mapIDselected, bpIDselected).item(i).t1) & _
              " " & GoodHex(Backpack(mapIDselected, bpIDselected).item(i).t2) & " ; "
              
    End If
    currentList = currentList & currentItem
  Next i
  lblItems.Caption = currentList
  End If
End Sub
Private Sub cmbBackpackID_Click()
  bpIDselected = cmbBackpackID.ListIndex
  UpdateItemList
End Sub

Private Sub cmdSearchItems_Click()
  Dim lon As Long
  Dim b1 As Byte
  Dim b2 As Byte
  Dim s1 As Byte
  Dim s2 As Byte
  Dim pos As Long
  Dim strRes As String
  Dim res As TypeSearchItemResult2
  #If FinalMode Then
  On Error GoTo exitS
  #End If
  s1 = FromHexToDec(Mid(txtTileID.Text, 1, 1))
  s2 = FromHexToDec(Mid(txtTileID.Text, 2, 1))
  b1 = (s1 * 16) + s2
  s1 = FromHexToDec(Mid(txtTileID.Text, 4, 1))
  s2 = FromHexToDec(Mid(txtTileID.Text, 5, 1))
  b2 = (s1 * 16) + s2
  res = SearchItem(mapIDselected, b1, b2)
  If res.foundcount > 0 Then
    lblResult.Caption = "Found " & CStr(res.foundcount) & " items. Last at : bp " & _
     CStr(CLng(res.bpID)) & " , slot " & CStr(CLng(res.slotID))
  Else
    lblResult.Caption = "None found"
  End If
  Exit Sub
exitS:
  MsgBox "Bad format"
End Sub

Private Sub cmdUpdateBP_Click()
  UpdateBPlist
End Sub

Private Sub cmdUpdateItems_Click()
  UpdateItemList
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub



Public Sub AddItem(clientID As Integer, bpID As Long, t1 As Byte, t2 As Byte, t3 As Byte, Optional t4 As Byte = &H0)
  'add item to a bp
  Dim i As Integer
  Dim cap As Long
  cap = Backpack(clientID, bpID).cap
  For i = cap - 1 To 1 Step -1
    Backpack(clientID, bpID).item(i).t1 = _
     Backpack(clientID, bpID).item(i - 1).t1
    Backpack(clientID, bpID).item(i).t2 = _
     Backpack(clientID, bpID).item(i - 1).t2
    Backpack(clientID, bpID).item(i).t3 = _
     Backpack(clientID, bpID).item(i - 1).t3
    Backpack(clientID, bpID).item(i).t4 = _
     Backpack(clientID, bpID).item(i - 1).t4
  Next i
  Backpack(clientID, bpID).item(0).t1 = t1
  Backpack(clientID, bpID).item(0).t2 = t2
  Backpack(clientID, bpID).item(0).t3 = t3
  Backpack(clientID, bpID).item(0).t4 = t4
  Backpack(clientID, bpID).used = Backpack(clientID, bpID).used + 1
End Sub
Public Sub UpdateItem(clientID As Integer, bpID As Long, slot As Long, t1 As Byte, t2 As Byte, t3 As Byte, Optional t4 As Byte = &H0)
  'update item in a given slot of a bp
  If slot <= HIGHEST_ITEM_BPSLOT Then
    Backpack(clientID, bpID).item(slot).t1 = t1
    Backpack(clientID, bpID).item(slot).t2 = t2
    Backpack(clientID, bpID).item(slot).t3 = t3
    Backpack(clientID, bpID).item(slot).t4 = t4
  End If
End Sub

Public Sub RemoveItem(clientID As Integer, bpID As Long, slot As Long)
  'remove item from a given slot of a bp
 Dim i As Integer
  Dim cap As Long
  Dim aRes As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  cap = Backpack(clientID, bpID).cap
  For i = slot To cap - 2
    Backpack(clientID, bpID).item(i).t1 = _
     Backpack(clientID, bpID).item(i + 1).t1
    Backpack(clientID, bpID).item(i).t2 = _
     Backpack(clientID, bpID).item(i + 1).t2
    Backpack(clientID, bpID).item(i).t3 = _
     Backpack(clientID, bpID).item(i + 1).t3
  Next i
  If (cap - 1) > 0 Then
    Backpack(clientID, bpID).item(cap - 1).t1 = &H0
    Backpack(clientID, bpID).item(cap - 1).t2 = &H0
    Backpack(clientID, bpID).item(cap - 1).t3 = &H0
    Backpack(clientID, bpID).used = Backpack(clientID, bpID).used - 1
  Else
  LogOnFile "errors.txt", "Warning at Removeitem (" & clientID & ", " & bpID & "," & slot & " )   : Container with cap 0!"
  End If
  Exit Sub
goterr:
  LogOnFile "errors.txt", "Unexpected error at Removeitem (" & clientID & ", " & bpID & "," & slot & " ) Cap=" & cap & "  : " & Err.Description
  aRes = GiveGMmessage(clientID, "Unexpected error in backpack module, please report to blackd. Received call: RemoveItem(" & clientID & "," & bpID & "," & slot & ") Error description: " & Err.Description, "Blackdproxy")
End Sub

Public Function totalbpsOpen(clientID As Integer) As Long
  Dim i As Integer
  Dim res As Long
  res = 0
  For i = 0 To HIGHEST_BP_ID
    If Backpack(clientID, i).open = True Then
      res = res + 1
    End If
  Next i
  totalbpsOpen = res
End Function
