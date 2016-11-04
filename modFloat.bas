Attribute VB_Name = "modFloat"
Option Explicit

Private Type TSingle
  Sng As Single
End Type

Private Type TLong
  Lng As Long
End Type

Public Function Long2Float(ByVal value As Long) As Single
Dim s As TLong, d As TSingle
  s.Lng = value
  LSet d = s
  Long2Float = d.Sng
End Function

Public Function Float2Long(ByVal value As Single) As Long
Dim s As TSingle, d As TLong
  s.Sng = value
  LSet d = s
  Float2Long = d.Lng
End Function
