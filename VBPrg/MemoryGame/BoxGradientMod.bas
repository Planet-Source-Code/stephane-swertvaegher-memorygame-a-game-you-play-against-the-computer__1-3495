Attribute VB_Name = "BoxGradientMod"
Option Explicit

Public Sub BoxGradient(OBJ As Object, R%, G%, B%, RStep%, GStep%, BStep%, Direc As Boolean)
Dim s%, xpos%, ypos%
OBJ.ScaleMode = 3 'pixel
If Direc = True Then
RStep% = -RStep%
GStep% = -GStep%
BStep% = -BStep%
End If
DoBox:
s% = s% + 1
If xpos% < Int(OBJ.ScaleWidth / 2) Then xpos% = s%
If ypos% < Int(OBJ.ScaleHeight / 2) Then ypos% = s%
OBJ.Line (xpos%, ypos%)-(OBJ.ScaleWidth - xpos%, OBJ.ScaleHeight - ypos%), RGB(R%, G%, B%), B
R% = R% - RStep%
If R% < 0 Then R% = 0
If R% > 255 Then R% = 255
G% = G% - GStep%
If G% < 0 Then G% = 0
If G% > 255 Then G% = 255
B% = B% - BStep%
If B% < 0 Then B% = 0
If B% > 255 Then B% = 255
If xpos% = Int(OBJ.ScaleWidth / 2) And ypos% = Int(OBJ.ScaleHeight / 2) Then Exit Sub
GoTo DoBox
End Sub
