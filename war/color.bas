Option Explicit


'Type PALETTEENTRY
'    peRed As String * 1
'    peGreen As String * 1
'    peBlue As String * 1
'    peFlags As String * 1
'End Type
    
    
Type logpalette
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(256) As Long ' Array length is arbitrary; may be changed
End Type

Declare Function GetSystemPaletteEntries Lib "GDI" (ByVal h As Integer, ByVal start As Integer, ByVal entries As Integer, pal As Long) As Integer
Declare Function CreatePalette Lib "GDI" (lpLogPalette As logpalette) As Integer
Declare Function SetPaletteEntries Lib "GDI" (ByVal hPalette As Integer, ByVal wStartIndex As Integer, ByVal wNumEntries As Integer, lpPaletteEntries As Long) As Integer
Declare Function RealizePalette Lib "User" (ByVal hDC As Integer) As Integer
Declare Function SelectPalette Lib "User" (ByVal hDC As Integer, ByVal hPalette As Integer, ByVal bForceBackground As Integer) As Integer




