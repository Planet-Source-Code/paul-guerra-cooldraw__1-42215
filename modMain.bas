Attribute VB_Name = "modMain"
Option Explicit
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal dwBytes As Long)
Declare Sub Sleep Lib "kernel32" (ByVal dwMs As Long)
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Enum STYLE
  dsNone = -1
  dsPen
  dsLine
  dsEllipse
  dsRectangle
  dsFloodFill
  dsPicker
  dsLineDoted
  dsEllipseDoted
  dsRectangleDoted
End Enum
Type DRAWENTRY
  DrawType As Byte 'STYLE
  BrushSize As Byte
  P1 As Integer
  P2 As Integer
  P3 As Integer
  P4 As Integer
  RGBColor(2) As Byte 'RGBColor As Long
  WaitTime As Integer
End Type
Public DrawData() As DRAWENTRY
Public MaxBound As Long

Sub Main()
  frmMain.Show
  frmTools.Left = 7020
  frmTools.Show
  frmDraw.Top = 645
  frmDraw.Show
End Sub

Function AddNewEntry() As Long
  ReDim Preserve DrawData(MaxBound)
  AddNewEntry = MaxBound
  MaxBound = MaxBound + 1
End Function

Sub ClearDraw()
  MaxBound = 0
  Erase DrawData
End Sub

Sub PlayDraw(ByVal UseTiming As Boolean, ByVal DobSpeed As Boolean)
  Dim i As Long, BrushSize As Long, ToolSel As Long

  With frmDraw
    BrushSize = .DrawWidth
    ToolSel = .DrawingStyle
    If UseTiming Then
      frmMain.Hide
      frmTools.Hide
      .Enabled = False
      .Caption = "Playing..."
    End If
    .DrawingStyle = dsNone
    .Cls
    For i = 0 To MaxBound - 1
      If UseTiming Then
        DoEvents
        If DobSpeed Then
          Sleep DrawData(i).WaitTime \ 2
        Else
          Sleep DrawData(i).WaitTime
        End If
        .LoadData i
        .Caption = "Playing... " & (i + 1) * Len(DrawData(0)) & " bytes played"
      ElseIf DrawData(i).DrawType < dsLineDoted Then
        .LoadData i
      End If
    Next i
    .DrawingStyle = ToolSel
    .DrawWidth = BrushSize
    frmDraw.RGBColor = frmTools.picSelColor.BackColor
    If UseTiming Then
      frmMain.Show
      frmTools.Show
      .Enabled = True
      .Caption = "Drawing area"
    End If
  End With
End Sub
