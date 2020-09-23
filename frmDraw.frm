VERSION 5.00
Begin VB.Form frmDraw 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drawing area"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   413
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   Begin VB.Line lnLine 
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   160
      X2              =   160
      Y1              =   104
      Y2              =   136
   End
   Begin VB.Shape shpShape 
      BorderStyle     =   3  'Dot
      Height          =   495
      Left            =   2520
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DrawingStyle As STYLE
Public RGBColor As Long, InitTime As Long
Public Recording As Boolean
Dim stX As Long, stY As Long
Dim Drawing As Boolean

Private Sub Form_Load()
  DrawingStyle = dsNone
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim hBrush As Long

  If Button <> 1 Or DrawingStyle = dsNone Then Exit Sub
  stX = X
  stY = Y
  Drawing = True
  If DrawingStyle = dsFloodFill Then
    Drawing = False
    hBrush = CreateSolidBrush(RGBColor)
    SelectObject hdc, hBrush
    ExtFloodFill hdc, stX, stY, Point(X, Y), 1
    DeleteObject hBrush
    Refresh
    If Recording Then SaveData stX, stY, 0, 0
  ElseIf DrawingStyle = dsLine Then
    With lnLine
      .X1 = stX
      .Y1 = stY
      .X2 = stX
      .Y2 = stY
      .Visible = True
    End With
  ElseIf DrawingStyle = dsRectangle Then
    With shpShape
      .Left = stX
      .Top = stY
      .Width = 0
      .Height = 0
      .Shape = 0
      .Visible = True
    End With
  ElseIf DrawingStyle = dsEllipse Then
    With shpShape
      .Left = stX
      .Top = stY
      .Width = 0
      .Height = 0
      .Shape = 2
      .Visible = True
    End With
  ElseIf DrawingStyle = dsPen Then
    With Me
      .CurrentX = stX
      .CurrentY = stY
    End With
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim TmpLng As Long

  If Not Drawing Then Exit Sub
  If DrawingStyle = dsLine Then
    If Recording Then SaveData stX, stY, X, Y, dsLineDoted
    With lnLine
      .X2 = X
      .Y2 = Y
    End With
  ElseIf DrawingStyle = dsPen Then
    If Recording Then SaveData CurrentX, CurrentY, X, Y
    Me.Line -(X, Y), RGBColor
  ElseIf DrawingStyle = dsRectangle Then
    If Recording Then SaveData stX, stY, X, Y, dsRectangleDoted
    With shpShape
      If Y - stY >= 0 And X - stX >= 0 Then
        .Top = stY
        .Left = stX
        .Height = Y - stY + 1
        .Width = X - stX + 1
      ElseIf Y - stY < 0 And X - stX >= 0 Then
        .Top = Y
        .Left = stX
        .Height = stY - Y + 1
        .Width = X - stX + 1
      ElseIf Y - stY >= 0 And X - stX < 0 Then
        .Top = stY
        .Left = X
        .Height = Y - stY + 1
        .Width = stX - X + 1
      ElseIf Y - stY < 0 And X - stX < 0 Then
        .Top = Y
        .Left = X
        .Height = stY - Y + 1
        .Width = stX - X + 1
      End If
    End With
  ElseIf DrawingStyle = dsEllipse Then
    If Recording Then SaveData stX, stY, X, Y, dsEllipseDoted
    With shpShape
      If Y - stY >= 0 And X - stX >= 0 Then
        .Top = stY
        .Left = stX
        .Height = Y - stY + 1
        .Width = X - stX + 1
      ElseIf Y - stY < 0 And X - stX >= 0 Then
        .Top = Y
        .Left = stX
        .Height = stY - Y + 1
        .Width = X - stX + 1
      ElseIf Y - stY >= 0 And X - stX < 0 Then
        .Top = stY
        .Left = X
        .Height = Y - stY + 1
        .Width = stX - X + 1
      ElseIf Y - stY < 0 And X - stX < 0 Then
        .Top = Y
        .Left = X
        .Height = stY - Y + 1
        .Width = stX - X + 1
      End If
    End With
  ElseIf DrawingStyle = dsPicker Then
    TmpLng = Point(X, Y)
    frmTools.picSelColor.BackColor = TmpLng
    RGBColor = TmpLng
  End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim ResX As Long, ResY As Long, MidX As Long, MidY As Long

  On Error Resume Next
  If DrawingStyle = dsNone Then Exit Sub
  If Recording And DrawingStyle <> dsPen And DrawingStyle <> dsPicker And DrawingStyle <> dsFloodFill Then SaveData stX, stY, X, Y
  Drawing = False
  If DrawingStyle = dsEllipse Then
    ResX = Abs(stX - X)
    ResY = Abs(stY - Y)
    X = X - stX
    Y = Y - stY
    MidX = stX + X / 2
    MidY = stY + Y / 2
    Circle (MidX, MidY), IIf(ResX > ResY, ResX / 2, ResY / 2), RGBColor, , , ResY / ResX
    shpShape.Visible = False
  ElseIf DrawingStyle = dsRectangle Then
    Line (stX, stY)-(X, Y), RGBColor, B
    shpShape.Visible = False
  ElseIf DrawingStyle = dsLine Then
    Line (stX, stY)-(X, Y), RGBColor
    lnLine.Visible = False
  End If
End Sub

Sub LoadData(ByVal Index As Long)
  On Error Resume Next
  With DrawData(Index)
    CopyMemory RGBColor, .RGBColor(0), 3
    DrawingStyle = .DrawType
    DrawWidth = .BrushSize
    stX = .P1
    stY = .P2
    Drawing = True
    If .DrawType = dsPen Then
      CurrentX = stX
      CurrentY = stY
      Form_MouseMove 1, 0, (.P3), (.P4)
    ElseIf .DrawType = dsEllipseDoted Then
      DrawingStyle = dsEllipse
      Form_MouseDown 1, 0, (stX), (stY)
      Form_MouseMove 1, 0, (.P3), (.P4)
    ElseIf .DrawType = dsLineDoted Then
      DrawingStyle = dsLine
      Form_MouseDown 1, 0, (stX), (stY)
      Form_MouseMove 1, 0, (.P3), (.P4)
    ElseIf .DrawType = dsRectangleDoted Then
      DrawingStyle = dsRectangle
      Form_MouseDown 1, 0, (stX), (stY)
      Form_MouseMove 1, 0, (.P3), (.P4)
    ElseIf .DrawType = dsFloodFill Then
      DrawingStyle = dsFloodFill
      Form_MouseDown 1, 0, (stX), (stY)
    Else
      Form_MouseUp 0, 0, (.P3), (.P4)
    End If
    Drawing = False
  End With
End Sub

Private Sub SaveData(ByVal P1 As Long, ByVal P2 As Long, ByVal P3 As Long, ByVal P4 As Long, Optional ByVal Styling As STYLE = -1)
  Dim TmpLng As Long

  TmpLng = AddNewEntry()
  With DrawData(TmpLng)
    If Styling <> -1 Then .DrawType = Styling Else .DrawType = DrawingStyle
    .BrushSize = DrawWidth
    .P1 = P1
    .P2 = P2
    .P3 = P3
    .P4 = P4
    CopyMemory .RGBColor(0), RGBColor, 3
    '.RGBColor = RGBColor
    .WaitTime = GetTickCount() - InitTime
    InitTime = GetTickCount()
  End With
  Me.Caption = "Recording... " & MaxBound * Len(DrawData(0)) & " bytes recorded"
End Sub
