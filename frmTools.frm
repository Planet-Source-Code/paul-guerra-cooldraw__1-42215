VERSION 5.00
Begin VB.Form frmTools 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tools"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   223
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   48
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTool 
      Height          =   375
      Index           =   5
      Left            =   360
      MaskColor       =   &H000000FF&
      Picture         =   "frmTools.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdTool 
      Height          =   375
      Index           =   4
      Left            =   0
      MaskColor       =   &H000000FF&
      Picture         =   "frmTools.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdVarTam 
      Caption         =   "-"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton cmdVarTam 
      Caption         =   "+"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox picSelColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   690
      TabIndex        =   7
      Top             =   2640
      Width           =   720
   End
   Begin VB.PictureBox picColor 
      AutoSize        =   -1  'True
      Height          =   1500
      Left            =   0
      Picture         =   "frmTools.frx":0204
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   6
      Top             =   1110
      Width           =   735
   End
   Begin VB.CommandButton cmdTool 
      Height          =   375
      Index           =   3
      Left            =   360
      MaskColor       =   &H000000FF&
      Picture         =   "frmTools.frx":069B
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdTool 
      Height          =   375
      Index           =   2
      Left            =   0
      MaskColor       =   &H000000FF&
      Picture         =   "frmTools.frx":079D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdTool 
      Height          =   375
      Index           =   1
      Left            =   360
      MaskColor       =   &H000000FF&
      Picture         =   "frmTools.frx":089F
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdTool 
      Height          =   375
      Index           =   0
      Left            =   0
      MaskColor       =   &H000000FF&
      Picture         =   "frmTools.frx":09A1
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Label lblTam 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      ToolTipText     =   "Brush size"
      Top             =   3000
      Width           =   255
   End
End
Attribute VB_Name = "frmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Coloring As Boolean
Dim LastSel As Long

Private Sub cmdTool_Click(Index As Integer)
  If LastSel <> -1 Then cmdTool(LastSel).UseMaskColor = True
  If LastSel = Index Then
    LastSel = -1
  Else
    LastSel = Index
    cmdTool(Index).UseMaskColor = False
  End If
  frmDraw.DrawingStyle = LastSel
End Sub

Private Sub cmdVarTam_Click(Index As Integer)
  Dim BrushSize As Long

  BrushSize = frmDraw.DrawWidth
  If Index Then
    If BrushSize < 50 Then BrushSize = BrushSize + 1
  Else
    If BrushSize > 1 Then BrushSize = BrushSize - 1
  End If
  frmDraw.DrawWidth = BrushSize
  lblTam.Caption = BrushSize
End Sub

Private Sub Form_Load()
  LastSel = -1
End Sub

Private Sub picColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Coloring = True
  picColor_MouseMove Button, Shift, X, Y
End Sub

Private Sub picColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim TmpLng As Long

  If Not Coloring Then Exit Sub
  TmpLng = picColor.Point(X, Y)
  If TmpLng = -1 Then Exit Sub
  frmDraw.RGBColor = TmpLng
  picSelColor.BackColor = TmpLng
End Sub

Private Sub picColor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Coloring = False
End Sub
