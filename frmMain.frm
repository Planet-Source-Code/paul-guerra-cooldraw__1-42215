VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Drawer"
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load"
      End
      Begin VB.Menu mnuHash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuDraw 
      Caption         =   "&Draw"
      Begin VB.Menu mnuRecord 
         Caption         =   "&Record"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "&Play"
      End
      Begin VB.Menu mnuClearDraw 
         Caption         =   "Clear &drawing"
      End
      Begin VB.Menu mnuClearMem 
         Caption         =   "Clear &memory"
      End
      Begin VB.Menu mnuHash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTiming 
         Caption         =   "&Ignore timing"
      End
      Begin VB.Menu mnuDoubleSpeed 
         Caption         =   "&x 2 Play"
      End
      Begin VB.Menu mnuCompress 
         Caption         =   "&Compress"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const File = "Draw.dat", TempFile = "Draw.tmp"

Private Sub mnuClearDraw_Click()
  frmDraw.Cls
End Sub

Private Sub mnuClearMem_Click()
  ClearDraw
End Sub

Private Sub mnuCompress_Click()
  mnuCompress.Checked = Not mnuCompress.Checked
End Sub

Private Sub mnuDoubleSpeed_Click()
  mnuDoubleSpeed.Checked = Not mnuDoubleSpeed.Checked
End Sub

Private Sub mnuDraw_Click()
  Dim TmpBool As Boolean

  TmpBool = frmDraw.Recording
  mnuStop.Enabled = TmpBool
  mnuRecord.Enabled = Not TmpBool
  mnuClearMem.Enabled = Not TmpBool
  mnuClearDraw.Enabled = Not TmpBool
  mnuPlay.Enabled = MaxBound <> 0 And Not TmpBool
End Sub

Private Sub mnuExit_Click()
  End
End Sub

Private Sub mnuLoad_Click()
  Dim DebeBorrar As Boolean

  Open File For Binary As 1
  Get 1, , MaxBound
  If MaxBound Then
    Close 1
    DecompressFile TempFile, File
    Open TempFile For Binary As 1
    DebeBorrar = True
  End If
  Get 1, 5, MaxBound
  If MaxBound Then
    ReDim DrawData(MaxBound - 1)
    Get 1, , DrawData
  Else
    Erase DrawData
  End If
  Close 1
  If DebeBorrar Then Kill TempFile
  MsgBox "Drawing loaded", vbInformation
End Sub

Private Sub mnuPlay_Click()
  PlayDraw Not mnuTiming.Checked, mnuDoubleSpeed.Checked
End Sub

Private Sub mnuRecord_Click()
  With frmDraw
    .Cls
    PlayDraw False, False
    .Recording = True
    .Caption = "Recording..."
    .InitTime = GetTickCount()
  End With
End Sub

Private Sub mnuSave_Click()
  Open File For Binary As 1
  Put 1, , CLng(0)
  Put 1, , MaxBound
  Put 1, , DrawData
  Close 1
  If mnuCompress.Checked Then CompressFile File
  MsgBox "Drawing saved to " & File, vbInformation
End Sub

Private Sub mnuStop_Click()
  With frmDraw
    .Recording = False
    .Caption = "Drawing area"
  End With
End Sub

Private Sub mnuTiming_Click()
  mnuTiming.Checked = Not mnuTiming.Checked
End Sub
