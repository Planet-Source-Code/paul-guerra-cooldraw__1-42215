Attribute VB_Name = "modCompression"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal dwBytes As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Const RedimInterval = 16384, MaxDict = 766, MaxEntrySize = 600
Const Signal = 1023, BlockSize = 2800, BlockStop = 2801, Limit = 100
Type DICTORYENTRY
  Entry() As Byte
  Size As Long
End Type
Dim Dictory(MaxDict) As DICTORYENTRY
Dim DestArray() As Byte, OutBuff(4) As Byte
Dim BytePos As Long, CurSize As Long, CurPos As Long, CurPosDec As Long, MaxEntries As Long

Private Sub InitValues()
  BytePos = 0
  CurPos = 4
  CurPosDec = 0
  MaxEntries = -1
  CurSize = RedimInterval
  ReDim DestArray(RedimInterval)
  Erase Dictory
End Sub

Function CompressData(Buffer() As Byte) As Boolean
  Dim MaxBound As Long, i As Long, ip As Long, Ret As Long, LastRet As Long, Size As Long, Count As Long
  Dim Char As Byte, NewChar As Byte, TempArray(MaxEntrySize) As Byte
  Dim MustFlush As Boolean, Blocked As Boolean
  
  On Error GoTo HuboErr
  InitValues
  MaxBound = UBound(Buffer)
  CopyMemory DestArray(0), MaxBound, Len(MaxBound)
  For i = 0 To MaxBound
    Count = Count + 1
    If Count = BlockSize Then
      Blocked = True
    ElseIf Count = BlockStop Then
      Count = 0
      InitDict
      Blocked = False
      AddByteToStream Signal, False
    End If
    Char = Buffer(i)
    If i = MaxBound Then
      AddByteToStream Char, True
    Else
      Size = 1
      TempArray(0) = Char
      ip = 1
      LastRet = -1
      Do
        If i + ip > MaxBound Then
          MustFlush = True
          Exit Do
        End If
        NewChar = Buffer(i + ip)
        AddToArray TempArray, Size, NewChar
        If Blocked Then Exit Do
        Ret = SearchInDictory(TempArray, Size)
        If Ret = -1 Then Exit Do
        LastRet = Ret
        ip = ip + 1
      Loop
      i = i + ip - 1
      AddToDictory TempArray, Size
      If LastRet = -1 Then
        AddByteToStream Char, False
      Else
        AddByteToStream LastRet + 256, False
      End If
    End If
  Next i
  If MustFlush Then AddByteToStream -1, True
  ReDim Preserve DestArray(CurPos - 1)
  Buffer = DestArray
  CompressData = True
HuboErr:
End Function

Function DecompressData(Buffer() As Byte) As Boolean
  Dim MaxBound As Long, DecompSize As Long, Char As Long, LastChar As Long, Size As Long
  Dim TempArray(MaxEntrySize) As Byte
  Dim EraseDict As Boolean

  On Error GoTo HuboErr
  InitValues
  MaxBound = UBound(Buffer)
  CopyMemory DecompSize, Buffer(0), Len(DecompSize)
  LastChar = GetNextByteFromStream(Buffer)
  While CurPosDec <= DecompSize
    If EraseDict Then
      InitDict
      EraseDict = False
    End If
    Do
      Char = GetNextByteFromStream(Buffer)
      If Char = Signal Then EraseDict = True Else Exit Do
    Loop
    If LastChar < 256 Then
      Size = 1
      TempArray(0) = LastChar
      AddByteToArray LastChar
    Else
      With Dictory(LastChar - 256)
        Size = .Size
        CopyMemory TempArray(0), .Entry(0), Size
      End With
      AddEntryToArray LastChar - 256
    End If
    If Char < 256 Then
      AddToArray TempArray, Size, Char
    ElseIf Char - 257 = MaxEntries Then
      If LastChar < 256 Then
        AddToArray TempArray, Size, LastChar
      Else
        AddToArray TempArray, Size, Dictory(LastChar - 256).Entry(0)
      End If
    ElseIf Char - 257 > MaxEntries Then
      Exit Function
    Else
      AddToArray TempArray, Size, Dictory(Char - 256).Entry(0)
    End If
    LastChar = Char
    AddToDictory TempArray, Size
  Wend
  ReDim Preserve DestArray(CurPosDec - 1)
  Buffer = DestArray
  DecompressData = True
HuboErr:
End Function

'decode from 10-bit value
Private Function GetNextByteFromStream(Stream() As Byte) As Long
  Select Case BytePos
    Case 0
      GetNextByteFromStream = CLng(Stream(CurPos)) * 4 + CLng(Stream(CurPos + 1) And 192) \ 64
    Case 1
      GetNextByteFromStream = CLng(Stream(CurPos) And 63) * 16 + CLng(Stream(CurPos + 1) And 240) \ 16
    Case 2
      GetNextByteFromStream = CLng(Stream(CurPos) And 15) * 64 + CLng(Stream(CurPos + 1) And 252) \ 4
    Case 3
      GetNextByteFromStream = CLng(Stream(CurPos) And 3) * 256 + CLng(Stream(CurPos + 1))
  End Select
  BytePos = BytePos + 1
  If BytePos = 4 Then
    BytePos = 0
    CurPos = CurPos + 1
  End If
  CurPos = CurPos + 1
End Function

'encode to 10-bit value
Private Sub AddByteToStream(ByVal Value As Long, ByVal Flush As Boolean)
  If Value <> -1 Then
    Select Case BytePos
      Case 0
        OutBuff(0) = CByte((Value And 1020) \ 4)
        OutBuff(1) = CByte((Value And 3) * 64)
      Case 1
        OutBuff(1) = OutBuff(1) Or CByte((Value And 1008) \ 16)
        OutBuff(2) = CByte((Value And 15) * 16)
      Case 2
        OutBuff(2) = OutBuff(2) Or CByte((Value And 960) \ 64)
        OutBuff(3) = CByte((Value And 63) * 4)
      Case 3
        OutBuff(3) = OutBuff(3) Or CByte((Value And 768) \ 256)
        OutBuff(4) = CByte(Value And 255)
    End Select
    BytePos = BytePos + 1
  End If
  If BytePos = 4 Or Flush Then
    If CurSize < CurPos + 5 Then
      CurSize = CurSize + RedimInterval
      ReDim Preserve DestArray(CurSize)
    End If
    CopyMemory DestArray(CurPos), OutBuff(0), 5
    CurPos = CurPos + 5
    BytePos = 0
  End If
End Sub

Private Sub AddToDictory(Entry() As Byte, ByVal Size As Long)
  If MaxEntries = MaxDict Then Exit Sub
  MaxEntries = MaxEntries + 1
  With Dictory(MaxEntries)
    .Size = Size
    ReDim .Entry(.Size - 1)
    CopyMemory .Entry(0), Entry(0), Size
  End With
End Sub

Private Function SearchInDictory(Entry() As Byte, ByVal Size As Long) As Long
  Dim i As Long

  For i = 0 To MaxEntries
    If CompareArray(Dictory(i).Entry, Dictory(i).Size, Entry, Size) Then
      SearchInDictory = i
      Exit Function
    End If
  Next i
  SearchInDictory = -1
End Function

Private Function CompareArray(Array1() As Byte, ByVal Size1 As Long, Array2() As Byte, ByVal Size2 As Long) As Boolean
  Dim i As Long

  If Size1 <> Size2 Then Exit Function
  For i = 0 To Size1 - 1
    If Array1(i) <> Array2(i) Then Exit Function
  Next i
  CompareArray = True
End Function

Private Sub AddToArray(Ary() As Byte, Size As Long, ByVal Value As Byte)
  Ary(Size) = Value
  Size = Size + 1
End Sub

Private Sub AddByteToArray(ByVal Value As Byte)
  If CurSize < CurPosDec Then
    CurSize = CurSize + RedimInterval
    ReDim Preserve DestArray(CurSize)
  End If
  DestArray(CurPosDec) = Value
  CurPosDec = CurPosDec + 1
End Sub

Private Sub AddEntryToArray(ByVal Index As Long)
  Dim EntrySize As Long

  EntrySize = Dictory(Index).Size
  If CurSize < CurPosDec + EntrySize Then
    CurSize = CurSize + RedimInterval
    ReDim Preserve DestArray(CurSize)
  End If
  CopyMemory DestArray(CurPosDec), Dictory(Index).Entry(0), EntrySize
  CurPosDec = CurPosDec + EntrySize
End Sub

Private Sub InitDict()
  If MaxEntries <> MaxDict Then Exit Sub
  MaxEntries = Limit
End Sub

Sub CompressFile(Path As String)
  Dim F1 As Integer
  Dim Ary() As Byte

  F1 = FreeFile()
  Open Path For Binary As F1
  If LOF(F1) Then
    ReDim Ary(LOF(F1) - 1)
    Get F1, , Ary
    Close F1
    CompressData Ary
    Kill Path
    Open Path For Binary As F1
    Put F1, , Ary
  End If
  Close F1
End Sub

Sub DecompressFile(Dst As String, Src As String)
  Dim F1 As Integer
  Dim Ary() As Byte

  F1 = FreeFile()
  Open Src For Binary As F1
  If LOF(F1) Then
    ReDim Ary(LOF(F1) - 1)
    Get F1, , Ary
    Close F1
    DecompressData Ary
  Else
    Close F1
  End If
  Open Dst For Binary As F1
  Put F1, , Ary
  Close F1
End Sub
