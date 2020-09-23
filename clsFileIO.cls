VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const Buf As Integer = 32767

Dim Temp As String * Buf
Dim FileLen As Long
Dim Multiply As Long
Dim Plus As Long
Dim Point As Long
Dim FileNo As Byte
Dim Counter As Long

Public FileName As String
Public Content As String

Public Sub ReadFile()
FileNo = FreeFile

Open FileName For Random As #FileNo Len = Buf

FileLen = LOF(FileNo)
Multiply = FileLen \ Buf
Content = Space$(FileLen)
Plus = FileLen - (Multiply * Buf)
Point = 1
  
  If Multiply = 0 Then
    Plus = FileLen: Counter = 1: GoTo Jump1
  End If
  
  For Counter = 1 To Multiply
    Get #FileNo, Counter, Temp
      Mid$(Content, Point, Buf) = Temp
      Point = Point + Buf
  Next
  
Jump1:
  
  If Plus > 0 Then
    Get #FileNo, Counter, Temp
      Mid$(Content, Point, Plus) = Left$(Temp, Plus)
  End If

Close #FileNo
End Sub
