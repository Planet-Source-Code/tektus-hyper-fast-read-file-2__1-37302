VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVBBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FileName As String
Public Content As String

Public Sub ReadFile()
Dim FileNum As Byte
Dim Lenght As Long
Dim Counter As Long
Dim bContent() As Byte

    FileNum = FreeFile: Lenght = FileLen(FileName)
    ReDim bContent(Lenght): Content = Space$(Lenght)
    
    Open FileName For Binary Access Read As #FileNum
    
    For Counter = 1 To Lenght
        Get #FileNum, Counter, bContent(Counter)
        Mid$(Content, Counter, 1) = Chr$(bContent(Counter))
    Next
      
    Close #FileNum
    
End Sub

