VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Reading / Writing Methodes"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   5895
      Begin VB.TextBox txtFileContent 
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   12
         Top             =   840
         Width           =   5655
      End
      Begin VB.Label Label4 
         Caption         =   $"Form1.frx":0442
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin MSComDlg.CommonDialog cDialog 
         Left            =   5400
         Top             =   3000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton btnMyMethode 
         Caption         =   "NORMAL: Only VB Code (my own methode)"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   5415
      End
      Begin VB.CommandButton btnApi 
         Caption         =   "FAST: Read Using Windows Api Read File Functions"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   5415
      End
      Begin VB.CommandButton btnVBBinary 
         Caption         =   "SLOW: Read Using VB Binary Read Methode"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   5415
      End
      Begin VB.CommandButton btnBrowse 
         Caption         =   "Browse"
         Height          =   300
         Left            =   4800
         TabIndex        =   3
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   1350
         TabIndex        =   2
         Top             =   315
         Width           =   3375
      End
      Begin VB.Label txtResult 
         BackStyle       =   0  'Transparent
         Caption         =   "0 milliseconds.."
         Height          =   495
         Left            =   1440
         TabIndex        =   9
         Top             =   2160
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "Test Result:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PLEASE DON'T FORGET TO VOTE FOR ME"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   7
         Top             =   2520
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Target File:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount& Lib "kernel32" ()

Dim FileIO As clsFileIO
Dim FileApi As clsApi
Dim FileVB As clsVBBin

Dim TimeEl As Long ' For timer calculations

Private Sub btnApi_Click()
If CheckFile = False Then Exit Sub

Set FileApi = New clsApi
    
    FileApi.FileName = txtFileName
    
  ' Begin Reading and Start Timer
    TimeEl = GetTickCount&
    FileApi.ApiReadFile
    txtResult = Len(FileApi.Content) & " bytes read in " & GetTickCount& - TimeEl & "ms using Api methode"
  ' End test
  
    txtFileContent = FileApi.Content

Set FileApi = Nothing
End Sub

Private Sub btnMyMethode_Click()
If CheckFile = False Then Exit Sub

Set FileIO = New clsFileIO
    
    FileIO.FileName = txtFileName
    
    TimeEl = GetTickCount&
    FileIO.ReadFile
    txtFileContent = FileIO.Content
    txtResult = Len(FileIO.Content) & " bytes read in " & GetTickCount& - TimeEl & "ms using my methode"
    
Set FileIO = Nothing
End Sub

Private Sub btnVBBinary_Click()
If CheckFile = False Then Exit Sub

Set FileVB = New clsVBBin
        
    FileVB.FileName = txtFileName
    
    TimeEl = GetTickCount&
    FileVB.ReadFile
    txtResult = Len(FileVB.Content) & " bytes read in " & GetTickCount& - TimeEl & "ms using VBBinary Open methode"
    txtFileContent = FileVB.Content

Set FileVB = Nothing
End Sub

Private Function CheckFile() As Boolean
Dim Result As Long
    
    If txtFileName = "" Or Dir$(txtFileName) = "" Then
       Result = MsgBox("Please select a file to read..", vbExclamation, "File Not Defined or Not Found")
       CheckFile = False
    Else
       CheckFile = True
    End If

txtResult = "Processing... Please wait...": DoEvents
End Function

Private Sub btnBrowse_Click()
cDialog.ShowOpen: txtFileName = cDialog.FileName
End Sub

Private Sub Form_Load()
TimeEl = MsgBox("All of these methodes are created by my own and belongs to me." & vbCrLf & "Please use them giving credit to me and don't forget to vote for me !" & vbCrLf & vbCrLf & "Ozan Yasin Dogan, Istanbul / Turkey", vbInformation, "File Tests")
TimeEl = MsgBox("THIS METHODES GETS THE FILE IN MEMORY ! " & vbCrLf & "IF YOU CHOOSE A BIGGER FILE THEN YOUR FREEMEMORY," & vbCrLf & "WINDOWS WILL OPEN TEMPORARY MEMORY IN THE HARD DISK" & vbCrLf & "AND THIS MAY SLOW DOWN THE PROCESS A LOT !", vbCritical, "WARNING!")
End Sub
