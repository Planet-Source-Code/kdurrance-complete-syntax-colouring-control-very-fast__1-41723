VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin Project1.CodeEdit rtftext 
      Height          =   6495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   11456
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      ceKeyWords      =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load File"
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   6600
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_VSCROLL = &H115
Private Sub Command1_Click()

Dim sfile As String
    
    sfile = InputBox("FileName and Path?", "Enter a FileName to Load")
    If sfile <> "" Then
        If Dir(sfile) <> "" Then
            rtftext.LoadFile (sfile)
        Else
            MsgBox "File Not Found!", vbCritical
        End If
    End If

End Sub

Private Sub Form_Load()
    rtftext.LoadFile App.Path & "\Source.txt"
End Sub
