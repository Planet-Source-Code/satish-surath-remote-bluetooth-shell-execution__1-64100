VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Remote Bluetooth  Shell Execution"
   ClientHeight    =   8145
   ClientLeft      =   600
   ClientTop       =   840
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Tag             =   "`"
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   1560
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   3015
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   240
      Pattern         =   "*.txt"
      TabIndex        =   11
      Top             =   5280
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2040
      Picture         =   "remote bluetooth shell exe.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   1935
      TabIndex        =   5
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   7080
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disable"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enable"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   960
   End
   Begin VB.Label Label4 
      Caption         =   $"remote bluetooth shell exe.frx":08B6
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1680
      Left            =   4320
      TabIndex        =   16
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label lblheading 
      AutoSize        =   -1  'True
      Caption         =   "Shell Execution"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   525
      Index           =   1
      Left            =   4080
      TabIndex        =   14
      Top             =   0
      Width           =   3330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Text Files Present In The Above Folder Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   12
      Top             =   4920
      Width           =   4590
   End
   Begin VB.Label Label2 
      Caption         =   $"remote bluetooth shell exe.frx":093E
      Height          =   3015
      Left            =   4320
      TabIndex        =   10
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Instructions :"
      BeginProperty Font 
         Name            =   "WST_Engl"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4320
      TabIndex        =   9
      Top             =   1200
      Width           =   2100
   End
   Begin VB.Label Label11 
      Caption         =   "2006 Satish Surath"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label10 
      Caption         =   "©"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "This Programme Was Designed For Demonstration Purpose Only With No Warranty/Liability Either Implied Or Expressed."
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   7680
      Width           =   4335
   End
   Begin VB.Label lblrefresh 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4440
      TabIndex        =   3
      Top             =   7080
      Width           =   60
   End
   Begin VB.Label lblheading 
      Caption         =   "Remote"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pin As Variant
Dim t1 As Variant
Dim t2 As Variant
Dim t3 As Variant
Dim t4 As Variant
Dim t5 As Variant
Dim t6 As Variant
Dim t7 As Variant
Dim t8 As Variant
Dim t9 As Variant
Dim t10 As Variant
Dim binpin As Integer
Dim strFileName As String
Dim m As Variant
Dim flag As Integer
Dim filenum As Variant
Dim tim As Integer
Dim counter As Integer
Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Sub cmdexit_Click()

End
End Sub


Private Sub Command1_Click()
Timer1.Enabled = True
Timer3.Enabled = True
Command2.Enabled = True
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Timer3.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
End Sub


Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1.Refresh
End Sub



Private Sub Drive1_Change()
On Error GoTo Trap
Dir1.Path = Drive1.Drive
Trap:
Exit Sub
End Sub

Private Sub Form_Load()
strFileName = "c:\out.por"
counter = 1
Command2.Enabled = False
Command1.Enabled = True
tim = 1
Dir1.Path = "c:\"
File1.Path = "c:\"


 
End Sub


Private Sub Timer1_Timer()
Dim r As Long
On Error GoTo Trap

filenum = FreeFile

Open strFileName For Input As filenum

m = Input(LOF(filenum), filenum)

 
    r = ShellExecute(Me.hwnd, "open", m, 0&, fGetWinDir, 1)
        
        
        
        

Close filenum

Kill (strFileName)



Trap:
Exit Sub

End Sub


Private Sub Timer3_Timer()
File1.Refresh
If File1.ListCount > 0 Then
File1.Selected(0) = True
End If
strFileName = File1.Path + File1.Filename



If counter <= 10 Then
lblrefresh.Caption = ""
counter = counter + 1
End If
If counter = 4 Then
lblrefresh.Caption = "Refreshing system info.."
counter = 0
End If
End Sub
Private Function fGetWinDir() As String
' Wrapper to return OS Path
Dim lRet As Long, lSize As Long, sBuf As String * 512
    
    
    lSize = 512
    lRet = GetWindowsDirectory(sBuf, lSize)
    fGetWinDir = Left(sBuf, InStr(1, sBuf, Chr(0)) - 1)


    
End Function

