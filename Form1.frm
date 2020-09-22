VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FRUITASTIC FRUIT MACHINE"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7575
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0442
   ScaleHeight     =   5670
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   8040
      Picture         =   "Form1.frx":90460
      ScaleHeight     =   1575
      ScaleWidth      =   2055
      TabIndex        =   22
      Top             =   6000
      Width           =   2055
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   5520
      Picture         =   "Form1.frx":91A08
      ScaleHeight     =   1575
      ScaleWidth      =   2055
      TabIndex        =   21
      Top             =   6000
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   3000
      Picture         =   "Form1.frx":92BEC
      ScaleHeight     =   1575
      ScaleWidth      =   2055
      TabIndex        =   20
      Top             =   6000
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   480
      Picture         =   "Form1.frx":95E52
      ScaleHeight     =   1575
      ScaleWidth      =   2055
      TabIndex        =   19
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3360
      Top             =   0
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Unregister"
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   4800
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   135
      Left            =   240
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   1e38
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Collect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7920
      TabIndex        =   15
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Gamble"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7920
      TabIndex        =   14
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      MaxLength       =   1
      TabIndex        =   12
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1920
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2400
      Top             =   0
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Claim"
      Default         =   -1  'True
      Height          =   735
      Left            =   6240
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   1e38
   End
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   6240
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NUDGE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NUDGE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NUDGE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE VOTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1440
      TabIndex        =   27
      Top             =   5040
      Width           =   3615
   End
   Begin VB.Label Label12 
      Caption         =   "78"
      Height          =   375
      Left            =   8040
      TabIndex        =   26
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "10"
      Height          =   375
      Left            =   5520
      TabIndex        =   25
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "142"
      Height          =   375
      Left            =   3000
      TabIndex        =   24
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "43"
      Height          =   375
      Left            =   480
      TabIndex        =   23
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7920
      TabIndex        =   16
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   7920
      TabIndex        =   13
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   7680
      X2              =   7680
      Y1              =   0
      Y2              =   5760
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   6600
      TabIndex        =   10
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FRUITASTIC FRUIT MACHINE  MADE BY BEN RANSHAW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   7
      Top             =   2880
      Width           =   5055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "300"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   480
      Top             =   2040
      Width           =   5055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   4080
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu sl 
         Caption         =   "Save+Load"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^S
      End
      Begin VB.Menu reg 
         Caption         =   "Register"
         Index           =   0
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer2.Enabled = True
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Timer5.Enabled = True
End Sub

Private Sub Command2_Click()
Timer4.Enabled = True
Label3.Caption = Int(Rnd * 8)
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Timer5.Enabled = True
End Sub

Private Sub Command3_Click()
Timer3.Enabled = True
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Timer5.Enabled = True
End Sub

Private Sub Command4_Click()
MsgBox "come back soon", vbInformation
PlaySound "\3sevens1"
End
End Sub

Private Sub Command5_Click()
If Label1.Caption = Label2.Caption Then
If Label1.Caption = Label2.Caption = Label3.Caption Then
Form1.Height = 8670
Form1.Width = 10100
Else
If Label1.Caption = Label2.Caption Then
Form1.Width = 10100
End If
Command5.Enabled = False
End If
End If
End Sub

Private Sub Command6_Click()
Label7.Caption = Int(Rnd * 4)
If Text1.Text = Label7.Caption Then
ProgressBar2.Value = ProgressBar2.Value + ProgressBar2.Value
Else
ProgressBar2.Value = 0
Form1.Width = 7777
End If
PlaySound "\spin"
End Sub

Private Sub Command7_Click()
ProgressBar1.Value = ProgressBar1.Value + Label8.Caption
ProgressBar2.Value = 0
Form1.Width = 7777
PlaySound "\coll"
End Sub

Private Sub Command8_Click()
Form3.reg.Text = ""
Form3.reg.SaveFile ("c:\wintoe.rtf")
End Sub

Private Sub Form_Load()
Randomize
ProgressBar1.Value = 300
End Sub



Private Sub Label6_Click()
PlaySound "\spin"
If Timer3.Enabled = False Then
Command5.Enabled = True
Timer6.Enabled = True
ProgressBar1.Value = ProgressBar1.Value - 25
Label4.Caption = ProgressBar1.Value
Label1.Caption = Int(Rnd * 7)
Label2.Caption = Int(Rnd * 7)
Label3.Caption = Int(Rnd * 8)
Label10.Caption = Int(Rnd * 150)
Label11.Caption = Int(Rnd * 150)
Label12.Caption = Int(Rnd * 150)
Label9.Caption = Int(Rnd * 150)
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Command5.Visible = True
ProgressBar2.Value = 0
Timer4.Enabled = True
Timer5.Enabled = True
If ProgressBar1.Value = 50 Then
ProgressBar1.Value = ProgressBar1.Value + 300
End If
End If
End Sub





Private Sub Picture1_Click()
ProgressBar1.Value = ProgressBar1.Value + Label9.Caption
Form1.Height = 6570
MsgBox "You have won " + Label9.Caption, vbInformation
End Sub

Private Sub Picture2_Click()
ProgressBar1.Value = ProgressBar1.Value + Label10.Caption
MsgBox "You have won " + Label10.Caption, vbInformation
Form1.Height = 6570
End Sub

Private Sub Picture3_Click()
ProgressBar1.Value = ProgressBar1.Value + Label11.Caption
MsgBox "You have won " + Label11.Caption, vbInformation
Form1.Height = 6570
End Sub

Private Sub Picture5_Click()
ProgressBar1.Value = ProgressBar1.Value + Label12.Caption
MsgBox "You have won " + Label12.Caption, vbInformation
Form1.Height = 6570
End Sub

Private Sub reg_Click(Index As Integer)
Form3.Visible = True
Form3.Command2.Enabled = False
End Sub

Private Sub sl_Click(Index As Integer)
Form2.Visible = True
End Sub

Private Sub Timer1_Timer()
Label4.Caption = ProgressBar1.Value
Label8.Caption = ProgressBar2.Value
End Sub

Private Sub Timer2_Timer()
Label1.Caption = Int(Rnd * 8)
End Sub

Private Sub Timer3_Timer()
Label2.Caption = Int(Rnd * 8)
End Sub

Private Sub Timer4_Timer()
Label3.Caption = Int(Rnd * 8)
End Sub

Private Sub Timer5_Timer()
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
End Sub


