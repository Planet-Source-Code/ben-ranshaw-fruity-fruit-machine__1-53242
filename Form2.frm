VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save/Load"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3990
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Save Score"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Save Score"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Save Score"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Load"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Load"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Load"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   4200
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Hide"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   4200
      Width           =   735
   End
   Begin Project1.chameleonButton command7 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   4200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&Load Scores"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form2.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"Form2.frx":001C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Index           =   10
      Left            =   2640
      TabIndex        =   9
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393217
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form2.frx":009B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"Form2.frx":0112
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Index           =   4
      Left            =   2640
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393217
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form2.frx":0191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   12
      Top             =   3000
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"Form2.frx":0208
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Index           =   6
      Left            =   2640
      TabIndex        =   13
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393217
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form2.frx":0287
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Save Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Money:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line3 
      X1              =   3960
      X2              =   0
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   4080
      X2              =   -720
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   3960
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   3960
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
RichTextBox1(10).Text = Form1.Label4.Caption
RichTextBox1(0).SaveFile ("c:\high12.rtf")
RichTextBox1(10).SaveFile ("c:\high22.rtf")
Form4.Visible = False
MsgBox "Saved!", vbInformation
End Sub

Private Sub Command2_Click()
On Error Resume Next
RichTextBox1(4).Text = Form1.Label4.Caption
RichTextBox1(3).SaveFile ("c:\high32.rtf")
RichTextBox1(4).SaveFile ("c:\high42.rtf")
MsgBox "Saved!", vbInformation
Form4.Visible = False
End Sub



Private Sub Command3_Click()
On Error Resume Next
Form1.ProgressBar1.Value = RichTextBox1(10).Text
Form4.Visible = False
MsgBox "Load Sucsussful!", vbInformation
End Sub

Private Sub Command4_Click()
On Error Resume Next
RichTextBox1(6).Text = Form1.Label4.Caption
RichTextBox1(5).SaveFile ("c:\high72.rtf")
RichTextBox1(6).SaveFile ("c:\high82.rtf")
MsgBox "Saved!", vbInformation
Form4.Visible = False
End Sub


Private Sub Command7_Click()
On Error Resume Next
RichTextBox1(0).LoadFile ("c:\high12.rtf")
RichTextBox1(10).LoadFile ("c:\high22.rtf")
RichTextBox1(3).LoadFile ("c:\high32.rtf")
RichTextBox1(4).LoadFile ("c:\high42.rtf")
RichTextBox1(5).LoadFile ("c:\high72.rtf")
RichTextBox1(6).LoadFile ("c:\high82.rtf")
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
End Sub

Private Sub Command8_Click()
Form2.Visible = False
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If RichTextBox1(0).Text = "" Or RichTextBox1(10).Text = "" Then
Command3.Enabled = False
Else
Command3.Enabled = True
If RichTextBox1(3).Text = "" Or RichTextBox1(4).Text = "" Then
Command5.Enabled = False
Else
Command5.Enabled = True
If RichTextBox1(5).Text = "" Or RichTextBox1(6).Text = "" Then
Command6.Enabled = False
Else
Command6.Enabled = True
End If
End If
End If
End Sub

