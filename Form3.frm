VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   5265
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Play Game"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox reg 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393217
      MultiLine       =   0   'False
      TextRTF         =   $"Form3.frx":0000
   End
   Begin VB.Label Label1 
      Caption         =   "Registration Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If reg.Text = "91371994" Then
reg.SaveFile ("c:\wintoe.rtf")
MsgBox "Thank you for registering", vbInformation
End
Else
MsgBox "Incorrect code", vbCritical
End If
End Sub

Private Sub Command2_Click()
PlaySound "\please register"
Form3.Visible = False
Form1.Visible = True
End Sub

Private Sub Form_Load()
reg.LoadFile ("c:\wintoe.rtf")
If reg.Text = "91371994" Then
Form3.Visible = False
Form1.Visible = True
Form1.Command6.Enabled = True
Form1.sl(0).Enabled = True
End If
End Sub
