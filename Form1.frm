VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Love Calculator        by Didar"
   ClientHeight    =   3765
   ClientLeft      =   2235
   ClientTop       =   1785
   ClientWidth     =   5580
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   3765
   ScaleWidth      =   5580
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   480
      Top             =   600
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2880
      Picture         =   "Form1.frx":61FB
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   1320
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      Picture         =   "Form1.frx":6AC5
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   1320
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4800
      Picture         =   "Form1.frx":738F
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   1890
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1890
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Congratulation!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Possibility"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   3360
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   3000
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, b, c As Integer
On Error Resume Next
Label3.Caption = "No Comment.."
Timer1.Enabled = False
Label8.Visible = False
Label8.Left = 3840
Picture1.Visible = False
Text1.Text = UCase(Text1.Text)
Text2.Text = UCase(Text2.Text)
b = Asc(Text1.Text)
c = Asc(Text2.Text)
a = (b + c) / 2

Label1.Caption = (b + c) / 2
If b = c Then
Label1.Caption = "0"
Label3.Caption = "God Bless You"
Else
If a = 66 Then
Label1.Caption = "10"
Label3.Caption = "Do You Really Fall In Love?"
Else
If a = 70 Then
Label1.Caption = "20"
Label3.Caption = "Very Sad.Bad Luck..."
Else
If a = 75 Then
Label1.Caption = "30"
Label3.Caption = "Are You Sure,You Fall In Love?"
Else
If a = 80 Then
Label1.Caption = "99.99"
Label3.Caption = "Wow!! You Are A Lucky Man."
Picture1.Visible = True
Timer1.Enabled = True
Else
If a = 69 Then
Label1.Caption = "42"
Label3.Caption = "You Are Joking,You don't Love ..."
Else
If a = 72 Then
Label1.Caption = "55"
Label3.Caption = "God Bless You!! May Be Success.."
Else
If a = 88 Then
Label1.Caption = "35"
Label3.Caption = "Love Is Not So Easy Man..."
Else
If a = 85 Then
Label1.Caption = "99.99"
Label3.Caption = "Wow!! You Are A Lucky Man."
Picture1.Visible = True
Timer1.Enabled = True
Else
If a = 87 Then
Label1.Caption = "91"
Label3.Caption = "Hey! You Will Be A Happy Man."
Picture1.Visible = True
Timer1.Enabled = True
Else
If a = 74 Then
Label1.Caption = "5"
Label3.Caption = "So Sorry!! You Are Not A Lover."
Else
If a = 79 Then
Label1.Caption = "98"
Label3.Caption = "Wow!! You Are A Lucky Man."
Picture1.Visible = True
Timer1.Enabled = True
Else
If a = 73 Then
Label1.Caption = "97"
Label3.Caption = "Wow!! You Are A Lucky Man.."
Picture1.Visible = True
Timer1.Enabled = True
Else
If a = 81 Then
Label1.Caption = "99"
Label3.Caption = "Hi!! You Are A Real Lover.."
Picture1.Visible = True
Timer1.Enabled = True
Else
If a = 90 Then
Label1.Caption = "94"
Label3.Caption = "Wow!! You Are A Lucky Man.."
Picture1.Visible = True
Timer1.Enabled = True
Else
If a = 91 Then
Label1.Caption = "96"
Label3.Caption = "Hey!! You Are A Real Lover.."
Picture1.Visible = True
Timer1.Enabled = True
Else
If a = 68 Then
Label1.Caption = "92"
Label3.Caption = "Congratulation!! You Are A Real Lover.."
Picture1.Visible = True
Timer1.Enabled = True
Else
If a = 76 Then
Label1.Caption = "60"
Label3.Caption = "Try More.You Can Be A Real Lover.."
Else
If a = 67 Then
Label1.Caption = "85"
Label3.Caption = "OH! Good.You Must Success.."
Else
If a = 71 Then
Label1.Caption = "3"
Label3.Caption = "Bad Luck.God Bless You!!."
Else
If a = 77 Then
Label1.Caption = "49"
Label3.Caption = "Try More.You Can Be A Real Lover.."
Else
If a = 78 Then
Label1.Caption = "90"
Label3.Caption = "Hey Man You Are Really Lucky.."
Picture1.Visible = True
Timer1.Enabled = True
Else
If a = 82 Then
Label1.Caption = "60"
Label3.Caption = "Try More.You Can Be A Real Lover.."






End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Timer1_Timer()
Label8.Visible = True
Label8.Left = Label8.Left - 10
End Sub
