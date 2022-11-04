VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command20 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "x^1/2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command19 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "x^2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command18 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Caption         =   "<="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command17 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4280
      Width           =   735
   End
   Begin VB.CommandButton Command16 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4275
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4275
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2685
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2685
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2685
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   240
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2690
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      CausesValidation=   0   'False
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   240
      MaxLength       =   9
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public a, b, c, d, e

Private Sub Command1_Click()
d = 0
e = 0
a = Val(Text1.Text)
c = 1
Text1.Text = "0"
End Sub

Private Sub Command10_Click()
If Text1.Text = "0" Then Text1.Text = ""
If d <> 1 Then Text1.Text = Text1.Text + "6"
End Sub

Private Sub Command11_Click()
If Text1.Text = "0" Then Text1.Text = ""
If d <> 1 Then Text1.Text = Text1.Text + "7"
End Sub

Private Sub Command12_Click()
If Text1.Text = "0" Then Text1.Text = ""
If d <> 1 Then Text1.Text = Text1.Text + "8"
End Sub

Private Sub Command13_Click()
If Text1.Text = "0" Then Text1.Text = ""
If d <> 1 Then Text1.Text = Text1.Text + "9"
End Sub

Private Sub Command14_Click()
If Text1.Text <> "0" And d = 0 Then Text1.Text = Text1.Text + "0"
End Sub

Private Sub Command15_Click()
d = 0
e = 0
Text1.Text = "0"
End Sub

Private Sub Command16_Click()
If d <> 1 And e = 0 Then Text1.Text = Text1.Text + "."
e = 1
End Sub

Private Sub Command17_Click()
d = 1
b = Val(Text1.Text)
If c = 1 Then Text1.Text = a + b
If c = 2 Then Text1.Text = a - b
If c = 3 Then Text1.Text = a * b
If c = 4 And b <> 0 Then Text1.Text = a / b
If c = 4 And a = 0 Then Text1.Text = a
If c = 4 And b = 0 Then Text1.Text = "Nane"
End Sub

Private Sub Command18_Click()
y = Text1.Text
x = Len(y)
If d = 0 Then Text1.Text = Mid(y, 1, x - 1)
If Text1.Text = "" Then Text1.Text = "0"
End Sub

Private Sub Command19_Click()
Text1.Text = Val(Text1.Text) * Val(Text1.Text)
d = 1
End Sub

Private Sub Command2_Click()
d = 0
e = 0
a = Val(Text1.Text)
c = 2
Text1.Text = "0"
End Sub

Private Sub Command20_Click()
Text1.Text = Sqr(Val(Text1.Text))
d = 1
End Sub

Private Sub Command3_Click()
d = 0
e = 0
a = Val(Text1.Text)
c = 3
Text1.Text = "0"
End Sub

Private Sub Command4_Click()
d = 0
e = 0
a = Val(Text1.Text)
c = 4
Text1.Text = "0"
End Sub

Private Sub Command5_Click()
If Text1.Text = "0" Then Text1.Text = ""
If d <> 1 Then Text1.Text = Text1.Text + "1"
End Sub

Private Sub Command6_Click()
If Text1.Text = "0" Then Text1.Text = ""
If d <> 1 Then Text1.Text = Text1.Text + "2"
End Sub

Private Sub Command7_Click()
If Text1.Text = "0" Then Text1.Text = ""
If d <> 1 Then Text1.Text = Text1.Text + "3"
End Sub

Private Sub Command8_Click()
If Text1.Text = "0" Then Text1.Text = ""
If d <> 1 Then Text1.Text = Text1.Text + "4"
End Sub

Private Sub Command9_Click()
If Text1.Text = "0" Then Text1.Text = ""
If d <> 1 Then Text1.Text = Text1.Text + "5"
End Sub
