VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caclulator"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      BackColor       =   &H00008000&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00008000&
      Caption         =   "<--"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00008000&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00008000&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00008000&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00008000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00008000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00008000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00008000&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00008000&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00008000&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00008000&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00008000&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00008000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00008000&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00008000&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00008000&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008000&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
      Begin VB.Menu About_auther 
         Caption         =   "About_auther"
      End
   End
   Begin VB.Menu Color 
      Caption         =   "Color"
      Begin VB.Menu Defult 
         Caption         =   "Default"
      End
      Begin VB.Menu Blue 
         Caption         =   "Blue"
      End
      Begin VB.Menu White 
         Caption         =   "White"
      End
      Begin VB.Menu Yellow 
         Caption         =   "Yellow"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op, fn As Integer


Private Sub About_auther_Click()
Form3.Show


End Sub

Private Sub Black_Click()
Form1.BackColor = vbBlack

End Sub

Private Sub Blue_Click()
Form1.BackColor = vbBlue
End Sub

Private Sub Command1_Click()
Text1.Text = Text1.Text & 7


End Sub

Private Sub Command10_Click()
Text1.Text = Text1.Text & 4
End Sub

Private Sub Command11_Click()
op = 4
fn = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command13_Click()
Text1.Text = Text1.Text & 9
End Sub

Private Sub Command14_Click()
Text1.Text = Text1.Text & 1
End Sub

Private Sub Command15_Click()
If op = 1 Then
Text1.Text = Val(fn) + Val(Text1.Text)
ElseIf op = 2 Then
Text1.Text = Val(fn) - Val(Text1.Text)
ElseIf op = 3 Then
Text1.Text = Val(fn) * Val(Text1.Text)
ElseIf op = 4 Then
Text1.Text = Val(fn) / Val(Text1.Text)
End If

End Sub

Private Sub Command16_Click()
Text1.Text = Text1.Text & 0
End Sub

Private Sub Command17_Click()
Text1.Text = Text1.Text & (".")

End Sub

Private Sub Command18_Click()
op = 1
fn = Text1.Text
Text1.Text = ""

End Sub

Private Sub Command19_Click()
Text1.Text = ""

End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text & 8
End Sub

Private Sub Command3_Click()
Text1.Text = Text1.Text & 3
End Sub

Private Sub Command4_Click()
Text1.Text = Text1.Text & 2
End Sub

Private Sub Command5_Click()
 Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
End Sub

Private Sub Command6_Click()
op = 2
fn = Text1.Text
Text1.Text = ""

End Sub

Private Sub Command7_Click()
op = 3
fn = Text1.Text
Text1.Text = ""

End Sub

Private Sub Command8_Click()
Text1.Text = Text1.Text & 6
End Sub

Private Sub Command9_Click()
Text1.Text = Text1.Text & 5
End Sub

Private Sub Defult_Click()
Form1.BackColor = vbBlack


End Sub

Private Sub Exit_Click()
Unload Me

End Sub

Private Sub Form_Load()
Text1.FontSize = "20"
Text1.FontBold = True
End Sub

Private Sub Text1_Change()
Text1.SetFocus
Text1.Locked = True

End Sub

Private Sub White_Click()
Form1.BackColor = vbWhite


End Sub

Private Sub Yello_Click()
Form1.BackColor = vbYellow
End Sub

Private Sub Yellow_Click()
Form1.BackColor = vbYellow
End Sub
