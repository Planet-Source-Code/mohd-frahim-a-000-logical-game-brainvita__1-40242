VERSION 5.00
Begin VB.Form high 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brainvita - Highscores"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox highname 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      MaxLength       =   40
      TabIndex        =   0
      ToolTipText     =   "Enter your name here"
      Top             =   960
      Width           =   3855
   End
   Begin VB.ListBox highlist 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Current highscores"
      Top             =   2280
      Width           =   3855
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      ToolTipText     =   "Click to go back"
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   1
      ToolTipText     =   "Click to save your high score"
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   10
      ToolTipText     =   "Date scored on (MM/DD/YY)"
      Top             =   1920
      Width           =   345
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   9
      ToolTipText     =   "Name of scorer"
      Top             =   1920
      Width           =   405
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   8
      ToolTipText     =   "Elapsed time"
      Top             =   1920
      Width           =   210
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Sc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Score"
      Top             =   1920
      Width           =   165
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Enter your name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Your score is among the 10 highest."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Highscores:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1140
   End
End
Attribute VB_Name = "high"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
highname.Text = ""
Unload Me
End Sub

Private Sub cmdok_Click()

If highname.Text = "" Then
    MsgBox "You have not entered the name.", vbOKOnly + vbExclamation, "Brainvita"
    Exit Sub
End If

Open "BVHigh.brh" For Append As #1
Print #1, score
Print #1, elapsedtime
Print #1, highname.Text
Print #1, Date

Close #1

Unload Me
End Sub

Private Sub Form_Load()

'load the highscore form
If highshow = True Then Exit Sub

If firsthigh = False Then
    Open "BVHigh.brh" For Input As #1
    Line Input #1, temp
    Line Input #1, temp
    highitem = temp + Space(4 - Len(temp))
    Line Input #1, temp
    highitem = highitem + temp + Space(7 - Len(temp))
    Line Input #1, temp
    highitem = highitem + temp + Space(42 - Len(temp))
    Line Input #1, temp
    highitem = highitem + temp
    highlist.AddItem highitem

    While Not EOF(1)
        Line Input #1, temp
        highitem = temp + Space(4 - Len(temp))
        Line Input #1, temp
        highitem = highitem + temp + Space(7 - Len(temp))
        Line Input #1, temp
        highitem = highitem + temp + Space(42 - Len(temp))
        Line Input #1, temp
        highitem = highitem + temp
        highlist.AddItem highitem
    Wend
    Close #1
End If
End Sub


