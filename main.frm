VERSION 5.00
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brainvita"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   DrawMode        =   6  'Mask Pen Not
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdscore 
      Caption         =   "Highscores"
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
      Left            =   5280
      TabIndex        =   6
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
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
      Left            =   4800
      TabIndex        =   1
      ToolTipText     =   "Back to previous move"
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdpause 
      Caption         =   "Pause"
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
      Left            =   5880
      TabIndex        =   0
      ToolTipText     =   "Pause/Resume Game"
      Top             =   2760
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox mainpict 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   5880
      Picture         =   "main.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   3960
      Width           =   510
   End
   Begin VB.Frame Frame1 
      Caption         =   "Legends"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4800
      TabIndex        =   62
      ToolTipText     =   "Legends"
      Top             =   3240
      Width           =   1935
      Begin VB.PictureBox hintpict 
         Height          =   510
         Left            =   240
         Picture         =   "main.frx":0614
         ScaleHeight     =   496.552
         ScaleMode       =   0  'User
         ScaleWidth      =   389.189
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hint"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   66
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label3 
         Caption         =   "Selected"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   65
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdscoring 
      Caption         =   "Scoring"
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
      Left            =   5880
      TabIndex        =   4
      ToolTipText     =   "Scoring pattern"
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdabout 
      Caption         =   "About"
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
      Left            =   4800
      TabIndex        =   5
      ToolTipText     =   "About Designer"
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit "
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
      Left            =   5880
      TabIndex        =   3
      ToolTipText     =   "Exit Brainvita"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
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
      Left            =   4800
      TabIndex        =   2
      ToolTipText     =   "New game"
      Top             =   720
      Width           =   855
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   48
      Left            =   3960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   47
      Left            =   3360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   46
      Left            =   2760
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   4440
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   45
      Left            =   2160
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   4440
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   44
      Left            =   1560
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   4440
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   43
      Left            =   960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   42
      Left            =   360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   41
      Left            =   3960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   40
      Left            =   3360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   39
      Left            =   2760
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   3840
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   38
      Left            =   2160
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   3840
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   37
      Left            =   1560
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   3840
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   36
      Left            =   960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   35
      Left            =   360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   34
      Left            =   3960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3240
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   33
      Left            =   3360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3240
      Width           =   533
   End
   Begin VB.PictureBox emptypict 
      Height          =   535
      Left            =   4200
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   535
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   32
      Left            =   2760
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3240
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   31
      Left            =   2160
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3240
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   30
      Left            =   1560
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3240
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   29
      Left            =   960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3240
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   28
      Left            =   360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3240
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   27
      Left            =   3960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2640
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   26
      Left            =   3360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2640
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   25
      Left            =   2760
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2640
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   24
      Left            =   2160
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2640
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   23
      Left            =   1560
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2640
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   22
      Left            =   960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2640
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   21
      Left            =   360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2640
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   20
      Left            =   3960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2040
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   19
      Left            =   3360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2040
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   18
      Left            =   2760
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2040
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   17
      Left            =   2160
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2040
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   16
      Left            =   1560
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2040
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   15
      Left            =   960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2040
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   14
      Left            =   360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2040
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   13
      Left            =   3960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   12
      Left            =   3360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   11
      Left            =   2760
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1440
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   10
      Left            =   2160
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1440
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   9
      Left            =   1560
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1440
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   8
      Left            =   960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   7
      Left            =   360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   6
      Left            =   3960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   5
      Left            =   3360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   4
      Left            =   2760
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   840
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   3
      Left            =   2160
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   840
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   2
      Left            =   1560
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   840
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   1
      Left            =   960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   533
   End
   Begin VB.PictureBox board 
      Height          =   535
      Index           =   0
      Left            =   360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   533
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   240
      X2              =   4640
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   4640
      X2              =   4640
      Y1              =   720
      Y2              =   5120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   240
      X2              =   4640
      Y1              =   5120
      Y2              =   5120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   240
      X2              =   240
      Y1              =   720
      Y2              =   5120
   End
   Begin VB.Label lbltime 
      AutoSize        =   -1  'True
      Caption         =   "0 seconds"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Left            =   5400
      TabIndex        =   68
      ToolTipText     =   "Time elapsed"
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Time Elapsed:"
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
      Left            =   5160
      TabIndex        =   67
      Top             =   2160
      Width           =   1305
   End
   Begin VB.Label lblscore 
      AutoSize        =   -1  'True
      Caption         =   "32 remaining out of 32"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Left            =   4920
      TabIndex        =   61
      ToolTipText     =   "Your current score"
      Top             =   1920
      Width           =   1950
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Score: "
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
      Left            =   5520
      TabIndex        =   60
      Top             =   1680
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Action:"
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
      Left            =   360
      TabIndex        =   59
      Top             =   5280
      Width           =   705
   End
   Begin VB.Label lblstatus 
      AutoSize        =   -1  'True
      Caption         =   "Please select a box."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   58
      ToolTipText     =   "User action to be taken"
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "BRAINVITA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   57
      Top             =   0
      Width           =   2355
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub board_Click(index As Integer)
If cmdpause.Caption = "Resume" Then
    MsgBox "Game has been paused." & vbCrLf & _
    "To resume click the 'Resume' button.", vbOKOnly + vbExclamation, "Brainvita"
    Exit Sub
End If

For i = 0 To 48
    pictrowcol (i)  'calculate row and column of the box
                    ' make appearance of the box 3D if there is ball on that box
    If boxfilled(row, col) = True Then
        board(i).Appearance = 1
        board(i).Picture = mainpict.Picture
        board(i).BackColor = &H8000000F
    Else
        board(i).Appearance = 1
        board(i).Picture = emptypict.Picture
        board(i).BackColor = &H8000000F
    End If
Next i

pictrowcol (index)  ' find row and column of the box user selected

If boxfilled(row, col) = True Then  ' if box is not empty then show it flat
    board(index).Appearance = 0
    board(index).BackColor = &HFF00&
Else                                 ' else show error message
    If selected = False Then
        MsgBox "You can not select this box.", vbOKOnly + vbExclamation, "Brainvita"
        Exit Sub
    End If
    
End If

' If user selects a box and then again click on that box then unselect that box
If previous = index And selected = True Then
    board(index).Appearance = 1
    board(index).BackColor = &H8000000F
    selected = False
    If backcount > 0 Then cmdback.Enabled = True
    lblstatus.Caption = "Please select a box."
    previous = index
    Exit Sub
End If

' if selection made (i.e. value of selected was false before user click ) then show possible moves
' if box was selected previously then check that move is valid or not
'                        if move is valid then update the board and score

If selected = False Then
    selected = True
    cmdback.Enabled = False
    lblstatus.Caption = "Now select a valid move. Hint moves are shown here." & vbCrLf & _
                        "To unselect click on the selected box."
    Call showmoves(row, col)
    If movespossible = False Then    'show possible moves to the user
        MsgBox "No moves are possible from this box.", vbOKOnly + vbExclamation, "Brainvita"
        board(index).Appearance = 1
        board(index).BackColor = &H8000000F
        selected = False
        lblstatus.Caption = "Please select a box."
        cmdback.Enabled = True
    Else
    'if valid moves are there then show moves as hint
    For i = 1 To 4
        If movevalid(i) = True Then
            board(rowcolpict(moverow(i), movecol(i))).Picture = hintpict.Picture
        End If
    Next i
    End If
Else
    selected = False
    cmdback.Enabled = True
    lblstatus.Caption = "Please select a box."
    If checkmove(row, col) = False Then       '  check that the move is valid or not
    MsgBox "Invalid move !", vbOKOnly + vbCritical, "Brainvita"
    board_Click (previous)
    Exit Sub
    End If
    current = index
    updboard (previous)         '  if valid move then update the board and score
End If

previous = index
End Sub

Private Sub cmdabout_Click()
MsgBox "Brianvita - Version 1.0" & vbCrLf & _
       "Made by Mohd. Frahim, mdfrahim@yahoo.com" & vbCrLf & _
       "Tata Consultancy Services, New Delhi, India", vbOKOnly + vbExclamation, "Brainvita"
End Sub

Private Sub cmdback_Click()
Dim temp1 As Integer
Dim temp2 As Integer
backcount = backcount - 1
If backcount < 0 Then
    cmdback.Enabled = False
    Exit Sub
End If

If gameend = True Then
    Timer1.Enabled = True
    paused = False
    cmdpause.Enabled = True
End If

board(rowcolpict(fromboxrow(backcount + 1), fromboxcol(backcount + 1))).Picture = mainpict.Picture
boxfilled(fromboxrow(backcount + 1), fromboxcol(backcount + 1)) = True

board(rowcolpict(toboxrow(backcount + 1), toboxcol(backcount + 1))).Picture = emptypict.Picture
boxfilled(toboxrow(backcount + 1), toboxcol(backcount + 1)) = False

temp1 = (fromboxrow(backcount + 1) + toboxrow(backcount + 1)) / 2
temp2 = (fromboxcol(backcount + 1) + toboxcol(backcount + 1)) / 2

board(rowcolpict(temp1, temp2)).Picture = mainpict.Picture
boxfilled(temp1, temp2) = True
score = score + 1
lblscore.Caption = score & " remaining out of 32"
If backcount <= 0 Then cmdback.Enabled = False

End Sub

Private Sub cmdnew_Click()
If MsgBox("This will end this game." & vbCrLf & _
    "Are you sure to start new game ?", vbYesNo + vbQuestion, "Brainvita") = vbYes Then Form_Load
End Sub

Private Sub cmdpause_Click()
If paused = True Then
    paused = False
    Timer1.Enabled = True
    cmdpause.Caption = "Pause"
    If backcount > 0 Then cmdback.Enabled = True
    Me.Caption = "Brainvita - " & elapsedtime
Else
    paused = True
    Timer1.Enabled = False
    cmdpause.Caption = "Resume"
    cmdback.Enabled = False
    Me.Caption = "Brainvita - Paused"
End If

End Sub

Private Sub cmdquit_Click()
If MsgBox("Are you sure to quit this game ?", vbYesNo + vbQuestion, "Brainvita") = vbYes Then

MsgBox "Thanks for playing." & vbCrLf & _
       "Did you like the game ?" & vbCrLf & _
       "Tell MOHD FRAHIM at mdfrahim@yahoo.com", vbOKOnly + vbExclamation, "Brainvita"
End
End If
End Sub


Private Sub cmdscore_Click()
Dim temp As String

Open "BVHigh.brh" For Input As #1
Line Input #1, temp

If temp = "Start" Then
    MsgBox "There are no highscores registered yet.", vbOKOnly + vbExclamation, "Brainvita"
    Close #1
    Exit Sub
End If

Line Input #1, temp
highitem = temp + Space(4 - Len(temp))
Line Input #1, temp
highitem = highitem + temp + Space(7 - Len(temp))
Line Input #1, temp
highitem = highitem + temp + Space(42 - Len(temp))
Line Input #1, temp
highitem = highitem + temp
viewhigh.highlist.AddItem highitem

While Not EOF(1)
    Line Input #1, temp
    highitem = temp + Space(4 - Len(temp))
    Line Input #1, temp
    highitem = highitem + temp + Space(7 - Len(temp))
    Line Input #1, temp
    highitem = highitem + temp + Space(42 - Len(temp))
    Line Input #1, temp
    highitem = highitem + temp
    viewhigh.highlist.AddItem highitem
Wend

Close #1
viewhigh.Show
End Sub

Private Sub cmdscoring_Click()
MsgBox "Brianvita - Scoring" & vbCrLf & "If balls remaining = " & vbCrLf & _
       "    1 - GENIUS" & vbCrLf & _
       "    2 - OUTSTANDING " & vbCrLf & _
       "    3 - GOOD" & vbCrLf & _
       " > 4 - Need practice.", vbOKOnly + vbExclamation, "Brainvita"
End Sub

Private Sub Form_Load()
    
cmdback.Enabled = False
Me.Caption = "Brainvita"
For i = 0 To 48
    board(i).Enabled = True
Next i
    
gameend = False
score = 32
elapsedtime = 0
paused = False
cmdpause.Caption = "Pause"
Timer1.Enabled = True
cmdback.Enabled = False
backcount = 0
lbltime.Caption = "0 seconds"
lblscore.Caption = "32 remaining out of 32"
lblstatus = "Please select a box."
selected = False
previous = 100
cmdpause.Enabled = True
For i = 0 To 8
    For j = 0 To 8
        boxfilled(i, j) = True
    Next j
Next i
boxfilled(4, 4) = False

For i = 0 To 48
    board(i).Appearance = 1
    board(i).Picture = mainpict.Picture
    board(i).BackColor = &H8000000F
Next i

board(24).Picture = emptypict.Picture

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Are you sure to quit this game ?", vbYesNo + vbQuestion, "Brainvita") = vbYes Then

MsgBox "Thanks for playing." & vbCrLf & _
       "Did you like the game ?" & vbCrLf & _
       "Tell MOHD FRAHIM at mdfrahim@yahoo.com", vbOKOnly + vbExclamation, "Brainvita"
End
End If
Cancel = True
End Sub
Private Sub Timer1_Timer()
elapsedtime = elapsedtime + 1
lbltime.Caption = elapsedtime & " seconds"
Me.Caption = "Brainvita - " & elapsedtime
If elapsedtime > 1800 Then
    MsgBox "Half an hour passed. What the hell are you doing ?" & vbCrLf & _
           "It seems you are not intrested in the game." & vbCrLf & _
           "Brainvita will now close.", vbOKOnly + vbExclamation, "Brainvita"
    End
End If
End Sub

