VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crazy Keyboard"
   ClientHeight    =   8010
   ClientLeft      =   7455
   ClientTop       =   4050
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   6690
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   5520
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Crazy Keyboard"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   5160
      Width           =   6255
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   2160
         TabIndex        =   10
         Top             =   1080
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF80&
         Caption         =   "READ FIRST!"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   12
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         MaskColor       =   &H0080FF80&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "START"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Coded by Vikas Singh"
         Height          =   255
         Left            =   3960
         TabIndex        =   7
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "      Start Code"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "         Name   "
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CRAZY KEYBOARD"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   4440
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      Height          =   7815
      Left            =   120
      Top             =   120
      Width           =   6495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Computer Society of India, DMCE - 2013-2014"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3480
      TabIndex        =   6
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Computer Association for Technological Trendz, DMCE- 2013-2014"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3000
      Left            =   3480
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3000
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3000
      Left            =   240
      Picture         =   "Form1.frx":3CAA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Private Sub Command1_Click()
If Text2.Text = "asd" Then
fname = Trim(Text1.Text)
Form3.Show
Form1.Hide


End If

End Sub

Private Sub Command2_Click()
Form5.Show
Command1.Enabled = True


End Sub

Private Sub Form_Load()
Text2.PasswordChar = "*"
End Sub

