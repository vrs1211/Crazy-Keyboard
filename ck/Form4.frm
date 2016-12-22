VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   8370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3930
   LinkTopic       =   "Form4"
   ScaleHeight     =   8370
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Score Board"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "DONOT CLOSE UNTIL YOU REPORT THIS SCORE"
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   7440
         Width           =   2895
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   6720
         Width           =   2055
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000C0&
         BorderStyle     =   5  'Dash-Dot-Dot
         FillColor       =   &H00C0FFFF&
         FillStyle       =   7  'Diagonal Cross
         Height          =   615
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   6600
         Width           =   3135
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "            TOTAL            "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   6120
         Width           =   3135
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   7695
         Left            =   120
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Roboto Cn"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   11
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Level 1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "-------------"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   9
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Level 2"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "-------------"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   7
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "-------------"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   6
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label18 
         Caption         =   "Level 3"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Level 4"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Level 5"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "-------------"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   3
         Left            =   1560
         TabIndex        =   2
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "-------------"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   4
         Left            =   1560
         TabIndex        =   1
         Top             =   4560
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Label2.Caption = fname
Dim tscore As Integer
Dim x As Integer
tscore = 0
For x = 1 To 5
Label4(x - 1) = "+" & lvlscore(x) & " seconds"
tscore = tscore + lvlscore(x)
Label8.Caption = tscore
Next
'Label4(0).Caption = Form2.Label4(0).Caption

End Sub

