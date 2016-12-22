VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   11385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15390
   LinkTopic       =   "Form6"
   ScaleHeight     =   11385
   ScaleWidth      =   15390
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Crazy"
      BeginProperty Font 
         Name            =   "Roboto Cn"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   10815
      Left            =   4320
      TabIndex        =   12
      Top             =   360
      Width           =   10815
      Begin RichTextLib.RichTextBox RichTextBox2 
         Height          =   2895
         Left            =   480
         TabIndex        =   23
         Top             =   1800
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5106
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Form6.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3135
         Left            =   480
         TabIndex        =   22
         Top             =   5160
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5530
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Form6.frx":0083
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   10080
         Top             =   360
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00008000&
         Caption         =   "Submit"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   12
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   8400
         Width           =   9015
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Type below"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   21
         Top             =   4800
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "LEVEL"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   18
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         TabIndex        =   20
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   18
            Charset         =   255
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5280
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Text"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   18
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label15"
         BeginProperty Font 
            Name            =   "Corbel"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   4920
         TabIndex        =   17
         Top             =   9960
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Pending"
         BeginProperty Font 
            Name            =   "Corbel"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   5040
         TabIndex        =   16
         Top             =   9360
         Width           =   4215
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Answer Status"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   600
         TabIndex        =   15
         Top             =   9480
         Width           =   4095
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Time Remaining"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   600
         TabIndex        =   14
         Top             =   10080
         Width           =   3135
      End
      Begin VB.Shape Shape2 
         Height          =   495
         Left            =   480
         Top             =   9960
         Width           =   9015
      End
      Begin VB.Shape Shape3 
         Height          =   495
         Left            =   480
         Top             =   9360
         Width           =   9015
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   2
         Height          =   855
         Left            =   2880
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Score Board"
      BeginProperty Font 
         Name            =   "Roboto Cn"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   10695
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   480
         TabIndex        =   11
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Level 1"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1920
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
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
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
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Level 4"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Level 5"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   4560
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
   Begin VB.Shape Shape1 
      BorderStyle     =   2  'Dash
      BorderWidth     =   3
      Height          =   11175
      Left            =   120
      Top             =   120
      Width           =   15135
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If RichTextBox1.Text = Trim(lvla(5, 1)) Then
Timer1.Enabled = False
Label6.Caption = "Success"
lvlscore(5) = score
Label4(0).Caption = score & vbCrLf & " seconds"
Label6.Caption = success
MsgBox ("all levels completer.press ok to view final score card")
Form4.Show
Me.Hide
Else
Label6.Caption = "try again"
End If




End Sub

Private Sub Form_Load()
score = 0
Label4(0).Caption = lvlscore(1)
Label4(1).Caption = lvlscore(2)
Label4(2).Caption = lvlscore(3)
Label4(3).Caption = lvlscore(4)
lvl = 5
lvlscore(lvl) = 0
Label2.Caption = fname
Label7.Caption = lvl

Randomize
qrnd = Int((5 - 1 + 1) * Rnd + 1)
RichTextBox2.Text = Trim(lvlq(5, qrnd))
Timer1.Enabled = True
End Sub

Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then KeyCode = 0
If KeyCode = 45 And Shift = 1 Then KeyCode = 0 'paste(shift + insert)
If KeyCode = 45 And Shift = 2 Then KeyCode = 0 'copy(ctrl + insert)
If KeyCode = 86 And Shift = 2 Then KeyCode = 0 'paste(ctrl + v)
If KeyCode = 67 And Shift = 2 Then KeyCode = 0 'copy(ctrl + c)
If KeyCode = 88 And Shift = 2 Then KeyCode = 0 'cut(ctrl + x)

End Sub

Private Sub RichTextBox2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then KeyCode = 0
If KeyCode = 45 And Shift = 1 Then KeyCode = 0 'paste(shift + insert)
If KeyCode = 45 And Shift = 2 Then KeyCode = 0 'copy(ctrl + insert)
If KeyCode = 86 And Shift = 2 Then KeyCode = 0 'paste(ctrl + v)
If KeyCode = 67 And Shift = 2 Then KeyCode = 0 'copy(ctrl + c)
If KeyCode = 88 And Shift = 2 Then KeyCode = 0 'cut(ctrl + x)

End Sub

Private Sub Timer1_Timer()

lvltmrcnt(5) = lvltmrcnt(5) - 1
If lvltmrcnt(5) = 0 Then
MsgBox ("TIME UP. GAMEOVE" & vbCrLf & "press ok to view final score card")
Form4.Show
Else
score = score + 1
Label15.Caption = lvltmrcnt(5) & "   seconds"
End If

End Sub

