VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form5"
   ClientHeight    =   9720
   ClientLeft      =   7455
   ClientTop       =   1605
   ClientWidth     =   5775
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   5775
   Begin VB.Frame Frame1 
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin VB.Frame Frame2 
         Caption         =   "CREDITS"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   240
         TabIndex        =   3
         Top             =   6360
         Width           =   4935
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Roboto Cn"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            TabIndex        =   7
            Top             =   1440
            Width           =   4215
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label 1"
            BeginProperty Font 
               Name            =   "Roboto Cn"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            TabIndex        =   6
            Top             =   2040
            Width           =   4215
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Roboto Cn"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            TabIndex        =   5
            Top             =   840
            Width           =   4215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label 1"
            BeginProperty Font 
               Name            =   "Roboto Cn"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   4215
         End
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   4215
         Left            =   240
         TabIndex        =   1
         Top             =   2040
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   7435
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Form5.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Roboto Cn"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Crazy Keyboard"
         BeginProperty Font 
            Name            =   "Roboto Lt"
            Size            =   39.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   480
         TabIndex        =   2
         Top             =   120
         Width           =   4215
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim rule As String
rule = "*It consists of 5 Levels." & vbCrLf & "*In Level 1, what you type is not what you get... type the crazy text to match the target text" & vbCrLf & "*Each level would show you an example of --target text example-- and --crazy text example-- " & vbCrLf & "*These two examples will give you the hint of how (by what logic)the text is jumbled. You have to understand this logic and use the same logic to type in your --crazy answer-- that results in the --target text-- of that level." & vbCrLf & "However, any punctuations will remain the same ireespective of the level" & vbCrLf & "*The total time for completing all 5 levels is 10 minutes. " & vbCrLf & "*Try to complete all five levels in the minimum possible time." & vbCrLf & vbCrLf & vbCrLf & "*Additional instructions(if any) about any level will be given before start of each level" & vbCrLf
Label1.Caption = "Coded/Designed/Tested By Vikas singh BE/A/2013-2014"
Label2.Caption = "Crazy Logic Provided vy Vishakha Patil BE/A/2013-2014"
Label4.Caption = "Crazy Text Provided by Ashwin Iyer BE/A/2013-2014"
Label5.Caption = "Special Thanks to Rohan Sapkal BE/A/2013-2014 for providing the level 1 snippet"




RichTextBox1.Text = rule


End Sub

