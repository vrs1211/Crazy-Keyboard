VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crazy Keyboard"
   ClientHeight    =   8910
   ClientLeft      =   3750
   ClientTop       =   2220
   ClientWidth     =   14970
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   14970
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1815
      Left            =   6840
      TabIndex        =   25
      Top             =   3240
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3201
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form3.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Cn"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   8175
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3855
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
         TabIndex        =   23
         Top             =   4560
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
         Index           =   3
         Left            =   1560
         TabIndex        =   22
         Top             =   3840
         Width           =   2175
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
         Left            =   120
         TabIndex        =   21
         Top             =   4680
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
         TabIndex        =   20
         Top             =   3960
         Width           =   1215
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
         Left            =   120
         TabIndex        =   19
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Roboto Cn"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   480
         TabIndex        =   18
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
         Left            =   120
         TabIndex        =   17
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
         TabIndex        =   16
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
         Left            =   120
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   3240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   8175
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      Begin RichTextLib.RichTextBox RichTextBox2 
         Height          =   1575
         Left            =   2760
         TabIndex        =   26
         Top             =   5040
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   2778
         _Version        =   393217
         ReadOnly        =   -1  'True
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Form3.frx":007F
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   10080
         Top             =   360
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   2
         Height          =   855
         Left            =   2760
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label14 
         Caption         =   "What You Get"
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
         Left            =   480
         TabIndex        =   24
         Top             =   5280
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "LEVEL"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   18
            Charset         =   255
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
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
         Height          =   495
         Left            =   5040
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Target Text"
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
         Left            =   480
         TabIndex        =   9
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   2760
         TabIndex        =   8
         Top             =   1440
         Width           =   6615
      End
      Begin VB.Label Label10 
         Caption         =   "What you typed"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label11 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label13 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   5
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label15 
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
         Left            =   5040
         TabIndex        =   4
         Top             =   7440
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
         Left            =   5160
         TabIndex        =   3
         Top             =   6720
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
         Left            =   720
         TabIndex        =   2
         Top             =   6840
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
         Left            =   720
         TabIndex        =   1
         Top             =   7560
         Width           =   3135
      End
      Begin VB.Shape Shape2 
         Height          =   495
         Left            =   600
         Top             =   7440
         Width           =   8775
      End
      Begin VB.Shape Shape3 
         Height          =   495
         Left            =   600
         Top             =   6720
         Width           =   8775
      End
   End
   Begin VB.Label Label21 
      Caption         =   "Label21"
      Height          =   255
      Left            =   6840
      TabIndex        =   28
      Top             =   8640
      Width           =   3375
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   15
      Left            =   4440
      TabIndex        =   27
      Top             =   8760
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      Height          =   8415
      Left            =   0
      Top             =   0
      Width           =   14775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l As Integer
Dim cnt As Integer
Dim temp As Integer
Dim current As String


Dim tempst As String




Private Sub Command2_Click()
Form2.Show

End Sub

Private Sub Form_Load()
Label2.Caption = fname
score = 0
lvlq(1, 1) = "Anyone who doesn't take truth seriously in small matters cannot be trusted in large ones either."
lvlq(1, 2) = "The pursuit of truth and beauty is a sphere of activity in which we are permitted to remain children all our lives."
lvlq(1, 3) = "Any man who reads too much and uses his own brain too little falls into lazy habits of thinking."
lvlq(1, 4) = "Only two things are infinite, the universe and human stupidity, and I'm not sure about the former."
lvlq(1, 5) = "If people are good only because they fear punishment, and hope for reward, then we are a sorry lot indeed"
lvlq(1, 6) = "If people are good only because they fear punishment, and hope for reward, then we are a sorry lot indeed"
lvla(1, 1) = "Zmblmv dsl wlvhm'g gzpv gifgs hvirlfhob rm hnzoo nzggvih xzmmlg yv gifhgvw rm ozitv lmvh vrgsvi."
lvla(1, 2) = "Gsv kfihfrg lu gifgs zmw yvzfgb rh z hksviv lu zxgrergb rm dsrxs dv ziv kvinrggvw gl ivnzrm xsrowivm zoo lfi orevh."
lvla(1, 3) = "Zmb nzm dsl ivzwh gll nfxs zmw fhvh srh ldm yizrm gll orggov uzooh rmgl ozab szyrgh lu gsrmprmt."
lvla(1, 4) = "Lmob gdl gsrmth ziv rmurmrgv, gsv fmrevihv zmw sfnzm hgfkrwrgb, zmw R'n mlg hfiv zylfg gsv ulinvi."
lvla(1, 5) = "Ru kvlkov ziv tllw lmob yvxzfhv gsvb uvzi kfmrhsnvmg, zmw slkv uli ivdziw, gsvm dv ziv z hliib olg rmwvvw."
lvla(1, 6) = "Ru kvlkov ziv tllw lmob yvxzfhv gsvb uvzi kfmrhsnvmg, zmw slkv uli ivdziw, gsvm dv ziv z hliib olg rmwvvw."
c1 = 600
lvlscore(1) = 0
lvl = 1
lvltmrcnt(1) = 0
lvltmrcnt(2) = 0
lvltmrcnt(3) = 0
lvltmrcnt(4) = 0
lvltmrcnt(5) = 0
lvltmrcnt(lvl) = 600
Label7.Caption = lvl
Randomize
qrnd = (5 - 1 + 1) * Rnd + 1
Label9.Caption = lvlq(1, qrnd)
Timer1.Enabled = True



End Sub










Private Sub RichTextBox1_Change()
If 1 > 2 Or RichTextBox1.Text = lvla(lvl, qrnd) Then
Timer1.Enabled = False
lvltmrcnt(lvl + 1) = lvltmrcnt(lvl)
Label21.Caption = c1
Label6.Caption = "Success"
lvlscore(1) = score
Label4(0).Caption = lvlscore(1)
MsgBox ("Level Completed.In the Next Level" & vbCrLf & "you have to press submit button" & vbCrLf & ". Press ok for next level.")
Form3.Hide
Form2.Show



Else
RichTextBox2.Text = ""
l = Len(RichTextBox1.Text)
If l > 0 Then
cnt = 1
Do While cnt <= l
temp = Asc(Mid(RichTextBox1.Text, cnt, 1))
   If (48 <= temp And temp <= 57) Then
    temp = 48 + (57 - temp)
    End If
    
   If (65 <= temp And temp <= 90) Then
    temp = Asc(Chr(65 + (90 - temp)))
    Else
    If (96 < temp And temp < 123) Then
    temp = Asc(Chr(97 + 122 - temp))
    End If
    End If
tempst = RichTextBox2.Text

RichTextBox2.Text = tempst + Chr(temp)
current = RichTextBox2.Text
cnt = cnt + 1
Loop
End If

End If
End Sub



Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then KeyCode = 0
If KeyCode = 45 And Shift = 1 Then KeyCode = 0 'paste(shift + insert)
If KeyCode = 45 And Shift = 2 Then KeyCode = 0 'copy(ctrl + insert)
If KeyCode = 86 And Shift = 2 Then KeyCode = 0 'paste(ctrl + v)
If KeyCode = 67 And Shift = 2 Then KeyCode = 0 'copy(ctrl + c)
If KeyCode = 88 And Shift = 2 Then KeyCode = 0 'cut(ctrl + x)
End Sub

Private Sub RichTextBox2_Change()
If (Len(RichTextBox2.Text) > Len(RichTextBox1.Text)) Then
RichTextBox2.Text = current
End If
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
lvltmrcnt(lvl) = lvltmrcnt(lvl) - 1
score = score + 1
Label15.Caption = lvltmrcnt(lvl) & "   seconds"

If lvltmrcnt(lvl) = 0 Then
    Timer1.Enabled = False
    MsgBox ("TIME UP. GAMEOVE" & vbCrLf & "press ok to view final score card")
    Form4.Show
    Unload Form1
    Unload Form2
    Form1.Show
    Unload Form3
End If

End Sub
