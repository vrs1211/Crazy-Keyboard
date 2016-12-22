VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crazy Keyboard"
   ClientHeight    =   11490
   ClientLeft      =   3750
   ClientTop       =   1605
   ClientWidth     =   15495
   BeginProperty Font 
      Name            =   "Corbel"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11490
   ScaleWidth      =   15495
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Crazy"
      BeginProperty Font 
         Name            =   "Roboto Cn"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   10815
      Left            =   4200
      TabIndex        =   2
      Top             =   360
      Width           =   10815
      Begin RichTextLib.RichTextBox RichTextBox2 
         Height          =   2415
         Left            =   3600
         TabIndex        =   27
         Top             =   3240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4260
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Form2.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Roboto Cn"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   2535
         Left            =   3600
         TabIndex        =   26
         Top             =   5760
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4471
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"Form2.frx":007F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Roboto Cn"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00008000&
         Caption         =   "Submit"
         Height          =   795
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   8400
         Width           =   9015
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
         Left            =   3600
         Top             =   720
         Width           =   4455
      End
      Begin VB.Shape Shape3 
         Height          =   495
         Left            =   480
         Top             =   9360
         Width           =   9015
      End
      Begin VB.Shape Shape2 
         Height          =   495
         Left            =   480
         Top             =   9960
         Width           =   9015
      End
      Begin VB.Label Label17 
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
         TabIndex        =   20
         Top             =   10080
         Width           =   3135
      End
      Begin VB.Label Label16 
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
         TabIndex        =   19
         Top             =   9480
         Width           =   4095
      End
      Begin VB.Label Label6 
         Caption         =   "Pending"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   5040
         TabIndex        =   18
         Top             =   9360
         Width           =   4215
      End
      Begin VB.Label Label15 
         Caption         =   "Label15"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   4920
         TabIndex        =   14
         Top             =   9960
         Width           =   2055
      End
      Begin VB.Label Label14 
         Caption         =   "enter your crazy answer"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   13
         Top             =   6840
         Width           =   2895
      End
      Begin VB.Label Label12 
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
         Left            =   600
         TabIndex        =   12
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   3600
         TabIndex        =   11
         Top             =   2520
         Width           =   5865
      End
      Begin VB.Label Label10 
         Caption         =   "answer"
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
         Left            =   600
         TabIndex        =   10
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   3600
         TabIndex        =   9
         Top             =   1680
         Width           =   5850
      End
      Begin VB.Label Label8 
         Caption         =   "Crazy hint"
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
         Left            =   600
         TabIndex        =   8
         Top             =   1800
         Width           =   1935
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
         Height          =   615
         Left            =   6600
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Left            =   4080
         TabIndex        =   6
         Top             =   840
         Width           =   1215
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         Left            =   240
         TabIndex        =   23
         Top             =   4560
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
         Left            =   240
         TabIndex        =   22
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
         Left            =   240
         TabIndex        =   21
         Top             =   3360
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
         Index           =   2
         Left            =   1560
         TabIndex        =   17
         Top             =   3240
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
         Index           =   1
         Left            =   1560
         TabIndex        =   16
         Top             =   2640
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
         TabIndex        =   5
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
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   4
         Top             =   1920
         Width           =   2175
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
         TabIndex        =   3
         Top             =   1920
         Width           =   1215
      End
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
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   960
         Width           =   2415
      End
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      Height          =   11295
      Left            =   120
      Top             =   120
      Width           =   15255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub Command1_Click()

'If lvl < 6 Then
lvl5:
    If lvl < 5 Then
        If Trim(LCase(RichTextBox1.Text)) = Trim(LCase(lvla(lvl, qrnd))) Then
            Timer1.Enabled = False
            Label6.Caption = "Success"
            lvltmrcnt(lvl + 1) = lvltmrcnt(lvl)
            RichTextBox1.Text = ""
            Label4(lvl - 1).Caption = lvlscore(lvl)
            Label6.ForeColor = &HC000&
            MsgBox ("press ok for next level")
            lvl = lvl + 1
            If lvl < 5 Then
                If lvl = 4 Then
                Label14.Caption = "enter plain text"
                Label12.Caption = "crazy text"
                MsgBox ("In this Level 4, You have to" & vbccrlf & "ENTER PLAIN TEXT for given crazy Text")
                End If
                Randomize
                qrnd = Int((5 - 1 + 1) * Rnd + 1)
                Label7.Caption = lvl
                Label9.Caption = hintq(lvl, 1) 'replace 1 with qrnd after editing all strings
                Label11.Caption = hinta(lvl, 1) 'replace 1 with qrnd after editing all strings
                RichTextBox2.Text = lvlq(lvl, qrnd)  'replace 1 with qrnd after editing all strings
                lvlscore(lvl) = 0
                Label6.Caption = "Pending"
                Timer1.Enabled = True
        
            
            Else
            GoTo lvl5
            End If
        
        
        
        Else

            Label6.ForeColor = &HFF&
            Label6.Caption = "      Try Again"
        End If
    Else
        
        
        MsgBox ("This is the last Level.Just type the given text fast" & vbCrLf & "It is case sensitive")
        Form6.Show
        Me.Hide
    End If
'Else
'done:
    'MsgBox ("CONGRATULATIONS>>> All Levels Completed! Click ok to view Final score card")
    'Form2.Hide
    'Form4.Show
    
    
    
'End If


End Sub

Private Sub Command3_Click()
Form2.Show

End Sub

Private Sub Form_Load()
hintq(2, 1) = "datta meghe"
hinta(2, 1) = "attad ehgem"
hintq(3, 1) = "we are proud engineers"
hinta(3, 1) = "ew are oudpr ersengine"
hintq(4, 1) = "i leki tyconolhge"
hinta(4, 1) = "i like technology"

'string initializations
lvlq(2, 1) = "and thus i clothe my naked villainy.with old odd ends,stolen forth of holy writ;and seem a saint,when most i play the devil."
lvlq(2, 2) = "he who has injured thee was either stronger or weaker than thee.if weaker,then spare him;if stronger,spare thyself."
lvlq(2, 3) = "you must fight conceit,envy,and every kind of ill-feeling in your heart.the universe seems neither benign nor hostile,merely indifferent."
lvlq(2, 4) = "i dont have a problem believing in God.But in Genesis one has to wonder about these sentences that just go on and end without finishing."
lvlq(2, 5) = "dont forget that the only two things people read in a story are the first and last sentences.Give them blood in the eye on the first one."
lvlq(2, 6) = "dont forget that the only two things people read in a story are the first and last sentences.Give them blood in the eye on the first one."

lvla(2, 1) = "dna suht i ehtolc ym dekan ynialliv.htiw dlo ddo sdne,nelots htrof fo yloh tirw;dna mees a tnias,nehw tsom i yalp eht lived."
lvla(2, 2) = "eh ohw sah derujni eeht saw rehtie regnorts ro rekaew naht eeht.fi rekaew,neht eraps mih;fi regnorts,eraps flesyht."
lvla(2, 3) = "uoy tsum thgif tiecnoc,yvne,dna yreve dnik fo lli-gnileef ni ruoy traeh.eht esrevinu smees rehtien ngineb ron elitsoh,ylerem tnereffidni."
lvla(2, 4) = "i tnod evah a melborp gniveileb ni dog.tub ni siseneg eno sah ot rednow tuoba eseht secnetnes taht tsuj og no dna dne tuohtiw gnihsinif."
lvla(2, 5) = "tnod tegrof taht eht ylno owt sgniht elpoep daer ni a yrots era eth tsrif dna tsal secnetnes.evig meht doolb ni eht eye no eht tsrif eno."
lvla(2, 6) = "tnod tegrof taht eht ylno owt sgniht elpoep daer ni a yrots era eth tsrif dna tsal secnetnes.evig meht doolb ni eht eye no eht tsrif eno."


lvlq(3, 1) = "free from gross passion or of mirth or anger constant in spirit,not swerving with the blood."
lvlq(3, 2) = "garnished and decked in modest compliment,not working with the eye without the ear."
lvlq(3, 3) = "free from gross passion or of mirth or anger constant in spirit,not swerving with the blood."
lvlq(3, 4) = "garnished and decked in modest compliment,not working with the eye without the ear."
lvlq(3, 5) = "free from gross passion or of mirth or anger constant in spirit,not swerving with the blood."
lvlq(3, 6) = "free from gross passion or of mirth or anger constant in spirit,not swerving with the blood."

lvla(3, 1) = "reef romf ossgr ionpass ro fo rthmir ro geran antconst ni ritspi, not ingswerv ithw the oodbl."
lvla(3, 2) = "hedgarnis and keddec ni estmod entcomplim,not ingwork ithw the eye outwith the ear."
lvla(3, 3) = "reef romf ossgr ionpass ro fo rthmir ro geran antconst ni ritspi, not ingswerv ithw the oodbl."
lvla(3, 4) = "hedgarnis and keddec ni estmod entcomplim,not ingwork ithw the eye outwith the ear."
lvla(3, 5) = "reef romf ossgr ionpass ro fo rthmir ro geran antconst ni ritspi, not ingswerv ithw the oodbl."
lvla(3, 6) = "reef romf ossgr ionpass ro fo rthmir ro geran antconst ni ritspi, not ingswerv ithw the oodbl."

lvlq(4, 1) = "tyeh wlli baerk penas of gsals and ssamh the wwndois of ceachos,and aosl kconk you dnwo wuthoit the ssithgelt comtuncpion.on the cynartro,tyeh wlli rrao whti lruthgea."
lvlq(4, 2) = "tyeh wlli baerk penas of gsals and ssamh the wwndois of ceachos,and aosl kconk you dnwo wuthoit the ssithgelt comtuncpion.on the cynartro,tyeh wlli rrao whti lruthgea."
lvlq(4, 3) = "tyeh wlli baerk penas of gsals and ssamh the wwndois of ceachos,and aosl kconk you dnwo wuthoit the ssithgelt comtuncpion.on the cynartro,tyeh wlli rrao whti lruthgea."
lvlq(4, 4) = "tyeh wlli baerk penas of gsals and ssamh the wwndois of ceachos,and aosl kconk you dnwo wuthoit the ssithgelt comtuncpion.on the cynartro,tyeh wlli rrao whti lruthgea."
lvlq(4, 5) = "tyeh wlli baerk penas of gsals and ssamh the wwndois of ceachos,and aosl kconk you dnwo wuthoit the ssithgelt comtuncpion.on the cynartro,tyeh wlli rrao whti lruthgea."
lvlq(4, 6) = "tyeh wlli baerk penas of gsals and ssamh the wwndois of ceachos,and aosl kconk you dnwo wuthoit the ssithgelt comtuncpion.on the cynartro,tyeh wlli rrao whti lruthgea."

lvla(4, 1) = "they will break panes of glass and smash the windows of coaches,and also knock you down without the slightest compunction.on the contrary,they will roar with laughter."
lvla(4, 2) = "they will break panes of glass and smash the windows of coaches,and also knock you down without the slightest compunction.on the contrary,they will roar with laughter."
lvla(4, 3) = "they will break panes of glass and smash the windows of coaches,and also knock you down without the slightest compunction.on the contrary,they will roar with laughter."
lvla(4, 4) = "they will break panes of glass and smash the windows of coaches,and also knock you down without the slightest compunction.on the contrary,they will roar with laughter."
lvla(4, 5) = "they will break panes of glass and smash the windows of coaches,and also knock you down without the slightest compunction.on the contrary,they will roar with laughter."
lvla(4, 6) = "they will break panes of glass and smash the windows of coaches,and also knock you down without the slightest compunction.on the contrary,they will roar with laughter."

lvlq(5, 1) = "I felt the wall of the tunnel shiver. The master alarm squealed through my earphones. Almost simultaneously, Jack yelled down to me that there was a warning light on. Fleeting but spectacular sights snapped into ans out of view, the snow, the shower of debris, the moon, looming close and big, the dazzling sunshine for once unfiltered by layers of air.The last twelve hours before re-entry were particular bone-chilling. During this period, I had to go up in to command module. Even after the fiery re-entry splashing down in 81 degrees water in south pacific, we could still see our frosty breath inside the command module."
lvlq(5, 2) = "I felt the wall of the tunnel shiver. The master alarm squealed through my earphones. Almost simultaneously, Jack yelled down to me that there was a warning light on. Fleeting but spectacular sights snapped into ans out of view, the snow, the shower of debris, the moon, looming close and big, the dazzling sunshine for once unfiltered by layers of air.The last twelve hours before re-entry were particular bone-chilling. During this period, I had to go up in to command module. Even after the fiery re-entry splashing down in 81 degrees water in south pacific, we could still see our frosty breath inside the command module."
lvlq(5, 3) = "I felt the wall of the tunnel shiver. The master alarm squealed through my earphones. Almost simultaneously, Jack yelled down to me that there was a warning light on. Fleeting but spectacular sights snapped into ans out of view, the snow, the shower of debris, the moon, looming close and big, the dazzling sunshine for once unfiltered by layers of air.The last twelve hours before re-entry were particular bone-chilling. During this period, I had to go up in to command module. Even after the fiery re-entry splashing down in 81 degrees water in south pacific, we could still see our frosty breath inside the command module."
lvlq(5, 4) = "I felt the wall of the tunnel shiver. The master alarm squealed through my earphones. Almost simultaneously, Jack yelled down to me that there was a warning light on. Fleeting but spectacular sights snapped into ans out of view, the snow, the shower of debris, the moon, looming close and big, the dazzling sunshine for once unfiltered by layers of air.The last twelve hours before re-entry were particular bone-chilling. During this period, I had to go up in to command module. Even after the fiery re-entry splashing down in 81 degrees water in south pacific, we could still see our frosty breath inside the command module."
lvlq(5, 5) = "I felt the wall of the tunnel shiver. The master alarm squealed through my earphones. Almost simultaneously, Jack yelled down to me that there was a warning light on. Fleeting but spectacular sights snapped into ans out of view, the snow, the shower of debris, the moon, looming close and big, the dazzling sunshine for once unfiltered by layers of air.The last twelve hours before re-entry were particular bone-chilling. During this period, I had to go up in to command module. Even after the fiery re-entry splashing down in 81 degrees water in south pacific, we could still see our frosty breath inside the command module."
lvlq(5, 6) = "I felt the wall of the tunnel shiver. The master alarm squealed through my earphones. Almost simultaneously, Jack yelled down to me that there was a warning light on. Fleeting but spectacular sights snapped into ans out of view, the snow, the shower of debris, the moon, looming close and big, the dazzling sunshine for once unfiltered by layers of air.The last twelve hours before re-entry were particular bone-chilling. During this period, I had to go up in to command module. Even after the fiery re-entry splashing down in 81 degrees water in south pacific, we could still see our frosty breath inside the command module."

lvla(5, 1) = "I felt the wall of the tunnel shiver. The master alarm squealed through my earphones. Almost simultaneously, Jack yelled down to me that there was a warning light on. Fleeting but spectacular sights snapped into ans out of view, the snow, the shower of debris, the moon, looming close and big, the dazzling sunshine for once unfiltered by layers of air.The last twelve hours before re-entry were particular bone-chilling. During this period, I had to go up in to command module. Even after the fiery re-entry splashing down in 81 degrees water in south pacific, we could still see our frosty breath inside the command module."
lvla(5, 2) = "I felt the wall of the tunnel shiver. The master alarm squealed through my earphones. Almost simultaneously, Jack yelled down to me that there was a warning light on. Fleeting but spectacular sights snapped into ans out of view, the snow, the shower of debris, the moon, looming close and big, the dazzling sunshine for once unfiltered by layers of air.The last twelve hours before re-entry were particular bone-chilling. During this period, I had to go up in to command module. Even after the fiery re-entry splashing down in 81 degrees water in south pacific, we could still see our frosty breath inside the command module."
lvla(5, 3) = "I felt the wall of the tunnel shiver. The master alarm squealed through my earphones. Almost simultaneously, Jack yelled down to me that there was a warning light on. Fleeting but spectacular sights snapped into ans out of view, the snow, the shower of debris, the moon, looming close and big, the dazzling sunshine for once unfiltered by layers of air.The last twelve hours before re-entry were particular bone-chilling. During this period, I had to go up in to command module. Even after the fiery re-entry splashing down in 81 degrees water in south pacific, we could still see our frosty breath inside the command module."
lvla(5, 4) = "I felt the wall of the tunnel shiver. The master alarm squealed through my earphones. Almost simultaneously, Jack yelled down to me that there was a warning light on. Fleeting but spectacular sights snapped into ans out of view, the snow, the shower of debris, the moon, looming close and big, the dazzling sunshine for once unfiltered by layers of air.The last twelve hours before re-entry were particular bone-chilling. During this period, I had to go up in to command module. Even after the fiery re-entry splashing down in 81 degrees water in south pacific, we could still see our frosty breath inside the command module."
lvla(5, 5) = "I felt the wall of the tunnel shiver. The master alarm squealed through my earphones. Almost simultaneously, Jack yelled down to me that there was a warning light on. Fleeting but spectacular sights snapped into ans out of view, the snow, the shower of debris, the moon, looming close and big, the dazzling sunshine for once unfiltered by layers of air.The last twelve hours before re-entry were particular bone-chilling. During this period, I had to go up in to command module. Even after the fiery re-entry splashing down in 81 degrees water in south pacific, we could still see our frosty breath inside the command module."
lvla(5, 6) = "I felt the wall of the tunnel shiver. The master alarm squealed through my earphones. Almost simultaneously, Jack yelled down to me that there was a warning light on. Fleeting but spectacular sights snapped into ans out of view, the snow, the shower of debris, the moon, looming close and big, the dazzling sunshine for once unfiltered by layers of air.The last twelve hours before re-entry were particular bone-chilling. During this period, I had to go up in to command module. Even after the fiery re-entry splashing down in 81 degrees water in south pacific, we could still see our frosty breath inside the command module."

'first laod params
Label4(0).Caption = Form3.Label4(0).Caption

lvl = 2
lvlscore(lvl) = 0
Label2.Caption = fname
Label7.Caption = lvl

Randomize
qrnd = Int((5 - 1 + 1) * Rnd + 1)
Label9.Caption = LCase(Trim(hintq(lvl, 1)))
Label11.Caption = LCase(Trim(hinta(lvl, 1)))
RichTextBox2.Text = Trim(LCase(lvlq(lvl, qrnd)))

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
lvltmrcnt(lvl) = lvltmrcnt(lvl) - 1
lvlscore(lvl) = lvlscore(lvl) + 1
Label15.Caption = lvltmrcnt(lvl) & "  seconds"

If lvltmrcnt(lvl) = 0 Then
Timer1.Enabled = False
MsgBox ("TIME UP. GAMEOVE" & vbCrLf & "press ok to view final score card")
Form4.Show
Unload Form1
Unload Form3
Form2.Hide
Unload Form2

End If

End Sub

