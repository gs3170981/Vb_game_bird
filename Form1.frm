VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "                                                                               "
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7455
   ScaleWidth      =   15285
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "拜拜"
      Height          =   495
      Left            =   3240
      Picture         =   "Form1.frx":141BA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "再来一次"
      Height          =   495
      Left            =   1080
      MaskColor       =   &H0000FF00&
      OLEDropMode     =   1  'Manual
      Picture         =   "Form1.frx":1D98F
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Timer Timer7 
      Interval        =   1
      Left            =   7440
      Top             =   0
   End
   Begin VB.Timer Timer6 
      Interval        =   1
      Left            =   8040
      Top             =   0
   End
   Begin VB.CommandButton z6 
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   5775
      Left            =   22440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton z5 
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   1455
      Left            =   22440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton z4 
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   3615
      Left            =   15720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton z3 
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   3495
      Left            =   15720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton z2 
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   5055
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton z1 
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   1935
      Left            =   9120
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Timer Siwang 
      Interval        =   20
      Left            =   10800
      Top             =   0
   End
   Begin VB.Timer Timer5 
      Interval        =   50
      Left            =   11160
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   11520
      Top             =   0
   End
   Begin VB.PictureBox Tu1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   600
      Picture         =   "Form1.frx":27164
      ScaleHeight     =   855
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   4800
      Width           =   975
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   11880
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   12600
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   12240
      Top             =   0
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FFFF&
      Caption         =   "Game Over"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "最高分："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "分数："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i%, j%, f%, sum%


Public Function Pz(x1 As Long, y1 As Long, w1 As Long, h1 As Long, x2 As Long, y2 As Long, w2 As Long, h2 As Long) As Boolean

    If y2 + h2 >= y1 And y2 <= y1 + h1 And x2 + w2 >= x1 And x2 <= x1 + w1 Then Pz = True
    
End Function



Private Sub Command1_Click()
    Tu1.Left = 450: Tu1.Top = 4680
    Command1.Visible = False: Command2.Visible = False
    z1.Height = 1935: z1.Left = 9120:  z1.Top = 0:    z1.Width = 1095
    z2.Height = 5055: z2.Left = 9120:  z2.Top = 4920: z2.Width = 1095
    z3.Height = 3495: z3.Left = 15720: z3.Top = 0:    z3.Width = 1095
    z4.Height = 5055: z4.Left = 15720: z4.Top = 6360: z4.Width = 1095
    z5.Height = 1455: z5.Left = 22440: z5.Top = 0:    z5.Width = 1095
    z6.Height = 10560: z6.Left = 22440: z6.Top = 4200: z6.Width = 1095
    i = 0: j = 8
    
        Timer1.Enabled = True
        Timer2.Enabled = True
        Timer3.Enabled = True
        Timer6.Enabled = True
        Label5.Visible = False
        Timer5.Enabled = True
        Siwang.Enabled = True
        Timer7.Enabled = True
    
End Sub

Private Sub Command2_Click()
End
End Sub


Private Sub Command3_Click()
 Tu1.Left = 450: Tu1.Top = 4680
    Command1.Visible = False: Command2.Visible = False
    z1.Height = 1935: z1.Left = 9120:  z1.Top = 0:    z1.Width = 1095
    z2.Height = 5055: z2.Left = 9120:  z2.Top = 4920: z2.Width = 1095
    z3.Height = 3495: z3.Left = 15720: z3.Top = 0:    z3.Width = 1095
    z4.Height = 5055: z4.Left = 15720: z4.Top = 6360: z4.Width = 1095
    z5.Height = 1455: z5.Left = 22440: z5.Top = 0:    z5.Width = 1095
    z6.Height = 10560: z6.Left = 22440: z6.Top = 4200: z6.Width = 1095
    i = 0: j = 8
    
        Timer1.Enabled = True
        Timer2.Enabled = True
        Timer3.Enabled = True
        Timer6.Enabled = True
        Label5.Visible = False
        Timer5.Enabled = True
        Siwang.Enabled = True
        Timer7.Enabled = True
End Sub

Private Sub Siwang_Timer()
    
    If Tu1.Top >= 7650 Or Tu1.Top <= 0 Then
        
        Label5.Visible = True
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
        Timer4.Enabled = False
      
        Siwang.Enabled = False
        Timer6.Enabled = False
        Command1.Visible = True: Command2.Visible = True
        Command1.SetFocus
        Timer7.Enabled = False
        Label4.Caption = sum
        If f > sum Then Label4.Caption = f
        Label2.Caption = ""
        f = 0
    End If
    
    If Pz(Tu1.Left, Tu1.Top, Tu1.Width, Tu1.Height, z1.Left, z1.Top, z1.Width, z1.Height) = True Or Pz(Tu1.Left, Tu1.Top, Tu1.Width, Tu1.Height, z2.Left, z2.Top, z2.Width, z2.Height) = True Or Pz(Tu1.Left, Tu1.Top, Tu1.Width, Tu1.Height, z3.Left, z3.Top, z3.Width, z3.Height) = True Or Pz(Tu1.Left, Tu1.Top, Tu1.Width, Tu1.Height, z4.Left, z4.Top, z4.Width, z4.Height) = True Or Pz(Tu1.Left, Tu1.Top, Tu1.Width, Tu1.Height, z5.Left, z5.Top, z5.Width, z5.Height) = True Or Pz(Tu1.Left, Tu1.Top, Tu1.Width, Tu1.Height, z6.Left, z6.Top, z6.Width, z6.Height) = True Then

        Label5.Visible = True
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
        Timer4.Enabled = False
       
        Siwang.Enabled = False
        Timer6.Enabled = False
        Command1.Visible = True: Command2.Visible = True
        Command1.SetFocus
        Timer7.Enabled = False
        Label4.Caption = sum
        If f > sum Then Label4.Caption = f
        Label2.Caption = ""
        f = 0
    End If
    
    
End Sub

Private Sub Timer4_Timer()
    Tu1.Top = Tu1.Top - 500
    Timer5.Enabled = True
    Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
    i = i + 1
    Tu1.Top = Tu1.Top + 15 * i
    If j = 5 Then i = 0: j = 0
End Sub


Private Sub Timer6_Timer()

    z1.Left = z1.Left - 45
    z2.Left = z2.Left - 45
    z3.Left = z3.Left - 45
    z4.Left = z4.Left - 45
    z5.Left = z5.Left - 45
    z6.Left = z6.Left - 45
    
    If z1.Left <= 0 - z1.Width Then z1.Left = Form1.Width: z1.Top = 0: z1.Height = Int(Rnd * (5000 - 1000 + 1000) + 1): z2.Top = z1.Height + 2900
    If z2.Left <= 0 - z2.Width Then z2.Left = Form1.Width
    If z3.Left <= 0 - z3.Width Then z3.Left = Form1.Width: z3.Top = 0: z3.Height = Int(Rnd * (5000 - 1000 + 1000) + 1): z4.Top = z3.Height + 2900
    If z4.Left <= 0 - z4.Width Then z4.Left = Form1.Width
    If z5.Left <= 0 - z5.Width Then z5.Left = Form1.Width: z5.Top = 0: z5.Height = Int(Rnd * (5000 - 1000 + 1000) + 1): z6.Top = z5.Height + 2900
    If z6.Left <= 0 - z6.Width Then z6.Left = Form1.Width
    
    

        
    
End Sub

Private Sub Timer7_Timer()

    f = f + 1
    Label2.Caption = f
    
    
End Sub

Private Sub Tu1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 32 Then Timer3.Enabled = True
    Tu1.Picture = LoadPicture(App.Path & "\2.jpg")
End Sub



Private Sub Timer1_Timer()
Tu1.Picture = LoadPicture(App.Path & "\2.jpg")
Timer2.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()

Tu1.Picture = LoadPicture(App.Path & "\1.jpg")
Timer1.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
    Tu1.Top = Tu1.Top - 600
    j = 5
    Timer4.Enabled = True
    Timer3.Enabled = False
End Sub
