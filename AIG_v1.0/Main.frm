VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00AF6136&
   Caption         =   " AIG - Artificial Intelligence "
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00D8A794&
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "Main.frx":0442
   MousePointer    =   99  'Custom
   ScaleHeight     =   2655
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAnswer 
      Appearance      =   0  'Flat
      BackColor       =   &H00AF6136&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D8A794&
      Height          =   1695
      HideSelection   =   0   'False
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox txtQuestion 
      Appearance      =   0  'Flat
      BackColor       =   &H00AF6136&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D8A794&
      Height          =   405
      HideSelection   =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblCredits 
      AutoSize        =   -1  'True
      BackColor       =   &H00AF6136&
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Francois van Niekerk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D8A794&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   2730
   End
   Begin VB.Label cmdExit 
      AutoSize        =   -1  'True
      BackColor       =   &H00AF6136&
      BackStyle       =   0  'Transparent
      Caption         =   " Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D8A794&
      Height          =   195
      Left            =   3360
      TabIndex        =   2
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label cmdAsk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D8A794&
      Height          =   195
      Left            =   3480
      MouseIcon       =   "Main.frx":074C
      TabIndex        =   1
      Top             =   210
      Width           =   270
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------AIG v1.0---------'
'
'--------Created by--------'
'----Francois van Niekerk--'
'
'  Please do not use code
' or copy this code without
' the proper permission from
'    its author, you may
'   however use ideas from
'  this code with pleasure,
'  just please include the
' author's name. Thank you.


Private Sub cmdAsk_Click()

If Trim(txtQuestion.Text) = "" Then Exit Sub

Dim Question As String
Dim PAnswer As Integer

Question = Trim(txtQuestion.Text)
txtQuestion.SelStart = 0
txtQuestion.SelLength = Len(txtQuestion.Text)
txtAnswer.Text = "Thinking..."

Do
    If Mid(Question, Len(Question), 1) = "?" Or Mid(Question, Len(Question), 1) = "." Or Mid(Question, Len(Question), 1) = " " Then
        Question = Mid(Question, 1, Len(Question) - 1)
    Else
        Exit Do
    End If
Loop

PAnswer = Check2(Question)

If PAnswer = 0 Then
    txtAnswer.Text = "I don't understand you. Please try explain yourself better."
Else
    'txtAnswer.Text = Answers(PAnswer * 2)
    PrintAns (Answers(PAnswer * 2))
End If

End Sub

Private Sub cmdAsk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAsk.ForeColor = vbWhite
End Sub

Private Sub cmdAsk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAsk.ForeColor = &HD8A794
End Sub

Private Sub cmdExit_Click()
Unload Me
End
End Sub

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdExit.ForeColor = vbWhite
End Sub

Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdExit.ForeColor = &HD8A794
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call cmdAsk_Click
End If
End Sub

Private Sub Form_Load()
Initialize
Randomize
Dim nGreeting As Integer
nGreeting = CInt(Rnd * TotGreetings)
If nGreeting = 0 Then nGreeting = 1
txtAnswer.Text = "Hi! " + Greetings(nGreeting) + " My name is AIG, and I have " + Str(TotAnswers) + " ideas in my brain."
End Sub

Private Sub Form_Resize()
If Me.Width < 3600 Then Me.Width = 3600
If Me.Height < 795 + 480 + 495 Then Me.Height = 795 + 480 + 495
txtAnswer.Height = Me.Height - (795 + 600)
txtAnswer.Width = Me.Width - 360
txtQuestion.Height = 285
txtQuestion.Width = Me.Width - 840
cmdAsk.Left = Me.Width - (495 + 115)
cmdExit.Top = Me.Height - (495 + 195)
cmdExit.Left = Me.Width - (495 + 195)
lblCredits.Top = cmdExit.Top
End Sub

Private Sub lblCredits_Click()
Call ShellExecute(hWnd, "Open", "mailto:flash_slash@hotmail.com?subject=AIG_v1.0", 0, 0, 0)
End Sub
