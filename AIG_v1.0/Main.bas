Attribute VB_Name = "modMain"
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

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Greetings(1 To 500) As String
Public Keywords(1 To 1000) As String
Public Answers(1 To 1000) As String
Public TotKeywords As Integer
Public TotGreetings As Integer
Public TotAnswers As Integer

Public Sub Initialize()
Dim Data As String
Dim CharData As String
Dim cCharData As Long
Dim Mode As String
Dim cGreetings As Integer
Dim cKeywords As Integer
Dim cAnswers As Integer

For i = 1 To 50
    Greetings(i) = ""
Next i

For i = 1 To 1000
    Keywords(i) = ""
Next i

For i = 1 To 1000
    Answers(i) = ""
Next i

cGreetings = 1
cKeywords = 1
cAnswers = 1
TotKeywords = 0
TotGreetings = 0
TotAnswers = 0

Open App.Path & "\AIG.brain" For Input As #1
Do
Line Input #1, Data
    Select Case Data
        Case "[Greetings]"
            Mode = "Greetings"
        Case "[Keywords]"
            Mode = "Keywords"
        Case "[Answers]"
            Mode = "Answers"
        Case Else
            Select Case Mode
                Case "Greetings"
                    If cGreetings <= 500 Then
                        Greetings(cGreetings) = Trim(Data)
                        cGreetings = cGreetings + 1
                        TotGreetings = TotGreetings + 1
                    End If
                Case "Keywords"
                    If cKeywords <= 999 Then
                        cCharData = 0
                        Do
                        cCharData = cCharData + 1
                        CharData = Mid(Data, cCharData, 1)
                        Loop Until CharData = "#"
                        cKeywords = Int(Mid(Trim(LCase(Data)), 1, cCharData - 1))
                        Keywords(cKeywords * 2 - 1) = Int(Mid(Trim(LCase(Data)), 1, cCharData - 1))
                        Keywords(cKeywords * 2) = Keywords(cKeywords * 2) + Mid(Trim(LCase(Data)), cCharData + 1, Len(Data) - cCharData) + " "
                        If Keywords(cKeywords * 2 - 1) > CInt(TotKeywords) Then
                            TotKeywords = TotKeywords + 1
                        End If
                    End If
                Case "Answers"
                    If cAnswers <= 999 Then
                        cCharData = 0
                        Do
                        cCharData = cCharData + 1
                        CharData = Mid(Data, cCharData, 1)
                        Loop Until CharData = "#"
                        cAnswers = Int(Mid(Trim(LCase(Data)), 1, cCharData - 1))
                        Answers(cAnswers * 2 - 1) = Int(Mid(Trim(LCase(Data)), 1, cCharData - 1))
                        Answers(cAnswers * 2) = Mid(Trim(Data), cCharData + 1, Len(Data) - cCharData)
                        If Int(Answers(cAnswers * 2 - 1)) > TotAnswers Then
                            TotAnswers = TotAnswers + 1
                        End If
                    End If
            End Select
    End Select
Loop Until EOF(1)

End Sub

Public Function Check(WList As String, KWord As Integer) As Integer

WList = Trim(LCase(WList))

If WList = "" Then
    Check = 0
    Exit Function
End If

Dim SepWList(1 To 999) As String
Dim KWList(1 To 100) As String
Dim cWList As Integer
Dim TotWords As Integer
Dim TotKWords As Integer
Dim Letter As String
Dim cLetter As Integer
Dim oldcLetter As Integer

For i = 1 To 100
    KWList(i) = ""
Next i

For i = 1 To 999
    SepWList(i) = ""
Next i

Check = 0

cLetter = 0
oldcLetter = 1
cWList = 1
TotWords = 0
TotKWords = 0

Do
    Do
        cLetter = cLetter + 1
        Letter = Mid(WList, cLetter, 1)
    Loop Until Letter = " " Or cLetter > Len(WList)
    SepWList(cWList) = Mid(WList, oldcLetter, cLetter - oldcLetter)
    oldcLetter = cLetter + 1
    cWList = cWList + 1
    TotWords = TotWords + 1
    If cWList >= 999 Then Exit Do
Loop Until cLetter >= Len(WList)

KWord = KWord * 2
cLetter = 0
oldcLetter = 1
cWList = 1

Do
    Do
        cLetter = cLetter + 1
        Letter = Mid(Keywords(KWord), cLetter, 1)
    Loop Until Letter = " " Or cLetter > Len(Keywords(KWord))
    KWList(cWList) = Mid(Keywords(KWord), oldcLetter, cLetter - oldcLetter)
    oldcLetter = cLetter + 1
    cWList = cWList + 1
    TotKWords = TotKWords + 1
    If cWList >= 100 Then Exit Do
Loop Until cLetter >= Len(Keywords(KWord))

For i = 1 To TotKWords
    For a = 1 To TotWords
        If KWList(i) = SepWList(a) And Not KWList(i) = "" And Not CommonWord(KWList(i)) And Not CommonWord(SepWList(a)) Then
            Check = Check + 1
            Exit For
        End If
    Next a
Next i

End Function

Public Function CommonWord(Word As String) As Boolean

Word = Trim(LCase(Word))

If Word = "what" Or Word = "what've" Or Word = "what's" Or Word = "what'll" Or _
Word = "you" Or Word = "you've" Or Word = "you'll" Or Word = "your" Or Word = "yours" Or _
Word = "have" Or Word = "haven't" Or _
Word = "will" Or _
Word = "would" Or Word = "wouldn't" Or _
Word = "won't" Or Word = "can't" Or Word = "don't" Or _
Word = "is" Or Word = "isn't" Or _
Word = "it" Or _
Word = "how" Or _
Word = "me" Or _
Word = "i" Or Word = "i've" Or Word = "i'll" Then
    CommonWord = True
Else
    CommonWord = False
End If

End Function

Public Function Check2(Question As String) As Integer

Dim PossAns(1 To 2) As Integer
Dim Check1 As Integer

PossAns(1) = 0

For i = 1 To TotKeywords
    Check1 = Check(Question, Int(i))
    If Check1 > PossAns(2) Then
        PossAns(1) = i
        PossAns(2) = Check1
    End If
Next i

Check2 = PossAns(1)

End Function

Public Sub PrintAns(Message As String)

Dim SepWords(1 To 999) As String
Dim TotWords As Integer
Dim cWords As Integer
Dim Letter As String
Dim cLetter As Integer
Dim oldcLetter As Integer

frmMain.txtAnswer.Text = ""
TotWords = 0
cWords = 1
cLetter = 0
oldcLetter = 1

Do
    Do
        cLetter = cLetter + 1
        Letter = Mid(Message, cLetter, 1)
    Loop Until Letter = " " Or cLetter > Len(Message)
    SepWords(cWords) = Mid(Message, oldcLetter, cLetter - oldcLetter)
    oldcLetter = cLetter + 1
    cWords = cWords + 1
    TotWords = TotWords + 1
    If cWList >= 999 Then Exit Do
Loop Until cLetter >= Len(Message)

For i = 1 To TotWords
    If SepWords(i) = "@time@" Then
        frmMain.txtAnswer.Text = frmMain.txtAnswer.Text + CStr(Format(Time(), "Short Time"))
    ElseIf SepWords(i) = "@date@" Then
        frmMain.txtAnswer.Text = frmMain.txtAnswer.Text + CStr(Date)
    ElseIf SepWords(i) = "@old@" Then
        frmMain.txtAnswer.Text = frmMain.txtAnswer.Text + CStr(Int((Date - DateSerial(2002, 5, 31)) / 365))
    ElseIf SepWords(i) = "@tot_answers@" Then
        frmMain.txtAnswer.Text = frmMain.txtAnswer.Text + CStr(TotAnswers)
    ElseIf SepWords(i) = "@tot_keywords@" Then
        frmMain.txtAnswer.Text = frmMain.txtAnswer.Text + CStr(TotKeywords)
    ElseIf SepWords(i) = "@tot_greetings@" Then
        frmMain.txtAnswer.Text = frmMain.txtAnswer.Text + CStr(TotGreetings)
    ElseIf SepWords(i) = "@greeting@" Then
        Randomize
        nGreeting = CInt(Rnd * TotGreetings)
        If nGreeting = 0 Then nGreeting = 1
        frmMain.txtAnswer.Text = frmMain.txtAnswer.Text + Greetings(nGreeting)
    Else
        frmMain.txtAnswer.Text = frmMain.txtAnswer.Text + SepWords(i) + " "
    End If
Next i

End Sub
