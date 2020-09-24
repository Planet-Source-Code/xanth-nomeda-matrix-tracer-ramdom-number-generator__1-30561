VERSION 5.00
Begin VB.Form frmTrace 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "The Matrix Tracer Utility - Type 'help' or 'command'"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10350
   ControlBox      =   0   'False
   Icon            =   "frmTrace.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   480
      MousePointer    =   1  'Arrow
      TabIndex        =   24
      Top             =   5520
      Width           =   10095
   End
   Begin VB.TextBox txtLength 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2160
      TabIndex        =   22
      Text            =   "7"
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox txtFinalNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   49
      TabIndex        =   21
      Top             =   960
      Width           =   7455
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Debug Info"
      Height          =   495
      Left            =   5880
      TabIndex        =   14
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Timer timFormat 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4680
      Top             =   5880
   End
   Begin VB.CommandButton cmdHalt 
      Caption         =   "Halt Trace"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Timer timScroll 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5280
      Top             =   5880
   End
   Begin VB.TextBox txtTrace 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   3855
      Index           =   6
      Left            =   8760
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "frmTrace.frx":030A
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtTrace 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   3855
      Index           =   5
      Left            =   7320
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "frmTrace.frx":0313
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtTrace 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   3855
      Index           =   4
      Left            =   5880
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "frmTrace.frx":031C
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtTrace 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   3855
      Index           =   3
      Left            =   4440
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frmTrace.frx":0325
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtTrace 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   3855
      Index           =   2
      Left            =   3000
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "frmTrace.frx":032E
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtTrace 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   3855
      Index           =   1
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmTrace.frx":0337
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtTrace 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   3855
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmTrace.frx":0340
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdTrace 
      Caption         =   "Begin Trace"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   255
      Left            =   7800
      TabIndex        =   2
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   7800
      TabIndex        =   1
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset Trace"
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   0
      Picture         =   "frmTrace.frx":0349
      Top             =   0
      Width           =   2100
   End
   Begin VB.Label lblTracing 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblError 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   5160
      Width           =   9855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "$:>"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   8880
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   6
      Left            =   9840
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   5
      Left            =   9600
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   4
      Left            =   9360
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   3
      Left            =   9120
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   8640
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   8400
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   8160
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblComplete 
      BackStyle       =   0  'Transparent
      Caption         =   "!!Trace Complete!!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   3855
   End
End
Attribute VB_Name = "frmTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************
'*This is the completly re-written file by Xanth Nomeda*
'*******************************************************

Dim setNumber As String 'The random number
Dim tCount As Integer
Dim tracing As Boolean
Dim halted As Boolean


Private Sub cmdAbout_Click()

    frmAbout.Show 1, Me

End Sub

Private Sub cmdExit_Click()

    End

End Sub

Private Sub cmdHalt_Click()

    'Stop everything!
    timScroll.Enabled = False
    timFormat.Enabled = False
    halted = True
    tracing = False

End Sub

Private Sub cmdInfo_Click()

'Just debugging info
    Dim msg As String

    For i = 0 To txtTrace.Count - 1
        msg = msg & vbCrLf & "txtTrace(" & i & ").Tag = " & txtTrace(i).Tag
    Next
    
    MsgBox msg, vbInformation, "Trace - Debug Info"

    lblError = ""

End Sub

Private Sub cmdReset_Click()
    
    timScroll.Enabled = False
    timFormat.Enabled = False
    For i = 0 To 6
        txtTrace(i) = ""
        lblNumber(i) = "0"
    Next
    tCount = 0
    lblComplete.Visible = False
    txtLength.Locked = False
    txtFinalNumber = ""
    lblTracing.Visible = False
    tracing = False
    halted = True
    Form_Load

End Sub

Private Sub cmdTrace_Click()

    'New Routine
    Call runTrace
    tracing = True
    halted = False

End Sub


Private Sub Form_Activate()

    txtInput.SetFocus

End Sub

Private Sub Form_Load()
    
    Me.Caption = "The Matrix Tracer Utility v" & App.Major & "." & App.Minor & "." & App.Revision & " - Type 'help' or 'command'"
    
    For i = 0 To 6
        txtTrace(i) = "0000000"
        For j = 1 To 15
            txtTrace(i) = txtTrace(i) & vbCrLf & "0000000"
        Next j
        txtTrace(i).Tag = "#######"
    Next i

End Sub

Private Function runTrace() As Boolean

    setNumber = generateNumber(txtLength, 9) 'Set the Random number thats going to be 'traced'

    lblTracing.Visible = True
    lblTracing = "Tracing: " & txtLength & "..."

    txtLength.Locked = True
    timScroll.Enabled = True
    timFormat.Enabled = True

End Function
'** Generates Random Numbers (duh)
Private Function generateNumber(numLen As Integer, max As Integer, Optional zero As Boolean = True) As String

    For i = 1 To numLen
        Randomize
        If zero = True Then generateNumber = generateNumber & CStr(CInt(max * Rnd(Timer)))
        If zero = False Then generateNumber = generateNumber & CStr(CInt(max * Rnd(Timer) + 1))
    Next

End Function

Private Sub timFormat_Timer()

    Dim start As Integer
    Dim myTag As String
    Dim newStr As String
    Dim pre As String
    Dim tIndex As Integer
    
    'Taken out to make it random which column shrinks
    'Static formatIndex As Integer
    
    'I know, I know, spaghetti code, but hey, Im not goin for code of the millenium here...
NewtIndex:
    tIndex = generateNumber(1, 6)
    If InStr(1, txtTrace(tIndex).Tag, "#") < 1 And tCount < Val(txtLength) Then GoTo NewtIndex

reNum:
    start = CInt((7 * Rnd(Timer)) + 1)

    If start > 7 Then start = 7

    If tCount < Val(txtLength) Then
'Taken out to make it random which column shrinks
'    If formatIndex <= 6 And tCount < Val(txtLength) Then
        'Get the current format
        myTag = txtTrace(tIndex).Tag

        'if the starting place in the format is already a space, try again
        If Mid(myTag, start, 1) = " " Then GoTo reNum

        'get thecharacters before the starting position, since replace doesnt return the full string
        pre = Left(myTag, start - 1)

        'replace the # in the format with a space
        newStr = Replace(myTag, "#", " ", start, 1)

        'Put them back together to form the new format
        txtTrace(tIndex).Tag = pre & newStr

        'Put up the next digit in the generated number
        txtFinalNumber = txtFinalNumber & Mid(setNumber, tCount + 1, 1) 'formatIndex + 1, 1)
        lblTracing = "Tracing: " & txtLength - (tCount + 1) & "..."
        'Taken out to make it random which column shrinks
        'lblNumber(formatIndex) = Mid(setNumber, formatIndex + 1, 1)
        formatIndex = formatIndex + 1
        tCount = tCount + 1
    'Taken out to make it random which column shrinks
    'ElseIf (Val(txtLength) > 7) And (tCount < Val(txtLength)) Then
    '    formatIndex = 0
    Else
        cmdHalt_Click
        tCount = 0
        lblComplete.Visible = True
        lblTracing.Visible = False
    End If

End Sub

Private Sub timScroll_Timer()

    Dim num As String
    Dim newText As String
    num = generateNumber(7, 9)

    For i = 0 To txtTrace.Count - 1
        newText = reFormat(num, txtTrace(i).Tag)
        txtTrace(i) = newText & vbCrLf & txtTrace(i)
        txtTrace(i) = Left(txtTrace(i), Len(txtTrace(i)) - 9)
    Next i

End Sub

'** I needed a format function that would allow me to remove 'digits' and replace them with spaces
Private Function reFormat(str As String, format As String) As String

    For i = 1 To Len(format)
        If Mid(format, i, 1) <> " " Then
            reFormat = reFormat & Mid(str, i, 1)
        Else
            reFormat = reFormat & " "
        End If
    Next

End Function

Private Sub txtInput_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyReturn:
            parseCommand (txtInput)
    End Select

End Sub

Private Sub txtInput_LostFocus()

    If frmAbout.Visible <> True Then txtInput.SetFocus

End Sub

Private Sub txtLength_Validate(Cancel As Boolean)

    If Val(txtLength) > 49 Then txtLength = 49

End Sub

'** Parses the command-line commands
Private Sub parseCommand(comStr As String)

    Dim space As Integer
    Dim command As String
    Dim argument As String
    
    space = InStr(1, comStr, " ")
    If space > 0 Then
        command = Mid(comStr, 1, space - 1)
        argument = Mid(comStr, space + 1, Len(comStr) - space)
    Else
        command = comStr
    End If
    
    If LCase(command) = "trace" Then
        If LCase(argument) <> "" Then
            txtLength = Mid(txtInput, 6)
            cmdReset_Click
            cmdTrace_Click
            lblError = ""
            txtInput = ""
        Else
            lblError = "You must enter a number between 1 and 49 with trace..."
            txtInput = ""
        End If
    ElseIf LCase(command) = "exit" Or command = "leave" Or command = "bye" Or command = "logout" Or command = "quit" Then
        If LCase(argument) = "" Then
            lblError = "Exiting, goodbye..."
            cmdExit_Click
        Else
            lblError = "Just type " & command & ", no arguments..."
            txtInput = ""
        End If
    ElseIf LCase(command) = "help" Or command = "?" Then
        If LCase(argument) = "" Then
            lblError = "help <subject> - trace | exit | debug | halt | reset | pause | about | who"
            txtInput = ""
        Else
            lblError = doHelp(argument)
            txtInput = ""
        End If
    ElseIf LCase(command) = "debug" Then
        If LCase(argument) = "" Then
            lblError = "Password please..."
            txtInput = ""
        ElseIf LCase(argument) = "agentsmith" Then
            lblError = "Welcome, Agent Smith..."
            cmdInfo_Click
            txtInput = ""
        ElseIf LCase(argument) = "matrix" Or argument = "neo" Then
            lblError = "Nice try..."
            txtInput = ""
        End If
    ElseIf LCase(command) = "halt" Or LCase(command) = "pause" Then
        If halted = True Then
            lblError = "Program alreadsy halted..."
            txtInput = ""
        Else
            lblError = "Halting..."
            cmdHalt_Click
            lblError = "Trace Halted..."
            txtInput = ""
        End If
    ElseIf LCase(command) = "reset" Then
        If halted = True Then
            lblError = "Resetting program..."
            cmdReset_Click
            lblError = "Program Reset..."
            txtInput = ""
        Else
            lblError = "Halting Program..."
            cmdHalt_Click
            lblError = "Trace Halted..."
            lblError = "Resetting program..."
            cmdReset_Click
            lblError = "Program Reset..."
            txtInput = ""
        End If
    ElseIf LCase(command) = "about" Then
        cmdAbout_Click
        txtInput = ""
    ElseIf LCase(command) = "who" Then
        lblError = "Xanth Nomeda : xnomeda@nomeda.com"
        txtInput = ""
    ElseIf LCase(command) = "command" Then
        lblError = "trace | exit | debug | halt | reset | pause | about | who"
        txtInput = ""
    Else
        lblError = "That was not a command!"
        txtInput = ""
    End If

End Sub

'** Gives out the help for each subject
Private Function doHelp(subject As String) As String

    Select Case subject
        Case "trace":
            doHelp = "trace <number of digits>: begin trace on <number of digits>"
        Case "exit":
            doHelp = "exit | bye | leave | logout | quit: quit the prorgam"
        Case "debug":
            doHelp = "debug <password>: gives debug info"
        Case "halt":
            doHelp = "halt | pause: stops the current trace"
        Case "pause":
            doHelp = "pause | halt: stops the current trace"
        Case "reset":
            doHelp = "reset: resets the trace"
        Case "about":
            doHelp = "about: shows the about dialog"
        Case "who":
            doHelp = "who: shows the author"
        Case Else:
            doHelp = "There is no help on that subject -" & subject & "-"
    End Select
End Function
