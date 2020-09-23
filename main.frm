VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E - MINDER"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox altime 
      Height          =   285
      Left            =   4680
      TabIndex        =   4
      ToolTipText     =   "Time in 24 hour format"
      Top             =   2040
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin VB.PictureBox traypict 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   530
      Left            =   120
      Picture         =   "main.frx":0442
      ScaleHeight     =   525
      ScaleMode       =   0  'User
      ScaleWidth      =   525
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   530
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5760
      Top             =   0
   End
   Begin VB.CommandButton cmdshowall 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Picture         =   "main.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Click to see all reminders"
      Top             =   4440
      Width           =   735
   End
   Begin VB.PictureBox showpict 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   530
      Left            =   4920
      ScaleHeight     =   525
      ScaleMode       =   0  'User
      ScaleWidth      =   525
      TabIndex        =   23
      Top             =   720
      Width           =   530
   End
   Begin VB.PictureBox miscpict 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   530
      Left            =   840
      Picture         =   "main.frx":0CC6
      ScaleHeight     =   525
      ScaleMode       =   0  'User
      ScaleWidth      =   525
      TabIndex        =   22
      Top             =   3360
      Visible         =   0   'False
      Width           =   530
   End
   Begin VB.PictureBox meetingpict 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   530
      Left            =   240
      Picture         =   "main.frx":1108
      ScaleHeight     =   525
      ScaleMode       =   0  'User
      ScaleWidth      =   525
      TabIndex        =   21
      Top             =   3360
      Visible         =   0   'False
      Width           =   530
   End
   Begin VB.PictureBox mailpict 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   530
      Left            =   840
      Picture         =   "main.frx":154A
      ScaleHeight     =   525
      ScaleMode       =   0  'User
      ScaleWidth      =   525
      TabIndex        =   20
      Top             =   3840
      Visible         =   0   'False
      Width           =   530
   End
   Begin VB.PictureBox callpict 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   530
      Left            =   240
      Picture         =   "main.frx":198C
      ScaleHeight     =   525
      ScaleMode       =   0  'User
      ScaleWidth      =   525
      TabIndex        =   19
      Top             =   3840
      Visible         =   0   'False
      Width           =   530
   End
   Begin VB.PictureBox birthdaypict 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   530
      Left            =   5640
      Picture         =   "main.frx":1DCE
      ScaleHeight     =   525
      ScaleMode       =   0  'User
      ScaleWidth      =   525
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   530
   End
   Begin VB.CommandButton cmdabout 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      Picture         =   "main.frx":2210
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Click to know about E - Minder"
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cmdhelp 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Picture         =   "main.frx":2652
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Click to take E - Minder Help"
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cmdexit 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      Picture         =   "main.frx":2A94
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Click to exit E - Minder"
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cmdclear 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      Picture         =   "main.frx":2ED6
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Click to clear all the fields"
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cmdsave 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Picture         =   "main.frx":3318
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Click to save entry to database"
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox comments 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   1560
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Comments"
      Top             =   3000
      Width           =   4455
   End
   Begin VB.TextBox multi 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   5
      ToolTipText     =   "Other information"
      Top             =   2520
      Width           =   4455
   End
   Begin VB.TextBox subnam 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   2
      ToolTipText     =   "Name or Subject"
      Top             =   1560
      Width           =   4455
   End
   Begin VB.ComboBox comtypes 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      ToolTipText     =   "Reminder type"
      Top             =   880
      Width           =   1935
   End
   Begin MSMask.MaskEdBox aldate 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      ToolTipText     =   "Date in MM/DD/YYYY format"
      Top             =   2040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label9 
      Caption         =   "HH:MM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   26
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "MM/DD/YYYY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   25
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "E - MINDER - An Easy Reminder Alerter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   840
      TabIndex        =   24
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label6 
      Caption         =   "Comments"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Relation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Select Reminder Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdabout_Click()
MsgBox "E - Minder Version 1.0" & vbCrLf & _
"Developed by PARMENDER DAHIYA" & vbCrLf & _
"TATA CONSULTANCY SERVICES, India" & vbCrLf & _
"ps_dahiya@yahoo.com", vbOKOnly + vbInformation, "E - Minder - About"
End Sub

Private Sub cmdclear_Click()

'clear all the fields and populate default values
subnam.BackColor = &H80000005
subnam.Text = ""
multi.BackColor = &H80000005
multi.Text = ""
aldate.BackColor = &H80000005
aldate.Mask = "##/##/####"
temp = "__/__/____"
aldate.Text = temp
altime.BackColor = &H80000005
temp = "__:__"
altime.Text = temp
comments.Text = ""
subnam.SetFocus
End Sub

Private Sub cmdexit_Click()
If (MsgBox("Are to sure to exit E - Minder ?", vbYesNo + vbInformation, "E - Minder - Exit")) = vbYes Then
End
End If
End Sub

Private Sub cmdhelp_Click()
MsgBox "E - Minder Version 1.0" & vbCrLf & _
"Developed by PARMENDER DAHIYA - ps_dahiya@yahoo.com" & vbCrLf & _
"" & vbCrLf & _
"Making of help is in progress. In the mean time to know the" & vbCrLf & _
"function of any of the button rest the mouse on the button," & vbCrLf & _
"a message will appear." & vbCrLf & _
"" & vbCrLf & _
"Thanks a lot for using my E - Minder." & vbCrLf & _
"For BUG reporting or comments please feel free to mail me.", vbOKOnly + vbInformation, "E - Minder - HELP"

End Sub

Private Sub cmdsave_Click()
Dim temp1 As String
subnam.BackColor = &H80000005
multi.BackColor = &H80000005
aldate.BackColor = &H80000005
altime.BackColor = &H80000005

    
'Name or subject validation
If subnam.Text = "" Then
    If rmcode = "BR" Or rmcode = "ML" Or rmcode = "CL" Then
    MsgBox "You have not entered the name.", vbOKOnly + vbCritical, "E - Minder - Error"
    Else
    MsgBox "You have not entered the subject.", vbOKOnly + vbCritical, "E - Minder - Error"
    End If
    subnam.BackColor = &HFFFF&
    subnam.SetFocus
    Exit Sub
End If

'Date validations for all types of reminders
If aldate.Text = "" Then ' if date not entered
    MsgBox "You have not entered the date.", vbOKOnly + vbCritical, "E - MINDER - Error"
    aldate.BackColor = &HFFFF&
    aldate.SetFocus
    Exit Sub
Else
    dateerror = ""
    checkdate (aldate.Text)
    If dateerror <> "" Then 'if there is error in date the dateerror will not be blank
    MsgBox dateerror, vbOKOnly + vbCritical, "E - MINDER - Error"
    aldate.BackColor = &HFFFF&
    aldate.SetFocus
    Exit Sub
    End If
End If

'Time validations for all types of reminders
If altime.Text = "" Then ' if  time not entered
    MsgBox "You have not entered the time.", vbOKOnly + vbCritical, "E - MINDER - Error"
    altime.BackColor = &HFFFF&
    altime.SetFocus
    Exit Sub
Else
    timeerror = ""
    checktime (altime.Text)
    If timeerror <> "" Then 'if there is error in time the timeerror will not be blank
    MsgBox timeerror, vbOKOnly + vbCritical, "E - MINDER - Error"
    altime.BackColor = &HFFFF&
    altime.SetFocus
    Exit Sub
    End If
End If
    
'Now validations according to the reminder types

Select Case rmcode

Case "BR":
        If multi.Text = "" Then
            MsgBox "You have not entered your relation with - " & subnam, vbOKOnly + vbCritical, "E - MINDER - Error"
            multi.BackColor = &HFFFF&
            multi.SetFocus
            Exit Sub
        End If
Case "CL":
         If multi.Text = "" Then   'check if number field is blank
            MsgBox "You have not entered number to call to  - " & subnam, vbOKOnly + vbCritical, "E - MINDER - Error"
            multi.BackColor = &HFFFF&
            multi.SetFocus
            Exit Sub
        Else
            ' check if it is non numeric, this can be done by the IsNumeric Function,
            ' but that will allow "." to be present
            For i = 1 To Len(multi.Text)
            temp1 = Mid(multi.Text, i, 1)
            If temp1 <> "0" And temp1 <> "1" And temp1 <> "2" And temp1 <> "3" _
               And temp1 <> "4" And temp1 <> "5" And temp1 <> "6" And temp1 <> "7" _
               And temp1 <> "8" And temp1 <> "9" Then
                MsgBox "Only integer values allowed in the number.", vbOKOnly + vbCritical, "E - MINDER - Error"
                multi.BackColor = &HFFFF&
                multi.SetFocus
                Exit Sub
            End If
            Next i
        End If

Case "ML":
        'Mail Validations only for "Mail Reminder"
        If multi.Text = "" Then ' if blank
            MsgBox "You have not entered E-mail address.", vbOKOnly + vbCritical, "E - MINDER - Error"
            multi.BackColor = &HFFFF&
            multi.SetFocus
            Exit Sub
        Else
            mailerror = checkMailVal(multi.Text) ' call mail validation function
            If mailerror <> "" Then ' if invalid mail entered
                MsgBox mailerror, vbOKOnly + vbCritical, "E - MINDER - Error"
                multi.BackColor = &HFFFF&
                multi.SetFocus
                Exit Sub
            End If
        End If
        
Case "MS": ' In this case this field is not visible
Case "MT":
            ' check if the location field is blank
        If multi.Text = "" Then
            MsgBox "You have not enterd the location of your meeting.", vbOKOnly + vbCritical, "E - MINDER - Error"
            multi.BackColor = &HFFFF&
            multi.SetFocus
            Exit Sub
        End If
End Select

'save to database
If (MsgBox("Save this reminder entry ?", vbYesNo + vbInformation, "E - Minder - Confirm")) = vbYes Then
savealarm (rmcode)
Call cmdclear_Click
End If

End Sub

Private Sub cmdshowall_Click()
ShowAll.Show
Me.Hide
End Sub

Private Sub comtypes_click()

'Set the  pictures and headings according to the reminder type
Label5.Visible = True
multi.Visible = True

Select Case comtypes.Text
Case "Birthday Reminder":
     showpict.Picture = birthdaypict.Picture
     Label5.Caption = "Relation"
     Label2.Caption = "Name"
     rmcode = "BR"
Case "Call Reminder":
     showpict.Picture = callpict.Picture
     Label5.Caption = "Number"
     Label2.Caption = "Name"
     rmcode = "CL"
Case "Mail Reminder":
     showpict.Picture = mailpict.Picture
     Label5.Caption = "E-mail"
     Label2.Caption = "Name"
     rmcode = "ML"
Case "Misc. Reminder":
     showpict.Picture = miscpict.Picture
     Label5.Visible = False
     multi.Visible = False
     Label2.Caption = "Subject"
     rmcode = "MS"
Case "Meeting Reminder":
     showpict.Picture = meetingpict.Picture
     Label5.Caption = "Location"
     Label2.Caption = "Subject"
     rmcode = "MT"

End Select
End Sub

Private Sub Form_Load()
'Give welcome message
MsgBox "Welcome to E - Minder, Version 1.0" & vbCrLf & _
"Developed by PARMENDER DAHIYA" & vbCrLf & _
"TATA CONSULTANCY SERVICES, India" & vbCrLf & _
"ps_dahiya@yahoo.com", vbOKOnly + vbInformation, "E - Minder - Welcome"

' If already running then give message and end
If App.PrevInstance = True Then
MsgBox "E - Minder is already running.", vbOKOnly + vbCritical, "E - Minder - Already Running."
End
End If

Call initdb ' Initialize database
Timer1.Enabled = True
' Add items to the combobox and show the "Birthday Reminder" as default
comtypes.AddItem "Birthday Reminder"
comtypes.AddItem "Call Reminder"
comtypes.AddItem "Mail Reminder"
comtypes.AddItem "Misc. Reminder"
comtypes.AddItem "Meeting Reminder"
comtypes.Text = "Birthday Reminder"
rmcode = "BR"
' show the picture for "Birthday Reminder"
showpict.Picture = birthdaypict.Picture
App.TaskVisible = False

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
Call systemtray ' Put in the system tray
main.Hide
End Sub

Private Sub Timer1_Timer()
Dim tmptype As String
Dim tmpsubnam As String
Dim tmpcomments As String
Dim tmpmulti As String
Dim tmpday As Integer
Dim tmpmonth As Integer
Dim tmpyear As Integer
Dim tmphour As Integer
Dim tmpminute As Integer
Dim tmpdone As Boolean
Dim AMorPM As String
Dim crnthour As Integer
Dim crntminute As Integer
Dim temptime As String
Dim i As Integer

'This timer checks for any remineder is there to show or not.
'It checks for all the reminders in the "Pending" state only


If rs.RecordCount = o Then Exit Sub

rs.MoveFirst

For i = 1 To rs.RecordCount ' Read all the records from database
tmptype = rs("type")
rmcodepop = tmptype
tmpsubnam = rs("sub_or_name")
tmpcomments = rs("com_or_desc")
tmpmulti = rs("multi_info")
tmpday = rs("alm_day")
tmpmonth = rs("alm_month")
tmpyear = rs("alm_year")
tmphour = rs("alm_hour")
tmpminute = rs("alm_minute")
tmpdone = rs("alm_done")

If tmpdone = False Then ' if reminder is in the "Pending" state
    
'get current system time and break in hour and minute as 24 hour format

temptime = Time
'this is done because if time can be of length 11 or 10
' e.g. "11:35:25 AM" and  "1:20:50 PM"
'so in both the cases we have to get the hour and minute differently
If Len(temptime) = 11 Then
    crnthour = Mid(temptime, 1, 2)
    crntminute = Mid(temptime, 4, 2)
    AMorPM = Mid(temptime, 10, 2)
Else
    crnthour = Mid(temptime, 1, 1)
    crntminute = Mid(temptime, 3, 2)
    AMorPM = Mid(temptime, 9, 2)
End If

'Convert to 24 hour format
If AMorPM = "PM" And crnthour <> 12 Then crnthour = crnthour + 12
If AMorPM = "AM" And crnthour = 12 Then crnthour = 0

' Check if this record is due for popup or not i.e.
' Its time and date is less than the current time and date or not
If (Year(Now) > tmpyear) Or ((Year(Now) = tmpyear) And (Month(Now) > tmpmonth)) Or _
   ((Year(Now) = tmpyear) And (Month(Now) = tmpmonth) And (Day(Now) > tmpday)) Or _
   ((Year(Now) = tmpyear) And (Month(Now) = tmpmonth) And (Day(Now) = tmpday) And (crnthour > tmphour)) Or _
   ((Year(Now) = tmpyear) And (Month(Now) = tmpmonth) And (Day(Now) = tmpday) And (crnthour = tmphour) And (crntminute >= tmpminute)) Then
       'Timer1.Enabled = False
       fillpopup ' fill the popup form
       PopUP.Show ' show popup form
       
       Exit Sub
    End If
End If
rs.MoveNext ' read next record
Next i


End Sub

Private Function fillpopup()
Dim temp As String
Dim temp1 As String

'this  function fills the popup form with the proper entries from the data base
PopUP.Label5.Visible = True
PopUP.multi.Visible = True

temp = rs("alm_month")
temp1 = rs("alm_day")
If Len(temp) = 1 Then temp = "0" & temp
If Len(temp1) = 1 Then temp1 = "0" & temp1
PopUP.aldate.Caption = temp & "/" & temp1 & "/" & rs("alm_year") ' date

temp = rs("alm_minute")
temp1 = rs("alm_hour")
If Len(temp) = 1 Then temp = "0" & temp
If Len(temp1) = 1 Then temp1 = "0" & temp1
PopUP.altime.Caption = temp1 & ":" & temp ' time

PopUP.subnam.Caption = rs("sub_or_name") ' subject or name
PopUP.multi.Caption = rs("multi_info")   ' other information
PopUP.comments.Caption = rs("com_or_desc") ' comments

' fill the heading as per the reminder type
Select Case rmcodepop
Case "BR"
    PopUP.type.Caption = "Birthday Reminder"
    PopUP.Label2.Caption = "Name"
    PopUP.Label5.Caption = "Relation"
    PopUP.showpict.Picture = birthdaypict.Picture
Case "CL"
    PopUP.type.Caption = "Call Reminder"
    PopUP.Label2.Caption = "Name"
    PopUP.Label5.Caption = "Number"
    PopUP.showpict.Picture = callpict.Picture
Case "ML"
    PopUP.type.Caption = " Mail Reminder"
    PopUP.Label2.Caption = "Name"
    PopUP.Label5.Caption = "E-mail"
    PopUP.showpict.Picture = mailpict.Picture
Case "MS"
    PopUP.type.Caption = "Misc. Reminder"
    PopUP.Label2.Caption = "Subject"
    PopUP.Label5.Visible = False
    PopUP.multi.Visible = False
    PopUP.showpict.Picture = miscpict.Picture
Case "MT"
    PopUP.type.Caption = "Meeting Reminder"
    PopUP.Label2.Caption = "Subject"
    PopUP.Label5.Caption = "Location"
    PopUP.showpict.Picture = meetingpict.Picture
End Select

End Function
Private Sub traypict_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Hex(x) = "1E3C" Then ' if right click
If (MsgBox("Are to sure to exit E - Minder ?", vbYesNo + vbInformation, "E - Minder - Exit")) = vbYes Then End
End If

If Hex(x) = "1E0F" Then ' if left click
main.Show
End If

End Sub
