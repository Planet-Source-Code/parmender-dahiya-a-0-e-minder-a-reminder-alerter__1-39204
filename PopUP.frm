VERSION 5.00
Begin VB.Form PopUP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E - Minder -  Alert"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox showpict 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   530
      Left            =   360
      ScaleHeight     =   525
      ScaleMode       =   0  'User
      ScaleWidth      =   525
      TabIndex        =   16
      Top             =   120
      Width           =   530
   End
   Begin VB.ComboBox comremind 
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
      Left            =   3360
      TabIndex        =   13
      ToolTipText     =   "Select time to remind after"
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton cmdremind 
      Caption         =   "Remind Me"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      Picture         =   "PopUP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Click to remind again"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmddone 
      Caption         =   "           OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Picture         =   "PopUP.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click to mark as done."
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "after"
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
      Left            =   2760
      TabIndex        =   15
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Minutes"
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
      Left            =   4080
      TabIndex        =   14
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label comments 
      Caption         =   "comments"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1320
      TabIndex        =   11
      Top             =   2880
      Width           =   3735
   End
   Begin VB.Label multi 
      Caption         =   "multi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   10
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label altime 
      Caption         =   "altime"
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
      Left            =   3840
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label aldate 
      Caption         =   "Aldate"
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
      Left            =   1320
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label subnam 
      Caption         =   "Subnam"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   7
      Top             =   840
      Width           =   3735
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2160
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   1680
      Width           =   615
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   615
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Label type 
      Caption         =   "Birthday Reminder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "PopUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmddone_Click()
'Mark this entry as done
rs.Edit
rs("alm_done") = True
rs.Update
Unload Me
End Sub

Private Sub cmdremind_Click()

If comremind.Text = "" Then
MsgBox "Select a value from the list ? ", vbOKOnly + vbCritical, "E - Minder - Error"
Exit Sub
End If

If remind = False Then  'call function "remind"
rs.Edit
rs("alm_done") = True
Else
rs("alm_done") = False
End If

'update the database
rs.Update
Unload Me
End Sub

Private Sub comremind_click()
'Enable the "Remind" button only when user selects an entry from combo box
If comremind.Text <> "" Then
cmdremind.Enabled = True
Else
cmdremind.Enabled = False
End If

End Sub

Private Sub Form_Load()
'Add items in the combo box
comremind.AddItem 10
comremind.AddItem 20
comremind.AddItem 30
comremind.AddItem 40
comremind.AddItem 50
comremind.AddItem 60
cmdremind.Enabled = False
main.Timer1.Enabled = False
SetWindowPos PopUP.hWnd, -1, 0, 0, 0, 0, &H1 Or &H2
End Sub

Private Function remind() As Boolean

Dim rmyr As Integer
Dim rmmon As Integer
Dim rmday As Integer
Dim rmmin As Integer
Dim rmhour As Integer
Dim rmAMorPM As String
Dim increment As Integer
Dim temp As Integer
Dim temptime As String

If comremind.Text = "" Then
remind = False
Exit Function
End If

'get current system time and break in hour and minute as 24 hour format
temptime = Time
'this is done because if time can be of length 11 or 10
' e.g. "11:35:25 AM" and  "1:20:50 PM"
'so in both the cases we have to get the hour and minute differently
If Len(temptime) = 11 Then
    rmhour = Mid(temptime, 1, 2)
    rmmin = Mid(temptime, 4, 2)
    rmAMorPM = Mid(temptime, 10, 2)
Else
    rmhour = Mid(temptime, 1, 1)
    rmmin = Mid(temptime, 3, 2)
    rmAMorPM = Mid(temptime, 9, 2)
End If

'Convert to 24 hour format
If rmAMorPM = "PM" And rmhour <> 12 Then rmhour = rmhour + 12
If rmAMorPM = "AM" And rmhour = 12 Then rmhour = 0

'get system date
rmyr = Year(Now)
rmmon = Month(Now)
rmday = Day(Now)
increment = comremind.Text

temp = rmmin
rmmin = rmmin + increment

If rmmin > 59 Then ' if minutes greater than 59 then increment hour
    rmhour = rmhour + 1
    rmmin = (increment - (60 - temp)) ' increment minutes accordingly
    If rmhour > 23 Then ' if hour more than 23
        rmday = rmday + 1   ' increment day and hour will be zero
        rmhour = 0
        Select Case rmmon   ' check that after increment in the day it is proper or not
        Case 1 Or 3 Or 5 Or 7 Or 8 Or 10:
            If rmday > 31 Then
                rmmon = rmmon + 1
                rmday = 1
            End If
        Case 4 Or 6 Or 9 Or 11:
            If rmday > 30 Then
                rmmon = rmmon + 1
                rmday = 1
            End If
            
        Case 2:
            If (rmyr Mod 400) = 0 Then ' if divisible by 400 then max days = 29
                If rmday > 29 Then
                    rmmon = rmmon + 1
                    rmday = 1
                End If
            Else
                If (rmyr Mod 100) = 0 Then 'if divisible by 100 and not by 400 then = 28
                    If rmday > 28 Then
                        rmmon = rmmon + 1
                        rmday = 1
                    End If
                Else
                    If (rmyr Mod 4) = 0 Then ' if by 4 but not by 100 and 400 then = 29
                        If rmday > 29 Then
                            rmmon = rmmon + 1
                            rmday = 1
                        End If
                    Else
                        If rmday > 28 Then ' if not by 4 and 100 and 400 then = 28
                            rmmon = rmmon + 1
                            rmday = 1
                        End If
                    End If
                End If
            End If
            
        Case 12: ' if december month then increment year
            If rmday > 31 Then
                rmyr = rmyr + 1
                rmmon = 1
                rmday = 1
            End If
        End Select
    End If
End If
' populate database fields
rs.Edit
rs("alm_minute") = rmmin
rs("alm_hour") = rmhour
rs("alm_day") = rmday
rs("alm_month") = rmmon
rs("alm_year") = rmyr
remind = True
End Function

Private Sub Form_Unload(Cancel As Integer)
main.Timer1.Enabled = True
SetWindowPos PopUP.hWnd, -2, 0, 0, 0, 0, &H1 Or &H2
End Sub
