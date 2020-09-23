VERSION 5.00
Begin VB.Form fNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Event"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   22
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Step"
      Height          =   375
      Left            =   2280
      TabIndex        =   21
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "Previous Step"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame frEvent 
      Caption         =   "1. Event Information"
      Height          =   1935
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   3615
      Begin VB.TextBox txtDes 
         Height          =   975
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtTitle 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Text            =   "Event Title"
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame frType 
      Caption         =   "2. Type of Time Table"
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3615
      Begin VB.OptionButton optByDOW 
         Caption         =   "Set By Day of Week"
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optByDate 
         Caption         =   "Set By Date"
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame frAlert 
      Caption         =   "5. Event Warning"
      Height          =   1935
      Left            =   0
      TabIndex        =   15
      Top             =   600
      Width           =   3615
      Begin VB.TextBox txtAlert 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         MaxLength       =   2
         TabIndex        =   17
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.CheckBox chkAlert 
         Caption         =   "Warn you of event"
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblWarnAid 
         Caption         =   "days before day of event"
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.Frame frDate 
      Caption         =   "4. Date of Event"
      Height          =   1935
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   3615
      Begin VB.TextBox txtWeek 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         MaxLength       =   1
         TabIndex        =   24
         Text            =   "1"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ComboBox cboDOW 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "1"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtYear 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   12
         Text            =   "2000"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblLI 
         Caption         =   "st"
         Height          =   255
         Left            =   1200
         TabIndex        =   26
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "On the"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   240
         X2              =   3360
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin VB.Frame frConsistancy 
      Caption         =   "3. Consistancy of Event"
      Height          =   1935
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   3615
      Begin VB.OptionButton optOnce 
         Caption         =   "Once"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optWeekly 
         Caption         =   "Weekly"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optMonthly 
         Caption         =   "Monthly"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optAnnually 
         Caption         =   "Annually"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Caption         =   "Enter the Title and Description of the event. Then, click 'Next Step' to continue."
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "fNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim new_event As Evt
Dim Step As Integer

Private Sub chkAlert_Click()
    If chkAlert.Value = 0 Then txtAlert.Enabled = False Else txtAlert.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLast_Click()
    Step = Step - 1
    Call PrepareStep
End Sub

Private Sub cmdNext_Click()
    Step = Step + 1
    Call PrepareStep
End Sub

Private Sub Form_Load()
    'setup month/DOW combo box
    With cboMonth
        .AddItem "January"
        .AddItem "February"
        .AddItem "March"
        .AddItem "April"
        .AddItem "May"
        .AddItem "June"
        .AddItem "July"
        .AddItem "August"
        .AddItem "September"
        .AddItem "October"
        .AddItem "November"
        .AddItem "December"
        .ListIndex = 0
    End With
    With cboDOW
        .AddItem "Sunday"
        .AddItem "Monday"
        .AddItem "Tuesday"
        .AddItem "Wednesday"
        .AddItem "Thursday"
        .AddItem "Friday"
        .AddItem "Saturday"
        .ListIndex = 0
    End With
    'initilize vars
    Step = 1
    With new_event
        .alert = True
        .annual = False
        .by_day_of_week = False
        .day = 1
        .day_of_week = 1
        .days_ahead_to_alert = 1
        .description = "Description of Event"
        .Month = 1
        .monthly = False
        .once = True
        .Title = "Event Title"
        .Week = 1
        .weekly = False
        .Year = 2000
    End With
    chkAlert.Value = 1
    optOnce.Value = True
    optByDate.Value = True
    cmdLast.Enabled = False
End Sub

Public Sub PrepareStep()
Dim OS As Integer
Select Case Step
  Case 1
    frEvent.ZOrder
    cmdLast.Enabled = False
    lblMsg.Caption = "Enter the Title and Description of the event. Then, click 'Next Step' to continue."
  Case 2
    If txtTitle.Text = "" Or txtDes.Text = "" Then
      MsgBox "The title box AND the description box need to be filled out. Please go back and fill them."
      Step = 1
      Exit Sub
    End If
    frType.ZOrder
    cmdLast.Enabled = True
    lblMsg.Caption = "'Set By Date' Example: March 9, 1984. 'Set By Day of Week' Example: 1st Tuesday of July"
  Case 3
    frConsistancy.ZOrder
    lblMsg.Caption = "Select how often you would like the event to be triggered."
    If optByDate.Value Then optWeekly.Enabled = False Else optWeekly.Enabled = True
    If optWeekly.Enabled = False And optWeekly.Value = True Then
      optWeekly.Value = False
      optOnce.Value = True
    End If
  Case 4
    frDate.ZOrder
    lblMsg.Caption = "Enter the Month, Day, Year, Day of Week, etc. that applies."
    cmdNext.Caption = "Next Step"
    If Not optOnce.Value Then txtYear.Enabled = False Else txtYear.Enabled = True
    If optByDOW.Value Or optWeekly.Value Then txtDay.Enabled = False Else txtDay.Enabled = True
    If optMonthly.Value Or optWeekly.Value Then cboMonth.Enabled = False Else cboMonth.Enabled = True
    If optByDate.Value Or optWeekly.Value Then txtWeek.Enabled = False Else txtWeek.Enabled = True
    If optByDate.Value Then cboDOW.Enabled = False Else cboDOW.Enabled = True
  Case 5
    If (txtDay.Text = "" And txtDay.Enabled = True) Or (txtYear.Text = "" And txtYear.Enabled = True) Or (txtWeek.Text = "" And txtWeek.Enabled = True) Then
      MsgBox "One or more boxes have not been filled out. Please go back and fill them."
      Step = 4
      Exit Sub
    End If
    frAlert.ZOrder
    lblMsg.Caption = "Decide if PC Calendar will alert you to this event and how many days in advance it will alert you."
    cmdNext.Caption = "Finish"
  Case 6 'Finish
    If txtAlert.Text = "" And txtAlert.Enabled = True Then
      MsgBox "The box has not been filled out. Please go back and fill them."
      Step = 5
      Exit Sub
    End If
    'Write event
    With new_event
      If chkAlert.Value = 0 Then .alert = False Else .alert = True
      .annual = optAnnually.Value
      .by_day_of_week = optByDOW.Value
      .day = txtDay.Text
      .day_of_week = cboDOW.ListIndex + 1
      .days_ahead_to_alert = txtAlert.Text
      .description = txtDes.Text
      .Month = cboMonth.ListIndex + 1
      .monthly = optMonthly.Value
      .once = optOnce.Value
      .Title = txtTitle.Text
      .Week = txtWeek.Text
      .weekly = optWeekly.Value
      .Year = txtYear.Text
    End With
    If Not Inserting Then
        AddNewEvent new_event
    Else
        InsertEvent InsertSlot, new_event
        Inserting = False
    End If
    If fZEdit.Visible Then fZEdit.UpdateList
    Unload Me
    frm.ShowNewMonth
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Inserting = False
End Sub

Private Sub txtAlert_Change()
    If txtAlert.Text = "" Then Exit Sub
    On Error GoTo Error
    If Int(txtAlert.Text) >= 0 And Int(txtAlert.Text) <= 100 Then GoTo nxt Else GoTo Error
nxt:
Exit Sub
Error:
    MsgBox "You must enter a number between 0 and 100"
    txtAlert.Text = "1"
End Sub

Private Sub txtDay_Change()
    If txtDay.Text = "" Then Exit Sub
    On Error GoTo Error
    If Not cboMonth.Enabled Then
      If Int(txtDay.Text) >= 1 And Int(txtDay.Text) <= 28 Then GoTo nxt Else GoTo Error
    Else
      Select Case cboMonth.ListIndex
        Case 0, 2, 4, 6, 7, 9, 11
          If Int(txtDay.Text) >= 1 And Int(txtDay.Text) <= 31 Then GoTo nxt Else GoTo Error
        Case 3, 5, 8, 10
          If Int(txtDay.Text) >= 1 And Int(txtDay.Text) <= 30 Then GoTo nxt Else GoTo Error
        Case 1
          If txtYear.Enabled Then
            If Int((Int(txtYear.Text) - 2000) / 4) = (Int(txtYear.Text) - 2000) / 4 Then
              If Int(txtDay.Text) >= 1 And Int(txtDay.Text) <= 29 Then GoTo nxt Else GoTo Error
            Else
              If Int(txtDay.Text) >= 1 And Int(txtDay.Text) <= 28 Then GoTo nxt Else GoTo Error
            End If
          Else
            If Int(txtDay.Text) >= 1 And Int(txtDay.Text) <= 28 Then GoTo nxt Else GoTo Error
          End If
      End Select
    End If
nxt:
    
Exit Sub
Error:
    MsgBox "Either you did not type in a valid number or you entered a number that does not exist in the given month."
    txtDay.Text = "1"
End Sub

Private Sub txtWeek_Change()
    If txtWeek.Text = "" Then Exit Sub
    On Error GoTo Error
    If Int(txtWeek.Text) > 0 And Int(txtWeek.Text) < 6 Then GoTo nxt Else GoTo Error
nxt:
    Select Case Int(Right(txtWeek.Text, 1))
      Case 1
        lblLI.Caption = "st"
      Case 2
        lblLI.Caption = "nd"
      Case 3
        lblLI.Caption = "rd"
      Case Else
        lblLI.Caption = "th"
    End Select
Exit Sub
Error:
    MsgBox "You must enter a number between 1 and 5"
    txtWeek.Text = "1"
End Sub

Private Sub txtYear_Change()
    If txtYear.Text = "" Then Exit Sub
    On Error GoTo Error
    If Int(txtYear.Text) > 0 And Int(txtYear.Text) < 10000 Then GoTo nxt Else GoTo Error
nxt:

Exit Sub
Error:
    MsgBox "You must enter a number greater than 0."
    txtYear.Text = "2000"
End Sub
