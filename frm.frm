VERSION 5.00
Begin VB.Form frm 
   Caption         =   "PC Calendar"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10125
   Icon            =   "frm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLast 
      Caption         =   "&Last Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "N&ext Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label day 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label DOW 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sunday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label DOW 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label DOW 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tuesday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label DOW 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wednesday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label DOW 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thursday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2520
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label DOW 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Friday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   3120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label DOW 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saturday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   3720
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label month_year 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Click to Change"
      Top             =   0
      Width           =   6135
   End
   Begin VB.Menu mnuEvent 
      Caption         =   "&Event"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuManage 
         Caption         =   "&Manage"
      End
      Begin VB.Menu mnuBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuColors 
         Caption         =   "&Colors"
      End
      Begin VB.Menu mnuStartup 
         Caption         =   "&Startup"
      End
      Begin VB.Menu mnuBreak3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnableBT 
         Caption         =   "Backtracking Disabled"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuBackTrack 
      Caption         =   "&Back Track"
      Enabled         =   0   'False
      Begin VB.Menu mnuNoSave 
         Caption         =   "Don't Save Anything"
      End
      Begin VB.Menu mnuNoSaveCfg 
         Caption         =   "Don't Save Options"
      End
      Begin VB.Menu mnuNoSaveEvt 
         Caption         =   "Don't Save Events"
      End
      Begin VB.Menu mnuBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResetAll 
         Caption         =   "Reset All"
      End
      Begin VB.Menu mnuResetCfg 
         Caption         =   "Reset Options"
      End
      Begin VB.Menu mnuResetEvt 
         Caption         =   "Reset Events"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuTips 
         Caption         =   "&Display Tips"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub PlaceBoxes(ColWidth As Integer, RowWidth As Integer)
    If Me.day.UBound = 0 Then Exit Sub
    Dim DOW As Integer
    Dim Week As Integer
    Dim box As Integer
    box = 0
    For Week = 1 To 6
        For DOW = 1 To 7
            day(box).Visible = True
            day(box).Move LEFT_MARGIN + ((DOW - 1) * ColWidth), TOP_MARGIN + ((Week - 1) * RowWidth), ColWidth - SPACING, RowWidth - SPACING
            box = box + 1
        Next DOW
    Next Week
End Sub

Public Sub SpawnBoxes()
    Dim A As Integer
    
    If day.UBound > 0 Then Exit Sub
    On Error Resume Next
    
    For A = 1 To 41
        Load day(A)
    Next A
End Sub

Public Function GetColWidth() As Integer
    If frm.Width <= LEFT_MARGIN + RIGHT_MARGIN Then
        GetColWidth = 10
    Else
        GetColWidth = (Me.Width - LEFT_MARGIN - RIGHT_MARGIN) / 7
    End If
End Function

Public Function GetRowWidth()
    If frm.Height <= TOP_MARGIN + BOTTOM_MARGIN Then
        GetRowWidth = 10
    Else
        GetRowWidth = (Me.Height - TOP_MARGIN - BOTTOM_MARGIN) / 6
    End If
End Function

Public Sub PlaceHeader()
    Dim A As Integer
    Dim CW As Integer
    If LEFT_MARGIN + RIGHT_MARGIN + Me.cmdLast.Width + Me.cmdNext.Width >= Me.Width Then Exit Sub
    
    Me.cmdLast.Move LEFT_MARGIN, 0
    Me.month_year.Move Me.cmdLast.Left + Me.cmdLast.Width, 0, Me.Width - LEFT_MARGIN - RIGHT_MARGIN - Me.cmdLast.Width - Me.cmdNext.Width
    Me.cmdNext.Move Me.month_year.Left + Me.month_year.Width, 0
    
    CW = GetColWidth
    
    For A = 0 To 6
        Me.DOW(A).Move LEFT_MARGIN + (A * (CW + SPACING)), Me.month_year.Height, CW
    Next A
End Sub

Private Sub cmdLast_Click()
    Dim New_Month As Integer
    Dim New_Year As Long
    
    New_Month = MonthShown - 1
    Select Case New_Month
        Case 0
            New_Month = 12
            New_Year = YearShown - 1
            If New_Year = -1 Then
                New_Month = 1
                New_Year = 0
            End If
        Case 1 To 12
            New_Year = YearShown
        Case Else
            MsgBox "Your system is trying to load a month that does not exist. Something is wrong.", vbCritical, "Error"
            Exit Sub
    End Select
    
    If MonthShown <> New_Month Or YearShown <> New_Year Then
        MonthShown = New_Month
        YearShown = New_Year
        ShowNewMonth
    End If
End Sub

Private Sub cmdNext_Click()
    Dim New_Month As Integer
    Dim New_Year As Double
    
    New_Month = MonthShown + 1
    Select Case New_Month
        Case 1 To 12
            New_Year = YearShown
        Case 13
            New_Month = 1
            New_Year = YearShown + 1
            If New_Year = 10000 Then
                New_Year = 9999
                New_Month = 12
            End If
        Case Else
            MsgBox "Your system is trying to load a month that does not exist. Something is wrong."
        Exit Sub
    End Select
    
    If MonthShown <> New_Month Or YearShown <> New_Year Then
        MonthShown = New_Month
        YearShown = New_Year
        ShowNewMonth
    End If
End Sub

Private Sub day_Click(Index As Integer)
    If day(Index).Tag > 0 Then
        Load fDay
        fDay.Start MonthShown & "/" & day(Index).Tag & "/" & YearShown
    End If
End Sub

Private Sub Form_Load()
    SpawnBoxes
    UpdateColor
    ShowNewMonth
    frm.WindowState = 2
    frm.Visible = True
    
    If ShowTips Then frmTip.Show
End Sub

Private Sub Form_Resize()
    PlaceHeader
    PlaceBoxes GetColWidth, GetRowWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "PC Calendar", "Options", "LA", Now
    
    If AnyChanges Then
        If CfgChanges Then SaveAll FA_CFG
        If EvtChanges Then SaveAll FA_EVENTS
    End If
    End
End Sub

Public Sub UpdateColor()
Attribute UpdateColor.VB_Description = "Updates color settings after they have been changed"
    Dim A As Integer
    
    If MonthOnDef Then
        Me.month_year.ForeColor = vbButtonText
        Me.month_year.BackColor = vbButtonFace
    Else
        Me.month_year.ForeColor = MonthFont
        Me.month_year.BackColor = MonthBg
    End If
    
    For A = 0 To 6
        If DOWOnDef Then
            Me.DOW(A).ForeColor = vbButtonText
            Me.DOW(A).BackColor = vbButtonFace
        Else
            Me.DOW(A).ForeColor = DOWFont
            Me.DOW(A).BackColor = DOWBg
        End If
    Next A
End Sub
Public Sub SetupColorCode()
Dim StartBox As Integer
Dim NO As String
StartBox = Weekday(MonthShown & "/1/" & YearShown) - 1
Dim n As Integer

'plot days of the month
For n = 0 To StartBox - 1
    Me.day(n).BackColor = IIf(BlankOnDef, vbInactiveTitleBar, BlankBg)
    Me.day(n).Tag = 0
    Me.day(n).Caption = ""
Next n
For n = StartBox To days_in_month(MonthShown) + StartBox - 1
    Me.day(n).BackColor = IIf(RegularOnDef, vbWhite, RegularBg)
    Me.day(n).ForeColor = IIf(RegularOnDef, vbBlack, RegularFont)
    Me.day(n).Tag = n - StartBox + 1
    Me.day(n).Caption = Me.day(n).Tag
Next n
For n = days_in_month(MonthShown) + StartBox To 41
    Me.day(n).BackColor = IIf(BlankOnDef, vbInactiveTitleBar, BlankBg)
    Me.day(n).Tag = 0
    Me.day(n).Caption = ""
Next n
'write in other data
If chkHol Then
    If Me.day(0).Tag <> 0 Then
        Me.day(0).BackColor = IIf(WeekEndsOnDef, vbMagenta, WeekEndsBg)
        Me.day(0).ForeColor = IIf(WeekEndsOnDef, vbBlack, WeekEndsFont)
    End If
    Me.day(7).BackColor = IIf(WeekEndsOnDef, vbMagenta, WeekEndsBg)
    Me.day(7).ForeColor = IIf(WeekEndsOnDef, vbBlack, WeekEndsFont)
    Me.day(14).BackColor = IIf(WeekEndsOnDef, vbMagenta, WeekEndsBg)
    Me.day(14).ForeColor = IIf(WeekEndsOnDef, vbBlack, WeekEndsFont)
    Me.day(21).BackColor = IIf(WeekEndsOnDef, vbMagenta, WeekEndsBg)
    Me.day(21).ForeColor = IIf(WeekEndsOnDef, vbBlack, WeekEndsFont)
    If Me.day(28).Tag <> 0 Then
        Me.day(28).BackColor = IIf(WeekEndsOnDef, vbMagenta, WeekEndsBg)
        Me.day(28).ForeColor = IIf(WeekEndsOnDef, vbBlack, WeekEndsFont)
    End If
    If Me.day(35).Tag <> 0 Then
        Me.day(35).BackColor = IIf(WeekEndsOnDef, vbMagenta, WeekEndsBg)
        Me.day(35).ForeColor = IIf(WeekEndsOnDef, vbBlack, WeekEndsFont)
    End If
    If Me.day(6).Tag <> 0 Then
        Me.day(6).BackColor = IIf(WeekEndsOnDef, vbMagenta, WeekEndsBg)
        Me.day(6).ForeColor = IIf(WeekEndsOnDef, vbBlack, WeekEndsFont)
    End If
    Me.day(13).BackColor = IIf(WeekEndsOnDef, vbMagenta, WeekEndsBg)
    Me.day(13).ForeColor = IIf(WeekEndsOnDef, vbBlack, WeekEndsFont)
    Me.day(20).BackColor = IIf(WeekEndsOnDef, vbMagenta, WeekEndsBg)
    Me.day(20).ForeColor = IIf(WeekEndsOnDef, vbBlack, WeekEndsFont)
    Me.day(27).BackColor = IIf(WeekEndsOnDef, vbMagenta, WeekEndsBg)
    Me.day(27).ForeColor = IIf(WeekEndsOnDef, vbBlack, WeekEndsFont)
    If Me.day(34).Tag <> 0 Then
        Me.day(34).BackColor = IIf(WeekEndsOnDef, vbMagenta, WeekEndsBg)
        Me.day(34).ForeColor = IIf(WeekEndsOnDef, vbBlack, WeekEndsFont)
    End If
    If Me.day(41).Tag <> 0 Then
        Me.day(41).BackColor = IIf(WeekEndsOnDef, vbMagenta, WeekEndsBg)
        Me.day(41).ForeColor = IIf(WeekEndsOnDef, vbBlack, WeekEndsFont)
    End If
    Select Case MonthShown
        Case 1 'NewYears
            Me.day(StartBox).BackColor = IIf(HolidayOnDef, vbRed, HolidayBg)
            Me.day(StartBox).ForeColor = IIf(HolidayOnDef, vbBlack, HolidayFont)
            Me.day(StartBox).Caption = Me.day(StartBox).Caption & Chr(13) & "New Year's Day"
        Case 2 'Valentines
            Me.day(StartBox + 13).BackColor = IIf(HolidayOnDef, vbRed, HolidayBg)
            Me.day(StartBox + 13).ForeColor = IIf(HolidayOnDef, vbBlack, HolidayFont)
            Me.day(StartBox + 13).Caption = Me.day(StartBox + 13).Caption & Chr(13) & "Valentine's Day"
        Case 7 'July 4th
            Me.day(StartBox + 3).BackColor = IIf(HolidayOnDef, vbRed, HolidayBg)
            Me.day(StartBox + 3).ForeColor = IIf(HolidayOnDef, vbBlack, HolidayFont)
            Me.day(StartBox + 3).Caption = Me.day(StartBox + 3).Caption & Chr(13) & "American Independance Day"
        Case 10 'Haloween
            Me.day(StartBox + 30).BackColor = IIf(HolidayOnDef, vbRed, HolidayBg)
            Me.day(StartBox + 30).ForeColor = IIf(HolidayOnDef, vbBlack, HolidayFont)
            Me.day(StartBox + 30).Caption = Me.day(StartBox + 30).Caption & Chr(13) & "Halloween"
        Case 11 'Thanksgiving 4th thr
            Me.day(Thanksgiving).BackColor = IIf(HolidayOnDef, vbRed, HolidayBg)
            Me.day(Thanksgiving).ForeColor = IIf(HolidayOnDef, vbBlack, HolidayFont)
            Me.day(Thanksgiving).Caption = Me.day(Thanksgiving).Caption & Chr(13) & "Thanksgiving"
        Case 12 'Christmas
            Me.day(StartBox + 24).BackColor = IIf(HolidayOnDef, vbRed, HolidayBg)
            Me.day(StartBox + 24).ForeColor = IIf(HolidayOnDef, vbBlack, HolidayFont)
            Me.day(StartBox + 24).Caption = Me.day(StartBox + 24).Caption & Chr(13) & "Christmas"
    End Select
Else
    Select Case MonthShown
        Case 1 'NewYears
            Me.day(StartBox).Caption = Me.day(StartBox).Caption & Chr(13) & "New Year's Day"
        Case 2 'Valentines
            Me.day(StartBox + 13).Caption = Me.day(StartBox + 13).Caption & Chr(13) & "Valentine's Day"
        Case 7 'July 4th
            Me.day(StartBox + 3).Caption = Me.day(StartBox + 3).Caption & Chr(13) & "American Independance Day"
        Case 10 'Haloween
            Me.day(StartBox + 30).Caption = Me.day(StartBox + 30).Caption & Chr(13) & "Halloween"
        Case 11 'Thanksgiving 4th thr
            Me.day(Thanksgiving).Caption = Me.day(Thanksgiving).Caption & Chr(13) & "Thanksgiving"
        Case 12 'Christmas
            Me.day(StartBox + 24).Caption = Me.day(StartBox + 24).Caption & Chr(13) & "Christmas"
    End Select
End If

month_year.Caption = month_name(MonthShown) & " " & YearShown

If MonthShown = ThisMonth And YearShown = ThisYear Then
    Me.day(StartBox - 1 + ThisDay).BackColor = IIf(TodayOnDef, vbActiveTitleBar, TodayBg)
    Me.day(StartBox - 1 + ThisDay).ForeColor = IIf(TodayOnDef, vbActiveTitleBarText, TodayFont)
End If

End Sub

Public Sub ShowNewMonth()
    CheckLeapYear
    SetupColorCode
    SetupShownEvents
End Sub

Public Sub SetupShownEvents()
    Dim A As Integer
    
    For A = 0 To Index
        If Events(A).Title = "" Then GoTo nextn
        If Not Events(A).by_day_of_week Then 'BY DATE
            If Events(A).once Then 'Once
                    If Events(A).Year = YearShown And Events(A).Month = MonthShown Then Call Place(Events(A).day, Events(A).Title)
                ElseIf Events(A).monthly Then 'Monthly
                    Call Place(Events(A).day, Events(A).Title)
                Else 'Annually
                    If MonthShown = Events(A).Month Then Call Place(Events(A).day, Events(A).Title)
                End If
            Else 'BY DAY OF WEEK
                If Events(A).once Then 'Once
                    If YearShown = Events(A).Year And MonthShown = Events(A).Month Then Call PlaceDOW(Events(A).day_of_week, Events(A).Week, Events(A).Title)
                ElseIf Events(A).weekly Then 'Weekly
                    Call PlaceDOW(Events(A).day_of_week, 0, Events(A).Title)
                ElseIf Events(A).monthly Then 'Monthly
                    Call PlaceDOW(Events(A).day_of_week, Events(A).Week, Events(A).Title)
                Else 'Annually
                    If MonthShown = Events(A).Month Then Call PlaceDOW(Events(A).day_of_week, Events(A).Week, Events(A).Title)
                End If
        End If
nextn:
    Next A
End Sub

Public Sub Place(day As Integer, Title As String)
    Dim SB As Integer
    
    SB = Weekday(MonthShown & "/1/" & YearShown) - 1
    Me.day(SB + day - 1).Caption = Me.day(SB + day - 1).Caption & Chr(13) & Title
End Sub

Public Sub PlaceDOW(DOW As Integer, Week As Integer, Title As String)
    Dim n As Integer
    
    If Week = 0 Then
        For n = 0 To 5
            If Me.day(n * 7 + DOW - 1).Tag > 0 Then Me.day(n * 7 + DOW - 1).Caption = Me.day(n * 7 + DOW - 1).Caption & Chr(13) & Title
        Next n
    Else
        If Me.day(DOW - 1).Tag > 0 Then
            Me.day((Week - 1) * 7 + DOW - 1).Caption = Me.day((Week - 1) * 7 + DOW - 1).Caption & Chr(13) & Title
        Else
            Me.day(Week * 7 + DOW - 1).Caption = Me.day(Week * 7 + DOW - 1).Caption & Chr(13) & Title
        End If
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuColors_Click()
    Load fColors
End Sub

Private Sub mnuEdit_Click()
    fEdit.Show
End Sub

Private Sub mnuEnableBT_Click()
    mnuEnableBT.Checked = Not mnuEnableBT.Checked
    mnuBackTrack.Enabled = Not mnuEnableBT.Checked
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuManage_Click()
    fZEdit.Show
End Sub

Private Sub mnuNew_Click()
    fNew.Show
End Sub

Private Sub mnuNoSave_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Are you sure you don't want PC Calendar to save any changes you have made so far?", vbQuestion + vbYesNo, "Confirmation")
    
    If Ans = 6 Then
        AnyChanges = False
        CfgChanges = False
        EvtChanges = False
        
        LoadAll
    End If
End Sub

Private Sub mnuNoSaveCfg_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Are you sure you don't want PC Calendar to save any option changes you have made so far?", vbQuestion + vbYesNo, "Confirmation")
    
    If Ans = 6 Then
        CfgChanges = False
        
        LoadAll FA_CFG
    End If
End Sub

Private Sub mnuNoSaveEvt_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Are you sure you don't want PC Calendar to save any event changes you have made so far?", vbQuestion + vbYesNo, "Confirmation")
    
    If Ans = 6 Then
        EvtChanges = False
        
        LoadAll FA_EVENTS
    End If
End Sub

Private Sub mnuResetAll_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Are you sure you want to reset everything to defualt? You will loose all your events and option settings perminantly.", vbQuestion + vbYesNo, "Confirmation")
    
    If Ans = 6 Then
        AnyChanges = False
        CfgChanges = False
        EvtChanges = False
        
        SetDefEvents
        SetDefCfg
        
        SaveAll
    End If
End Sub

Private Sub mnuResetCfg_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Are you sure you want to reset your options to defualt? You will loose all your  option settings perminantly.", vbQuestion + vbYesNo, "Confirmation")
    
    If Ans = 6 Then
        CfgChanges = False
        
        SetDefCfg
        
        SaveAll FA_CFG
    End If
End Sub

Private Sub mnuResetEvt_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Are you sure you want to reset your events to defualt? You will loose all your events perminantly.", vbQuestion + vbYesNo, "Confirmation")
    
    If Ans = 6 Then
        EvtChanges = False
        
        SetDefEvents
        
        SaveAll FA_EVENTS
    End If
End Sub

Private Sub mnuStartup_Click()
    Load fOptions
End Sub

Private Sub mnuTips_Click()
    frmTip.Show
End Sub

Private Sub month_year_Click()
    fChange.Show
End Sub
