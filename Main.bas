Attribute VB_Name = "Module1"
Public TodayOnDef As Boolean
Public TodayBg As Long
Public TodayFont As Long
Public BlankOnDef As Boolean
Public BlankBg As Long
Public RegularOnDef As Boolean
Public RegularBg As Long
Public RegularFont As Long
Public MonthOnDef As Boolean
Public MonthBg As Long
Public MonthFont As Long
Public DOWOnDef As Boolean
Public DOWBg As Long
Public DOWFont As Long
Public chkHol As Boolean
Public HolidayOnDef As Boolean
Public HolidayBg As Long
Public HolidayFont As Long
Public WeekEndsOnDef As Boolean
Public WeekEndsBg As Long
Public WeekEndsFont As Long

Public Enum PastEvents
    PE_DELETE = 0
    PE_ASK = 1
    PE_IGNORE = 2
End Enum
Public Enum FileAccess
    FA_BOTH = 0
    FA_CFG = 1
    FA_EVENTS = 2
End Enum

Public Inserting As Boolean
Public InsertSlot As Integer

Public ChoosenColor As Long
Public DoneChoosing As Boolean

Public Prompt As Boolean
Public OnceEv As PastEvents
Public DoubleWarning As Boolean
Public ActuallyChecked As Boolean

Public Const LEFT_MARGIN = 120
Public Const RIGHT_MARGIN = 240
Public Const TOP_MARGIN = 990
Public Const BOTTOM_MARGIN = 480
Public Const SPACING = 0

Public Type Evt
    by_day_of_week As Boolean
    day_of_week As Integer
    Week As Integer
    
    Month As Integer
    day As Integer
    Year As Long
    
    annual As Boolean
    weekly As Boolean
    monthly As Boolean
    once As Boolean
    
    Title As String
    description As String
    
    alert As Boolean
    days_ahead_to_alert As Integer
End Type

Public ClipBoard As Evt
Public AnythingCopied As Boolean

Public AnyChanges As Boolean
Public CfgChanges As Boolean
Public EvtChanges As Boolean

Public FiredEvents() As Integer
Public LastActivation As String
Public ShowTips As Boolean

Public Index As Integer
Public Events() As Evt
Public MonthShown As Integer
Public YearShown As Long
Public Today As Date
Public days_in_month(0 To 12) As Integer
Public month_name(0 To 12) As String
Public ThisMonth As Integer
Public ThisDay As Integer
Public ThisWeekday As Integer
Public ThisYear As Long

Public Sub Main()
    SetCurrentDate
    LoadAll
    CheckEvents
    DealWithEvents
End Sub

Public Sub ApplyCfg()
    frm.month_year.Caption = month_name(ThisMonth) & " " & ThisYear
    frm.UpdateColor
End Sub

Public Function NextOccurance(TEV As Evt, Optional FromWhen As String) As String
Attribute NextOccurance.VB_Description = "Returns the date of the next occurance of an event"
Dim vMonth As Integer
Dim vDay As Integer
Dim vWeekday As Integer
Dim vYear As Integer

If FromWhen = "" Then FromWhen = Now

vMonth = Month(FromWhen)
vDay = day(FromWhen)
vWeekday = Weekday(FromWhen)
vYear = Year(FromWhen)

Dim FDOW As Integer
If TEV.by_day_of_week = False Then
    If TEV.once Then 'Once
        Select Case TEV.Year
            Case Is < vYear
                NextOccurance = "1/1/0001"
                Exit Function
            Case Is = vYear
                Select Case TEV.Month
                    Case Is < vMonth
                        NextOccurance = "1/1/0001"
                        Exit Function
                    Case Is = vMonth
                        If TEV.day < vDay Then
                            NextOccurance = "1/1/0001"
                            Exit Function
                        Else
                            NextOccurance = TEV.Month & "/" & TEV.day & "/" & TEV.Year
                            Exit Function
                        End If
                    Case Is > vMonth
                        NextOccurance = TEV.Month & "/" & TEV.day & "/" & TEV.Year
                        Exit Function
                End Select
            Case Is > vYear
                NextOccurance = TEV.Month & "/" & TEV.day & "/" & TEV.Year
                Exit Function
        End Select
    ElseIf TEV.monthly Then 'Monthly
        If TEV.day < vDay Then
            If vMonth = 12 Then NextOccurance = "1/" & TEV.day & "/" & vYear + 1 Else NextOccurance = vMonth + 1 & "/" & TEV.day & "/" & vYear
            Exit Function
        Else
            NextOccurance = vMonth & "/" & TEV.day & "/" & vYear
            Exit Function
        End If
    Else 'Annually
        Select Case TEV.Month
            Case Is < vMonth
                NextOccurance = TEV.Month & "/" & TEV.day & "/" & vYear + 1
                Exit Function
            Case Is = vMonth
                Select Case TEV.day
                    Case Is < vDay
                        NextOccurance = TEV.Month & "/" & TEV.day & "/" & vYear + 1
                        Exit Function
                    Case Is >= vDay
                        NextOccurance = TEV.Month & "/" & TEV.day & "/" & vYear
                        Exit Function
                End Select
            Case Is > vMonth
                NextOccurance = TEV.Month & "/" & TEV.day & "/" & vYear
                Exit Function
        End Select
    End If
Else
    FDOW = Weekday(vMonth & "/1/" & vYear)
    If TEV.once Then 'Once
        Select Case TEV.Year
            Case Is < vYear
                NextOccurance = "1/1/0001"
                Exit Function
            Case Is = vYear
                Select Case TEV.Month
                    Case Is < vMonth
                        NextOccurance = "1/1/0001"
                        Exit Function
                    Case Is = vMonth
                        If FDOW <= TEV.day_of_week Then
                            If (TEV.Week - 1) * 7 + TEV.day_of_week - FDOW + 1 < vDay Then
                                NextOccurance = "1/1/0001"
                                Exit Function
                            Else
                                NextOccurance = vMonth & "/" & (TEV.Week - 1) * 7 + TEV.day_of_week - FDOW + 1 & "/" & vYear
                                Exit Function
                            End If
                        Else
                            If (TEV.Week) * 7 + TEV.day_of_week - FDOW + 1 < vDay Then
                                NextOccurance = "1/1/0001"
                                Exit Function
                            Else
                                NextOccurance = vMonth & "/" & (TEV.Week) * 7 + TEV.day_of_week - FDOW + 1 & "/" & vYear
                                Exit Function
                            End If
                        End If
                    Case Is > vMonth
                        If FDOW <= TEV.day_of_week Then
                            NextOccurance = vMonth & "/" & (TEV.Week - 1) * 7 + TEV.day_of_week - FDOW + 1 & "/" & vYear
                            Exit Function
                        Else
                            NextOccurance = vMonth & "/" & (TEV.Week) * 7 + TEV.day_of_week - FDOW + 1 & "/" & vYear
                            Exit Function
                        End If
                End Select
            Case Is > vYear
                If FDOW <= TEV.day_of_week Then
                    NextOccurance = vMonth & "/" & (TEV.Week - 1) * 7 + TEV.day_of_week - FDOW + 1 & "/" & vYear
                    Exit Function
                Else
                    NextOccurance = vMonth & "/" & (TEV.Week) * 7 + TEV.day_of_week - FDOW + 1 & "/" & vYear
                    Exit Function
                End If
        End Select
    ElseIf TEV.weekly Then 'Weekly
        If vWeekday > TEV.day_of_week Then
            If vDay + 7 - vWeekday + TEV.day_of_week > days_in_month(vMonth) Then
                If vMonth = 12 Then
                    NextOccurance = "1/" & vDay + 7 - vWeekday + TEV.day_of_week - days_in_month(vMonth) & "/" & vYear + 1
                    Exit Function
                Else
                    NextOccurance = vMonth + 1 & "/" & vDay + 7 - vWeekday + TEV.day_of_week - days_in_month(vMonth) & "/" & vYear
                    Exit Function
                End If
            Else
                NextOccurance = vMonth & "/" & vDay + 7 - vWeekday + TEV.day_of_week & "/" & vYear
                Exit Function
            End If
        Else
            If vDay + TEV.day_of_week - vWeekday > days_in_month(vMonth) Then
                If vMonth = 12 Then
                    NextOccurance = "1/" & vDay + TEV.day_of_week - vWeekday - days_in_month(vMonth) & "/" & vYear + 1
                    Exit Function
                Else
                    NextOccurance = vMonth + 1 & "/" & vDay + TEV.day_of_week - vWeekday - days_in_month(vMonth) & "/" & vYear
                    Exit Function
                End If
            Else
                NextOccurance = vMonth & "/" & vDay + TEV.day_of_week - vWeekday & "/" & vYear
                Exit Function
            End If
        End If
    ElseIf TEV.monthly Then 'Monthly
        If FDOW <= TEV.day_of_week Then
            If (TEV.Week - 1) * 7 + TEV.day_of_week - FDOW + 1 < vDay Then
less:
                If vMonth = 12 Then
                    FDOW = Weekday("1/1/" & vYear + 1)
                    NextOccurance = "1/" & (TEV.Week - 1) * 7 + TEV.day_of_week - FDOW + 1 & "/" & vYear + 1
                Else
                    FDOW = Weekday(vMonth + 1 & "/1/" & vYear)
                    NextOccurance = vMonth + 1 & "/" & (TEV.Week - 1) * 7 + TEV.day_of_week - FDOW + 1 & "/" & vYear
                End If
                If Not FDOW <= TEV.day_of_week Then GoTo more
            Else
                NextOccurance = vMonth & "/" & (TEV.Week - 1) * 7 + TEV.day_of_week - FDOW + 1 & "/" & vYear
                Exit Function
            End If
        Else
            If (TEV.Week) * 7 + TEV.day_of_week - FDOW + 1 < vDay Then
more:
                If vMonth = 12 Then
                    FDOW = Weekday("1/1/" & vYear + 1)
                    NextOccurance = "1/" & (TEV.Week) * 7 + TEV.day_of_week - FDOW + 1 & "/" & vYear + 1
                Else
                    FDOW = Weekday(vMonth + 1 & "/1/" & vYear)
                    NextOccurance = vMonth + 1 & "/" & (TEV.Week) * 7 + TEV.day_of_week - FDOW + 1 & "/" & vYear
                End If
                If FDOW <= TEV.day_of_week Then GoTo less
            Else
                NextOccurance = vMonth & "/" & (TEV.Week) * 7 + TEV.day_of_week - FDOW + 1 & "/" & vYear
                Exit Function
            End If
        End If
    Else 'Annually
        Select Case TEV.Month
            Case Is < vMonth
                GoTo nextyear
            Case Is = vMonth
                FDOW = Weekday(TEV.Month & "/1/" & vYear)
                If FDOW <= TEV.day_of_week Then
                    If (TEV.Week - 1) * 7 + TEV.day_of_week - FDOW + 1 > vDay Then GoTo vYear Else GoTo nextyear
                Else
                    If (TEV.Week) * 7 + TEV.day_of_week - FDOW + 1 > vDay Then GoTo vYear Else GoTo nextyear
                End If
            Case Is > vMonth
                GoTo vYear
        End Select
        Exit Function
vYear:
        FDOW = Weekday(TEV.Month & "/1/" & vYear)
        If FDOW <= TEV.day_of_week Then
            NextOccurance = TEV.Month & "/" & (TEV.Week - 1) * 7 + TEV.day_of_week - FDOW + 1 & "/" & vYear
        Else
            NextOccurance = TEV.Month & "/" & (TEV.Week) * 7 + TEV.day_of_week - FDOW + 1 & "/" & vYear
        End If
        Exit Function
nextyear:
        FDOW = Weekday(TEV.Month & "/1/" & vYear + 1)
        If FDOW <= TEV.day_of_week Then
            NextOccurance = TEV.Month & "/" & (TEV.Week - 1) * 7 + TEV.day_of_week - FDOW + 1 & "/" & vYear + 1
        Else
            NextOccurance = TEV.Month & "/" & (TEV.Week) * 7 + TEV.day_of_week - FDOW + 1 & "/" & vYear + 1
        End If
    End If
End If
End Function

Public Function DaysAway(Evnt As Evt) As Long
Attribute DaysAway.VB_Description = "Returns the number of days away an event is"
    Dim NO As String
    'On Error GoTo ErrorHandler
    
    NO = NextOccurance(Evnt)
    If NO = "1/1/0001" Then
        DaysAway = Evnt.days_ahead_to_alert + 2
    Else
        DaysAway = Abs(DateDiff("d", NO, Today))
    End If
Exit Function
ErrorHandler:
    DaysAway = 0
End Function

Public Function IsAlarmed(Evnt As Evt) As Boolean
Attribute IsAlarmed.VB_Description = "Determines if the event needs to be brought to the user's attention"
    If DaysAway(Evnt) <= Evnt.days_ahead_to_alert _
        And Evnt.alert Then _
        IsAlarmed = True _
    Else _
        IsAlarmed = False
End Function

Public Function AuthorBDay() As Evt
    AuthorBDay.alert = True
    AuthorBDay.annual = True
    AuthorBDay.by_day_of_week = False
    AuthorBDay.day = 9
    AuthorBDay.day_of_week = 1
    AuthorBDay.days_ahead_to_alert = 7
    AuthorBDay.description = "The Birthday of the author of this program is coming up. You can wish me a happy birthday by e-mailing me at johnf@teleport.com"
    AuthorBDay.Month = 3
    AuthorBDay.monthly = False
    AuthorBDay.once = False
    AuthorBDay.Title = "Johnathan's Birthday"
    AuthorBDay.Week = 1
    AuthorBDay.weekly = False
    AuthorBDay.Year = 1984
End Function

Public Function NeedToCheckEvents() As Boolean
    Dim DaysPast As Long
    DaysPast = DateDiff("d", Today, LastActivation)
    If DaysPast = 0 And Not DoubleWarning Then _
        NeedToCheckEvents = False _
    Else _
        NeedToCheckEvents = True
End Function

Public Sub SetCurrentDate()
    AnythingCopied = False

    AnyChanges = False
    CfgChanges = False
    EvtChanges = False
    
    Today = Now
    ThisMonth = Month(Today)
    ThisDay = day(Today)
    ThisWeekday = Weekday(Today)
    ThisYear = Year(Today)
    MonthShown = ThisMonth
    YearShown = ThisYear
End Sub

Public Sub CheckEvents()
    Dim A As Integer
    Dim Ans As Integer
    
    ReDim FiredEvents(0)
    FiredEvents(0) = -1
    
    Dim FirstRound As Boolean
    FirstRound = True
    
    ActuallyChecked = False
    
    If NeedToCheckEvents Then
        ActuallyChecked = True
        For A = 0 To Index
            If IsAlarmed(Events(A)) Then
                If FirstRound Then
                    FirstRound = False
                    FiredEvents(0) = A
                Else
                    ReDim Preserve FiredEvents(UBound(FiredEvents) + 1)
                    FiredEvents(UBound(FiredEvents)) = A
                End If
            ElseIf Events(A).once Then
                Select Case OnceEv
                    Case PE_DELETE
                        Ans = 6
                    Case PE_ASK
                        Ans = MsgBox(Events(A).description & ", a past event, will never trigger again. Would you like to delete it?", vbYesNo, "Old Event")
                    Case PE_IGNORE
                        Ans = 7
                End Select
                If Ans = 6 Then
                    RemoveEvent A
                End If
            End If
        Next A
    End If
End Sub

Public Function AnyEventsFired() As Boolean
    If FiredEvents(0) = -1 Then AnyEventsFired = False Else AnyEventsFired = True
End Function

Public Sub DealWithEvents()
    Dim Ans As Integer
    If AnyEventsFired Then
        Load fAlert
    Else
        If Prompt Then
            If ActuallyChecked Then
                Ans = MsgBox("No events were triggered. Would you like to quit?", vbYesNo + vbQuestion, "No Events")
            Else
                Ans = MsgBox("You have already checked your events today. Would you like to quit?", vbYesNo + vbQuestion, "Already Checked")
            End If
            If Ans = 6 Then
                SaveSetting "PC Calendar", "Options", "LA", Now
                End
            End If
        End If
        Load frm
    End If
End Sub

Public Sub CheckLeapYear()
    If (YearShown - 2000) Mod 4 = 0 And MonthShown = 2 Then days_in_month(2) = 29 Else days_in_month(2) = 28
End Sub
Public Function Thanksgiving()
    Dim FDOW As Integer
    FDOW = Weekday("11/1/" & YearShown)
    
    If FDOW <= 5 Then 'Included in first week?
        Thanksgiving = 25
    Else
        Thanksgiving = 32
    End If
End Function

Public Sub RemoveEvent(Number As Integer)
    AnyChanges = True
    EvtChanges = True
    Dim A As Integer
    
    For A = Number To Index - 1
        Events(A) = Events(A + 1)
    Next A
    
    Index = Index - 1
    ReDim Preserve Events(Index)
End Sub

Public Sub OverlapEvent(Number As Integer, Evnt As Evt)
    AnyChanges = True
    EvtChanges = True
    Events(Number) = Evnt
End Sub

Public Sub AddNewEvent(Evnt As Evt)
    AnyChanges = True
    EvtChanges = True
    
    Index = Index + 1
    ReDim Preserve Events(Index)
    
    Events(Index) = Evnt
End Sub

Public Sub SwitchEvents(Number1 As Integer, Number2 As Integer)
    Dim Temp_Event As Evt
    AnyChanges = True
    EvtChanges = True
    
    Temp_Event = Events(Number1)
    Events(Number1) = Events(Number2)
    Events(Number2) = Temp_Event
End Sub

Public Function GetUserColor(DefualtColor As Long)
    DoneChoosing = False
    
    Load fChooseColor
    fChooseColor.Start DefualtColor
    
    Do Until DoneChoosing = True
        DoEvents
    Loop
    
    GetUserColor = ChoosenColor
End Function

Public Sub CopyEvent(Number As Integer)
    AnythingCopied = True
    ClipBoard = Events(Number)
End Sub

Public Sub PasteEvent()
    AnyChanges = True
    EvtChanges = True

    AddNewEvent ClipBoard
End Sub

Public Sub CutEvent(Number As Integer)
    AnyChanges = True
    EvtChanges = True

    AnythingCopied = True
    CopyEvent Number
    RemoveEvent Number
End Sub

Public Sub InsertEvent(Number As Integer, Evnt As Evt)
    Dim A As Integer
    
    AnyChanges = True
    EvtChanges = True
    
    Index = Index + 1
    ReDim Preserve Events(Index)
    
    For A = Index - 1 To Number Step -1
        Events(A + 1) = Events(A)
    Next A
    
    OverlapEvent Number, Evnt
End Sub

Public Function EventOnDay(Number As Integer, Date1 As String) As Boolean
On Error GoTo ErrorW
    If Events(Number).Title = "Church" Then
    Dim X As String
    X = NextOccurance(Events(Number), Date1)
    X = ""
    End If
    
    If DateDiff("d", Date1, NextOccurance(Events(Number), Date1)) = 0 Then _
        EventOnDay = True _
    Else _
        EventOnDay = False
Exit Function
ErrorW:
    EventOnDay = False
End Function
