Attribute VB_Name = "Module2"
Function WriteSingle(FileNumber As Integer, Variable As Single)
    Put #FileNumber, Seek(FileNumber), Variable
End Function
Function WriteByte(FileNumber As Integer, Variable As Byte)
    Put #FileNumber, Seek(FileNumber), Variable
End Function
Function ReadSingle(FileNumber As Integer) As Single
    Get #FileNumber, Seek(FileNumber), ReadSingle
End Function
Function ReadByte(FileNumber As Integer) As Byte
    Get #FileNumber, Seek(FileNumber), ReadByte
End Function
Function ReadLong(FileNumber As Integer) As Long
    Get #FileNumber, Seek(FileNumber), ReadLong
End Function
Function ReadInteger(FileNumber As Integer) As Integer
    Get #FileNumber, Seek(FileNumber), ReadInteger
End Function
Function ReadString(FileNumber As Integer) As String
    Dim CurrentByte As Byte
    Do While EOF(FileNumber) = False
        CurrentByte = ReadByte(FileNumber)
        If CurrentByte <> 0 Then
            ReadString = ReadString & Chr(CurrentByte)
        Else
            Exit Function
        End If
    Loop
End Function
Function WriteLong(FileNumber As Integer, Variable As Long)
    Put #FileNumber, Seek(FileNumber), Variable
End Function
Function WriteInteger(FileNumber As Integer, Variable As Integer)
    Put #FileNumber, Seek(FileNumber), Variable
End Function
Function WriteString(FileNumber As Integer, Variable As String)
    Dim A As Integer
    For A = 1 To Len(Variable)
        Call WriteByte(FileNumber, Asc(Mid(Variable, A, 1)))
    Next A
    Call WriteByte(FileNumber, 0)
End Function
Function WriteBoolean(FileNumber As Integer, Variable As Boolean)
    If Variable Then
        WriteByte FileNumber, 0
    Else
        WriteByte FileNumber, 1
    End If
End Function
Function ReadBoolean(FileNumber As Integer) As Boolean
    Dim Temp_Byte As Byte
    Temp_Byte = ReadByte(FileNumber)
    If Temp_Byte = 0 Then ReadBoolean = True Else ReadBoolean = False
End Function

Public Function GetEvent(FileNumber As Integer) As Evt
    GetEvent.by_day_of_week = ReadBoolean(FileNumber)
    GetEvent.day_of_week = ReadInteger(FileNumber)
    GetEvent.Week = ReadInteger(FileNumber)
    GetEvent.Month = ReadInteger(FileNumber)
    GetEvent.day = ReadInteger(FileNumber)
    GetEvent.Year = ReadLong(FileNumber)
    GetEvent.annual = ReadBoolean(FileNumber)
    GetEvent.weekly = ReadBoolean(FileNumber)
    GetEvent.monthly = ReadBoolean(FileNumber)
    GetEvent.once = ReadBoolean(FileNumber)
    GetEvent.Title = ReadString(FileNumber)
    GetEvent.description = ReadString(FileNumber)
    GetEvent.alert = ReadBoolean(FileNumber)
    GetEvent.days_ahead_to_alert = ReadInteger(FileNumber)
End Function
Public Function SetEvent(FileNumber As Integer, Variable As Evt)
    WriteBoolean FileNumber, Variable.by_day_of_week
    WriteInteger FileNumber, Variable.day_of_week
    WriteInteger FileNumber, Variable.Week
    WriteInteger FileNumber, Variable.Month
    WriteInteger FileNumber, Variable.day
    WriteLong FileNumber, Variable.Year
    WriteBoolean FileNumber, Variable.annual
    WriteBoolean FileNumber, Variable.weekly
    WriteBoolean FileNumber, Variable.monthly
    WriteBoolean FileNumber, Variable.once
    WriteString FileNumber, Variable.Title
    WriteString FileNumber, Variable.description
    WriteBoolean FileNumber, Variable.alert
    WriteInteger FileNumber, Variable.days_ahead_to_alert
End Function

Public Sub WriteCfg(FileNumber As Integer)
    Dim A As Integer
    
    For A = 1 To 12
        WriteInteger FileNumber, days_in_month(A)
        WriteString FileNumber, month_name(A)
    Next A

    WriteBoolean FileNumber, TodayOnDef
    WriteLong FileNumber, TodayBg
    WriteLong FileNumber, TodayFont
    
    WriteBoolean FileNumber, BlankOnDef
    WriteLong FileNumber, BlankBg
    
    WriteBoolean FileNumber, RegularOnDef
    WriteLong FileNumber, RegularBg
    WriteLong FileNumber, RegularFont
    
    WriteBoolean FileNumber, MonthOnDef
    WriteLong FileNumber, MonthBg
    WriteLong FileNumber, MonthFont
    
    WriteBoolean FileNumber, DOWOnDef
    WriteLong FileNumber, DOWBg
    WriteLong FileNumber, DOWFont
    
    WriteBoolean FileNumber, chkHol
    
    WriteBoolean FileNumber, HolidayOnDef
    WriteLong FileNumber, HolidayBg
    WriteLong FileNumber, HolidayFont
    
    WriteBoolean FileNumber, WeekEndsOnDef
    WriteLong FileNumber, WeekEndsBg
    WriteLong FileNumber, WeekEndsFont
    
    WriteBoolean FileNumber, Prompt
    WriteByte FileNumber, CByte(OnceEv)
    
    WriteBoolean FileNumber, DoubleWarning
    
    WriteBoolean FileNumber, ShowTips
End Sub

Public Sub ReadCfg(FileNumber As Integer)
    Dim A As Integer
    
    For A = 1 To 12
        days_in_month(A) = ReadInteger(FileNumber)
        month_name(A) = ReadString(FileNumber)
    Next A

    TodayOnDef = ReadBoolean(FileNumber)
    TodayBg = ReadLong(FileNumber)
    TodayFont = ReadLong(FileNumber)
    
    BlankOnDef = ReadBoolean(FileNumber)
    BlankBg = ReadLong(FileNumber)
    
    RegularOnDef = ReadBoolean(FileNumber)
    RegularBg = ReadLong(FileNumber)
    RegularFont = ReadLong(FileNumber)
    
    MonthOnDef = ReadBoolean(FileNumber)
    MonthBg = ReadLong(FileNumber)
    MonthFont = ReadLong(FileNumber)
    
    DOWOnDef = ReadBoolean(FileNumber)
    DOWBg = ReadLong(FileNumber)
    DOWFont = ReadLong(FileNumber)
    
    chkHol = ReadBoolean(FileNumber)
    
    HolidayOnDef = ReadBoolean(FileNumber)
    HolidayBg = ReadLong(FileNumber)
    HolidayFont = ReadLong(FileNumber)
    
    WeekEndsOnDef = ReadBoolean(FileNumber)
    WeekEndsBg = ReadLong(FileNumber)
    WeekEndsFont = ReadLong(FileNumber)
    
    Prompt = ReadBoolean(FileNumber)
    OnceEv = ReadByte(FileNumber)
    
    LastActivation = GetSetting("PC Calendar", "Options", "LA", "1/1/2000")
    If LastActivation = "" Then LastActivation = Now
    DoubleWarning = ReadBoolean(FileNumber)
    
    ShowTips = ReadBoolean(FileNumber)
End Sub

Public Sub ReadEvents(FileNumber As Integer)
    Dim A As Integer
    Index = ReadInteger(FileNumber)
    ReDim Events(Index)
    For A = 0 To Index
        Events(A) = GetEvent(FileNumber)
    Next A
End Sub

Public Sub WriteEvents(FileNumber As Integer)
    Dim A As Integer
    WriteInteger FileNumber, Index
    For A = 0 To Index
        SetEvent FileNumber, Events(A)
    Next A
End Sub

Public Sub SetDefCfg()
    days_in_month(1) = 31
    month_name(1) = "January"
    days_in_month(2) = 28
    month_name(2) = "February"
    days_in_month(3) = 31
    month_name(3) = "March"
    days_in_month(4) = 30
    month_name(4) = "April"
    days_in_month(5) = 31
    month_name(5) = "May"
    days_in_month(6) = 30
    month_name(6) = "June"
    days_in_month(7) = 31
    month_name(7) = "July"
    days_in_month(8) = 31
    month_name(8) = "August"
    days_in_month(9) = 30
    month_name(9) = "September"
    days_in_month(10) = 31
    month_name(10) = "October"
    days_in_month(11) = 30
    month_name(11) = "November"
    days_in_month(12) = 31
    month_name(12) = "December"

    TodayOnDef = True
    TodayBg = vbActiveTitleBar
    TodayFont = vbActiveTitleBarText
    
    BlankOnDef = True
    BlankBg = vbInactiveTitleBar
    
    RegularOnDef = True
    RegularBg = vbWhite
    RegularFont = vbBlack
    
    MonthOnDef = True
    MonthBg = vbButtonFace
    MonthFont = vbButtonText
    
    DOWOnDef = True
    DOWBg = vbButtonFace
    DOWFont = vbButtonText
    
    chkHol = False
    
    HolidayOnDef = True
    HolidayBg = vbRed
    HolidayFont = vbBlack
    
    WeekEndsOnDef = True
    WeekEndsBg = vbMagenta
    WeekEndsFont = vbBlack
    
    Minimize = False
    Prompt = False
    OnceEv = PE_ASK
    DoubleWarning = True
    
    LastActivation = Now
    ShowTips = True
End Sub

Public Sub LoadAll(Optional FileAccess As FileAccess)
    On Error Resume Next
    Dim FF As Integer
    FF = FreeFile
    
    If IsMissing(FileAccess) Or FileAccess = FA_BOTH Or FileAccess = FA_CFG Then
        If FileLen(App.Path & "\Options.dat") = 0 Then
            MsgBox "This is your first time running this program. PC Calendar will now automatically build the files it needs.", , "First Time Use"
            SetDefCfg
            SaveAll FA_CFG
            Kill App.Path & "\Events.dat"
            SetDefEvents
            SaveAll FA_EVENTS
        Else
            Open App.Path & "\Options.dat" For Binary Access Read Write As #FF
                ReadCfg FF
            Close #FF
        End If
    End If
    
    If IsMissing(FileAccess) Or FileAccess = FA_BOTH Or FileAccess = FA_EVENTS Then
        If FileLen(App.Path & "\Events.dat") = 0 Then
            MsgBox "You have no events loaded on PC Calendar. You should never see this error message.", vbCritical, "Error"
            SetDefEvents
            SaveAll FA_EVENTS
        Else
            FF = FreeFile
            Open App.Path & "\Events.dat" For Binary Access Read As #FF
                ReadEvents FF
            Close #FF
        End If
    End If
End Sub

Public Sub SaveAll(Optional FileAccess As FileAccess)
    Dim FF As Integer
    On Error Resume Next
    If IsMissing(FileAccess) Or FileAccess = FA_BOTH Or FileAccess = FA_CFG Then
        FF = FreeFile
        Kill App.Path & "\Options.dat"
        Open App.Path & "\Options.dat" For Binary Access Write As #FF
            WriteCfg FF
        Close #FF
    End If
    
    If IsMissing(FileAccess) Or FileAccess = FA_BOTH Or FileAccess = FA_EVENTS Then
        FF = FreeFile
        Kill App.Path & "\Events.dat"
        Open App.Path & "\Events.dat" For Binary Access Write As #FF
            WriteEvents FF
        Close #FF
    End If
End Sub

Public Sub SetDefEvents()
    Index = 0
    ReDim Events(Index)
    Events(0) = AuthorBDay
End Sub
