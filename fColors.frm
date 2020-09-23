VERSION 5.00
Begin VB.Form fColors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Options"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7200
      TabIndex        =   45
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   8760
      TabIndex        =   44
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   5640
      TabIndex        =   43
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame Frame 
      Caption         =   "Weekends"
      Height          =   855
      Index           =   6
      Left            =   5280
      TabIndex        =   37
      Top             =   2160
      Width           =   5175
      Begin VB.CheckBox chkOnDef 
         Caption         =   "Use Defualt Color"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdFont 
         Height          =   255
         Index           =   6
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdBackground 
         Height          =   255
         Index           =   6
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Font Color:"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   42
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Background Color:"
         Height          =   255
         Index           =   6
         Left            =   3360
         TabIndex        =   41
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Important Holidays"
      Height          =   855
      Index           =   5
      Left            =   5280
      TabIndex        =   31
      Top             =   1200
      Width           =   5175
      Begin VB.CheckBox chkOnDef 
         Caption         =   "Use Defualt Color"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdFont 
         Height          =   255
         Index           =   5
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdBackground 
         Height          =   255
         Index           =   5
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Font Color:"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   36
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Background Color:"
         Height          =   255
         Index           =   5
         Left            =   3360
         TabIndex        =   35
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CheckBox chkUseSpecial 
      Caption         =   "&Highlight Special Days"
      Height          =   255
      Left            =   6360
      TabIndex        =   30
      Top             =   960
      Width           =   2655
   End
   Begin VB.Frame Frame 
      Caption         =   "Days of the Week Heading"
      Height          =   855
      Index           =   4
      Left            =   0
      TabIndex        =   24
      Top             =   3120
      Width           =   5175
      Begin VB.CheckBox chkOnDef 
         Caption         =   "Use Defualt Color"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdFont 
         Height          =   255
         Index           =   4
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdBackground 
         Height          =   255
         Index           =   4
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Font Color:"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Background Color:"
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   28
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Month and Year Heading"
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   18
      Top             =   2160
      Width           =   5175
      Begin VB.CheckBox chkOnDef 
         Caption         =   "Use Defualt Color"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdFont 
         Height          =   255
         Index           =   3
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdBackground 
         Height          =   255
         Index           =   3
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Font Color:"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Background Color:"
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Today"
      Height          =   855
      Index           =   2
      Left            =   0
      TabIndex        =   12
      Top             =   1200
      Width           =   5175
      Begin VB.CheckBox chkOnDef 
         Caption         =   "Use Defualt Color"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdFont 
         Height          =   255
         Index           =   2
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdBackground 
         Height          =   255
         Index           =   2
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Font Color:"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Background Color:"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Regular Days"
      Height          =   855
      Index           =   1
      Left            =   5280
      TabIndex        =   6
      Top             =   0
      Width           =   5175
      Begin VB.CheckBox chkOnDef 
         Caption         =   "Use Defualt Color"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdFont 
         Height          =   255
         Index           =   1
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdBackground 
         Height          =   255
         Index           =   1
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Font Color:"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Background Color:"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Blank Days"
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton cmdBackground 
         Height          =   255
         Index           =   0
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdFont 
         Caption         =   "N/A"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkOnDef 
         Caption         =   "Use Defualt Color"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Background Color:"
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Font Color:"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "fColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FontC(0 To 6) As Long
Dim BackgroundC(0 To 6) As Long
Dim OnDef(0 To 6) As Boolean

Private Enum ColorType
    CT_BLANK = 0
    CT_REGULAR = 1
    CT_TODAY = 2
    CT_MONTH_YEAR = 3
    CT_DOW = 4
    CT_HOLIDAYS = 5
    CT_WEEKENDS = 6
End Enum

Public Sub EnableFrame(Index As Integer, Enabled As Boolean)
    Frame(Index).Enabled = Enabled
    chkOnDef(Index).Enabled = Enabled
    Label1(Index).Enabled = Enabled
    Label2(Index).Enabled = Enabled
    cmdFont(Index).Enabled = Enabled
    cmdBackground(Index).Enabled = Enabled
End Sub

Public Sub SetBackToReal()
    chkUseSpecial.Value = IIf(chkHol, 1, 0)

    OnDef(CT_BLANK) = BlankOnDef
    BackgroundC(CT_BLANK) = BlankBg
    
    OnDef(CT_REGULAR) = RegularOnDef
    BackgroundC(CT_REGULAR) = RegularBg
    FontC(CT_REGULAR) = RegularFont
    
    OnDef(CT_TODAY) = TodayOnDef
    BackgroundC(CT_TODAY) = TodayBg
    FontC(CT_TODAY) = TodayFont
    
    OnDef(CT_MONTH_YEAR) = MonthOnDef
    BackgroundC(CT_MONTH_YEAR) = MonthBg
    FontC(CT_MONTH_YEAR) = MonthFont
    
    OnDef(CT_DOW) = DOWOnDef
    BackgroundC(CT_DOW) = DOWBg
    FontC(CT_DOW) = DOWFont
    
    OnDef(CT_HOLIDAYS) = HolidayOnDef
    BackgroundC(CT_HOLIDAYS) = HolidayBg
    FontC(CT_HOLIDAYS) = HolidayFont
    
    OnDef(CT_WEEKENDS) = WeekEndsOnDef
    BackgroundC(CT_WEEKENDS) = WeekEndsBg
    FontC(CT_WEEKENDS) = WeekEndsFont
End Sub

Public Sub SetRealToBack()
    chkHol = (chkUseSpecial.Value = 1)
    
    BlankOnDef = OnDef(CT_BLANK)
    BlankBg = BackgroundC(CT_BLANK)
    
    RegularOnDef = OnDef(CT_REGULAR)
    RegularBg = BackgroundC(CT_REGULAR)
    RegularFont = FontC(CT_REGULAR)
    
    TodayOnDef = OnDef(CT_TODAY)
    TodayBg = BackgroundC(CT_TODAY)
    TodayFont = FontC(CT_TODAY)
    
    MonthOnDef = OnDef(CT_MONTH_YEAR)
    MonthBg = BackgroundC(CT_MONTH_YEAR)
    MonthFont = FontC(CT_MONTH_YEAR)
    
    DOWOnDef = OnDef(CT_DOW)
    DOWBg = BackgroundC(CT_DOW)
    DOWFont = FontC(CT_DOW)
    
    HolidayOnDef = OnDef(CT_HOLIDAYS)
    HolidayBg = BackgroundC(CT_HOLIDAYS)
    HolidayFont = FontC(CT_HOLIDAYS)
    
    WeekEndsOnDef = OnDef(CT_WEEKENDS)
    WeekEndsBg = BackgroundC(CT_WEEKENDS)
    WeekEndsFont = FontC(CT_WEEKENDS)
End Sub

Public Sub SetFont(Number As Integer, Color As Long)
    FontC(Number) = Color
    cmdFont(Number).BackColor = Color
End Sub

Public Sub SetBackground(Number As Integer, Color As Long)
    BackgroundC(Number) = Color
    cmdBackground(Number).BackColor = Color
End Sub

Private Sub chkOnDef_Click(Index As Integer)
    OnDef(Index) = (chkOnDef(Index).Value = 1)
End Sub

Private Sub chkUseSpecial_Click()
    Dim Temp As Boolean
    Temp = (chkUseSpecial.Value = 1)
    
    EnableFrame 5, Temp
    EnableFrame 6, Temp
End Sub

Private Sub cmdApply_Click()
    AnyChanges = True
    CfgChanges = True
    SetRealToBack
    frm.UpdateColor
    frm.ShowNewMonth
End Sub

Private Sub cmdBackground_Click(Index As Integer)
    Dim Temp As Long
    
    Temp = BackgroundC(Index)
    SetBackground Index, GetUserColor(Temp)
    
    If Temp <> BackgroundC(Index) Then
        AnyChanges = True
        CfgChanges = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFont_Click(Index As Integer)
    Dim Temp As Long
    
    Temp = FontC(Index)
    SetFont Index, GetUserColor(Temp)

    If Temp <> FontC(Index) Then
        AnyChanges = True
        CfgChanges = True
    End If
End Sub

Private Sub cmdOK_Click()
    AnyChanges = True
    CfgChanges = True
    SetRealToBack
    frm.UpdateColor
    frm.ShowNewMonth
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim A As Integer
    
    SetBackToReal

    For A = 0 To 6
        SetFont A, FontC(A)
        SetBackground A, BackgroundC(A)
        chkOnDef(A).Value = IIf(OnDef(A), 1, 0)
    Next A
    
    If Not chkHol Then
        EnableFrame 5, False
        EnableFrame 6, False
    End If
    
    Me.Visible = True
End Sub
