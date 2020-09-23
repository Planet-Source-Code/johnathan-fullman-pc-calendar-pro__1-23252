VERSION 5.00
Begin VB.Form fChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jump To Date"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsbMonth 
      Height          =   255
      Left            =   120
      Max             =   11
      TabIndex        =   7
      Top             =   960
      Width           =   2895
   End
   Begin VB.VScrollBar vsbYear 
      Height          =   1815
      Index           =   3
      Left            =   4440
      Max             =   9
      TabIndex        =   6
      Top             =   960
      Width           =   255
   End
   Begin VB.VScrollBar vsbYear 
      Height          =   1815
      Index           =   2
      Left            =   4080
      Max             =   9
      TabIndex        =   5
      Top             =   960
      Width           =   255
   End
   Begin VB.VScrollBar vsbYear 
      Height          =   1815
      Index           =   1
      Left            =   3720
      Max             =   9
      TabIndex        =   4
      Top             =   960
      Width           =   255
   End
   Begin VB.VScrollBar vsbYear 
      Height          =   1815
      Index           =   0
      Left            =   3360
      Max             =   9
      TabIndex        =   3
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblYD 
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "fChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim New_Month As Integer
Dim New_Year As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    YearShown = New_Year
    MonthShown = New_Month
    frm.ShowNewMonth
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Move frm.Left + frm.month_year.Left, frm.Top + 120

    lblYD.Caption = month_name(MonthShown) & " " & YearShown
    hsbMonth.Value = MonthShown - 1
    vsbYear(0).Value = Int(YearShown / 1000)
    vsbYear(1).Value = Int(YearShown / 100) - 10 * Int(YearShown / 1000)
    vsbYear(2).Value = Int(YearShown / 10) - 10 * Int(YearShown / 100)
    vsbYear(3).Value = YearShown - 10 * Int(YearShown / 10)
End Sub

Private Sub hsbMonth_Change()
    New_Month = hsbMonth.Value + 1
    Generate
End Sub

Public Sub Generate()
    lblYD.Caption = month_name(New_Month) & " " & New_Year
End Sub

Private Sub vsbYear_Change(Index As Integer)
    Dim Digit(0 To 3) As Integer
    Dim n As Integer
    
    For n = 0 To 3
      Digit(n) = vsbYear(n).Value
    Next n
    
    New_Year = Digit(0) * 1000 + Digit(1) * 100 + Digit(2) * 10 + Digit(3)
    
    Generate
End Sub
