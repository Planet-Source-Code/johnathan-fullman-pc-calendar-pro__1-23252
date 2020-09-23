VERSION 5.00
Begin VB.Form fAlert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Event Timer"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   2925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Completely Exit"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ComboBox cboTitle 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   2895
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close Window"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblDes 
      Alignment       =   2  'Center
      Caption         =   "Des"
      Height          =   1455
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblEvents 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   2895
   End
End
Attribute VB_Name = "fAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboTitle_Change()
    lblDes(cboTitle.ListIndex).Visible = True
    lblDes(cboTitle.ListIndex).ZOrder
End Sub

Private Sub cboTitle_Click()
    lblDes(cboTitle.ListIndex).Visible = True
    lblDes(cboTitle.ListIndex).ZOrder
End Sub

Private Sub cmdClose_Click()
    Unload fAlert
    Load frm
End Sub

Private Sub cmdEnd_Click()
    SaveSetting "PC Calendar", "Options", "LA", Now
    End
End Sub

Private Sub Form_Load()
    Dim A As Integer
    
    FirstRound = True
    For A = 0 To UBound(FiredEvents)
        cboTitle.AddItem Events(FiredEvents(A)).Title
        
        If A > 0 Then Load lblDes(A)
        lblDes(A).Caption = GetPrelude(Events(FiredEvents(A))) & Chr(13) & Chr(13) & Events(FiredEvents(A)).description
        lblDes(A).Visible = True
    Next A
    
    cboTitle.ListIndex = 0
    lblDes(0).ZOrder
    
    Me.lblEvents.Caption = UBound(FiredEvents) + 1 & " event(s)"
    Me.Visible = True
End Sub

Private Function GetPrelude(Evnt As Evt)
    Select Case DaysAway(Evnt)
        Case 0
            GetPrelude = "Happening Today."
        Case 1
            GetPrelude = "Happening Tomorrow."
        Case Else
            GetPrelude = "Happening in " & DaysAway(Evnt) & " days."
    End Select
End Function
