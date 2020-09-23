VERSION 5.00
Begin VB.Form fOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Startup Options"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CheckBox chkShowTips 
      Caption         =   "&Show Tips at Startup"
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   600
      Width           =   2295
      Begin VB.OptionButton Option2 
         Caption         =   "No"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optDouble 
         Caption         =   "Yes"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
      Begin VB.OptionButton optOnce 
         Caption         =   "Ignore them"
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optOnce 
         Caption         =   "Prompt you to delete them"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton optOnce 
         Caption         =   "Delete Them"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1935
      Begin VB.OptionButton Option1 
         Caption         =   "No"
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optPrompt 
         Caption         =   "Yes"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Should PC Calendar warn you of events twice in the same day if you already checked them?"
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "What should PC Calendar do with events that happen only once and have past?"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Should PC Calendar prompt you to quit if no events are triggered at startup?"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "fOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    AnyChanges = True
    CfgChanges = True
    SaveOpts
    
    Unload Me
End Sub

Private Sub Form_Load()
    LoadOpts
    Me.Visible = True
End Sub

Public Sub LoadOpts()
    Dim A As Integer
    
    optPrompt.Value = Prompt
    
    optDouble.Value = DoubleWarning
    
    chkShowTips.Value = IIf(ShowTips, 1, 0)
    
    For A = 0 To 2
        If A = OnceEv Then optOnce(A).Value = True
    Next A
End Sub

Public Sub SaveOpts()
    Dim A As Integer
    
    Prompt = optPrompt.Value
    
    DoubleWarning = optDouble.Value
    
    ShowTips = (chkShowTips.Value = 1)
    
    For A = 0 To 2
        If optOnce(A).Value Then
            Select Case A
                Case 0
                    OnceEv = PE_DELETE
                Case 1
                    OnceEv = PE_ASK
                Case 2
                    OnceEv = PE_IGNORE
            End Select
        End If
    Next A
End Sub
