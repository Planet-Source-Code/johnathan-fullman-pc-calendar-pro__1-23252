VERSION 5.00
Begin VB.Form fZEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Event Management"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Copy / Paste Events"
      Height          =   1215
      Left            =   2760
      TabIndex        =   13
      Top             =   1320
      Width           =   3375
      Begin VB.CommandButton cmdInsertPaste 
         Caption         =   "Insert"
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   375
         Left            =   1680
         TabIndex        =   19
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdCut 
         Caption         =   "Cut"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   375
         Left            =   840
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblCopiedEvt 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Event In Clipboard:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.ListBox lstEvents 
      Height          =   4155
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add / Remove Events"
      Height          =   855
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   3375
      Begin VB.CommandButton cmdInsertNew 
         Caption         =   "Insert New"
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add New"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Event Order"
      Height          =   1815
      Left            =   2760
      TabIndex        =   1
      Top             =   2640
      Width           =   3375
      Begin VB.CommandButton cmdMoveDown 
         Caption         =   "Move Down"
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdMoveUp 
         Caption         =   "Move Up"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdSwitch 
         Caption         =   "Switch"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdSelectClear 
         Caption         =   "Select"
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   120
         X2              =   3240
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblEvent2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "will switch with"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblEvent1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Event Selection:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "fZEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SwitchNum1 As Integer
Public SwitchNum2 As Integer
Public Step As Integer

Private Sub cmdAdd_Click()
    fNew.Show
End Sub

Private Sub cmdCopy_Click()
    CopyEvent lstEvents.ListIndex
    CheckClipboard
End Sub

Private Sub cmdCut_Click()
    CutEvent lstEvents.ListIndex
    CheckClipboard
    UpdateList
End Sub

Private Sub cmdInsertNew_Click()
    Inserting = True
    InsertSlot = lstEvents.ListIndex
    fNew.Show
End Sub

Private Sub cmdInsertPaste_Click()
    InsertEvent lstEvents.ListIndex, ClipBoard
    UpdateList lstEvents.ListIndex
End Sub

Private Sub cmdMoveDown_Click()
    Dim NewFocus
    NewFocus = lstEvents.ListIndex + 1
    
    SwitchEvents lstEvents.ListIndex, lstEvents.ListIndex + 1
    UpdateList
    
    lstEvents.ListIndex = NewFocus
End Sub

Private Sub cmdMoveUp_Click()
    Dim NewFocus As Integer
    NewFocus = lstEvents.ListIndex - 1
    
    SwitchEvents lstEvents.ListIndex, lstEvents.ListIndex - 1
    UpdateList
    
    lstEvents.ListIndex = NewFocus
End Sub

Private Sub cmdPaste_Click()
    PasteEvent
    UpdateList Index
End Sub

Private Sub cmdRemove_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Are you sure you wish to delete " & Events(lstEvents.ListIndex).Title & "?", vbYesNo + vbQuestion, "Confirmation")
    If Ans = 6 Then
        RemoveEvent lstEvents.ListIndex
        UpdateList
    End If
End Sub

Private Sub cmdSelectClear_Click()
    Select Case Step
        Case 0 'None Selected
            Step = 1
            SwitchNum1 = lstEvents.ListIndex
            lblEvent1.Caption = Events(SwitchNum1).Title
            cmdSwitch.Enabled = False
            cmdSelectClear.Caption = "Select"
        Case 1 'One Selected
            Step = 2
            SwitchNum2 = lstEvents.ListIndex
            lblEvent2.Caption = Events(SwitchNum2).Title
            cmdSwitch.Enabled = True
            cmdSelectClear.Caption = "Clear"
        Case 2 'Both Selected
            Step = 0
            lblEvent1.Caption = ""
            lblEvent2.Caption = ""
            cmdSwitch.Enabled = False
            cmdSelectClear.Caption = "Select"
    End Select
End Sub

Private Sub cmdSwitch_Click()
    SwitchEvents SwitchNum1, SwitchNum2
    UpdateList
    
    Step = 0
    lblEvent1.Caption = ""
    lblEvent2.Caption = ""
    cmdSwitch.Enabled = False
    cmdSelectClear.Caption = "Select"
End Sub

Private Sub Form_Load()
    UpdateList
    CheckClipboard
    lstEvents.ListIndex = 0
    Step = 0
End Sub

Public Sub UpdateList(Optional NewIndex As Integer)
    Dim A As Integer
    lstEvents.Clear
    For A = 0 To Index
        lstEvents.AddItem Events(A).Title
    Next A
    
    If IsMissing(NewIndex) Then
        lstEvents.ListIndex = 0
    Else
        lstEvents.ListIndex = NewIndex
    End If
    
    frm.ShowNewMonth
End Sub

Private Sub lstEvents_Click()
    Select Case lstEvents.ListIndex
        Case 0
            cmdMoveUp.Enabled = False
            If lstEvents.ListCount = 1 Then
                cmdMoveDown.Enabled = False
            Else
                cmdMoveDown.Enabled = True
            End If
        Case lstEvents.ListCount - 1
            cmdMoveDown.Enabled = False
            If lstEvents.ListCount = 1 Then
                cmdMoveUp.Enabled = False
            Else
                cmdMoveUp.Enabled = True
            End If
        Case Else
            cmdMoveUp.Enabled = True
            cmdMoveDown.Enabled = True
    End Select
End Sub

Public Sub CheckClipboard()
    If AnythingCopied Then
        lblCopiedEvt.Caption = ClipBoard.Title
        cmdPaste.Enabled = True
    Else
        lblCopiedEvt.Caption = "None"
        cmdPaste.Enabled = False
    End If
    
    cmdInsertPaste.Enabled = cmdPaste.Enabled
End Sub
