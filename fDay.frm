VERSION 5.00
Begin VB.Form fDay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Day's Events"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstEvents 
      Height          =   1035
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblNO 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lblDes 
      Height          =   975
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblHeader 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "fDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Start(Feche As String)
    Dim DOW As Integer
    DOW = Weekday(Feche)
    lblHeader.Caption = "The things happening on " & frm.DOW(DOW - 1).Caption & ", " & month_name(Month(Feche)) & " " & day(Feche) & ", " & Year(Feche) & ":"
    
    SetupAll Feche
    If lstEvents.ListCount > 0 Then
        lstEvents.ListIndex = 0
        Me.Visible = True
    Else
        MsgBox "There are no events happening on " & frm.DOW(DOW - 1).Caption & ", " & month_name(Month(Feche)) & " " & day(Feche) & ", " & Year(Feche)
        Unload Me
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub SetupAll(Feche As String)
On Error Resume Next
    Dim A As Integer
    Dim Num As Integer
    
    Num = 0
    For A = 0 To Index
        If EventOnDay(A, Feche) Then
            If Num > 0 Then
                Load lblDes(Num)
                lblDes(Num).Visible = True
                Load lblNO(Num)
                lblNO(Num).Visible = True
            End If

            lstEvents.AddItem Events(A).Title
            lblDes(Num).Caption = Events(A).description
            lblNO(Num).Caption = "Next occurrence is in " & Abs(DateDiff("d", NextOccurance(Events(A)), Today)) & " days."

            Num = Num + 1
        End If
    Next A
End Sub

Private Sub lstEvents_Click()
    If lstEvents.ListCount > 0 Then
        lblDes(lstEvents.ListIndex).ZOrder
        lblNO(lstEvents.ListIndex).ZOrder
    End If
End Sub
