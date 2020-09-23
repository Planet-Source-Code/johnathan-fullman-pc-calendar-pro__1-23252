VERSION 5.00
Begin VB.Form fChooseColor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Color"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "fChooseColor.frx":0000
   ScaleHeight     =   3930
   ScaleWidth      =   2625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblSample 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   2415
   End
End
Attribute VB_Name = "fChooseColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const IMG_H = 3045
Dim DefColor As Long
Dim UsedX As Boolean

Public Sub Start(DefualtColor As Long)
    DefColor = DefualtColor
    
    lblSample.BackColor = DefColor
    
    UsedX = True
    Me.Visible = True
End Sub

Private Sub cmdCancel_Click()
    UsedX = False
    ChoosenColor = DefColor
    DoneChoosing = True
    Unload Me
End Sub

Private Sub cmdOk_Click()
    UsedX = False
    ChoosenColor = lblSample.BackColor
    DoneChoosing = True
    Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < IMG_H Then lblSample.BackColor = Me.Point(X, Y)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If UsedX Then
        ChoosenColor = DefColor
        DoneChoosing = True
    End If
End Sub
