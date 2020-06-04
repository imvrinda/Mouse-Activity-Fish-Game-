VERSION 5.00
Begin VB.Form frmGameFish 
   BackColor       =   &H80000009&
   Caption         =   "Mouse Activity (Developed By Shah Infotech Students)"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form6"
   Picture         =   "game.frx":0000
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3720
      TabIndex        =   6
      Text            =   "Level"
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton cmdres 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   5
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Timer Timer4 
      Left            =   7440
      Top             =   7440
   End
   Begin VB.Timer Timer3 
      Left            =   6960
      Top             =   7440
   End
   Begin VB.Timer Timer2 
      Left            =   6360
      Top             =   7440
   End
   Begin VB.Timer Timer1 
      Left            =   5760
      Top             =   7440
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "SCORE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      TabIndex        =   2
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   585
      Left            =   3960
      Picture         =   "game.frx":1B942
      Top             =   8760
      Width           =   990
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   1320
      Picture         =   "game.frx":1C298
      Top             =   3960
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Image Image5 
      Height          =   570
      Left            =   11760
      Picture         =   "game.frx":1CB78
      Top             =   5640
      Width           =   1005
   End
   Begin VB.Image Image4 
      Height          =   675
      Left            =   5760
      Picture         =   "game.frx":1D3C8
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   960
      TabIndex        =   0
      Top             =   7560
      Width           =   15
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   120
      Picture         =   "game.frx":1DC59
      Top             =   240
      Width           =   1005
   End
End
Attribute VB_Name = "frmGameFish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim inter As Integer

Private Sub cmdres_Click()

Timer1.Interval = 200
Timer2.Interval = 200
Timer3.Interval = 1000
Timer4.Interval = 2000
cmdres.Visible = False
Combo1.Visible = False
If Combo1.Text = "Level1" Then
    inter = 200
ElseIf Combo1.Text = "Level2" Then
    inter = 300
ElseIf Combo1.Text = "Level3" Then
    inter = 400
ElseIf Combo1.Text = "Level4" Then
    inter = 500
ElseIf Combo1.Text = "Level5" Then
    inter = 600


End If

    
    
    
End Sub

'Private Sub Combo1_Change()
'If Combo1.Text = "Level1" Then
'    inter = 500
'ElseIf Combo1.Text = "Level2" Then
'    inter = 400
'ElseIf Combo1.Text = "Level3" Then
'    inter = 300
'End If
'
'End Sub

Private Sub Form_Load()
Timer1.Interval = 0
Timer2.Interval = 0
Timer3.Interval = 0
Timer4.Interval = 0
Combo1.Text = "Level1"
Combo1.AddItem ("Level1")
Combo1.AddItem ("Level2")
Combo1.AddItem ("Level3")
Combo1.AddItem ("Level4")
Combo1.AddItem ("Level5")




End Sub

Private Sub Image1_Click()
Label2.Caption = Val(Label2.Caption) + 1
Image1.Visible = False
End Sub

Private Sub Image2_Click()
Label2.Caption = Val(Label2.Caption) + 1
Image2.Visible = False
End Sub

Private Sub Image3_Click()
Label2.Caption = Val(Label2.Caption) + 1
Image3.Visible = False
End Sub

Private Sub Image4_Click()
Label2.Caption = Val(Label2.Caption) + 1
Image4.Visible = False
End Sub

Private Sub Image5_Click()
Label2.Caption = Val(Label2.Caption) + 1
Image5.Visible = False
End Sub


Private Sub Timer1_Timer()
If Image1.Left > Me.Width Then
Image1.Left = 0 - Image1.Width
Else
 Image1.Left = Image1.Left + inter
 End If
 If Image4.Left > Me.Width Then
Image4.Left = 0 - Image4.Width
Else
 Image4.Left = Image4.Left + inter
 
 End If
 If Image5.Left > Me.Width Then
Image5.Left = 0 - Image5.Width
Else
 Image5.Left = Image5.Left + inter
 
 End If
End Sub
 


Private Sub Timer2_Timer()

If Image2.Left + Image2.Width < Me.Left Then
Image2.Left = Me.Width

Else
 Image2.Left = Image2.Left - inter
 End If
 
 
If Image3.Left + Image3.Width < Me.Left Then
    Image3.Left = Me.Width
Else
    Image3.Left = Image3.Left - inter
End If
End Sub

Private Sub Timer3_Timer()
Label3.Caption = Val(Label3.Caption) + 1
If Val(Label3.Caption) <= 60 And Val(Label2.Caption) >= 50 Then
   
   Label3.Caption = 0
  
   Label2.Caption = 0
   Timer3.Interval = 0
           MsgBox "Congratulation!! You Have Won"
           cmdres.Visible = True
           Combo1.Visible = True
        
ElseIf Val(Label3.Caption) >= 60 And Val(Label2.Caption) < 50 Then
        MsgBox "You Have Lose"
        MsgBox "Your Score is = " + Label2.Caption
        Label3.Caption = 0
        Label2.Caption = 0
        Timer3.Interval = 0
        cmdres.Visible = True
        Combo1.Visible = True
        
End If

End Sub

Private Sub Timer4_Timer()
Image1.Visible = True
Image2.Visible = True
Image3.Visible = True
Image4.Visible = True
Image5.Visible = True
End Sub
