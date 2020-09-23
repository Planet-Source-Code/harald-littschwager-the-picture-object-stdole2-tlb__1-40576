VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "(c) Litschi-Soft"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   3585
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   6
      Top             =   990
      Width           =   645
   End
   Begin VB.CheckBox Check2 
      Caption         =   "&Loop"
      Height          =   195
      Left            =   2160
      TabIndex        =   3
      Top             =   495
      Width           =   1320
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Play/Pause"
      Height          =   375
      Left            =   2160
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   45
      Width           =   1365
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   135
      Top             =   135
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   285
      LargeChange     =   10
      Left            =   45
      Max             =   184
      TabIndex        =   1
      Top             =   1710
      Width           =   2085
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1620
      Left            =   45
      ScaleHeight     =   1560
      ScaleWidth      =   1980
      TabIndex        =   0
      Top             =   45
      Width           =   2040
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      Caption         =   "Framerate"
      Height          =   240
      Left            =   2160
      TabIndex        =   7
      Top             =   1035
      Width           =   690
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fest Einfach
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   1710
      Width           =   1365
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fest Einfach
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   1350
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pic(184) As Picture         'size Picarray

Private Sub Check1_Click()
Timer1.Enabled = IIf(Check1.Value, True, False) 'Play
End Sub

Private Sub Form_Load()
Dim Tmp As Integer              'Counter
    For Tmp = 0 To 184          'Load Pictures
        Set Pic(Tmp) = LoadPicture(App.Path & "\seq\" & Tmp & ".jpg")
    Next

Picture1.Picture = Pic(0)   'show 1 image
Text1.Text = 29.97                 'Set framerate (1000ms / Framerate)=Interval
Timer1.Interval = 1000 / Abs(Text1.Text)
End Sub


Private Sub HScroll1_Change()
    Show_Pic
End Sub

Private Sub HScroll1_Scroll()
    Show_Pic
End Sub

Private Sub Text1_Change()
Timer1.Interval = 1000 / Abs(Text1.Text) 'Set framerate (1000ms / Framerate)=Interval
End Sub

Private Sub Timer1_Timer()
If Check2.Value = 0 And HScroll1.Value = UBound(Pic) Then  'play once
    Check1.Value = False
    Timer1.Enabled = False
End If

If HScroll1.Value = UBound(Pic) Then HScroll1.Value = 0
If Check1.Value = 1 Then HScroll1.Value = HScroll1.Value + 1
End Sub
Sub Show_Pic()
    Picture1.Picture = Pic(HScroll1.Value)                                                  'show Picture
    Label1.Caption = "Frame : " & HScroll1.Value                                            'show position
    Label2.Caption = "Time : " & Format(HScroll1.Value / Abs(Text1.Text), "0.0" & " sec")   'show time
End Sub
