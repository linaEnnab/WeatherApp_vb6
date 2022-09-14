VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChangeInterval 
      Caption         =   "Change"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Interval 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Text            =   "5"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Timer UpdateTimer 
      Left            =   720
      Top             =   3480
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2160
      List            =   "Form1.frx":0013
      TabIndex        =   0
      Text            =   "Select City"
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fetch every "
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Select city"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.Label apiStatus 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label result 
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents weather As weatheApi.weatherClass
Attribute weather.VB_VarHelpID = -1
Dim Sec As Integer

 Dim timerInterval As Integer


Private Sub CallWeatherApi()
city = Combo1.List(Combo1.ListIndex)
weather.GetTempreture (city)
End Sub



Private Sub Combo1_Click()
Me.UpdateTimer.Enabled = True
CallWeatherApi
End Sub

Private Sub Form_Load()
Sec = 5
timerInterval = 5
Me.UpdateTimer.Interval = 1000
Me.UpdateTimer.Enabled = False
Set weather = New weatheApi.weatherClass



End Sub
Private Sub cmdChangeInterval_Click()
'change timer interval

timerInterval = Val(Me.Interval.Text)

End Sub

Private Sub Interval_KeyUp(KeyCode As Integer, Shift As Integer)
If Not IsNumeric(Interval.Text) Then
Me.Interval.Text = timerInterval
End If

End Sub





Private Sub Interval_Validate(Cancel As Boolean)
Cancel = Not IsNumeric(Me.Interval.Text)
End Sub

Private Sub UpdateTimer_Timer()
' call method every N sec

Sec = Sec - 1
If Sec = 0 Then
Me.apiStatus.Caption = "updating..."
Call CallWeatherApi
Sec = timerInterval
End If
End Sub


Private Sub weather_Completed(ByVal result As String)
Me.result.Caption = "Current degree in " & Combo1.List(Combo1.ListIndex) & "  is:  " & result
Me.apiStatus.Caption = ""
End Sub
