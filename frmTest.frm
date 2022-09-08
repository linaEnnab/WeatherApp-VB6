VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Weather App"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChangeInterval 
      Caption         =   "Change"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Interval 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Text            =   "5"
      Top             =   1200
      Width           =   375
   End
   Begin VB.ComboBox SelectCity 
      BackColor       =   &H8000000A&
      Height          =   315
      ItemData        =   "frmTest.frx":058A
      Left            =   2400
      List            =   "frmTest.frx":059D
      TabIndex        =   1
      Text            =   "Select city"
      Top             =   600
      Width           =   1935
   End
   Begin VB.Timer UpdateTimer 
      Left            =   480
      Top             =   2880
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000A&
      Caption         =   "Interval in seconds "
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Result 
      BackColor       =   &H8000000A&
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label apiStatus 
      BackColor       =   &H8000000A&
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "Select city "
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' VBJSON is a VB6 adaptation of the VBA JSON project at http://code.google.com/p/vba-json/
' Some bugs fixed, speed improvements added for VB6 by Michael Glaser (vbjson@ediy.co.nz)
' BSD Licensed

 Public WithEvents weather As WeatherApiCall.WeatherApi
Attribute weather.VB_VarHelpID = -1
 Dim Sec As Integer
 Dim timerInterval As Integer
 Dim Json As New JsonSerializer.JsonClass

 
Private Sub cmdChangeInterval_Click()
'change timer interval

timerInterval = Val(Me.Interval.Text)

End Sub



Private Sub weather_OnChange(Result As String)
 'api call completed
 
   Set resultObj = Json.parse(Result)
    
   city = Me.SelectCity.List(Me.SelectCity.ListIndex)
   
   Me.Result.Caption = "Current degree in " & city & " is: " & resultObj.Item("main").Item("temp")
   Me.apiStatus.Caption = ""

End Sub



Private Sub SelectCity_Click()
Me.apiStatus.Caption = "updating..."
Call getWeather
Me.UpdateTimer.Enabled = True
Sec = Val(Me.Interval.Text)
End Sub

 Private Sub Form_Load()
 ' initialize timer , weather object
 
Me.UpdateTimer.Enabled = False
Me.UpdateTimer.Interval = 1000
timerInterval = Val(Me.Interval.Text)
Set weather = New WeatherApiCall.WeatherApi
End Sub
 
Private Function getWeather()
'call api

   Dim weatherStr As String
   Dim city As String
   
   city = Me.SelectCity.List(Me.SelectCity.ListIndex)
   weatherStr = weather.getWeather(city)
 
   
End Function




Private Sub UpdateTimer_Timer()
' call method every N sec

Sec = Sec - 1
If Sec = 0 Then
Me.apiStatus.Caption = "updating..."
Call getWeather
Sec = Val(Me.Interval.Text)

End If


End Sub


