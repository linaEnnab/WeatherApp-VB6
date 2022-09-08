VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeatherApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Event OnChange(result As String)

Public Function getWeather(city As String) As String

  Set http = CreateObject("WinHttp.WinHttprequest.5.1")
  URL = "http://api.openweathermap.org/data/2.5/weather?q=" & city & "&appid=fa45908ebb85a1177ddac6c466c4b87b"
  http.Open "Get", URL, False
  http.send
  
  Dim TimeStart As Currency
Dim TimeStop As Currency
Dim TimePassed As Currency
Dim TimeWait As Currency

  RaiseEvent OnChange(http.responseText)

End Function



