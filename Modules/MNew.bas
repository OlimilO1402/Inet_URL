Attribute VB_Name = "MNew"
Option Explicit

Public Function InternetURL(ByVal sURL As String) As InternetURL
    Set InternetURL = New InternetURL: InternetURL.New_ sURL
End Function
