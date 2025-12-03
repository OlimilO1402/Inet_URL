VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13725
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   13725
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnGetMyIP 
      Caption         =   "GetMyIP"
      Height          =   375
      Left            =   8760
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton BtnReadURL 
      Caption         =   "Read"
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox TxtURL 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3615
      Left            =   0
      TabIndex        =   4
      Top             =   375
      Width           =   13575
      ExtentX         =   23945
      ExtentY         =   6376
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox TxtHTML 
      Height          =   3615
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   3960
      Width           =   13575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'TxtURL.Text = "http://checkip.dyndns.org"
    'TxtURL.Text = "http://www.activevb.de/rubriken/apikatalog/deklarationen/internetclosehandle.html"
    TxtURL.Text = "http://foren.activevb.de/forum/vb-classic/"
    WebBrowser1.Navigate2 "about:blank"
End Sub

Private Sub BtnReadURL_Click()
    Dim u As InternetURL: Set u = MNew.InternetURL(TxtURL.Text)
    Dim s As String: s = u.Read
    TxtHTML.Text = s
    WebBrowser1.Document.body.innerHTML = s
    WriteHtmlfile s
End Sub

Private Sub WriteHtmlfile(sContent As String)
    Dim FNr As Integer: FNr = FreeFile
    Dim FNm As String:  FNm = "C:\TestDir\AVB.html"
Try: On Error GoTo Finally
    Open FNm For Binary Access Write As FNr
    Put FNr, , sContent
Finally:
    Close FNr
End Sub

Private Sub BtnGetMyIP_Click()
    BtnGetMyIP.Caption = GetMyIP
End Sub

Private Function GetMyIP() As String
    Dim s As String: s = MNew.InternetURL("http://checkip.dyndns.org").Read
    Dim i As Long:   i = InStr(1, s, "IP Address: "): If i = 0 Then Exit Function
    Dim L As Long:   L = InStr(1, s, "</body>"):      If L = 0 Then Exit Function
    i = i + 12:    L = L - i
    GetMyIP = Mid(s, i, L)
End Function

Private Sub Form_Resize()
    Dim L As Single: 'L = TxtHTML.Left
    Dim t As Single: t = WebBrowser1.Top
    Dim W As Single: W = Me.ScaleWidth - L
    Dim H As Single: H = (Me.ScaleHeight - t) / 2
    If W > 0 And H > 0 Then WebBrowser1.Move L, t, W, H
    t = t + H
    If W > 0 And H > 0 Then TxtHTML.Move L, t, W, H
End Sub
