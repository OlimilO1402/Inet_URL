VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14295
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   14295
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TxtHTML 
      Height          =   2655
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   10335
   End
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
End Sub

Private Sub BtnReadURL_Click()
    Dim u As InternetURL: Set u = MNew.InternetURL(TxtURL.Text)
    Dim s As String: s = u.Read
    TxtHTML.Text = s
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
    Dim L As Single: L = TxtHTML.Left
    Dim T As Single: T = TxtHTML.Top
    Dim W As Single: W = Me.ScaleWidth - L
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then TxtHTML.Move L, T, W, H
End Sub
