VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmLKDB 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Lichtenknecker Database of the BAV"
   ClientHeight    =   11100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17910
   Icon            =   "frmLKDB.frx":0000
   LinkMode        =   1  'Quelle
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11100
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   10000
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdBeobList_close 
      Caption         =   "Liste schließen"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin SHDocVwCtl.WebBrowser Browser2 
      Height          =   1695
      Left            =   1560
      TabIndex        =   5
      Top             =   6840
      Visible         =   0   'False
      Width           =   10815
      ExtentX         =   19076
      ExtentY         =   2990
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
      Location        =   ""
   End
   Begin VB.CommandButton cmdPicSave 
      Caption         =   "Bild speichern"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdBeobList 
      Caption         =   "Beobachtungsliste"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdForw 
      Caption         =   ">>"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin SHDocVwCtl.WebBrowser wbrWebBrowser 
      Height          =   4575
      Left            =   1791
      TabIndex        =   0
      Top             =   233
      Width           =   5415
      ExtentX         =   9551
      ExtentY         =   8070
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
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
      Location        =   ""
   End
End
Attribute VB_Name = "frmLKDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fs As FileSystemObject
Dim einstrom As TextStream
Dim webready As Boolean
Dim star As String
Dim Gifname As String
Dim ListName As String
Dim result



Private Sub cmdBack_Click()
On Error Resume Next
cmdBeobList_close_Click
wbrWebBrowser.GoBack
End Sub

Private Sub cmdBeobList_Click()
Me.MousePointer = 11
'If fs.FileExists(App.Path & "\BeobListe.txt") Then fs.DeleteFile (App.Path & "\BeobListe.txt")
Browser2.Navigate ListName
Browser2.Visible = True
Form_Resize
cmdBeobList_close.Visible = True
cmdBeobList.Visible = False
 'result = URLDownloadToFile(0, Trim(ListName), _
  '  App.Path & "\BeobListe.txt", 0, 0)
  '  Shell ("notepad " & App.Path & "\BeobListe.txt"), vbNormalFocus
    Me.MousePointer = 1
End Sub

Private Sub cmdBeobList_close_Click()
Me.MousePointer = 11
Browser2.Visible = False
Form_Resize
cmdBeobList_close.Visible = False
cmdBeobList.Visible = True
     Me.MousePointer = 1
End Sub

Private Sub cmdForw_Click()
On Error Resume Next
wbrWebBrowser.GoForward
End Sub

Private Sub cmdPicSave_Click()
Me.MousePointer = 11
If fs.FileExists(App.Path & "\" & star & ".gif") Then fs.DeleteFile (App.Path & "\" & star & ".gif")
 result = URLDownloadToFile(0, Trim(Gifname), _
    App.Path & "\" & star & ".gif", 0, 0)
    Me.MousePointer = 1
End Sub



Public Sub Form_Load()
result = CheckInetConnection(Me.hwnd)
If result = False Then Exit Sub

Me.WindowState = vbNormal
'star = Trim(frmSterninfo.lblStern.Caption)
star = "AB And"
webready = False

On Error Resume Next
wbrWebBrowser.Navigate "http://www.bav-astro.de/LkDB/index.html"
Do While Not webready
DoEvents
Loop
   wbrWebBrowser.Document.Forms(0).stern.Value = star
   webready = False
   wbrWebBrowser.Document.Forms(0).submit.Click
   
   Do While Not webready
    DoEvents
    Loop
   Call quelltextsuche(App.Path & "\quelltext.txt")
 
End Sub

Private Sub Form_Resize()
On Error Resume Next
wbrWebBrowser.Width = Me.ScaleWidth - 1000
wbrWebBrowser.Height = Me.ScaleHeight - 500
Browser2.Width = wbrWebBrowser.Width
Browser2.Left = wbrWebBrowser.Left
Browser2.Top = wbrWebBrowser.Top + wbrWebBrowser.Height - Browser2.Height
If Browser2.Visible = True Then wbrWebBrowser.Height = wbrWebBrowser.Height - Browser2.Height

End Sub

Private Sub wbrWebBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
webready = True
End Sub


Private Sub wbrWebBrowser_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
'MsgBox "Es ist ein Fehler aufgetreten. EventuellLeider ist diese Option nicht verfübgar.."
End Sub

Sub quelltextsuche(ByVal sFilename As String)
Dim F As Integer

Dim zeile As String
Dim searchstring
Dim pos(1) As Integer

searchstring = "http://www.bavdata-astro.de/~tl/vs_out/"
Set fs = New FileSystemObject
  

  With wbrWebBrowser.Document.documentElement
    F = FreeFile
    Open sFilename For Output As #F
    Print #F, .outerHTML;
    Close #F
  End With
  
  Set einstrom = fs.OpenTextFile(sFilename)
  
  While Not einstrom.AtEndOfStream
  zeile = einstrom.ReadLine
  
  If InStr(1, zeile, searchstring, vbTextCompare) <> 0 Then
  pos(0) = InStr(1, zeile, searchstring, vbTextCompare)
  pos(1) = InStr(pos(0), zeile, "method=post", vbTextCompare)
  If pos(1) = 0 Then
    pos(1) = InStr(pos(0), zeile, Chr(34))
    Gifname = Mid(zeile, pos(0), pos(1) - pos(0))
    Else
    ListName = Mid(zeile, pos(0), pos(1) - pos(0))
  End If
  
 
  End If
  
  Wend
  
  einstrom.Close
  Set einstrom = Nothing
  If fs.FileExists(sFilename) Then fs.DeleteFile (sFilename)
  Set fs = Nothing
  
End Sub
