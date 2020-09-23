VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form2 
   Caption         =   "PSC Search minibrowser"
   ClientHeight    =   6570
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9390
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6570
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txturl 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9375
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   120
      Left            =   0
      TabIndex        =   1
      Top             =   6450
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
   End
   Begin SHDocVwCtl.WebBrowser wb1 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      ExtentX         =   14208
      ExtentY         =   11668
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
   Begin VB.Menu mnuback 
      Caption         =   "Back"
   End
   Begin VB.Menu mnuforward 
      Caption         =   "Forward"
   End
   Begin VB.Menu mnurefresh 
      Caption         =   "Refresh"
   End
   Begin VB.Menu mnunewwindow 
      Caption         =   "PopUps"
      WindowList      =   -1  'True
      Begin VB.Menu mnuBlockpopups 
         Caption         =   "Block"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAllowPopUps 
         Caption         =   "Allow"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
txturl.Top = 0
txturl.Left = 0
txturl.Width = ScaleWidth
wb1.Top = txturl.Height
wb1.Left = 0
wb1.Width = ScaleWidth
wb1.Height = ScaleHeight - (ProgressBar1.Height + txturl.Height)
End Sub

Private Sub mnuAllowPopUps_Click()
mnuBlockpopups.Checked = False
mnuAllowPopUps.Checked = True
End Sub

Private Sub mnuBlockpopups_Click()
mnuBlockpopups.Checked = True
mnuAllowPopUps.Checked = False
End Sub
Private Sub mnuback_Click()
On Error Resume Next
wb1.GoBack
End Sub
Private Sub mnuforward_Click()
On Error Resume Next
wb1.GoForward
End Sub

Private Sub mnurefresh_Click()
On Error Resume Next
wb1.Refresh
End Sub

Private Sub txturl_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Debug.Print KeyCode
If KeyCode = 13 Then
    wb1.Navigate txturl
    KeyCode = 0
End If
End Sub

Private Sub wb1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
ProgressBar1 = 0
txturl = wb1.Document.URL
End Sub

Private Sub wb1_NewWindow2(ppDisp As Object, Cancel As Boolean)
If mnuBlockpopups.Checked = True Then
Cancel = True
End If
End Sub

Private Sub wb1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
ProgressBar1.Max = ProgressMax
ProgressBar1 = Progress

End Sub

