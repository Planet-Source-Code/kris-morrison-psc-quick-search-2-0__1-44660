VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "PSC Quick Search"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11205
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   542
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   747
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicSearch 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7320
      Picture         =   "Form1.frx":27A2
      ScaleHeight     =   285
      ScaleWidth      =   1440
      TabIndex        =   9
      Top             =   120
      Width           =   1440
      Begin VB.Label LblSearch 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Quick Search"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   30
         Width           =   975
      End
   End
   Begin VB.PictureBox PicOver 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7680
      Picture         =   "Form1.frx":3D44
      ScaleHeight     =   270
      ScaleWidth      =   1425
      TabIndex        =   8
      Top             =   6480
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.PictureBox PicDown 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7680
      Picture         =   "Form1.frx":51C6
      ScaleHeight     =   270
      ScaleWidth      =   1425
      TabIndex        =   7
      Top             =   6480
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.PictureBox Picnormal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7680
      Picture         =   "Form1.frx":6648
      ScaleHeight     =   270
      ScaleWidth      =   1425
      TabIndex        =   6
      Top             =   6480
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":7ACA
      Left            =   4680
      List            =   "Form1.frx":7AEF
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   135
      Width           =   2565
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   7995
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   11145
      TabIndex        =   2
      Top             =   7395
      Width           =   11205
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   11145
      TabIndex        =   1
      Top             =   7695
      Width           =   11205
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10320
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7B5E
            Key             =   "b0"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7EB2
            Key             =   "updated"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":820E
            Key             =   "b.5"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8562
            Key             =   "b1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":88B6
            Key             =   "b1.5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8C0A
            Key             =   "b2"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8F5E
            Key             =   "b2.5"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":92B2
            Key             =   "b3"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9606
            Key             =   "b3.5"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":995A
            Key             =   "b4"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9CAE
            Key             =   "b4.5"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A002
            Key             =   "b5"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A356
            Key             =   "prog"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A4C2
            Key             =   "vote"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A816
            Key             =   "novote"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2760
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   4868
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "ProgramURL"
         Text            =   "ProgramURL"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Compatability"
         Text            =   "Compatability"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Level"
         Text            =   "Level"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Views"
         Text            =   "Views"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "NewViews"
         Text            =   "NewViews"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "DateAdded"
         Text            =   "DateAdded"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "Rating"
         Text            =   "Rating"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "Votes"
         Text            =   "Votes"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "NewVotes"
         Text            =   "NewVotes"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Key             =   "ExcellentVotes"
         Text            =   "ExcellentVotes"
         Object.Width           =   38100
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10320
      Top             =   4680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Search for:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "In language:"
      Height          =   195
      Left            =   3360
      TabIndex        =   11
      Top             =   135
      Width           =   885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents wsc1 As DGSwsHTTP
Attribute wsc1.VB_VarHelpID = -1
Dim timercount As Long, totalvotes As Long, totalRating As Double, totalViews As Long, totalGlobes As Long, totalHalfGlobes As Long, PrevVotes As Long, PrevViews As Long
Dim rateicon As String, voteicon As String
Dim rx1 As RegX
Dim dicPrevVote As Dictionary
Dim dicPrevView As Dictionary

Private Sub Form_Load()
Set rx1 = New RegX
Set dicPrevVote = New Dictionary
Set dicPrevView = New Dictionary
' set alarm so I can wake up in the morning

loadSettings
Set wsc1 = New DGSwsHTTP
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
saveSettings
End Sub
Private Sub Form_Resize()
If Me.WindowState = 1 Or 0 Then Exit Sub
If Me.Width < 8500 Then Me.Width = 8500
If Me.Height < 3000 Then Me.Height = 3000
Me.ListView1.Left = 0
Me.ListView1.Height = Me.ScaleHeight - Me.ListView1.Top - Me.Picture1.Height - Me.Picture2.Height - Me.ProgressBar1.Height
Me.ListView1.Width = Me.ScaleWidth
'Me.txtauthorurl.Left = 0
'Me.txtauthorurl.Width = ScaleWidth
'Me.cmdmonitor.Left = ScaleWidth - cmdmonitor.Width
End Sub
Sub status(strmsg As String)
    Picture1.Cls
    Picture1.Print strmsg
End Sub
Sub status2(strmsg As String)
    Picture2.Cls
    Picture2.Print strmsg
End Sub
Sub loadSettings()
On Error Resume Next
Me.Width = GetIni("PSCsearch.ini", "settings", "formwidth", Me.Width)
Me.Height = GetIni("PSCsearch.ini", "settings", "formheight", Me.Height)
PrevViews = GetIni("PSCsearch.ini", "settings", "prevviews", 0)
PrevVotes = GetIni("PSCsearch.ini", "settings", "prevvotes", 0)
Me.WindowState = GetIni("PSCsearch.ini", "settings", "windowstate", Me.WindowState)

'load list view settings
Dim colhead As ColumnHeader
For Each colhead In Me.ListView1.ColumnHeaders
   colhead.Width = GetIni("PSCsearch.ini", "ListView1-" & colhead.Key, "Width", "90")
   colhead.Position = GetIni("PSCsearch.ini", "ListView1-" & colhead.Key, "Position", colhead.Position)
Next

'load dictionaries
Dim x As Long
Dim count As String
Dim pscitem As String
Dim psckey As String
count = GetIni("PSCsearch.ini", "dicPrevView", "count", "0")
For x = 0 To count - 1
   psckey = GetIni("PSCsearch.ini", "dicPrevView", "Key" & x, "")
   pscitem = GetIni("PSCsearch.ini", "dicPrevView", "Item" & x, "0")
   dicPrevView.Add psckey, pscitem
   Debug.Print psckey & " " & pscitem
Next x

count = GetIni("PSCsearch.ini", "dicPrevVote", "count", "0")
For x = 0 To count - 1
   psckey = GetIni("PSCsearch.ini", "dicPrevVote", "Key" & x, "")
   pscitem = GetIni("PSCsearch.ini", "dicPrevVote", "Item" & x, "0")
   dicPrevVote.Add psckey, pscitem
Next x
End Sub
Sub saveSettings()
' save settings for form size
PutIni "PSCsearch.ini", "settings", "formwidth", Me.Width
PutIni "PSCsearch.ini", "settings", "formheight", Me.Height
PutIni "PSCsearch.ini", "settings", "windowstate", Me.WindowState
PutIni "PSCsearch.ini", "settings", "prevviews", CStr(PrevViews)
PutIni "PSCsearch.ini", "settings", "prevvotes", CStr(PrevVotes)
' Save listview settings
Dim colhead As ColumnHeader
For Each colhead In Me.ListView1.ColumnHeaders
    PutIni "PSCsearch.ini", "ListView1-" & colhead.Key, "Width", colhead.Width
    PutIni "PSCsearch.ini", "ListView1-" & colhead.Key, "Position", colhead.Position
Next

End Sub



Sub parsepage()
   On Error Resume Next
    ListView1.ListItems.Clear
    totalViews = 0
    totalRating = 0
    totalvotes = 0
    totalGlobes = 0
    totalHalfGlobes = 0
status "parsing page" ' probably wont see this, because it parses so fast!!!

Dim tmpstr As String
Dim x As Long

' get a matches collection for programs
Dim programs As MatchCollection
Dim program As String

Set programs = rx1.RegX(wsc1.filedata, "<!--descrip-->[\w\W]*?<!--compat-->", True, False, False)

' get a matches collection for ratings
Dim ratings As MatchCollection
Set ratings = rx1.RegX(wsc1.filedata, "<!--user rating-->[\w\W]*?<!description>", True, False, False)
Dim ratingdetails As String
Dim rating As String

' get a matches collection for compatabilities
Dim compatmodes As MatchCollection
Dim compatmode As String
Set compatmodes = rx1.RegX(wsc1.filedata, "<!--code compat-->[\w\W]*?<!--level/author-->", True, False, False)

' get a matches collection for levels
Dim levels As MatchCollection
Dim level As String
Set levels = rx1.RegX(wsc1.filedata, "<!--level-->[\w\W]*?/", True, False, False)

' get a matches collection for views
Dim views As MatchCollection
Dim view As String
Set views = rx1.RegX(wsc1.filedata, "<!--views/date submitted-->[\w\W]*?<!--user rating-->", True, False, False)
Dim dateadded As String
' Now we have each item we want to report on in seperate match collections
' So all we have to do now is loop through them, do a little cleanup and add them to the list

For x = 0 To programs.count - 1
Dim lstitm As ListItem
'cleanup programs and add to list
    program = rx1.HTMLtag.Replace(programs(x), "")
    program = rx1.LeadingWhitespace.Replace(program, "")
    program = rx1.TrailingWhitespace.Replace(program, "")
    Dim programurl As String
    '<A href="http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=29426&amp;lngWId=1">ezUndoRedo</A>
    programurl = rx1.strSubmatch(programs(x), "<A HREF=([^>]+)>" & rx1.Replace(program, "([\(\)])", "\$1", True, True, False) & "</A>", 0, False, False, False)
    programurl = "http://www.planet-source-code.com" & programurl
    ' add to listview
     Set lstitm = ListView1.ListItems.Add(, program, program, "prog", "prog")
         lstitm.SubItems(1) = programurl
'cleanup compatmodes and add to list
    compatmode = rx1.HTMLtag.Replace(compatmodes(x), "")
    compatmode = rx1.HTMLchar.Replace(compatmode, "")
        lstitm.SubItems(2) = compatmode


'cleanup levels and add to list
    level = rx1.HTMLtag.Replace(levels(x), "")
    level = rx1.Replace(level, "/", "", True, True, False)
    level = rx1.HTMLchar.Replace(level, "")
        lstitm.SubItems(3) = level
        
'cleanup views/dateadded and add to list
    view = rx1.HTMLtag.Replace(views(x), "")
    dateadded = rx1.strRegx(view, "[\d\/]+\s[\d:]+\s\w\w", 0, True, True, False)
    lstitm.SubItems(6) = dateadded
    
    view = rx1.strRegx(view, "^\d*", 0, True, False, False)
' Remember view count
    totalViews = totalViews + view
        lstitm.SubItems(4) = view
        
        Dim newprogviews As Long
        newprogviews = view - CLng(dicPrevView(program))
        If newprogviews = view Then newprogviews = 0 ' we have prevviews info
            lstitm.SubItems(5) = newprogviews
            If newprogviews > 0 Then
                lstitm.SmallIcon = "updated"
                lstitm.ListSubItems(4).Bold = True
                lstitm.ListSubItems(4).ForeColor = &H8000&
                lstitm.ListSubItems(5).Bold = True
                lstitm.ListSubItems(5).ForeColor = &H8000&
            Else
                lstitm.ListSubItems(4).Bold = False
                lstitm.ListSubItems(4).ForeColor = vbBlack
                lstitm.ListSubItems(5).Bold = False
                lstitm.ListSubItems(5).ForeColor = vbBlack
        End If



' Get matches for the number of globes and full globes
Dim globe As MatchCollection
Set globe = rx1.RegX(ratings(x), "RatingSmall.jpg", True, True, False)
Dim halfglobe As MatchCollection
Set halfglobe = rx1.RegX(ratings(x), "RatingHalfSmall.jpg", True, False, False)
rating = (globe.count) + (halfglobe.count * 0.5)
totalRating = totalRating + rating
totalGlobes = totalGlobes + globe.count
totalHalfGlobes = totalHalfGlobes + halfglobe.count
rateicon = "b" & rating

    lstitm.SubItems(7) = rating
    lstitm.ListSubItems(7).ReportIcon = rateicon
' get votes out of rating details
Dim vote As String
    vote = rx1.strSubmatch(ratings(x), "By\s(\d+)\sUsers", 0, False, False, False)
    If vote & "" = "" Then vote = "0"
    If vote = 0 Then voteicon = "novote" Else: voteicon = "vote"
    lstitm.SubItems(8) = vote
    lstitm.ListSubItems(8).ReportIcon = voteicon
    totalvotes = totalvotes + vote
    
        Dim newprogvotes As Long
        newprogvotes = vote - CLng(dicPrevVote(program))
        If newprogvotes = vote Then newprogvotes = 0 ' we have prevVotes info
            lstitm.SubItems(9) = newprogvotes
            If newprogvotes > 0 Then
                lstitm.SmallIcon = "updated"
                lstitm.ListSubItems(8).Bold = True
                lstitm.ListSubItems(8).ForeColor = &H8000&
                lstitm.ListSubItems(9).Bold = True
                lstitm.ListSubItems(9).ForeColor = &H8000&
            Else
                lstitm.ListSubItems(8).Bold = False
                lstitm.ListSubItems(8).ForeColor = vbBlack
                lstitm.ListSubItems(9).Bold = False
                lstitm.ListSubItems(9).ForeColor = vbBlack
            End If
            
'cleanup ratingdetails and add to list
Dim excellentratings As String
    excellentratings = rx1.strSubmatch(ratings(x), "(\d*)\s+Excellent\s+Ratings", 0, False, False, False)
    If excellentratings & "" = "" Then excellentratings = "0"
    lstitm.SubItems(10) = excellentratings
Next

If programs.count < 0 Then
    status "Couldn't parse page, either this is not a PSCprograms list page, or PSC had a major revamp"
End If
Dim newviews As Long
Dim newvotes As Long
newviews = totalViews - dicPrevView("pscmon-totalviews")
newvotes = totalvotes - dicPrevVote("pscmon-totalvotes")

status2 "[New] Views:" & newviews & "  Votes:" & newvotes & "   [Totals] Programs:" & programs.count & "  Views:" & totalViews & "  Votes:" & totalvotes & "  Rating:" & totalRating & "  Globes:" & totalGlobes & "  HalfGlobes:" & totalHalfGlobes
If newvotes > 0 Or newviews > 0 Then Beep
    ' should probably add use configurable sound options,
    ' and allow different sound for viewsc1hanged, voteschanged etc
    ' maybe in the next release, for now just

' save prev counts to dictionary
Dim tmpitem As ListItem
dicPrevView.RemoveAll
dicPrevVote.RemoveAll
    For Each tmpitem In ListView1.ListItems
        dicPrevVote.Add tmpitem.Text, tmpitem.SubItems(8)
        dicPrevView.Add tmpitem.Text, tmpitem.SubItems(4)
    Next
' save prev totals to dictionary (use obscure name so it doesn't conflic with users program name"
        dicPrevView.Add "pscmon-totalviews", totalViews
        dicPrevVote.Add "pscmon-totalvotes", totalvotes
status "done"
Timer1.Enabled = True ' only start timer if page parses

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next

Debug.Print ListView1.SortKey & "  " & ColumnHeader.Index - 1
If ListView1.SortKey = ColumnHeader.Index - 1 Then
    If ListView1.SortOrder = lvwAscending Then
        ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
Else
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.SortOrder = lvwAscending
End If

ListView1.Sorted = True
End Sub

Private Sub ListView1_ItemClick(ByVal item As MSComctlLib.ListItem)
    Load Form2
    Form2.wb1.Navigate item.SubItems(1)
    Form2.Show vbModal

End Sub



Private Sub wsc1_DownloadComplete()
Me.ProgressBar1 = 0
 parsepage
End Sub

Private Sub wsc1_ProgressChanged(ByVal bytesreceived As Long)
On Error Resume Next
'Show current download status
Dim percentcomplete As Long
percentcomplete = (bytesreceived / wsc1.FileSize) * 100
ProgressBar1 = percentcomplete
status "Downloading " & bytesreceived & " byres received " & percentcomplete & "%"
End Sub

Private Sub LblSearch_Click()
Dim String1 As String
ListView1.SetFocus

Select Case Combo1.Text
    
    Case ""
        MsgBox "Please select a Programming language.", vbInformation, App.Title
    
    Case "Visual Basic"
        
        String1 = "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&lngWId=" & "1" & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&txtCriteria=" & Text1.Text
         wsc1.geturl (String1)
         Picture1.Cls
         Picture1.Print "Searching....."
         
         
    Case "Java / Javascript"
         
        String1 = "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&lngWId=" & "2" & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&txtCriteria=" & Text1.Text
        wsc1.geturl (String1)
        Picture1.Cls
        Picture1.Print "Searching....."
        
    Case "C / C++"
        
        String1 = "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&lngWId=" & "3" & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&txtCriteria=" & Text1.Text
        wsc1.geturl (String1)
        Picture1.Cls
        Picture1.Print "Searching....."
        
    Case "ASP / VbScript"
        
        String1 = "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&lngWId=" & "4" & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&txtCriteria=" & Text1.Text
        wsc1.geturl (String1)
        Picture1.Cls
        Picture1.Print "Searching....."
        
    Case "SQL"
        
        String1 = "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&lngWId=" & "5" & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&txtCriteria=" & Text1.Text
        wsc1.geturl (String1)
        Picture1.Cls
        Picture1.Print "Searching....."
        
    Case "Perl"
        
        String1 = "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&lngWId=" & "6" & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&txtCriteria=" & Text1.Text
        wsc1.geturl (String1)
        Picture1.Cls
        Picture1.Print "Searching....."
        
    Case "Delphi"
        
        String1 = "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&lngWId=" & "7" & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&txtCriteria=" & Text1.Text
        wsc1.geturl (String1)
        Picture1.Cls
        Picture1.Print "Searching....."
        
    Case "PHP"
        
        String1 = "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&lngWId=" & "8" & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&txtCriteria=" & Text1.Text
        wsc1.geturl (String1)
        Picture1.Cls
        Picture1.Print "Searching....."
        
    Case "Cold Fusion"
        
        String1 = "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&lngWId=" & "9" & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&txtCriteria=" & Text1.Text
        wsc1.geturl (String1)
        Picture1.Cls
        Picture1.Print "Searching....."
        
    Case ".Net"
        
        String1 = "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&lngWId=" & "10" & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&txtCriteria=" & Text1.Text
        wsc1.geturl (String1)
        Picture1.Cls
        Picture1.Print "Searching....."
        
    Case "LISP"
        
        String1 = "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&lngWId=" & "13" & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&txtCriteria=" & Text1.Text
        wsc1.geturl (String1)
        Picture1.Cls
        Picture1.Print "Searching....."
        
    End Select
    
End Sub


Private Sub Picsearch_Click()
    LblSearch_Click
End Sub

Private Sub Picsearch_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    PicSearch.Picture = PicDown.Picture
End Sub
Private Sub Picsearch_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    PicSearch.Picture = PicOver.Picture
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    PicSearch.Picture = PicDown.Picture
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    PicSearch.Picture = Picnormal.Picture
End Sub

