VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Regx"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public LeadingWhitespace As RegExp
Public TrailingWhitespace As RegExp
Public AllWhitespace As RegExp
Public HTMLtag As RegExp
Public HTMLchar As RegExp
' Define Private Variables
Private regx1 As New RegExp
Private matches1 As MatchCollection
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent regxError[(arg1, arg2, ... , argn)]
Public Event regxError(ByVal Number As Long, ByVal Description As String)
'local variable(s) to hold property value(s)
Private mvarHideErrors As Boolean 'local copy
Public Property Let HideErrors(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HideErrors = 5
    mvarHideErrors = vData
End Property
Public Property Get HideErrors() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HideErrors
    HideErrors = mvarHideErrors
End Property
Public Function RegX(rx_string As String, rx_pattern As String, rx_global As Boolean, rx_ignorecase As Boolean, rx_multiline As Boolean) As MatchCollection
Attribute RegX.VB_Description = "Runs generic regx on the fly and returns matches collection"
On Error GoTo errorhandler
    'Runs generic regx on the fly and returns matches collection
    Set regx1 = New RegExp
    regx1.Pattern = rx_pattern
    regx1.Global = rx_global
    regx1.MultiLine = rx_multiline
    regx1.IgnoreCase = rx_ignorecase
    Set RegX = regx1.Execute(rx_string)
    
    Exit Function
errorhandler:
    If mvarHideErrors = True Then RaiseEvent regxError(Err.Number, Err.Description)
    Resume Next
End Function
Public Function Replace(source_string As String, rx_pattern As String, replace_string As String, rx_global As Boolean, rx_ignorecase As Boolean, rx_multiline As Boolean) As String
Attribute Replace.VB_Description = "Runs generic regx replace on the fly and returns string"
On Error GoTo errorhandler
    'Runs generic regx replace on the fly and returns string
    Set regx1 = New RegExp
    regx1.Pattern = rx_pattern
    regx1.Global = rx_global
    regx1.MultiLine = rx_multiline
    regx1.IgnoreCase = rx_ignorecase
    Replace = regx1.Replace(source_string, replace_string)
    
    Exit Function
errorhandler:
    If mvarHideErrors = True Then RaiseEvent regxError(Err.Number, Err.Description)
    Resume Next
End Function
Public Function strRegx(rx_string As String, rx_pattern As String, rx_matchindex As Long, rx_ignorecase As Boolean, rx_global As Boolean, rx_multiline As Boolean) As String
Attribute strRegx.VB_Description = "Runs generic regx on the fly and returns string"
On Error GoTo errorhandler
    'Runs generic regx on the fly and returns match string
    regx1.Pattern = rx_pattern
    regx1.Global = rx_global
    regx1.MultiLine = rx_multiline
    regx1.IgnoreCase = rx_ignorecase
    Set matches1 = regx1.Execute(rx_string)
    strRegx = matches1(rx_matchindex)
    
    Exit Function
errorhandler:
    If mvarHideErrors = True Then RaiseEvent regxError(Err.Number, Err.Description)
    Resume Next
End Function
Public Function strSubmatch(rx_string As String, rx_pattern As String, rx_submatchindex As Long, rx_ignorecase As Boolean, rx_global As Boolean, rx_multiline As Boolean) As String
Attribute strSubmatch.VB_Description = "Runs regx on the fly and returns submatch as string"
On Error GoTo errorhandler
    'Runs generic regx on the fly and returns submatch string
    regx1.Pattern = rx_pattern
    regx1.Global = rx_global
    regx1.MultiLine = rx_multiline
    regx1.IgnoreCase = rx_ignorecase
    Set matches1 = regx1.Execute(rx_string)
    strSubmatch = matches1(0).SubMatches(rx_submatchindex)
    
    Exit Function
errorhandler:
    If mvarHideErrors = True Then RaiseEvent regxError(Err.Number, Err.Description)
    Resume Next
End Function

Private Sub Class_Initialize()
' regx to remove leading whitespace
    Set LeadingWhitespace = New RegExp
    LeadingWhitespace.Global = False
    LeadingWhitespace.MultiLine = False
    LeadingWhitespace.IgnoreCase = True
    LeadingWhitespace.Pattern = "^\s*"

' regx to remove trailing whitespace
    Set TrailingWhitespace = New RegExp
    TrailingWhitespace.Global = False
    TrailingWhitespace.MultiLine = False
    TrailingWhitespace.IgnoreCase = True
    TrailingWhitespace.Pattern = "\s*$"
    
' regx to remove all whitespace
    Set AllWhitespace = New RegExp
    AllWhitespace.Global = True
    AllWhitespace.MultiLine = True
    AllWhitespace.IgnoreCase = True
    AllWhitespace.Pattern = "\s*"
    
' regx to remove HTML tags
    Set HTMLtag = New RegExp
    HTMLtag.Global = True
    HTMLtag.MultiLine = False
    HTMLtag.IgnoreCase = True
    HTMLtag.Pattern = "<[^>]*?>"

' regx to remove HTML special characters
    Set HTMLchar = New RegExp
    HTMLchar.Global = True
    HTMLchar.MultiLine = False
    HTMLchar.IgnoreCase = True
    HTMLchar.Pattern = "&[a-zA-Z0-9]{2,6};"
End Sub

Private Sub Class_Terminate()
Set regx1 = Nothing
Set matches1 = Nothing
Set LeadingWhitespace = Nothing
Set TrailingWhitespace = Nothing
Set AllWhitespace = Nothing
Set HTMLtag = Nothing
Set HTMLchar = Nothing
End Sub
