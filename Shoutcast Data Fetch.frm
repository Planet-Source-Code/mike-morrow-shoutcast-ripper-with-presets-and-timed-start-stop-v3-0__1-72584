VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataFetch 
   Caption         =   "Shoutcast Stations List and Station Selector"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15540
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   15540
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClearSaves 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Clear Saved"
      Height          =   435
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7920
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pbStations 
      Height          =   300
      Left            =   10440
      TabIndex        =   21
      Top             =   5820
      Visible         =   0   'False
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdNextStream 
      Caption         =   "&>"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Move to higher number stream on this station."
      Top             =   5865
      Width           =   495
   End
   Begin VB.CommandButton cmdPreviousStream 
      Caption         =   "&<"
      Enabled         =   0   'False
      Height          =   345
      Left            =   510
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Move to lower number stream on this station."
      Top             =   5865
      Width           =   495
   End
   Begin VB.CommandButton cmdAddNew 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Remember Selected Station"
      Enabled         =   0   'False
      Height          =   435
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Add selected/displayed station to the saved stations list in the next available open row and exit this screen."
      Top             =   7920
      Width           =   2355
   End
   Begin VB.TextBox txtGenreSearch 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6480
      TabIndex        =   0
      ToolTipText     =   "Enter the beginning letters of a Genre then press Enter (case insensitive)."
      Top             =   6900
      Width           =   8955
   End
   Begin VB.TextBox txtStationSearch 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6480
      TabIndex        =   1
      ToolTipText     =   "Enter any part of a station name then press Enter (case insensitive) to search for stations in the displayed list."
      Top             =   7560
      Width           =   8955
   End
   Begin VB.ListBox lstStations 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5160
      Left            =   2100
      Sorted          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "After you search for a Genre, the list of all stations for that Genre will be displayed here."
      Top             =   660
      Width           =   13275
   End
   Begin VB.ListBox lstGenres 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5160
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "A listing of all Shoutcast Genres."
      Top             =   660
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   15060
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FF00FF&
      Cancel          =   -1  'True
      Caption         =   "&We're Done Here!"
      Height          =   435
      Left            =   13920
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "You know!  Saves any Remembered Station(s) in Station Memory(ies)"
      Top             =   7920
      Width           =   1515
   End
   Begin VB.Label lblAddedStationsCt 
      Alignment       =   2  'Center
      Caption         =   "No stations added yet."
      Height          =   255
      Left            =   9120
      TabIndex        =   22
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Label lblFoundStations 
      Height          =   195
      Left            =   6480
      TabIndex        =   20
      Top             =   5880
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.Label lblMultiStream 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   990
      TabIndex        =   17
      ToolTipText     =   "Number of the current stread/Number of streams or indication of streams information error."
      Top             =   5865
      Width           =   4290
   End
   Begin VB.Label txtStationName 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Station Name as supplied by Shoutcast."
      Top             =   6240
      Width           =   6240
   End
   Begin VB.Label lblSearch 
      Caption         =   "Enter any part of a station name then press the Enter key."
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   7260
      Width           =   5595
   End
   Begin VB.Label lblGenreSearch 
      Caption         =   "Enter the beginning letters of a Genre then press the Enter key."
      Height          =   255
      Left            =   6480
      TabIndex        =   13
      Top             =   6600
      Width           =   5595
   End
   Begin VB.Label Label1 
      Caption         =   "List of stations for selected Genre"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2160
      TabIndex        =   12
      Top             =   240
      Width           =   8715
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select a Genre to see all stations"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   60
      Width           =   1695
   End
   Begin VB.Label lblURL 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "IP address and Port of Shoutcast server."
      Top             =   8040
      Width           =   6240
   End
   Begin VB.Label lblListeners 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3360
      TabIndex        =   9
      ToolTipText     =   "Current listener count for this stream."
      Top             =   7380
      Width           =   3000
   End
   Begin VB.Label lblBitRate 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Stream bitrate (in kilobits/second)"
      Top             =   7380
      Width           =   3000
   End
   Begin VB.Label lblID 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Shoutcast internal station ID.  Used to retrieve URL."
      Top             =   7080
      Width           =   3000
   End
   Begin VB.Label lblGenres 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "One of more Genres that this station plays."
      Top             =   7740
      Width           =   6240
   End
   Begin VB.Label lblFormat 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3360
      TabIndex        =   5
      ToolTipText     =   "Stream format.  As of 2009, only mp3 or aac were being streamed."
      Top             =   6720
      Width           =   3000
   End
End
Attribute VB_Name = "frmDataFetch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
  Dim sWholeData As String    ' Winsock puts it here
 'Dim sGenreData As String    ' Copied from sWholeData to process
  Dim sStationData As String  ' Copied from sWholeData to process
  Dim sEndTag As String
  Dim sShoutCastURL As String
  
  Private bDataReceived As Boolean  ' Set to True in Winsock1.DataArrival when the end marker is received.

  Private iFoundStation As Long
  Private iFoundGenre As Long

  Private bIgnoreChange As Boolean
  
  Private iMultiStreamCt As Long
  Private sMultiStreamURL() As String    ' Multistream URLs
  Private sMultiStreamLC() As String     ' Multistream ListenerCounts
  Private sMultiStreamTitle() As String  ' This multistream's title
  Private iMultiShown As Long            ' Which of the multistream infos are showing
  
  Private iNewStation2Add As Long
  
  
Sub FetchGenreInfo()

  Dim ipos As Long
  Dim i As Long
  Dim iEndpos As Long
  Dim tCount As Integer
  Dim p As Long
  Dim j As Long
  Dim sStation As String
  Dim sID As String
  Dim sGenre As String
  Dim sCurrTrack As String
  Dim sBitRate As String
  Dim iSpinCt As Long
  
  Winsock1.RemoteHost = "yp.shoutcast.com"
  Winsock1.RemotePort = 80
  
  sShoutCastURL = "http://yp.shoutcast.com/sbin/newxml.phtml?"
  sEndTag = "</genrelist>"
  sWholeData = ""
  bDataReceived = False
  Winsock1.Connect
  
  iSpinCt = 0
  Do
    Wait 20
    iSpinCt = iSpinCt + 1
    DoEvents
    If iSpinCt > 400 Then
      MsgBox "Waiting too long for Genre list.  Stopping waiting now."
      bDataReceived = True
    End If
  Loop Until bDataReceived = True
  If gbDebugLogic Then Debug.Print MyTime() & "</genrelist> wait spin count: " & iSpinCt
  
  Winsock1.Close

  bDataReceived = False
  
  i = 1  ' Start at the beginning...
  Do
    j = InStr(i, sWholeData, "genre name", vbTextCompare)
    If j > 0 Then
      ipos = InStr(j, sWholeData, sQuote) + 1
      iEndpos = InStr(ipos, sWholeData, sQuote)
      sGenre = Mid$(sWholeData, ipos, iEndpos - ipos)

' from Mike Morrow, the egomaniac...
' NOTE: There is reported problem with the Shoutcast server but they have chosen not to fix it.
'       I tried replacing the "&amp;" with an "&" resulting in "R&b" but when that is sent in,
'       it gets no response.  And that is Shoutcasts way of telling you that it does not like
'       what you are doing.  It would be nice if they would return an error packet but, alas...
'       So to avoid looking like an idiot at Shoutcast's expense and for their error, I skip it.
' NOTE: Shoutcast should have asked me before desigining the system!

     'Skip this known Shoutcast database index (genre) error.
      If sGenre <> "R&amp;b" Then lstGenres.AddItem sGenre     ' Add it to the listbox where it does some good.
      i = iEndpos  ' This is where to start on the next round looking for the next Genre.
    End If
  Loop Until j = 0

End Sub

Sub FindStations()

  If iFoundStation = -1 Then iFoundStation = 0
  
  iFoundStation = SearchList(txtStationSearch, lstStations, iFoundStation + 1)
  
  If iFoundStation <> -1 Then
    lstStations.ListIndex = iFoundStation
  Else
    MsgBox "There are no stations matching " & sQuote & txtStationSearch & sQuote & "."
  End If
  

End Sub

Sub FindGenre()

  cmdAddNew.Enabled = False
  
  txtStationName = "": lblFormat = "": lblID = "": lblBitRate = "": lblGenres = ""
  lblMultiStream = "...": txtStationSearch = ""
  lblListeners = "": lblURL = "": DoEvents
  
  If iFoundGenre = -1 Then iFoundGenre = 0
  
  iFoundGenre = SearchList(txtGenreSearch, lstGenres, iFoundGenre + 1)
  
  If iFoundGenre <> -1 Then
    lstGenres.ListIndex = iFoundGenre
  Else
    MsgBox "There are not more Genre entries matching " & sQuote & txtGenreSearch & sQuote & "."
  End If
  
End Sub

Sub AddNewToArray()

  Dim i As Integer
  
  If txtStationName <> "" Then  ' Be sure something is selected
    
    iNewStation2Add = iNewStation2Add + 1
    
    If iNewStation2Add > 1 Then ReDim Preserve NewShoutcastStation(iNewStation2Add)
    
    NewShoutcastStation(iNewStation2Add).StationName = txtStationName
    If gbDebugLogic Then Debug.Print MyTime() & "Passing Back Station Name: " & NewShoutcastStation(iNewStation2Add).StationName
    
    i = InStr(1, lblFormat, ":")
    NewShoutcastStation(iNewStation2Add).Format = Trim$(Mid$(lblFormat, i + 1))
    
    i = InStr(1, lblID, ":")
    NewShoutcastStation(iNewStation2Add).ID = Trim$(Mid$(lblID, i + 1))
    If gbDebugLogic Then Debug.Print MyTime() & "Passing Back Station ID: " & NewShoutcastStation(iNewStation2Add).ID
    
    i = InStr(1, lblBitRate, ":")
    NewShoutcastStation(iNewStation2Add).BitRate = Trim$(Mid$(lblBitRate, i + 1))
    
    i = InStr(1, lblURL, ":")
    NewShoutcastStation(iNewStation2Add).URL = Trim$(Mid$(lblURL, i + 1))
    
    If gbDebugLogic Then Debug.Print MyTime() & "Listener count: " & lblListeners
    NewShoutcastStation(iNewStation2Add).ListenerCount = lblListeners
    If gbDebugLogic Then Debug.Print MyTime() & "Genre: " & lblGenres
    NewShoutcastStation(iNewStation2Add).Genre = lblGenres
    NewShoutcastStation(iNewStation2Add).StartDate = ""
    NewShoutcastStation(iNewStation2Add).StartHour = ""
    NewShoutcastStation(iNewStation2Add).StartMin = ""
    NewShoutcastStation(iNewStation2Add).Duration = ""
    NewShoutcastStation(iNewStation2Add).CurrentTrack = ""
    NewShoutcastStation(iNewStation2Add).MyFilePrefix = ""
    If iNewStation2Add = 1 Then
      lblAddedStationsCt = "1 station added."
    Else
      lblAddedStationsCt = iNewStation2Add & " stations added."
    End If
   'Unload Me
  
  Else
    MsgBox "No station selected.  Please try again or select We're Done Here."
  End If
  
End Sub

Sub UpdateStreamInfo()

  Dim sStreamName As String
  
  lblURL = sMultiStreamURL(iMultiShown)
  aSplit = Split(sMultiStreamLC(iMultiShown), "/")
  If UBound(aSplit) < 1 Then
    MsgBox "Program Error USI1 in UpdateStreamInfo splitting " & sMultiStreamLC(iMultiShown)
    Exit Sub
  End If
  
  lblListeners = aSplit(0) & " of " & aSplit(1) & " (max) listeners."
   
  sStreamName = sMultiStreamTitle(iMultiShown)
  
 'The Shoutcast database is really ugly with lots of ASCII fluff.  I had just as soon not see all that junque!
 'Here, I filter it out.  And it takes some good filtering.  VB's Replace only makes one pass and this junk takes
 'multiple passes to filter out all the visual distractions so I go to a repeating Replace routine to do it right.
  sStreamName = MultiReplace(sStreamName, "]]", "]")
  sStreamName = MultiReplace(sStreamName, "[[", "[")
  sStreamName = MultiReplace(sStreamName, "^^", "^")
  sStreamName = MultiReplace(sStreamName, "~~", "~")
  sStreamName = MultiReplace(sStreamName, "!!", "!")
  sStreamName = MultiReplace(sStreamName, "((", "(")
  sStreamName = MultiReplace(sStreamName, "))", ")")
  sStreamName = MultiReplace(sStreamName, " -", "")
  sStreamName = MultiReplace(sStreamName, "::", ":")
  sStreamName = MultiReplace(sStreamName, "..", ".")
  sStreamName = MultiReplace(sStreamName, "\", " ")
  sStreamName = MultiReplace(sStreamName, "/", " ")
  sStreamName = MultiReplace(sStreamName, ":", " ")
  sStreamName = MultiReplace(sStreamName, "*", " ")
  sStreamName = MultiReplace(sStreamName, "?", " ")
  sStreamName = MultiReplace(sStreamName, sQuote, " ")
  sStreamName = MultiReplace(sStreamName, "<", " ")
  sStreamName = MultiReplace(sStreamName, ">", " ")
  sStreamName = MultiReplace(sStreamName, "|", " ")
  sStreamName = MultiReplace(sStreamName, "  ", " ")
  sStreamName = MultiReplace(sStreamName, "--", "-")
  sStreamName = MultiReplace(sStreamName, "==", "=")
  If Left(sStreamName, 1) = "." Then sStreamName = Mid$(sStreamName, 2)
  txtStationName = Trim$(sStreamName)
  lblMultiStream = "Stream " & iMultiShown & " of " & iMultiStreamCt
  
  cmdAddNew.Enabled = True
  
End Sub

Private Sub cmdAddNew_Click()
  AddNewToArray
End Sub

Private Sub cmdClearSaves_Click()

  ReDim Preserve NewShoutcastStation(1)
  NewShoutcastStation(1).StationName = ""
  lblAddedStationsCt = "No stations added yet."
  
End Sub

Private Sub cmdExit_Click()
  
  Winsock1.Close
  
  Unload Me
  
End Sub

Sub FetchStationsFor(sWhichGenre As String)

 'This routine fetches all of the stations which claim the Genre in sWhichGenre
 'Every little line of it has " - [SHOUTcast.com]" in it and is redundant.  I take it out later.
   
  Dim ipos As Long
  Dim i As Long
  Dim iEndpos As Long
  Dim tCount As Integer
  Dim p As Long
  Dim j As Long
  Dim sStation As String
  Dim sFormat As String
  Dim sID As String
  Dim sGenre As String
  Dim sCurrTrack As String
  Dim sBitRate As String
  Dim sLC As String
  Dim lStartWaitSeconds As Long
  Dim lWaitseconds As Long
  Dim lRetryCt As Long
  Dim iSpinCt As Long
  Dim iStationCOunt As Long
  
  Winsock1.Close: DoEvents
  
  gtFetchedStationCt = 0
  i = 0
  sWholeData = ""
  
  sShoutCastURL = "http://yp.shoutcast.com/sbin/newxml.phtml?genre=" & sWhichGenre
  sEndTag = "</stationlist>"
  Winsock1.Connect
  
  iSpinCt = 0
  Do
    Wait 20
    iSpinCt = iSpinCt + 1
    DoEvents
    If iSpinCt > 400 Then
      MsgBox "Waiting too long for Stations list.  Stopping waiting now."
      bDataReceived = True
    End If
  Loop Until bDataReceived = True
  If gbDebugLogic Then Debug.Print MyTime() & "</stationlist> wait spin count: " & iSpinCt

  Winsock1.Close
  bDataReceived = False
  
  sStationData = sWholeData
  If sStationData <> "" Then pbStations.Max = Len(sWholeData)
  pbStations.Value = 0
  pbStations.Visible = True
  
  sID = Replace(sStationData, " - [SHOUTcast.com]", "")
  sStationData = Replace(sID, "&amp;", "&")
  lstStations.Clear
  
  Do
    j = InStr(i + 1, sStationData, "<station name=", vbTextCompare)
    If j > 0 Then
      If j <= pbStations.Max Then pbStations.Value = j
      ipos = InStr(j, sStationData, sQuote) + 1
      iEndpos = InStr(ipos, sStationData, sQuote)
      If iEndpos = 0 Then  ' This means that the input data is corrupt.  Opening quote without end quote.
        MsgBox "The fetched data for this genre is corrupt.  You may want to try clicking again."
        j = 0
      Else
        sStation = Mid$(sStationData, ipos, iEndpos - ipos)
       'Now clean it up for viewing by normal human eyes
        If InStr(sStation, "]]") Then sStation = MultiReplace(sStation, "]]", "]")
        If InStr(sStation, "[[") Then sStation = MultiReplace(sStation, "[[", "[")
        If InStr(sStation, "^^") Then sStation = MultiReplace(sStation, "^^", "^")
        If InStr(sStation, "~~") Then sStation = MultiReplace(sStation, "~~", "~")
        If InStr(sStation, "!!") Then sStation = MultiReplace(sStation, "!!", "!")
        If InStr(sStation, "((") Then sStation = MultiReplace(sStation, "((", "(")
        If InStr(sStation, "))") Then sStation = MultiReplace(sStation, "))", ")")
        If InStr(sStation, " -") Then sStation = MultiReplace(sStation, " -", "")
        If InStr(sStation, "::") Then sStation = MultiReplace(sStation, "::", ":")
        If InStr(sStation, "..") Then sStation = MultiReplace(sStation, "..", ".")
        If InStr(sStation, "\") Then sStation = MultiReplace(sStation, "\", " ")
        If InStr(sStation, "/") Then sStation = MultiReplace(sStation, "/", " ")
        If InStr(sStation, "/") Then sStation = MultiReplace(sStation, ":", " ")
        If InStr(sStation, "*") Then sStation = MultiReplace(sStation, "*", " ")
        If InStr(sStation, "?") Then sStation = MultiReplace(sStation, "?", " ")
        If InStr(sStation, sQuote) Then sStation = MultiReplace(sStation, sQuote, " ")
        If InStr(sStation, "<") Then sStation = MultiReplace(sStation, "<", " ")
        If InStr(sStation, ">") Then sStation = MultiReplace(sStation, ">", " ")
        If InStr(sStation, "|") Then sStation = MultiReplace(sStation, "|", " ")
        If InStr(sStation, "  ") Then sStation = MultiReplace(sStation, "  ", " ")
        If InStr(sStation, "--") Then sStation = MultiReplace(sStation, "--", "-")
        If InStr(sStation, "==") Then sStation = MultiReplace(sStation, "==", "=")
        If Left(sStation, 1) = "." Then sStation = Mid$(sStation, 2)
        sStation = Trim$(sStation)
        
        ipos = InStr(iEndpos, sStationData, "mt=") + 4
        iEndpos = InStr(ipos, sStationData, sQuote)
        sFormat = Mid$(sStationData, ipos, iEndpos - ipos)
        
        ipos = InStr(iEndpos, sStationData, "id=") + 4
        iEndpos = InStr(ipos, sStationData, sQuote)
        sID = Mid$(sStationData, ipos, iEndpos - ipos)
        
        ipos = InStr(iEndpos, sStationData, "br=") + 4
        iEndpos = InStr(ipos, sStationData, sQuote)
        sBitRate = Mid$(sStationData, ipos, iEndpos - ipos)
        
        ipos = InStr(iEndpos, sStationData, "genre=") + 7
        iEndpos = InStr(ipos, sStationData, sQuote)
        sGenre = Mid$(sStationData, ipos, iEndpos - ipos)
        
        ipos = InStr(iEndpos, sStationData, "ct=") + 4
        iEndpos = InStr(ipos, sStationData, sQuote)
        sCurrTrack = Mid$(sStationData, ipos, iEndpos - ipos)
        
        ipos = InStr(iEndpos, sStationData, "lc=") + 4
        iEndpos = InStr(ipos, sStationData, sQuote)
        sLC = Mid$(sStationData, ipos, iEndpos - ipos)
        
        i = iEndpos
      
        gtFetchedStationCt = gtFetchedStationCt + 1
        If gtFetchedStationCt > MAX_SC_STATIONS Then
          MsgBox "Too many stations, please update MAX_SC_STATIONS and recompile"
          Unload Me
          Exit Sub
        End If
        gtFetchedStations(gtFetchedStationCt).StationName = sStation
        gtFetchedStations(gtFetchedStationCt).Format = sFormat
        gtFetchedStations(gtFetchedStationCt).ID = sID
        gtFetchedStations(gtFetchedStationCt).BitRate = sBitRate
        gtFetchedStations(gtFetchedStationCt).Genre = sGenre
        gtFetchedStations(gtFetchedStationCt).ListenerCount = sLC
        lstStations.AddItem sStation
        lstStations.ItemData(lstStations.NewIndex) = Val(sID)
        If lstStations.ListCount Mod 10 = 0 Then
          lblFoundStations = "Found " & lstStations.ListCount & " stations."
          DoEvents
        End If
      End If
    End If
  Loop Until j = 0
  
  pbStations.Visible = False
  
End Sub

Private Function SearchList(ToSearch As String, lstList As ListBox, iStart As Long) As Integer

  Dim i As Integer
  
  SearchList = -1

  For i = iStart To lstList.ListCount - 1
    If InStr(1, LCase(lstList.List(i)), LCase(ToSearch)) Then
      SearchList = i
      Exit For
    End If
    
  Next i

End Function

Private Sub cmdNextStream_Click()

  cmdPreviousStream.Enabled = True
  
  iMultiShown = iMultiShown + 1
  
  If iMultiShown = iMultiStreamCt Then cmdNextStream.Enabled = False
  
  UpdateStreamInfo
  
End Sub

Private Sub cmdPreviousStream_Click()

  cmdNextStream.Enabled = True
  
  
  If iMultiShown > 1 Then iMultiShown = iMultiShown - 1
  If iMultiShown < 1 Then iMultiShown = 1  ' First press fixup
  If iMultiShown = 1 Then cmdPreviousStream.Enabled = False
  
  UpdateStreamInfo
  
End Sub

Private Sub Form_Load()

  iFoundGenre = 0   ' For some reason this one is not going back to 0 on repeat form load.
  iFoundStation = 0 ' Something holding it in memory. That is bad!
  iNewStation2Add = 0  ' No stations added to array yet
  NewShoutcastStation(1).StationName = ""  ' The indication that nothing was added since the UBound will never be 0.
  
  Me.Top = GetSetting(App.EXEName, "Form", "frmDataFetch_Top", frmMain.Top)
  If Me.Top < 0 Then Me.Top = 0
  If Me.Top + Me.Height > Screen.Height Then Me.Top = Screen.Height - Me.Height
  
  Me.Left = GetSetting(App.EXEName, "Form", "frmDataFetch_Left", frmMain.Left)
  If Me.Left < 0 Then Me.Top = 0
  If Me.Left + Me.Width > Screen.Width Then Me.Left = Screen.Width - Me.Width
  
  FetchGenreInfo
  
  Me.MousePointer = vbDefault  ' This was set before I got here.  Reset it now.
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  SaveSetting App.EXEName, "Form", "frmDataFetch_Top", Me.Top
  SaveSetting App.EXEName, "Form", "frmDataFetch_Left", Me.Left

End Sub

Private Sub Form_Resize()

  If Me.Width < 11700 Then Me.Width = 11700
  
  lstStations.Width = Me.Width - lstGenres.Left - lstGenres.Width - 300
  
  lblURL.Top = Me.Height - 900
  lblGenres.Top = lblURL.Top - lblGenres.Height - 30
  
  lblListeners.Top = lblGenres.Top - lblListeners.Height - 30
  lblBitRate.Top = lblListeners.Top
  
  lblFormat.Top = lblListeners.Top - lblFormat.Height - 30
  lblID.Top = lblFormat.Top
  
  txtStationName.Top = lblID.Top - txtStationName.Height - 30
  cmdPreviousStream.Top = txtStationName.Top - cmdPreviousStream.Height - 30
  cmdNextStream.Top = cmdPreviousStream.Top
  lblMultiStream.Top = cmdPreviousStream.Top
  lblFoundStations.Top = lblMultiStream.Top
  pbStations.Top = lblFoundStations.Top
  
  lstGenres.Height = cmdNextStream.Top - 700
  
  lstStations.Height = lstGenres.Height
  
  lblGenreSearch.Top = txtStationName.Top
  lblGenreSearch.Left = txtStationName.Left + txtStationName.Width + 100
  
  txtGenreSearch.Left = lblGenreSearch.Left
  txtGenreSearch.Top = lblGenreSearch.Top + lblGenreSearch.Height
  txtGenreSearch.Width = Me.Width - txtStationName.Width - 500
  pbStations.Width = txtGenreSearch.Left + txtGenreSearch.Width - pbStations.Left
  
  lblFoundStations.Left = txtGenreSearch.Left
  lblFoundStations.Width = txtGenreSearch.Width
  
  lblSearch.Top = txtGenreSearch.Top + txtGenreSearch.Height + 300
  lblSearch.Left = txtGenreSearch.Left
  txtStationSearch.Width = txtGenreSearch.Width
  
  txtStationSearch.Top = lblSearch.Top + lblSearch.Height
  txtStationSearch.Left = txtGenreSearch.Left
    
  cmdAddNew.Top = txtStationSearch.Top + txtStationSearch.Height + 180
  cmdAddNew.Left = txtStationSearch.Left
  cmdExit.Top = cmdAddNew.Top - 10
  
  cmdExit.Left = Me.Width - cmdExit.Width - 300

End Sub

Private Sub lstGenres_Click()
  
  Me.MousePointer = vbHourglass
  Me.Enabled = False
  
  txtStationName = "": lblFormat = "": lblID = "": lblBitRate = "": lblGenres = ""
  lblMultiStream = "...": txtStationSearch = ""
  lblListeners = "": lblURL = "": DoEvents
  
  lblFoundStations = "Retreiving stations list from Shoutcast server..."
  lblFoundStations.Visible = True: DoEvents
  lstStations.Visible = False
  
  FetchStationsFor lstGenres.Text
  
  lblFoundStations.Visible = False
  lstStations.Visible = True
  
  Me.MousePointer = vbDefault
  
  iFoundStation = -1  ' Will be used for station location
  Me.Enabled = True
  
End Sub


Private Sub lstGenres_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub lstStations_Click()
  
 'This code gets the item data for the station that was clicked on in lstStations via lstStations.ListIndex
 'It then displays the information for that station in the text box and labels at the bottom of the screen.
 
 'It also goes out and gets 0 or more streams in a [playlist] (see Winamp docs) structure.
 'If there are more than 1 streams, they are detailed into a few of the boxes.  See the code. I am tired of typing now!
 
  Dim iURLpos As Long
  Dim iTitlepos As Long
  Dim sSearchTitle As String
  Dim iFilepos As Long
  Dim sSearchFile As String
  Dim sSearchNE As String
  Dim sSearchLC As String
  Dim iLCpos As Long
  Dim iNSPos As Long  ' Starting pos for "numberofentries=" clause.  Really want the number right after it.
  Dim sNS As String   ' The bytes, normally numeric, telling the number of entries.
  Dim sLC As String
  Dim iEndpos As Long
  Dim sFile As String
  Dim i As Long
  Dim iSpinCt As Long
  Dim iClickedStationID As Long
  
  If gbDebugLogic Then Debug.Print MyTime() & "Click on station: " & lstStations.ListIndex
  
  cmdAddNew.Enabled = False
  iMultiShown = 0
  
  txtStationName = "": lblFormat = "": lblID = "": lblBitRate = "": lblGenres = ""
  lblMultiStream = "..."
  lblListeners = "": lblURL = ""
  
  i = lstStations.ListIndex + 1
  sFile = lstStations.Text
  txtStationName = Trim$(sFile)
  If gbDebugLogic Then Debug.Print MyTime() & "Station ID for listbox item " & lstStations.ListIndex; " is " & lstStations.ItemData(lstStations.ListIndex)
 
 'Now, loop throught gtFetchedStations for a match on the StationID which is the key to everything.
  iClickedStationID = lstStations.ItemData(lstStations.ListIndex)
  For i = 1 To gtFetchedStationCt
    If gtFetchedStations(i).ID = iClickedStationID Then Exit For
  Next
  
  If gbDebugLogic Then Debug.Print MyTime() & "The next two should match:"
  If gbDebugLogic Then Debug.Print MyTime() & txtStationName
  If gbDebugLogic Then Debug.Print MyTime() & gtFetchedStations(i).StationName
  
  lblFormat = "Format: " & gtFetchedStations(i).Format
  lblID = "ID: " & gtFetchedStations(i).ID
  lblBitRate = "Bitrate: " & gtFetchedStations(i).BitRate & "kb/sec"
  lblGenres = "Genre(s): " & gtFetchedStations(i).Genre
  lblListeners = gtFetchedStations(i).ListenerCount & " Listening"
  
 'Now it is time to retrieve the [playlist] formatted info which lists several things.
 ' 1. The standard HTTP header stuff including a date/time that the [playlist] info will expire.
 '    To be totally nice, I should keep the info saved off somewhere but this would be a hassle.
 '    Maybe...  someday...  or never.
 ' 2. The "[playlist]" header
 ' 3. the "numberofentries=" identifier followed by numeric digits (I hope!) with the number of
 '    streams from 0 to whatever that this stations provides.
 ' 4. The URL for the stream.  It may be that the stations does stream balancing across the streams.
 ' 5. The title of the stream.  So far, this has matched the stations title exactly on all streams.
 ' 6. An normally unused Length indicator.  This is really for file data, not streams.
  
  sShoutCastURL = "http://yp.shoutcast.com/sbin/tunein-station.pls?id=" & gtFetchedStations(i).ID
  sEndTag = "Version="
  sWholeData = ""
  Winsock1.Close: DoEvents
  Winsock1.Connect
  bDataReceived = False
  
  iSpinCt = 0
  Do
    Wait 20
    iSpinCt = iSpinCt + 1
    DoEvents
    If iSpinCt > 400 Then
      MsgBox "Waiting too long for Station URL.  Stopping waiting now."
      bDataReceived = True
    End If
  Loop Until bDataReceived = True
  If gbDebugLogic Then Debug.Print MyTime() & "Title1 wait spin count: " & iSpinCt

  Winsock1.Close
  bDataReceived = False
  
  If gbDebugLogic Then Debug.Print MyTime() & "--------- Stream Playlist starts here ---------'"
  If gbDebugLogic Then Debug.Print sWholeData
  If gbDebugLogic Then Debug.Print MyTime() & "---------- Stream Playlist ends here ----------'"
  
  sSearchNE = "numberofentries="
  iNSPos = InStr(1, sWholeData, sSearchNE)
  If iNSPos = 0 Then
    MsgBox "Data Error LSC2 -- Returned playlist does not contain 'numberofentries='"
    Exit Sub
  End If
  iEndpos = InStr(iNSPos, sWholeData, Chr$(10))
  sNS = Mid$(sWholeData, iNSPos + Len(sSearchNE), iEndpos - iNSPos - Len(sSearchNE))
  If Not IsNumeric(sNS) Then
    MsgBox "Data Error LSC3 -- Number of streams indicator is not numeric.  Received: '" & sNS & "' instead.  The transmission was corrupted. Try it again, if you dare..."
    Exit Sub
  End If
  iMultiStreamCt = Val(sNS)
  If iMultiStreamCt = 0 Then
    lblMultiStream = "Station has no streams."
    lblURL = lblMultiStream
    Exit Sub
  End If
  
  ReDim sMultiStreamURL(iMultiStreamCt)  ' "Preserve" not needed here.
  ReDim sMultiStreamLC(iMultiStreamCt)
  ReDim sMultiStreamTitle(iMultiStreamCt)

  For i = 1 To iMultiStreamCt
    
    sSearchTitle = "Title" & Trim$(CStr(i))
    iTitlepos = InStr(1, sWholeData, sSearchTitle)
    
   'Get Stream Title (for now, includes listener count and maximum listeners)
    iEndpos = InStr(iTitlepos, sWholeData, Chr$(10))
    sFile = Mid$(sWholeData, iTitlepos + Len(sSearchTitle) + 1, iEndpos - iTitlepos - (Len(sSearchTitle) + 1))
    
   'OK, Title still has some stuff up front.  Let's get the listener count and max listeners out of there.
    iTitlepos = InStr(1, sFile, ")")  ' A ")" ends up the Listener Count/Max Listeners phrase
   'Will use this later to chop off the leading stuff after capture.
    iLCpos = InStrRev(sFile, " ", iTitlepos)
    sLC = Mid$(sFile, iLCpos + 1, iTitlepos - iLCpos - 1)
    sFile = Mid$(sFile, iTitlepos + 1)
    
    sMultiStreamTitle(i) = Trim$(sFile)
    sMultiStreamLC(i) = Trim$(sLC)
    
   'Now, find URL for this stream
    sSearchFile = "File" & Trim$(CStr(i))
    iFilepos = InStr(1, sWholeData, sSearchFile)
    iEndpos = InStr(iFilepos, sWholeData, Chr$(10))
    sFile = Mid$(sWholeData, iFilepos + Len(sSearchFile) + 1, iEndpos - iFilepos - (Len(sSearchFile) + 1))
    sFile = Replace$(sFile, "http://", "")
    sMultiStreamURL(i) = "URL: " & sFile
   
    iTitlepos = InStr(sWholeData, "Title" & Trim$(CStr(i)))
    
  Next
  
'--------- Stream Playlist [.pls] sample starts here ---------'

'HTTP/1.1 200 OK
'Date: Mon, 09 Nov 2009 05:02:59 GMT
'Server: Apache
'Cache -Control: Max -age = 86400
'Expires: Tue, 10 Nov 2009 05:02:59 GMT
'Content-Length: 397
'Connection: Close
'Content-Type: audio/x-scpls
'
'[Playlist]
'numberofentries = 3
'File1=http://213.251.190.150:8750
'Title1=(#1 - 1/100) ))) POLSKASTACJA .PL ))) - JAZZ (Polskie Radio),aacplus
'Length1 = -1
'File2=http://87.98.236.207:80
'Title2=(#2 - 2/150) ))) POLSKASTACJA .PL ))) - JAZZ (Polskie Radio),aacplus
'Length2 = -1
'File3=http://213.251.138.82:8750
'Title3=(#3 - 2/150) ))) POLSKASTACJA .PL ))) - JAZZ (Polskie Radio),aacplus
'Length3 = -1
'version = 2
'
'---------- Stream Playlist sample ends here ----------'
  
  If iMultiStreamCt = 1 Then
    lblMultiStream = "This station has 1 stream."
    cmdNextStream.Enabled = False
    cmdPreviousStream.Enabled = False
    lblURL = sMultiStreamURL(1)
    aSplit = Split(sMultiStreamLC(1), "/")
    If UBound(aSplit) < 1 Then
      MsgBox "Program Error LSC1 in UpdateStreamInfo splitting " & sMultiStreamLC(iMultiShown)
      Exit Sub
    End If
    lblListeners = aSplit(0) & " of " & aSplit(1) & " (max) listeners."
  Else
    lblMultiStream = "Total listeners on " & iMultiStreamCt & " streams."
    cmdNextStream.Enabled = True
    cmdPreviousStream.Enabled = True
    cmdAddNew.Enabled = False
    lblURL = "Click the buttons, above, to select a stream."
    iSpinCt = 1000 ' Fake out the .enabled, below.
  End If
    
  If iSpinCt < 501 Then cmdAddNew.Enabled = True
  
End Sub

Private Sub lstStations_DblClick()
 'AddNewToArray  ' Only if the defaut stream is selected there.  Currently not, so not here, too.
End Sub

Private Sub txtGenreSearch_GotFocus()

  If Len(txtGenreSearch) > 0 Then
    txtGenreSearch.SelStart = 0
    txtGenreSearch.SelLength = Len(txtGenreSearch)
  End If
  
End Sub

Private Sub txtGenreSearch_KeyPress(KeyAscii As Integer)

  If gbDebugLogic Then Debug.Print MyTime() & "Genre Search Keypress " & KeyAscii
  
  If KeyAscii = 13 Then
    KeyAscii = 0  ' Kill off input
    Me.MousePointer = vbHourglass
    
    FindGenre
    
    Me.MousePointer = vbDefault
  End If
  
End Sub

Private Sub txtStationSearch_GotFocus()

  If Len(txtStationSearch) > 0 Then
    txtStationSearch.SelStart = 0
    txtStationSearch.SelLength = Len(txtStationSearch)
  End If
  
End Sub

Public Function AutoComplete(Word As String, List As ListBox, Skip As Integer) As Long
 
 'Function: Completes a word by searching through a specified listbox
 '          Will skip as many matches as the number you type in for Skip
  
  Dim i As Long
  Dim j As Long
  Dim SkipAmount As Integer
  Dim k As Long
  
  k = Len(Word)
  
  SkipAmount = Skip
  For j = 0 To List.ListCount - 1
    If UCase(Word) = UCase(Left(List.List(j), k)) Then
      If SkipAmount > 0 Then
        SkipAmount = SkipAmount - 1
      Else
        AutoComplete = j  ' List.List(J)
        Exit Function
      End If
    End If
  Next

End Function

Private Sub txtStationSearch_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
    KeyAscii = 0
    Me.MousePointer = vbHourglass
    
    FindStations
    
    Me.MousePointer = vbDefault
  End If
  
End Sub

Private Sub Winsock1_Connect()
  
  Winsock1.SendData "GET " & sShoutCastURL & " HTTP/1.0" & vbCrLf & "Accept: */*" & vbCrLf & "Accept: text/html" & vbCrLf & vbCrLf
  If gbDebugLogic Then Debug.Print MyTime() & "Connection string: " & sShoutCastURL
  
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

 'Sorta like the old days with modems.  You have to sit around and
 'wait for the tail to come in.  So, we do.  When sEndTag is in the
 'string, we declare it a done deal and notify the caller that he
 'can proceed with business.  But while waiting, one has to piece
 'together the incoming bits until the end arrives.
 
 'The end is known by finding sEndTag in the input stream.
 
  Dim sData As String
  
  Winsock1.GetData sData, vbString
  
  sWholeData = sWholeData & sData
  lblFoundStations = "Data bytes received: " & Len(sWholeData)
  
  If gbDebugBuffer Then Debug.Print MyTime() & "Received: " & Len(sWholeData) & " bytes waiting for " & sEndTag
  If gbDebugBuffer Then Debug.Print MyTime() & sWholeData
  
  If InStr(1, sWholeData, sEndTag, vbTextCompare) Then
    bDataReceived = True
    Winsock1.Close
  End If

  If gbDebugFile And bDataReceived Then
    Dim iFile As Integer
    iFile = FreeFile()
    Open App.Path & "\Received Data.txt" For Append Access Write As #iFile
    Print #iFile, sWholeData
    Print #iFile, "---------------------------------------------------------"
    Close #iFile
  End If
  
End Sub


