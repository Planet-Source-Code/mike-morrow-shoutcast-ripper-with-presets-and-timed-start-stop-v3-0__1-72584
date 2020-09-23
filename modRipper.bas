Attribute VB_Name = "modRipper"
Option Explicit

  Public bExitNow As Boolean  ' This is set by either the reboot or shutdown form and, if set, I will cleanly exit.
  
  Public Enum TimedForce
    F_None
    F_On
    F_Off
  End Enum
  Public giForceTimed As TimedForce
 '-f=off or -f=on -- if missing, use Getsetting info (mostly for debugging while already running)
 'Required minimum is 2 characters of options: 'on' for on, 'of' for off but it can be all 3 for 'off'
  
  Public giSavedStationNo As Long
 '-s=12 -- "Tune" to this stream immediately.  This will be a saved station number from the list.
  
  Public gbMinimizeMe As Boolean
 '-m=y(es) or -m=n(o) (default)  Minimize me after startup is done and -f=on is set.
 
  Public gbRecordImmediately As Boolean  ' (may prove to be unneccessary and be a biproduct of gsSavedStationNo)
 '-i=y or -i=n -- record immediately upon starting
  
  Public gbQuickExit As Boolean
  
  Public gbDebugLogic As Boolean
  Public gbDebugFile As Boolean
  Public gbDebugBuffer As Boolean
  Public gbStopAllFunctions As Boolean  ' Goes true when Exit pressed.
  
  Public bAutoStartPlay As Boolean
  
  Public giNextToRecord As Long
  Public gbNowRecording As Boolean
  Public gbStopContinuousPlay As Boolean
  
  Public gsPassedDayName As String  ' Passed to frmSelectStartDate from frmStationMemory
  
 'This is a FIFO queue.  Items are added to the end and played form the top (array element 0) and removed as played.
  Public sPlayFilesQueue() As String
  Public iPlayFilesQueue As Long  ' This is where the new item to play gets added to the queue.
  
  Public Enum WMPStates
    Undefined = 0      ' Windows Media Player is in an undefined state.
    Stopped = 1        ' Playback of the current media item is stopped.
    Paused = 2         ' Playback of the current media item is paused. When a media item is paused, resuming playback begins from the same location.
    Playing = 3        ' The current media item is playing.
    ScanForward = 4    ' The current media item is fast forwarding.
    ScanReverse = 5    ' The current media item is fast rewinding.
    Buffering = 6      ' The current media item is getting additional data from the server.
    Waiting = 7        ' Connection is established, but the server is not sending data. Waiting for session to begin.
    MediaEnded = 8     ' Media item has completed playback.
    Transitioning = 9  ' Preparing new media item or shucking old media item.
    Ready = 10         ' Ready to begin playing.  You can issue Play command here or at Stopped.
    Reconnecting = 11  ' Reconnecting to stream.
  End Enum
  
  Public sWMPStates(11) As String
    
  Public sParmsFileVersion
  
 'V2 format file added UseICYSongTitle to the end of the others
 'Not done yet, maybe never -- V3 format file added a directory for each stream so the user can sort the recordings automatically.
  Public Const CURRENT_PARMS_VERSION = "V2"  ' Writing file version 2 now.
  
  Public iFinallyTheStationToTune As Integer
  
  Public aSplit() As String
  
  Public Const MP3Specifier = "audio/mpeg"
  Public Const MP3FileType = ".mp3"
  
  Public Const AACSpecifier = "audio/aacp"
  Public Const AACFileType = ".aac"
  
  Dim aLC() As String
  
  Public Const NO_DATE = "No Date"
  Public Const NOT_SET = "Not Set"
  Public Const USE_FILE_PREFIX = "Use this file prefix: "
  
  Public gbIgnoreClicks As Boolean
  
  Public Enum ReturnedDateType
    UseDate
    UseDay
    UseNone
    NoChange
  End Enum
 
  Public giUseReturnedOrNot As Integer
  
  Public gdtStartDate As Date
  Public giReturnedDay As Long
  
  Public gsDayNames(6) As String
  
  Public Const IN_BUFFERS = 10
  Public sInBuffer(IN_BUFFERS) As String
  
  Public giCurrentInBuffer As Long  ' This ranges from 1 to 10 to show which buffer was last received.  0 not used for now.
  Public giBuffersReceived As Long  ' Just for my interest, how many buffers received.

  Public Const sQuote = """"

  Public Const MyBirth = "October 15, 2009 00:00:00"
  Public gdtEpoch As Date

  Private Const MAX_SC_GENRE = 5000  ' Maximum Shoutcast Genre entries
  Public gsGenre(MAX_SC_GENRE) As String
 'Public gsGenreCt As Long  ' How many genre there are.
  
 'Increment this when anything added to file and be SURE to add it after anything already there.
  Public SavedFileVersion As Integer
 ' Version 1 was the first released version (some people may have it)
 ' Version 2 added UseICYSongTitle (boolean) to the saved station information
  
  Public Const MAX_SC_STATIONS = 10000  ' Max Shoutcast Stations/Genre
  Public Const MAX_MEMORY_STATIONS = 1000  'Max Stations in memory
  Public Type Station
    StationName As String      ' Human readable station ID from Shoutcast
    MyFilePrefix As String     ' What I want to prefix the recorded filename with (required)
    Format As String           ' Encoding format
    ID As String               ' Numeric Shoutcast ID
    BitRate As String          ' Encoded stream bitrate
    Genre As String            ' Could have multiple genre here
    StartDate As String        ' Start Day Name or Date to start recording or "Not Set" (optional)
    StartHour As String        ' (Entered) Start Hour (0-23) to start recording (optional)
    StartMin As String         ' (Entered) Start Minute (0-59) to start recording (optional)
    Duration As String         ' (Entered) Duration of recording in minutes
    CurrentTrack As String     ' Name of what is playing now
    ListenerCount As String    ' How many are listening
    URL As String              ' Address and port of station
    UseICYSongTitle As Integer ' If vbChecked, use ICY song title and break songs up by title (kinda stupid on most stations)
   'Anything after this is not saved but an internal calculation refreshed in several places.
    dtAutoRecordStart As Date  ' (Calculated) Date & Time for autorecord start (as typed in StartDate, Hour and Minute)
    dtAutoRecordStop As Date   ' (Calculated) Date & Time for autorecord stop (above + duration minutes)
  End Type
 
  Public gtFetchedStations(MAX_SC_STATIONS) As Station  ' Temporary use in frmDataFetch to hold information on fetched stations
  Public gtFetchedStationCt As Long
  
  Public gtMemorizedStations(MAX_MEMORY_STATIONS) As Station  ' Permanent use in frmStationMemory and saved between executions
  Public giMemorizedStationsCt As Long  ' Highest used station
  
  Public iStationsInfoFN As Integer
  Public iStationDetailFN As Integer
  Public iGenreDataFN As Integer

  Public NewShoutcastStation() As Station

 'The Wait Timer Declarations ---------------------------------------------

  Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
  End Type

  Private Const WAIT_ABANDONED& = &H80&
  Private Const WAIT_ABANDONED_0& = &H80&
  Private Const WAIT_FAILED& = -1&
  Private Const WAIT_IO_COMPLETION& = &HC0&
  Private Const WAIT_OBJECT_0& = 0
  Private Const WAIT_OBJECT_1& = 1
  Private Const WAIT_TIMEOUT& = &H102&
  
  Private Const INFINITE = &HFFFF
  Private Const ERROR_ALREADY_EXISTS = 183&
  
  Private Const QS_HOTKEY& = &H80
  Private Const QS_KEY& = &H1
  Private Const QS_MOUSEBUTTON& = &H4
  Private Const QS_MOUSEMOVE& = &H2
  Private Const QS_PAINT& = &H20
  Private Const QS_POSTMESSAGE& = &H8
  Private Const QS_SENDMESSAGE& = &H40
  Private Const QS_TIMER& = &H10
  
  Private Const QS_MOUSE& = (QS_MOUSEMOVE _
                          Or QS_MOUSEBUTTON)
  
  Private Const QS_INPUT& = (QS_MOUSE _
                          Or QS_KEY)
  
  Private Const QS_ALLEVENTS& = (QS_INPUT _
                              Or QS_POSTMESSAGE _
                              Or QS_TIMER _
                              Or QS_PAINT _
                              Or QS_HOTKEY)
  
  Private Const QS_ALLINPUT& = (QS_SENDMESSAGE _
                             Or QS_PAINT _
                             Or QS_TIMER _
                             Or QS_POSTMESSAGE _
                             Or QS_MOUSEBUTTON _
                             Or QS_MOUSEMOVE _
                             Or QS_HOTKEY _
                             Or QS_KEY)

  Private Declare Function CreateWaitableTimer Lib "kernel32" _
    Alias "CreateWaitableTimerA" ( _
    ByVal lpSemaphoreAttributes As Long, _
    ByVal bManualReset As Long, _
    ByVal lpName As String) As Long
    
  Private Declare Function OpenWaitableTimer Lib "kernel32" _
    Alias "OpenWaitableTimerA" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal lpName As String) As Long
    
  Private Declare Function SetWaitableTimer Lib "kernel32" ( _
    ByVal hTimer As Long, _
    lpDueTime As FILETIME, _
    ByVal lPeriod As Long, _
    ByVal pfnCompletionRoutine As Long, _
    ByVal lpArgToCompletionRoutine As Long, _
    ByVal fResume As Long) As Long
    
  Private Declare Function CancelWaitableTimer Lib "kernel32" ( _
    ByVal hTimer As Long)
    
  Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long
    
  Private Declare Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
    
  Private Declare Function MsgWaitForMultipleObjects Lib "user32" ( _
    ByVal nCount As Long, _
    pHandles As Long, _
    ByVal fWaitAll As Long, _
    ByVal dwMilliseconds As Long, _
    ByVal dwWakeMask As Long) As Long

  Public Const FOF_ALLOWUNDO = &H40
  Private Const FO_DELETE = &H3
  Private Const FOF_NOCONFIRMATION = &H10
  
  Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
  End Type
  
  Public SH As New Shell  'reference to shell32.dll class
  Public ShBFF As Folder  'Shell Browse For Folder
  
  Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
  
 'BrowseInfo is a type used with the SHBrowseForFolder API call
  Private Type BrowseInfo
    hwndOwner As Long
    pidlroot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
  End Type
  
  Enum BrowseForFolderFlags
    ReturnFileSystemFoldersOnly = &H1
    DontGoBelowDomain = &H2
    IncludeStatusText = &H4
    BrowseForComputer = &H1000
    BrowseForPrinter = &H2000
    BrowseIncludeFiles = &H4000
    IncludeTextBox = &H10
    ReturnFileSystemAncestors = &H8
  End Enum
  
  Enum Folders
    Desktop = &H0
    Internet = &H1
    Programs = &H2
    ControlsFolder = &H3
    Printers = &H4
    Personal = &H5
    Favorites = &H6
    StartUp = &H7
    Recent = &H8
    SendTo = &H9
    RecycleBin = &HA
    StartMenu = &HB
    DesktopDirectory = &H10
    Drives = &H11
    network = &H12
    Nethood = &H13
    Fonts = &H14
    Templates = &H15
    Common_StartMenu = &H16
    Common_Programs = &H17
    Common_StartUp = &H18
    Common_DesktopDirectory = &H19
    ApplicationData = &H1A
    PrintHood = &H1B
    AltStartUp = &H1D
    Common_AltStartUp = &H1E
    Common_Favorites = &H1F
    InternetCache = &H20
    Cookies = &H21
    History = &H22
  End Enum
  
  Private Declare Function SHBrowseForFolder Lib "Shell32" (lpbi As BrowseInfo) As Long
  
 'Declare GetTickCount API for timing (replaces timer controls)
  Declare Function GetTickCount Lib "kernel32.dll" () As Long


Function MultiReplace(ByVal sInStr As String, sSearched As String, sReplaceWith As String) As String

  Dim i As Long
  
  Do While InStr(1, sInStr, sSearched)
    sInStr = Replace(sInStr, sSearched, sReplaceWith)
  Loop

  MultiReplace = sInStr
  
End Function

Function CleanSpace(ByVal strIn As String) As String
  
 'Remove leading or trailing spaces
  strIn = Trim(strIn)
 'Replace all double space pairings with single spaces
  Do While InStr(strIn, "  ")
    strIn = Replace(strIn, "  ", " ")
  Loop
 'Return the result
  CleanSpace = strIn

End Function

Public Function CharCount(OrigString As String, Chars As String, Optional CaseSensitive As Boolean = False) As Long

'**********************************************
'PURPOSE: Returns Number of occurrences of a character or
'or a character sequencence within a string

'PARAMETERS:
    'OrigString: String to Search in
    'Chars: Character(s) to search for
    'CaseSensitive (Optional): Do a case sensitive search
    'Defaults to false

'RETURNS:
    'Number of Occurrences of Chars in OrigString

'EXAMPLES:
'Debug.Print CharCount("FreeVBCode.com", "E") -- returns 3
'Debug.Print CharCount("FreeVBCode.com", "E", True) -- returns 0
'Debug.Print CharCount("FreeVBCode.com", "co") -- returns 2
''**********************************************

Dim lLen As Long
Dim lCharLen As Long
Dim lAns As Long
Dim sInput As String
Dim sChar As String
Dim lCtr As Long
Dim lEndOfLoop As Long
Dim bytCompareType As Byte

sInput = OrigString
If sInput = "" Then Exit Function
lLen = Len(sInput)
lCharLen = Len(Chars)
lEndOfLoop = (lLen - lCharLen) + 1
bytCompareType = IIf(CaseSensitive, vbBinaryCompare, _
   vbTextCompare)

    For lCtr = 1 To lEndOfLoop
        sChar = Mid(sInput, lCtr, lCharLen)
        If StrComp(sChar, Chars, bytCompareType) = 0 Then _
            lAns = lAns + 1
    Next

CharCount = lAns

End Function

Public Sub ShellDeleteOne(sFile As String, ActionFlag As Long)

  Dim SHFileOp As SHFILEOPSTRUCT
  Dim r As Long
  
  sFile = sFile & Chr$(0)
  
  With SHFileOp
    .wFunc = FO_DELETE
    .pFrom = sFile
    .fFlags = ActionFlag Or FOF_NOCONFIRMATION  ' Delete to Recycle Bin without confirmation message.
  End With
  
  r = SHFileOperation(SHFileOp)

End Sub

Public Sub Wait(lNumberOfMilliSeconds As Long)
    
  Dim ft As FILETIME
  Dim lBusy As Long
  Dim lRet As Long
  Dim dblDelay As Double
  Dim dblDelayLow As Double
  Dim dblUnits As Double
  Dim hTimer As Long
  
  hTimer = CreateWaitableTimer(0, True, App.EXEName & "Timer")
  
  If err.LastDllError = ERROR_ALREADY_EXISTS Then
   'If the timer already exists, it does not hurt to open it
   'as long as the person who is trying to open it has the
   'proper access rights.
  Else
    ft.dwLowDateTime = -1
    ft.dwHighDateTime = -1
    lRet = SetWaitableTimer(hTimer, ft, 0, 0, 0, 0)
  End If
  
 'Convert the Units to nanoseconds.
  dblUnits = CDbl(&H10000) * CDbl(&H10000)
 'dblDelay = CDbl(lNumberOfSeconds) * 1000 * 10000
  dblDelay = CDbl(lNumberOfMilliSeconds) * 1000 * 10  ' 000
 'For 1 second, dblDelay = 1,000,000 (1 million)  ' 100,000 = 1/10th sec, 10000 = 1/100th sec, 1000 = 1 millisecond
  
 'By setting the high/low time to a negative number, it tells
 'the Wait (in SetWaitableTimer) to use an offset time as
 'opposed to a hardcoded time. If it were positive, it would
 'try to convert the value to GMT.
  ft.dwHighDateTime = -CLng(dblDelay / dblUnits) - 1
  dblDelayLow = -dblUnits * (dblDelay / dblUnits - _
    Fix(dblDelay / dblUnits))
  
  If dblDelayLow < CDbl(&H80000000) Then
   '&H80000000 is MAX_LONG, so you are just making sure
   'that you don't overflow when you try to stick it into
   'the FILETIME structure.
    dblDelayLow = dblUnits + dblDelayLow
    ft.dwHighDateTime = ft.dwHighDateTime + 1
  End If
  
  ft.dwLowDateTime = CLng(dblDelayLow)
  lRet = SetWaitableTimer(hTimer, ft, 0, 0, 0, False)
  
 'QS_ALLINPUT means that MsgWaitForMultipleObjects will
 'return every time the thread in which it is running gets
 'a message. If you wanted to handle messages in here you could,
 'but by calling Doevents you are letting DefWindowProc
 'do its normal windows message handling---Like DDE, etc.
  Do
    lBusy = MsgWaitForMultipleObjects(1, hTimer, False, INFINITE, QS_ALLINPUT&)
    DoEvents
  Loop Until lBusy = WAIT_OBJECT_0
  
 'Close the handles when you are done with them.
  CloseHandle hTimer
  
End Sub

Function ValidURL(sURL As String) As Boolean
  ValidURL = True
End Function

Function DateCompare(date1 As Date, date2 As Date)

  Dim d1 As Date
  Dim d2 As Date
   
 'Returns True if they are equal or False if they are different
  d1 = CDate(date1)
  d2 = CDate(date2)
  
  If Day(d1) = Day(d2) And Month(d1) = Month(d2) And Year(d1) = Year(d2) Then
    DateCompare = True
  Else
    DateCompare = False
  End If

End Function

Function IsWeekend(tempDate)
 
  If IsDate(tempDate) Then
    Select Case Weekday(tempDate)
      Case 1, 7
       'Sunday or Saturday
        IsWeekend = True
      Case 2, 3, 4, 5, 6
       'Monday through Friday
        IsWeekend = False
    End Select
  Else
    MsgBox "Bad Date passed to IsWeekend: " & tempDate
  End If
  
End Function
 
Function NextBusinessDay(tempDate)
 
' 'Initially, just add one day
'  myDate = CDate(tempDate) + 1
'
' 'Continue to add one day until the date is not a holiday or weekend.
'  Do While IsHoliday(myDate) Or IsWeekend(myDate)
'    myDate = CDate(myDate) + 1
'  Loop
'
'  NextBusinessDay = myDate
 
End Function
Public Function MyTime() As String
  MyTime = Format(Now, "yyyy-MMM-dd HH:nn:ss") & "." & Right(Format(Timer, "#0.00"), 2) & " "
End Function
