VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "GetShoutcast by M. Morrow/F. Aldea"
   ClientHeight    =   9870
   ClientLeft      =   3090
   ClientTop       =   2970
   ClientWidth     =   6030
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   6030
   Tag             =   "-1"
   Begin VB.Frame Frame1 
      Caption         =   "Shoutcast/Saved Station Options (Read-Only)"
      Enabled         =   0   'False
      Height          =   1935
      Left            =   60
      TabIndex        =   22
      Top             =   5100
      Width           =   5895
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   1560
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "0.0.0.0"
         ToolTipText     =   "URL or IP address of Shoutcast/ICY station to record."
         Top             =   300
         Width           =   4200
      End
      Begin VB.CheckBox chkFileBySong 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Use ICY song title for filename (overrides filename prefix, below)"
         Height          =   255
         Left            =   150
         TabIndex        =   26
         Tag             =   "-1"
         ToolTipText     =   "If checked, break stream into separate recordings by received filename (may not always be good to do)."
         Top             =   1260
         Width           =   5555
      End
      Begin VB.CheckBox optTimedStop 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Stop recording at: not set"
         Height          =   255
         Left            =   150
         TabIndex        =   25
         Tag             =   "-1"
         ToolTipText     =   "If checked, recording will stop after this many minutes."
         Top             =   990
         Width           =   5555
      End
      Begin VB.CheckBox optSupplyFilename 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Use this filename prefix: not set"
         Height          =   255
         Left            =   150
         TabIndex        =   24
         Tag             =   "-1"
         ToolTipText     =   "If checked, use supplied name for filename prefix, else use generic filename."
         Top             =   1530
         Width           =   5555
      End
      Begin VB.CheckBox optTimedStart 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Start recording at: not set"
         Height          =   255
         Left            =   150
         TabIndex        =   23
         Tag             =   "-1"
         ToolTipText     =   "If checked, use supplied station (stream) information for recording."
         Top             =   720
         Width           =   5555
      End
      Begin VB.Label lblServer 
         AutoSize        =   -1  'True
         Caption         =   "Server IP/Path:Port:"
         Height          =   195
         Left            =   60
         TabIndex        =   28
         Top             =   330
         Width           =   1440
      End
   End
   Begin VB.Frame fraManualControls 
      Caption         =   "Manual Recording Controls"
      Height          =   1155
      Left            =   60
      TabIndex        =   21
      Top             =   7140
      Width           =   5895
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FF00FF&
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   360
         Left            =   3900
         MaskColor       =   &H00C0FFFF&
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Close all files and cease all operation."
         Top             =   675
         UseMaskColor    =   -1  'True
         Width           =   1800
      End
      Begin VB.CommandButton cmdDisconnectFromServer 
         BackColor       =   &H008080FF&
         Caption         =   "&Disconnect from Server"
         Enabled         =   0   'False
         Height          =   360
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Disconnect from server."
         Top             =   300
         Width           =   1970
      End
      Begin VB.CommandButton cmdStartRecording 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Start Recording"
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Start saving the received data to the Output Path folder with default or entered filename."
         Top             =   675
         UseMaskColor    =   -1  'True
         Width           =   1800
      End
      Begin VB.CommandButton cmdStopRecording 
         BackColor       =   &H008080FF&
         Caption         =   "Stop Re&cording"
         Enabled         =   0   'False
         Height          =   360
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Stop recording stream in the selected file."
         Top             =   660
         UseMaskColor    =   -1  'True
         Width           =   1970
      End
      Begin VB.CommandButton cmdSelectMemory 
         BackColor       =   &H0080C0FF&
         Caption         =   "Saved St&ations"
         Height          =   360
         Left            =   3900
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Display a Msgbox with partial or full information on the displayed station.  Varies with how the station info was acquired."
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   1800
      End
      Begin VB.CommandButton cmdConnectToServer 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Connect to Server"
         Height          =   360
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Attempt to connect to shown server."
         Top             =   300
         Width           =   1800
      End
   End
   Begin VB.Timer tmrUpdateAll 
      Interval        =   1000
      Left            =   5940
      Top             =   840
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Recording/Playback Options"
      Height          =   2775
      Left            =   60
      TabIndex        =   18
      Top             =   2220
      Width           =   5895
      Begin VB.Frame fraPlayControl 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   3430
         TabIndex        =   34
         Top             =   660
         Visible         =   0   'False
         Width           =   1830
         Begin VB.CheckBox chkPlayStop 
            DownPicture     =   "frmMain.frx":0442
            Height          =   660
            Left            =   1200
            Picture         =   "frmMain.frx":0AEA
            Style           =   1  'Graphical
            TabIndex        =   37
            Tag             =   "-1"
            ToolTipText     =   "Play or stop the current stream.  If you stop it, then play again, it wills start from the very beginning."
            Top             =   0
            Width           =   675
         End
         Begin VB.CheckBox chkMutePlayback 
            DownPicture     =   "frmMain.frx":11B2
            Height          =   660
            Left            =   600
            Picture         =   "frmMain.frx":1874
            Style           =   1  'Graphical
            TabIndex        =   36
            Tag             =   "-1"
            ToolTipText     =   "Mute the playback of the stream without changing the volume level.  Click again to hear recording (will not pause)."
            Top             =   0
            Width           =   675
         End
         Begin VB.CheckBox chkPausePlayback 
            DownPicture     =   "frmMain.frx":1F36
            Height          =   660
            Left            =   0
            Picture         =   "frmMain.frx":253F
            Style           =   1  'Graphical
            TabIndex        =   35
            Tag             =   "-1"
            ToolTipText     =   "Pause the playback of the stream.  Click again to resume from paused position."
            Top             =   0
            Width           =   675
         End
      End
      Begin VB.CommandButton cmdBFF 
         Caption         =   "..."
         Height          =   285
         Left            =   5460
         TabIndex        =   32
         Tag             =   "-1"
         ToolTipText     =   "Click to browse the file system for a folder into which to save the Shoutcast stream files."
         Top             =   310
         Width           =   315
      End
      Begin VB.CheckBox chkRecordAllMemorized 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Record all ti&med entries in the saved stations list to the Output Path, above."
         Height          =   780
         Left            =   660
         TabIndex        =   2
         Tag             =   "-1"
         ToolTipText     =   "When checked, program is in ""Automatic Record"" mode.  Uncheck for manual recording."
         Top             =   690
         Width           =   2235
      End
      Begin VB.TextBox txtOutPath 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Text            =   "c:\"
         ToolTipText     =   "Output folder for recordings.  Click ""..."" button at right to select folder."
         Top             =   310
         Width           =   4335
      End
      Begin MSComctlLib.Slider sliVolume 
         Height          =   255
         Left            =   60
         TabIndex        =   3
         ToolTipText     =   "Adjust playback volume, Low to High."
         Top             =   1500
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   450
         _Version        =   393216
         MousePointer    =   9
         Enabled         =   0   'False
         Max             =   100
         SelStart        =   50
         TickFrequency   =   2
         Value           =   50
         TextPosition    =   1
      End
      Begin MSComctlLib.Slider sliBalance 
         Height          =   255
         Left            =   60
         TabIndex        =   30
         ToolTipText     =   "Adjust balance Left to Right."
         Top             =   2100
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   450
         _Version        =   393216
         MousePointer    =   9
         Enabled         =   0   'False
         LargeChange     =   10
         Min             =   -100
         Max             =   100
         TickFrequency   =   4
         TextPosition    =   1
      End
      Begin VB.Label lblPlayControl 
         Caption         =   " Pause    Mute   Play/Stop "
         Height          =   195
         Left            =   3520
         TabIndex        =   33
         Tag             =   "-1"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L------------------------------------- Balance --------------------------------------R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   31
         Tag             =   "-1"
         ToolTipText     =   "Adjust balance Left to Right."
         Top             =   2400
         Width           =   5625
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L------------------------------- Playback Volume -------------------------------H"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   29
         ToolTipText     =   "Adjust playback volume, Low to High."
         Top             =   1800
         Width           =   5625
      End
      Begin VB.Label Label1 
         Caption         =   "O&utput Path:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "-1"
         Top             =   340
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   195
         Left            =   150
         Picture         =   "frmMain.frx":2B48
         Tag             =   "-1"
         Top             =   1800
         Width           =   5625
      End
      Begin VB.Image Image2 
         Height          =   195
         Left            =   150
         Picture         =   "frmMain.frx":3A33
         Tag             =   "-1"
         Top             =   2400
         Width           =   5625
      End
      Begin VB.Label lblPlayControlsComing 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Play Controls will appear here when appropriate."
         Height          =   630
         Left            =   3465
         TabIndex        =   38
         Top             =   850
         Width           =   1785
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2055
      Left            =   60
      ScaleHeight     =   1995
      ScaleWidth      =   5835
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Information panel.  Point to items herein for specific information."
      Top             =   120
      Width           =   5895
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size: n/a"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   17
         ToolTipText     =   "Filesize of output file (stream data)."
         Top             =   1665
         Width           =   645
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File: n/a"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   16
         ToolTipText     =   "Output filename for stream data."
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Running time: n/a"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3195
         TabIndex        =   15
         ToolTipText     =   "Total recording time unless files are being broken into titled segments, then title time."
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "REC"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   45
         TabIndex        =   14
         ToolTipText     =   "Will blink, slowly, when recording active."
         Top             =   1215
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblBitrate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bitrate: n/a"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   13
         ToolTipText     =   "Server bitrate for stream."
         Top             =   1395
         Width           =   795
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title: n/a"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   12
         ToolTipText     =   "Song title if titles being sent in ICY headers."
         Top             =   810
         Width           =   645
      End
      Begin VB.Label lblRadio 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stream ID and Status Messages"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   735
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Stream and server status messages."
         Top             =   0
         UseMnemonic     =   0   'False
         Width           =   5595
      End
   End
   Begin MSWinsockLib.Winsock wsShoutcastReceiver 
      Left            =   5940
      Top             =   420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   5940
      Top             =   1260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Orientation     =   2
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP 
      Height          =   1215
      Left            =   60
      TabIndex        =   20
      Top             =   8580
      Width           =   5895
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   0   'False
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   -1  'True
      _cx             =   10398
      _cy             =   2143
   End
   Begin VB.Label lblCurrTime 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Current Time here"
      Height          =   300
      Left            =   60
      TabIndex        =   19
      Tag             =   "-1"
      ToolTipText     =   "Well, duhhhhh...."
      Top             =   8280
      Width           =   5880
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSR 
      Caption         =   "&Shutdown/Reboot"
      Begin VB.Menu mnuSRS 
         Caption         =   "Shut&down after recordings end"
      End
      Begin VB.Menu mnuSRR 
         Caption         =   "&Reboot after recordings end"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  ''''''''''''''''''''''''''''''''''''''''''''''
  ''    Module written by Fernando Aldea G.   ''
  ''    e-mail: fernando_aldea@terra.cl       ''
  ''    Release October, 2004                 ''
  ''                                          ''
  ''    sorry for not comment the code        ''
  ''    & sorry for my English!               ''
  ''                                          ''
  ''    Your English is NO problem.           ''
  ''    Mike Morrow added some comments.      ''
  ''''''''''''''''''''''''''''''''''''''''''''''
  
  Const ReqHeader = "GET $ HTTP/1.0" & vbCrLf & _
                    "Host: %" & vbCrLf & _
                    "User-Agent: WinampMPEG/2.7" & vbCrLf & _
                    "Accept: */*" & vbCrLf & _
                    "Icy-MetaData:1" & vbCrLf & _
                    "Connection: close" & vbCrLf & vbCrLf
                  
  
  Dim IcyReceived As Boolean
  Dim sIcyHeader As String
  
  Dim sData As String
  Dim sMeta As String
  
  Dim DataLen As Long
  Dim MetaLen As Long
  
  Dim nData As Long
  Dim bMeta As Boolean
  
 'Next var is both a switch and a DSECT pointer.  Extra points if you know what a DSECT is!
 'When = 0, no recording is taking place and that is the switch part.
 'When <> 0, recording happening to that file (by number).
  Dim giFile As Integer
  Dim gFileLen As Long
  Dim gsPath As String
  
  Dim dtStartTime As Date
  
  Private bNowRecording As Boolean
  
 'Private Const LOCAL_PORT = 8000
  Private gStartOnTime As Boolean
  Private gbDataComing As Boolean
  
  Private gdtRecordStartTime As Date
  
  Private gsCurrentOutputFile As String
  
  Private giEnablePlayRecording As Long
  Private gbNowPlaying As Boolean
  
  Private gsStationFormat As String
Sub CalculateStartEndTimes(iWhich As Long)

  Dim s As String
  Dim dt As Date
  Dim i As Integer
  Dim j As Integer
  
 'This routine calculates dtAutoRecordStart and dtAutoRecordStop.
 'These are the automatic start and stop times for each station stream entry.
 
  If IsDate(gtMemorizedStations(iWhich).StartDate) Then
    gtMemorizedStations(iWhich).dtAutoRecordStart = CDate(gtMemorizedStations(iWhich).StartDate & " " & _
                              gtMemorizedStations(iWhich).StartHour & ":" & gtMemorizedStations(iWhich).StartMin)
  Else
    s = Format(Now, "dddd") & "s"
    If gtMemorizedStations(iWhich).StartDate = s Then  ' This looks dumb but I have to find today before the loop, else I find a week from today.
      dt = Now  ' This looks dumb, too, but I don't want any chance of it changing while I am fetching it 3 times.
      gtMemorizedStations(iWhich).dtAutoRecordStart = CDate(Year(dt) & "/" & Month(dt) & "/" & Day(dt) & " " & _
                          gtMemorizedStations(iWhich).StartHour & ":" & gtMemorizedStations(iWhich).StartMin)
    Else
      dt = Now
      For i = 0 To 5
        j = i + 1
        dt = dt + 1
        If Format(dt, "dddd") & "s" = gtMemorizedStations(iWhich).StartDate Then
          gtMemorizedStations(iWhich).dtAutoRecordStart = CDate(Year(dt) & "/" & Month(dt) & "/" & Day(dt) & " " & _
                              gtMemorizedStations(iWhich).StartHour & ":" & gtMemorizedStations(iWhich).StartMin)
        End If
      Next
    End If
  End If
  
  gtMemorizedStations(iWhich).dtAutoRecordStop = DateAdd("n", gtMemorizedStations(iWhich).Duration, gtMemorizedStations(iWhich).dtAutoRecordStart)
  
End Sub

Sub ConnectToServer()

  Dim sServer As String
  Dim sPath As String
  Dim sPort As String
  Dim ipos As Integer
  
 'Remove possible http://
  sServer = Replace(txtServer.Text, "http://", "")
  ipos = InStr(1, sServer, ":")
  If ipos = 0 Then
    MsgBox "Port number not found.  Please enter 'ServerIP' field as 'IPaddr/Path:Port' (Note: Path is needed only if required by desured Shotcast server."
    Exit Sub
  End If
  sPort = Mid(sServer, ipos + 1)
  sServer = Left(sServer, ipos - 1)
  
 'split url to get any existent path after a slash
  ipos = InStr(1, sServer, "/")
  If ipos > 0 Then
    sPath = Mid(sServer, ipos)
    sServer = Left(sServer, ipos - 1)
  End If
  
  ipos = InStr(1, sPort, "/")
  If ipos > 0 Then
    sPath = Mid$(sPort, ipos + 1)
    sPort = Left(sPort, ipos - 1)
  End If
  
  If ValidURL(sServer) Then
    Tune sServer, sPort, sPath
  Else
    MsgBox "Reenter the URL or IP Address.  It does not appear to be valid."
  End If
  
End Sub

Sub DisconnectFromServer()

  wsShoutcastReceiver.Close
  
  lblRadio.Caption = "Disconnected"
  
  StopRecording
  
  cmdStartRecording.Enabled = False  ' Override StopRecording
  cmdDisconnectFromServer.Enabled = False
  
End Sub

Sub FindNextToRecord()
      
  Dim i As Integer
  
  For i = 1 To giMemorizedStationsCt
    If gtMemorizedStations(i).StartDate <> "" Then
      CalculateStartEndTimes (i)
    Else
      gtMemorizedStations(i).dtAutoRecordStart = gdtEpoch
      gtMemorizedStations(i).dtAutoRecordStop = gdtEpoch
    End If
  Next
 
 'Now, all of the start times, if any, have been calculated.  See if any station has a record start time >= now.
 'If so, set bSomethingToRecord to True and update main with the information.
 
  giNextToRecord = 0  ' If this is still 0 at the end, there is nothing to record.
  For i = 1 To giMemorizedStationsCt
    If gtMemorizedStations(i).dtAutoRecordStart <= Now() And Now() < gtMemorizedStations(i).dtAutoRecordStop Then
   'We have a winner.  We are in this time (start/stop) window.  Stop now and use what we have
      giNextToRecord = i  ' This is our guy!  Get out and start recording.
      Exit For
    Else  ' Now test for starting and ending some time in the future
      If gtMemorizedStations(i).dtAutoRecordStart <> gdtEpoch And gtMemorizedStations(i).dtAutoRecordStart >= Now() And _
         gtMemorizedStations(i).dtAutoRecordStart < gtMemorizedStations(giNextToRecord).dtAutoRecordStart Then _
         giNextToRecord = i ' Remember sooner start than previously captured and where the rest of the needed information is.
    End If
  Next
  
  If giNextToRecord > 0 Then  ' If there is something to record...
    
    chkRecordAllMemorized.Enabled = True
    chkRecordAllMemorized.Value = vbChecked
    
    optTimedStart.Caption = "Start recording at: " & gtMemorizedStations(giNextToRecord).dtAutoRecordStart
    optTimedStart.Value = vbChecked
    
    optTimedStop.Caption = "Stop recording at: " & gtMemorizedStations(giNextToRecord).dtAutoRecordStop
    optTimedStop.Value = vbChecked
    
    gsStationFormat = gtMemorizedStations(giNextToRecord).Format
    
    If gtMemorizedStations(giNextToRecord).MyFilePrefix <> "" Then
      optSupplyFilename.Caption = USE_FILE_PREFIX & gtMemorizedStations(giNextToRecord).MyFilePrefix
      optSupplyFilename.Value = vbChecked
    End If
    
    
    txtServer = gtMemorizedStations(giNextToRecord).URL
    
    If gtMemorizedStations(giNextToRecord).UseICYSongTitle = vbChecked Then
      chkFileBySong.Value = vbChecked
    Else
      chkFileBySong.Value = vbUnchecked
    End If
    
    lblRadio = "Next recording is stream: " & gtMemorizedStations(giNextToRecord).StationName
    
    cmdConnectToServer.Enabled = False
    chkRecordAllMemorized.Enabled = True
    cmdSelectMemory.Enabled = False
  Else  ' There is nothing to record.  I am mortified!  Why am I here???  I have NOTHING TO DO!!!  I am going home!!  Release me you brute!!!!!!
        '   ...or kindly pick a stream, please.
    chkRecordAllMemorized.Value = vbUnchecked  ' If it is still disabled, uncheck it.
    optTimedStart.Caption = "Start recording at: " & NOT_SET
    optTimedStart.Value = vbUnchecked
    optTimedStop.Caption = "Stop recording at: " & NOT_SET
    optTimedStop.Value = vbUnchecked
        
    txtServer = ""
    
  End If
  
End Sub

Sub InitVars()

  Dim i As Long
  Dim iFile As Integer
  Dim sWork As String
  Dim iInternalCount As Long
  
  gbIgnoreClicks = True
  gbStopAllFunctions = False
  gbStopContinuousPlay = True  ' Initially, no playing please.
  
  ReDim Preserve sPlayFilesQueue(100)
  ReDim Preserve NewShoutcastStation(0)
  
  iPlayFilesQueue = 0
  
  gsDayNames(0) = WeekdayName(1)
  gsDayNames(1) = WeekdayName(2)
  gsDayNames(2) = WeekdayName(3)
  gsDayNames(3) = WeekdayName(4)
  gsDayNames(4) = WeekdayName(5)
  gsDayNames(5) = WeekdayName(6)
  gsDayNames(6) = WeekdayName(7)
  
  sWMPStates(0) = "Undefined"
  sWMPStates(1) = "Stopped"
  sWMPStates(2) = "Paused"
  sWMPStates(3) = "Playing"
  sWMPStates(4) = "Scan Forward"
  sWMPStates(5) = "Scan Reverse"
  sWMPStates(6) = "Buffering"
  sWMPStates(7) = "Waiting"
  sWMPStates(8) = "Media Ended"
  sWMPStates(9) = "Transitioning"
  sWMPStates(10) = "Ready"
  sWMPStates(11) = "Reconnecting"

  giNextToRecord = 0  ' Point to the "immediate" station -- 0.  If manually recording, use the information from array element 0
  giCurrentInBuffer = 0  ' It will be added to before each use.
  
  gdtEpoch = CDate(MyBirth)  ' This will be the comparison, if needed.  May not be needed...
  
  Me.Top = GetSetting(App.EXEName, "Form", "frmMain_Top", 0)
  If Me.Top < 0 Then Me.Top = 0
  If Me.Top + Me.Height > Screen.Height Then Me.Top = Screen.Height - Me.Height
  
  Me.Left = GetSetting(App.EXEName, "Form", "frmMain_Left", 0)
  If Me.Left < 0 Then Me.Left = 0
  If Me.Left + Me.Width > Screen.Width Then Me.Left = Screen.Width - Me.Width
  
  txtServer = GetSetting(App.EXEName, "LastTuned", "IP", "0.0.0.0:8000")
  
  txtOutPath = GetSetting(App.EXEName, "Options", "OutFilePath", "C:\Temp")
  
  chkFileBySong.Value = GetSetting(App.EXEName, "Options", "BySong", vbUnchecked)
  
  optSupplyFilename.Value = GetSetting(App.EXEName, "Options", "UsePrefix", vbUnchecked)
  sWork = GetSetting(App.EXEName, "Options", "TextPrefix", "")
  If sWork = "" Then
    optSupplyFilename.Caption = USE_FILE_PREFIX & NOT_SET
    optSupplyFilename.Value = vbUnchecked
  Else
    optSupplyFilename.Caption = sWork
    optSupplyFilename.Value = vbChecked
  End If
  
  optTimedStart.Value = GetSetting(App.EXEName, "Options", "StartOnTime", vbUnchecked)
  gsStationFormat = GetSetting(App.EXEName, "Options", "StationFormat", "")
  NewShoutcastStation(0).StationName = GetSetting(App.EXEName, "Options", "StationName", "")
  NewShoutcastStation(0).ID = GetSetting(App.EXEName, "Options", "SCID", "")
  
  sliVolume.Value = GetSetting(App.EXEName, "Options", "PlayVolume", 50)
  sliBalance.Value = GetSetting(App.EXEName, "Options", "Balance", 0)
  
  mnuSRR.Checked = GetSetting(App.EXEName, "BootShut", "Reboot", vbUnchecked)
  mnuSRS.Checked = GetSetting(App.EXEName, "BootShut", "Shutdown", vbUnchecked)
  
  bNowRecording = False  ' Initially not recording.
  gStartOnTime = False  ' Timer set and waiting for the time.
  gbDataComing = False  ' Nothing coming in at the moment.  But that is soon to change, we hope.
  
  lblCurrTime = "It is now " & Now
  lblRadio = "Idle state.  Not connected, not waiting to record."

 'Now, read all of the saved Memorized Stations into memory array.
  sWork = Dir$(App.Path & "\SavedParms.txt")
  If sWork <> "" Then
    iFile = FreeFile
    Open App.Path & "\SavedParms.txt" For Input Access Read As iFile
    Line Input #iFile, sWork
   
   'OK, so I forgot to put in the file version on the first go around.  I actually thought
   'about it but decided to leave it out.  I thought  "...I won't ever need this..." but
   'WRONG AGAIN, SOFTWARE BREATH!
   'So version 1 parms file will not have a V in the first byte.  All others will.
   'If there is a V, then save it off and read the next line in so the rest of the routine
   'will be in sync.  Otherwise, use what we read in from the first line, it is a V1 file
   'and that first line is the Stations Count.
   
    If Left(sWork, 1) = "V" Then  ' If this is a file version greater than 1, it will have a "V" followed by a number.
      sParmsFileVersion = Val(Mid$(sWork, 2))
      Line Input #iFile, sWork  ' Now read the number of streams in the file to get ready for the loop.
    Else
      sParmsFileVersion = 1  ' If no "V", then this is a version 1 file and won't have the UseICYSongTitle check field in it.
    End If
    giMemorizedStationsCt = Val(sWork)
    
    For i = 1 To giMemorizedStationsCt And Not EOF(iFile)  ' The EOF is a little safety valve for corrupt/incomplete file.
      
      iInternalCount = iInternalCount + 1  ' Count this set.
      
      Line Input #iFile, gtMemorizedStations(i).StationName
      Line Input #iFile, gtMemorizedStations(i).MyFilePrefix
      Line Input #iFile, gtMemorizedStations(i).Format
      Line Input #iFile, gtMemorizedStations(i).ID
      Line Input #iFile, gtMemorizedStations(i).URL
      Line Input #iFile, gtMemorizedStations(i).BitRate
      Line Input #iFile, gtMemorizedStations(i).Genre
      Line Input #iFile, gtMemorizedStations(i).StartDate
      Line Input #iFile, gtMemorizedStations(i).StartHour
      Line Input #iFile, gtMemorizedStations(i).StartMin
      Line Input #iFile, gtMemorizedStations(i).Duration
      If sParmsFileVersion > 1 Then  ' Version 2 file has an extra field per stream in it.
        Line Input #iFile, sWork
        gtMemorizedStations(i).UseICYSongTitle = Val(sWork)
      End If
    
    Next
    Close iFile
    If gbDebugLogic Then Debug.Print MyTime() & "Expecting " & giMemorizedStationsCt & " and read " & iInternalCount & "."
    If giMemorizedStationsCt <> iInternalCount Then _
      MsgBox "The memorized stations file is quite probably corrupted.  Only " & iInternalCount & " of the expected " & giMemorizedStationsCt & " stream sets were found."
      
    gtMemorizedStations(0).dtAutoRecordStart = CDate("12/31/2999 23:59:59")
    gtMemorizedStations(0).MyFilePrefix = optSupplyFilename.Caption
    
    If giForceTimed = F_None Then
      chkRecordAllMemorized.Value = GetSetting(App.EXEName, "Options", "RecordAll", vbUnchecked)  ' Will run FindNextToRecord if needed.
    Else
      If giForceTimed = F_On Then
        chkRecordAllMemorized.Value = vbChecked
      Else
        chkRecordAllMemorized.Value = vbUnchecked
      End If
    End If
  End If
  
  gbIgnoreClicks = False
  
End Sub
Sub StartRecording()
      
  If gbDebugLogic Then Debug.Print MyTime() & "Starting Recording"

  If CreateOutFile Then
    cmdStartRecording.Enabled = False
    cmdConnectToServer.Enabled = False
    cmdSelectMemory.Enabled = False
    cmdStopRecording.Enabled = True
    lblStatus.Visible = True
    
    lblTitle.Enabled = True
    lblFile.Enabled = True
    lblBitrate.Enabled = True
    lblSize.Enabled = True
    lblTime.Enabled = True
    
    bNowRecording = True
    gdtRecordStartTime = Now
    
    giEnablePlayRecording = 0  ' After 5 seconds, can enable the play button
    gbNowPlaying = False  ' Just started recording. Can't play yet.  Need some buffer.
    gbNowRecording = True
  Else
    DisconnectFromServer
  End If

End Sub

Sub StopRecording()

  If gbDebugLogic Then Debug.Print MyTime() & "Stopping Recording"
  
  CloseOutFile
  
  cmdStartRecording.Enabled = True
  cmdConnectToServer.Enabled = True
  cmdSelectMemory.Enabled = True
  cmdStopRecording.Enabled = False
  
  lblStatus.Visible = False  ' Turn off REC "light"
  
  lblTitle.Enabled = False
  lblTitle = "Title: n/a"
  
  lblFile.Enabled = False
  lblFile = "File: n/a"
  
  lblBitrate.Enabled = False
  lblBitrate = "Bitrate: n/a"
 
  lblSize.Enabled = False
  lblSize = "Size: n/a"
  
  lblTime.Enabled = False
  lblTime = "Running time: n/a"
  
  bNowRecording = False
  
  lblRadio = "Recording stopped at: " & Now
    
  If gbStopContinuousPlay Then
    fraPlayControl.Visible = False
    lblPlayControl.Visible = False
    lblPlayControlsComing.Visible = True
  End If
  
End Sub

Sub StopWaiting()

  gStartOnTime = False

  cmdConnectToServer.Enabled = True
  cmdSelectMemory.Enabled = True
  
  lblRadio = "Timed record wait cancelled at " & Now
  
End Sub

Public Sub Tune(ByVal ServerIP As String, ByVal Port As Long, Optional sPath As String)
  
  Dim sCompletePath As String
  
  IcyReceived = False
  sIcyHeader = ""
  sData = ""
  sMeta = ""
  DataLen = 0
  MetaLen = 0
  nData = 0
  bMeta = False
  giFile = 0
  gFileLen = 0
  gsPath = IIf(sPath = "", "/", sPath)
  
  If sPath <> "" Then ServerIP = ServerIP & "/" & sPath
  
  wsShoutcastReceiver.Close
  wsShoutcastReceiver.Connect ServerIP, Port

End Sub

Private Sub chkMutePlayback_Click()
  
  If chkMutePlayback.Value = vbChecked Then
    WMP.settings.mute = True
  Else
    WMP.settings.mute = False
  End If
  
End Sub

Private Sub chkPausePlayback_Click()

  If chkPausePlayback.Value = vbChecked Then
    WMP.Controls.pause
  Else
    WMP.Controls.play
  End If

End Sub

Private Sub chkPlayStop_Click()
  
  If chkPlayStop.Value = vbChecked Then
    gbStopContinuousPlay = False  ' Allow continuous play
    StartPlaying
  Else
    gbStopContinuousPlay = True  ' No More playing, please.
    StopPlaying
  End If
  
End Sub

Private Sub chkRecordAllMemorized_Click()

  If gbNowRecording Then Exit Sub  ' If I am now recording, the buttons will be handled later by someone else.
  
  If chkRecordAllMemorized.Value = vbChecked Then
    cmdConnectToServer.Enabled = False
    FindNextToRecord
  Else
    cmdConnectToServer.Enabled = True
    cmdSelectMemory.Enabled = True
    lblRadio = "Idle state.  Not connected, not waiting to record."
  End If
  
End Sub

Private Sub cmdBFF_Click()

  On Error GoTo UserCancelError
 
 'set object, i.e., when the geflagamach is engaged, the remflaglation enunsifies the malitunjite and the whole thing gets refalgomized.
 'Clear now?  If so, you are as sick as I am!  I feel for you...  GET HELP!!!   NOW!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
 
  Set ShBFF = SH.BrowseForFolder(hwnd, "Select Folder for Saving Shoutcast Stream Files.", 1)
  
  With ShBFF.Items.Item
  Debug.Print MyTime() & ShBFF.Items.Item.Name
 'get folder props
  
  If .Path = "" Then
    txtOutPath = "C:\"
  Else
    txtOutPath = .Path
    On Error GoTo BadOutPath
    ChDrive Left(txtOutPath, 1)
    ChDir txtOutPath
    On Error GoTo 0
  End If
  If gbDebugLogic Then _
    Debug.Print MyTime() & "Name: " & .Name & vbCrLf & "Type: " & .Type & vbCrLf & "Last Modified: " & .ModifyDate & vbCrLf & "Parent: " & .Parent & vbCrLf
  End With
  On Error GoTo 0
  Exit Sub
  
BadOutPath:
  On Error GoTo 0
  MsgBox "Directory Error TOP1 Cannot find the folder " & txtOutPath & ".  Please change the path or create it."
  Exit Sub

UserCancelError:
  On Error GoTo 0
  MsgBox "User cancel or filesystem error UCE1.  Path not changed."
  Exit Sub

End Sub

Private Sub cmdConnectToServer_Click()
  ConnectToServer
End Sub

Private Sub cmdDisconnectFromServer_Click()
  DisconnectFromServer
End Sub

Private Sub cmdExit_Click()
  DoExit
End Sub

Sub DoExit()

  gbStopAllFunctions = True  ' Going home
  If bNowRecording Then StopRecording
  WMP.Controls.stop: Wait 50
  Unload Me
  
End Sub
Sub StartPlaying()

  Dim i As Integer
  
  If gbDebugLogic Then Debug.Print MyTime() & "Trying for: " & gsCurrentOutputFile
  gbNowPlaying = True  ' Stop the button from being re-enabled.
  sliVolume.Enabled = True
  sliBalance.Enabled = True
  
 'Play first song.  The rest will be pushed along by WMP.PlayStateChange
  
  If gbDebugLogic Then Debug.Print MyTime() & "Playing first song: " & sPlayFilesQueue(0)
  WMP.URL = sPlayFilesQueue(0): Wait 20
  WMP.Controls.play: Wait 20
  
  If iPlayFilesQueue > 0 Then  ' Obvious maybe but this overlays the file in item 0 and copies all the rest down one toward 0.  File 0 played next.
    If gbDebugLogic Then Debug.Print MyTime() & "There are " & iPlayFilesQueue & " songs in queue."
    For i = 0 To iPlayFilesQueue - 1
      sPlayFilesQueue(i) = sPlayFilesQueue(i + 1)
      sPlayFilesQueue(i + 1) = ""
      If gbDebugLogic Then Debug.Print MyTime() & "PlayNextSong: Queue item " & i & ": " & sPlayFilesQueue(i)
    Next
    iPlayFilesQueue = iPlayFilesQueue - 1
  End If

End Sub

Private Sub cmdFakeLoad_Click()

  Dim i As Long
  Dim j As Long
  Dim k As Long
  
  j = GetTickCount
  For i = 1 To 100000
    DoEvents
    k = GetTickCount
    If k <> j Then
      Debug.Print i, j, k
      j = k
    End If
  Next
  
  Debug.Print MyTime() & "There are " & iPlayFilesQueue & " songs in queue now.  Adding....."
  fraPlayControl.Visible = True
  gbStopContinuousPlay = False
  
  sPlayFilesQueue(iPlayFilesQueue) = "C:\Shoutcasts\Ron Grainer - Dr. Who Theme.aac"
  iPlayFilesQueue = iPlayFilesQueue + 1
  sPlayFilesQueue(iPlayFilesQueue) = "C:\Shoutcasts\Jwwwthis is.aac"
  iPlayFilesQueue = iPlayFilesQueue + 1
  sPlayFilesQueue(iPlayFilesQueue) = "C:\Shoutcasts\Ron Grainer - Dr. Who Theme.aac"
  iPlayFilesQueue = iPlayFilesQueue + 1
  sPlayFilesQueue(iPlayFilesQueue) = "C:\Shoutcasts\Jwwwthis is.aac"
  iPlayFilesQueue = iPlayFilesQueue + 1
  
  Debug.Print MyTime() & "There are " & iPlayFilesQueue & " songs in queue now.  Done."
 
 End Sub

Private Sub cmdSelectMemory_Click()
  
 'tmrUpdateAll.Enabled = False  ' Waste of CPU time and screws up debugging.
  
  optTimedStart.Caption = "Start recording at: " & NOT_SET
  optTimedStart.Value = vbUnchecked
  optTimedStop.Caption = "Stop recording at: " & NOT_SET
  optTimedStop.Value = vbUnchecked

  frmStationMemory.Show vbModal
  
  txtServer = gtMemorizedStations(iFinallyTheStationToTune).URL
  gsStationFormat = gtMemorizedStations(iFinallyTheStationToTune).Format
  chkFileBySong.Value = gtMemorizedStations(iFinallyTheStationToTune).UseICYSongTitle
  
  If gtMemorizedStations(iFinallyTheStationToTune).MyFilePrefix <> "" Then
    gtMemorizedStations(iFinallyTheStationToTune).MyFilePrefix = MultiReplace(gtMemorizedStations(iFinallyTheStationToTune).MyFilePrefix, USE_FILE_PREFIX, "")
    optSupplyFilename.Caption = USE_FILE_PREFIX & gtMemorizedStations(iFinallyTheStationToTune).MyFilePrefix
    optSupplyFilename.Value = vbChecked
  End If
  
 'tmrUpdateAll.Enabled = True

  If chkRecordAllMemorized.Value = vbChecked Then FindNextToRecord  ' Only do this if requested.

End Sub

Private Sub cmdStartRecording_Click()
  StartRecording
End Sub

Sub StopPlaying()

  gbNowPlaying = False
  
  WMP.Controls.stop
  
  sliVolume.Enabled = False
  sliBalance.Enabled = False

End Sub
Private Sub cmdStopRecording_Click()
  StopRecording
End Sub

Private Sub Form_Load()
    
  gbDebugLogic = True  ' Controls lots of Debug.Print statements.
  gbDebugFile = False  ' Controls one more to create/not create the file "Received Data.txt " of received data for debugging.
  gbDebugBuffer = False  ' Dump ProcessBuffer packets or not.
  
  gbIgnoreClicks = True
  
  Me.Caption = Me.Caption & " - v" & App.Major & "." & App.Minor & "." & App.Revision
  
  gbQuickExit = False
  ProcessCommands
  
  InitVars
  
  gbIgnoreClicks = False
  
  DO_CtrlOutline Me
  
  If gbMinimizeMe And giForceTimed = F_On Then Me.WindowState = vbMinimized
  
End Sub
Sub ProcessCommands()

  Dim sCommands As String
  Dim aCommands() As String
  Dim aSplit() As String
  Dim i As Integer
  
  sCommands = UCase(Command$)
  aCommands = Split(sCommands, "-")
  If gbDebugLogic Then Debug.Print "There is " & UBound(aCommands) & " option on the command line."
  If UBound(aCommands) < 0 Then Exit Sub
  
  giForceTimed = F_None  ' Set the default
  giSavedStationNo = -1  ' An impossible station to tune so is a marker that none has been set.
  
  For i = 1 To UBound(aCommands)
    
    If gbDebugLogic Then Debug.Print "Parm " & i & ": " & aCommands(i)
    
    aSplit = Split(aCommands(i), "=")
    If gbDebugLogic Then Debug.Print "Found " & UBound(aSplit) & " parm on " & aSplit(0) & " option."
    
    If UBound(aSplit) < 1 Then
      MsgBox "There is an error in the input parms near -" & aSplit(0) & ". No option specified.  Please correct and rerun."
      gbQuickExit = True
      Unload Me
      End
    End If
    
    If UBound(aSplit) > -1 Then
      If gbDebugLogic Then Debug.Print "Processing option " & aSplit(0)
    End If
    
    Select Case Left(aSplit(0), 1)
      Case "F"
        If Left(aSplit(1), 2) = "ON" Then
          giForceTimed = F_On
          If gbDebugLogic Then Debug.Print "Set RecordNow"
        ElseIf Left(aSplit(1), 2) = "OF" Then
          giForceTimed = F_Off
          If gbDebugLogic Then Debug.Print "Clear RecordNow"
        Else
          MsgBox "Error: Unrecognized option " & aSplit(1) & " on parm " & aSplit(0) & "."
          gbQuickExit = True
          Unload Me
          End
        End If
      Case "S"
        If Not IsNumeric(aSplit(1)) Then
          MsgBox "S option found with non-numeric parm.  Please correct and rerun."
          gbQuickExit = True
          Unload Me
          End
        End If
        giSavedStationNo = Val(aSplit(1))
        If gbDebugLogic Then Debug.Print "S Parm directs tuning to stream # " & giSavedStationNo
      Case "M"
        gbMinimizeMe = False
        If Left(aSplit(1), 1) = "Y" Then gbMinimizeMe = True
      Case Else
        MsgBox "Unknown command line parameter '" & aSplit(0) & "'."
        gbQuickExit = True
        Unload Me
        End
    End Select
  Next
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Dim i As Long
  Dim iFile As Integer
  Dim sFile As String
  
  If gbQuickExit Then Exit Sub
  
  wsShoutcastReceiver.Close
  
  SaveSetting App.EXEName, "Form", "frmMain_Top", Me.Top
  SaveSetting App.EXEName, "Form", "frmMain_Left", Me.Left
  
  SaveSetting App.EXEName, "LastTuned", "IP", txtServer
  
  SaveSetting App.EXEName, "Options", "OutFilePath", txtOutPath
  
  SaveSetting App.EXEName, "Options", "BySong", chkFileBySong.Value
  
  SaveSetting App.EXEName, "Options", "UsePrefix", optSupplyFilename.Value
  SaveSetting App.EXEName, "Options", "TextPrefix", optSupplyFilename.Caption
  
  SaveSetting App.EXEName, "Options", "StartOnTime", optTimedStart.Value
  SaveSetting App.EXEName, "Options", "StationFormat", gsStationFormat
  SaveSetting App.EXEName, "Options", "StationName", NewShoutcastStation(0).StationName
  SaveSetting App.EXEName, "Options", "SCID", NewShoutcastStation(0).ID
  
  SaveSetting App.EXEName, "Options", "PlayVolume", sliVolume.Value
  SaveSetting App.EXEName, "Options", "Balance", sliBalance.Value
  SaveSetting App.EXEName, "Options", "RecordAll", chkRecordAllMemorized.Value
  
  SaveSetting App.EXEName, "BootShut", "Reboot", mnuSRR.Checked
  SaveSetting App.EXEName, "BootShut", "Shutdown", mnuSRS.Checked
  
  iFile = FreeFile()
  Open App.Path & "\SavedParms.$$new$$" For Output Access Write As iFile
  
  Print #iFile, CURRENT_PARMS_VERSION  ' Save file version for posterity...
  Print #iFile, giMemorizedStationsCt  ' Let me know how many to read back next time.
  
  For i = 1 To giMemorizedStationsCt
    Print #iFile, gtMemorizedStations(i).StationName
    Print #iFile, gtMemorizedStations(i).MyFilePrefix
    Print #iFile, gtMemorizedStations(i).Format
    Print #iFile, gtMemorizedStations(i).ID
    Print #iFile, gtMemorizedStations(i).URL
    Print #iFile, gtMemorizedStations(i).BitRate
    Print #iFile, gtMemorizedStations(i).Genre
    If gtMemorizedStations(i).StartDate = NOT_SET Then gtMemorizedStations(i).StartDate = ""
    Print #iFile, gtMemorizedStations(i).StartDate
    Print #iFile, gtMemorizedStations(i).StartHour
    Print #iFile, gtMemorizedStations(i).StartMin
    Print #iFile, gtMemorizedStations(i).Duration
    Print #iFile, gtMemorizedStations(i).UseICYSongTitle
  Next
  Close iFile
  
  sFile = Dir$(App.Path & "\SavedParms.bak")
  If sFile <> "" Then ShellDeleteOne App.Path & "\SavedParms.bak", FOF_ALLOWUNDO
  
  sFile = Dir$(App.Path & "\SavedParms.txt")
  If sFile <> "" Then Name App.Path & "\SavedParms.txt" As App.Path & "\SavedParms.bak"
  
  Name App.Path & "\SavedParms.$$new$$" As App.Path & "\SavedParms.txt"

End Sub

Private Sub Form_Resize()

  'Debug.Print "Height: " & Me.Height
  'Debug.Print "Width:  " & Me.Width
  If Me.WindowState = vbMinimized Then Exit Sub
  
  Me.Width = 6150
  
  If Me.Height < 9560 Then Exit Sub
  
  WMP.Height = Me.Height - WMP.Top - 340
  
End Sub

Private Sub mnuFileExit_Click()
  DoExit
End Sub

Private Sub mnuSRR_Click()

  If Not mnuSRR.Checked Then
    If MsgBox("Selecting this option enables rebooting after each recording completes." & vbCrLf & vbCrLf & "Are you sure this is what you want?", vbYesNo) = vbYes Then
      mnuSRR.Checked = True
      mnuSRS.Checked = False
    End If
  Else
    MsgBox "Rebooting after recording cancelled."
    mnuSRR.Checked = False
  End If
  
End Sub

Private Sub mnuSRS_Click()
  
  If Not mnuSRS.Checked Then
    If MsgBox("Selecting this option enables computer shutdown after the next recording completes." & vbCrLf & vbCrLf & "Are you sure this is what you want?", vbYesNo) = vbYes Then
      mnuSRS.Checked = True
      mnuSRR.Checked = False
    End If
  Else
    MsgBox "Shutdown after recording cancelled."
    mnuSRS.Checked = False
  End If
  
End Sub

Private Sub optTimedStart_Click()
  If optTimedStart.Value <> vbChecked Then gStartOnTime = False
End Sub

Private Sub wsShoutcastReceiver_Close()
  
  lblTitle = "Title: n/a"
  lblFile = "File: n/a"
  lblBitrate = "Bitrate: n/a"
  lblSize = "Size: n/a"
  
  cmdSelectMemory.Enabled = True
  cmdConnectToServer.Enabled = True
  cmdStartRecording.Enabled = False
  cmdDisconnectFromServer.Enabled = False
  
  lblRadio.Caption = "Disconnected"
  gbDataComing = False

End Sub

Private Sub wsShoutcastReceiver_Connect()
    
  Dim sRequest As String
  
  lblRadio = "Connected"
  
  cmdStartRecording.Enabled = True
  cmdConnectToServer.Enabled = False
  cmdSelectMemory.Enabled = True
  cmdDisconnectFromServer.Enabled = True
  
  sRequest = Replace$(ReqHeader, "%", wsShoutcastReceiver.RemoteHostIP)
  sRequest = Replace$(sRequest, "$", gsPath)
  wsShoutcastReceiver.SendData sRequest

End Sub

Private Sub wsShoutcastReceiver_ConnectionRequest(ByVal requestID As Long)
  lblRadio = "Connecting ..."
End Sub

Private Sub wsShoutcastReceiver_DataArrival(ByVal bytesTotal As Long)
    
  Dim pos As Long
  Dim pos2 As Long
  Dim seconds As String

  If Not gbDataComing Then  ' Success is so sweet.  Tell my user that things have started to happen and he should enjoy the stream.
    gbDataComing = True
    cmdConnectToServer.Enabled = False
    cmdSelectMemory.Enabled = True
  End If
  
  giBuffersReceived = giBuffersReceived + 1
  giCurrentInBuffer = giCurrentInBuffer + 1
  If giCurrentInBuffer = 11 Then giCurrentInBuffer = 1
  
  wsShoutcastReceiver.GetData sInBuffer(giCurrentInBuffer), , bytesTotal
  If gbDebugBuffer Then Debug.Print MyTime() & giCurrentInBuffer & "-" & giBuffersReceived & ": Received buffer length: " & bytesTotal
  
  If Not IcyReceived Then
    sData = sData & sInBuffer(giCurrentInBuffer)
    pos = InStr(1, sData, vbCrLf & vbCrLf)
    
    If pos > 0 Then
      If InStr(1, sData, "ICY 200 OK") Then
        sIcyHeader = Left(sData, pos + Len(vbCrLf & vbCrLf) - 1)
       'Find "metaint" in input stream
        pos = InStr(1, sData, "icy-metaint:") + Len("icy-metaint:")
        pos2 = InStr(pos, sData, vbCrLf)
        DataLen = Mid(sData, pos, pos2 - pos + 1)
        sInBuffer(giCurrentInBuffer) = Mid(sData, Len(sIcyHeader) + 1)
        ShowInfo sIcyHeader
        If gbDebugLogic Then Debug.Print MyTime() & "Icy Header: " & sIcyHeader
        IcyReceived = True
        bMeta = False
      End If
   'Else  ' No longer needed.  No code for the "Else"
   'Some time out waiting for Icy header??
    End If
  End If


  If IcyReceived Then
    
    If gbDebugBuffer Then Debug.Print MyTime() & "Buffer size: " & Len(sInBuffer(giCurrentInBuffer))
    While sInBuffer(giCurrentInBuffer) <> ""
      sInBuffer(giCurrentInBuffer) = ProcessBuffer(sInBuffer(giCurrentInBuffer), bMeta)
    Wend
    
  End If

  If giFile Then
    seconds = DateDiff("s", dtStartTime, Now)
    lblSize = "Size: " & Format$(gFileLen \ 1024, "###,###,##0") & " kb"
    lblTime = "Time: " & (seconds \ 60) & ":" & Format((seconds Mod 60), "0#")
    If optTimedStop.Value = vbChecked And Now > gtMemorizedStations(giNextToRecord).dtAutoRecordStop Then
      gbStopContinuousPlay = True
      DisconnectFromServer  ' Will do a StopRecording, too.
      If chkRecordAllMemorized.Value = vbChecked Then
        FindNextToRecord
        If TimeToRecordSomething > 5 Then
          bExitNow = False  ' Will be massaged by one of the following forms.
          If mnuSRS.Checked Then frmShutdown.Show vbModal
          If mnuSRR.Checked Then frmReboot.Show vbModal
          If bExitNow Then
            Close
            Unload Me
            End
          End If
        End If
      Else
        bExitNow = False  ' Will be massaged by one of the following forms.
        If mnuSRS.Checked Then frmShutdown.Show vbModal
        If mnuSRR.Checked Then frmReboot.Show vbModal
        If bExitNow Then
          Close
          Unload Me
          End
        End If
      End If
    End If
  End If
  
End Sub

Function ProcessBuffer(ByVal sBuffer As String, ByRef esMeta As Boolean) As String
    
  Dim Remain As Long
  
  If gbDebugBuffer Then Debug.Print MyTime() & "ProcessBuffer got: " & sBuffer
 
  If esMeta = False Then  ' Incoming buffer is data
    Remain = DataLen - nData
    If (Remain <= Len(sBuffer)) Then
      nData = nData + Remain
      Call WriteOutFile(Left(sBuffer, Remain))
      nData = 0
      esMeta = True
      ProcessBuffer = Mid(sBuffer, Remain + 1)
    Else
      nData = nData + Len(sBuffer)
      Call WriteOutFile(sBuffer)
      ProcessBuffer = ""
    End If
            
  Else  ' Incoming buffer is metadata
  
    If MetaLen = 0 Then
      'get length of metadata (first byte of block * 16)
      MetaLen = Asc(Left(sBuffer, 1)) * 16
    End If
    
    Remain = MetaLen - Len(sMeta)
    
    If Remain = 0 Then
      esMeta = False
      ProcessBuffer = Mid(sBuffer, 2)
    
    ElseIf Remain <= Len(sBuffer) Then
      
      sMeta = sMeta & Mid(sBuffer, 2, Remain)
      
      ShowTitle sMeta
      If gbDebugBuffer Then Debug.Print MyTime() & "sMeta has: " & sMeta
      sMeta = ""
      MetaLen = 0
      esMeta = False
      ProcessBuffer = Mid(sBuffer, Remain + 2)
  
    Else
        
      sMeta = sMeta & sBuffer
      ProcessBuffer = ""
    
    End If
  
  End If
  
End Function

Private Sub wsShoutcastReceiver_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  
  lblRadio = "Winsock reported that an error " & Number & " ocurred -- " & Description & "--" & Scode & "--" & Source
  Debug.Print MyTime() & lblRadio
  
End Sub

Sub ShowTitle(ByVal sMetaData As String)
  
  Dim pos As Long, pos2 As Long
  
 'The title is only sent once, at the start of a new file (song).  Display it.
 'And, if recording by file, change it now.
 
 'Find the stream (song) title in the incoming data.
  pos = InStr(1, sMetaData, "StreamTitle=") + Len("StreamTitle=")
  If pos > 0 Then
    pos2 = InStr(pos, sMetaData, ";")
    lblTitle = "Title: " & Mid(sMetaData, pos, pos2 - pos)
    If Len(lblTitle) = 9 Then lblTitle = "Title: " & "Title not available"
    If gbDebugLogic Then Debug.Print MyTime() & "New Title received: " & lblTitle
  Else
    lblTitle = "Title: " & "Title not available"
  End If
  
  If bNowRecording And chkFileBySong.Value = vbChecked Then
    CloseOutFile
    CreateOutFile
  End If
    
End Sub

Sub ShowInfo(ByVal sIcy As String)
    
  Dim pos As Long
  Dim pos2 As Long
  
 'seek station name
  pos = InStr(1, sIcy, "icy-name:") + Len("icy-name:")
  If pos > 0 Then
    pos2 = InStr(pos, sIcy, vbCrLf)
    lblRadio = Mid(sIcy, pos, pos2 - pos + 1)
  Else
    lblRadio = "Unidentified station"
  End If
  
 'seek bit rate
  pos = InStr(1, sIcy, "icy-br:") + Len("icy-br:")
  If pos > 0 Then
    pos2 = InStr(pos, sIcy, vbCrLf)
    lblBitrate = "Bitrate: " & Mid(sIcy, pos, pos2 - pos) & " Kbps"
  Else
    lblBitrate = "Bitrate: " & "Unspecified bitrate"
  End If
  
End Sub

Function CreateOutFile() As Boolean
    
  Dim sPath As String
  Dim sFile As String
  Dim sTemp As String
  Dim iAns As Integer
  Dim sFileExt As String
  
  If gbDebugLogic Then Debug.Print MyTime() & "Proposed stream format: " & gsStationFormat
  Select Case gsStationFormat
    Case MP3Specifier
      sFileExt = MP3FileType
    Case AACSpecifier
      sFileExt = AACFileType
    Case Else
      sFileExt = ".UNK"
      iAns = MsgBox("Program error COF1 in CreateOutFile. Unknown data type: '" & gsStationFormat & _
             "' The program must be updated to receive this type of stream." & vbCrLf & vbCrLf & _
             "Would you like to continue with a filetype of '.UNK' and change it later? " & _
             "(The stream will not be playable by this program.)", vbYesNo, "Unknown Stream Type")
       If iAns = vbNo Then
         chkRecordAllMemorized.Value = vbUnchecked
         Exit Function
       End If
      CreateOutFile = False
      Exit Function
  End Select
  
  If giFile <> 0 Then CloseOutFile
  
  gFileLen = 0
  If chkFileBySong.Value = vbUnchecked Then  ' Do not file by song title
    
    If optSupplyFilename.Value = vbChecked Then  ' Use supplied title for all, if present
      If Right(gtMemorizedStations(giNextToRecord).MyFilePrefix, 1) <> "_" Then
        sTemp = gtMemorizedStations(giNextToRecord).MyFilePrefix & "_"
      Else
        sTemp = gtMemorizedStations(giNextToRecord).MyFilePrefix
      End If
      sFile = sTemp & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Format(Now, "hhmmss") & sFileExt
    Else
     'Set a default name for quick acceptance
      sFile = "File_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Format(Now, "hhmmss") & sFileExt
      sFile = InputBox("Enter file name: ", "New filename", sFile)
    End If
    If sFile = "" Then Exit Function
  Else  ' It IS checked.  Use songnames for filenames.
    sTemp = Trim$(Replace(lblTitle, "Title:", ""))
    sFile = Replace(sTemp, ":", ".") & sFileExt
    If sFile = "" Then sFile = "Initial Partial Song Capture" & sFileExt
  End If
  
  sFile = Replace(sFile, USE_FILE_PREFIX, "")  ' This is a little fixup done for manual recording.  It is a little dirty, clean it up!
  
  If Left(sFile, 1) = "'" Then  ' Leading Quote, probably a trailing one, too.
    sFile = Mid$(sFile, 2)
    If Right(sFile, 5) = "'" & sFileExt Then
      sFile = Left(sFile, Len(sFile) - 5) & sFileExt
    End If
  End If
  
  giFile = FreeFile()
  sPath = txtOutPath
  If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
  If Dir$(sPath & sFile) <> "" Then
    iAns = MsgBox("Output file " & sFile & " already exists, do you want to delete and overwrite it?", vbYesNo, "File Replace Verification")
    If iAns = vbNo Then
      CreateOutFile = False
      Exit Function
    End If
    Kill sPath & sFile
  End If
  
 'Fix up the filename to please the File System.
  sFile = Replace(sFile, "/", "-")
  sFile = Replace(sFile, "\", "-")
  sFile = Replace(sFile, ":", "-")
  sFile = Replace(sFile, "*", "-")
  sFile = Replace(sFile, "?", "-")
  sFile = Replace(sFile, sQuote, "-")
  sFile = Replace(sFile, "<", "-")
  sFile = Replace(sFile, ">", "-")
  sFile = Replace(sFile, "|", "-")
  
  Open sPath & sFile For Binary Access Write As #giFile
  If gbDebugLogic Then Debug.Print MyTime() & "CreateOutFile: Opening " & sPath & sFile & " for saving song."
  sPlayFilesQueue(iPlayFilesQueue) = sPath & sFile
  iPlayFilesQueue = iPlayFilesQueue + 1
  If iPlayFilesQueue > UBound(sPlayFilesQueue) Then ReDim Preserve sPlayFilesQueue(UBound(sPlayFilesQueue) + 100)
  
  If gbDebugLogic Then Debug.Print MyTime() & "CreateOutFile: Items in the play queue = " & iPlayFilesQueue
  If gbDebugLogic Then Debug.Print MyTime() & "---------------"
  
  lblFile = "File: " & sFile
  lblFile.ToolTipText = sFile  ' Just in case it is very long and overflows the screen.
  
  gsCurrentOutputFile = sPath & sFile
  
  CreateOutFile = True

  dtStartTime = Now

End Function

Sub WriteOutFile(ByVal sBuff As String)
    
    If giFile = 0 Then Exit Sub
    Put #giFile, , sBuff
    
    gFileLen = gFileLen + Len(sBuff)
    
End Sub

Sub CloseOutFile()
  
  Close #giFile
  giFile = 0
  
  gbDataComing = False
  gbNowRecording = False
  
End Sub

Private Sub SliBalance_Change()

  WMP.settings.balance = sliBalance.Value  ' -100 left, 0 middle, 100 right

End Sub

Private Sub sliVolume_Change()

  Dim d As Double
  
 'd = CDbl(sliVolume.Value)
  
 'The documentation is very strange on this one.  One place says 0-10, another 0-100 and one says double 0 to 1.
 'I don't know but this works.  Close enough for now.
 'I just set a slider with 0 to 100 and it seems to work just fine.  Beats me!!  If you figure it out, let me know.
 
  WMP.settings.volume = sliVolume.Value
  
End Sub

Private Sub tmrUpdateAll_Timer()

  Dim i As Integer
  
  lblCurrTime = "It is now " & Now
  If gbDebugBuffer And giBuffersReceived > 0 Then Debug.Print MyTime() & "Received buffers: " & giBuffersReceived
  
 'Here is the routine to flash the big rec "REC" label.
  If bNowRecording Then
    If lblStatus.Visible Then
      lblStatus.Visible = False
    Else
      lblStatus.Visible = True
    End If
    
   'Wait for 3 seconds before enabling the play button for adequate buffering.
   'Note: If the recording is stopped by the usser,
    giEnablePlayRecording = giEnablePlayRecording + 1
    If giEnablePlayRecording > 2 And Not gbNowPlaying Then
      fraPlayControl.Visible = True
      lblPlayControl.Visible = True
      lblPlayControlsComing.Visible = False
    End If
      
 'Control only gets to this ElseIf if I am not NowRecording.  And only continues if the user wants recording and there is one to record.
  ElseIf chkRecordAllMemorized.Value = vbChecked And giNextToRecord > 0 Then ' And Not NowRecording Then ' Reduce timer requirement by 50%.  Do everything here.
  
   'NOTE: The ending time is detected in the Winsock (wsShoutcastReceiver) data arrival event.
    i = TimeToRecordSomething ' Don't care about the return code here.
  End If  ' There is something to record and we are not recording.  See if it is time to record.
  
End Sub
Function TimeToRecordSomething() As Long

  TimeToRecordSomething = -1  ' Assume there is nothing to record.
  If Now() >= gtMemorizedStations(giNextToRecord).dtAutoRecordStart Then  ' If the time to record has passed...
    If Now() < gtMemorizedStations(giNextToRecord).dtAutoRecordStop Then  ' ...and time to end has not, we are in the window, start recording.
      If gbDataComing Then
        StartRecording  ' If data is here, we can start recording it.
      Else
        ConnectToServer  ' This will eventually, (we hope) set gbDataComing
      End If  ' Got to connect first, get data coming, then start recording.
    End If  ' If the record stop time is in front of us...
  End If  ' If the record start time is behind us...

  TimeToRecordSomething = DateDiff("n", Now(), gtMemorizedStations(giNextToRecord).dtAutoRecordStart)  ' Minutes till next rec start

End Function

Private Sub txtOutPath_KeyDown(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
    Case 35, 36, 37, 39
    Case Else
      KeyCode = 0
  End Select
  
End Sub

Private Sub txtOutPath_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub txtOutPath_LostFocus()
  
 'NOTE: I tried using Dir$ here, first, but it failed on empty directories.
 '      I could probably could have added the specifier for looking for directories but
 '      I just changed to this method instead (using ChDrive and ChDir).  It may be better
 '      to do it this way, anyway.  Yeah, I know I could have created it.  That's what BrowseForFolder is for!  GO AWAY!!!
 
  If Right(txtOutPath, 1) <> "\" Then txtOutPath = txtOutPath & "\"
  
  On Error GoTo BadOutPath
  ChDrive Left(txtOutPath, 1)
  ChDir txtOutPath
  On Error GoTo 0
  Exit Sub
  
BadOutPath:
  MsgBox "Directory Error TOP1 Cannot find the folder " & txtOutPath & ".  Please change the path or create it."
  On Error GoTo 0
  Exit Sub
  
End Sub

Private Sub txtServer_KeyDown(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
    Case 35, 36, 37, 39
    Case Else
      KeyCode = 0
  End Select
  
End Sub

Private Sub txtServer_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub WMP_PlayStateChange(ByVal NewState As Long)
  
  Dim i As Integer
  
  If gbDebugLogic Then Debug.Print MyTime() & "WMP_PlayStateChange: Playstate changed to " & NewState & " -- " & sWMPStates(NewState)
  
  If gbStopAllFunctions Then Exit Sub  ' Going home

  If (WMP.playState = WMPStates.Stopped) Then   ' First or subsequent song has finshed playing.
    
    If gbDebugLogic Then Debug.Print MyTime() & "WMP_PlayStateChange: Ready for next song."
    
   'Play next song.
    
   'If Not bNowRecording Then chkPlayStop.Value = vbUnchecked
    
    If gbStopContinuousPlay Then
      If Not bNowRecording Then
        fraPlayControl.Visible = False
        lblPlayControl.Visible = False
        lblPlayControlsComing.Visible = True
        Exit Sub
      End If
    End If
  
    If Not gbStopContinuousPlay Then
      If gbDebugLogic Then Debug.Print MyTime() & "WMP.PlayStateChange setting my URL to: " & sPlayFilesQueue(0)
      WMP.URL = sPlayFilesQueue(0):
      
      For i = 1 To 50
        Wait 20    ' Setting the .URL will also start it playing if it is in right state.
        If gbDebugLogic Then Debug.Print MyTime() & "PlayStateChange spinning, waiting for Playing status.  Current status: " & sWMPStates(WMP.playState)
        If WMP.playState = WMPStates.Playing Then Exit For
        If gbStopAllFunctions Then Exit Sub
      Next
      
      If WMP.playState <> WMPStates.Playing Then
        If gbDebugLogic Then Debug.Print MyTime() & "WMP issuing Play command to self."
        WMP.Controls.play: Wait 20
      Else
        If gbDebugLogic Then Debug.Print MyTime() & "Already playing.  Duplicate play command not issued."
      End If
      
      If iPlayFilesQueue > 0 Then  ' Obvious maybe but this overlays the file in item 0 and copies all the rest down one toward 0.  File 0 played next.
        For i = 0 To iPlayFilesQueue - 1
          sPlayFilesQueue(i) = sPlayFilesQueue(i + 1)
          sPlayFilesQueue(i + 1) = ""
          If gbDebugLogic Then Debug.Print MyTime() & "PlayNextSong: Queue item " & i & ": " & sPlayFilesQueue(i)
        Next
        iPlayFilesQueue = iPlayFilesQueue - 1
      End If
      If gbDebugLogic Then Debug.Print MyTime() & "PlayStateChange says " & iPlayFilesQueue & " files still in queue."
      
    End If
  
  If gbDebugLogic Then Debug.Print MyTime() & "PlayStateChange exiting with status " & sWMPStates(NewState)

  End If
  
End Sub

Sub PlayNextSong()

  Debug.Print MyTime() & "!@#$%^&*()!@#$%^&*()_!@#$%^&*()_  SHOULD NOT BE IN PlayNextSong!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
  Exit Sub
  
  Dim i As Integer
  
 'WMP.DLL is very slow about changing files.  It goes to a Transitioning state for a file, then Ready, then Transitioning again, then....
 '... finally, Stopped State.  Then, I can tell it to play the next song.  Otherwise, it is NOT listening and won't hear the request.
 'So, I have to wait out his states with a little state follower and when he is good and ready, I tell him to play the next one.
 
  If gbDebugLogic Then Debug.Print MyTime() & "PlayNextSong: Trying to play " & sPlayFilesQueue(0) & " with " & iPlayFilesQueue & " files in queue."
  
 'NOTE: Wait does a DoEvents.  No need to do one here.
  If WMP.playState <> WMPStates.Stopped Then
    WMP.Controls.stop
    For i = 1 To 40  ' Wait up to 2 seconds for this slow poke to get ready to play the next file.  I should live so long!
      If WMP.playState = WMPStates.Stopped Then Exit For
      Wait 50  ' milliseconds.  Normally takes 2 of these, best I have seen so far.
      If gbDebugLogic Then Debug.Print MyTime() & "PlayNextSong waiting for Stopped Status - WMP Status = " & sWMPStates(WMP.playState)
    Next
  End If
  
  If Not gbStopContinuousPlay Then WMP.URL = sPlayFilesQueue(0): Wait 20    ' Setting the .URL will also start it playing if it is in right state.
 
  If gbDebugLogic Then Debug.Print MyTime() & "PlayNextSong Initial Temporary Status: WMP.PlayState = " & sWMPStates(WMP.playState)
  
  For i = 1 To 40  ' Wait up to 2 seconds for this slow poke to get ready to play the next file.  I should live so long!
    If WMP.playState = WMPStates.Playing Or WMP.playState = WMPStates.Ready Or WMP.playState = WMPStates.Stopped Then Exit For
    Wait 50  ' milliseconds.  Normally takes 2 of these, best I have seen so far.
    If gbDebugLogic Then Debug.Print MyTime() & "PlayNextSong waiting - WMP Status = " & sWMPStates(WMP.playState)
  Next
  
  If gbDebugLogic Then Debug.Print MyTime() & "PlayNextSong Second Temporary Status: WMP.PlayState = " & sWMPStates(WMP.playState) & " after " & i & " spins."
  
  DoEvents  ' Just for the hell of it.  Completely unnecessary.  (I think.  You try it.  I know, let's let Mikey try it.  He hates everything!
  
  If ((WMP.playState = WMPStates.Stopped) Or (WMP.playState = WMPStates.Ready)) And (Not gbStopContinuousPlay) Then
    WMP.Controls.play: DoEvents
    If gbDebugLogic Then Debug.Print MyTime() & "PlayNextSong: Issuing Controls.play command."
  
    If iPlayFilesQueue > 0 Then iPlayFilesQueue = iPlayFilesQueue - 1
  
    If gbDebugLogic Then Debug.Print MyTime() & "PlayNextSong: There are now " & iPlayFilesQueue & " items in the play queue."
   
    If gbDebugLogic Then Debug.Print MyTime() & "---------------"
  
  End If
 
  If iPlayFilesQueue > 0 Then  ' Obvious maybe but this overlays the file in item 0 and copies all the rest down one toward 0.  File 0 played next.
    For i = 0 To iPlayFilesQueue - 1
      sPlayFilesQueue(i) = sPlayFilesQueue(i + 1)
      sPlayFilesQueue(i + 1) = ""
      If gbDebugLogic Then Debug.Print MyTime() & "PlayNextSong: Queue item " & i & ": " & sPlayFilesQueue(i)
    Next
    iPlayFilesQueue = iPlayFilesQueue - 1
  End If
    
End Sub

