VERSION 5.00
Begin VB.Form frmStationMemory 
   Caption         =   "Saved Stations Entry & Selection"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   12975
   Begin VB.VScrollBar vsStreams 
      Height          =   6540
      LargeChange     =   20
      Left            =   12480
      TabIndex        =   253
      Top             =   720
      Width           =   315
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   19
      Left            =   8640
      TabIndex        =   252
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   7050
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   18
      Left            =   8640
      TabIndex        =   251
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   6720
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   17
      Left            =   8640
      TabIndex        =   250
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   6390
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   16
      Left            =   8640
      TabIndex        =   249
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   6060
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   15
      Left            =   8640
      TabIndex        =   248
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   5730
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   14
      Left            =   8640
      TabIndex        =   247
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   5400
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   13
      Left            =   8640
      TabIndex        =   246
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   5070
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   12
      Left            =   8640
      TabIndex        =   245
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   4740
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   11
      Left            =   8640
      TabIndex        =   244
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   4410
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   10
      Left            =   8640
      TabIndex        =   243
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   4080
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   9
      Left            =   8640
      TabIndex        =   242
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   3750
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   8
      Left            =   8640
      TabIndex        =   241
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   3420
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   7
      Left            =   8640
      TabIndex        =   240
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   3090
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   6
      Left            =   8640
      TabIndex        =   239
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   2760
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   5
      Left            =   8640
      TabIndex        =   238
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   2430
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   4
      Left            =   8640
      TabIndex        =   237
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   2100
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   3
      Left            =   8640
      TabIndex        =   236
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   1770
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   2
      Left            =   8640
      TabIndex        =   235
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   1440
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   1
      Left            =   8640
      TabIndex        =   234
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   1110
      Width           =   195
   End
   Begin VB.CheckBox chkFileBySong 
      Height          =   195
      Index           =   0
      Left            =   8640
      TabIndex        =   232
      ToolTipText     =   "Break up stream into file name by ICY title (song titles usually)"
      Top             =   780
      Width           =   195
   End
   Begin VB.CommandButton cmdShowShoutcast 
      BackColor       =   &H0080C0FF&
      Caption         =   "Select S&houtcast"
      Height          =   360
      Left            =   8580
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   230
      ToolTipText     =   "Show the Shoutcast station selection screen."
      Top             =   7500
      UseMaskColor    =   -1  'True
      Width           =   1485
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   19
      Left            =   8940
      TabIndex        =   177
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   6990
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   18
      Left            =   8940
      TabIndex        =   169
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   6660
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   17
      Left            =   8940
      TabIndex        =   161
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   6330
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   16
      Left            =   8940
      TabIndex        =   153
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   6000
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   15
      Left            =   8940
      TabIndex        =   145
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   5670
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   14
      Left            =   8940
      TabIndex        =   137
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   5340
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   13
      Left            =   8940
      TabIndex        =   129
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   5010
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   12
      Left            =   8940
      TabIndex        =   121
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   4680
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   11
      Left            =   8940
      TabIndex        =   113
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   4350
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   10
      Left            =   8940
      TabIndex        =   105
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   4020
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   9
      Left            =   8940
      TabIndex        =   97
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   3690
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   8
      Left            =   8940
      TabIndex        =   89
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   3360
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   7
      Left            =   8940
      TabIndex        =   81
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   3030
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   6
      Left            =   8940
      TabIndex        =   73
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   2700
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   5
      Left            =   8940
      TabIndex        =   65
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   2370
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   4
      Left            =   8940
      TabIndex        =   57
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   2040
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   3
      Left            =   8940
      TabIndex        =   49
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   1710
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   2
      Left            =   8940
      TabIndex        =   41
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   1380
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   1
      Left            =   8940
      TabIndex        =   33
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   1050
      Width           =   425
   End
   Begin VB.CommandButton cmdClearStationInfo 
      Caption         =   "<-->"
      Height          =   330
      Index           =   0
      Left            =   8940
      TabIndex        =   25
      ToolTipText     =   "Remove Station and Timed Record Info"
      Top             =   720
      Width           =   425
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   10
      Left            =   720
      TabIndex        =   10
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   4020
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   10
      Left            =   1080
      TabIndex        =   103
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   4020
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   255
      Index           =   11
      Left            =   720
      TabIndex        =   11
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   4350
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   11
      Left            =   1080
      TabIndex        =   111
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   4350
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   12
      Left            =   720
      TabIndex        =   12
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   4680
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   12
      Left            =   1080
      TabIndex        =   119
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   4680
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   13
      Left            =   720
      TabIndex        =   13
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   5010
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   13
      Left            =   1080
      TabIndex        =   127
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   5010
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   14
      Left            =   720
      TabIndex        =   14
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   5340
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   14
      Left            =   1080
      TabIndex        =   135
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   5340
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   15
      Left            =   720
      TabIndex        =   15
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   5670
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   15
      Left            =   1080
      TabIndex        =   143
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   5670
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   16
      Left            =   720
      TabIndex        =   16
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   6000
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   16
      Left            =   1080
      TabIndex        =   151
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   6000
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   17
      Left            =   720
      TabIndex        =   17
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   6330
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   17
      Left            =   1080
      TabIndex        =   159
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   6330
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   18
      Left            =   720
      TabIndex        =   18
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   6660
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   18
      Left            =   1080
      TabIndex        =   167
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   6660
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   19
      Left            =   720
      TabIndex        =   19
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   6990
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   19
      Left            =   1080
      TabIndex        =   175
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   6990
      Width           =   5000
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   19
      Left            =   6135
      TabIndex        =   176
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   6990
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   18
      Left            =   6135
      TabIndex        =   168
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   6660
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   17
      Left            =   6135
      TabIndex        =   160
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   6330
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   16
      Left            =   6135
      TabIndex        =   152
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   6000
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   15
      Left            =   6135
      TabIndex        =   144
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   5670
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   14
      Left            =   6135
      TabIndex        =   136
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   5340
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   13
      Left            =   6135
      TabIndex        =   128
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   5010
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   12
      Left            =   6135
      TabIndex        =   120
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   4680
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   11
      Left            =   6135
      TabIndex        =   112
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   4350
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   10
      Left            =   6135
      TabIndex        =   104
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   4020
      Width           =   2400
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   10
      Left            =   9840
      TabIndex        =   107
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   4020
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   11
      Left            =   9840
      TabIndex        =   115
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   4350
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   12
      Left            =   9840
      TabIndex        =   123
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   13
      Left            =   9840
      TabIndex        =   131
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   5010
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   14
      Left            =   9840
      TabIndex        =   139
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   5340
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   15
      Left            =   9840
      TabIndex        =   147
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   5670
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   16
      Left            =   9840
      TabIndex        =   155
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   17
      Left            =   9840
      TabIndex        =   163
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   6330
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   18
      Left            =   9840
      TabIndex        =   171
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   6660
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   19
      Left            =   9840
      TabIndex        =   179
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   6990
      Width           =   1095
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   10
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   108
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   4030
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   10
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   109
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   4030
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   11
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   116
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   4360
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   11
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   117
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   4360
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   12
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   124
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   4690
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   12
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   125
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   4690
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   13
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   132
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   5020
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   13
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   133
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   5020
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   14
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   140
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   5350
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   14
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   141
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   5350
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   15
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   148
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   5680
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   15
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   149
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   5680
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   16
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   156
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   6000
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   16
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   157
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   6000
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   17
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   164
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   6340
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   17
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   165
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   6340
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   18
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   172
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   6670
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   18
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   173
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   6670
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   19
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   180
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   6990
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   19
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   181
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   6990
      Width           =   270
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   10
      Left            =   11760
      TabIndex        =   110
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   4030
      Width           =   600
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   11
      Left            =   11760
      TabIndex        =   118
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   4360
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   12
      Left            =   11760
      TabIndex        =   126
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   4690
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   13
      Left            =   11760
      TabIndex        =   134
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   5020
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   14
      Left            =   11760
      TabIndex        =   142
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   5350
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   15
      Left            =   11760
      TabIndex        =   150
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   5680
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   16
      Left            =   11760
      TabIndex        =   158
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   17
      Left            =   11760
      TabIndex        =   166
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   6340
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   18
      Left            =   11760
      TabIndex        =   174
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   6670
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   19
      Left            =   11760
      TabIndex        =   182
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   6990
      Width           =   615
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   10
      Left            =   9375
      TabIndex        =   106
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   4020
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   11
      Left            =   9375
      TabIndex        =   114
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   4350
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   12
      Left            =   9375
      TabIndex        =   122
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   4680
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   13
      Left            =   9375
      TabIndex        =   130
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   5010
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   14
      Left            =   9375
      TabIndex        =   138
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   5340
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   15
      Left            =   9375
      TabIndex        =   146
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   5670
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   16
      Left            =   9375
      TabIndex        =   154
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   6000
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   17
      Left            =   9375
      TabIndex        =   162
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   6330
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   18
      Left            =   9375
      TabIndex        =   170
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   6660
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   19
      Left            =   9375
      TabIndex        =   178
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   6990
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   9
      Left            =   9375
      TabIndex        =   98
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   3690
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   8
      Left            =   9375
      TabIndex        =   90
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   3360
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   7
      Left            =   9375
      TabIndex        =   82
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   3030
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   6
      Left            =   9375
      TabIndex        =   74
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   2700
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   5
      Left            =   9375
      TabIndex        =   66
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   2370
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   4
      Left            =   9375
      TabIndex        =   58
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   2040
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   3
      Left            =   9375
      TabIndex        =   50
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   1710
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   2
      Left            =   9375
      TabIndex        =   42
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   1380
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   1
      Left            =   9375
      TabIndex        =   34
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   1050
      Width           =   425
   End
   Begin VB.CommandButton cmdClearTimedInfo 
      Caption         =   "-->>"
      Height          =   330
      Index           =   0
      Left            =   9375
      TabIndex        =   26
      ToolTipText     =   "Remove only Timed Record Info"
      Top             =   720
      Width           =   425
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   9
      Left            =   11760
      TabIndex        =   102
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   3700
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   8
      Left            =   11760
      TabIndex        =   94
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   3370
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   7
      Left            =   11760
      TabIndex        =   86
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   3040
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   6
      Left            =   11760
      TabIndex        =   78
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   2710
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   5
      Left            =   11760
      TabIndex        =   70
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   2380
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   4
      Left            =   11760
      TabIndex        =   62
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   2050
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   3
      Left            =   11760
      TabIndex        =   54
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   1720
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   2
      Left            =   11760
      TabIndex        =   46
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   1390
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   1
      Left            =   11760
      TabIndex        =   38
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   1060
      Width           =   615
   End
   Begin VB.TextBox txtRecordMinutes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   0
      Left            =   11760
      TabIndex        =   30
      Text            =   "0"
      ToolTipText     =   "Duration, in minutes, of the recording."
      Top             =   730
      Width           =   600
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   9
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   101
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   3700
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   9
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   100
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   3700
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   8
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   93
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   3370
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   8
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   92
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   3370
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   7
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   85
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   3040
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   7
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   84
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   3040
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   6
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   77
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   2710
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   6
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   76
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   2710
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   5
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   69
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   2380
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   5
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   68
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   2380
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   4
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   61
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   2050
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   4
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   60
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   2050
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   3
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   53
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   1720
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   3
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   52
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   1720
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   2
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   45
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   1390
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   2
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   44
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   1390
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   1
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   37
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   1060
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   1
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   36
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   1060
      Width           =   270
   End
   Begin VB.TextBox txtStartMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   0
      Left            =   11400
      MaxLength       =   2
      TabIndex        =   29
      Text            =   "00"
      ToolTipText     =   "Minute of the hour to start recording."
      Top             =   730
      Width           =   270
   End
   Begin VB.TextBox txtStartHour 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   0
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   28
      Text            =   "00"
      ToolTipText     =   "Hour to start recording."
      Top             =   730
      Width           =   270
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   9
      Left            =   9840
      TabIndex        =   99
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   3690
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   8
      Left            =   9840
      TabIndex        =   91
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   7
      Left            =   9840
      TabIndex        =   83
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   3030
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   6
      Left            =   9840
      TabIndex        =   75
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   5
      Left            =   9840
      TabIndex        =   67
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   2370
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   4
      Left            =   9840
      TabIndex        =   59
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   3
      Left            =   9840
      TabIndex        =   51
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   1710
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   2
      Left            =   9840
      TabIndex        =   43
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   1380
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   1
      Left            =   9840
      TabIndex        =   35
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   1050
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecDate 
      Caption         =   "12/31/2009"
      Height          =   330
      Index           =   0
      Left            =   9840
      TabIndex        =   27
      ToolTipText     =   "If a date, record once on that date.  If a weekday name, record multiply on that day of the week and date/time specified."
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   0
      Left            =   6135
      TabIndex        =   24
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   720
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   1
      Left            =   6135
      TabIndex        =   32
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   1050
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   2
      Left            =   6135
      TabIndex        =   40
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   1380
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   3
      Left            =   6135
      TabIndex        =   48
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   1710
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   4
      Left            =   6135
      TabIndex        =   56
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   2040
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   5
      Left            =   6135
      TabIndex        =   64
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   2370
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   6
      Left            =   6135
      TabIndex        =   72
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   2700
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   7
      Left            =   6135
      TabIndex        =   80
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   3030
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   8
      Left            =   6135
      TabIndex        =   88
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   3360
      Width           =   2400
   End
   Begin VB.TextBox txtFN_Prefix 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   9
      Left            =   6135
      TabIndex        =   96
      ToolTipText     =   "Recorded filename prefix for this station.  You can have a station in here multiple times with different prefixes and date/times."
      Top             =   3690
      Width           =   2400
   End
   Begin VB.CommandButton cmdSaveNow 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Save List Now"
      Height          =   360
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Just save the current information to the memory array."
      Top             =   7500
      UseMaskColor    =   -1  'True
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FF00FF&
      Cancel          =   -1  'True
      Caption         =   "&Close without Saving"
      Height          =   345
      Left            =   10320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "The doorbell just rung.  Better go answer it."
      Top             =   7500
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton cmdUse 
      BackColor       =   &H00FFFF80&
      Caption         =   "Save &List, Use Selected and Close"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   360
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Save the stations to the internal list, use the one selected (if no timed starts exist) and close this form."
      Top             =   7500
      UseMaskColor    =   -1  'True
      Width           =   2715
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   9
      Left            =   1080
      TabIndex        =   95
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   3690
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   9
      Left            =   720
      TabIndex        =   9
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   3690
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   8
      Left            =   1080
      TabIndex        =   87
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   3360
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   8
      Left            =   720
      TabIndex        =   8
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   3360
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   7
      Left            =   1080
      TabIndex        =   79
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   3030
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   7
      Left            =   720
      TabIndex        =   7
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   3030
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   6
      Left            =   1080
      TabIndex        =   71
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   2700
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   6
      Left            =   720
      TabIndex        =   6
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   2700
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   63
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   2370
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   5
      Left            =   720
      TabIndex        =   5
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   2370
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   55
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   2040
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   4
      Left            =   720
      TabIndex        =   4
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   2040
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   47
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   1710
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   1710
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   39
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   1380
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   1380
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   31
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   1050
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   1050
      Width           =   195
   End
   Begin VB.TextBox txtStationName 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   23
      ToolTipText     =   "Station name as provided by Shoutcast.  Not changable."
      Top             =   720
      Width           =   5000
   End
   Begin VB.OptionButton OptUseMe 
      Caption         =   "Option1"
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Tag             =   "-1"
      ToolTipText     =   "Use the station that is on this row when ""Save and Use"" is clicked, below."
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "File by Song Title"
      Height          =   555
      Left            =   8520
      TabIndex        =   233
      Top             =   60
      Width           =   495
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   2580
      Top             =   7440
      Width           =   2835
   End
   Begin VB.Label lblScrollLoc 
      Alignment       =   2  'Center
      Caption         =   "1000 / 1000"
      Height          =   195
      Left            =   2625
      TabIndex        =   231
      ToolTipText     =   "Current and maximum stations counters."
      Top             =   7590
      Width           =   2745
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "200"
      Height          =   195
      Index           =   19
      Left            =   135
      TabIndex        =   229
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   7005
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "9999"
      Height          =   195
      Index           =   18
      Left            =   135
      TabIndex        =   228
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   6675
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "18."
      Height          =   195
      Index           =   17
      Left            =   135
      TabIndex        =   227
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   6345
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "17."
      Height          =   195
      Index           =   16
      Left            =   135
      TabIndex        =   226
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   6015
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "16."
      Height          =   195
      Index           =   15
      Left            =   135
      TabIndex        =   225
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   5685
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "15."
      Height          =   195
      Index           =   14
      Left            =   135
      TabIndex        =   224
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   5355
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "14."
      Height          =   195
      Index           =   13
      Left            =   135
      TabIndex        =   223
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   5025
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "13."
      Height          =   195
      Index           =   12
      Left            =   135
      TabIndex        =   222
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   4695
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "12."
      Height          =   195
      Index           =   11
      Left            =   135
      TabIndex        =   221
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   4365
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "11."
      Height          =   195
      Index           =   10
      Left            =   135
      TabIndex        =   220
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   4035
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "10."
      Height          =   195
      Index           =   9
      Left            =   135
      TabIndex        =   219
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   3705
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "9."
      Height          =   195
      Index           =   8
      Left            =   135
      TabIndex        =   218
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   3375
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "8."
      Height          =   195
      Index           =   7
      Left            =   135
      TabIndex        =   217
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   3045
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "7."
      Height          =   195
      Index           =   6
      Left            =   135
      TabIndex        =   216
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   2715
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "6."
      Height          =   195
      Index           =   5
      Left            =   135
      TabIndex        =   215
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   2385
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "5."
      Height          =   195
      Index           =   4
      Left            =   135
      TabIndex        =   214
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   2055
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "4."
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   213
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   1725
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "3."
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   212
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   1395
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "2."
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   211
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   1065
      Width           =   405
   End
   Begin VB.Label lblSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "1."
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   210
      ToolTipText     =   "Number of memorized station up to maximum stations."
      Top             =   750
      Width           =   405
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   19
      Left            =   11310
      TabIndex        =   209
      Top             =   4035
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   18
      Left            =   11310
      TabIndex        =   208
      Top             =   4365
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   11310
      TabIndex        =   207
      Top             =   4695
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   11310
      TabIndex        =   206
      Top             =   5025
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   11310
      TabIndex        =   205
      Top             =   5355
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   11310
      TabIndex        =   204
      Top             =   5685
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   11310
      TabIndex        =   203
      Top             =   6030
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   11310
      TabIndex        =   202
      Top             =   6345
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   11310
      TabIndex        =   201
      Top             =   6675
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   11310
      TabIndex        =   200
      Top             =   7020
      Width           =   45
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Click to Clear"
      Height          =   375
      Index           =   6
      Left            =   8955
      TabIndex        =   199
      Top             =   270
      Width           =   830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Duration"
      Height          =   195
      Index           =   5
      Left            =   11760
      TabIndex        =   198
      Top             =   450
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Time"
      Height          =   195
      Index           =   4
      Left            =   11040
      TabIndex        =   197
      Top             =   450
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Click to set Start Date"
      Height          =   375
      Index           =   3
      Left            =   9840
      TabIndex        =   196
      Top             =   270
      Width           =   1095
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   11310
      TabIndex        =   195
      Top             =   3705
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   11310
      TabIndex        =   194
      Top             =   3375
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   11310
      TabIndex        =   193
      Top             =   3045
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   11310
      TabIndex        =   192
      Top             =   2715
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   11310
      TabIndex        =   191
      Top             =   2385
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   11310
      TabIndex        =   190
      Top             =   2055
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   11310
      TabIndex        =   189
      Top             =   1725
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   11310
      TabIndex        =   188
      Top             =   1395
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   11310
      TabIndex        =   187
      Top             =   1065
      Width           =   45
   End
   Begin VB.Label lblColon 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   11310
      TabIndex        =   186
      Top             =   750
      Width           =   45
   End
   Begin VB.Shape shpOptions 
      Height          =   6675
      Left            =   660
      Shape           =   4  'Rounded Rectangle
      Top             =   660
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Recorded Filename Prefix"
      Height          =   195
      Index           =   0
      Left            =   6195
      TabIndex        =   185
      Top             =   420
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Original Station Name"
      Height          =   195
      Index           =   2
      Left            =   1140
      TabIndex        =   184
      Top             =   420
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Use"
      Height          =   195
      Index           =   1
      Left            =   660
      TabIndex        =   183
      Top             =   420
      Width           =   285
   End
End
Attribute VB_Name = "frmStationMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private bIgnoreClicks As Boolean
  
  Private iSelectedStation As Long
  
  Private iCurrentTopStation As Long

  Private tMemorizedStations(MAX_MEMORY_STATIONS) As Station
  Private iMemorizedStationsCt As Long
  
Function AfterChange() As Boolean

  Dim i As Long
  
  AfterChange = False
  
  ClearScreenForReloading

  FillGrid
  
  For i = 0 To 19
    If txtStationName(i) = "" And txtFN_Prefix(i) = "" Then
      cmdClearStationInfo(i).Enabled = False
    Else
      cmdClearStationInfo(i).Enabled = True
    End If

    If txtFN_Prefix(i) <> "" And txtStationName(i) = "" Then
      MsgBox "There is something in the Recorded Filename Prefix on line " & i + iCurrentTopStation & _
             " but no Station Name.  Please correct."
      Exit Function
    End If
    
    If cmdRecDate(i).Caption = NOT_SET Then
      cmdClearTimedInfo(i).Enabled = False
      txtStartHour(i).Enabled = False
      txtStartMin(i).Enabled = False
      txtRecordMinutes(i).Enabled = False
    Else
      If txtStartHour(i) = "" Or txtStartMin(i) = "" Or Val(txtRecordMinutes(i)) = 0 Then
        MsgBox "There is a start date for stream " & i + iCurrentTopStation + 1 & " but incomplete start time & duration information."
        Exit Function
      End If
      cmdClearTimedInfo(i).Enabled = True
      txtStartHour(i).Enabled = True
      txtStartMin(i).Enabled = True
      txtRecordMinutes(i).Enabled = True
    End If

  Next
  
  AfterChange = True
  
End Function
Sub ClearOpts()

  Dim i As Long
  
  For i = 0 To 19
    OptUseMe(i).Value = False
  Next
  
End Sub

Sub ClearScreenForReloading()

  Dim i As Long
  
  bIgnoreClicks = True
  
  For i = 0 To 19
    txtStationName(i) = ""
    txtFN_Prefix(i) = ""
    cmdClearTimedInfo(i).Enabled = False
    cmdRecDate(i).Enabled = False
    cmdRecDate(i).Caption = NOT_SET
    txtStartHour(i) = ""
    txtStartMin(i) = ""
    txtRecordMinutes(i) = ""
    chkFileBySong(i).Value = vbUnchecked
  Next
  
  bIgnoreClicks = False
  
End Sub

Sub FillGrid()

  Dim i As Long
  
  bIgnoreClicks = True
  
  If iMemorizedStationsCt > 0 Then
    vsStreams.Value = iCurrentTopStation
    For i = 0 To iMemorizedStationsCt - 1
      If i = 20 Then Exit For  ' Double ending here.  If more than 20 stations, get out.  No more room now.
      txtStationName(i) = tMemorizedStations(i + iCurrentTopStation + 1).StationName
      txtStationName(i).ToolTipText = txtStationName(i) & " IP: " & tMemorizedStations(i + iCurrentTopStation + 1).URL
      
      cmdRecDate(i).Enabled = True
      
      txtFN_Prefix(i) = tMemorizedStations(i + iCurrentTopStation + 1).MyFilePrefix
      aSplit = Split(tMemorizedStations(i + iCurrentTopStation + 1).Genre, ":")
      If UBound(aSplit) > 0 Then  ' See if the Genre: is still there.  Will be coming back from Shoutcast selector.
        txtFN_Prefix(i).ToolTipText = aSplit(1)
      Else
        txtFN_Prefix(i).ToolTipText = tMemorizedStations(i + iCurrentTopStation + 1).Genre
      End If
      
      If tMemorizedStations(i + iCurrentTopStation + 1).StartDate = "" Then
        cmdRecDate(i).Caption = NOT_SET
      Else
        cmdRecDate(i).Caption = tMemorizedStations(i + iCurrentTopStation + 1).StartDate
      End If
      txtStartHour(i) = Val(tMemorizedStations(i + iCurrentTopStation + 1).StartHour)
      txtStartMin(i) = Val(tMemorizedStations(i + iCurrentTopStation + 1).StartMin)
      txtRecordMinutes(i) = Val(tMemorizedStations(i + iCurrentTopStation + 1).Duration)
      chkFileBySong(i).Value = tMemorizedStations(i + iCurrentTopStation + 1).UseICYSongTitle
    Next
  End If
  
  i = iCurrentTopStation + 20
  If i > giMemorizedStationsCt Then i = giMemorizedStationsCt
    
  lblScrollLoc = "Showing " & iCurrentTopStation + 1 & " to " & i & " of " & giMemorizedStationsCt
  
  UpdateSequenceNumbers
  
  bIgnoreClicks = False
  
End Sub

Sub FindNextRecording()

 'Go through list of stations and find next one to record.
 
End Sub

Sub SaveList()

  Dim i As Integer
  
  If AfterChange Then
    giMemorizedStationsCt = iMemorizedStationsCt
    For i = 1 To iMemorizedStationsCt
      gtMemorizedStations(i).StationName = tMemorizedStations(i).StationName
      gtMemorizedStations(i).MyFilePrefix = tMemorizedStations(i).MyFilePrefix
      gtMemorizedStations(i).Format = tMemorizedStations(i).Format
      gtMemorizedStations(i).ID = tMemorizedStations(i).ID
      gtMemorizedStations(i).BitRate = tMemorizedStations(i).BitRate
      gtMemorizedStations(i).Genre = tMemorizedStations(i).Genre
      gtMemorizedStations(i).StartDate = tMemorizedStations(i).StartDate
      gtMemorizedStations(i).StartHour = tMemorizedStations(i).StartHour
      gtMemorizedStations(i).StartMin = tMemorizedStations(i).StartMin
      gtMemorizedStations(i).Duration = tMemorizedStations(i).Duration
      gtMemorizedStations(i).URL = tMemorizedStations(i).URL
      gtMemorizedStations(i).UseICYSongTitle = tMemorizedStations(i).UseICYSongTitle
    Next
  
  End If
  
End Sub

Sub UpdateSequenceNumbers()

  Dim i As Integer
  
  For i = 0 To 19
    lblSeq(i) = i + iCurrentTopStation + 1
  Next
  
End Sub

Private Sub chkFileBySongTitle_Click(Index As Integer)

  If bIgnoreClicks Then Exit Sub
  
  If chkFileBySong(Index).Value = vbChecked Then
    txtFN_Prefix(Index).Enabled = False
  Else
     txtFN_Prefix(Index).Enabled = True
  End If
  
  tMemorizedStations(Index + iCurrentTopStation + 1).UseICYSongTitle = chkFileBySong(Index).Value

End Sub

Private Sub chkFileBySong_Click(Index As Integer)
  
  If bIgnoreClicks Then Exit Sub
  
  tMemorizedStations(Index + iCurrentTopStation + 1).UseICYSongTitle = chkFileBySong(Index).Value
  
End Sub

Private Sub cmdClearStationInfo_Click(Index As Integer)

    tMemorizedStations(Index + iCurrentTopStation + 1).StationName = ""
    tMemorizedStations(Index + iCurrentTopStation + 1).MyFilePrefix = ""
    tMemorizedStations(Index + iCurrentTopStation + 1).Format = ""
    tMemorizedStations(Index + iCurrentTopStation + 1).ID = ""
    tMemorizedStations(Index + iCurrentTopStation + 1).BitRate = ""
    tMemorizedStations(Index + iCurrentTopStation + 1).Genre = ""
    tMemorizedStations(Index + iCurrentTopStation + 1).StartDate = ""
    tMemorizedStations(Index + iCurrentTopStation + 1).StartHour = ""
    tMemorizedStations(Index + iCurrentTopStation + 1).StartMin = ""
    tMemorizedStations(Index + iCurrentTopStation + 1).Duration = ""
    tMemorizedStations(Index + iCurrentTopStation + 1).UseICYSongTitle = vbUnchecked
    
  AfterChange
  
End Sub

Private Sub cmdClearTimedInfo_Click(Index As Integer)

  tMemorizedStations(Index + iCurrentTopStation + 1).StartDate = ""
  tMemorizedStations(Index + iCurrentTopStation + 1).StartHour = ""
  tMemorizedStations(Index + iCurrentTopStation + 1).StartMin = ""
  tMemorizedStations(Index + iCurrentTopStation + 1).Duration = ""
  
 'txtStartHour(Index + 1).Enabled = False
 'txtStartMin(Index + 1).Enabled = False
 'txtRecordMinutes(Index + 1).Enabled = False
 'cmdClearTimedInfo(Index + 1).Enabled = False
  
  AfterChange
  
End Sub


Private Sub cmdDown_Click()
  
  iCurrentTopStation = iCurrentTopStation + 20
  If iCurrentTopStation > (MAX_MEMORY_STATIONS - 20) Then iCurrentTopStation = MAX_MEMORY_STATIONS - 20
  If iCurrentTopStation > iMemorizedStationsCt Then iCurrentTopStation = iCurrentTopStation - 20
  If iCurrentTopStation < 0 Then iCurrentTopStation = 0
  
  AfterChange
  
End Sub

Private Sub cmdExit_Click()
  iFinallyTheStationToTune = 0
  Unload Me
End Sub

Private Sub cmdRecDate_Click(Index As Integer)

  
  If cmdRecDate(Index).Caption = NOT_SET Then
    giUseReturnedOrNot = ReturnedDateType.UseNone
  ElseIf IsDate(cmdRecDate(Index).Caption) Then
    giUseReturnedOrNot = ReturnedDateType.UseDate
    gdtStartDate = cmdRecDate(Index).Caption
  Else
    giUseReturnedOrNot = ReturnedDateType.UseDay
    gsPassedDayName = cmdRecDate(Index).Caption
  End If
  
  frmSelectStartDate.Show vbModal
    
  Select Case giUseReturnedOrNot
    Case ReturnedDateType.UseDate
      cmdRecDate(Index).Caption = gdtStartDate
      txtStartHour(Index).Enabled = True
      txtStartMin(Index).Enabled = True
      txtRecordMinutes(Index).Enabled = True
      cmdClearTimedInfo(Index).Enabled = True
      txtStartHour(Index).SetFocus
    Case ReturnedDateType.UseDay
      cmdRecDate(Index).Caption = gsDayNames(giReturnedDay) & "s"
      txtStartHour(Index).Enabled = True
      txtStartMin(Index).Enabled = True
      cmdClearTimedInfo(Index).Enabled = True
      txtRecordMinutes(Index).Enabled = True
      txtStartHour(Index).SetFocus
    Case ReturnedDateType.UseNone
      cmdRecDate(Index).Caption = NOT_SET
      txtStartHour(Index) = ""
      txtStartMin(Index) = ""
      txtRecordMinutes(Index) = ""
    Case ReturnedDateType.NoChange
     'NOP
  End Select
  
End Sub

Private Sub cmdRecDate_LostFocus(Index As Integer)
  
  tMemorizedStations(Index + iCurrentTopStation + 1).StartDate = cmdRecDate(Index).Caption

End Sub


Private Sub cmdSaveNow_Click()
  SaveList
End Sub

Private Sub cmdShowShoutcast_Click()
  
  Dim i As Long
  Dim iStation2Add As Long
  
  ReDim Preserve NewShoutcastStation(1)  ' 0 is a special case.  It hold the manual fields on frmMain.  Start new stations at 1 each time.
  If gbDebugLogic Then Debug.Print MyTime() & "ShowShoutcast: says upper bound of NewShoutcastStation Array is: "; UBound(NewShoutcastStation)
  
  Me.MousePointer = vbHourglass
  frmDataFetch.Show vbModal
 '
 'To make it possible to add multiple stations...
 'Make NewShoutcastStation a redimmable array.
 'When returning from frmDataFetch, loop on the new station array UBound and add all that are passed back.
 'If none were passed back, the UBound will be 1 but the StationName in that one will be blank.
 'If 1 was passed back, then UBound will still be 1 but StationName(1) will not be blank so use it.
 'If more than 1 was passed back, then loop to UBound of the returned array.  It will be ReDim Preserve'd over there.
 'If too many are passed back to fit in the memorized stations list, excess ones will be skipped over and a msgbox will appear.
 '
  For iStation2Add = 1 To UBound(NewShoutcastStation)
    If NewShoutcastStation(iStation2Add).StationName <> "" Then  ' If anything was selected... (there is no cancel variable used or needed this way)
      For i = 1 To MAX_MEMORY_STATIONS  ' Find a blank spot to put it based on the unchangable SC Station Name (as munched).
        If tMemorizedStations(i).StationName = "" Then
          tMemorizedStations(i).StationName = NewShoutcastStation(iStation2Add).StationName
          tMemorizedStations(i).MyFilePrefix = Trim$(Left(NewShoutcastStation(iStation2Add).StationName & Space(25), 25))
          tMemorizedStations(i).Genre = NewShoutcastStation(iStation2Add).Genre
          tMemorizedStations(i).Format = NewShoutcastStation(iStation2Add).Format
          tMemorizedStations(i).ID = NewShoutcastStation(iStation2Add).ID
          tMemorizedStations(i).BitRate = NewShoutcastStation(iStation2Add).BitRate
          tMemorizedStations(i).URL = NewShoutcastStation(iStation2Add).URL
          Exit For
        End If
       Next
    End If
  Next
 
 'At this point, i has number of the station just added.  I want to scroll the display down to see that new add.
  iCurrentTopStation = (i \ 20) * 20
  AfterChange
  
  If i = MAX_MEMORY_STATIONS + 1 Then
    MsgBox "Too many saved stations.  Cannot add any more."
    iMemorizedStationsCt = MAX_MEMORY_STATIONS
    Exit Sub
  End If
  
  For iMemorizedStationsCt = MAX_MEMORY_STATIONS To 1 Step -1
    If tMemorizedStations(iMemorizedStationsCt).StationName <> "" Then Exit For
  Next
  
  AfterChange
  
  Me.MousePointer = vbDefault
  
End Sub

Private Sub cmdUp_Click()
  
  iCurrentTopStation = iCurrentTopStation - 20
  If iCurrentTopStation < 0 Then iCurrentTopStation = 0
  
  AfterChange
  
End Sub

Private Sub cmdUse_Click()

    
  If AfterChange Then  ' If all validity checks pass, then fill in Main form and unload.
    
    SaveList
    
    If iSelectedStation = 0 Then
      MsgBox "No station selected.  Please select a station and try again."
      Exit Sub
    End If
    
    iFinallyTheStationToTune = iSelectedStation '+ iCurrentTopStation + 1
    
    Unload Me
    
  End If
  
End Sub

Private Sub Form_Load()

  Dim i As Long
 'Dim iMax As Long
  
 'Me.Show
  
  bIgnoreClicks = True
  
  iSelectedStation = 0
  
  Me.Top = GetSetting(App.EXEName, "Form", "frmMemory_Top", frmMain.Top)
  If Me.Top < 0 Then Me.Top = 0
  If Me.Top + Me.Height > Screen.Height Then Me.Top = Screen.Height - Me.Height
  
  Me.Left = GetSetting(App.EXEName, "Form", "frmMemory_Left", frmMain.Left)
  If Me.Left < 0 Then Me.Top = 0
  If Me.Left + Me.Width > Screen.Width Then Me.Left = Screen.Width - Me.Width
  vsStreams.Max = giMemorizedStationsCt - 1
  iMemorizedStationsCt = giMemorizedStationsCt
  For i = 1 To giMemorizedStationsCt
    tMemorizedStations(i).StationName = gtMemorizedStations(i).StationName
    tMemorizedStations(i).MyFilePrefix = gtMemorizedStations(i).MyFilePrefix
    tMemorizedStations(i).Format = gtMemorizedStations(i).Format
    tMemorizedStations(i).ID = gtMemorizedStations(i).ID
    tMemorizedStations(i).BitRate = gtMemorizedStations(i).BitRate
    tMemorizedStations(i).Genre = gtMemorizedStations(i).Genre
    tMemorizedStations(i).StartDate = gtMemorizedStations(i).StartDate
    tMemorizedStations(i).StartHour = gtMemorizedStations(i).StartHour
    tMemorizedStations(i).StartMin = gtMemorizedStations(i).StartMin
    tMemorizedStations(i).Duration = gtMemorizedStations(i).Duration
    tMemorizedStations(i).URL = gtMemorizedStations(i).URL
    tMemorizedStations(i).UseICYSongTitle = gtMemorizedStations(i).UseICYSongTitle
  Next
  
 'ClearScreenForReloading  ' Now done in AfterChange
  
  iCurrentTopStation = 0
  
  DO_CtrlOutline Me

  AfterChange
  
  bIgnoreClicks = False
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  SaveSetting App.EXEName, "Form", "frmMemory_Top", Me.Top
  SaveSetting App.EXEName, "Form", "frmMemory_Left", Me.Left

End Sub


Private Sub OptUseMe_Click(Index As Integer)
  
  If bIgnoreClicks Then Exit Sub
  
  iSelectedStation = Index + iCurrentTopStation + 1
  
  If tMemorizedStations(iSelectedStation).StationName = "" Then
    ClearOpts
    iSelectedStation = 1
    MsgBox "You should not select an empty station.  That's just silly!"
  End If
  
  If gbDebugLogic Then Debug.Print MyTime() & "Clicked opt on station: " & iSelectedStation
  
End Sub


Private Sub OptUseMe_GotFocus(Index As Integer)
  shpOptions.Visible = True
End Sub


Private Sub OptUseMe_LostFocus(Index As Integer)
  shpOptions.Visible = False
End Sub


Private Sub txtRecordMinutes_LostFocus(Index As Integer)

    tMemorizedStations(Index + iCurrentTopStation + 1).Duration = txtRecordMinutes(Index)

End Sub

Private Sub txtFN_Prefix_GotFocus(Index As Integer)

  If Len(txtFN_Prefix(Index)) > 0 Then
    txtFN_Prefix(Index).SelStart = 0
    txtFN_Prefix(Index).SelLength = Len(txtFN_Prefix(Index))
  End If

End Sub


Private Sub txtFN_Prefix_LostFocus(Index As Integer)

  Dim i As Integer
  
  If txtFN_Prefix(Index) = "" Then
    i = 20
    If Len(txtStationName(Index)) < 20 Then i = Len(txtStationName(Index))
    txtFN_Prefix(Index) = Left(txtStationName(Index), i)
  End If
  
  tMemorizedStations(Index + iCurrentTopStation + 1).MyFilePrefix = txtFN_Prefix(Index)
    
  AfterChange

End Sub


Private Sub txtRecordMinutes_Change(Index As Integer)

  If Len(txtRecordMinutes(Index)) = 0 Then Exit Sub  ' Nothing there.  No need to check validity.
  
  If Not IsNumeric(txtRecordMinutes(Index)) Then
    txtRecordMinutes(Index) = Left(txtRecordMinutes(Index), Len(txtRecordMinutes(Index)) - 1) ' Chop off offending appendage.
    txtRecordMinutes(Index).SelStart = Len(txtRecordMinutes(Index))  ' Leave input cursor at the end of the good stuff.
    Beep
  Else  ' It is numeric
    txtRecordMinutes(Index) = Val(txtRecordMinutes(Index))
    tMemorizedStations(Index + iCurrentTopStation + 1).Duration = Val(txtRecordMinutes(Index))
  End If
  
End Sub

Private Sub txtRecordMinutes_GotFocus(Index As Integer)

  If Len(txtRecordMinutes(Index)) > 0 Then
    txtRecordMinutes(Index).SelStart = 0
    txtRecordMinutes(Index).SelLength = Len(txtRecordMinutes(Index))
  End If

End Sub


Private Sub txtStartHour_Change(Index As Integer)

  If bIgnoreClicks Or Len(txtStartHour(Index)) = 0 Then Exit Sub   ' Nothing there.  No need to check validity.
  
  If Not IsNumeric(txtStartHour(Index)) Then
    txtStartHour(Index) = Left(txtStartHour(Index), Len(txtStartHour(Index)) - 1) ' Chop off offending appendage.
    txtStartHour(Index).SelStart = Len(txtStartHour(Index))  ' Leave input cursor at the end of the good stuff.
    Beep
    Exit Sub
  Else  ' It is numeric
   'txtStartHour(Index) = Val(txtStartHour(Index))
    tMemorizedStations(Index + iCurrentTopStation + 1).StartHour = Val(txtStartHour(Index))
  End If
  
  If Len(txtStartHour(Index)) = 2 Then
    txtStartHour(Index) = Val(txtStartHour(Index))
    txtStartMin(Index).SetFocus
  End If

End Sub


Private Sub txtStartHour_GotFocus(Index As Integer)

  If Len(txtStartHour(Index)) > 0 Then
    txtStartHour(Index).SelStart = 0
    txtStartHour(Index).SelLength = Len(txtStartHour(Index))
  End If
  
End Sub


Private Sub txtStartHour_LostFocus(Index As Integer)

  If Len(txtStartHour(Index)) > 0 Then
    If txtStartHour(Index) > 23 Then
      MsgBox "Please select an hour between 0 and 23."
      txtStartHour(Index).SetFocus
    Else
      tMemorizedStations(Index + iCurrentTopStation + 1).StartHour = txtStartHour(Index)
    End If
  End If
  
End Sub

Private Sub txtStartMin_Change(Index As Integer)

  If bIgnoreClicks Or Len(txtStartMin(Index)) = 0 Then Exit Sub   ' Nothing there.  No need to check validity.
  
  If Not IsNumeric(txtStartMin(Index)) Then
    txtStartMin(Index) = Left(txtStartMin(Index), Len(txtStartMin(Index)) - 1)  ' Chop off offending appendage.
    txtStartMin(Index).SelStart = Len(txtStartMin(Index))  ' Leave input cursor at the end of the good stuff.
    Beep
    Exit Sub
  Else  ' It is numeric
    tMemorizedStations(Index + iCurrentTopStation + 1).StartMin = Val(txtStartMin(Index))
  End If
  
  If Len(txtStartMin(Index)) = 2 Then
    txtStartMin(Index) = Val(txtStartMin(Index))
    txtRecordMinutes(Index).SetFocus
  End If
  
End Sub


Private Sub txtStartMin_GotFocus(Index As Integer)

  If Len(txtStartMin(Index)) > 0 Then
    txtStartMin(Index).SelStart = 0
    txtStartMin(Index).SelLength = Len(txtStartMin(Index))
  End If
  
End Sub


Private Sub txtStartMin_LostFocus(Index As Integer)

  If Len(txtStartMin(Index)) > 0 Then
    If txtStartMin(Index) > 59 Then
      MsgBox "Please select an minute between 0 and 59."
      txtStartMin(Index).SetFocus
    Else
      tMemorizedStations(Index + iCurrentTopStation + 1).StartMin = txtStartMin(Index)
    End If
  End If
  
End Sub


Private Sub txtStationName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  
  If gbDebugLogic Then Debug.Print MyTime() & "Keydown: " & KeyCode
  
  Select Case KeyCode
    Case 35, 36, 37, 39
    Case Else
      KeyCode = 0
  End Select
  
End Sub

Private Sub txtStationName_KeyPress(Index As Integer, KeyAscii As Integer)
  If gbDebugLogic Then Debug.Print MyTime(); "txtStationName_KeyPressed.  Ignoring: " & KeyAscii
  KeyAscii = 0
End Sub


Private Sub vsStreams_Change()
  
  Dim iDirection As Long
  
  iDirection = iCurrentTopStation - vsStreams.Value
  iCurrentTopStation = vsStreams.Value
  
  If iSelectedStation - iCurrentTopStation - 1 < 0 Or iSelectedStation - iCurrentTopStation - 1 > 19 Then
    ClearOpts
  Else
    bIgnoreClicks = True
    OptUseMe(iSelectedStation - iCurrentTopStation - 1).Value = True
    bIgnoreClicks = False
  End If
  
  If iCurrentTopStation > (MAX_MEMORY_STATIONS - 20) Then iCurrentTopStation = MAX_MEMORY_STATIONS - 20
  If iCurrentTopStation > iMemorizedStationsCt Then iCurrentTopStation = iCurrentTopStation - 20
  If iCurrentTopStation < 0 Then iCurrentTopStation = 0
  
  AfterChange

End Sub


