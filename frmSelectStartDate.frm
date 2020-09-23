VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmSelectStartDate 
   Caption         =   "Select Day/Date for automatic recording"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   5985
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF00FF&
      Cancel          =   -1  'True
      Caption         =   "&No Change"
      Height          =   400
      Left            =   4020
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Your dinner is ready."
      Top             =   3570
      UseMaskColor    =   -1  'True
      Width           =   1755
   End
   Begin VB.Frame Frame1 
      Caption         =   "Record every..."
      Height          =   2655
      Left            =   3900
      TabIndex        =   6
      ToolTipText     =   "Click on a day name to record repeatedly on that day of the week."
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton OptRepeatDay 
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   13
         Tag             =   "-1"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.OptionButton OptRepeatDay 
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   12
         Tag             =   "-1"
         Top             =   1960
         Width           =   1575
      End
      Begin VB.OptionButton OptRepeatDay 
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   11
         Tag             =   "-1"
         Top             =   1640
         Width           =   1575
      End
      Begin VB.OptionButton OptRepeatDay 
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   10
         Tag             =   "-1"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton OptRepeatDay 
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Tag             =   "-1"
         Top             =   1000
         Width           =   1575
      End
      Begin VB.OptionButton OptRepeatDay 
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Tag             =   "-1"
         Top             =   680
         Width           =   1575
      End
      Begin VB.OptionButton OptRepeatDay 
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Tag             =   "-1"
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.OptionButton OptRepeating 
      Caption         =   "Repeat every week"
      Height          =   315
      Left            =   3900
      TabIndex        =   5
      Tag             =   "-1"
      ToolTipText     =   "Repeat every week on the day selected."
      Top             =   2900
      Width           =   1935
   End
   Begin VB.OptionButton optOneTime 
      Caption         =   "Record once on this date"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Tag             =   "-1"
      ToolTipText     =   "Use one-time recording on selected date."
      Top             =   2900
      Width           =   2355
   End
   Begin VB.CommandButton cmdAccept 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Set Start Day/Date"
      Default         =   -1  'True
      Height          =   400
      Left            =   180
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Remember entry and close form."
      Top             =   3570
      UseMaskColor    =   -1  'True
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Clear Date/Time"
      Height          =   400
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Your dinner is ready."
      Top             =   3570
      UseMaskColor    =   -1  'True
      Width           =   1755
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2835
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Change date and click on a day to record, once, on that date."
      Top             =   60
      Width           =   3675
      _Version        =   524288
      _ExtentX        =   6482
      _ExtentY        =   5001
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2009
      Month           =   10
      Day             =   15
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   0
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblSelectedDate 
      Alignment       =   2  'Center
      Caption         =   "Selected Date: 12/31/2009"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   3240
      Width           =   5880
   End
End
Attribute VB_Name = "frmSelectStartDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private gRepeatDay As Long

Private Sub Calendar1_Click()
  lblSelectedDate = "Record once on: " & Calendar1.Value
  optOneTime.Value = True
End Sub

Private Sub Calendar1_DblClick()

  gdtStartDate = Calendar1.Value
  giUseReturnedOrNot = ReturnedDateType.UseDate
  
  Unload Me
  
End Sub

Private Sub cmdAccept_Click()

  Dim i As Integer
  
  If optOneTime = False And OptRepeating = False Then
    MsgBox "No selection made.  Please pick 'Record once on this date' or 'Repeat every week'."
    Exit Sub
  End If
  
  If optOneTime Then
    gdtStartDate = Calendar1.Value
    giUseReturnedOrNot = ReturnedDateType.UseDate
  Else
    giReturnedDay = gRepeatDay
    giUseReturnedOrNot = ReturnedDateType.UseDay
  End If
  
  If gdtStartDate < Date Then
    i = MsgBox("You asked for a start date before today.  Is that what you meant to do?", vbYesNo, "Early Start Verification")
    If i = vbNo Then Exit Sub
  End If
    
  Unload Me
  
End Sub

Private Sub cmdCancel_Click()

  giUseReturnedOrNot = ReturnedDateType.UseNone
  Unload Me

End Sub

Private Sub Command1_Click()

  giUseReturnedOrNot = ReturnedDateType.NoChange
  Unload Me
  
End Sub

Private Sub Form_Load()

  Dim i As Integer
  
  Me.Top = GetSetting(App.EXEName, "Form", "frmDate_Top", frmMain.Top)
  If Me.Top < 0 Then Me.Top = 0
  If Me.Top + Me.Height > Screen.Height Then Me.Top = Screen.Height - Me.Height
  
  Me.Left = GetSetting(App.EXEName, "Form", "frmDate_Left", frmMain.Left)
  If Me.Left < 0 Then Me.Top = 0
  If Me.Left + Me.Width > Screen.Width Then Me.Left = Screen.Width - Me.Width
  
  OptRepeatDay(0).Caption = "&" & gsDayNames(0)
  OptRepeatDay(1).Caption = "&" & gsDayNames(1)
  OptRepeatDay(2).Caption = "&" & gsDayNames(2)
  OptRepeatDay(3).Caption = "&" & gsDayNames(3)
  OptRepeatDay(4).Caption = "&" & gsDayNames(4)
  OptRepeatDay(5).Caption = "&" & gsDayNames(5)
  OptRepeatDay(6).Caption = "&" & gsDayNames(6)
  
  Select Case giUseReturnedOrNot
    Case ReturnedDateType.UseNone
      gdtStartDate = Now
      Calendar1.Value = Now
      lblSelectedDate = "Please select a day or date for recording."
      
    Case ReturnedDateType.UseDate
      Calendar1.Value = gdtStartDate
      optOneTime.Value = True
      lblSelectedDate = "Record once on: " & Calendar1.Value

    Case ReturnedDateType.UseDay
      For i = 0 To 6
        If gsPassedDayName = gsDayNames(i) & "s" Then
          OptRepeatDay(i).Value = True
          OptRepeating.Value = True
          Exit For
        End If
      Next
    Case Else
     'NOP
  End Select
    
  DO_CtrlOutline Me
  
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  SaveSetting App.EXEName, "Form", "frmDate_Top", Me.Top
  SaveSetting App.EXEName, "Form", "frmDate_Left", Me.Left

End Sub


Private Sub OptRepeatDay_Click(Index As Integer)
  
  gRepeatDay = Index
  lblSelectedDate = "Record every " & OptRepeatDay(Index).Caption
  OptRepeating.Value = True

End Sub

