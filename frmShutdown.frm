VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShutdown 
   BackColor       =   &H008080FF&
   Caption         =   "Reboot Notice"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3915
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAbort 
      Caption         =   "&Abort Shutdown"
      Height          =   495
      Left            =   2633
      TabIndex        =   4
      Top             =   3180
      Width           =   1215
   End
   Begin VB.CommandButton cmdReboot 
      Caption         =   "&Shutdown Now"
      Height          =   495
      Left            =   833
      TabIndex        =   3
      Top             =   3180
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4260
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   420
      TabIndex        =   2
      Top             =   2640
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Enabled         =   0   'False
      Scrolling       =   1
   End
   Begin VB.Label lblWhatsLeft 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your computer will shutdown in 60 seconds.  Press Shutdown to reboot now, Abort to abort shutdown."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1395
      Left            =   420
      TabIndex        =   1
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label lblFlash 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Warning!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   413
      TabIndex        =   0
      Top             =   300
      Width           =   3855
   End
End
Attribute VB_Name = "frmShutdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Dim seconds As Long
  
Private Sub cmdAbort_Click()
  Unload Me
End Sub

Private Sub cmdReboot_Click()
  
  bExitNow = True
  Unload Me
  
  Shell "Shutdown -s"

End Sub

Private Sub Form_Load()

  ProgressBar1.Max = 60
  ProgressBar1.Value = 60
  seconds = 60
  
End Sub

Private Sub Timer1_Timer()

  If lblFlash.Visible Then
    lblFlash.Visible = False
  Else
    lblFlash.Visible = True
  End If
  
  seconds = seconds - 1
  If seconds <= 0 Then
    
    Shell "Shutdown -s -t 60"
    
    bExitNow = True
    Unload Me
  
  End If
  
  ProgressBar1.Value = seconds
  lblWhatsLeft = "Your computer will shutdown in " & seconds & " seconds.  Press Shutdown to reboot now, Abort to abort rebooting."
  
End Sub
