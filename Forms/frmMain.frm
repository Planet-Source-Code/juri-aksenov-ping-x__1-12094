VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ping X"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imglMain 
      Left            =   1920
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0624
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   405
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   405
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   405
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame frmeMain 
      BackColor       =   &H80000004&
      ForeColor       =   &H80000007&
      Height          =   4935
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7815
      Begin RichTextLib.RichTextBox txtNumber 
         Height          =   285
         Left            =   5160
         TabIndex        =   1
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         MaxLength       =   10
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":0D9B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtIP 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   503
         _Version        =   393217
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":0E7D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtOutPut 
         Height          =   3735
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1080
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6588
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":0F5F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdPing 
         Caption         =   "&Ping"
         Default         =   -1  'True
         Height          =   420
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblMain 
         Caption         =   "Number of times:"
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblMain 
         Caption         =   "IP Address:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblMain 
         Caption         =   "Ping Results:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************
'* Developer: Juri Aksenov                                 *
'* Date Started: Sunday, October 15, 2000                  *
'* Date Last Modified: Monday, October 16, 2000            *
'* Description: Ping X is a network utility to check the   *
'*              the internet/intranet connections of other *
'*              computers in your office or home. This is  *
'*              for those people that don't like the       *
'*              command prompt and prefer the GUI.         *
'***********************************************************
Option Explicit

Const SYNCHRONIZE = &H100000
Const INFINITE = &HFFFF
Const WAIT_OBJECT_0 = 0
Const WAIT_TIMEOUT = &H102

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Sub cmdAbout_Click()
    frmAbout.Show
End Sub

Private Sub cmdClear_Click()
    txtIP.Text = ""
    txtNumber.Text = ""
    Open "C:\log.txt" For Output As #1
    Close #1
    txtOutPut.Text = ""
End Sub

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub cmdPing_Click()
Dim ShellX As String
Dim lPid As Long
Dim lHnd As Long
Dim lRet As Long
Dim VarX As String

  frmMain.MousePointer = 11
  If txtIP.Text <> "" Then
    DoEvents
    ShellX = Shell("command.com /c ping -n " & txtNumber.Text & " " & txtIP.Text & " > C:\log.txt", vbHide)
    
    lPid = ShellX
    If lPid <> 0 Then
        lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
        If lHnd <> 0 Then
            lRet = WaitForSingleObject(lHnd, INFINITE)
            CloseHandle (lHnd)
        End If
            Beep
            frmMain.MousePointer = 0
            Open "C:\log.txt" For Input As #1
            txtOutPut.Text = Input(LOF(1), 1)
            Close #1
    End If
  Else
    frmMain.MousePointer = 0
    VarX = MsgBox("You have not entered an ip address or the number of times you want to ping.", vbCritical, "Error has occured")
  End If
End Sub

Private Sub Form_Load()
  frmMain.Icon = imglMain.ListImages.Item(1).Picture
  Open "C:\log.txt" For Output As #1
  Close #1

    SendMessage cmdPing.hWnd, &HF4&, &H0&, 0&
    SendMessage cmdAbout.hWnd, &HF4&, &H0&, 0&
    SendMessage cmdExit.hWnd, &HF4&, &H0&, 0&
    SendMessage cmdClear.hWnd, &HF4&, &H0&, 0&

End Sub

Private Sub SelectText(ByRef textObj As RichTextBox)
    textObj.SelStart = 0
    textObj.SelLength = Len(textObj)
End Sub

Private Sub txtIP_GotFocus()
    Call SelectText(txtIP)
End Sub

Private Sub txtNumber_GotFocus()
    Call SelectText(txtNumber)
End Sub

Private Sub txtOutput_GotFocus()
    Call SelectText(txtOutPut)
End Sub

Private Sub txtStatus_Click()
    txtIP.SetFocus
End Sub
