VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Ping X"
   ClientHeight    =   3600
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2484.784
   ScaleMode       =   0  'User
   ScaleWidth      =   5113.137
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmeMain 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      Begin VB.PictureBox picMain 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   1335
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblProductName 
         Caption         =   "Copyright © 2000 Minds Imagination. All right reserved."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   4215
      End
      Begin VB.Label lblProductName 
         Caption         =   "Ping X is a pinging utility for network users who seek simplicity and have a small understanding of the command prompt."
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   2070
         Width           =   4935
      End
      Begin VB.Label lblProductName 
         Caption         =   "Version 1.0"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   4
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblProductName 
         Caption         =   "Ping X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   3983
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1294
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  Dim VarX As Variant
    picMain.Picture = frmMain.imglMain.ListImages.Item(2).Picture
    SendMessage cmdClose.hWnd, &HF4&, &H0&, 0&
    'VarX = App.Path & "\" & "xtechnology.gif"
    'picMain.Picture = LoadPicture(VarX)
End Sub
