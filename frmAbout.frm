VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3345
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5310
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2308.779
   ScaleMode       =   0  'User
   ScaleWidth      =   4986.365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2258
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2745
      Width           =   795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   528.686
      X2              =   4444.532
      Y1              =   1708.289
      Y2              =   1708.289
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   2
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   542.772
      X2              =   4444.532
      Y1              =   1718.642
      Y2              =   1708.289
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  
    Unload Me
    Set frmAbout = Nothing
    
End Sub

Private Sub Form_Load()
    
    FrmAboutOpen = True
    Me.Caption = "About " & App.EXEName
    Me.Icon = frmMain.Icon
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.EXEName
    lblDescription.Caption = "Put all your shortcuts in one folder.  Give eXeCute the path to that folder and have access to every shortcut from a system tray icon."
    picIcon.Picture = frmMain.Icon
        
End Sub

Private Sub Form_Paint()
'# set the form colours to my personal choice
    
    ColourIn Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FrmAboutOpen = False
    
End Sub
