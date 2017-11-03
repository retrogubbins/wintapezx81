VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3810
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2627.221
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   1752
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmAbout.frx":0000
      Top             =   1080
      Width           =   5532
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      ClipControls    =   0   'False
      Height          =   540
      Left            =   120
      Picture         =   "frmAbout.frx":1CFE
      ScaleHeight     =   336.791
      ScaleMode       =   0  'User
      ScaleWidth      =   336.791
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   528
      Left            =   4380
      TabIndex        =   0
      Top             =   3180
      Width           =   1260
   End
   Begin VB.Label lblVersion 
      Caption         =   "Ver."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   600
      TabIndex        =   6
      Top             =   720
      Width           =   1488
   End
   Begin VB.Label Label1 
      Caption         =   "based on original ZXD Decoder by Steven McDonald"
      ForeColor       =   &H00000000&
      Height          =   348
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   3876
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   56.343
      X2              =   5281.227
      Y1              =   2023.856
      Y2              =   2023.856
   End
   Begin VB.Label lblTitle 
      Caption         =   "WinTape"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   1548
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   45.074
      X2              =   5255.873
      Y1              =   2147.977
      Y2              =   2147.977
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Windows GUI and Code Development by Pete Todd"
      ForeColor       =   &H00000000&
      Height          =   348
      Left            =   180
      TabIndex        =   2
      Top             =   3240
      Width           =   3876
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
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub

