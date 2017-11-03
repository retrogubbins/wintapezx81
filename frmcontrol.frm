VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmcontrol 
   Caption         =   "Wintape - Controller"
   ClientHeight    =   6480
   ClientLeft      =   1932
   ClientTop       =   564
   ClientWidth     =   8364
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   8364
   Begin VB.CommandButton comExit 
      Caption         =   "Exit"
      Height          =   432
      Left            =   6180
      TabIndex        =   52
      Top             =   6000
      Width           =   1632
   End
   Begin VB.CommandButton comDefaults 
      Caption         =   "Use Defaults"
      Height          =   252
      Left            =   5520
      TabIndex        =   50
      Top             =   1560
      Width           =   2772
   End
   Begin VB.CheckBox chkInvert 
      Caption         =   "Invert"
      Height          =   312
      Left            =   7560
      TabIndex        =   48
      Top             =   120
      Width           =   672
   End
   Begin VB.CommandButton comLoadpfile 
      Caption         =   "Load P-FILE"
      Height          =   252
      Left            =   6420
      TabIndex        =   47
      Top             =   4380
      Width           =   1152
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   2040
      Top             =   120
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   327680
   End
   Begin VB.CommandButton comAbout 
      Caption         =   "About/Help"
      Height          =   432
      Left            =   3420
      TabIndex        =   46
      Top             =   6000
      Width           =   1272
   End
   Begin VB.TextBox txtStartsel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   720
      TabIndex        =   45
      Text            =   "0"
      Top             =   1080
      Width           =   1632
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Filter"
      Height          =   312
      Left            =   7560
      TabIndex        =   44
      Top             =   1080
      Width           =   672
   End
   Begin VB.CheckBox chkRectify 
      Caption         =   "Rectify"
      Height          =   312
      Left            =   7560
      TabIndex        =   43
      Top             =   600
      Width           =   792
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   612
      Left            =   60
      TabIndex        =   41
      Top             =   5760
      Width           =   2352
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   252
         Left            =   120
         TabIndex        =   42
         Text            =   "Idle"
         Top             =   240
         Width           =   2112
      End
   End
   Begin VB.TextBox txtSpikefilter 
      Height          =   252
      Left            =   6900
      TabIndex        =   35
      Text            =   "8"
      Top             =   1260
      Width           =   552
   End
   Begin VB.TextBox txtListing 
      Height          =   1932
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   34
      Top             =   1860
      Width           =   4152
   End
   Begin VB.CommandButton comLoadFile 
      Caption         =   "Load RAW-FILE"
      Height          =   432
      Left            =   6180
      TabIndex        =   31
      Top             =   3900
      Width           =   1632
   End
   Begin VB.CommandButton comSavepfile 
      Caption         =   "Save P-FILE"
      Height          =   432
      Left            =   6180
      TabIndex        =   30
      Top             =   5460
      Width           =   1632
   End
   Begin VB.TextBox txtFilename 
      Height          =   252
      Left            =   4080
      TabIndex        =   26
      Top             =   4260
      Width           =   1272
   End
   Begin VB.TextBox txtZerocycles 
      Height          =   252
      Left            =   6900
      TabIndex        =   14
      Text            =   "4"
      Top             =   960
      Width           =   552
   End
   Begin VB.TextBox txtOnecycles 
      Height          =   252
      Left            =   6900
      TabIndex        =   12
      Text            =   "7"
      Top             =   660
      Width           =   552
   End
   Begin VB.TextBox txtCyclethresh 
      Height          =   252
      Left            =   6900
      TabIndex        =   11
      Text            =   "16"
      Top             =   360
      Width           =   552
   End
   Begin VB.TextBox txtTrigthresh 
      Height          =   252
      Left            =   6900
      TabIndex        =   10
      Text            =   "160"
      Top             =   60
      Width           =   552
   End
   Begin VB.TextBox txtOutput 
      Height          =   1932
      Left            =   4320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1860
      Width           =   3972
   End
   Begin VB.CommandButton cmdReadTape 
      Caption         =   "Read Tape"
      Height          =   672
      Left            =   6180
      TabIndex        =   0
      Top             =   4680
      Width           =   1632
   End
   Begin VB.Label Label16 
      Caption         =   "File Name"
      Height          =   252
      Left            =   4140
      TabIndex        =   53
      Top             =   3960
      Width           =   1092
   End
   Begin VB.Label Label17 
      Caption         =   "Header Name"
      Height          =   252
      Left            =   2640
      TabIndex        =   51
      Top             =   3960
      Width           =   1152
   End
   Begin VB.Label Label14 
      Caption         =   "ZX81 System Variables"
      Height          =   192
      Left            =   3000
      TabIndex        =   49
      Top             =   60
      Width           =   1752
   End
   Begin VB.Line Line6 
      X1              =   5520
      X2              =   5520
      Y1              =   3900
      Y2              =   6480
   End
   Begin VB.Line Line5 
      X1              =   2460
      X2              =   2460
      Y1              =   3840
      Y2              =   6480
   End
   Begin VB.Line Line4 
      X1              =   5460
      X2              =   5460
      Y1              =   1800
      Y2              =   0
   End
   Begin VB.Line Line3 
      X1              =   2460
      X2              =   2460
      Y1              =   1800
      Y2              =   0
   End
   Begin VB.Label Label15 
      Caption         =   "Expected Bytes"
      Height          =   252
      Left            =   2640
      TabIndex        =   40
      Top             =   5040
      Width           =   1152
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1740
      TabIndex        =   39
      Top             =   1440
      Width           =   612
   End
   Begin VB.Label lblByteno 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   720
      TabIndex        =   38
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label13 
      Caption         =   "Byte No"
      Height          =   252
      Left            =   60
      TabIndex        =   37
      Top             =   1500
      Width           =   612
   End
   Begin VB.Label Label12 
      Caption         =   "SPIKE FILTER"
      Height          =   192
      Left            =   5640
      TabIndex        =   36
      Top             =   1260
      Width           =   1212
   End
   Begin VB.Label lblRawsamples 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   960
      TabIndex        =   33
      Top             =   3900
      Width           =   1332
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Samples"
      Height          =   372
      Left            =   120
      TabIndex        =   32
      Top             =   3960
      Width           =   732
   End
   Begin VB.Label lblWaveerrors 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   960
      TabIndex        =   29
      Top             =   4620
      Width           =   1332
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Errors"
      Height          =   252
      Left            =   360
      TabIndex        =   28
      Top             =   4680
      Width           =   492
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000012&
      X1              =   60
      X2              =   2340
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   60
      X2              =   2340
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label lblFname 
      Alignment       =   1  'Right Justify
      Height          =   252
      Left            =   2640
      TabIndex        =   27
      Top             =   4260
      Width           =   1152
   End
   Begin VB.Label lblFrames 
      Caption         =   "*"
      Height          =   252
      Left            =   4080
      TabIndex        =   25
      Top             =   1380
      Width           =   1212
   End
   Begin VB.Label lblStkend 
      Caption         =   "*"
      Height          =   252
      Left            =   4080
      TabIndex        =   24
      Top             =   900
      Width           =   1212
   End
   Begin VB.Label lblStkbot 
      Caption         =   "*"
      Height          =   252
      Left            =   4080
      TabIndex        =   23
      Top             =   420
      Width           =   1212
   End
   Begin VB.Label lblFsize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   2640
      TabIndex        =   22
      Top             =   5340
      Width           =   1212
   End
   Begin VB.Label lblEline 
      Caption         =   "*"
      Height          =   252
      Left            =   2580
      TabIndex        =   21
      Top             =   1380
      Width           =   1212
   End
   Begin VB.Label lblVars 
      Caption         =   "*"
      Height          =   252
      Left            =   2580
      TabIndex        =   20
      Top             =   900
      Width           =   1212
   End
   Begin VB.Label Label10 
      Caption         =   "ZX81 Tape Reader"
      Height          =   252
      Left            =   60
      TabIndex        =   19
      Top             =   600
      Width           =   1512
   End
   Begin VB.Label Label9 
      Caption         =   "Wintape 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   60
      TabIndex        =   18
      Top             =   180
      Width           =   1932
   End
   Begin VB.Label lblDfile 
      Caption         =   "*"
      Height          =   252
      Left            =   2580
      TabIndex        =   17
      Top             =   420
      Width           =   1212
   End
   Begin VB.Label lblBytesd 
      Alignment       =   1  'Right Justify
      Caption         =   "Bytes"
      Height          =   252
      Left            =   120
      TabIndex        =   16
      Top             =   5400
      Width           =   732
   End
   Begin VB.Label lblBytesRead 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   960
      TabIndex        =   15
      Top             =   5340
      Width           =   1332
   End
   Begin VB.Label Label7 
      Caption         =   "ZERO-CYCLES"
      Height          =   192
      Left            =   5640
      TabIndex        =   13
      Top             =   960
      Width           =   1212
   End
   Begin VB.Label Label6 
      Caption         =   "CYCLE-THRESH"
      Height          =   192
      Left            =   5520
      TabIndex        =   9
      Top             =   360
      Width           =   1332
   End
   Begin VB.Label Label5 
      Caption         =   "TRIG-THRESH"
      Height          =   192
      Left            =   5640
      TabIndex        =   8
      Top             =   60
      Width           =   1212
   End
   Begin VB.Label Label4 
      Caption         =   "ONE-CYCLES"
      Height          =   192
      Left            =   5760
      TabIndex        =   7
      Top             =   660
      Width           =   1092
   End
   Begin VB.Label lblCyclesRead 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   960
      TabIndex        =   6
      Top             =   4260
      Width           =   1332
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Cycles"
      Height          =   312
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   612
   End
   Begin VB.Label lblBitsRead 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   960
      TabIndex        =   4
      Top             =   4980
      Width           =   1332
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Bits"
      Height          =   372
      Left            =   480
      TabIndex        =   3
      Top             =   5040
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "Sample"
      Height          =   252
      Left            =   60
      TabIndex        =   2
      Top             =   1140
      Width           =   552
   End
End
Attribute VB_Name = "frmcontrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Main Control Form Code
Option Explicit

Private Sub cmdReadTape_Click()
    'copy values off form into global vars
    GTRIG_POINT = Val(txtTrigthresh)
    GCYCLE_LIMIT = Val(txtCyclethresh)
    GONE_CYCLES = Val(txtOnecycles)
    GZERO_CYCLES = Val(txtZerocycles)
    GSPIKE_FILTER = Val(txtSpikefilter)
    
    reset_arrays
    
    'convert tape RAW into list of cycle times
    digitise_tape
    
    get_bytes
    get_progname 'from header bytes
    get_prog (Len(lblFname.Caption)) 'param is start of listing
    frmGraph.Refresh
    
End Sub

Private Sub comDefaults_Click()
txtTrigthresh = 160
txtCyclethresh = 16
txtOnecycles = 7
txtZerocycles = 4
txtSpikefilter = 8
chkInvert.Value = vbUnchecked
chkFilter.Value = vbUnchecked
chkRectify.Value = vbUnchecked

End Sub

Private Sub comExit_Click()

Unload Me

End Sub

Private Sub comSavepfile_Click()

'save a P file
Dim h
Dim fname As String
'Make up a suggested filename
fname = ""
'Open a Dialog box to get the load name
'Text (*.txt)|*.txt|Pictures (*.bmp;*.ico)|*.bmp;*.ico
CMDialog1.Filter = "P Files (*.P)|*.P|All Files (*.*)|*.*"
fname = GetFileName(fname, "Save")
If fname <> "" Then
    txtFilename.Text = fname
    'Test file to see what type it is
    ' Could be  X86 Binary
    '           SUN Binary
    '           ASCII
    h = WriteDataFile(Trim(fname), Val(lblFsize))
    Caption = "WinTape" & txtFilename
End If
End Sub


Private Sub comAbout_Click()

frmAbout.Show vbModal


End Sub

Private Sub comLoadpfile_Click()

'load a P file in

Dim fname As String
'Make up a suggested filename
fname = ""
'Open a Dialog box to get the load name
'Text (*.txt)|*.txt|Pictures (*.bmp;*.ico)|*.bmp;*.ico
CMDialog1.Filter = "P Files (*.P)|*.P|All Files (*.*)|*.*"
fname = GetFileName(fname, "Load")

If fname <> "" Then
    txtFilename.Text = fname
    'Test file to see what type it is
    ' Could be  X86 Binary
    '           SUN Binary
    '           ASCII
    ReadPFile (Trim(txtFilename.Text))
    Caption = "WinTape" & txtFilename
    'load in the zx byte file *.p
End If
    lblFname.Caption = ""
   
    get_prog (0)
    frmGraph.Refresh


End Sub


Private Sub Form_Load()

'should really load up registry defaults for
'the stuff here


charset_init 'set up the zx81 character set
Show

End Sub

Private Sub Form_Unload(Cancel As Integer)

Unload frmGraph
Unload Me


End Sub



Private Sub txtOutput_Click()
    'Find the current insertion point
    Gselected_byte = frmcontrol.txtOutput.SelStart

    'convert to a sample position
    GStart_Sel = Gbytepos(Gselected_byte)
    txtStartsel.Text = Str(GStart_Sel)
    lblByteno.Caption = Str(Gselected_byte)
    lblValue = Str(Gbyte(Gselected_byte))
    frmGraph.Refresh

End Sub

Private Sub txtOutput_KeyUp(KeyCode As Integer, Shift As Integer)

txtOutput_Click

End Sub

Private Sub comLoadFile_Click()

Dim fname As String
'Make up a suggested filename
fname = txtFilename.Text
'Open a Dialog box to get the load name
'Text (*.txt)|*.txt|Pictures (*.bmp;*.ico)|*.bmp;*.ico
CMDialog1.Filter = "RAW Files (*.raw)|*.raw|All Files (*.*)|*.*"
fname = GetFileName(fname, "Load")
If fname <> "" Then
    txtFilename.Text = fname
    'Test file to see what type it is
    ' Could be  X86 Binary
    '           SUN Binary
    '           ASCII
    ReadDataFile (Trim(txtFilename.Text))
    Caption = "WinTape" & txtFilename
    'load in the zx tape file zxtape.raw
    GStart_Sel = 0
    GEnd_Sel = 0
    frmGraph.Show
    frmGraph.Refresh
End If

End Sub

Function GetFileName(Filename As Variant, Mode As Variant)
    ' Display a Save As dialog box and return a filename.
    ' If the user chooses Cancel, return an empty string.
'    On Error Resume Next
    CMDialog1.Filename = Filename
    If Mode = "Save" Then
        CMDialog1.ShowSave
    Else
        CMDialog1.ShowOpen
    End If
    '32755
    If Err <> vbCancel Then    ' User chose Cancel.
        GetFileName = CMDialog1.Filename
    Else
        GetFileName = ""
    End If
End Function
