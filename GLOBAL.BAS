Attribute VB_Name = "global"
'GLOBAL module for Wintape

Option Explicit

Const constMAXCYCLES = 4500000
Const twiceMAXCYCLES = 9500000

Const constMAXBYTES = 32768
Const tenMAXBYTES = 327680

Public Gplot_points As Long
Public Gtrig_points As Long
Public Gprog_start As Integer

'This next bit is all awful
'should be either dynamic or random access to RAW file
'hope people have enough memory
Public Gwave(twiceMAXCYCLES) As Byte
Public Gtrig(constMAXCYCLES) As Long

Public Gbitpos(tenMAXBYTES) As Long
Public Gbit(tenMAXBYTES) As Byte 'this is probably overkill pete

Public Gbyte(constMAXBYTES) As Byte
Public Gbytepos(constMAXBYTES) As Long


Public Gselected_byte As LoadPictureConstants

Public otext As String
Public Gbytes As Integer
Public Gbits As Long

Public GStart_Sel As Long
Public GEnd_Sel As Long
Public GSel_Points As Long
Public GWave_Errors As Long

Public max As Byte
Public GTRIG_POINT As Integer
Public GCYCLE_LIMIT As Integer
Public GONE_CYCLES As Integer
Public GZERO_CYCLES As Integer
Public GSPIKE_FILTER As Integer



Public Function ReadDataFile(fname As String)
Dim i, f As Long
Dim fileno

i = 0
fileno = FreeFile(0)
Open fname For Binary Access Read As fileno
If Err Then MsgBox "Couldn't Read Data File:" & fname: End

Get #fileno, , Gwave
Gplot_points = LOF(fileno) - 1

frmcontrol.lblRawsamples.Caption = Str(Gplot_points)
frmcontrol.lblRawsamples.Refresh

Close fileno


frmcontrol.txtStatus.Text = "Idle"

End Function

Public Function WriteDataFile(fname As String, pnts As Integer)
Dim i, fileno

'Write an output file in .P format
fileno = FreeFile(0)
Open fname For Binary Access Write As fileno
If Err Then MsgBox "Couldn't Write Data File:" & fname: End

For i = 0 To pnts - 1
    Put #fileno, , Gbyte(Gprog_start + i)
Next i
Close fileno

frmcontrol.txtStatus.Text = "Idle"
End Function

Public Function ReadPFile(fname As String)
Dim fileno

fileno = FreeFile(0)
Open fname For Binary Access Read As fileno
If Err Then MsgBox "Couldn't Read P File:" & fname: End

Get #fileno, , Gbyte
Gbytes = LOF(fileno)

Close fileno

frmcontrol.txtStatus.Text = "Idle"

End Function

Sub digitise_tape()
Dim ht, x, f, g As Long
Dim last_val As Byte
Dim removed As Integer

'This reads the RAW file and turns it into bits
'The Gtrig array holds sample point positions
'of all the points triggered

'every time we get a low to high transition
'past TRIG_POINT
'mark it
If frmcontrol.chkRectify.Value = vbChecked Then
    frmcontrol.txtStatus.Text = "Rectifying Data"
    frmcontrol.txtStatus.Refresh
    For f = 0 To Gplot_points
        If Gwave(f) < 128 Then Gwave(f) = 255 - Gwave(f)
    Next f
End If

If frmcontrol.chkInvert.Value = vbChecked Then
    frmcontrol.txtStatus.Text = "Inverting Data"
    frmcontrol.txtStatus.Refresh
    For f = 0 To Gplot_points
        Gwave(f) = (255 - Gwave(f)) Mod 256
    Next f
End If



frmcontrol.txtStatus.Text = "Digitising Cycles"

last_val = 255
x = 0
For f = 0 To Gplot_points
        
    If Gwave(f) >= GTRIG_POINT And last_val < GTRIG_POINT Then
        If x < constMAXCYCLES Then
            Gtrig(x) = f
            x = x + 1
        End If
        If (x Mod 100) = 0 Then
            frmcontrol.lblCyclesRead.Caption = Str(x)
            frmcontrol.lblCyclesRead.Refresh
        End If
    End If
    last_val = Gwave(f)
Next f
Gtrig_points = x


If frmcontrol.chkFilter.Value = vbChecked Then

    frmcontrol.txtStatus.Text = "Filtering Noise"
frmcontrol.txtStatus.Refresh

    'sweep through array getting rid of lone spikes
    For f = 1 To Gtrig_points - 2
        removed = 0
        If (Gtrig(f) - Gtrig(f - 1) > GSPIKE_FILTER) And (Gtrig(f + 1) - Gtrig(f) > GSPIKE_FILTER) Then
            'remove this point from the list
            'try simply copying the next one
            Gtrig(f) = Gtrig(f - 1)
        End If
    Next f
End If
Beep
frmcontrol.txtStatus.Text = "Tape Digitised"

End Sub

Sub reset_arrays()
Dim f

frmcontrol.txtStatus.Text = "Clearing Arrays"

'clear out bit array
For f = 0 To tenMAXBYTES
    Gbit(f) = 0
Next f
'clear out byte array
For f = 0 To constMAXBYTES
    Gbyte(f) = 0
Next f
frmcontrol.txtStatus.Text = "Arrays Cleared"

Gtrig_points = 0
Gbits = 0
Gbytes = 0
End Sub

Sub get_bytes()
Dim x, f, g, cycles As Long
Dim last_trig As Long
Dim cycle_start As Long


frmcontrol.txtStatus.Text = "Getting BITS"

    'turn digitise array into individual bytes
    ' put in prog array
'Step through gtrig array

'if next array entry is outside cycle thresh
'limit then it is counted as the end of the current bit
x = 0
cycles = 0
GWave_Errors = 0
cycle_start = Gtrig(0)
For f = 1 To Gtrig_points - 1
    If Gtrig(f) - Gtrig(f - 1) > GCYCLE_LIMIT Then
        'End of a series of triggers
        'Now work out number of cycles
        If cycles >= GZERO_CYCLES Then Gbit(x) = 0
        If cycles >= GONE_CYCLES Then Gbit(x) = 1
        If cycles < GZERO_CYCLES Then
            GWave_Errors = GWave_Errors + 1
            frmcontrol.lblWaveerrors.Caption = Str(GWave_Errors)
            frmcontrol.lblWaveerrors.Refresh
        End If
        Gbitpos(x) = cycle_start
        x = x + 1
        cycles = 0
        cycle_start = Gtrig(f)
    End If
    cycles = cycles + 1
Next f

Gbits = x

frmcontrol.txtStatus.Text = "Getting BYTES"


'Wrap bits up into bytes
'awful language
Gbytes = 0
x = 0
For f = 0 To Gbits Step 8
    Gbyte(x) = 0
    Gbyte(x) = Gbyte(x) + Gbit(f) * 128
    Gbyte(x) = Gbyte(x) + Gbit(f + 1) * 64
    Gbyte(x) = Gbyte(x) + Gbit(f + 2) * 32
    Gbyte(x) = Gbyte(x) + Gbit(f + 3) * 16
    Gbyte(x) = Gbyte(x) + Gbit(f + 4) * 8
    Gbyte(x) = Gbyte(x) + Gbit(f + 5) * 4
    Gbyte(x) = Gbyte(x) + Gbit(f + 6) * 2
    Gbyte(x) = Gbyte(x) + Gbit(f + 7) * 1
    Gbytepos(x) = Gbitpos(f)
    x = x + 1
    Gbytes = x
    If (x Mod 100) = 0 Then
        frmcontrol.lblBytesRead.Caption = Str(Gbytes)
        frmcontrol.lblBytesRead.Refresh
    End If
Next f
 
End Sub

Sub get_progname()
Dim x, f, g As Long
Dim zxchr As String
Dim c As Byte

zxchr = "\~$:?()><=+-*/;,.0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

frmcontrol.txtStatus.Text = "Scanning Tape Header"
frmcontrol.txtStatus.Refresh

'get filename from tape data
frmcontrol.lblFname.Caption = ""

For f = 0 To 20
    c = Gbyte(f)
    If c > 128 Then
        c = c - 128 'invert the last byte of the name
        If ((c >= 11) And (c <= 63)) Then
            frmcontrol.lblFname.Caption = frmcontrol.lblFname.Caption & Mid$(zxchr, (c - 10), 1)
                           
            g = g + 1
            Exit For
        Else
            Exit For
        End If
    End If
    If ((c >= 11) And (c <= 63)) Then
        frmcontrol.lblFname.Caption = frmcontrol.lblFname.Caption & Mid$(zxchr, (c - 10), 1)
        g = g + 1
    End If
Next f
If g = 0 Then frmcontrol.lblFname.Caption = ""


End Sub



Sub get_prog(st As Integer)

'st is the starting byte of the program within the byte array
Dim x, f, g As Long
Dim zxchr As String
Dim lnum, lleng As Long
Dim linebuf As String
Dim prog_bytes As Long
zxchr = "\œ$:?()><=+-*/;,.0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

frmcontrol.txtStatus.Text = "Reading ZX81 Program"
frmcontrol.txtStatus.Refresh

'put prog bytes into text for output box
otext = ""
For f = 0 To Gbytes - 1
    If ((Gbyte(f) >= 11) And (Gbyte(f) <= 63)) Then
        otext = otext & Mid$(zxchr, (Gbyte(f) - 10), 1)
    Else
       otext = otext & "?"
    End If
    'next line gives vertical listing
    'otext = otext & Chr(8) & Str(Gbyte(f)) & Chr(8) & Gcharset(Gbyte(f)) & (Chr(13) + Chr(10))
Next f

frmcontrol.txtOutput.Text = otext
     
'Get zx81 variables to display on screen
'must be done here as its dependent upon the length
'of the filename
process_vars

frmcontrol.txtStatus.Text = "Creating Listing"
frmcontrol.txtStatus.Refresh

'convert to listing
frmcontrol.txtListing.Text = ""
f = Gprog_start + 116
'f = 122 ' Skip past filename,system variables etc.
'hmmm, don't really know what the format is here

x = CLng(CLng(Gbyte(Gprog_start + 4)) * 256) + Gbyte(Gprog_start + 3)
prog_bytes = (x - 16509) + 116 + 6

On Error Resume Next
'I hated writing that line - bodge
Do While f < prog_bytes - 1
        lnum = Gbyte(f) * 256 + Gbyte(f + 1)
        lleng = Gbyte(f + 2) + 256 * Gbyte(f + 3)
        'get line characters to expand
        f = f + 4
        If lleng < 1 Then Exit Do
        linebuf = ""
        For g = 0 To lleng - 1
            If Gbyte(f) <> 234 And Gbyte(f + g) = 126 Then
                g = g + 5
            Else
                linebuf = linebuf + Gcharset(Gbyte(f + g))
            End If
        Next g
        f = f + lleng
        frmcontrol.txtListing.Text = frmcontrol.txtListing.Text + _
            " " & Str(lnum) & " : " & linebuf & Chr(13) & Chr(10)
Loop
frmcontrol.txtListing.Text = frmcontrol.txtListing.Text + _
        "<END" & Chr(10) & Chr(13)
Beep
frmcontrol.txtStatus.Text = "Idle"


frmcontrol.lblBitsRead.Caption = Str(Gbits)
frmcontrol.lblBytesRead.Caption = Str(Gbytes - Len(frmcontrol.lblFname.Caption))
End Sub


Sub process_vars()
Dim x, s

Gprog_start = Len(frmcontrol.lblFname.Caption)
s = Gprog_start + 3
   x = CLng(CLng(Gbyte(s + 1)) * 256) + Gbyte(s)
   frmcontrol.lblDfile = "D-FILE: " & Str(x)
   x = CLng(CLng(Gbyte(s + 5)) * 256) + Gbyte(s + 4)
   frmcontrol.lblVars = "VARS: " & Str(x)
   x = CLng(CLng(Gbyte(s + 9)) * 256) + Gbyte(s + 8)
   frmcontrol.lblEline = "E-LINE: " & Str(x)
   x = CLng(CLng(Gbyte(s + 9)) * 256) + Gbyte(s + 8) - 16392
   frmcontrol.lblFsize = Str(x)
   x = CLng(CLng(Gbyte(s + 15)) * 256) + Gbyte(s + 14)
   frmcontrol.lblStkbot = "STKBOT: " & Str(x)
   x = CLng(CLng(Gbyte(s + 17)) * 256) + Gbyte(s + 16)
   frmcontrol.lblStkend = "STKEND: " & (x)
   x = CLng(CLng(Gbyte(s + 25)) * 256) + Gbyte(s + 24)
   frmcontrol.lblFrames = "FRAMES: " & Str(x)
End Sub

