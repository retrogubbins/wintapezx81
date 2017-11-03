VERSION 5.00
Begin VB.Form frmGraph 
   Caption         =   "Wintape - Graph"
   ClientHeight    =   1008
   ClientLeft      =   1884
   ClientTop       =   7404
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1008
   ScaleWidth      =   8220
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Paint()
Dim x, f, g As Long
Dim ht
Cls

'this next line sets the number of samples to plot
GEnd_Sel = GStart_Sel + 1500
GSel_Points = GEnd_Sel - GStart_Sel

ht = ScaleHeight
PSet (0, ht - ((Gwave(0) / 255) * ht))

'plot to RAW file
x = GStart_Sel
For f = GStart_Sel To GEnd_Sel - 1
        Line -((f - GStart_Sel) / (GSel_Points - 1) * ScaleWidth, ht - ((Gwave(x) / 255) * ht)), RGB(255, 255, 255)
    x = x + 1
Next f

'Mark the currently found trigger points
For f = 0 To Gtrig_points - 1
    If Gtrig(f) > GStart_Sel And Gtrig(f) <= GEnd_Sel Then
       'work out position on graph for marker
        PSet (((Gtrig(f) - GStart_Sel) / (GSel_Points - 1)) * ScaleWidth, ht - (ht * (GTRIG_POINT / 255)))
        Line -(((Gtrig(f) - GStart_Sel) / (GSel_Points - 1)) * ScaleWidth, ht - (ht * (GTRIG_POINT / 255) + (ScaleHeight / 10))), RGB(255, 25, 255)
    End If
Next f

'Mark the currently found bits
g = 0
For f = 0 To Gbits - 1
    If Gbitpos(f) >= GStart_Sel And Gbitpos(f) < GEnd_Sel Then
       'work out position on graph for marker
        PSet (((Gbitpos(f) - GStart_Sel) / (GSel_Points - 1)) * ScaleWidth, ht - (ht * (GTRIG_POINT / 255) + (ScaleHeight / 10) / 2))
        Print Str(Gbit(f));
        g = g + 1
        If g = 8 Then
            PSet (((Gbitpos(f) - GStart_Sel) / (GSel_Points - 1)) * 1.1 * ScaleWidth, ht)
            Line -(((Gbitpos(f) - GStart_Sel) / (GSel_Points - 1)) * 1.1 * ScaleWidth, 0), RGB(25, 255, 255)
        End If
    End If
Next f

End Sub

Private Sub Form_Resize()

Form_Paint

End Sub

Private Sub Form_Unload(Cancel As Integer)

Unload frmcontrol

End Sub

