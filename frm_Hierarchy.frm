VERSION 5.00
Begin VB.Form frm_Hierarchy 
   Caption         =   "Hierarchy View - copyright (C) 2001/2002 Carriage Return software"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11895
   Icon            =   "FRM_HI~1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   11895
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Top             =   5640
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2280
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   12
      Top             =   5040
      Width           =   495
   End
   Begin VB.ListBox LSTtmp1 
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox PICcorner 
      Height          =   255
      Left            =   11640
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   4680
      Width           =   255
   End
   Begin VB.PictureBox PICrefresh 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   7800
      ScaleHeight     =   705
      ScaleWidth      =   2625
      TabIndex        =   5
      Top             =   3600
      Width           =   2655
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "REFRESHING..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.PictureBox PICeasymove 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   10800
      ScaleHeight     =   1455
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   2760
      Width           =   560
      Begin VB.CommandButton CMDwrite 
         Enabled         =   0   'False
         Height          =   495
         Left            =   0
         Picture         =   "FRM_HI~1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton CMDzoom 
         Height          =   495
         Left            =   0
         Picture         =   "FRM_HI~1.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox PICzoom 
         BackColor       =   &H00808080&
         Height          =   495
         Left            =   0
         Picture         =   "FRM_HI~1.frx":091E
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox PICwrite 
         BackColor       =   &H00808080&
         Height          =   495
         Left            =   0
         Picture         =   "FRM_HI~1.frx":0C28
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   8
         Top             =   480
         Width           =   495
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H0080C0FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   120
         Top             =   300
         Width           =   255
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H0080C0FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   120
         Top             =   60
         Width           =   255
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   240
         Top             =   180
         Width           =   255
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   0
         Top             =   180
         Width           =   255
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   100
      Left            =   0
      SmallChange     =   10
      TabIndex        =   3
      Top             =   4680
      Width           =   11655
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4695
      LargeChange     =   100
      Left            =   11640
      Max             =   15000
      SmallChange     =   10
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox PICbg 
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4635
      ScaleWidth      =   11595
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.PictureBox PICdraw 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   50000
         Left            =   0
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   49995
         ScaleWidth      =   40005
         TabIndex        =   2
         Top             =   -240
         Width           =   40000
         Begin VB.Timer SCROLLtimer 
            Interval        =   250
            Left            =   120
            Top             =   4080
         End
         Begin VB.Shape Shape1 
            BorderStyle     =   2  'Dash
            BorderWidth     =   2
            FillColor       =   &H008080FF&
            Height          =   735
            Left            =   2280
            Top             =   3720
            Width           =   1455
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            Index           =   0
            X1              =   2160
            X2              =   2160
            Y1              =   4440
            Y2              =   3240
         End
      End
   End
   Begin VB.Label Label4 
      Caption         =   "To add a person, type a name then right-click, hold and drag."
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   6000
      Width           =   3015
   End
   Begin VB.Label Label3 
      Height          =   1455
      Left            =   7680
      TabIndex        =   15
      Top             =   5040
      Width           =   3735
   End
   Begin VB.Label Label2 
      Height          =   1575
      Left            =   3240
      TabIndex        =   14
      Top             =   5040
      Width           =   4215
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   1
      Left            =   7680
      Picture         =   "FRM_HI~1.frx":0F32
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   0
      Left            =   7080
      Picture         =   "FRM_HI~1.frx":123C
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   10920
      Picture         =   "FRM_HI~1.frx":1546
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   10305
      Picture         =   "FRM_HI~1.frx":1850
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   9720
      Picture         =   "FRM_HI~1.frx":1B5A
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   9240
      Picture         =   "FRM_HI~1.frx":1E64
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   8760
      Picture         =   "FRM_HI~1.frx":216E
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frm_Hierarchy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! #
'#' Hierarchy display v1.0  copyright (C) 2001-2002 Carriage Return software            #
'#'                                                                                     #
'#' Hierarchy display is a set of function I wrote for a previous application and that  #
'#' is isolated here for demonstration purposes.                                        #
'#' Basically, these functions let you create, modify, store and retrieve a simple      #
'#' hierarchy design.                                                                   #
'#'                                                                                     #
'#' Originally, these function were written to be used with SQL server. For the sake    #
'#' of this example, I rewrote them to use plain text files instead. With a little      #
'#' thinking, it should be no problem to modify them to work with Access, SQL or any    #
'#' other database system you may use.                                                  #
'#'                                                                                     #
'#' The principle is simple: 2 tables. One contains the components of the hierarchy     #
'#' (in this case, persons) as well as their caracteristics (ID, name, position from    #
'#' top and from left of their icon, height and width of their icon) The second table   #
'#' contains a reference for each join established and is compsoed of the IDs of the    #
'#' two persons linked.                                                                 #
'#'                                                                                     #
'#' It has a basic zoom function, together with recentering features.                   #
'#'                                                                                     #
'#' Objects can be added, removed, moved, linked or unlinked.                           #
'#'                                                                                     #
'#' I hope you find this code useful to build upon....                                  #
'#'                                                                                     #
'#' You can reach me at chris@gillent.com                                               #
'#'                                                                                     #
'#' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! #


Private MAY_LINK_TO_MULTIPLE_SUPERIORS As Boolean

Private varBehaviour, varZoom
Private varPlacedLocs

Private varDrawX, varDrawY As Long

Private formHeight, formWidth
Private varDragClicX, varDragClicY As Long
Private arrayFramesPos(5000, 7)
Private arrayJoins(5000, 6)
Private varPrevSelID, varSelID, varClicID
Private varActFor
Private varMayDrop, varJoining As Boolean
Private varAutoScroll, varScrollOn As Integer
Private varRedrawed
Private varDragPicture
Private varRecenter, varInfo As Boolean

Private Const NORMAL = 0
Private Const ACTION1 = 1
Private Const ACTION2 = 2

Private Const constADD = 0
Private Const constREMOVE = 1
Private Const constUPDATE = 2

Private Const constUP = 0
Private Const constDOWN = 1
Private Const constALL = 2

Private Const ADD = 0
Private Const RESET = 1
Private Const COUNTALL = 2
Private Const COUNTREMAIN = 3
Private Const GETNEXTONE = 4

Private Sub CMDzoom_Click()
CMDzoom.Visible = False
'PICdraw.Enabled = False
varZoom = 4

'#' now that the zoom factor is set to 4, we need to redraw the contents of the screen
Call subHch_RedrawVisibleArea(0)

PICdraw.Left = 0
PICdraw.Top = 0
'#' And adapt the scrollbars values
VScroll1.Value = Abs(PICdraw.Top) / 15
HScroll1.Value = Abs(PICdraw.Left) / 15

End Sub

Private Sub Form_Load()

'#' tell the user what he can do
Label2 = "To highlight a person, left-click it." & vbCrLf & "To link two persons, hold down SHIFT and click the two persons one after another." & vbCrLf & "To unlink, repeat the same operation." & vbCrLf & "To move a person, right-click it, hold and drag."
Label3 = "Left-click (and hold if you want) on the orange direction pad to scroll." & vbCrLf & "Click the grey surface to refresh." & vbCrLf & "Click the magnifier button to enable/disable zoom." & vbCrLf & "When in zoom mode, double-click the hierarchy to recenter on a specific place."

Picture1.Picture = Image1(0).Picture

MAY_LINK_TO_MULTIPLE_SUPERIORS = False

'#' Zoom is off
varZoom = 1
varBehavior = NORMAL
varRecenter = False

formHeight = 6840
formWidth = 12000

'#' The scrolltimer is used to automatically scroll the screen when you drag a person
'#' near the edges of the display screen
SCROLLtimer.Enabled = True
SCROLLtimer.Interval = 125
varAutoScroll = 0

Me.Height = formHeight
Me.Width = formWidth
'Me.WindowState = 2

PICrefresh.Visible = False
PICrefresh.Left = (PICbg.Width / 2) - (PICrefresh.Width / 2)
PICrefresh.Top = (PICbg.Height / 2) - (PICrefresh.Height / 2)

'#' Shape1 is used to show what person is currently selected in the hierarchy
Shape1.Visible = False
'#' Line that figures the Join animation
Line1(0).Visible = False

varPrevSelID = ""

'#' read from the textfiles (db) the hierarchy objects and the existing joins
Call subHch_LoadObjects

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
varAutoScroll = 0
End Sub

Private Sub Form_Resize()
If Me.WindowState = 0 Then
    Me.Height = formHeight
    Me.Width = formWidth
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'#' kill arrays
Erase arrayFramesPos
Erase arrayJoins
End Sub

Private Sub HScroll1_Change()
'#' let's scroll
HSV = HScroll1.Value
PICdraw.Left = (HSV * 15) * -1
End Sub

Private Sub HScroll1_Scroll()
HSV = HScroll1.Value
PICdraw.Left = (HSV * 15) * -1
DoEvents
End Sub

Private Sub PICcorner_DblClick()
'#' When the square at the intersection of the scrollbars is double-clicked, we refresh
Call subHch_RedrawVisibleArea(0)
End Sub

Private Sub PICdraw_DblClick()
'#' If the zoom is on, double-clicking lets the user recenter the screen on a
'#' specific area
If varZoom > 2 Then
    Shape1.Width = PICbg.Width / varZoom
    Shape1.Height = PICbg.Height / varZoom
    Shape1.Left = varDrawX - (Shape1.Width / 2)
    Shape1.Top = varDrawY - (Shape1.Height / 2)
    Shape1.FillStyle = 5
    Shape1.Visible = True
    PICdraw.MousePointer = 15
    varRecenter = True
End If
End Sub

Private Sub PICdraw_DragDrop(Source As Control, X As Single, Y As Single)

'#' Something was dropped on the hierarchy area
Randomize Timer
varAutoScroll = 0
If varMayDrop = False Then Exit Sub

If Source.Name = "Picture1" Then
    '#' if we've just added a person.....
    varWidth = Image1(0).Width
    varHeight = Image1(0).Height
    '
    '
    '#' invent an ID (for a working application, we should ensure the ID is unique)
    varID = CStr(Int(Rnd * 1000000))
    varName = Text1.Text
    Text1 = ""
    
    '#' draw it on the screen
    PICdraw.PaintPicture Image1(varDragPicture).Picture, X - varDragClicX, Y - varDragClicY
    PICdraw.CurrentX = (X - varDragClicX + (varWidth / 2)) - (PICdraw.TextWidth(varName) / 2)
    PICdraw.CurrentY = Y - varDragClicY + Image1(varDragPicture).Height
    PICdraw.Print varName
    '#' and update the records, in memory as well as in the DB
    Call subHch_UpdateFrameArray(varID, varName, (X - varDragClicX) * varZoom, (Y - varDragClicY) * varZoom, varWidth * varZoom, varHeight * varZoom, constADD)
    Call subHch_UpdateFrameSQL(varID, (X - varDragClicX) * varZoom, (Y - varDragClicY) * varZoom, varWidth * varZoom, varHeight * varZoom)
    

ElseIf Source.Name = "PICdraw" Then
    
    '#' otherwise it's a person we've moved from one place to another
    '#' retrieve coordinates of the selected object (Shape)
    varX1 = Shape1.Left
    varY1 = Shape1.Top
    varX2 = varX1 + Shape1.Width
    varY2 = varY1 + Shape1.Height
    
    varID = PICdraw.Tag
        PICdraw.Tag = ""
    '#' retrieve its name
    varCaption = fnHch_GetObjectInfo(varID, 2)
    
    '#' and erase the joins that might exist between it and soemthing else
    Call subHch_EraseJoins(varID)
    
    '#' then eliminate it (no mercy)
    PICdraw.DrawMode = 13
    PICdraw.Line (varX1, varY1)-(varX2, varY2), vbWhite, BF
    
    Shape1.Visible = False
    
    '#' now let's draw it at its new location
    varWidth = Image1(varDragPicture).Width
    varHeight = Image1(varDragPicture).Height
    PICdraw.PaintPicture Image1(varDragPicture).Picture, X - varDragClicX, Y - varDragClicY
    PICdraw.CurrentX = (X - varDragClicX + (Image1(varDragPicture).Width / 2)) - (PICdraw.TextWidth(varCaption) / 2)
    PICdraw.CurrentY = Y - varDragClicY + Image1(varDragPicture).Height
    PICdraw.ForeColor = vbBlack
    PICdraw.DrawMode = 15
    PICdraw.Print varCaption
    '#' and update both memory arrays and database
    Call subHch_UpdateFrameArray(varID, Null, X - varDragClicX, Y - varDragClicY, varWidth, varHeight, constUPDATE)
    Call subHch_UpdateFrameSQL(varID, X - varDragClicX, Y - varDragClicY, varWidth, varHeight)
    '#' joins need to be calculated again
    Call subHch_RecalculateJoins(varID)
    '#' and of course redrawed
    Call subHch_RedrawJoins(varID)
    
End If
End Sub

Private Sub subHch_UpdateFrameArray(fID, fName, fLeft, fTop, fWidth, fHeight, fAct)
'#' this function will update (Add, remove and update) the memory array containing
'#' the hierarchy objects (persons) information
If fAct = constADD Then
    For xyz = 1 To 5000
        If arrayFramesPos(xyz, 1) = "" Then
            arrayFramesPos(xyz, 1) = fID
            arrayFramesPos(xyz, 2) = fName
            arrayFramesPos(xyz, 3) = fLeft
            arrayFramesPos(xyz, 4) = fTop
            arrayFramesPos(xyz, 5) = fWidth
            arrayFramesPos(xyz, 6) = fHeight
            Exit For
        End If
    Next xyz
ElseIf fAct = constREMOVE Then
    For xyz = 1 To 5000
        If arrayFramesPos(xyz, 1) = fID Then
            arrayFramesPos(xyz, 1) = ""
            arrayFramesPos(xyz, 2) = ""
            arrayFramesPos(xyz, 3) = ""
            arrayFramesPos(xyz, 4) = ""
            arrayFramesPos(xyz, 5) = ""
            arrayFramesPos(xyz, 6) = ""
            Exit For
        End If
    Next xyz
ElseIf fAct = constUPDATE Then
    For xyz = 1 To 5000
        If arrayFramesPos(xyz, 1) = fID Then
            If IsNull(fName) = False Then arrayFramesPos(xyz, 2) = fName
            arrayFramesPos(xyz, 3) = fLeft
            arrayFramesPos(xyz, 4) = fTop
            arrayFramesPos(xyz, 5) = fWidth
            arrayFramesPos(xyz, 6) = fHeight
            Exit For
        End If
    Next xyz
End If
End Sub

Private Sub subHch_UpdateJoinArray(fTopID, fTopX, fTopY, fLowID, fLowX, fLowY, fAct)
'#' this function will update (Add, remove and update) the memory array containing
'#' the join information (ID of both persons plus coordinates of both ends of the
'#' join line
If fAct = constADD Then
    For xyz = 1 To 5000
        If arrayJoins(xyz, 1) = "" Then
            arrayJoins(xyz, 1) = fTopID
            arrayJoins(xyz, 2) = fTopX
            arrayJoins(xyz, 3) = fTopY
            arrayJoins(xyz, 4) = fLowID
            arrayJoins(xyz, 5) = fLowX
            arrayJoins(xyz, 6) = fLowY
            Exit For
        End If
    Next xyz
ElseIf fAct = constREMOVE Then
    For xyz = 1 To 5000
        If (arrayJoins(xyz, 1) = fLowID And arrayJoins(xyz, 4) = fTopID) Or (arrayJoins(xyz, 1) = fTopID And arrayJoins(xyz, 4) = fLowID) Then
            arrayJoins(xyz, 1) = ""
            arrayJoins(xyz, 2) = ""
            arrayJoins(xyz, 3) = ""
            arrayJoins(xyz, 4) = ""
            arrayJoins(xyz, 5) = ""
            arrayJoins(xyz, 6) = ""
            Exit For
        End If
    Next xyz
ElseIf fAct = constUPDATE Then
    For xyz = 1 To 5000
        If arrayJoins(xyz, 1) = fTopID And arrayJoins(xyz, 4) = fLowID Then
            arrayJoins(xyz, 1) = fTopID
            arrayJoins(xyz, 2) = fTopX
            arrayJoins(xyz, 3) = fTopY
            arrayJoins(xyz, 4) = fLowID
            arrayJoins(xyz, 5) = fLowX
            arrayJoins(xyz, 6) = fLowY
            Exit For
        End If
    Next xyz
End If
End Sub


Private Sub PICdraw_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'#' we're dragging something over the hierarchy area.
If fnHch_IsThereFrame(X, Y) <> "||||" Then
    '#' if we're over another person, no drop possible
    Source.DragIcon = Image1(2).Picture
    varMayDrop = False
Else
    Source.DragIcon = Image1(varDragPicture).Picture
    varMayDrop = True
End If

'#'  Should we scroll the screen?
'#' if we reach the limits of the visible area, turn on AutoScroll
varTop = Abs(PICdraw.Top)
varBottom = varTop + PICbg.Height
varLeft = Abs(PICdraw.Left)
varRight = varLeft + PICbg.Width
varAutoScroll = 0
If Abs(Y) - varTop < 200 Then
    varAutoScroll = 2
End If
If Abs(Y) - varTop < 100 Then
    varAutoScroll = 1
End If
If varBottom - Abs(Y) < 200 Then
    varAutoScroll = 4
End If
If varBottom - Abs(Y) < 100 Then
    varAutoScroll = 3
End If

If Abs(X) - varLeft < 200 Then
    varAutoScroll = 6
End If
If Abs(X) - varLeft < 100 Then
    varAutoScroll = 5
End If
If varRight - Abs(X) < 200 Then
    varAutoScroll = 8
End If
If varRight - Abs(X) < 100 Then
    varAutoScroll = 7
End If

End Sub

Private Sub PICdraw_KeyDown(KeyCode As Integer, Shift As Integer)
'#' this is to respond to the DELETE key.
If varZoom <> 1 Then Exit Sub
'#' if tehre's someone selected.....
If Shape1.Visible = True Then
    If KeyCode = 46 Then
        varID = varClicID
        varNom = fnHch_GetObjectInfo(varID, 2)
        varSortBoth = Format(varLevel, "0#") & varNom
        '#' let's remove the person both from the array and the DB
        Call subHch_UpdateFrameArray(varClicID, "", 0, 0, 0, 0, constREMOVE)
        Call subHch_DeleteFrameSQL(varClicID)
        '#' delete any existing joins with this person
        Call subHch_DeleteJoin(varClicID)
        Call subHch_DeleteJoinSQL(varClicID)
        '#' and redraw everything
        Call subHch_RedrawVisibleArea(0)
        Shape1.Visible = False
    End If
End If
End Sub

Private Sub PICdraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'#' someone clicked the hierarchy area
If varRecenter = True Then
    '#' if we're in zoom mode and are willing to recenter.....
    Shape1.Visible = False
    Shape1.FillStyle = 1
    PICdraw.MousePointer = 0
    varRecenter = False
    
    varLMargin = Abs(PICdraw.Left) * varZoom
    varFromLeft = (Shape1.Left - Abs(PICdraw.Left)) * varZoom
    varNewX = (varLMargin + varFromLeft) * -1
    
    varTMargin = Abs(PICdraw.Top) * varZoom
    varFromTop = (Shape1.Top - Abs(PICdraw.Top)) * varZoom
    varNewY = (varTMargin + varFromTop) * -1
    
    varZoom = 1
    Call subHch_RedrawVisibleArea(0)
    PICdraw.Top = varNewY
    PICdraw.Left = varNewX
    
    CMDzoom.Visible = True
    
    VScroll1.Value = Abs(PICdraw.Top) / 15
    HScroll1.Value = Abs(PICdraw.Left) / 15
    
End If

If varZoom <> 1 Then Exit Sub '#' if in Zoom mode, go no further

'#' if we're pressing shift and are not joining yet, let's start to join
If (Shift And 1) = 1 And varJoining = False Then varJoining = True

If Button = 1 Then   '#' left click

    '#' get info about the clicked person
    varTmp = fnHch_IsThereFrame(X, Y)
    thisFrame = Split(varTmp, "|")
    '#' if we've clicked someone
    If thisFrame(0) <> "" Then
    
        If varBehaviour = NORMAL Then
            
            varClicID = thisFrame(0)
            If varSelID = "" And varPrevSelID <> "" Then
                varSelID = thisFrame(0)
            End If
            
            If varPrevSelID = "" And varJoining = True Then
                varPrevSelID = thisFrame(0)
            End If
            
            varIL = CLng(thisFrame(2)) / varZoom
            varIT = CLng(thisFrame(3)) / varZoom
            varIW = CLng(thisFrame(4)) / varZoom
            varIH = CLng(thisFrame(5)) / varZoom
            
            '#' highlight the clicked person by adjusting the shape position
            '#' and making it visible
            With Shape1
                .Left = (varIL + (varIW / 2)) - ((PICdraw.TextWidth(thisFrame(1)) + 300) / 2) 'thisFrame(2)
                .Top = varIT  'thisFrame(3)
                .Width = PICdraw.TextWidth(thisFrame(1)) + 300 'thisFrame(4)
                .Height = varIH + PICdraw.TextHeight(thisFrame(1)) + 100 'thisFrame(5)
                .Visible = True
            End With
            
            '#' if we're joining...
            If varJoining = True Then
                '#' if the line not visible yet, it means we've not selected
                '#' the first element of the join yet
                If Line1(0).Visible = False Then
                    '#' we're showing the line object that will visually represent the
                    '#' join in progress
                    With Line1(0)
                        .Visible = True
                        .Tag = thisFrame(0)
                        .X1 = Shape1.Left + (Shape1.Width / 2)
                        .Y1 = Shape1.Top + (Shape1.Height / 2)
                        .X2 = X
                        .Y2 = Y
                    End With
                    
                Else
                
                    '#' otherwise, we've just selected the second element
                    '#' hide the animation (fake) line
                    Line1(0).Visible = False
                    
                    '#' if we're not joining someone with himself, ok
                    If varPrevSelID <> varSelID Then
                        
                        '#' calculate the shortest line for the join
                        varTmp = fnHch_GetBestJoin(Line1(0).X1, Line1(0).Y1, Shape1.Left + (Shape1.Width / 2), Shape1.Top + (Shape1.Height / 2))
                        varTmp = Split(varTmp, "|")
                        
                        '#' and draw the join line
                        '#' IMPORTANT: we're using Drawmode 6 which is 'invert'
                        '#' This way, we're getting a black line on a white background.
                        '#' We'll use the same operation to erase the join, returning to
                        '#' our original white background
                        PICdraw.DrawMode = 6
                            PICdraw.Line (varTmp(0), varTmp(1))-(varTmp(2), varTmp(3)), varDrawColor
                        PICdraw.DrawMode = 13
                        
                        '#' if there was already a join between those 2
                        If fnHch_IsJoin(varPrevSelID, varSelID) = True Then
                            '#' we erase the join
                            Call subHch_UpdateJoinArray(varPrevSelID, varTmp(0), varTmp(1), varSelID, varTmp(2), varTmp(3), constREMOVE)
                        Else
                            '#' otherwise we create it in memory
                            Call subHch_UpdateJoinArray(varPrevSelID, varTmp(0), varTmp(1), varSelID, varTmp(2), varTmp(3), constADD)
                        End If
                        '#' and update the join database
                        Call subHch_UpdateJoinSQL(varPrevSelID, varSelID)
                    
                    End If
                    
                    varJoining = False
                    Shape1.Visible = False
                    Shape1.Tag = ""
                    varPrevSelID = ""
                    varSelID = ""
                End If
            End If
        
        End If
    
    Else
        '#' noone was clicked
        If varBehaviour = NORMAL Then
            
            varClicID = ""
            varJoining = False
            Shape1.Visible = False
            CMDwrite.Enabled = False
            Shape1.Tag = ""
            varPrevSelID = ""
            varSelID = ""
            
        End If
    End If

ElseIf Button = 2 Then  '#' right button. Probably the start of a drag and drop

    varJoining = False '#' if we were joining, stop the join now
    varTmp = fnHch_IsThereFrame(X, Y)
    thisFrame = Split(varTmp, "|")
    If thisFrame(0) <> "" Then
        varIL = CLng(thisFrame(2))
        varIT = CLng(thisFrame(3))
        varIW = CLng(thisFrame(4))
        varIH = CLng(thisFrame(5))
        varDragClicX = X - varIL
        varDragClicY = Y - varIT
        
        '#' highlight the clicked person
        With Shape1
            .Left = (varIL + (varIW / 2)) - ((PICdraw.TextWidth(thisFrame(1)) + 300) / 2) 'thisFrame(2)
            .Top = varIT  'thisFrame(3)
            .Width = PICdraw.TextWidth(thisFrame(1)) + 300 'thisFrame(4)
            .Height = varIH + PICdraw.TextHeight(thisFrame(1)) + 100 'thisFrame(5)
            .Visible = True
        End With
        
        If InStr(thisFrame(0), "X") <> 0 Then varDragPicture = 1 Else varDragPicture = 0
        '#' show a drag image and let's start to drag
        PICdraw.DragIcon = Image1(varDragPicture).Picture
        PICdraw.Tag = thisFrame(0)
        PICdraw.Drag vbBeginDrag
    End If
    
End If
End Sub

Private Function fnHch_IsThereFrame(fX, fY) As String
'#' this function verifies if there is a person at the clicked location
'#' if yes, its informations are retrieved
varFrame = "||||"

For xyz = 1 To 5000
    If arrayFramesPos(xyz, 1) <> "" Then
        If (fX >= (arrayFramesPos(xyz, 3) / varZoom) And fX <= (arrayFramesPos(xyz, 3) + arrayFramesPos(xyz, 5)) / varZoom) And (fY >= arrayFramesPos(xyz, 4) / varZoom And fY <= (arrayFramesPos(xyz, 4) + arrayFramesPos(xyz, 6)) / varZoom) Then
            '#' Yesssssss!!!
            varFrame = arrayFramesPos(xyz, 1) & "|"
            varFrame = varFrame & arrayFramesPos(xyz, 2) & "|"
            varFrame = varFrame & arrayFramesPos(xyz, 3) & "|"
            varFrame = varFrame & arrayFramesPos(xyz, 4) & "|"
            varFrame = varFrame & arrayFramesPos(xyz, 5) & "|"
            varFrame = varFrame & arrayFramesPos(xyz, 6)
            Exit For
        End If
    End If
Next xyz
fnHch_IsThereFrame = varFrame
End Function

Private Sub PICdraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
varDrawX = X: varDrawY = Y

'#' if we're busy making a join, the animated line position needs to be updated
If varJoining = True Then
    With Line1(0)
        .X2 = X
        .Y2 = Y
    End With
Else
    Line1(0).Visible = False
End If

'#' and if we're recentering, the target shape must move together with the mouse
If varRecenter = True Then
    Shape1.Left = X - (Shape1.Width / 2)
    Shape1.Top = Y - (Shape1.Height / 2)
End If

'#' Should we scroll the screen?
If varJoining = True Or varRecenter = True Then
    varTop = Abs(PICdraw.Top)
    varBottom = varTop + PICbg.Height
    varLeft = Abs(PICdraw.Left)
    varRight = varLeft + PICbg.Width
    varAutoScroll = 0
    If Abs(Y) - varTop < 200 Then
        varAutoScroll = 2
    End If
    If Abs(Y) - varTop < 100 Then
        varAutoScroll = 1
    End If
    If varBottom - Abs(Y) < 200 Then
        varAutoScroll = 4
    End If
    If varBottom - Abs(Y) < 100 Then
        varAutoScroll = 3
    End If
    
    If Abs(X) - varLeft < 200 Then
        varAutoScroll = 6
    End If
    If Abs(X) - varLeft < 100 Then
        varAutoScroll = 5
    End If
    If varRight - Abs(X) < 200 Then
        varAutoScroll = 8
    End If
    If varRight - Abs(X) < 100 Then
        varAutoScroll = 7
    End If
End If

End Sub

Private Sub PICeasymove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'#' If the Easymove picture is clicked, we scroll the screen in the appropriate direction
varScrollOn = True
varAutoScroll = 0
If (X > Shape2(0).Left And X < Shape2(0).Left + Shape2(0).Width) And (Y > Shape2(0).Top And Y < Shape2(0).Top + Shape2(0).Height) Then
    ' ~~ Left
    varAutoScroll = 5
End If
If (X > Shape2(1).Left And X < Shape2(1).Left + Shape2(1).Width) And (Y > Shape2(1).Top And Y < Shape2(1).Top + Shape2(1).Height) Then
    ' ~~ Right
    varAutoScroll = 7
End If
If (X > Shape2(2).Left And X < Shape2(2).Left + Shape2(2).Width) And (Y > Shape2(2).Top And Y < Shape2(2).Top + Shape2(2).Height) Then
    ' ~~ Top
    varAutoScroll = 1
End If
If (X > Shape2(3).Left And X < Shape2(3).Left + Shape2(3).Width) And (Y > Shape2(3).Top And Y < Shape2(3).Top + Shape2(3).Height) Then
    ' ~~ Bottom
    varAutoScroll = 3
End If

If varAutoScroll = 0 Then
    Call subHch_RedrawVisibleArea(0)
End If

End Sub

Private Sub PICeasymove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'#' and we adapt the direction if the clicking perdures
If varScrollOn = True Then
    If (X > Shape2(0).Left And X < Shape2(0).Left + Shape2(0).Width) And (Y > Shape2(0).Top And Y < Shape2(0).Top + Shape2(0).Height) Then
        ' ~~ Left
        varAutoScroll = 5
    End If
    If (X > Shape2(1).Left And X < Shape2(1).Left + Shape2(1).Width) And (Y > Shape2(1).Top And Y < Shape2(1).Top + Shape2(1).Height) Then
        ' ~~ Right
        varAutoScroll = 7
    End If
    If (X > Shape2(2).Left And X < Shape2(2).Left + Shape2(2).Width) And (Y > Shape2(2).Top And Y < Shape2(2).Top + Shape2(2).Height) Then
        ' ~~ Top
        varAutoScroll = 1
    End If
    If (X > Shape2(3).Left And X < Shape2(3).Left + Shape2(3).Width) And (Y > Shape2(3).Top And Y < Shape2(3).Top + Shape2(3).Height) Then
        ' ~~ Bottom
        varAutoScroll = 3
    End If
End If
End Sub

Private Sub PICeasymove_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
varAutoScroll = 0
varScrollOn = False
End Sub

Private Function fnHch_GetBestJoin(fX1, fY1, fX2, fY2)
'#' this function calculates what the best start and end point should be for the join line

'#' get the coordinates of the two objects to link
varFr1 = fnHch_IsThereFrame(fX1, fY1)
varFr2 = fnHch_IsThereFrame(fX2, fY2)

varFrm1 = Split(varFr1, "|")
varFrm2 = Split(varFr2, "|")

'#' create 2 work arrays
ReDim arraySquare(2, 8, 2) '1 = left; 2 = top
    arraySquare(1, 1, 1) = CLng(varFrm1(2))
        arraySquare(1, 1, 2) = CLng(varFrm1(3))
    arraySquare(1, 2, 1) = CLng(varFrm1(2)) + (CLng(varFrm1(4)) / 2)
        arraySquare(1, 2, 2) = CLng(varFrm1(3))
    arraySquare(1, 3, 1) = CLng(varFrm1(2)) + CLng(varFrm1(4))
        arraySquare(1, 3, 2) = CLng(varFrm1(3))
    arraySquare(1, 4, 1) = CLng(varFrm1(2)) + CLng(varFrm1(4))
        arraySquare(1, 4, 2) = CLng(varFrm1(3)) + (CLng(varFrm1(5)) / 2)
    arraySquare(1, 5, 1) = CLng(varFrm1(2)) + CLng(varFrm1(4))
        arraySquare(1, 5, 2) = CLng(varFrm1(3)) + CLng(varFrm1(5))
    arraySquare(1, 6, 1) = CLng(varFrm1(2)) + (CLng(varFrm1(4)) / 2)
        arraySquare(1, 6, 2) = CLng(varFrm1(3)) + CLng(varFrm1(5))
    arraySquare(1, 7, 1) = CLng(varFrm1(2))
        arraySquare(1, 7, 2) = CLng(varFrm1(3)) + CLng(varFrm1(5))
    arraySquare(1, 8, 1) = CLng(varFrm1(2))
        arraySquare(1, 8, 2) = CLng(varFrm1(3)) + (CLng(varFrm1(5)) / 2)
    
    arraySquare(2, 1, 1) = CLng(varFrm2(2))
        arraySquare(2, 1, 2) = CLng(varFrm2(3))
    arraySquare(2, 2, 1) = CLng(varFrm2(2)) + (CLng(varFrm2(4)) / 2)
        arraySquare(2, 2, 2) = CLng(varFrm2(3))
    arraySquare(2, 3, 1) = CLng(varFrm2(2)) + CLng(varFrm2(4))
        arraySquare(2, 3, 2) = CLng(varFrm2(3))
    arraySquare(2, 4, 1) = CLng(varFrm2(2)) + CLng(varFrm2(4))
        arraySquare(2, 4, 2) = CLng(varFrm2(3)) + (CLng(varFrm2(5)) / 2)
    arraySquare(2, 5, 1) = CLng(varFrm2(2)) + CLng(varFrm2(4))
        arraySquare(2, 5, 2) = CLng(varFrm2(3)) + CLng(varFrm2(5))
    arraySquare(2, 6, 1) = CLng(varFrm2(2)) + (CLng(varFrm2(4)) / 2)
        arraySquare(2, 6, 2) = CLng(varFrm2(3)) + CLng(varFrm2(5))
    arraySquare(2, 7, 1) = CLng(varFrm2(2))
        arraySquare(2, 7, 2) = CLng(varFrm2(3)) + CLng(varFrm2(5))
    arraySquare(2, 8, 1) = CLng(varFrm2(2))
        arraySquare(2, 8, 2) = CLng(varFrm2(3)) + (CLng(varFrm2(5)) / 2)
    

'#' Remember Pythagore ???? we'll use this formula to find the shortest path
'#' between the two objects
    
    ReDim varBest(3)
    varBest(1) = 1000000
    For xyz = 1 To 8
        For abc = 1 To 8
            varH = Abs(arraySquare(1, xyz, 2) - arraySquare(2, abc, 2))
            varW = Abs(arraySquare(1, xyz, 1) - arraySquare(2, abc, 1))
            varHyp = Int(Sqr((varH * varH) + (varW * varW)))
            If varHyp < varBest(1) Then
                varBest(1) = varHyp
                varBest(2) = xyz
                varBest(3) = abc
            ElseIf varHyp = varBest(1) Then
                '#' Give favor to medium values
                If (xyz / 2 = xyz \ 2) And (varBest(2) / 2 <> varBest(2) \ 2) Then
                    varBest(1) = varHyp
                    varBest(2) = xyz
                    varBest(3) = abc
                End If
            End If
        Next abc
    Next xyz

fnHch_GetBestJoin = arraySquare(1, varBest(2), 1) & "|" & arraySquare(1, varBest(2), 2) & "|" & arraySquare(2, varBest(3), 1) & "|" & arraySquare(2, varBest(3), 2)

End Function

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    '#' We're starting the dragging of a new object
    If Text1 = "" Then
        MsgBox "Please give this person a name"
    Else
        Picture1.Drag vbBeginDrag
    End If
End If
End Sub

Private Sub PICzoom_Click()
'#' zoom is requested
varRecenter = False
CMDzoom.Visible = True
PICdraw.Enabled = True
Shape1.Visible = False
varZoom = 1
Call subHch_RedrawVisibleArea(0)
End Sub

Private Sub SCROLLtimer_Timer()
'#' if we're auto-scrolling, the timer will update the scrolling every 0,125 seconds
If varAutoScroll <> 0 Then
    Select Case varAutoScroll
        Case 1, 2
        If varAutoScroll = 1 Then vAS = 6 Else vAS = 1
        varTop = Abs(PICdraw.Top)
        If varTop > (50 * vAS) Then
            PICdraw.Top = PICdraw.Top + (50 * vAS)
            VScroll1.Value = Abs(PICdraw.Top) / 15
        End If
        
        Case 3, 4
        If varAutoScroll = 3 Then vAS = 6 Else vAS = 1
        varTop = Abs(PICdraw.Top)
        If (varTop + PICbg.Height) < PICdraw.Height Then
            PICdraw.Top = PICdraw.Top - (50 * vAS)
            VScroll1.Value = Abs(PICdraw.Top) / 15
        End If
        
        Case 5, 6
        If varAutoScroll = 5 Then vAS = 6 Else vAS = 1
        varLeft = Abs(PICdraw.Left)
        If varLeft > (50 * vAS) Then
            PICdraw.Left = PICdraw.Left + (50 * vAS)
            HScroll1.Value = Abs(PICdraw.Left) / 15
        End If
        
        Case 7, 8
        If varAutoScroll = 7 Then vAS = 6 Else vAS = 1
        varLeft = Abs(PICdraw.Left)
        If (varLeft + PICbg.Width) < PICdraw.Width Then
            PICdraw.Left = PICdraw.Left - (50 * vAS)
            HScroll1.Value = Abs(PICdraw.Left) / 15
        End If
        
        Case Else
    End Select
End If
End Sub

Private Sub VScroll1_Change()
VSV = VScroll1.Value
PICdraw.Top = (VSV * 15) * -1
If varInfo = True Then
    varAbs = Shape1.Left - Abs(PICdraw.Left)
    If varAbs > (PICbg.Width / 2) Then
        PICpopup.Left = Abs(PICdraw.Left) + 100
    Else
        PICpopup.Left = Abs(PICdraw.Left) + PICbg.Width - 220 - PICpopup.Width
    End If
    PICpopup.Top = Abs(PICdraw.Top) + 120
    PICshadow.Left = PICpopup.Left + 120
    PICshadow.Top = PICpopup.Top + 120
End If
End Sub

Private Function fnHch_IsJoin(fID1, fID2)
'#' returns if a join already exists between two specified objects
varFound = False
For xyz = 1 To 5000
    If arrayJoins(xyz, 1) <> "" Then
        If arrayJoins(xyz, 1) = fID1 Then
            If arrayJoins(xyz, 4) = fID2 Then
                varFound = True
                Exit For
            End If
        ElseIf arrayJoins(xyz, 1) = fID2 Then
            If arrayJoins(xyz, 4) = fID1 Then
                varFound = True
                Exit For
            End If
        End If
    End If
Next xyz
fnHch_IsJoin = varFound
End Function

Private Sub VScroll1_Scroll()
VSV = VScroll1.Value
PICdraw.Top = (VSV * 15) * -1
If varInfo = True And PICpopup.Visible = True Then
    varAbs = Shape1.Left - Abs(PICdraw.Left)
    If varAbs > (PICbg.Width / 2) Then
        PICpopup.Left = Abs(PICdraw.Left) + 100
    Else
        PICpopup.Left = Abs(PICdraw.Left) + PICbg.Width - 220 - PICpopup.Width
    End If
    PICpopup.Top = Abs(PICdraw.Top) + 120
    PICshadow.Left = PICpopup.Left + 120
    PICshadow.Top = PICpopup.Top + 120
End If
DoEvents
End Sub

Private Sub subHch_EraseJoins(fRelatedTo)
'#' this sub erases Joint lines from the Screen,
'#' not from the memory array, if they are related to
'#' 'fRelatedTo'
For xyz = 1 To 5000
    If arrayJoins(xyz, 1) <> "" Then
        If (arrayJoins(xyz, 1) = fRelatedTo Or arrayJoins(xyz, 4) = fRelatedTo) Or fRelatedTo = "ALL" Then
            varX1 = arrayJoins(xyz, 2)
            varY1 = arrayJoins(xyz, 3)
            varX2 = arrayJoins(xyz, 5)
            varY2 = arrayJoins(xyz, 6)
            PICdraw.DrawMode = vbInvert
            PICdraw.Line (varX1, varY1)-(varX2, varY2), vbRed
        End If
    End If
Next xyz
PICdraw.DrawMode = 13
End Sub

Private Sub subHch_RecalculateJoins(fRelatedTo)
'#' this sub is called after objects have been moved on
'#' the screen, and recalculates Joint lines related to
'#' 'fRelatedTo'
For xyz = 1 To 5000
    If arrayJoins(xyz, 1) <> "" Then
        If (arrayJoins(xyz, 1) = fRelatedTo Or arrayJoins(xyz, 4) = fRelatedTo) Or fRelatedTo = "ALL" Then
            varID1 = arrayJoins(xyz, 1)
            varID2 = arrayJoins(xyz, 4)
            '#' Get a point from into each object
            varTmp = fnHch_GetObjectInfo(varID1, 3)
                If varTmp <> "" Then
                    varX1 = CLng(fnHch_GetObjectInfo(varID1, 3)) + 10
                    varY1 = CLng(fnHch_GetObjectInfo(varID1, 4)) + 10
                Else
                    GoTo nextloop
                End If
            varTmp = fnHch_GetObjectInfo(varID2, 3)
                If varTmp <> "" Then
                    varX2 = CLng(fnHch_GetObjectInfo(varID2, 3)) + 10
                    varY2 = CLng(fnHch_GetObjectInfo(varID2, 4)) + 10
                Else
                    GoTo nextloop
                End If
            
            '#' and calculate the best joint again
            varTmp = fnHch_GetBestJoin(varX1 / varZoom, varY1 / varZoom, varX2 / varZoom, varY2 / varZoom)
            varTmp = Split(varTmp, "|")
            
            arrayJoins(xyz, 2) = varTmp(0)
            arrayJoins(xyz, 3) = varTmp(1)
            arrayJoins(xyz, 5) = varTmp(2)
            arrayJoins(xyz, 6) = varTmp(3)
        End If
    End If
    
nextloop:

Next xyz
PICdraw.DrawMode = 13
End Sub

Private Sub subHch_RedrawJoins(fRelatedTo, Optional ByVal fDir = 2)
'#' this sub erases Joint lines from the Screen,
'#' not from the memory array, if they are related to
'#' 'fRelatedTo'
For xyz = 1 To 5000
    If arrayJoins(xyz, 1) <> "" Then
        If fDir = constUP Then
            
            If arrayJoins(xyz, 1) = fRelatedTo Or fRelatedTo = "ALL" Then
                If varRedrawed <> "" Then
                    If InStr(varRedrawed, "|" & xyz & "|") <> 0 Then GoTo nextloop
                End If
                    
                varX1 = arrayJoins(xyz, 2) / varZoom
                varY1 = arrayJoins(xyz, 3) / varZoom
                varX2 = arrayJoins(xyz, 5) / varZoom
                varY2 = arrayJoins(xyz, 6) / varZoom
                PICdraw.DrawMode = vbInvert
                PICdraw.Line (varX1, varY1)-(varX2, varY2), vbRed
                    
                If varRedrawed <> "" Then varRedrawed = varRedrawed & xyz & "|"
            
            End If
        
        ElseIf fDir = constDOWN Then
            
            If arrayJoins(xyz, 4) = fRelatedTo Or fRelatedTo = "ALL" Then
                If varRedrawed <> "" Then
                    If InStr(varRedrawed, "|" & xyz & "|") <> 0 Then GoTo nextloop
                End If
                    
                varX1 = arrayJoins(xyz, 2) / varZoom
                varY1 = arrayJoins(xyz, 3) / varZoom
                varX2 = arrayJoins(xyz, 5) / varZoom
                varY2 = arrayJoins(xyz, 6) / varZoom
                PICdraw.DrawMode = vbInvert
                PICdraw.Line (varX1, varY1)-(varX2, varY2), vbRed
                    
                If varRedrawed <> "" Then varRedrawed = varRedrawed & xyz & "|"
            
            End If
        
        ElseIf fDir = constALL Then
        
            If (arrayJoins(xyz, 1) = fRelatedTo Or arrayJoins(xyz, 4) = fRelatedTo) Or fRelatedTo = "ALL" Then
                If varRedrawed <> "" Then
                    If InStr(varRedrawed, "|" & xyz & "|") <> 0 Then GoTo nextloop
                End If
                    
                varX1 = arrayJoins(xyz, 2) / varZoom
                varY1 = arrayJoins(xyz, 3) / varZoom
                varX2 = arrayJoins(xyz, 5) / varZoom
                varY2 = arrayJoins(xyz, 6) / varZoom
                PICdraw.DrawMode = vbInvert
                PICdraw.Line (varX1, varY1)-(varX2, varY2), vbRed
                    
                If varRedrawed <> "" Then varRedrawed = varRedrawed & xyz & "|"
            
            End If
        
        End If
    End If

nextloop:

Next xyz
PICdraw.DrawMode = 13
End Sub

Private Function fnHch_GetObjectInfo(fID, fIdx)
'#' get info about an object
varCoord = ""
For abc = 1 To 5000
    If arrayFramesPos(abc, 1) = fID Then
        varCoord = arrayFramesPos(abc, fIdx)
        Exit For
    End If
Next abc
fnHch_GetObjectInfo = varCoord
End Function

Private Sub subHch_RedrawVisibleArea(fMod)
'#' Sometimes things get messed up. This sub ensures
'#' that the visible area is always clean.

'#' Determine dimensions of the View window

PICrefresh.Visible = True
Me.MousePointer = 13
DoEvents

varViewX1 = Abs(PICdraw.Left)
varViewY1 = Abs(PICdraw.Top)
varViewX2 = varViewX1 + PICbg.Width
varViewY2 = varViewY1 + PICbg.Height

PICdraw.Cls
PICdraw.FontSize = 6.75 / varZoom

'#' Redraw all objects that are in the visible area
'#' and the Joints to those to which they are linked
varIDList = ""
For xyz = 1 To 5000
    If arrayFramesPos(xyz, 1) <> "" Then
                            
        varCaption = arrayFramesPos(xyz, 2)
        varID = arrayFramesPos(xyz, 1)
        If InStr(varID, "X") <> 0 Then varPic = 1 Else varPic = 0
        
        varDiv = IIf(varZoom = 1, 1, 2)
        If fMod = 0 Then
            PICdraw.PaintPicture Image1(varPic).Picture, arrayFramesPos(xyz, 3) / varZoom, arrayFramesPos(xyz, 4) / varZoom, Image1(varPic).Width / varDiv, Image1(varPic).Height / varDiv
        Else
            PICdraw.PaintPicture Image3(varPic).Picture, arrayFramesPos(xyz, 3) / varZoom, arrayFramesPos(xyz, 4) / varZoom, Image1(varPic).Width / varDiv, Image1(varPic).Height / varDiv
        End If
        PICdraw.ForeColor = vbBlack
        PICdraw.DrawMode = 13
        varMode = IIf(varZoom = 1, 1, -1)
        PICdraw.CurrentX = (arrayFramesPos(xyz, 3) / varZoom) - ((PICdraw.TextWidth(varCaption) / 2) / varZoom) + ((Image1(0).Width / 2) * varMode)
        PICdraw.CurrentY = (arrayFramesPos(xyz, 4) / varZoom) + (Image1(0).Height / varDiv)
        PICdraw.Print varCaption
        
        If InStr(varIDList, varID) = 0 Then varIDList = varIDList & varID & "|"
    
    End If
Next xyz

If varIDList <> "" And fMod = 0 Then
    varRedrawed = "|"
    varIDList = Left(varIDList, Len(varIDList) - 1)
    varList = Split(varIDList, "|")
    For xyz = 0 To UBound(varList)
        Call subHch_RedrawJoins(varList(xyz))
    Next xyz
    varRedrawed = ""
End If

Me.MousePointer = 0
PICrefresh.Visible = False

End Sub

Private Sub subHch_UpdateFrameSQL(fID, fLeft, fTop, fWidth, fHeight)
'#' this sub updates or creates the table (in this case the textfile) containing the
'#' hierarchy objects names and infos
If Dir("new_table_people.txt") <> "" Then
    Kill "new_table_people.txt"
End If
varFound = False
Open "table_people.txt" For Input As #1
Open "new_table_people.txt" For Output As #2
Do Until EOF(1)
    Line Input #1, l$
    
    If l$ <> "" Then
        If Left(l$, 1) <> "/" Then
            tmpArray = Split(l$, ";")
            If tmpArray(0) = fID Then
                varFound = True
                Print #2, fID & ";" & tmpArray(1) & ";" & fLeft & ";" & fTop & ";" & fWidth & ";" & fHeight
            Else
                Print #2, l$
            End If
        Else
            Print #2, l$
        End If
    End If
Loop
If varFound = False Then
    Print #2, fID & ";" & fnHch_GetObjectInfo(fID, 2) & ";" & fLeft & ";" & fTop & ";" & fWidth & ";" & fHeight
End If
Close #1
Close #2
Kill "table_people.txt"
Name "new_table_people.txt" As "table_people.txt"
End Sub

Private Sub subHch_DeleteFrameSQL(fID)
'#' this sub deletes a hierarchy object from the database
If Dir("new_table_people.txt") <> "" Then
    Kill "new_table_people.txt"
End If
varFound = False
Open "table_people.txt" For Input As #1
Open "new_table_people.txt" For Output As #2
Do Until EOF(1)
    Line Input #1, l$
    
    If l$ <> "" Then
        If Left(l$, 1) <> "/" Then
            tmpArray = Split(l$, ";")
            If tmpArray(0) <> fID Then
                Print #2, l$
            End If
        Else
            Print #2, l$
        End If
    End If
    
Loop
Close #1
Close #2
Kill "table_people.txt"
Name "new_table_people.txt" As "table_people.txt"
End Sub

Private Sub subHch_UpdateJoinSQL(fID1, fID2)
'#' this sub updates or creates the table (in this case the textfile) containing the
'#' joins informations
If Dir("new_table_join.txt") <> "" Then
    Kill "new_table_join.txt"
End If
varFound = False
Open "table_join.txt" For Input As #1
Open "new_table_join.txt" For Output As #2
Do Until EOF(1)
    Line Input #1, l$
    
    If l$ <> "" Then
        If Left(l$, 1) <> "/" Then
        
            tmpArray = Split(l$, ";")
            If (tmpArray(0) = fID1 And tmpArray(1) = fID2) Or (tmpArray(1) = fID1 And tmpArray(0) = fID2) Then
                varFound = True
            Else
                Print #2, l$
            End If
        Else
            Print #2, l$
        End If
    End If
Loop
If varFound = False Then
    Print #2, fID1 & ";" & fID2
End If
Close #1
Close #2
Kill "table_join.txt"
Name "new_table_join.txt" As "table_join.txt"
End Sub

Private Sub subHch_DeleteJoinSQL(fRelatedTo)
'#' this sub deletes a join from the database
If Dir("new_table_join.txt") <> "" Then
    Kill "new_table_join.txt"
End If
varFound = False
Open "table_join.txt" For Input As #1
Open "new_table_join.txt" For Output As #2
Do Until EOF(1)
    Line Input #1, l$
    
    If l$ <> "" Then
        If Left(l$, 1) <> "/" Then
            tmpArray = Split(l$, ";")
            If (tmpArray(0) <> fRelatedTo) And (tmpArray(1) <> fRelatedTo) Then
                Print #2, l$
            End If
        Else
            Print #2, l$
        End If
    End If
Loop
Close #1
Close #2
Kill "table_join.txt"
Name "new_table_join.txt" As "table_join.txt"
End Sub

Private Sub subHch_LoadObjects()
'#' this is the sub that loads both hierarchy objects and join information into
'#' memory
varPlacedLocs = ""
ReDim varHighest(2)
    varHighest(1) = 1000000
    varHighest(2) = 0
    
varCounter = 0

Open "table_people.txt" For Input As #1
Do Until EOF(1)
    Line Input #1, l$
    
    If l$ <> "" Then
        If Left(l$, 1) <> "/" Then   '#' ignore comment lines
        
            tmpArray = Split(l$, ";")
            varCounter = varCounter + 1
            
            arrayFramesPos(varCounter, 1) = tmpArray(0)           ' ID
            arrayFramesPos(varCounter, 2) = tmpArray(1)           ' Name
            arrayFramesPos(varCounter, 3) = CLng(tmpArray(2))     ' Pos from Left
            arrayFramesPos(varCounter, 4) = CLng(tmpArray(3))     ' Pos from Top
            arrayFramesPos(varCounter, 5) = CLng(tmpArray(4))     ' Width (480)
            arrayFramesPos(varCounter, 6) = CLng(tmpArray(5))     ' Height (480)
            
            If varTop < varHighest(1) Then
                varHighest(1) = varTop
                varHighest(2) = varLeft
            End If
        End If
    End If
Loop
Close #1


varCounter = 0

Open "table_join.txt" For Input As #1
Do Until EOF(1)
    Line Input #1, l$

    If l$ <> "" Then
        If Left(l$, 1) <> "/" Then   '#' ignore comment lines
        
            tmpArray = Split(l$, ";")
            varCounter = varCounter + 1
                
            arrayJoins(varCounter, 1) = tmpArray(0)
            arrayJoins(varCounter, 4) = tmpArray(1)
        End If
    End If
Loop
Close #1

If varHighest(2) <> 0 Then
    PICdraw.Left = (varHighest(2) - (PICbg.Width / 2)) * -1
    PICdraw.Top = (varHighest(1) - 300) * -1
    VScroll1.Value = Abs(PICdraw.Top) / 15
    HScroll1.Value = Abs(PICdraw.Left) / 15
End If

HScroll1.Max = (PICdraw.Width - PICbg.Width) / 15
VScroll1.Max = (PICdraw.Height - PICbg.Height) / 15

'#' joins are calculated and the area is drawed
Call subHch_RecalculateJoins("ALL")
Call subHch_RedrawVisibleArea(0)

End Sub

Private Sub subHch_DeleteJoin(fRelatedTo)
'#' this sub deletes a join from memory

    For xyz = 1 To 5000
        If arrayJoins(xyz, 1) = fRelatedTo Or arrayJoins(xyz, 4) = fRelatedTo Then
            arrayJoins(xyz, 1) = ""
            arrayJoins(xyz, 2) = ""
            arrayJoins(xyz, 3) = ""
            arrayJoins(xyz, 4) = ""
            arrayJoins(xyz, 5) = ""
            arrayJoins(xyz, 6) = ""
        End If
    Next xyz
End Sub

Private Function fnHch_GetMaxLevel()
fnHch_GetMaxLevel = 8
End Function

Private Sub subHch_RedrawObject(ByVal fID)
'#' used to redraw a specific object
For xyz = 1 To 5000
    If arrayFramesPos(xyz, 1) = fID Then
                            
        varCaption = arrayFramesPos(xyz, 2)
        varID = arrayFramesPos(xyz, 1)
        If InStr(varID, "X") <> 0 Then varPic = 1 Else varPic = 0
        
        varDiv = IIf(varZoom = 1, 1, 2)
        PICdraw.PaintPicture Image1(varPic).Picture, arrayFramesPos(xyz, 3) / varZoom, arrayFramesPos(xyz, 4) / varZoom, Image1(varPic).Width / varDiv, Image1(varPic).Height / varDiv
        PICdraw.ForeColor = vbBlack
        PICdraw.DrawMode = 13
        varMode = IIf(varZoom = 1, 1, -1)
        PICdraw.CurrentX = (arrayFramesPos(xyz, 3) / varZoom) - ((PICdraw.TextWidth(varCaption) / 2) / varZoom) + ((Image1(0).Width / 2) * varMode)
        PICdraw.CurrentY = (arrayFramesPos(xyz, 4) / varZoom) + (Image1(0).Height / varDiv)
        PICdraw.Print varCaption
        
        If InStr(varIDList, varID) = 0 Then varIDList = varIDList & varID & "|"
        
        Exit Sub
    End If
Next xyz
End Sub
