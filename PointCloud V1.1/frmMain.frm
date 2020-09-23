VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Point Cloud - By Erkan Sanli"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10710
   DrawWidth       =   2
   FillStyle       =   0  'Solid
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   498
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   714
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      Height          =   6930
      Left            =   0
      ScaleHeight     =   458
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   10
      Top             =   540
      Width           =   3615
      Begin VB.CommandButton cmdRefreshBorder 
         Caption         =   "Refresh"
         Height          =   270
         Left            =   2160
         TabIndex        =   48
         Top             =   5160
         Width           =   975
      End
      Begin VB.CheckBox chkShowHideDot 
         Caption         =   "Show/Hide Dot"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   5880
         Width           =   1455
      End
      Begin VB.CheckBox chkMeshBorder 
         Caption         =   "Show Mesh Border"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   5160
         Width           =   1695
      End
      Begin VB.CheckBox chkBackFace 
         Caption         =   "Show Backface"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   5400
         Width           =   1575
      End
      Begin VB.CheckBox chkBigDot 
         Caption         =   "Big Dot"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   6120
         Width           =   855
      End
      Begin VB.CheckBox chkClip 
         Caption         =   "Clip Far"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   6360
         Width           =   855
      End
      Begin VB.CheckBox chkShowBox 
         Caption         =   "Show Box"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton cmdClipZ 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   2130
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   6360
         Width           =   375
      End
      Begin VB.CommandButton cmdClipZ 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   2760
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   6360
         Width           =   375
      End
      Begin VB.CommandButton cmdClipZ 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   2490
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   6360
         Width           =   255
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mesh Operations"
         Height          =   2895
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   3375
         Begin VB.CommandButton cmdCancelEdit 
            Caption         =   "Cancel Edit"
            Height          =   375
            Left            =   240
            TabIndex        =   46
            Top             =   2280
            Width           =   2895
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Reverse Face"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   45
            Top             =   1800
            Width           =   1450
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Delete Face"
            Height          =   375
            Index           =   5
            Left            =   1680
            TabIndex        =   44
            Top             =   1800
            Width           =   1450
         End
         Begin VB.OptionButton opt 
            Height          =   375
            Index           =   1
            Left            =   2400
            Picture         =   "frmMain.frx":3452
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Select point of rectangular area"
            Top             =   360
            Width           =   375
         End
         Begin VB.OptionButton opt 
            Height          =   375
            Index           =   0
            Left            =   2025
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMain.frx":37DC
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Arrow"
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.OptionButton opt 
            Height          =   375
            Index           =   2
            Left            =   2775
            Picture         =   "frmMain.frx":3F46
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Select point of poligonal area "
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "New Mesh"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Create Mesh"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   30
            Top             =   840
            Width           =   2895
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Reverse Mesh"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   29
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Delete Mesh"
            Height          =   375
            Index           =   3
            Left            =   1680
            TabIndex        =   28
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame fra 
         Caption         =   "Local"
         Height          =   1815
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   3375
         Begin VB.CommandButton cmdParamsM 
            Caption         =   "Go"
            Height          =   300
            Index           =   1
            Left            =   2200
            TabIndex        =   20
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox txtTra 
            Height          =   285
            Index           =   0
            Left            =   480
            TabIndex        =   19
            Text            =   "traX"
            Top             =   480
            Width           =   910
         End
         Begin VB.TextBox txtTra 
            Height          =   285
            Index           =   1
            Left            =   480
            TabIndex        =   18
            Text            =   "traY"
            Top             =   720
            Width           =   910
         End
         Begin VB.TextBox txtTra 
            Height          =   285
            Index           =   2
            Left            =   480
            TabIndex        =   17
            Text            =   "traZ"
            Top             =   960
            Width           =   910
         End
         Begin VB.TextBox txtRot 
            Height          =   285
            Index           =   0
            Left            =   1440
            TabIndex        =   16
            Text            =   "rotX"
            Top             =   480
            Width           =   910
         End
         Begin VB.TextBox txtRot 
            Height          =   285
            Index           =   1
            Left            =   1440
            TabIndex        =   15
            Text            =   "rotY"
            Top             =   720
            Width           =   910
         End
         Begin VB.TextBox txtRot 
            Height          =   285
            Index           =   2
            Left            =   1440
            TabIndex        =   14
            Text            =   "rotZ"
            Top             =   960
            Width           =   910
         End
         Begin VB.TextBox txtSca 
            Height          =   285
            Index           =   0
            Left            =   2400
            TabIndex        =   13
            Text            =   "scaX"
            Top             =   480
            Width           =   910
         End
         Begin VB.CommandButton cmdParamsM 
            Caption         =   "Reset"
            Height          =   300
            Index           =   0
            Left            =   360
            TabIndex        =   12
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "Scale"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   26
            Top             =   240
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "Rotation"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   25
            Top             =   240
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "Translation"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   24
            Top             =   240
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "Z"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   23
            Top             =   1000
            Width           =   300
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "Y"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   22
            Top             =   750
            Width           =   300
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "X"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   300
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   710
      TabIndex        =   1
      Top             =   0
      Width           =   10710
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":42D0
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Open"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdExportOBJ 
         Height          =   495
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":4A3A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Export Wavefront OBJ file"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdImportPoint 
         Height          =   495
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":51A4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Import Points File"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdOpen 
         Height          =   495
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":590E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Open"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNew 
         Height          =   495
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":6078
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "New"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdVStyle 
         Height          =   495
         Index           =   0
         Left            =   2640
         Picture         =   "frmMain.frx":67E2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Dot"
         Top             =   0
         Width           =   500
      End
      Begin VB.CommandButton cmdVStyle 
         Height          =   495
         Index           =   1
         Left            =   3120
         Picture         =   "frmMain.frx":7124
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Wireframe"
         Top             =   0
         Width           =   500
      End
      Begin VB.CommandButton cmdVStyle 
         Height          =   495
         Index           =   3
         Left            =   4080
         Picture         =   "frmMain.frx":7A66
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Smooth"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdVStyle 
         Height          =   495
         Index           =   2
         Left            =   3600
         Picture         =   "frmMain.frx":83A8
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Facet"
         Top             =   0
         Width           =   500
      End
   End
   Begin VB.Timer tmrProcess 
      Interval        =   1
      Left            =   11640
      Top             =   10920
   End
   Begin VB.PictureBox picCanvas 
      BackColor       =   &H8000000C&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00808080&
      Height          =   2040
      Left            =   3720
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   140
      TabIndex        =   0
      Top             =   600
      Width           =   2160
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu tire 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import"
         Begin VB.Menu mnuImportPoints 
            Caption         =   "ASCII Points File"
         End
         Begin VB.Menu mnuImportOBJ 
            Caption         =   "OBJ"
         End
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
         Begin VB.Menu mnuExportOBJ 
            Caption         =   "OBJ"
         End
      End
      Begin VB.Menu tire2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuVisualstyleC 
      Caption         =   "&Visual Style"
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Dot"
         Index           =   0
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Wireframe"
         Index           =   1
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Facets"
         Index           =   2
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Smooth"
         Index           =   3
      End
   End
   Begin VB.Menu cHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopCancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu mnuPopClosePolygon 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub chkShowHideDot_Click()
    ShowHideDot = chkShowHideDot.Value
End Sub

Private Sub cmdRefreshBorder_Click()
    
    Dim idx         As Long

    With Mesh1.Meshs(Mesh1.NumMeshs)
        ReDim .BorderEdges(1)
        For idx = 1 To .NumFaces
            With .Faces(idx)
                Call SearchEdge(.A, .B)
                Call SearchEdge(.B, .C)
                Call SearchEdge(.C, .A)
            End With
        Next
        ReDim Preserve .BorderEdges(UBound(.BorderEdges) - 1)
    End With

End Sub

Private Sub mnuFileNew_Click()
    
    Dots1.NumDot = 0
    Erase Dots1.Dots
    Erase Dots1.Box
    Unsaved = False
    Call Reset
    Call Render(picCanvas.hDC)
    
End Sub

Private Sub mnuFileOpen_Click()
    
    Set cdiLoad = New clsCommonDialog
    Set cfEPJ = New clsFileEPJ
    With cdiLoad
        .Filter = "Point Cloud Job |*.EPJ"
        .InitDir = App.Path & "\JOB"
        .FileName = ""
        .ShowOpen
        DoEvents
        Me.MousePointer = vbHourglass
        If FileExist(.FileName) Then cfEPJ.ReadEPJ .FileName
        Me.MousePointer = vbDefault
        If LoadComplete Then
            Call Init
            tmrProcess.Enabled = True
            cmdEdit(0).Enabled = True
        Else
            tmrProcess.Enabled = False
            cmdEdit(0).Enabled = False
        End If
    End With
    Set cdiLoad = Nothing
    Set cfEPJ = Nothing
    Unsaved = False
    Call mnuVisualStyle_Click(Facet)
    frmMain.picCanvas.SetFocus

End Sub

Private Sub mnuFileSave_Click()
    
    Set cdiLoad = New clsCommonDialog
    Set cfEPJ = New clsFileEPJ
    With cdiLoad
        .Filter = "Pointcloud Job |*.epj"
        .InitDir = App.Path & "\JOB"
        .DefaultExt = "*.epj"
        .ShowSave
        DoEvents
        Me.MousePointer = vbHourglass
        If Len(.FileName) Then cfEPJ.WriteEPJ .FileName
        Me.MousePointer = vbDefault
    End With
    Set cdiLoad = Nothing
    Set cfEPJ = Nothing
    Unsaved = False
    frmMain.picCanvas.SetFocus

End Sub

Private Sub mnuHelp_Click()

    Call ShellExecute(hWnd, "open", App.Path & "\help.pdf", vbNullString, vbNullString, 0&)

End Sub

Private Sub mnuImportPoints_Click()
    
    Set cdiLoad = New clsCommonDialog
    Set cfPoints = New clsFilePOINTS
    tmrProcess.Enabled = False
    frmImportPoint.Show vbModal
    If LoadComplete Then
        Call Init
        tmrProcess.Enabled = True
        cmdEdit(0).Enabled = True
    Else
        tmrProcess.Enabled = False
        cmdEdit(0).Enabled = False
    End If
    Set cfPoints = Nothing
    Set cdiLoad = Nothing
    Unsaved = False
    frmMain.picCanvas.SetFocus

End Sub

Private Sub mnuImportOBJ_Click()

    Set cdiLoad = New clsCommonDialog
    Set cfOBJ = New clsFileOBJ
    With cdiLoad
        .Filter = "Wavefront Object File |*.obj"
        .InitDir = App.Path & "\Export"
        .DefaultExt = "*.obj"
        .FileName = ""
        .ShowOpen
        DoEvents
        Me.MousePointer = vbHourglass
        cfOBJ.ReadOBJ .FileName
        Me.MousePointer = vbDefault
    End With
    If LoadComplete Then
        Call Init
        tmrProcess.Enabled = True
        cmdEdit(0).Enabled = True
    Else
        tmrProcess.Enabled = False
        cmdEdit(0).Enabled = False
    End If
    Set cdiLoad = Nothing
    Set cfOBJ = Nothing
    Unsaved = False
    frmMain.picCanvas.SetFocus

End Sub

Private Sub mnuExportOBJ_Click()
    
    Set cdiLoad = New clsCommonDialog
    Set cfOBJ = New clsFileOBJ
    With cdiLoad
        .Filter = "Wavefront Object File |*.obj"
        .InitDir = App.Path & "\Export"
        .DefaultExt = "*.obj"
        .FileName = ""
        .ShowSave
        DoEvents
        Me.MousePointer = vbHourglass
        cfOBJ.WriteOBJ .FileName
        Me.MousePointer = vbDefault
    End With
    Set cdiLoad = Nothing
    Set cfOBJ = Nothing
    frmMain.picCanvas.SetFocus

End Sub

Private Sub mnuFileExit_Click()
    
    Unload Me
    
End Sub

Private Sub mnuPopCancel_Click()
    
    EditMF = NewMesh

End Sub

Private Sub mnuPopClosePolygon_Click()
    
    StartPolygon = False
    SelectOp = True
    Call SelectedPoints

End Sub

Public Sub mnuVisualStyle_Click(Index As Integer)
    
    Dim idx As Integer

    VStyle = Index
    For idx = 0 To mnuVisualStyle.Count - 1
        mnuVisualStyle(idx).Checked = IIf(idx = Index, True, False)
    Next
    picCanvas.SetFocus
    tmrProcess.Enabled = True
                                                 
End Sub

Private Sub mnuAbout_Click()
    
    frmAbout.Show vbModal

End Sub

Private Sub cmdExportOBJ_Click()
    
    Call mnuExportOBJ_Click

End Sub

Private Sub cmdNew_Click()

    Call mnuFileNew_Click

End Sub

Private Sub cmdOpen_Click()
    
    Call mnuFileOpen_Click

End Sub

Private Sub cmdSave_Click()
    
    Call mnuFileSave_Click

End Sub

Private Sub chkMeshBorder_Click()

    ShowMeshBorder = chkMeshBorder.Value

End Sub

Private Sub chkShowBox_Click()
    
    ShowBox = chkShowBox.Value
    frmMain.picCanvas.SetFocus

End Sub

Private Sub cmdCancelEdit_Click()
    
    EditMF = NewMesh

End Sub

Public Sub cmdClipZ_Click(Index As Integer)
    
    Select Case Index
        Case 0
            Dots1.Center.Vector.Z = Dots1.Center.Vector.Z + Dots1.ClpZ
        Case 1
            Dots1.Center.Vector.Z = Dots1.Center.Vector.Z - Dots1.ClpZ
        Case 2
            Dots1.Center.Vector.Z = 0
    End Select
    frmMain.picCanvas.SetFocus

End Sub

Private Sub cmdImportPoint_Click()
    
    Call mnuImportPoints_Click

End Sub

Private Sub Form_Load()
    
    Set Geometry = New clsPolygonal
    Call MouseInit
    Call Init
    
End Sub

Private Sub Form_Resize()
    
    On Error GoTo err:
    picCanvas.Move Picture2.ScaleWidth, Picture1.ScaleHeight
    picCanvas.Width = Me.ScaleWidth - Picture2.ScaleWidth
    picCanvas.Height = Me.ScaleHeight - Picture1.ScaleHeight
    cWidth = picCanvas.Width
    cHeight = picCanvas.Height
    HalfWidth = cWidth / 2
    HalfHeight = cHeight / 2
    Call CanBuffer.Create(cWidth, cHeight)
    Call BackBuffer.Create(cWidth, cHeight)
    Call DrawGradientRectangle
err:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    DoEvents
    tmrProcess.Enabled = False
    Set CanBuffer = Nothing
    Set BackBuffer = Nothing
    Set Geometry = Nothing
    End

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim RetVal As Long
    
    If Unsaved Then
        RetVal = MsgBox("Save changes to Job file", vbYesNoCancel)
        Select Case RetVal
            Case vbYes: mnuFileSave_Click: m_blnWheelPresent = False
            Case vbNo: mnuFileExit_Click: m_blnWheelPresent = False
            Case vbCancel: Cancel = 1
        End Select
    Else
        RetVal = MsgBox("Do you really exit?", vbYesNo)
        Select Case RetVal
            Case vbYes: m_blnWheelPresent = False
            Case vbNo: Cancel = 1
        End Select
    End If

End Sub

Private Sub cmdVStyle_Click(Index As Integer)
    
    mnuVisualStyle_Click (Index)

End Sub

Private Sub picCanvas_DblClick()

    Position.Sca = 1
    InvScl = 1
    Position.Tra = VectorSet(0, 0, 0)
    
End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Select Case Button
        
        Case vbLeftButton
            Select Case SelType
                Case Rectangular
                    Rect1.X = X
                    Rect1.Y = Y
                    Rect2.X = X
                    Rect2.Y = Y
                    SelectOp = True
                Case Polygonal
                    If StartPolygon Then
                        Geometry.AddVertex CLng(X), CLng(Y)
                    Else
                        StartPolygon = True
                        SelectOp = False
                        Geometry.ClearDots
                        Geometry.AddVertex CLng(X), CLng(Y)
                    End If
                Case Else
                    SelectOp = False
            End Select
            
            If EditMF > DeleteMesh Then
                If Button = vbLeftButton Then
                    Select Case EditMF
                        Case DeleteFace
                            Call DeleteSelFace(SelectedMeshIndex, SelectedFaceIndex)
                        Case ReverseFace
                            Call ReverseBackFace(SelectedMeshIndex, SelectedFaceIndex)
                    End Select
                End If
            End If
        
        Case vbRightButton
            PopupMenu mnuPop

    End Select
        
End Sub

Private Sub tmrProcess_Timer()
    
    If LoadComplete Then Call RenderProcess
    
End Sub

Private Sub chkBackFace_Click()

    ShowBackFace = CBool(chkBackFace.Value)
 
End Sub

Private Sub chkBigDot_Click()
    
    BigDot = chkBigDot.Value

End Sub

Private Sub chkClip_Click()
    
    ClipFar = chkClip.Value
    cmdClipZ(0).Enabled = ClipFar
    cmdClipZ(1).Enabled = ClipFar
    cmdClipZ(2).Enabled = ClipFar
    frmMain.picCanvas.SetFocus

End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Select Case Button
    
        Case vbLeftButton
            Select Case SelType
            
                Case NoSelect
'                    If EditMF <> DeleteFace Then
                        Position.Rot.X = Position.Rot.X - (LastY - Y) * 0.8
                        Position.Rot.Y = Position.Rot.Y - (LastX - X) * 0.8
'                    End If
                Case Rectangular
                    Rect2.X = X
                    Rect2.Y = Y
                    Call SelectedPoints
            End Select
            
        Case vbMiddleButton
            If SelType = NoSelect Then
                Position.Tra.X = (Position.Tra.X - (LastX - X) * 1.2)
                Position.Tra.Y = (Position.Tra.Y - (LastY - Y) * 1.2)
            End If
    End Select
    LastX = X
    LastY = Y

End Sub


Private Sub RenderProcess()
    
    On Error Resume Next
    
    Dim idx   As Long
    Dim idxMesh As Integer
    Dim idxFace As Long
    
    DoEvents
    Call RefreshEvents(frmMain.picCanvas.hWnd)
    
' Dots ___________________________________
    With Dots1
        For idx = 1 To .NumDot                                          'Dots
            Call NewDotPos(.Dots(idx))
        Next idx
        If .NumSelDot > 0 Then
            For idx = 1 To .NumSelDot                                   'Selected dots
                Call NewDotPos(.SelDots(idx))
            Next idx
        End If
        For idx = 1 To 8                                                'Box
            Call NewDotPos(.Box(idx))
        Next idx
        Call NewDotPos(.Center)                                         'Center
    End With
    
'Meshs____________________________________
    With Mesh1
        If .NumMeshs > 0 Then
            For idxMesh = 1 To .NumMeshs                                'Mesh Normals
                For idx = 1 To .Meshs(idxMesh).NumFaces
                    .Meshs(idxMesh).NormalsT(idx) = MatrixMultiplyVector(matOutput, .Meshs(idxMesh).Normals(idx))
                Next
            Next
        End If
    End With
            
    Call Render(picCanvas.hDC)

    With Position
        txtTra(0).Text = .Tra.X
        txtTra(1).Text = .Tra.Y
        txtTra(2).Text = .Tra.Z
        txtRot(0).Text = .Rot.X
        txtRot(1).Text = .Rot.Y
        txtRot(2).Text = .Rot.Z
        txtSca(0).Text = .Sca
    End With

End Sub

Private Sub NewDotPos(NewDot As DOT)
    
    With NewDot
        .VectorT = MatrixMultiplyVector(matOutput, .Vector)
        .Screen.X = .VectorT.X + HalfWidth
        .Screen.Y = .VectorT.Y + HalfHeight
    End With

End Sub

Private Sub SelectedPoints()

    Dim X1 As Single
    Dim Y1 As Single
    Dim idx As Long
    Dim MaxP As POINTAPI
    Dim MinP As POINTAPI

        Select Case SelType
            Case Rectangular
                MinP = Rect1
                MaxP = Rect2
                If MinP.X > Rect2.X Then MinP.X = Rect2.X
                If MinP.Y > Rect2.Y Then MinP.Y = Rect2.Y
                If MaxP.X < Rect1.X Then MaxP.X = Rect1.X
                If MaxP.Y < Rect1.Y Then MaxP.Y = Rect1.Y
                For idx = 1 To Dots1.NumDot
                    With Dots1.Dots(idx)
                        If .Visible Then
                            X1 = .Screen.X
                            Y1 = .Screen.Y
                            If X1 > MinP.X And X1 < MaxP.X And Y1 > MinP.Y And Y1 < MaxP.Y Then
                                .Selected = True
                            Else
                                .Selected = False
                            End If
                        End If
                    End With
                Next
            Case Polygonal
                For idx = 1 To Dots1.NumDot
                    With Dots1.Dots(idx)
                        If .Visible Then
                            X1 = .Screen.X
                            Y1 = .Screen.Y
                            If Geometry.VertexCount > 2 Then
                                .Selected = Geometry.IsInPolygon(X1, Y1)
                            Else
                                .Selected = False
                            End If
                        End If
                    End With
                Next
        End Select

End Sub

Private Sub Init()

    Dim idx As Byte
    
    InvScl = 1
    EditMF = NewMesh
    chkBackFace.Value = vbChecked
    chkBigDot.Value = vbChecked
    chkShowHideDot = vbChecked
    cmdClipZ(0).Enabled = False
    cmdClipZ(1).Enabled = False
    cmdClipZ(2).Enabled = False
    opt(0).Enabled = False
    opt(1).Enabled = False
    opt(2).Enabled = False
    cmdCancelEdit.Enabled = False
    For idx = 0 To 5
        cmdEdit(idx).Enabled = False
    Next
    With Color
        .rgbBack1 = ColorSet(100, 110, 170)
        .rgbBack2 = ColorSet(10, 20, 30)
        .lObjDots = RGB(0, 255, 255) 'BGR
        .lSelDots = RGB(0, 0, 255) ' BGR
        .lWireframe = RGB(250, 250, 250) ' BGR
        .rgbFaces = ColorSet(250, 250, 250)
        .rgbFacesPen = ColorSet(240, 240, 240)
        .rgbBackFace = ColorSet(255, 0, 0)
        .rgbSelFace = ColorSet(0, 0, 255)
        .rgbSelBackFace = ColorSet(255, 255, 0)
        .lBox = RGB(100, 50, 50) ' BGR
        .lSelGeo = RGB(255, 0, 0) ' BGR
    End With
    Call ResetCameraParameters
    Call ResetLightParameters

End Sub

Private Sub cmdParamsM_Click(Index As Integer)
    
    Select Case Index
    Case 0
        Call ResetMeshParameters
    Case 1
        With Position
            .Tra = VectorSet(VerifyText(txtTra(0)), VerifyText(txtTra(1)), VerifyText(txtTra(2)))
            .Rot = VectorSet(VerifyText(txtRot(0)), VerifyText(txtRot(1)), VerifyText(txtRot(2)))
            .Sca = VerifyText(txtSca(0))
        End With
    End Select
    frmMain.picCanvas.SetFocus
    frmMain.tmrProcess.Enabled = True

End Sub

Private Sub txtRot_GotFocus(Index As Integer)
    frmMain.tmrProcess.Enabled = False
End Sub
Private Sub txtRot_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then cmdParamsM_Click (1) ' 13=Enter
End Sub
Private Sub txtSca_GotFocus(Index As Integer)
    frmMain.tmrProcess.Enabled = False
End Sub
Private Sub txtSca_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then cmdParamsM_Click (1) ' 13=Enter
End Sub
Private Sub txtTra_GotFocus(Index As Integer)
    frmMain.tmrProcess.Enabled = False
End Sub
Private Sub txtTra_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then cmdParamsM_Click (1)  ' 13=Enter
End Sub

Private Sub opt_Click(Index As Integer)
    
    SelType = Index
    Select Case SelType
        Case Rectangular
            Rect1.X = 0
            Rect1.Y = 0
            Rect2.X = 0
            Rect2.Y = 0
        Case Polygonal
            Geometry.ClearDots
    End Select
    frmMain.picCanvas.SetFocus

End Sub

Private Sub Seperate()
    
    Dim idx As Long
    
    With Dots1
        .NumSelDot = 1
        ReDim .SelDots(.NumSelDot)
        For idx = 1 To .NumDot
            If .Dots(idx).Selected Then
                If .Dots(idx).Visible Then
                    If ClipFar Then
                        If .Center.VectorT.Z < .Dots(idx).VectorT.Z Then
                            ReDim Preserve .SelDots(.NumSelDot)
                            .SelDots(.NumSelDot) = .Dots(idx)
                            .SelDots(.NumSelDot).Index = idx
                            .NumSelDot = .NumSelDot + 1
                        End If
                    Else
                        ReDim Preserve .SelDots(.NumSelDot)
                        .SelDots(.NumSelDot) = .Dots(idx)
                        .SelDots(.NumSelDot).Index = idx
                        .NumSelDot = .NumSelDot + 1
                    End If
                End If
            End If
        Next
        .NumSelDot = UBound(.SelDots)
    End With
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    
    Dim idx As Long
    
    EditMF = Index
    Select Case EditMF
        
        Case NewMesh
            opt(2).Value = True
            opt(0).Enabled = True
            opt(1).Enabled = True
            opt(2).Enabled = True
            cmdEdit(0).Enabled = False
            cmdEdit(1).Enabled = True
            cmdEdit(2).Enabled = False
            cmdEdit(3).Enabled = False
            cmdEdit(4).Enabled = False
            cmdEdit(5).Enabled = False
            cmdCancelEdit.Enabled = False

        Case CreateMesh
            Call Seperate
            Call Triangulate
            Call mnuVisualStyle_Click(Facet)
            opt(0).Value = True
            opt(0).Enabled = False
            opt(1).Enabled = False
            opt(2).Enabled = False
            cmdEdit(0).Enabled = True
            cmdEdit(1).Enabled = False
            cmdEdit(2).Enabled = True
            cmdEdit(3).Enabled = True
            cmdEdit(4).Enabled = True
            cmdEdit(5).Enabled = True
            cmdCancelEdit.Enabled = False
           
        Case ReverseMesh
            If Mesh1.NumMeshs > 0 Then
                For idx = 1 To Mesh1.Meshs(Mesh1.NumMeshs).NumFaces
                    Call ReverseBackFace(Mesh1.NumMeshs, idx)
                Next
            End If
            cmdCancelEdit.Enabled = True

        Case DeleteMesh
        
            If Mesh1.NumMeshs > 1 Then
                Mesh1.NumMeshs = Mesh1.NumMeshs - 1
                ReDim Preserve Mesh1.Meshs(Mesh1.NumMeshs)
            Else
                Call Reset
            End If
            cmdCancelEdit.Enabled = True
            
        Case ReverseFace, DeleteFace
'            Action = False
            cmdCancelEdit.Enabled = True
    End Select
    frmMain.picCanvas.SetFocus

End Sub

Private Sub ReverseBackFace(idxMesh As Integer, idxFace As Long)
    
    Dim tmpLong As Long

    With Mesh1.Meshs(idxMesh).Faces(idxFace)
        tmpLong = .A
        .A = .B
        .B = tmpLong
        Mesh1.Meshs(idxMesh).Normals(idxFace) = _
            CalculateNormal(Dots1.Dots(.A).Vector, Dots1.Dots(.B).Vector, Dots1.Dots(.C).Vector)
    End With
    
End Sub

Private Sub DeleteSelFace(idxMesh As Integer, idxFace As Long)
    
    Dim idx As Long
    Dim idx2  As Long
    
    With Mesh1.Meshs(idxMesh)
        For idx = 1 To .NumFaces
            If idx <> idxFace Then
             idx2 = idx2 + 1
             .Faces(idx2) = .Faces(idx)
             .Normals(idx2) = .Normals(idx)
            End If
        Next
        .NumFaces = .NumFaces - 1
        If .NumFaces > 0 Then
            ReDim Preserve .Faces(.NumFaces)
            ReDim Preserve .Normals(.NumFaces)
        Else
            Erase .BorderEdges
            Erase .Normals
        End If
    End With
    
End Sub

Private Function VerifyText(txt As TextBox) As Single

    If IsNumeric(txt.Text) Then
        VerifyText = CSng(txt.Text)
    Else
        VerifyText = 0
    End If

End Function

Private Sub Reset()

    Mesh1.NumMeshs = 0
    Erase Mesh1.Meshs
    Erase Mesh1.FaceV
    opt(0).Value = True
    opt(0).Enabled = False
    opt(1).Enabled = False
    opt(2).Enabled = False
    cmdEdit(0).Enabled = True
    cmdEdit(1).Enabled = False
    cmdEdit(2).Enabled = False
    cmdEdit(3).Enabled = False
    cmdEdit(4).Enabled = False
    cmdEdit(5).Enabled = False

End Sub


