VERSION 5.00
Begin VB.Form frmImportPoint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Point"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   Icon            =   "frmImportPoint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Parse"
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Width           =   7095
      Begin VB.ComboBox cmbFormat 
         Height          =   315
         Left            =   1200
         TabIndex        =   17
         Text            =   "cmbFormat"
         Top             =   840
         Width           =   5655
      End
      Begin VB.CheckBox chkDelimiters 
         Caption         =   "Other"
         Height          =   375
         Index           =   4
         Left            =   5640
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkDelimiters 
         Caption         =   "Space"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkDelimiters 
         Caption         =   "Comma"
         Height          =   375
         Index           =   2
         Left            =   3480
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkDelimiters 
         Caption         =   "Semicolon"
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkDelimiters 
         Caption         =   "Tab"
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtDelimiter 
         Height          =   285
         Left            =   6480
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Format :"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Delimiters :"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Output"
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   7095
      Begin VB.ListBox lstOutput 
         Height          =   1230
         ItemData        =   "frmImportPoint.frx":000C
         Left            =   240
         List            =   "frmImportPoint.frx":000E
         TabIndex        =   9
         Top             =   360
         Width           =   6615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.ListBox lstInput 
         Height          =   1230
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   6615
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   315
         Left            =   5880
         TabIndex        =   3
         Top             =   600
         Width           =   1000
      End
      Begin VB.Label Label3 
         Caption         =   "Preview (first 100 lines)"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblFilename 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblFilename"
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   "Point file name"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmImportPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FileName As String

Private Sub Form_Load()
    
    With cmbFormat
        .AddItem "X , Y , Z"
        .AddItem "Number , X , Y , Z"
        .AddItem "X , Y , Z , Description"
        .AddItem "Number , X , Y , Z , Description"
        .Text = .List(0)
    End With
    txtDelimiter.Visible = False
    lblFilename.Caption = ""
    chkDelimiters(eComma).Value = vbChecked
    chkDelimiters(eSpace).Value = vbChecked
        
End Sub

Private Sub cmdBrowse_Click()
    
    With cdiLoad
        .Filter = "ASCII Points File (*.*)|*.*"
        .InitDir = App.Path & "\Import"
        .FileName = ""
        .ShowOpen
        If FileExist(.FileName) Then
            FileName = .FileName
            lblFilename.Caption = GetFileNameEx(FileName)
            Call cfPoints.InputList(FileName, lstInput)
            Call cfPoints.RefreshOutputList(lstOutput)
        Else
            lblFilename.Caption = "Cancelled"
            lstInput.Clear
            lstOutput.Clear
        End If
    End With
    
End Sub

Private Sub chkDelimiters_Click(Index As Integer)
 
    With cfPoints
        Select Case Index
            Case eTab:          delim.tTab = Not delim.tTab
            Case eSemicolon:    delim.tSemicolon = Not delim.tSemicolon
            Case eComma:        delim.tComma = Not delim.tComma
            Case eSpace:        delim.tSpace = Not delim.tSpace
            Case eOther:        delim.tOther = Not delim.tOther
        End Select
        If delim.tOther Then
            txtDelimiter.Visible = True
            txtDelimiter.SetFocus
        Else
            txtDelimiter.Visible = False
        End If
        Call .RefreshOutputList(lstOutput)
    End With
    
End Sub

Private Sub cmdOK_Click()

    If lblFilename.Caption <> "Cancelled" Then
        DoEvents
        frmMain.MousePointer = vbHourglass
        Call cfPoints.InputAll(FileName)
        frmMain.MousePointer = vbDefault
    End If
    Call cmdCancel_Click
    
End Sub

Private Sub cmdCancel_Click()
    
    Me.Hide
    frmMain.tmrProcess = True
    
End Sub

Private Sub lblFilename_Click()
    
    cmdBrowse_Click

End Sub

Private Sub txtDelimiter_Change()

    delim.tDelimChar = txtDelimiter.Text
    Call cfPoints.RefreshOutputList(lstOutput)

End Sub

Private Sub cmbFormat_Click()
    
    delim.tFormat = cmbFormat.ListIndex
    Call cfPoints.RefreshOutputList(lstOutput)

End Sub

