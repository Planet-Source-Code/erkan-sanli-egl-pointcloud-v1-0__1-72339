VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Point Cloud"
   ClientHeight    =   3825
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6105
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmAbout.frx":3452
   Picture         =   "frmAbout.frx":411C
   ScaleHeight     =   2640.083
   ScaleMode       =   0  'User
   ScaleWidth      =   5732.911
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4440
      TabIndex        =   0
      Top             =   3240
      Width           =   1260
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label lblLink 
      Caption         =   "Other Submissions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   2235
      Left            =   120
      Picture         =   "frmAbout.frx":816C
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1785
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Point Cloud"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2160
      TabIndex        =   1
      Top             =   1560
      Width           =   3645
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
    lblTitle.Caption = App.Title & " V " & App.Major & "." & App.Minor & "." & App.Revision
    Text1.Text = "Author  : Erkan Þanlý" & vbNewLine & _
                 "Country : Izmir / Turkey   " & vbNewLine & _
                 "e-mail  : erkansanli_70@hotmail.com "
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 0
End Sub

Private Sub lblLink_Click()

    Call ShellExecute(hWnd, "open", _
                     "http://www.pscode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=22111880271&strAuthorName=Erkan%20Sanli&txtMaxNumberOfEntriesPerPage=25", _
                     vbNullString, vbNullString, 0&)

End Sub

Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 99
End Sub

