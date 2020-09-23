VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   2775
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   3255
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAdd 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.ListBox lstTypes 
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   10
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   9
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "to the list:"
      Height          =   195
      Index           =   2
      Left            =   1920
      TabIndex        =   14
      Top             =   720
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add extension"
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   13
      Top             =   480
      Width           =   1005
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3120
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search for Files with these extensions:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   2700
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
If txtAdd.Text = "" Then Exit Sub
lstTypes.AddItem UCase(txtAdd.Text)
txtAdd.Text = ""
End Sub

Private Sub cmdRemove_Click()
On Error Resume Next
lstTypes.RemoveItem lstTypes.ListIndex
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If lstTypes.ListCount = 0 Then
    MsgBox "You need to search for at least one file type.", vbInformation, "Options"
Else
    frmMain.FileTypeList = ""
    x = 0
    Do Until x >= lstTypes.ListCount
        frmMain.FileTypeList = frmMain.FileTypeList & lstTypes.List(x) & ";"
        x = x + 1
    Loop
    WriteINI "Main", "Extensions", frmMain.FileTypeList, App.Path & "\settings.ini"
    frmMain.LoadMeForm
    Unload Me
End If
End Sub

Private Sub Form_Load()

listsplitup = Split(frmMain.FileTypeList, ";")
x = 0
Do Until x >= UBound(listsplitup)
    lstTypes.AddItem listsplitup(x)
    x = x + 1
Loop

End Sub
