VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FileFind"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6135
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.log"
      DialogTitle     =   "Save Log"
      Filter          =   "Log File|*.log"
   End
   Begin VB.TextBox txtOutput 
      Height          =   375
      Left            =   4200
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Log"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   5880
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0794
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E38
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":118A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9340
      _Version        =   393217
      Indentation     =   212
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5895
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FileTypeList As String

Dim FileCount As Dictionary
Dim FilePaths As Dictionary
Dim FileSizes As Dictionary

Dim AmSearching As Boolean
Dim addattrs As Boolean

Dim ElNumero As Integer


Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdSave_Click()
On Error GoTo finishup

listsplitup = Split(FileTypeList, ";")
x = 0
Do Until x >= UBound(listsplitup)
    txtOutput = txtOutput & "[" & listsplitup(x) & " Files - Count: " & FileCount(listsplitup(x)) & " - Size: " & Round(FileSizes(listsplitup(x)), 2) & " kb]" & vbCrLf
    txtOutput = txtOutput & FilePaths(listsplitup(x)) & vbCrLf
    x = x + 1
Loop

CommonDialog1.ShowOpen

Open CommonDialog1.filename For Output As #1
    Print #1, txtOutput
Close #1

MsgBox "Log Saved to:" & vbCrLf & "C:\FileFind.log", vbInformation, "Log Saved"

finishup:
End Sub

Private Sub cmdSearch_Click()
If cmdSearch.Caption = "Cancel Search" Then
    AmSearching = False
    MsgBox "Search Cancelled.", vbInformation, "Search Complete"
Else
    ReDim sarray(0) As String
        
    AmSearching = True
    
    cmdSave.Enabled = False
    cmdSearch.Caption = "Cancel Search"
    TreeView1.Enabled = False
    mnuOptions.Enabled = False
    
    LoadMeForm
    
    Call DirWalk(Dir1.Path, sarray)
    
    mnuOptions.Enabled = True
    TreeView1.Enabled = True
    cmdSearch.Caption = "Search"
    cmdSave.Enabled = True
    If AmSearching = True Then MsgBox "Search Complete.", vbInformation, "Search Complete"
    AmSearching = False
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Drive1_Change()
On Error GoTo endzor
TreeView1.Nodes.Item("hdd").Text = Drive1.Drive
Dir1.Path = Left(Drive1.Drive, 2) & "\"
File1.Path = Dir1.Path
LoadMeForm
Exit Sub
endzor:
Select Case Err.Number
    Case 68
        MsgBox Err.Description
        Drive1.Drive = "C:"
        Dir1.Path = "C:\"
        LoadMeForm
End Select
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Function CheckTheDir()

x = 0
Do Until x >= File1.ListCount
        
    right4 = UCase(Right(File1.List(x), 4))
    right5 = UCase(Right(File1.List(x), 5))
    
    If Not Right(File1.Path, 1) = "\" Then
        thefullpath = File1.Path & "\" ' & File1.List(c)
    Else
        thefullpath = File1.Path
    End If
    
    addattrs = False
    
    listsplitup = Split(FileTypeList, ";")
    lol = 0
    Do Until lol >= UBound(listsplitup)
        checkright = UCase(Right(File1.List(x), Len(listsplitup(lol))))
        If checkright = listsplitup(lol) Then
            ElNumero = ElNumero + 1
            addattrs = True
            FileCount(listsplitup(lol)) = FileCount(listsplitup(lol)) + 1
            FileSizes(listsplitup(lol)) = FileSizes(listsplitup(lol)) + GetFileSize2(thefullpath & File1.List(x))
            FilePaths(listsplitup(lol)) = FilePaths(listsplitup(lol)) & thefullpath & File1.List(x) & vbCrLf
            TreeView1.Nodes.Add listsplitup(lol), tvwChild, "X" & ElNumero, File1.List(x), 5
            TreeView1.Nodes.Item(listsplitup(lol)).Text = listsplitup(lol) & " Files (Count: " & FileCount(listsplitup(lol)) & ") (Size: " & Round(FileSizes(listsplitup(lol)), 2) & " kb)"
            GoTo continueon
        End If
        lol = lol + 1
    Loop
    
continueon:
    
    If addattrs = True Then
        TreeView1.Nodes.Add "X" & ElNumero, tvwChild, "X" & ElNumero & "_Path", "Path: " & thefullpath, 2
        TreeView1.Nodes.Add "X" & ElNumero, tvwChild, "X" & ElNumero & "_Size", "Size: " & GetFileSize(thefullpath & File1.List(x)), 2
    End If
    
    x = x + 1
Loop

End Function

Private Sub Form_Load()
Set FileCount = New Dictionary
Set FilePaths = New Dictionary
Set FileSizes = New Dictionary

FileTypeList = ReadINI("Main", "Extensions", App.Path & "\settings.ini")

If FileTypeList = "" Then FileTypeList = "PPS;MPE;MPEG;AVI;MOV;JPG;JPEG;BMP;ZIP;"

Dir1.Path = Left(Drive1.Drive, 2) & "\"
LoadMeForm
End Sub

Function LoadMeForm()
Dir1.Path = Left(Drive1.Drive, 2) & "\"
File1.Path = Dir1.Path
TreeView1.Nodes.Clear
TreeView1.Nodes.Add , , "hdd", Drive1.Drive, 4
TreeView1.Nodes.Item("hdd").Expanded = True

listsplitup = Split(FileTypeList, ";")
x = 0
Do Until x >= UBound(listsplitup)
    FileCount(listsplitup(x)) = 0
    FilePaths(listsplitup(x)) = ""
    FileSizes(listsplitup(x)) = 0
    TreeView1.Nodes.Add "hdd", tvwChild, listsplitup(x), listsplitup(x) & " Files (Count: 0) (Size: 0 kb)", 1
    x = x + 1
Loop

ElNumero = 0
End Function

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Public Function GetFileSize(filename) As String
    On Error GoTo Gfserror
    Dim TempStr As String
    TempStr = FileLen(filename)

    If TempStr >= "1024" Then
        'KB
    TempStr = Round(CCur(TempStr / 1024), 2) & "KB"
    Else
        If TempStr >= "1048576" Then
            'MB
            TempStr = Round(CCur(TempStr / (1024 * 1024)), 2) & "KB"
        Else
            TempStr = Round(CCur(TempStr), 2) & "B"
        End If
    End If
    GetFileSize = TempStr
    Exit Function
Gfserror:
    GetFileSize = "0B"
    Resume
End Function

Public Function GetFileSize2(filename) As Double
    On Error GoTo Gfserror
    Dim TempStr As String
    TempStr = FileLen(filename)
    TempStr = Round(CCur(TempStr / 1024), 2)
    GetFileSize2 = TempStr
    Exit Function
Gfserror:
    GetFileSize2 = 0
    Resume
End Function


Sub DirWalk(ByVal CurDir As String, SFound() As String)
    If AmSearching = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    On Error Resume Next
    Dim i As Integer
    Dim sCurPath As String
    Dim sFile As String
    Dim ii As Integer
    Dim ifiles As Integer
    Dim ilen As Integer

    If Right$(CurDir, 1) <> "\" Then
        Dir1.Path = CurDir & "\"
    Else
        Dir1.Path = CurDir
    End If

    For i = 0 To Dir1.ListCount
        If AmSearching = False Then Exit Sub
        If Dir1.List(i) <> "" Then
            If AmSearching = False Then Exit Sub
            DoEvents
            Call DirWalk(Dir1.List(i), SFound())
        Else
            If Right$(Dir1.Path, 1) = "\" Then
                sCurPath = Left$(Dir1.Path, Len(Dir1.Path) - 1)
            Else
                sCurPath = Dir1.Path
            End If
            File1.Path = sCurPath
            
            If AmSearching = False Then Exit Sub
            CheckTheDir

            ilen = Len(Dir1.Path)
            Do While Mid(Dir1.Path, ilen, 1) <> "\"
                ilen = ilen - 1
            Loop
            Dir1.Path = Mid(Dir1.Path, 1, ilen)
        End If
    Next i

    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuOptions_Click()
frmOptions.Show vbModal
End Sub
Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
If Node.Image = 3 Then
    Node.Image = 1
End If
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
If Node.Image = 1 Then
    Node.Image = 3
End If
End Sub
