VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Bytes to your program made by sharon elharar"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   -120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "AddKb"
      Filter          =   "Alle Dateien | *.*"
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "1000"
      Top             =   840
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   135
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Byte"
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CD1.ShowOpen
Text1.Text = CD1.FileName
End Sub

Private Sub Command2_Click()
If Text1.Text <> "" And Text2.Text > 0 Then
fsiz = ShowFileSize(Text1.Text)
PB1.Value = 0
PB1.Max = Text2.Text
PB1.Visible = True
Open Text1.Text For Binary As #1
For a = 1 To Text2.Text
Put #1, fsiz - 1 + a, 0
PB1.Value = a
Next
Close
End If
PB1.Visible = False
PB1.Value = 0
End Sub
Function ShowFileSize(file)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(file)
    ShowFileSize = f.Size
    's = UCase(f.Name) & " uses " & f.Size & " bytes."
    'MsgBox s, 0, "Folder Size Info"
End Function
'94208

Private Sub Form_Load()
Text1.Text = App.Path & "\"
End Sub
