VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPickApps 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Directory Check"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form2"
   ScaleHeight     =   2145
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   1800
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1680
      Width           =   2355
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6720
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   60
      Picture         =   "frmPickApps.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   900
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available Users:"
      Height          =   195
      Left            =   420
      TabIndex        =   5
      Top             =   1740
      Width           =   1140
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00404040&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   765
      Left            =   -60
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   420
      TabIndex        =   3
      Top             =   1260
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPickApps.frx":11C9
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   420
      TabIndex        =   2
      Top             =   60
      Width           =   5895
   End
   Begin VB.Label lblN 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   420
      TabIndex        =   1
      Top             =   960
      Width           =   6315
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   -120
      Top             =   0
      Width           =   7395
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   -60
      Top             =   60
      Width           =   7335
   End
End
Attribute VB_Name = "frmPickApps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NDir As String
Dim RootDir As String
Dim IDir As String

Private Sub Combo1_Click()
    Form_Load
End Sub

Private Sub Command1_Click()
    If Combo1.ListIndex = -1 Then
        MsgBox "Please select User Name."
        Combo1.SetFocus
        Exit Sub
    End If
    NetDir = NDir
    Open App.Path & "\settings.ini" For Output As #1
        Print #1, "N:" & NDir
    Close #1
    
    Unload Me
End Sub

Private Sub Form_Load()
    CurrentNSVersion = GetRegistryKey("CurrentVersion")
    NDir = GetRegistryKey("Install Directory", registryLocation & "\" & CurrentNSVersion & "\main")
    
    For X = Len(NDir) To 1 Step -1
        If Mid$(NDir, X, 1) = "\" Then
            spot = X - 1
            Exit For
        End If
    Next X
    NDir = Left$(NDir, spot) & "\Users"
    RootDir = NDir
    Dir1.Path = NDir
    
    If Combo1.ListCount > 0 Then GoTo 1
    For X = 0 To Dir1.ListCount - 1
        Combo1.AddItem UCase(Right(Dir1.List(X), (Len(Dir1.List(X)) - Len(NDir)) - 1))
    Next X
1
    NDir = NDir & "\" & Combo1.Text
    NDir = UCase(NDir)
    lblN = NDir
    If Dir(NDir & "\cookies.txt") <> "" Then
        lblStat = "Cookies.txt - Found"
    Else
        lblStat = "Cookies.txt - NOT Found"
    End If
End Sub

Private Sub Picture2_Click()
    MsgBox "Not Yet, Son!"
End Sub
