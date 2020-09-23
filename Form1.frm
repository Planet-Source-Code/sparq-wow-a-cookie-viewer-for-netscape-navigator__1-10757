VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Cookie Tosser"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   5160
      Width           =   4215
      Begin VB.CommandButton Command3 
         Caption         =   "<"
         Height          =   360
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   555
      End
      Begin VB.CommandButton Command2 
         Caption         =   ">"
         Height          =   360
         Left            =   660
         TabIndex        =   16
         Top             =   0
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1320
         TabIndex        =   18
         Top             =   60
         Width           =   75
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2595
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   4215
      Begin VB.TextBox txtValue 
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox txtExp 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2160
         Width           =   2355
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   2475
      End
      Begin VB.CheckBox chkSecure 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Secure"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1020
         TabIndex        =   7
         Top             =   180
         Width           =   855
      End
      Begin VB.CheckBox chkFlag 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Flag"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   615
      End
      Begin VB.Line Line2 
         X1              =   4185
         X2              =   4185
         Y1              =   1140
         Y2              =   1800
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   0
         Y1              =   1140
         Y2              =   1800
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1500
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name: "
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expires:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1950
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Path:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   -180
         Top             =   1140
         Width           =   4380
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4500
      TabIndex        =   0
      Top             =   5880
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   3503
      TabIndex        =   2
      Top             =   300
      Width           =   915
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   203
      TabIndex        =   1
      Top             =   300
      Width           =   3675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Occurances"
      Height          =   195
      Left            =   3510
      TabIndex        =   4
      Top             =   60
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Domains"
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   60
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NetscapeDir As String
Dim Domain(0 To 400)
Dim Occur(0 To 400)
Dim Flag(0 To 400)
Dim SpotString As String
Dim NumOf As Integer
Dim CurrPosition As Integer
Dim LinkLists As Boolean

Private Sub Command2_Click()
    NumOf = NumOf + 1
    If NumOf > Val(List2.Text) Then NumOf = Val(List2.Text)
    DisplayCurrent CurrPosition
    Label7 = "Cookie " & NumOf & "/" & List2.Text
End Sub

Private Sub Command3_Click()
  Dim Spos(0 To 50) As String
  Dim Cnt As Integer
  Dim TempSpot As Integer
  Dim CurrLett
  
  On Error GoTo Err
    NumOf = NumOf - 1
    If NumOf < 1 Then NumOf = 1
    Label7 = "Cookie " & NumOf & "/" & List2.Text
    
    Cnt = 0
    For X = 1 To Len(SpotString)
        CurrLett = Mid$(SpotString, X, 1)
        If CurrLett = "," Then Cnt = Cnt + 1: GoTo 5
        Spos(Cnt) = Spos(Cnt) & CurrLett
5
    Next X
6
    For X = 0 To Cnt
        If Val(CurrPosition) = Val(Spos(X)) Then
            TempSpot = X - 1
        End If
    Next X
    
    If TempSpot <> 0 Then
        DisplayCurrent Val(Spos(TempSpot - 1))
    Else
        DisplayCurrent Val(Spos(0))
    End If
    
    Exit Sub
Err:
    NumOf = 1
End Sub

Private Sub Form_Load()
  Dim ReadLn As String
    If Dir(App.Path & "\settings.ini") = "" Then
        DoINIFile
    Else
        Open App.Path & "\settings.ini" For Input As #1
            Line Input #1, ReadLn
            NetDir = Right$(ReadLn, Len(ReadLn) - 2)
        Close #1
    End If
    List1.Clear
    List2.Clear
    GetDomains
    Me.Visible = True
End Sub

Private Function GetDomains()
  Dim ReadLn As String
  Dim TempDomain
  Dim Hit As Boolean
  Dim X As Integer
  Dim Z As Integer
  
    X = 0
    Open NetDir & "\cookies.txt" For Input As #1
        Line Input #1, ReadLn
        Line Input #1, ReadLn
        Line Input #1, ReadLn
        Line Input #1, ReadLn
        Line Input #1, ReadLn
        Line Input #1, ReadLn

        Do While Not EOF(1)
            Line Input #1, ReadLn
            If Trim(ReadLn) = "" Then GoTo 10
            TempDomain = Left$(ReadLn, InStr(1, ReadLn, Chr(9)) - 1)
            For Z = 0 To 400
                If Trim(Domain(Z)) = Trim(TempDomain) Then
                   Occur(Z) = Occur(Z) + 1
                   Hit = True
                   Exit For
                Else
                   Hit = False
                End If
            Next Z
            If Not (Hit) Then
                X = X + 1
                Domain(X) = TempDomain
                Occur(X) = Occur(X) + 1
            End If
10
        Loop
    Close #1
    
    For Z = 0 To X
        If Trim(Domain(Z)) = "" Then GoTo 20
        List1.AddItem Domain(Z), Z - 1
        List2.AddItem Occur(Z), Z - 1
20
    Next Z
End Function

Private Function DoINIFile()
    Load frmPickApps
    frmPickApps.Show 1
End Function

Private Sub List2_Click()
    List1.ListIndex = List2.ListIndex
End Sub

Private Sub List1_Click()
    List2.ListIndex = List1.ListIndex
    NumOf = 1
    SpotString = ""
    DisplayCurrent
    Command2.Enabled = (Val(List2.Text) > 1)
    Command3.Enabled = Command2.Enabled
End Sub

Private Sub List1_Scroll()
    List2.TopIndex = List1.TopIndex
End Sub

Private Sub List2_Scroll()
    List1.TopIndex = List2.TopIndex
End Sub
 


Private Function DisplayCurrent(Optional Spot As Integer)
  On Error GoTo Err
  Dim ReadLn As String
  Dim CurrDomain As String
  Dim L As Integer
  Dim X As Integer
  Dim Spot1 As Integer
  Dim Spot2 As Integer
  Dim Flag As Boolean
  Dim Path As String
  Dim Secure As Boolean
  Dim Expiration As Date
  Dim Name As String
  Dim Value As String
  

    CurrDomain = Trim(List1.Text)
    L = Len(CurrDomain)
  
    X = Spot
    If Spot = 0 Then NumOf = 1
    Open NetDir & "\cookies.txt" For Input As #1
        If X < 1 Then GoTo 15
        For X = 1 To Spot
            Line Input #1, ReadLn
        Next X
15
        Do While Not EOF(1)
            Line Input #1, ReadLn
            X = X + 1
            If Left(ReadLn, L) = CurrDomain Then
                CurrPosition = X
                If InStr(1, SpotString, Trim(Str(Format(X, "000")))) = 0 Then
                   SpotString = SpotString & Format(X, "000") & ","
                End If
                Exit Do
            End If
        Loop
    Close #1
    
    Spot1 = L + 2
    Spot2 = InStr(Spot1, ReadLn, Chr(9))
    Flag = Mid$(ReadLn, Spot1, Spot2 - Spot1)
    
    Spot1 = Spot2 + 1
    Spot2 = InStr(Spot1, ReadLn, Chr(9))
    Path = Mid$(ReadLn, Spot1, Spot2 - Spot1)

    Spot1 = Spot2 + 1
    Spot2 = InStr(Spot1, ReadLn, Chr(9))
    Secure = Mid$(ReadLn, Spot1, Spot2 - Spot1)
    
    Spot1 = Spot2 + 1
    Spot2 = InStr(Spot1, ReadLn, Chr(9))
    Expiration = DateAdd("s", Val(Mid$(ReadLn, Spot1, Spot2 - Spot1)), "01/01/1970")

    Spot1 = Spot2 + 1
    Spot2 = InStr(Spot1, ReadLn, Chr(9))
    Name = Mid$(ReadLn, Spot1, Spot2 - Spot1)
    
    Spot1 = Spot2 + 1
    Value = Mid$(ReadLn, Spot1, Len(ReadLn) - Spot1)
    
    If Flag Then
        chkFlag.Value = 1
    Else
        chkFlag.Value = 0
    End If
    
    If Secure Then
        chkSecure.Value = 1
    Else
        chkSecure.Value = 0
    End If
    
    txtExp = Format(Expiration, "mmm dd, yyyy   hh:nn:s")
    txtPath = Path
    Label5 = "Name:   " & Name
    txtValue = Value
    Label7 = "Cookie " & NumOf & "/" & List2.Text
Err:
End Function
