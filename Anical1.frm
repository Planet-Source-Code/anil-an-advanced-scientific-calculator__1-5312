VERSION 5.00
Begin VB.Form standard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " # Standard CALC #"
   ClientHeight    =   3510
   ClientLeft      =   1650
   ClientTop       =   1950
   ClientWidth     =   3510
   ForeColor       =   &H00000000&
   Icon            =   "Anical1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3510
   ScaleWidth      =   3510
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Scientific"
      Height          =   255
      Left            =   2460
      TabIndex        =   20
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Anil's STANDARD"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3255
      Begin VB.CommandButton Command5 
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   19
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "CE"
         Height          =   615
         Left            =   2520
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "C"
         Height          =   615
         Left            =   1920
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "="
         Height          =   615
         Index           =   4
         Left            =   1320
         TabIndex        =   16
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "*"
         Height          =   615
         Index           =   3
         Left            =   2520
         TabIndex        =   15
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "/"
         Height          =   615
         Index           =   2
         Left            =   1920
         TabIndex        =   14
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "--"
         Height          =   615
         Index           =   1
         Left            =   2520
         TabIndex        =   13
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+"
         Height          =   615
         Index           =   0
         Left            =   1920
         TabIndex        =   12
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         Height          =   615
         Index           =   10
         Left            =   720
         TabIndex        =   11
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "9"
         Height          =   615
         Index           =   9
         Left            =   1320
         TabIndex        =   10
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "8"
         Height          =   615
         Index           =   8
         Left            =   720
         TabIndex        =   9
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "7"
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "6"
         Height          =   615
         Index           =   6
         Left            =   1320
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "5"
         Height          =   615
         Index           =   5
         Left            =   720
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "4"
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "3"
         Height          =   615
         Index           =   3
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "2"
         Height          =   615
         Index           =   2
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "0"
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   2040
         Width           =   615
      End
   End
   Begin VB.Label text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   120
      TabIndex        =   21
      Top             =   180
      Width           =   2295
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu ecut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu ecopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu epaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu eselectall 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu eexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu type 
      Caption         =   "&Type"
      Begin VB.Menu tstandard 
         Caption         =   "S&tanfdard"
         Shortcut        =   ^T
      End
      Begin VB.Menu tscientific 
         Caption         =   "S&cientific"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu hcontents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu habout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "standard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dflag As Integer
Dim i As Integer
Dim opnre As Integer
Dim prev As Double
Dim oflag As Integer
Dim ind As Integer
Private Sub Command1_Click(Index As Integer)
    If ind = 4 Then
        prev = 0
        Text1.Caption = " "
        ind = 0
    End If
    opnre = 0
    If oflag = 0 Then
        Text1.Caption = " "
    End If
    oflag = 1
    If Command1(Index).Caption <> "." Then
           If Text1.Caption <> " 0" Then
                Text1.Caption = Text1.Caption & Command1(Index).Caption
            Else
                Text1.Caption = " " & Command1(Index).Caption
            End If
    Else
            If dflag = 0 Then
                Text1.Caption = Text1.Caption & "."
                dflag = 1
            Else
                MsgBox ("ILLEGAL SAIRAM")
            End If
     End If
            

End Sub

Private Sub Command2_Click(Index As Integer)
        If opnre = 0 Or Index = 4 Then
            If ind = 0 Then
                 prev = prev + Val(Text1.Caption)
            ElseIf ind = 1 Then
                 prev = prev - Val(Text1.Caption)
            ElseIf ind = 2 Then
                If Val(Text1.Caption) = 0 Then
                    MsgBox ("SORRY DIVIDE ZERO")
                    Exit Sub
                Else
                 prev = prev / Val(Text1.Caption)
                End If
            ElseIf ind = 3 Then
                 prev = prev * Val(Text1.Caption)
            End If
            Text1.Caption = Str(prev)
            oflag = 0
        End If
        opnre = 1
        ind = Index
        dflag = 0
End Sub

Private Sub Command3_Click()
        Text1.Caption = " 0"
        
End Sub

Private Sub Command4_Click()
    dflag = 0
    prev = 0
    oflag = 0
    ind = 0
    opnre = 0
    Text1.Caption = " 0"

End Sub

Private Sub Command5_Click()
    Unload Me
    
End Sub

Private Sub ecopy_Click()
        Clipboard.Clear
        Clipboard.SetText Text1.Caption
    
End Sub

Private Sub ecut_Click()
        Clipboard.Clear
        Clipboard.SetText Text1.Caption
        Text1.Caption = ""
       
End Sub

Private Sub eexit_Click()
        Unload Me
       
End Sub

Private Sub epaste_Click()
               Text1.Caption = ""
               Text1.Caption = Clipboard.GetText()
End Sub

Private Sub eselectall_Click()
        Clipboard.Clear
        Clipboard.SetText Text1.Caption
        End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  
   If KeyAscii = Asc(".") Then
        i = 10
         Command1_Click (i)
         Beep
   ElseIf KeyAscii = Asc("0") Then
        i = 0
         Command1_Click (i)
         Beep
   ElseIf KeyAscii = Asc("1") Then
        i = 1
          Command1_Click (i)
          Beep
   ElseIf KeyAscii = Asc("2") Then
        i = 2
          Command1_Click (i)
        Beep
   ElseIf KeyAscii = Asc("3") Then
        i = 3
          Command1_Click (i)
          Beep
   ElseIf KeyAscii = Asc("4") Then
        i = 4
          Command1_Click (i)
          Beep
   ElseIf KeyAscii = Asc("5") Then
        i = 5
          Command1_Click (i)
          Beep
   ElseIf KeyAscii = Asc("6") Then
        i = 6
          Command1_Click (i)
          Beep
   ElseIf KeyAscii = Asc("7") Then
        i = 7
          Command1_Click (i)
          Beep
   ElseIf KeyAscii = Asc("8") Then
        i = 8
          Command1_Click (i)
          Beep
   ElseIf KeyAscii = Asc("9") Then
        i = 9
          Command1_Click (i)
          Beep
   ElseIf KeyAscii = Asc("0") Then
        i = 0
          Command1_Click (i)
          Beep
   ElseIf KeyAscii = Asc("+") Then
        i = 0
          Command2_Click (i)
          Beep
   ElseIf KeyAscii = Asc("+") Then
        i = 0
          Command2_Click (i)
          Beep
   ElseIf KeyAscii = Asc("-") Then
        i = 1
          Command2_Click (i)
          Beep
   ElseIf KeyAscii = Asc("/") Then
        i = 2
          Command2_Click (i)
          Beep
 
   ElseIf KeyAscii = Asc("*") Then
        i = 3
          Command2_Click (i)
          Beep
   ElseIf KeyAscii = Asc("=") Then
        i = 4
          Command2_Click (i)
          Beep
   ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
        dflag = 0
        prev = 0
        oflag = 0
        ind = 0
        opnre = 0
        Text1.Caption = " 0"
        Beep
        Beep
   ElseIf KeyAscii = Asc("d") Or KeyAscii = Asc("D") Then
        Text1.Caption = " 0"
        Beep
   End If
End Sub


Private Sub Form_Load()
  '  standard.Height = 4090
  '  standard.Width = 3430
    dflag = 0
    prev = 0
    oflag = 0
    ind = 0
    opnre = 0
    Clipboard.Clear
End Sub

Private Sub habout_Click()
nachelp.Show
End Sub

Private Sub hcontents_Click()
     nachelp.Show
End Sub

Private Sub Option1_Click()
    scientific.Show
    standard.Option1.Value = False
    standard.Hide
End Sub

Private Sub tscientific_Click()
    scientific.Show
    standard.Hide

End Sub

Private Sub tstandard_Click()
    standard.Show
    scientific.Hide
End Sub


