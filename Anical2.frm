VERSION 5.00
Begin VB.Form scientific 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "# Scientific Calc #"
   ClientHeight    =   4305
   ClientLeft      =   1230
   ClientTop       =   2220
   ClientWidth     =   6420
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Anical2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4305
   ScaleWidth      =   6420
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   2880
      TabIndex        =   65
      Top             =   90
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.OptionButton Option3 
      Caption         =   "STANDARD"
      Height          =   315
      Left            =   3030
      TabIndex        =   58
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "SCIENTIFIC"
      Height          =   3210
      Left            =   90
      TabIndex        =   0
      Top             =   990
      Width           =   6255
      Begin VB.CommandButton Command8 
         Caption         =   "View"
         Height          =   375
         Index           =   1
         Left            =   3810
         TabIndex        =   63
         Top             =   2610
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "xPOy"
         Height          =   375
         Index           =   8
         Left            =   3810
         TabIndex        =   62
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "npr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   3810
         TabIndex        =   61
         Top             =   1740
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ncr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3810
         TabIndex        =   60
         Top             =   1260
         Width           =   495
      End
      Begin VB.CommandButton Command14 
         Caption         =   "±"
         Height          =   375
         Left            =   2520
         TabIndex        =   59
         Top             =   1710
         Width           =   495
      End
      Begin VB.CommandButton Command13 
         Caption         =   "EXIT"
         Height          =   375
         Left            =   5610
         TabIndex        =   57
         Top             =   270
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "atan"
         Height          =   375
         Index           =   5
         Left            =   3120
         TabIndex        =   56
         Top             =   270
         Width           =   495
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Back"
         Height          =   375
         Left            =   3810
         TabIndex        =   55
         Top             =   270
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "acos"
         Height          =   375
         Index           =   4
         Left            =   2520
         TabIndex        =   54
         Top             =   270
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "asin"
         Height          =   375
         Index           =   3
         Left            =   1920
         TabIndex        =   53
         Top             =   270
         Width           =   495
      End
      Begin VB.CommandButton Command12 
         Caption         =   "NOT"
         Height          =   375
         Index           =   2
         Left            =   5610
         TabIndex        =   52
         Top             =   1710
         Width           =   495
      End
      Begin VB.CommandButton Command12 
         Caption         =   "OR"
         Height          =   375
         Index           =   1
         Left            =   5610
         TabIndex        =   51
         Top             =   1260
         Width           =   495
      End
      Begin VB.CommandButton Command12 
         Caption         =   "AND"
         Height          =   375
         Index           =   0
         Left            =   5610
         TabIndex        =   50
         Top             =   810
         Width           =   495
      End
      Begin VB.CommandButton Command9 
         Caption         =   "h"
         Height          =   375
         Index           =   4
         Left            =   2520
         TabIndex        =   49
         Top             =   2610
         Width           =   495
      End
      Begin VB.CommandButton Command9 
         Caption         =   "g"
         Height          =   375
         Index           =   3
         Left            =   1920
         TabIndex        =   48
         Top             =   2610
         Width           =   495
      End
      Begin VB.CommandButton Command9 
         Caption         =   "e"
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   47
         Top             =   2610
         Width           =   495
      End
      Begin VB.CommandButton Command9 
         Caption         =   "pi"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   46
         Top             =   2610
         Width           =   495
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Rnd"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   45
         Top             =   2610
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "%"
         Height          =   375
         Index           =   6
         Left            =   1920
         TabIndex        =   44
         Top             =   1710
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "00"
         Height          =   375
         Index           =   11
         Left            =   720
         TabIndex        =   43
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command10 
         Caption         =   "e(x]"
         Height          =   375
         Index           =   2
         Left            =   3120
         TabIndex        =   42
         Top             =   810
         Width           =   495
      End
      Begin VB.CommandButton Command10 
         Caption         =   "log"
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   41
         Top             =   1260
         Width           =   495
      End
      Begin VB.CommandButton Command10 
         Caption         =   "ln"
         Height          =   375
         Index           =   0
         Left            =   3810
         TabIndex        =   40
         Top             =   810
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "X ³"
         Height          =   375
         Index           =   8
         Left            =   3120
         TabIndex        =   39
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "copy"
         Height          =   375
         Index           =   3
         Left            =   5040
         TabIndex        =   38
         Top             =   2190
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "del"
         Height          =   375
         Index           =   2
         Left            =   5610
         TabIndex        =   37
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "push"
         Height          =   375
         Index           =   0
         Left            =   4410
         TabIndex        =   36
         Top             =   2190
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "MC"
         Height          =   375
         Index           =   5
         Left            =   5010
         TabIndex        =   35
         Top             =   1710
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "MR"
         Height          =   375
         Index           =   4
         Left            =   4410
         TabIndex        =   34
         Top             =   1710
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "M /"
         Height          =   375
         Index           =   3
         Left            =   5010
         TabIndex        =   29
         Top             =   1260
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "M *"
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   28
         Top             =   1260
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "M -"
         Height          =   375
         Index           =   1
         Left            =   5010
         TabIndex        =   27
         Top             =   810
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "M +"
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   26
         Top             =   810
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "n !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3120
         TabIndex        =   25
         Top             =   1710
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Sq[x]"
         Height          =   375
         Index           =   2
         Left            =   3120
         TabIndex        =   24
         Top             =   2610
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "X ²"
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   23
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "1/ X"
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   22
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "tan"
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   21
         Top             =   270
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "cos"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   20
         Top             =   270
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "sin"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   270
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "AC"
         Height          =   375
         Index           =   0
         Left            =   5010
         TabIndex        =   18
         Top             =   270
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "C"
         Height          =   375
         Index           =   0
         Left            =   4410
         TabIndex        =   17
         Top             =   270
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "="
         Height          =   375
         Index           =   4
         Left            =   4380
         TabIndex        =   16
         Top             =   2610
         Width           =   1725
      End
      Begin VB.CommandButton Command2 
         Caption         =   "*"
         Height          =   375
         Index           =   3
         Left            =   2520
         TabIndex        =   15
         Top             =   1260
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "/"
         Height          =   375
         Index           =   2
         Left            =   1950
         TabIndex        =   14
         Top             =   1260
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "-"
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   13
         Top             =   810
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+"
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   12
         Top             =   810
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         Height          =   375
         Index           =   10
         Left            =   1320
         TabIndex        =   11
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "9"
         Height          =   375
         Index           =   9
         Left            =   1320
         TabIndex        =   10
         Top             =   1710
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "8"
         Height          =   375
         Index           =   8
         Left            =   720
         TabIndex        =   9
         Top             =   1710
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "7"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   1710
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "6"
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   7
         Top             =   1260
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "5"
         Height          =   375
         Index           =   5
         Left            =   720
         TabIndex        =   6
         Top             =   1260
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "4"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1260
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "3"
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   4
         Top             =   810
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "2"
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   3
         Top             =   810
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   810
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "0"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   3  'Dot
         Height          =   2310
         Left            =   60
         Top             =   765
         Width           =   3615
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Height          =   2310
         Left            =   3750
         Top             =   750
         Width           =   2475
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Height          =   465
         Left            =   60
         Top             =   225
         Width           =   6105
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "angle"
      Height          =   495
      Left            =   90
      TabIndex        =   30
      Top             =   420
      Width           =   2055
      Begin VB.OptionButton Option2 
         Caption         =   "grd"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   33
         Top             =   180
         Width           =   645
      End
      Begin VB.OptionButton Option2 
         Caption         =   "rad"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   32
         Top             =   180
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "deg"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.Label Label2 
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
      Height          =   285
      Left            =   120
      TabIndex        =   66
      Top             =   90
      Width           =   2715
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4350
      TabIndex        =   64
      Top             =   150
      Width           =   165
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu ecut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu ecopy 
         Caption         =   "C&opy"
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
         Caption         =   "Select &All"
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
         Caption         =   "S&tandard"
         Shortcut        =   ^T
      End
      Begin VB.Menu tscientific 
         Caption         =   "S&cieintific"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu hcontents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu hsep 
         Caption         =   "-"
      End
      Begin VB.Menu habout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "scientific"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim que(25) As Double
Public qt As Integer
Public qh As Integer
Public qv As Integer
Public ang As Double

Public memo As Double
Dim dflag As Integer
Dim i As Integer
Dim opnre As Integer
Dim prev As Double
Dim oflag As Integer
Dim ind As Integer
Private Sub Command1_Click(Index As Integer)
    If ind = 4 Then
        prev = 0
        Text1.Text = " "
        ind = 0
    End If
    opnre = 0  ' conform operand
    If oflag = 0 Then   ' if previous operator clear text
        Text1.Text = " "
    End If
    oflag = 1
    If Command1(Index).Caption <> "." Then
           If Text1.Text <> " 0" Then
                Text1.Text = Text1.Text & Command1(Index).Caption
            Else
                Text1.Text = " " & Command1(Index).Caption
            End If
    Else
            If dflag = 0 Then
                Text1.Text = Text1.Text & "."
                dflag = 1
            Else
                MsgBox ("ILLEGAL SAIRAM")
            End If
     End If
            

End Sub

Private Sub Command10_Click(Index As Integer)
        If Index = 2 Then
            If Val(Text1.Text) < 700 Then
                Text1.Text = Str(Exp(Val(Text1.Text)))
            Else
                MsgBox (" OVERFLOW. VALUE TOO BIG ")
            End If
        ElseIf Index = 0 Then
            If Val(Text1.Text) > 0 Then
                Text1.Text = Str(Log(Val(Text1.Text)))
            Else
                MsgBox (" ILLEGAL.  LOG  NON  POSITIVE ")
            End If
        ElseIf Index = 1 Then
             If Val(Text1.Text) > 0 Then
                Text1.Text = Str((Log(Val(Text1.Text)) / Log(10)))
             Else
                MsgBox (" ILLEGAL. LOG  NON  POSITIVE ")
             End If

        End If
End Sub

Private Sub Command11_Click()
    If (Text1.Text <> "") Then
        Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text) - 1)
    End If
End Sub

Private Sub Command13_Click()
        Unload Me
End Sub

Private Sub Command14_Click()
            Text1.Text = Str(Val(Text1.Text) * -1)
End Sub
Function qnext(p As Integer) As Integer
       On Error Resume Next
       qnext = (p + 1) Mod qt
End Function

Function power(x As Double, Y As Long) As Double
        Dim i As Double
        i = 1
        If (Y > 0) Then
            While Y > 0
                Y = Y - 1
                i = i * x
            Wend
            power = i
        ElseIf (Y = 0) Then
            power = 1
        Else
            MsgBox ("ILLEGAL. POWER LESTHAN 0.")
        End If
End Function
Function fact(num As Long) As Long
    If (num < 0 Or num = 0) Then
        MsgBox ("ILLEGAL  NEAGETIVE  FACTORIAL")
        fact = num
    Else
        If (num > 12) Then
            MsgBox ("VALUE TOO LARGE")
            fact = num
        Else
            re = 1
            While (num > 0)
                re = re * num
                num = num - 1
            Wend
            fact = re
        End If
    End If
        
End Function
Private Sub Command2_Click(Index As Integer)
        Dim n As Long
        Dim r As Long
        If opnre = 0 Or Index = 4 Then
            If ind = 0 Then
                 prev = prev + Val(Text1.Text)
            ElseIf ind = 1 Then
                 prev = prev - Val(Text1.Text)
            ElseIf ind = 6 Then
                prev = prev Mod Val(Text1.Text)
            ElseIf ind = 7 Then
                r = Fix(Val(Text1.Text))
                n = Fix(Val(prev))
                If ((n > r Or n = r) And n > 0 And (r > 0 Or r = 0)) Then
                    prev = fact(n) / (fact(n - r))
                Else
                    MsgBox ("ILLEGAL  ENTRIES  of  N ,R ")
                End If

            ElseIf ind = 5 Then
                r = Fix(Val(Text1.Text))
                n = Fix(Val(prev))
                If ((n > r Or n = r) And n > 0 And (r > 0 Or r = 0)) Then
                    prev = fact(n) / (fact(n - r) * fact(r))
                Else
                    MsgBox ("ILLEGAL  ENTRIES  of  N ,R ")
                End If
            ElseIf ind = 8 Then
                If (Text1.Text = "" Or prev = 0) Then
                    MsgBox ("ILLEGAL.  INVALIED ENTRIES")
                Else
                        prev = (power(prev, Fix(Val(Text1.Text))))
                         
                End If
           
            ElseIf ind = 2 Then
                 If Val(Text1.Text) <> 0 Then
                     prev = prev / Val(Text1.Text)
                  Else
                      MsgBox (" ILLEGAL  DIVIDE  0 ")
                  End If
            ElseIf ind = 3 Then
                 prev = prev * Val(Text1.Text)
            End If
            Text1.Text = Str(prev)
            oflag = 0   ' operator or operand
        End If
        opnre = 1  ' multiple operators flag
        ind = Index
        dflag = 0
End Sub

Private Sub Command3_Click(Index As Integer)
        Text1.Text = " 0"
End Sub

Private Sub Command4_Click(Index As Integer)
    memo = 0
    dflag = 0
    prev = 0
    oflag = 0
    ind = 0
    opnre = 0
    qh = 0
    qt = 0
    Clipboard.Clear
    Text1.Text = " 0"
End Sub

Private Sub Command5_Click(Index As Integer)
        Select Case Index
            Case 0
                Text1.Text = Str(Sin(ang * Val(Text1.Text)))
            Case 1
                Text1.Text = Str(Cos(ang * Val(Text1.Text)))
            Case 2
                If (Cos(Val(Text1.Text))) <> 0 Then
                    Text1.Text = Str(Sin(ang * Val(Text1.Text)) / Cos(ang * Val(Text1.Text)))
                Else
                    MsgBox ("ILLEGAL.  DIVIDE  BY  ZERO ")
                End If
            Case 5
                Text1.Text = Str((Atn(Val(Text1.Text))) / ang)
            End Select
End Sub

Private Sub Command6_Click(Index As Integer)
            Dim re As Long
            Dim temp As Long
            
            temp = Val(scientific.Text1.Text)
            Select Case Index
            Case 2
'                temp = Val(Text1.Text)
                If temp > 0 Or temp = 0 Then
                    scientific.Text1.Text = Str(Sqr(Val(Text1.Text)))
                Else
                    MsgBox (" ILLEGAL  ATTEMPTING  NEGETIVE  ROOT")
                End If
            
            Case 0
                temp = Val(Text1.Text)
                If temp <> 0 Then
                    scientific.Text1.Text = Str(1 / temp)
                Else
                     MsgBox (" ILLEGAL  DIVIDE  0 ")
                    
                End If
            Case 1
                If Abs(Val(Text1.Text)) < 46300 Then
                    scientific.Text1.Text = Str((temp * temp))
                Else
                    MsgBox (" ILLEGAL  DIVIDE  0 ")
                End If
            Case 8
                If Abs(Val(Text1.Text)) < 1290 Then
                    scientific.Text1.Text = Str(temp * temp * temp)
                Else
                    MsgBox ("OVERFLOW. VALUE TOO LARGE ")
                End If
            Case 4
                    Text1.Text = Str(fact(Val(Text1.Text)))
                     End Select
                    
                
                    
End Sub

Private Sub Command7_Click(Index As Integer)
      Select Case Index
            Case 0
                 memo = memo + Val(Text1.Text)
            Case 1
                 memo = memo - Val(Text1.Text)
            Case 2
                 memo = memo * Val(Text1.Text)
            Case 3
                If Val(Text1.Text) <> 0 Then
                     memo = memo / Val(Text1.Text)
                Else
                    MsgBox ("ILLEGAL. DIVIDE 0 ERROR ")
                End If
            Case 4
                Text1.Text = Str(memo)
                prev = Val(Text1.Text)
            Case 5
                memo = 0
        End Select
End Sub

Private Sub Command8_Click(Index As Integer)
        Select Case Index
            Case 0
                que(qt) = Val(Text1.Text)
                qt = qt + 1
            Case 1
                qv = qnext(qv)
                Label1.Caption = Str(que(qv))
'                qv = qnext(qv)
            Case 2
                On Error GoTo anil
                que(qv) = que(qt - 1)
                Label1.Caption = Str(que(qv))
                qt = qt - 1
                
            Case 3
                Text1.Text = Label1.Caption
        End Select
anil:
     
End Sub

Private Sub Command9_Click(Index As Integer)
    Select Case Index
        Case 0
            Text1.Text = Str(Rnd)
        Case 1
            Text1.Text = 3.141592654
        Case 2
            Text1.Text = 2.718281828
        Case 3
            Text1.Text = 9.86
        Case 4
            Text1.Text = "6.625"
    End Select
        opnre = 0
'        If oflag = 0 Then
'             Text1.Text = " "
'        End If
        oflag = 1
End Sub


'Private Sub Command4_Click()
'    dflag = 0
'    prev = 0
'    oflag = 0
'    ind = 0
'    opnre = 0
'    Text1.Text = " 0"

'End Sub

'Private Sub Command5_Click()
'    Unload Me
'    End
'End Sub

Private Sub ecopy_Click()
        Clipboard.Clear
        Clipboard.SetText Text1.SelText
    
End Sub

Private Sub ecut_Click()
        Clipboard.Clear
        Clipboard.SetText Text1.SelText
        Text1.SelText = ""
       
End Sub

Private Sub eexit_Click()
        Unload Me
       
End Sub

Private Sub epaste_Click()
               Text1.Text = ""
               Text1.SelText = Clipboard.GetText()
End Sub

Private Sub eselectall_Click()
        Clipboard.Clear
        Clipboard.SetText Text1.Text
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
        Text1.Text = " 0"
        Beep
        Beep
   ElseIf KeyAscii = Asc("d") Or KeyAscii = Asc("D") Then
        Text1.Text = " 0"
        Beep
   End If
End Sub

Private Sub Form_Load()
    scientific.Width = 6540
    scientific.Height = 4735
    memo = 0
    dflag = 0
    prev = 0
    oflag = 0
    ind = 0
    opnre = 0
    qh = 0
    qt = 0
    Clipboard.Clear
    Option2(0).Value = True
    ang = 3.14 / 180
End Sub

Private Sub habout_Click()
         nachelp.Show
End Sub

Private Sub hcontents_Click()
naccalc.Show

End Sub

Private Sub Option2_Click(Index As Integer)
        Select Case Index
            Case 0
                ang = 3.141592654 / 180
            Case 1
                ang = 1
            Case 2
                ang = 3.141592654 / 200
            End Select
End Sub

Private Sub Option3_Click()
         standard.Show
        scientific.Hide
        scientific.Option3.Value = False
       

End Sub




Private Sub Text1_Change()
Label2.Caption = Text1.Text

End Sub

Private Sub tscientific_Click()
    scientific.Show
    standard.Hide

End Sub

Private Sub tstandard_Click()
    standard.Show
    scientific.Hide
End Sub


             
