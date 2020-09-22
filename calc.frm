VERSION 5.00
Begin VB.Form nachelp 
   Caption         =   "mail your comments to   anilfriend@hotmial.com"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "calc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "calc.frx":0442
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Height          =   285
      Left            =   10560
      TabIndex        =   0
      Top             =   8190
      Width           =   1005
   End
End
Attribute VB_Name = "nachelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
scientific.Show
Me.Hide
End Sub
