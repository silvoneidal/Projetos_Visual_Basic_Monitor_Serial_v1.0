VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Porta"
   ClientHeight    =   1260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2445
   LinkTopic       =   "Form2"
   ScaleHeight     =   1260
   ScaleWidth      =   2445
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "1"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Conectar na COM:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form1
Form1.Show
Unload Form2
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Unload Form1
Form1.Show
Unload Form2
End If
End Sub

Private Sub form_load()
Text1.SelLength = Len(Text1.Text)
End Sub
