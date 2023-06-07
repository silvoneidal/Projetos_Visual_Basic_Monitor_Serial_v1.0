VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desconectado"
   ClientHeight    =   7260
   ClientLeft      =   210
   ClientTop       =   540
   ClientWidth     =   13575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   13575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkHex 
      Caption         =   "Hexadecimal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Deleta a saída"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   11880
      TabIndex        =   6
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdSend 
      Appearance      =   0  'Flat
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox cboBaudRate 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Main.frx":169B2
      Left            =   9840
      List            =   "Main.frx":169DA
      TabIndex        =   4
      Text            =   "9600"
      Top             =   6840
      Width           =   1935
   End
   Begin VB.ComboBox cboEndStr 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Main.frx":16A2B
      Left            =   7560
      List            =   "Main.frx":16A3B
      TabIndex        =   3
      Text            =   "Nenhum final-de-linha"
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CheckBox chkAuto 
      Caption         =   "Auto-rolagem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox txtReceive 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   13335
   End
   Begin VB.TextBox txtSend 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      MaxLength       =   100
      TabIndex        =   0
      Top             =   120
      Width           =   11895
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6960
      Top             =   6840
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6240
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
   End
   Begin VB.Menu mMenu 
      Caption         =   "Menu"
      Begin VB.Menu mConectar 
         Caption         =   "Conectar"
         Begin VB.Menu mCOM1 
            Caption         =   "COM1"
            Enabled         =   0   'False
         End
         Begin VB.Menu mCOM2 
            Caption         =   "COM2"
            Enabled         =   0   'False
         End
         Begin VB.Menu mCOM3 
            Caption         =   "COM3"
            Enabled         =   0   'False
         End
         Begin VB.Menu mCOM4 
            Caption         =   "COM4"
            Enabled         =   0   'False
         End
         Begin VB.Menu mCOM5 
            Caption         =   "COM5"
            Enabled         =   0   'False
         End
         Begin VB.Menu mCOM6 
            Caption         =   "COM6"
            Enabled         =   0   'False
         End
         Begin VB.Menu mCOM7 
            Caption         =   "COM7"
            Enabled         =   0   'False
         End
         Begin VB.Menu mCOM8 
            Caption         =   "COM8"
            Enabled         =   0   'False
         End
         Begin VB.Menu mCOM9 
            Caption         =   "COM9"
            Enabled         =   0   'False
         End
         Begin VB.Menu mCOM10 
            Caption         =   "COM10"
            Enabled         =   0   'False
         End
         Begin VB.Menu mCOM11 
            Caption         =   "COM11"
            Enabled         =   0   'False
         End
         Begin VB.Menu mCOM12 
            Caption         =   "COM12"
            Enabled         =   0   'False
         End
         Begin VB.Menu mCOM13 
            Caption         =   "COM13"
            Enabled         =   0   'False
         End
         Begin VB.Menu mCOM14 
            Caption         =   "COM14"
            Enabled         =   0   'False
         End
         Begin VB.Menu mCOM15 
            Caption         =   "COM15"
            Enabled         =   0   'False
         End
         Begin VB.Menu mCOM16 
            Caption         =   "COM16"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mDesconectar 
         Caption         =   "Desconectar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mScanear 
         Caption         =   "Scanear"
         Begin VB.Menu mPortaCOM 
            Caption         =   "PortaCOM"
         End
      End
      Begin VB.Menu mGerenciador 
         Caption         =   "Gerênciador"
         Begin VB.Menu mDispositivos 
            Caption         =   "Dispositivos"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sleep
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim portaCOM As Integer

Private Sub Form_Load()
    
On Error GoTo Erro
    
   'Rotina para detectar portas COM disponíveis
   Dim i As Integer
   For i = 1 To 16
       If DetectaPortaCOM(i) <> 0 Then
           Select Case i
              Case 1
                 mCOM1.Enabled = True
              Case 2
                 mCOM2.Enabled = True
              Case 3
                 mCOM3.Enabled = True
              Case 4
                 mCOM4.Enabled = True
              Case 5
                 mCOM5.Enabled = True
              Case 6
                 mCOM6.Enabled = True
              Case 7
                 mCOM7.Enabled = True
              Case 8
                 mCOM8.Enabled = True
              Case 9
                 mCOM9.Enabled = True
              Case 10
                 mCOM10.Enabled = True
              Case 11
                 mCOM11.Enabled = True
              Case 12
                 mCOM12.Enabled = True
              Case 13
                 mCOM13.Enabled = True
              Case 14
                 mCOM14.Enabled = True
              Case 15
                 mCOM15.Enabled = True
              Case 16
                 mCOM16.Enabled = True
           End Select
           
       End If
   Next
        
Exit Sub
    
Erro:
    MsgBox "Erro " & Err & ". " & Error, vbApplicationModal, "DALÇOQUIO AUTOMAÇÃO"
    Beep
    
End Sub

Private Sub cmdConexao()

On Error GoTo Erro

   If MSComm1.PortOpen = True Then
      MSComm1.PortOpen = False
      Form1.Caption = "Desconectado"
      mConectar.Enabled = True
      mDesconectar.Enabled = False
      Call uncheckCOM
      Timer1.Enabled = True
      mScanear.Enabled = True
   Else
      MSComm1.CommPort = portaCOM
      MSComm1.Settings = cboBaudRate.Text & "n,8,1"
      MSComm1.PortOpen = True
      MSComm1.DTREnable = True
      MSComm1.RTSEnable = True
      Form1.Caption = "Conectado na COM" & portaCOM & " por DALÇOQUIO AUTOMAÇÃO"
      mConectar.Enabled = False
      mDesconectar.Enabled = True
      Timer1.Enabled = False
      mScanear.Enabled = False
   End If

Exit Sub
   
Erro:
    Beep
    MsgBox "Erro " & Err & ". " & Error, vbApplicationModal, "DALÇOQUIO AUTOMAÇÃO"

   
End Sub

Private Sub cboBaudRate_Click()
   If MSComm1.PortOpen = True Then
      MSComm1.PortOpen = False
      MSComm1.PortOpen = True
      MSComm1.DTREnable = True
      MSComm1.RTSEnable = True
      MSComm1.Settings = cboBaudRate.Text & "n,8,1"
   End If
    
End Sub

Private Sub cmdSend_Click()

On Error GoTo Erro

   If chkHex.Value = 1 Then
      Call sendHex
   Else
          If cboEndStr.Text = "Nenhum final-de-linha" Then
            MSComm1.Output = txtSend.Text
      ElseIf cboEndStr.Text = "Nova-linha" Then
            MSComm1.Output = txtSend.Text & vbLf
      ElseIf cboEndStr.Text = "Retorno de carro" Then
            MSComm1.Output = txtSend.Text & vbCr
      ElseIf cboEndStr.Text = "Ambos, NL e CR" Then
            MSComm1.Output = txtSend.Text & vbCrLf
      End If
    
    End If
    
Exit Sub
   
Erro:
    Beep
    MsgBox "Erro " & Err & ". " & Error, vbApplicationModal, "DALÇOQUIO AUTOMAÇÃO"

End Sub

Private Sub sendHex()

On Error GoTo Erro

   Dim A As String
   Dim B As Integer
   A = txtSend.Text
   A = Replace(A, " ", "")
   B = Len(A) / 2

   Dim byteData(100) As Byte '5A A5 05 82 20 00 00 01
   Dim x As Integer
   Dim y As Integer
   y = 1
   
   For x = 0 To B
      If Mid(txtSend.Text, y, 2) = Empty Then Exit For
      byteData(x) = "&H" & Mid(txtSend.Text, y, 2)
      y = y + 3
   Next x

   MSComm1.Output = byteData
   
Exit Sub
   
Erro:
    Beep
    MsgBox "Erro " & Err & ". " & Error, vbApplicationModal, "DALÇOQUIO AUTOMAÇÃO"

End Sub

Private Sub mDesconectar_Click()
   Call cmdConexao

End Sub

Private Sub mDispositivos_Click()
    Shell ("cmd.exe /c devmgmt.msc")

End Sub

Private Sub mPortaCOM_Click()
   mScanear.Enabled = False
   Timer1.Enabled = False
   Form1.Caption = "Scaneando, aguarde..."
   Call Form_Load
   Form1.Caption = "Desconectado"
   Timer1.Enabled = True
   mScanear.Enabled = True
   Beep
   
End Sub

Private Sub MSComm1_OnComm()

On Error GoTo Erro

    Dim receive As String
    Dim hexa As String

    receive = MSComm1.Input
    
    If chkHex.Value = 1 Then
      'HEX
      txtReceive.Text = txtReceive.Text + StringtoHex(receive) & "   "
    Else
      'CHR
      txtReceive.Text = txtReceive.Text + receive
    End If
    
    If chkAuto.Value = 1 Then
      'txtReceive.SelStart = 65535 'IDE Arduino
      txtReceive.SelStart = Len(txtReceive.Text)
    End If
    
Exit Sub
   
Erro:
    Beep
    MsgBox "Erro " & Err & ". " & Error, vbApplicationModal, "DALÇOQUIO AUTOMAÇÃO"
    
    '------------------------------------------------------------------------------------------
    'Mid: Retorna o número especificado de caracteres de uma string.
    'exemplo: mid(text1.text,1,5) -> retorna as letras 1,2,3,4,5 do text1.
    'exemplo: mid(text1.text,20,5) -> retorna  as ultimas 5 letras iniciando da posicai 20 do text1.
    
    'Left:Retorna o número especificado de caracteres a partir do início de uma string.
    'exemplo: left(text1.text,3) -> retorna as 3 primeiras letras do text1.
    
    'right:Retorna o número especificado de caracteres a partir do lado direito de uma string.
    'exemplo: right(text1.text, 4) -> retorna as quatro últimas letras do text1.
    '------------------------------------------------------------------------------------------
    
End Sub

Private Function StringtoHex(ByVal tmpStr As String) As String
Dim i As Integer
Dim tmpl As String
    For i = 1 To Len(tmpStr)
         If Len(Hex(Asc(Mid$(tmpStr, i, 1)))) > 1 Then
            tmpl = tmpl & Hex(Asc(Mid$(tmpStr, i, 1))) & " "
         Else
            tmpl = tmpl & "0" & Hex(Asc(Mid$(tmpStr, i, 1))) & " " 'Ex:  5A A5 06 83 10 00 01 00 01
         End If
    Next i
    StringtoHex = Trim$(tmpl)
    
'tmpl = tmpl & Hex(Asc(Mid$(tmpStr, i, 1))) & " " 'Ex:  5A A5 6 83 10 0 1 0 1
'tmpl = tmpl & "[" & Hex(Asc(Mid$(tmpStr, i, 1))) & "]" & " " 'Ex:  [5A] [A5] [6] [83] [10] [0] [1] [0 [1]
'tmpl = tmpl & "[" & i - 1 & "]" & Hex(Asc(Mid$(tmpStr, i, 1))) & " " 'Ex:  [1]5A [2]A5 [3]6 [4]83 [5]10 [6]0 [7]1 [8]0 [9]1
'tmpl = tmpl & i - 1 & "[" & Hex(Asc(Mid$(tmpStr, i, 1))) & "]" & " " 'Ex:  1[5A] 2[A5] 3[6] 4[83] 5[10] 6[0] 7[1] 8[0] 9[1]
    
End Function
Private Sub chkHex_Click()
   txtReceive.Text = Empty

End Sub

Private Sub chkAuto_Click()
   If chkAuto.Value = True Then
   Else
   End If
End Sub

Private Sub mCOM1_Click()
   Call uncheckCOM
   mCOM1.Checked = True
   portaCOM = 1
   Call cmdConexao
   
End Sub

Private Sub mCOM2_Click()
   Call uncheckCOM
   mCOM2.Checked = True
   portaCOM = 2
   Call cmdConexao
   
End Sub

Private Sub mCOM3_Click()
   Call uncheckCOM
   mCOM3.Checked = True
   portaCOM = 3
   Call cmdConexao
   
End Sub

Private Sub mCOM4_Click()
   Call uncheckCOM
   mCOM4.Checked = True
   portaCOM = 4
   Call cmdConexao
   
End Sub

Private Sub mCOM5_Click()
   Call uncheckCOM
   mCOM5.Checked = True
   portaCOM = 5
   Call cmdConexao
   
End Sub

Private Sub mCOM6_Click()
   Call uncheckCOM
   mCOM6.Checked = True
   portaCOM = 6
   Call cmdConexao
   
End Sub

Private Sub mCOM7_Click()
   Call uncheckCOM
   mCOM7.Checked = True
   portaCOM = 7
   Call cmdConexao
   
End Sub

Private Sub mCOM8_Click()
   Call uncheckCOM
   mCOM8.Checked = True
   portaCOM = 8
   Call cmdConexao
   
End Sub

Private Sub mCOM9_Click()
   Call uncheckCOM
   mCOM9.Checked = True
   portaCOM = 9
   Call cmdConexao
   
End Sub

Private Sub mCOM10_Click()
   Call uncheckCOM
   mCOM10.Checked = True
   portaCOM = 10
   Call cmdConexao
   
End Sub

Private Sub mCOM11_Click()
   Call uncheckCOM
   mCOM11.Checked = True
   portaCOM = 11
   Call cmdConexao
   
End Sub

Private Sub mCOM12_Click()
   Call uncheckCOM
   mCOM12.Checked = True
   portaCOM = 12
   Call cmdConexao
   
End Sub

Private Sub mCOM13_Click()
   Call uncheckCOM
   mCOM13.Checked = True
   portaCOM = 13
   Call cmdConexao
   
End Sub

Private Sub mCOM14_Click()
   Call uncheckCOM
   mCOM14.Checked = True
   portaCOM = 14
   Call cmdConexao
   
End Sub

Private Sub mCOM15_Click()
   Call uncheckCOM
   mCOM15.Checked = True
   portaCOM = 15
   Call cmdConexao
   
End Sub

Private Sub mCOM16_Click()
   Call uncheckCOM
   mCOM16.Checked = True
   portaCOM = 16
   Call cmdConexao
   
End Sub

Private Sub uncheckCOM()
   mCOM1.Checked = False
   mCOM2.Checked = False
   mCOM3.Checked = False
   mCOM4.Checked = False
   mCOM5.Checked = False
   mCOM6.Checked = False
   mCOM7.Checked = False
   mCOM8.Checked = False
   mCOM9.Checked = False
   mCOM10.Checked = False
   mCOM11.Checked = False
   mCOM12.Checked = False
   mCOM13.Checked = False
   mCOM14.Checked = False
   mCOM15.Checked = False
   mCOM16.Checked = False

End Sub

Private Sub cmdClear_Click()
   txtReceive.Text = Empty
   
End Sub

Private Sub Timer1_Timer()
   If Form1.Caption = "Desconectado" Then
      Form1.Caption = Empty
   Else
      Form1.Caption = "Desconectado"
   End If
   
End Sub

'Rotina para desconectar porta COM ao clickar em "X"

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True 'cliente esta fechando o sistema pelo "X"
    
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False 'fecha porta COM...
    End If
    
    End ' fecha o sistema...
End Sub

















