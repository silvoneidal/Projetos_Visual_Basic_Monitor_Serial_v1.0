VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conexão [Desconectado]"
   ClientHeight    =   10575
   ClientLeft      =   210
   ClientTop       =   540
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10575
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox CheckPausa 
      Caption         =   "Pausa"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   19
      Top             =   9720
      Width           =   1215
   End
   Begin VB.TextBox txtClear 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   240
      MaxLength       =   1
      TabIndex        =   17
      Top             =   9720
      Width           =   360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5280
      Top             =   7320
   End
   Begin VB.Frame Frame3 
      Caption         =   "Conexão"
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   7455
      Begin VB.ComboBox cboBaudRate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "MSComm.frx":0000
         Left            =   5160
         List            =   "MSComm.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cboPortaCom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "MSComm.frx":008B
         Left            =   2040
         List            =   "MSComm.frx":00CB
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label_BaudRate 
         Caption         =   "Baud Rate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label_PortaCom 
         Caption         =   "Porta COM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5880
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6480
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "LIMPAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   8280
      Width           =   7215
   End
   Begin VB.TextBox txtRecebido 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4200
      Width           =   7215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Recepção"
      Height          =   5775
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   7455
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "ENVIAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtEnviar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MaxLength       =   29
      TabIndex        =   0
      Top             =   1800
      Width           =   7215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transmissão"
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   7455
      Begin VB.CheckBox CheckAutoEnvio 
         Caption         =   "Envio automático"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox CheckLF 
         Caption         =   "LF"
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox CheckEnter 
         Caption         =   "Enter"
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox CheckCRLF 
         Caption         =   "CR/LF"
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox CheckCR 
         Caption         =   "CR"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   18
      ToolTipText     =   "Caracter para Clear Automatico da Recepção"
      Top             =   9840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "RX"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   16
      Top             =   9840
      Width           =   375
   End
   Begin VB.Shape ShapeRX 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   7080
      Top             =   9720
      Width           =   495
   End
   Begin VB.Menu mnuConectar 
      Caption         =   "Conectar"
   End
   Begin VB.Menu mnuDesconectar 
      Caption         =   "Desconectar"
   End
   Begin VB.Menu mnuAtualizar 
      Caption         =   "Atualizar"
      Begin VB.Menu mnuPortaCOM 
         Caption         =   "Porta COM"
      End
   End
   Begin VB.Menu mnuFormatarRecepcao 
      Caption         =   "Formatar"
      Begin VB.Menu mnuRecepcao 
         Caption         =   "Recepção"
         Begin VB.Menu mnuFonteRecepcao 
            Caption         =   "Fonte"
         End
         Begin VB.Menu mnuCorFonteRecepcao 
            Caption         =   "Cor Fonte"
         End
         Begin VB.Menu mnuCorEditorRecepcao 
            Caption         =   "Cor Editor"
         End
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public varBackEspace As Integer
Public varHex As String
Public conexao As Boolean



Private Sub Form_Load()

    'Busca configurações salvas de formatação
    txtRecebido.Font = ReadIniValue(App.Path & "\Config.ini", "FORMATAR", "txtRecebido.Font")
    txtRecebido.FontBold = ReadIniValue(App.Path & "\Config.ini", "FORMATAR", "txtRecebido.FontBold")
    txtRecebido.FontItalic = ReadIniValue(App.Path & "\Config.ini", "FORMATAR", "txtRecebido.FontItalic")
    txtRecebido.FontSize = ReadIniValue(App.Path & "\Config.ini", "FORMATAR", "txtRecebido.FontSize")
    txtRecebido.ForeColor = ReadIniValue(App.Path & "\Config.ini", "FORMATAR", "txtRecebido.ForeColor")
    txtRecebido.BackColor = ReadIniValue(App.Path & "\Config.ini", "FORMATAR", "txtRecebido.BackColor")
    
    'Configurações iniciais de conexão
    On Error GoTo Erro
        buscarConexao
        cboBaudRate.Text = "9600"
    Exit Sub

Erro:
    

End Sub

Private Sub buscarConexao()
    Dim I As Integer

    'Rotina para detectar portas COM disponíveis
    cboPortaCom.Clear
    For I = 1 To 32
        If DetectaPortaCOM(I) <> 0 Then
            cboPortaCom.AddItem "COM" & I
        End If
    Next
    
    On Error GoTo Erro
        cboPortaCom.ListIndex = 0
        If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
            mnuConectar.Enabled = True
            mnuDesconectar.Enabled = False
            Form1.Caption = "Conexão [Desconectado]"
        Else
            mnuConectar.Enabled = True
            mnuDesconectar.Enabled = False
            Form1.Caption = "Conexão [Desconectado]"
        End If
    Exit Sub
    
Erro:
    MsgBox "Erro " & Err & ". " & Error, vbApplicationModal, "Mensagem"
    Beep

End Sub

Private Sub cboBaudRate_Click()
    If conexao = True Then
    MSComm1.Settings = cboBaudRate.Text & "n,8,1"
    atualizaForm1
    End If
End Sub

Private Sub saveSettings()
    With MSComm1
        .CommPort = cboPortaCom.ListIndex + 1 'cboPortaCom.Text
        .Settings = cboBaudRate.Text & "n,8,1"
    End With
    
End Sub

Private Sub atualizaForm1()
    If conexao = True Then
        Form1.Caption = "Conexão [Conectado]" & "   " & "Porta " & "[" & cboPortaCom & "]" & "   " & "Baud " & "[" & cboBaudRate & "]"
    Else
        Form1.Caption = "Conexão [Desconectado]"
    End If
    
    If CheckPausa.Value = 1 Then
        Form1.Caption = "Conexão [Pausado]" & "   " & "Porta " & "[" & cboPortaCom & "]" & "   " & "Baud " & "[" & cboBaudRate & "]"
    End If
End Sub


Private Sub cmdEnviar_Click()

    If MSComm1.PortOpen = True Then
    
        If CheckCR.Value = 1 Then
            MSComm1.Output = txtEnviar.Text & vbCr
        ElseIf CheckLF.Value = 1 Then
            MSComm1.Output = txtEnviar.Text & vbLf
        ElseIf CheckCRLF.Value = 1 Then
            MSComm1.Output = txtEnviar.Text & vbCrLf
        ElseIf CheckEnter.Value = 1 Then
            MSComm1.Output = txtEnviar.Text & Chr$(13)
        Else
            MSComm1.Output = txtEnviar.Text
        End If
    
    End If

End Sub


Private Sub txtEnviar_change()
    Dim UltimaLetra As String
    
    If MSComm1.PortOpen = True Then
    
        If CheckAutoEnvio.Value = 1 Then
            UltimaLetra = Right(txtEnviar.Text, 1)
            
            If varBackEspace = 1 Then
                varBackEspace = 0
            Else
            
                If CheckCR.Value = 1 Then
                    MSComm1.Output = UltimaLetra & vbCr
                ElseIf CheckLF.Value = 1 Then
                    MSComm1.Output = UltimaLetra & vbLf
                ElseIf CheckCRLF.Value = 1 Then
                    MSComm1.Output = UltimaLetra & vbCrLf
                ElseIf CheckEnter.Value = 1 Then
                    MSComm1.Output = UltimaLetra & Chr$(13)
                Else
                        MSComm1.Output = UltimaLetra
                End If
            
            End If
        End If
    End If

End Sub

Private Sub txtEnviar_KeyPress(KeyAscii As Integer)

    If MSComm1.PortOpen = True Then
    
        If KeyAscii = vbKeyReturn Then
            MSComm1.Output = txtEnviar.Text
        End If
    
        If KeyAscii = 8 Then
            varBackEspace = 1
        End If
    
    End If

End Sub

Private Sub MSComm1_OnComm()
    Dim recebido As String

    recebido = MSComm1.Input
    
    If txtClear.Text <> "" Then
        If Left(recebido, 1) = txtClear.Text Then
            txtRecebido.Text = ""
        End If
    End If
    
    txtRecebido.Text = txtRecebido.Text + recebido
    txtRecebido.SelStart = Len(txtRecebido.Text)
    
    ShapeRX.BackColor = &HFF00&
    Timer1.Enabled = True
    
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

Private Sub Timer1_Timer()
    ShapeRX.BackColor = &H8000&
    Timer1.Enabled = False
End Sub

Private Sub cmdLimpar_Click()
    txtRecebido.Text = ""
End Sub

Private Sub CheckCR_Click()

    If CheckCR.Value = 1 Then
        CheckLF = 0
        CheckCRLF = 0
        CheckHex = 0
        CheckEnter = 0
    End If

End Sub

Private Sub CheckLF_Click()

    If CheckLF.Value = 1 Then
        CheckCR = 0
        CheckCRLF = 0
        CheckHex = 0
        CheckEnter = 0
    End If

End Sub

Private Sub CheckCRLF_Click()

    If CheckCRLF.Value = 1 Then
        CheckCR = 0
        CheckLF = 0
        CheckHex = 0
        CheckEnter = 0
    End If

End Sub

Private Sub CheckEnter_Click()

    If CheckEnter.Value = 1 Then
        CheckCR = 0
        CheckLF = 0
        CheckHex = 0
        CheckCRLF = 0
    End If

End Sub

Private Sub CheckAutoEnvio_Click()

    If CheckAutoEnvio.Value = 1 Then
        cmdEnviar.Enabled = False
    Else
        cmdEnviar.Enabled = True
    End If
    
End Sub

Private Sub CheckPausa_Click()
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
        mnuDesconectar.Enabled = False
        atualizaForm1
    Else
        MSComm1.PortOpen = True
        mnuDesconectar.Enabled = True
        atualizaForm1
    End If
End Sub

'Rotinas referente ao Menu
'--------------------------------------------------------------------------

Private Sub mnuConectar_Click()

    On Error GoTo Erro
        saveSettings
        MSComm1.PortOpen = True

        conexao = True
        CheckPausa.Enabled = True
        cboPortaCom.Enabled = False
        mnuAtualizar.Enabled = False
        mnuConectar.Enabled = False
        mnuDesconectar.Enabled = True
        atualizaForm1
    Exit Sub
    
Erro:
    MsgBox "Erro " & Err & ". " & Error, vbApplicationModal, "Mensagem"

End Sub

Private Sub mnuDesconectar_Click()

    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
        conexao = False
        CheckPausa.Enabled = False
        cboPortaCom.Enabled = True
        mnuAtualizar.Enabled = True
        mnuDesconectar.Enabled = False
        mnuConectar.Enabled = True
        atualizaForm1
    End If

End Sub
Private Sub mnuPortaCOM_Click()
    buscarConexao
End Sub

Private Sub mnuFonteRecepcao_Click()
    cd1.Flags = cd1CFBoth
    cd1.ShowFont
    txtRecebido.Font = cd1.FontName
    txtRecebido.FontBold = cd1.FontBold
    txtRecebido.FontItalic = cd1.FontItalic
    txtRecebido.FontSize = cd1.FontSize
    salvaFormatacao
End Sub

Private Sub mnuCorFonteRecepcao_Click()
    cd1.ShowColor
    txtRecebido.ForeColor = cd1.Color
    salvaFormatacao
End Sub

Private Sub mnuCorEditorRecepcao_Click()
    cd1.ShowColor
    txtRecebido.BackColor = cd1.Color
    salvaFormatacao
End Sub

Private Sub salvaFormatacao()
    WriteIniValue App.Path & "\Config.ini", "FORMATAR", "txtRecebido.Font", txtRecebido.Font
    WriteIniValue App.Path & "\Config.ini", "FORMATAR", "txtRecebido.FontBold", txtRecebido.FontBold
    WriteIniValue App.Path & "\Config.ini", "FORMATAR", "txtRecebido.FontItalic", txtRecebido.FontItalic
    WriteIniValue App.Path & "\Config.ini", "FORMATAR", "txtRecebido.FontSize", txtRecebido.FontSize
    WriteIniValue App.Path & "\Config.ini", "FORMATAR", "txtRecebido.ForeColor", txtRecebido.ForeColor
    WriteIniValue App.Path & "\Config.ini", "FORMATAR", "txtRecebido.BackColor", txtRecebido.BackColor
    'MsgBox "Dados salvo com sucesso...", , "Configuração"
End Sub

Private Sub mnuSair_Click()
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
        Unload Me
    Else
        Unload Me
    End If
End Sub


''--------------------------------------------------------------------------


