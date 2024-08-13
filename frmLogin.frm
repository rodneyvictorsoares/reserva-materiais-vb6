VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Sistema de reservas de Materiais"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogin 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   4
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox txtSenha 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2280
      Width           =   4935
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Acesso ao Sistema"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1680
      TabIndex        =   5
      Top             =   360
      Width           =   3720
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      Caption         =   "Senha: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   825
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   780
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogin_Click()
    'Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim email As String
    Dim senha As String
    Dim query As String
    
    email = txtUsuario.Text
    senha = txtSenha.Text
    
'    Set conn = New ADODB.Connection
'    conn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=school_db;Data Source=.\SQLEXPRESS"
    
    AbrirConexao
    
    query = "select * from Usuarios WHERE Email = '" & email & "' AND Senha = '" & senha & "'"
    Set rs = gConn.Execute(query)
    
    If Not rs.EOF Then
        MsgBox "Acesso autorizado com sucesso", vbInformation
        gstrUsuarioLogado = rs!Nome
        gintIdUsuarioLogado = CInt(rs!Id)
        gstrTipoUsuario = rs!Tipo
        
        MDIPrincipal.Show
        Unload Me
    Else
        MsgBox "Email ou Senha inválidos", vbCritical
        LimpaCampos
    End If
    
    rs.Close
    FecharConexao
    Set rs = Nothing
    'Set conn = Nothing
    
End Sub

Private Sub LimpaCampos()
    txtUsuario.Text = ""
    txtSenha.Text = ""
    txtUsuario.SetFocus
End Sub
