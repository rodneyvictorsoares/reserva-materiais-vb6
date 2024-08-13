VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmUsuarios 
   BorderStyle     =   0  'None
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12600
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   12600
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7680
      TabIndex        =   14
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   495
      Left            =   2880
      TabIndex        =   13
      Top             =   5640
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   6720
      Visible         =   0   'False
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=school_db;Data Source=.\SQLEXPRESS"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=school_db;Data Source=.\SQLEXPRESS"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Usuarios"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   495
      Left            =   10080
      TabIndex        =   12
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdAtualizar 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   11
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdicionar 
      Caption         =   "Novo"
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   5640
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmUsuarios.frx":0000
      Height          =   1935
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   24
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboTipo 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      Height          =   420
      ItemData        =   "frmUsuarios.frx":0015
      Left            =   7200
      List            =   "frmUsuarios.frx":0017
      TabIndex        =   4
      Top             =   2640
      Width           =   5055
   End
   Begin VB.TextBox txtSenha 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2640
      Width           =   4575
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   11175
   End
   Begin VB.TextBox txtNome 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   11175
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   4080
      X2              =   12720
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   4080
      X2              =   12720
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      Caption         =   "Tipo: "
      Height          =   300
      Left            =   6600
      TabIndex        =   8
      Top             =   2640
      Width           =   570
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      Caption         =   "Senha: "
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   825
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "Email: "
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   705
   End
   Begin VB.Label lblNome 
      AutoSize        =   -1  'True
      Caption         =   "Nome: "
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   750
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   4080
      X2              =   12720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Cadastro de Usuários"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3480
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim originalNome As String
Dim originalEmail As String
Dim originalSenha As String
Dim originalTipo As String

Dim tempNome As String
Dim tempEmail As String
Dim tempSenha As String
Dim tempTipo As String

Private Sub cmdAdicionar_Click()
    txtNome.Text = ""
    txtEmail.Text = ""
    txtSenha.Text = ""
    cboTipo.ListIndex = 0
    
    LiberaControles
    cmdEditar.Enabled = False
    tempNome = ""
    tempEmail = ""
    tempSenha = ""
    tempTipo = ""

End Sub

Private Sub cmdAtualizar_Click()
On Error GoTo errmsg
    If tempNome = "" Then
        ' Adicionar novo registro
        tempNome = txtNome.Text
        tempEmail = txtEmail.Text
        tempSenha = txtSenha.Text
        tempTipo = cboTipo.Text
        
        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields("Nome").Value = tempNome
        Adodc1.Recordset.Fields("Email").Value = tempEmail
        Adodc1.Recordset.Fields("Senha").Value = tempSenha
        Adodc1.Recordset.Fields("Tipo").Value = tempTipo
    Else
        tempNome = txtNome.Text
        tempEmail = txtEmail.Text
        tempSenha = txtSenha.Text
        tempTipo = cboTipo.Text
        
        Adodc1.Recordset.Fields("Nome").Value = tempNome
        Adodc1.Recordset.Fields("Email").Value = tempEmail
        Adodc1.Recordset.Fields("Senha").Value = tempSenha
        Adodc1.Recordset.Fields("Tipo").Value = tempTipo
        Adodc1.Recordset.Update
    End If

    BloqueiaControles
    cmdEditar.Enabled = True
    MsgBox "Usuário atualizado com sucesso!"
    Exit Sub
errmsg:
    MsgBox "Erro ao modificar o registro de usuário..."
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo errmsg
    Dim response As Integer
    response = MsgBox("Tem certeza de que deseja cancelar?", vbYesNo + vbQuestion, "Confirmar Cancelamento")
    If response = vbYes Then
        If Not tempNome = "" Then
            txtNome.Text = originalNome
            txtEmail.Text = originalEmail
            txtSenha.Text = originalSenha
            cboTipo.Text = originalTipo
        End If
        BloqueiaControles
        cmdEditar.Enabled = True
    End If
    
    Exit Sub
errmsg:
    MsgBox "Erro ao cancelar a edição/adição..."
End Sub

Private Sub cmdEditar_Click()
    txtNome.Text = Adodc1.Recordset.Fields("Nome").Value
    txtEmail.Text = Adodc1.Recordset.Fields("Email").Value
    txtSenha.Text = Adodc1.Recordset.Fields("Senha").Value
    cboTipo.Text = Adodc1.Recordset.Fields("Tipo").Value
    
    originalNome = txtNome.Text
    originalEmail = txtEmail.Text
    originalSenha = txtSenha.Text
    originalTipo = cboTipo.Text
    
    tempNome = txtNome.Text
    tempEmail = txtEmail.Text
    tempSenha = txtSenha.Text
    tempTipo = cboTipo.Text
    
    LiberaControles
    cmdEditar.Enabled = False
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo errmsg
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MoveFirst
    BloqueiaControles
    cmdEditar.Enabled = True
    MsgBox "Usuário removido com sucesso!"
    Exit Sub
errmsg:
    MsgBox "Erro ao remover usuário..."
End Sub

Private Sub Form_Load()
    
    cboTipo.AddItem "Professor"
    cboTipo.AddItem "Administrador"
    cboTipo.ListIndex = 0
    
End Sub

Public Sub LiberaControles()
    txtNome.Enabled = True
    txtNome.BackColor = &H80000005
    
    txtEmail.Enabled = True
    txtEmail.BackColor = &H80000005
    
    txtSenha.Enabled = True
    txtSenha.BackColor = &H80000005
    
    cboTipo.Enabled = True
    cboTipo.BackColor = &H80000005
    
    cmdAtualizar.Enabled = True
    cmdCancelar.Enabled = True
    txtNome.SetFocus
    
End Sub

Public Sub BloqueiaControles()
    txtNome.Text = ""
    txtEmail.Text = ""
    txtSenha.Text = ""
    cboTipo.ListIndex = 0
    
    txtNome.Enabled = False
    txtNome.BackColor = &H80000016
    
    txtEmail.Enabled = False
    txtEmail.BackColor = &H80000016
    
    txtSenha.Enabled = False
    txtSenha.BackColor = &H80000016
    
    cboTipo.Enabled = False
    cboTipo.BackColor = &H80000016
    
    cmdAtualizar.Enabled = False
    cmdCancelar.Enabled = False
End Sub
