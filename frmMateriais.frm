VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmMateriais 
   BorderStyle     =   0  'None
   Caption         =   "Cadastro de UsuáriosCadastro de UsuáriosCadastro de UsuáriosCadastro de UsuáriosCadastro de Usuários"
   ClientHeight    =   6420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
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
      Left            =   9840
      TabIndex        =   14
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
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
      Left            =   7560
      TabIndex        =   13
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Enabled         =   0   'False
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
      Left            =   5160
      TabIndex        =   12
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
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
      Left            =   2760
      TabIndex        =   11
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdicionar 
      Caption         =   "Novo"
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
      Left            =   480
      TabIndex        =   10
      Top             =   5760
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
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
      RecordSource    =   "Materiais"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmMateriais.frx":0000
      Height          =   3015
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   24
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin VB.TextBox txtQuantidade 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   420
      Left            =   8520
      TabIndex        =   8
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox txtDescricao 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   420
      Left            =   1440
      TabIndex        =   6
      Top             =   1560
      Width           =   5295
   End
   Begin VB.ComboBox cboTipo 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8520
      TabIndex        =   4
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox txtNome 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   420
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   5295
   End
   Begin VB.Label lblQuantidade 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade: "
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
      Left            =   7080
      TabIndex        =   7
      Top             =   1680
      Width           =   1365
   End
   Begin VB.Label lblDescricao 
      AutoSize        =   -1  'True
      Caption         =   "Descrição: "
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
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1185
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      Caption         =   "Tipo: "
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
      Left            =   7080
      TabIndex        =   3
      Top             =   960
      Width           =   570
   End
   Begin VB.Label lblNome 
      AutoSize        =   -1  'True
      Caption         =   "Nome: "
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
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   750
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Cadastro de Materiais"
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
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   4080
      X2              =   12720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   4080
      X2              =   12720
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   4080
      X2              =   12720
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "frmMateriais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim originalNome As String
Dim originalQuantidade As String
Dim originalDescricao As String
Dim originalTipo As String

Dim tempNome As String
Dim tempQuantidade As String
Dim tempDescricao As String
Dim tempTipo As String

Private Sub cmdAdicionar_Click()
    txtNome.Text = ""
    cboTipo.ListIndex = 0
    txtDescricao.Text = ""
    txtQuantidade.Text = ""
    
    LiberaControles
    cmdEditar.Enabled = False
    tempNome = ""
    tempTipo = ""
    tempDescricao = ""
    tempQuantidade = ""
    
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo errmsg
    Dim response As Integer
    response = MsgBox("Tem certeza de que deseja cancelar?", vbYesNo + vbQuestion, "Confirmar Cancelamento")
    If response = vbYes Then
        BloqueiaControles
        cmdEditar.Enabled = True
    End If
    
    Exit Sub
errmsg:
    MsgBox "Erro ao cancelar a edição/adição..."
End Sub

Private Sub cmdEditar_Click()
    txtNome.Text = Adodc1.Recordset.Fields("Nome").Value
    cboTipo.Text = Adodc1.Recordset.Fields("Tipo").Value
    txtDescricao.Text = Adodc1.Recordset.Fields("Descricao").Value
    txtQuantidade.Text = Adodc1.Recordset.Fields("Quantidade").Value
    
    originalNome = txtNome.Text
    originalTipo = cboTipo.Text
    originalDescricao = txtDescricao.Text
    originalQuantidade = txtQuantidade.Text
    
    tempNome = txtNome.Text
    tempTipo = cboTipo.Text
    tempDescricao = txtDescricao.Text
    tempQuantidade = txtQuantidade.Text
        
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

Private Sub cmdSalvar_Click()
On Error GoTo errmsg
    If tempNome = "" Then
        ' Adicionar novo registro
        tempNome = txtNome.Text
        tempTipo = cboTipo.Text
        tempDescricao = txtDescricao.Text
        tempQuantidade = txtQuantidade.Text
        
        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields("Nome").Value = tempNome
        Adodc1.Recordset.Fields("Tipo").Value = tempTipo
        Adodc1.Recordset.Fields("Quantidade").Value = tempQuantidade
        Adodc1.Recordset.Fields("Descricao").Value = tempDescricao
    Else
        tempNome = txtNome.Text
        tempTipo = cboTipo.Text
        tempDescricao = txtDescricao.Text
        tempQuantidade = txtQuantidade.Text
        
        Adodc1.Recordset.Fields("Nome").Value = tempNome
        Adodc1.Recordset.Fields("Tipo").Value = tempTipo
        Adodc1.Recordset.Fields("Quantidade").Value = tempQuantidade
        Adodc1.Recordset.Fields("Descricao").Value = tempDescricao
        Adodc1.Recordset.Update
    End If

    BloqueiaControles
    cmdEditar.Enabled = True
    MsgBox "Material atualizado com sucesso!"
    Exit Sub
errmsg:
    MsgBox "Erro ao modificar o registro de materiais..."
End Sub

Private Sub Form_Load()
    cboTipo.AddItem ""
    cboTipo.AddItem "Projetor"
    cboTipo.AddItem "Laptop"
    cboTipo.AddItem "Desktop"
    cboTipo.AddItem "Cabo"
    cboTipo.AddItem "Adaptador"
    cboTipo.AddItem "Outro"
    cboTipo.ListIndex = 0
End Sub

Public Sub LiberaControles()
    txtNome.Enabled = True
    txtNome.BackColor = &H80000005
    
    cboTipo.Enabled = True
    cboTipo.BackColor = &H80000005
    
    txtDescricao.Enabled = True
    txtDescricao.BackColor = &H80000005
    
    txtQuantidade.Enabled = True
    txtQuantidade.BackColor = &H80000005
    
    cmdSalvar.Enabled = True
    cmdCancelar.Enabled = True
    txtNome.SetFocus
    
End Sub

Public Sub BloqueiaControles()
    txtNome.Text = ""
    cboTipo.ListIndex = 0
    txtDescricao.Text = ""
    txtQuantidade.Text = ""
    
    txtNome.Enabled = False
    txtNome.BackColor = &H80000016
    
    cboTipo.Enabled = False
    cboTipo.BackColor = &H80000016
    
    txtDescricao.Enabled = False
    txtDescricao.BackColor = &H80000016
    
    txtQuantidade.Enabled = False
    txtQuantidade.BackColor = &H80000016
    
    cmdSalvar.Enabled = False
    cmdCancelar.Enabled = False
    
End Sub
