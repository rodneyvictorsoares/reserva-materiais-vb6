VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmGerenciarReservas 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEntregar 
      Caption         =   "Marcar como Entregue"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   6120
      Visible         =   0   'False
      Width           =   2200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5880
      TabIndex        =   6
      Top             =   1680
      Width           =   2200
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "Pesquisar Usuário"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   1680
      Width           =   2200
   End
   Begin VB.TextBox txtPesquisa 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      ToolTipText     =   "Digite o Id do Usuário"
      Top             =   960
      Width           =   1575
   End
   Begin MSComctlLib.ListView lstReservas 
      Height          =   3375
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5953
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Id"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Material"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Data Retirada"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Data Entrega"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Status"
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.Label lblUsuarioIdentity 
      AutoSize        =   -1  'True
      Caption         =   "-"
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
      Left            =   5400
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      Caption         =   "USUÁRIO: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3960
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label lblPesquisa 
      AutoSize        =   -1  'True
      Caption         =   "ID USUÁRIO: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1740
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   3840
      X2              =   12480
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   3840
      X2              =   12480
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   3840
      X2              =   12480
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Baixar Reservas"
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
      Width           =   2580
   End
End
Attribute VB_Name = "frmGerenciarReservas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    
    txtPesquisa.Text = ""
    txtPesquisa.Enabled = True
    cmdPesquisar.Enabled = True
    
    lblUsuario.Visible = False
    lblUsuarioIdentity.Visible = False
    lblUsuarioIdentity.Caption = "-"
    cmdCancelar.Enabled = False
    lstReservas.ListItems.Clear
    lstReservas.Enabled = False
    lstReservas.Visible = False
    cmdEntregar.Enabled = False
    cmdEntregar.Visible = False
    
    Me.Height = 2600
    
    txtPesquisa.SetFocus
    
End Sub

Private Sub cmdEntregar_Click()
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim entregues As Integer
    Dim reservaId As String
    Dim materialId As String
    Dim query As String
    
    entregues = 0
    
    AbrirConexao
    
    For i = 1 To lstReservas.ListItems.Count
        If lstReservas.ListItems(i).Selected Then
            reservaId = lstReservas.ListItems(i).Text
            
            query = "SELECT idMaterial FROM Reservas WHERE Id = " & reservaId
            Set rs = gConn.Execute(query)
            materialId = CStr(rs.Fields("idMaterial").Value)
            
            query = "UPDATE Reservas SET Status = 'Entregue' WHERE Id = " & reservaId
            gConn.Execute query
            
            query = "UPDATE Materiais SET Quantidade = Quantidade + 1 WHERE Id = " & materialId
            gConn.Execute query
                        
            lstReservas.ListItems(i).SubItems(4) = "Entregue"
            
            entregues = entregues + 1
        End If
    Next i
    
    If entregues > 0 Then MsgBox "Reservas selecionadas baixadas com sucesso!", vbInformation
        
    FecharConexao
    
End Sub

Private Sub cmdPesquisar_Click()
    Dim userId As String
    
    userId = Trim(txtPesquisa.Text)
    If userId = "" Then
        MsgBox "Por favor, insira o ID do usuário para consultar.", vbInformation
        txtPesquisa.SetFocus
        Exit Sub
    End If
    
    lstReservas.ListItems.Clear
    
    Call CarregarReservasUsuario(CLng(userId))
    
End Sub

Private Sub Form_Activate()
    txtPesquisa.SetFocus
End Sub

Private Sub Form_Load()
    'Me.Height = 2600
    
End Sub

Private Sub txtPesquisa_KeyPress(KeyAscii As Integer)
    ' Permite apenas números e a tecla Backspace
    If Len(txtPesquisa.Text) > 9 Then
        KeyAscii = 0
    End If
    
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub CarregarReservasUsuario(Optional ByVal userId As Long = 0)
    Dim rs As ADODB.Recordset
    Dim query As String
    Dim Item As ListItem

    AbrirConexao
    
    ' Consulta para selecionar as reservas
    query = "SELECT R.Id, U.Nome Usuario, M.Nome Material, R.datRetirada, R.datEntrega, R.Status FROM Reservas R " & _
            "INNER JOIN Materiais M ON R.idMaterial = M.Id INNER JOIN Usuarios U ON R.idUsuario = U.Id " & _
            "WHERE R.Status in ('Reservado','Atrasado') AND R.idUsuario = " & userId
    
    Set rs = gConn.Execute(query)
    
    If rs.EOF Then
        MsgBox "Não foram encontradas reservas para o usuário informado!", vbInformation
        Exit Sub
    End If
    
    txtPesquisa.Enabled = False
    cmdPesquisar.Enabled = False
    
    lblUsuario.Visible = True
    lblUsuarioIdentity.Visible = True
    lblUsuarioIdentity.Caption = rs.Fields("Usuario").Value
    cmdCancelar.Enabled = True
    lstReservas.Enabled = True
    lstReservas.Visible = True
    cmdEntregar.Enabled = True
    cmdEntregar.Visible = True
    
    Me.Height = 6800
    
    
    ' Preenche o ListView com as reservas encontradas
    Do While Not rs.EOF
        Set Item = lstReservas.ListItems.Add(, , CStr(rs.Fields("Id").Value))
        Item.SubItems(1) = rs.Fields("Material").Value
        Item.SubItems(2) = Format(rs.Fields("datRetirada").Value, "dd/MM/yyyy")
        Item.SubItems(3) = Format(rs.Fields("datEntrega").Value, "dd/MM/yyyy")
        Item.SubItems(4) = rs.Fields("Status").Value
        rs.MoveNext
    Loop
    
    FecharConexao
    
    lstReservas.SetFocus
End Sub


