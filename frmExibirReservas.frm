VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmExibirReservas 
   BorderStyle     =   0  'None
   Caption         =   "Reservas Realizadas"
   ClientHeight    =   6660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lstReservas 
      Height          =   4455
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7858
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Materiais"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Data Retirada"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Data Entrega"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      Caption         =   "Usuário: "
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
      TabIndex        =   2
      Top             =   960
      Width           =   1080
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
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   75
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Exibir Reservas"
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
      Width           =   2475
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   3840
      X2              =   12480
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   3840
      X2              =   12480
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   3840
      X2              =   12480
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "frmExibirReservas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CarregaReservasUsuario()
    Dim rs As ADODB.Recordset
    Dim query As String
    Dim Item As ListItem
    
    AbrirConexao
    
    query = "SELECT R.Id ID, M.Nome Material, R.datRetirada datRetirada, R.datEntrega datEntrega, R.Status FROM Reservas R INNER JOIN Materiais M ON R.idMaterial = M.Id WHERE R.idUsuario =" & gintIdUsuarioLogado
    Set rs = gConn.Execute(query)
    
    If rs.EOF Then
        MsgBox "Não foram encontradas reservas para o usuário!", vbInformation
        Exit Sub
    End If
        
    Do While Not rs.EOF
        Set Item = lstReservas.ListItems.Add(, , CStr(rs.Fields("Id").Value))
        Item.SubItems(1) = rs.Fields("Material").Value
        Item.SubItems(2) = Format(rs.Fields("datRetirada").Value, "dd/MM/yyyy")
        Item.SubItems(3) = Format(rs.Fields("datEntrega").Value, "dd/MM/yyyy")
        Item.SubItems(4) = rs.Fields("Status").Value
        rs.MoveNext
    Loop
    
    FecharConexao
    
End Sub

Private Sub Form_Load()
    lblUsuarioIdentity.Caption = gstrUsuarioLogado
    lstReservas.View = lvwReport
    
    Call CarregaReservasUsuario
End Sub
