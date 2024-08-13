VERSION 5.00
Begin VB.Form frmReservas 
   BorderStyle     =   0  'None
   Caption         =   "Reserva de Materiais"
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDatEntrega 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      ToolTipText     =   "Digite a data de entrega (dd/MM/yyyy)"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtDatRetirada 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "Digite a data de retirada (dd/MM/yyyy)"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.ComboBox cboMateriais 
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
      Left            =   1680
      TabIndex        =   3
      Top             =   2280
      Width           =   7455
   End
   Begin VB.CommandButton cmdReserva 
      Caption         =   "Efetuar Reserva"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   3120
      Width           =   4095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Meterial: "
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
      Left            =   360
      TabIndex        =   9
      Top             =   2400
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data Entrega: "
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
      Left            =   4920
      TabIndex        =   8
      Top             =   1680
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data Retirada: "
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
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   1860
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
      Left            =   1560
      TabIndex        =   6
      Top             =   960
      Width           =   75
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
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   1080
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
      Caption         =   "Reserva de Materiais"
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
      Width           =   3345
   End
End
Attribute VB_Name = "frmReservas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReserva_Click()
Dim strMaterial As String
Dim dtReserva As Date
Dim dtEntrega As Date

    If Not IsDate(txtDatRetirada.Text) Or Not IsDate(txtDatEntrega.Text) Then
        MsgBox "Favor inserir datas válidas no formato (dd/MM/yyyy)", vbExclamation
        txtDatRetirada.SetFocus
        Exit Sub
    End If
    
    If cboMateriais.ListIndex = -1 Then
        MsgBox "Por favor, selecione um material a ser reservado.", vbExclamation
        cboMateriais.SetFocus
        Exit Sub
    End If
    
    dtReserva = CDate(txtDatRetirada.Text)
    dtEntrega = CDate(txtDatEntrega.Text)
    
    If dtReserva < Date Then
        MsgBox "A data de retirada não pode ser inferior a data atual", vbExclamation
        txtDatRetirada.SetFocus
        Exit Sub
    End If
    
    If dtEntrega <= dtReserva Then
        MsgBox "A data de entrega deve ser posterior a data de retirada.", vbExclamation
        txtDatEntrega.SetFocus
        Exit Sub
    End If
    
    If DateDiff("d", dtReserva, dtEntrega) > 15 Then
        MsgBox "Tempo máximo de entrega é de 15 dias.", vbExclamation
        txtDatEntrega.SetFocus
        Exit Sub
    End If
    
    strMaterial = cboMateriais.Text
    
    EfetuarReserva strMaterial, dtReserva, dtEntrega
    
End Sub

Private Sub Form_Activate()
    txtDatRetirada.SetFocus
End Sub

Private Sub Form_Load()
 Dim rs As ADODB.Recordset
 Dim query As String
 
 lblUsuarioIdentity.Caption = gstrUsuarioLogado
 
 AbrirConexao
    
    query = "SELECT * FROM Materiais WHERE Quantidade > 0"
    Set rs = gConn.Execute(query)
    
    Do While Not rs.EOF
        cboMateriais.AddItem rs!Id & " - " & rs!Nome
        rs.MoveNext
    Loop
    
    rs.Close
    FecharConexao
    Set rs = Nothing
    
    
    
End Sub

Private Sub EfetuarReserva(ByVal strMaterial As String, ByVal dtReserva As Date, ByVal dtEntrega As Date)
Dim query As String
Dim intIdMaterial As Integer
    
On Error GoTo ErroSQL

    intIdMaterial = CInt(Trim(Left(strMaterial, (InStr(strMaterial, "-")))))
    
    If intIdMaterial < 0 Then intIdMaterial = intIdMaterial * -1
        
    If Not VerificarPermissao(intIdMaterial, dtReserva) Then
        MsgBox "Usuário já possui materiais reservados no período.", vbInformation
        Exit Sub
    End If
    
    AbrirConexao
    
    query = "INSERT INTO Reservas (idUsuario, idMaterial, datRetirada, datEntrega, status) VALUES (" & gintIdUsuarioLogado & "," & intIdMaterial & ",'" & dtReserva & "','" & dtEntrega & "', 'reservado')"
    gConn.Execute query
    
    query = "UPDATE Materiais SET Quantidade = Quantidade - 1 WHERE Id = " & intIdMaterial
    gConn.Execute query
    
    FecharConexao
    
    MsgBox "Reserva efetuada com sucesso", vbInformation
    Unload Me
    
    Exit Sub
    
ErroSQL:
    FecharConexao
    MsgBox "Não foi possível realizar a reserva.", vbCritical
End Sub


Private Function VerificarPermissao(ByVal idMaterial As Integer, ByVal dtReserva As Date) As Boolean
Dim rs As ADODB.Recordset
Dim query As String

    AbrirConexao
    
    query = "SELECT * FROM Reservas WHERE idUsuario = " & gintIdUsuarioLogado & " and idMaterial = " & idMaterial & " and datEntrega <= '" & dtReserva & "'"
    Set rs = gConn.Execute(query)
    
    If rs.EOF Then
        VerificarPermissao = True
    Else
        VerificarPermissao = False
    End If
    
    FecharConexao
    
End Function

Private Sub AplicarMascaraData(ByRef txtBox As TextBox, ByRef KeyAscii As Integer)
    ' Permite apenas números e a tecla Backspace
    If Len(txtBox.Text) > 9 Then
        KeyAscii = 0
    End If
    
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
    
    ' Aplica a máscara enquanto o usuário digita
    If Len(txtBox.Text) = 2 Or Len(txtBox.Text) = 5 Then
        txtBox.Text = txtBox.Text & "/"
        txtBox.SelStart = Len(txtBox.Text)
    End If
End Sub

Private Sub txtDatEntrega_KeyPress(KeyAscii As Integer)
    Call AplicarMascaraData(txtDatEntrega, KeyAscii)
End Sub

Private Sub txtDatRetirada_KeyPress(KeyAscii As Integer)
    Call AplicarMascaraData(txtDatRetirada, KeyAscii)
End Sub
