VERSION 5.00
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H80000008&
   Caption         =   "Sistema de Reserva Materiais"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   12660
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastro"
      Begin VB.Menu mnuUsuarios 
         Caption         =   "&Usuários"
      End
      Begin VB.Menu mnuMateriais 
         Caption         =   "&Materiais"
      End
   End
   Begin VB.Menu mnuReserva 
      Caption         =   "&Reserva"
      Begin VB.Menu mnuNovaReserva 
         Caption         =   "&Nova Reserva"
      End
      Begin VB.Menu mnuExibirReserva 
         Caption         =   "&Exibir Reserva"
      End
      Begin VB.Menu mnuBaixarReserva 
         Caption         =   "&Baixar Reserva"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public frmUsuariosInstance As frmUsuarios
Public frmMateriaisInstance As frmMateriais

Private Sub MDIForm_Load()
    Call VerificaPerfil
End Sub

Private Sub mnuBaixarReserva_Click()
    CloseAllChildForms
    Set frmGerenciarReservas = New frmGerenciarReservas
    CenterChildForm frmGerenciarReservas
    frmGerenciarReservas.Show
    
End Sub

Private Sub mnuExibirReserva_Click()
    CloseAllChildForms
    Set frmExibirReservas = New frmExibirReservas
    CenterChildForm frmExibirReservas
    frmExibirReservas.Show
    
End Sub

Private Sub mnuMateriais_Click()
    CloseAllChildForms
    Set frmMateriaisInstance = New frmMateriais
    CenterChildForm frmMateriaisInstance
    frmMateriaisInstance.Show
    
End Sub

Private Sub mnuNovaReserva_Click()
    CloseAllChildForms
    Set frmReservas = New frmReservas
    CenterChildForm frmReservas
    frmReservas.Show
End Sub

Private Sub mnuSair_Click()
    Unload Me
    
End Sub

Private Sub mnuUsuarios_Click()
    CloseAllChildForms
    Set frmUsuariosInstance = New frmUsuarios
    CenterChildForm frmUsuariosInstance
    frmUsuariosInstance.Show
    
End Sub

Private Sub CloseAllChildForms()
    If Not frmUsuariosInstance Is Nothing Then
        Unload frmUsuariosInstance
        Set frmUsuariosInstance = Nothing
    End If
    
    ' Verificar se frmMateriaisInstance está aberto e fechar
    If Not frmMateriaisInstance Is Nothing Then
        Unload frmMateriaisInstance
        Set frmMateriaisInstance = Nothing
    End If
End Sub

Private Sub CenterChildForm(frm As Form)
    ' Calcular a posição para centralizar o formulário filho
    frm.Left = (Me.ScaleWidth - frm.Width) / 2
    frm.Top = (Me.ScaleHeight - frm.Height) / 2
End Sub

Private Sub VerificaPerfil()
    If gstrTipoUsuario = "Administrador" Then
        mnuCadastro.Visible = True
        mnuBaixarReserva.Visible = True
    Else
        mnuCadastro.Visible = False
        mnuBaixarReserva.Visible = False
    End If
End Sub
