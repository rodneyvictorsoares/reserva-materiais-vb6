
# 📋 Projeto Reserva de Materiais

[![Visual Basic](https://img.shields.io/badge/Visual%20Basic-6.0-blue)](https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6)
[![SQL Server](https://img.shields.io/badge/SQL%20Server-2019-red)](https://www.microsoft.com/en-us/sql-server/sql-server-2019)

## 📝 Sobre o Projeto

A aplicação **Reserva de Materiais**! Este é um projeto desenvolvido em Visual Basic 6 para gerenciar de maneira básica a reserva e emprestimo de materiais acadêmicos de uma escola, permitindo o controle de recursos, usuários e reservas. Este pequeno sistema foi criado com o objetivo praticar os conhecimentos de VB6 e Sql Server.

## 🚀 Tecnologias Utilizadas

- **Visual Basic 6**: Linguagem de programação clássica amplamente usada em sistemas legado de todos os tipos.  ![Visual Basic](https://img.shields.io/badge/Visual%20Basic-6.0-blue)
- **SQL Server**: Banco de dados relacional robusto e facilmente integrável com o VB6  ![SQL Server](https://img.shields.io/badge/SQL%20Server-2019-red)

## 📋 Funcionalidades

- **Login de Usuário:**
  - Sistema de autenticação para restringir o acesso ao sistema.
  - Diferentes níveis de acesso para usuários, permitindo uma gestão segura e personalizada.

- **Gestão de Reservas:**
  - Exibir reservas existentes.
  - Gerenciar reservas, permitindo adicionar, editar ou cancelar reservas.
  - Funcionalidade para visualizar detalhes das reservas realizadas.

- **Gestão de Materiais:**
  - Cadastro de novos materiais disponíveis para reserva.
  - Edição e remoção de materiais.
  - Visualização dos materiais disponíveis no sistema.

- **Gestão de Usuários:**
  - Cadastro de novos usuários.
  - Edição de informações dos usuários.
  - Controle de permissões e acessos dos usuários.

## 📦 Estrutura do Projeto

```bash
ReservaMateriais/
│
├── database.sql               # Script SQL para criação do banco de dados
├── ReservaMateriais.vbp        # Arquivo de projeto do Visual Basic 6
├── Form1.frm                   # Formulário principal
├── frmExibirReservas.frm       # Formulário para exibir reservas
├── frmGerenciarReservas.frm    # Formulário para gerenciar reservas
├── frmLogin.frm                # Formulário de login de usuário
├── frmMateriais.frm            # Formulário de gestão de materiais
├── frmReservas.frm             # Formulário para realizar reservas
├── frmUsuarios.frm             # Formulário de gestão de usuários
├── MDIPrincipal.frm            # Formulário principal MDI (Multiple Document Interface)
└── mdlGlobal.bas               # Módulo global com funções e variáveis compartilhadas
```

## 🚀 Instalação

Para rodar o projeto, siga os seguintes passos:

1. **Configurar o Banco de Dados:**
   - Importe o arquivo `database.sql` em seu banco de dados SQL Server.
   - Configure as conexões no código fonte para apontar para seu banco de dados.

2. **Abrir o Projeto:**
   - Utilize o Visual Basic 6 para abrir o arquivo `ReservaMateriais.vbp`.

3. **Executar:**
   - Compile e execute o projeto através do Visual Basic 6.

## 🖥️ Pré-requisitos

- **Visual Basic 6:** Necessário para compilar e rodar o projeto.
- **Banco de Dados:** Servidor SQL Server compatível para rodar o script `database.sql`.

## 🤝 Contribuindo

Contribuições são bem-vindas! Para contribuir:

1. Faça um fork deste repositório.
2. Crie uma branch com sua feature: `git checkout -b minha-feature`.
3. Commit suas mudanças: `git commit -m 'Adiciona minha nova feature'`.
4. Push para a branch: `git push origin minha-feature`.
5. Abra um Pull Request.

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

---

**Autor**: [Rodney Victor](https://github.com/rodneyvictorsoares)

