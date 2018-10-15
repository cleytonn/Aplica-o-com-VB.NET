VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormCadAltCurso 
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15825
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   15825
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   615
      TabIndex        =   15
      Top             =   5325
      Width           =   5160
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10245
      Top             =   4335
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   480
      Left            =   11595
      TabIndex        =   14
      Top             =   6150
      Width           =   795
   End
   Begin VB.TextBox txt_id 
      Height          =   480
      Left            =   2715
      TabIndex        =   12
      Top             =   3090
      Width           =   570
   End
   Begin VB.CommandButton btn_excluir 
      Caption         =   "Excluir"
      Height          =   465
      Left            =   10470
      TabIndex        =   11
      Top             =   6195
      Width           =   630
   End
   Begin VB.CommandButton btn_editar 
      Caption         =   "Editar"
      Height          =   495
      Left            =   9495
      MouseIcon       =   "FormCadAlteracao.frx":0000
      Picture         =   "FormCadAlteracao.frx":5A5A2
      TabIndex        =   10
      Top             =   6195
      Width           =   630
   End
   Begin VB.CommandButton btnVoltar 
      Caption         =   "Voltar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   1
      Left            =   4665
      TabIndex        =   9
      Top             =   4020
      Width           =   1680
   End
   Begin VB.TextBox TextBoxData 
      Height          =   495
      Left            =   9900
      TabIndex        =   8
      Top             =   840
      Width           =   1650
   End
   Begin VB.TextBox TextBoxHora 
      Height          =   570
      Left            =   9900
      TabIndex        =   7
      Top             =   2070
      Width           =   1980
   End
   Begin VB.CommandButton btnSalvar1 
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   0
      Left            =   6810
      TabIndex        =   6
      Top             =   3975
      Width           =   1680
   End
   Begin VB.TextBox txt_nome 
      Height          =   540
      Left            =   645
      TabIndex        =   2
      Top             =   825
      Width           =   8520
   End
   Begin VB.TextBox txt_horario 
      Height          =   435
      Left            =   585
      TabIndex        =   1
      Top             =   3180
      Width           =   1485
   End
   Begin VB.TextBox cod_curso 
      Height          =   510
      Left            =   630
      TabIndex        =   0
      Top             =   2100
      Width           =   8565
   End
   Begin VB.Label LblId 
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2925
      TabIndex        =   13
      Top             =   2760
      Width           =   465
   End
   Begin VB.Shape Shape1 
      Height          =   8595
      Left            =   165
      Top             =   240
      Width           =   14265
   End
   Begin VB.Label lblNome 
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   675
      TabIndex        =   5
      Top             =   450
      Width           =   1965
   End
   Begin VB.Label lblHorario 
      Caption         =   "Carga horária"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   645
      TabIndex        =   4
      Top             =   2700
      Width           =   1785
   End
   Begin VB.Label lblCod 
      Caption         =   "Código do Curso"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   615
      TabIndex        =   3
      Top             =   1635
      Width           =   1710
   End
End
Attribute VB_Name = "FormCadAltCurso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   Me.btn_editar.Enabled = False
   Me.btn_excluir.Enabled = False
   
   Me.TextBoxData = Format(Now, "DD/MM/YYYY")
      
   Dim id As Integer
   Dim linhalistbox As Integer
   Dim comando_Sql As String
   
   Call Conectar_BD
   
   On Error Resume Next
   
   'Operação para copiar dados da tabela e lançar na listBox
   Set consulta = New ADODB.Recordset
   comando_Sql = "SELECT * FROM sistema_ceuma.cursos" 'Pegando todos os dados da tabela especifica
   consulta.Open comando_Sql, conexao, adOpenStatic, adLockReadOnly
   
   Me.List1.Clear    'ListBox do frame
   
   'Adicionando dados ao ListBox do Form
   While Not consulta.EOF 'Realiza a consult até o ultimo campo
      
      With Me.List1
      .AddItem
      .List(linhalistbox, 0) = consulta(0) 'id
                          id = consulta(0)
      .List(linhalistbox, 1) = consulta(1) 'codCurso
      .List(linhalistbox, 1) = consulta(2) 'carga horaria
      .List(linhalistbox, 2) = consulta(3) 'nome
      End With
         
      linhalistbox = linhalistbox + 1
   consula.MoveNext
   Wend
   
   consulta.Close          'Fechamento da consulta
   Set consulta = Nothing  'Limpa Banco de dados
   Call Desconectar_BD     'Desconectando do BD
End Sub


Private Sub ListBox1_Click()
   Dim valor_id As Integer
   Dim selecao As Integer
   
   selecao = ListBox1.ListIndex
   valoer_id = ListBox1.List(selecao, 0)
   
   Me.txt_id = valor_id
   
   Call pesquisa
End Sub

Private Sub btn_editar_Click()
   Dim id As Integer
   Dim comando_Sql As String
   
   If txt_id = "" Then
      Exit Sub
   Else
      id = txt_id
   End If
   
   Call Conectar_BD
   
   Set consulta = New ADODB.Recordset
   comando_Sql = "SELECT * FROM sistema_ceuma.cursos id like '" & id & "' "
   consulta.Open comando_Sql, conexao, , adLockOptimistic
   
   On Error Resume Next
   'Exibe nos campos do formulário, o conteúdo de cada campo encontrado na consulta
   
   consulta(1) = Me.txt_nome
   consulta(2) = Me.txt_nome
   consulta(3) = Me.txt_horario
   
   'Atualiza o BD
   consulta.Update
   
   'Exibe mensagem de sucesso na alteração de dados
   MsgBox "Registro Alterado com Sucesso!", vbDefaultButton1, "Alteração"
   
   'Chama a rotina que libera as variaveis de objeto do BD
   Call Desconectar_BD
   
   Call limpar_campos
   
End Sub

Private Sub btnSalvar1_Click(Index As Integer)
   Call Conectar_BD
   
   Dim comando_Sql As String
   
   Dim codCurso As Integer
   Dim nome As String
   Dim horario As String
   Dim data As String
   
   codCurso = Me.cod_curso
   
   nome = Me.txt_nome
   horario = Me.txt_horario
  ' data = Me.txtData
   'data = Year(data) & "/" & Month(data) & "/" & Day(data) 'Conversão de data para o formato de BD MYSQL ISO-8601
   
    
   '############Trabalhando com inserção de dados na tabela####################
   
   'Adiciona dados a tabela
   
   
   comando_Sql = "INSERT INTO sistema_ceuma.cursos(cod_curso, carga_horaria, nome, data_cad) VALUES ('" & codCurso & "', '" & horario & "', '" & nome & "', '" & data & "')"
   
   conexao.Execute comando_Sql
   
   Call Desconectar_BD
   
   MsgBox "Dados inseridos com sucesso"
   
   Call Desconectar_BD
      
   'FormPrincipal.Show
   'Unload Me
End Sub

Private Sub btnVoltar_Click(Index As Integer)
   FormPrincipal.Show
   Unload Me
End Sub

Private Sub pesquisa()
   Dim id As Integer
   Dim comando_Sql As String
   
   If txt_id = "" Then
      Exit Sub
   Else
      id = txt_id
   End If
   
   Call Conectar_BD
   
   Set consulta = New ADODB.Recordset
   comando_Sql = "SELECT * FROM sistema_ceuma.cursos id like '" & id & "' "
   consulta.Open comando_Sql, conexao, , adLockOptimistic
   
   On Error Resume Next
   'Exibe nos campos do formulário, o conteúdo de cada campo encontrado na consulta
   
   consulta(1) = Me.txt_nome
   consulta(2) = Me.txt_nome
   consulta(3) = Me.txt_horario
   
   'Habilita os botões editar e excluir, pois estará em um registro existente
   Me.btn_editar.Enabled = True
   Me.btn_excluir.Enabled = True
   
   'Desabilita o botão salvar para evitar duplicação de dados
   Me.btnSalvar1.Enable = True
   
   Call Desconectar_BD
End Sub
