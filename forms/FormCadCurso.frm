VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormCadCurso 
   Caption         =   "Formul�rio de Curso"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15660
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   15660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextBoxData 
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   3375
      Width           =   1650
   End
   Begin VB.TextBox TextBoxHora 
      Height          =   450
      Left            =   6090
      TabIndex        =   7
      Top             =   3405
      Width           =   1980
   End
   Begin VB.TextBox codCurso 
      Height          =   510
      Left            =   930
      TabIndex        =   6
      Top             =   2310
      Width           =   8565
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
      Height          =   570
      Index           =   1
      Left            =   9840
      TabIndex        =   5
      Top             =   2235
      Width           =   1680
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   435
      Left            =   8835
      Top             =   7740
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   767
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
      Left            =   9765
      TabIndex        =   4
      Top             =   960
      Width           =   1680
   End
   Begin VB.TextBox txtHorario 
      Height          =   435
      Left            =   885
      TabIndex        =   3
      Top             =   3390
      Width           =   2265
   End
   Begin VB.TextBox txtNome 
      Height          =   540
      Left            =   945
      TabIndex        =   1
      Top             =   1035
      Width           =   8520
   End
   Begin VB.ListBox ListBox1 
      Columns         =   3
      Height          =   2790
      ItemData        =   "FormCadCurso.frx":0000
      Left            =   840
      List            =   "FormCadCurso.frx":0002
      TabIndex        =   12
      Top             =   4920
      Width           =   6615
   End
   Begin VB.Shape Shape1 
      Height          =   8250
      Left            =   480
      Top             =   315
      Width           =   12150
   End
   Begin VB.Label lblHora 
      Caption         =   "Hora:"
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
      Index           =   3
      Left            =   4020
      TabIndex        =   11
      Top             =   3000
      Width           =   1965
   End
   Begin VB.Label lblData 
      Caption         =   "Data:"
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
      Index           =   2
      Left            =   6120
      TabIndex        =   10
      Top             =   3030
      Width           =   1965
   End
   Begin VB.Label lblHorario 
      Caption         =   "Carga Hor�ria:"
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
      Index           =   1
      Left            =   945
      TabIndex        =   9
      Top             =   2940
      Width           =   1965
   End
   Begin VB.Label lblCod 
      Caption         =   "C�digo do Curso:"
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
      Left            =   915
      TabIndex        =   2
      Top             =   1845
      Width           =   1710
   End
   Begin VB.Label lblNome 
      Caption         =   "Nome do Curso:"
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
      Index           =   0
      Left            =   975
      TabIndex        =   0
      Top             =   660
      Width           =   1965
   End
End
Attribute VB_Name = "FormCadCurso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
   
   Call Conectar_BD
   
   On Error Resume Next
   'Copia Dados da tabela no servidor e lan�a  e lan�a na listBox
   
   TextBoxData.Enabled = False
   TextBoxHora.Enabled = False
   
   Me.TextBoxData = Format(Now, "DD/MM/YYYY")
   
   Dim Comando_Sql As String
   
   Dim idCurso As Integer
   Dim linhaListBox As Integer
   
     
   On Error Resume Next
   
   'Opera��o para copiar dados da tabela e lan�ar na listBox
   Set consulta = New ADODB.Recordset
   Comando_Sql = "SELECT * FROM sistema_ceuma.cursos" 'Pegando todos os dados da tabela especifica
   consulta.Open Comando_Sql, conexao, adOpenStatic, adLockReadOnly
   
   Me.ListBox1.Clear    'ListBox do frame
   
   'Adicionando dados ao ListBox do Form
   While Not consulta.EOF 'Realiza a consult at� o ultimo campo
      
      With Me.ListBox1
         .List
         .List(linhaListBox, 0) = consulta(0) 'Id
                             idCurso = consulta(0)
         .List(linhaListBox, 1) = consulta(1) 'codCurso
         .List(linhaListBox, 1) = consulta(2) 'carga horaria
         .List(linhaListBox, 2) = consulta(3) 'nome
      End With
         
      linhaListBox = linhaListBox + 1
   consula.MoveNext
   Wend
   
   consulta.Close          'Fechamento da consulta
   Set consulta = Nothing  'Limpa Banco de dados
   Call Desconectar_BD     'Desconectando do BD
End Sub

Private Sub btnSalvar1_Click(Index As Integer)
   Call Conectar_BD
   
   Dim Comando_Sql As String
   
   Dim codCurso As Integer
   Dim nome As String
   Dim horario As String
   Dim data As String
   
   codCurso = Me.codCurso
   
   nome = Me.txtNome
   horario = Me.txtHorario
  ' data = Me.txtData
   'data = Year(data) & "/" & Month(data) & "/" & Day(data) 'Convers�o de data para o formato de BD MYSQL ISO-8601
   
    
   '############Trabalhando com inser��o de dados na tabela####################
   
   'Adiciona dados a tabela
   
   
   Comando_Sql = "INSERT INTO sistema_ceuma.cursos(cod_curso, carga_horaria, nome, data_cad) VALUES ('" & codCurso & "', '" & horario & "', '" & nome & "', '" & data & "')"
   
   conexao.Execute Comando_Sql
   
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

Sub limpar_campos()

   Me.txtNome = ""
   Me.txtHorario = ""
   Me.txtData = ""
   
   Call UserForm_Initialize
End Sub

Private Sub ListBox1_Click()
   Dim valor_id As Integer
   Dim selecao As Integer
   
   selecao = ListBox1.ListIndex
   valor_id = ListBox1.List(selecao, 0)
   
   Me.txt_id = valor_id
   
   Call pesquisa
End Sub


