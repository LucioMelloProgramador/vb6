VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artigos Populares (API TIMES)"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   8175
      Left            =   11640
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   11280
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      Caption         =   "VBJSON project at http://code.google.com/p/vba-json/"
      Height          =   255
      Index           =   1
      Left            =   7080
      TabIndex        =   4
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      Caption         =   "Lucio Mello"
      Height          =   255
      Left            =   7200
      TabIndex        =   3
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      Caption         =   "Visual Basic 6"
      Height          =   255
      Index           =   0
      Left            =   7080
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "The New York Times Developer Network"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Consumo da API ""Artigos Populares"""
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   2940
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variáveis para consumir api
Dim xhr, Method, url, Contents, FormatContent

'Para uso no scroll Form
Dim PosicaoAnterior As Integer

'Posição dos objetos e tamanho da tela
Dim posTop As Integer

'Array de urls
Dim urls() As String
   
'Criação dos componentes em tempo de execução
Private WithEvents lbl As Label
Attribute lbl.VB_VarHelpID = -1
Private pic As Image
Attribute pic.VB_VarHelpID = -1

'API do windows para download
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
        (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
         ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
         
'API do windows para execução de programas
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
         ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Download da imagem da url obtida no JSON
Public Function DownloadFile(url As String, LocalFilename As String) As Boolean
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, url, LocalFilename, 0, 0)
    If lngRetVal = 0 Then
        DownloadFile = True
    End If
End Function

'Abre o navegador com o link do artigo
Private Sub Command1_Click(index As Integer)
    ShellExecute hwnd, "open", urls(index), vbNullString, vbNullString, conSwNo
End Sub

'Monta a tela com a lista de artigos
'Cria os objets em tempo de execução para simular uma lista de itens
Private Sub MontarTela()
   
   On Error GoTo TrataErro
   
   Dim conteudo As String
   Dim url As String
   Dim urlImage As String
   Dim totalResults As Integer
   
   posTop = 1200
   If xhr.responseText <> "" Then
      Set p = JSON.parse(xhr.responseText)
      If Not (p Is Nothing) Then
         If JSON.GetParserErrors <> "" Then
            MsgBox JSON.GetParserErrors, vbInformation, "Parsing Error(s) occured"
         Else
            totalResults = p.Item("num_results")
            ReDim urls(totalResults)
            For i = 1 To totalResults
                url = p.Item("results").Item(i).Item("url")
                urls(i) = url
                'Faz o download do arquivo e mostra na tela
                If p.Item("results").Item(i).Item("media").Count > 0 Then
                    urlImage = p.Item("results").Item(i).Item("media").Item(1).Item("media-metadata").Item(1).Item("url")
                    DownloadFile urlImage, "C:\sistemas\VB6ConsomeAPI\pic.jpg"
                    Set pic = Controls.Add("VB.Image", "pic" & i)
                    With pic
                       .Visible = True
                       .Picture = LoadPicture("C:\sistemas\VB6ConsomeAPI\pic.jpg")
                       .Top = posTop
                       .Left = 100
                    End With
                    Kill "C:\sistemas\VB6ConsomeAPI\pic.jpg"
                End If
                'Cria o label de título do artigo
                Set lbl = Controls.Add("VB.Label", "lblTitle" & i)
                With lbl
                   .Visible = True
                   .Width = 10000
                   .Height = 400
                   .Caption = p.Item("results").Item(i).Item("title")
                   .Top = posTop
                   .Left = 1300
                   .Font = "ARIAL"
                   .BackColor = &H80000005
                   .FontSize = 12
                End With
                'Cria o label do abstract
                Set lbl = Controls.Add("VB.Label", "lblAbstract" & i)
                With lbl
                   .Visible = True
                   .Width = 10000
                   .Caption = p.Item("results").Item(i).Item("abstract")
                   .Top = posTop + 400
                   .Left = 1300
                   .Font = "ARIAL"
                   .BackColor = &H80000005
                End With
                'Cria o botão de link a partir do botão principal
                Load Command1(i)
                With Command1(i)
                   .Caption = "Link do artigo"
                   .Top = posTop
                   .Width = 1300
                   .Height = 400
                   .Left = 10000
                   .Visible = True
                End With
                posTop = posTop + 1500
            Next
         End If
      Else
         MsgBox "Ocorreu um erro ao ler JSON da API "
      End If
   End If

   Exit Sub
   
TrataErro:
    MsgBox "Ocorreu o seguinte erro ao obter os dados: " + Err.Description
    Resume Next
 
End Sub

Private Sub Form_Load()
   
    Set xhr = CreateObject("MSXML2.XMLHTTP")
    'Define o método HTTP de envio como GET
    Method = "GET"
    'Formato retornado pela API se usar outro, alterar aqui
    FormatContent = "application/json"
    'url da API
    url = "https://api.nytimes.com/svc/mostpopular/v2/emailed/1.json?api-key=dRI4qQlarq7FUOT5GFXvOOYPtLUp1nnR"
    xhr.Open Method, url, False
    xhr.send
    'Se houver algum erro, exibe a mensagem com a descrição do erro, se houver
    If xhr.Status < 200 Or xhr.Status >= 300 Then
        MsgBox "Erro HTTP:" & xhr.Status & " - Descrição: " & xhr.responseText
    End If
    
    MontarTela
    CriarScrollDaTela
    
End Sub
    
Private Sub CriarScrollDaTela()
    
    Dim MaxFormHeigth As Integer
    Dim MaxDisplayHeight  As Integer
    
    MaxFormHeigth = posTop + 200
    MaxDisplayHeight = 8670
    Me.Height = MaxDisplayHeight
    With VScroll1
        .Height = Me.ScaleHeight
        .Min = 0
        .Max = MaxFormHeigth - MaxDisplayHeight
        .SmallChange = Screen.TwipsPerPixelY * 10
        .LargeChange = .SmallChange
    End With

End Sub
    
Private Sub ScrollForm()
    
    Dim Objeto As Control
    
    For Each Objeto In Me.Controls
        If Not (TypeOf Objeto Is VScrollBar) And _
            Not (TypeOf Objeto Is Line) Then
                Objeto.Top = Objeto.Top + PosicaoAnterior - VScroll1.Value
        End If
        
        If (TypeOf Objeto Is Line) Then
            Objeto.Y1 = Objeto.Y1 + PosicaoAnterior - VScroll1.Value
            Objeto.Y2 = Objeto.Y2 + PosicaoAnterior - VScroll1.Value
        End If
    Next
    PosicaoAnterior = VScroll1.Value

End Sub

Private Sub VScroll1_Change()
    Call ScrollForm
End Sub

Private Sub VScroll1_Scroll()
    Call ScrollForm
End Sub
