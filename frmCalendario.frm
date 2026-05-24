VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendario 
   Caption         =   "Calendário"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4425
   OleObjectBlob   =   "frmCalendario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'====================================================================================
' frmCalendario
'====================================================================================
'
' DATEPICKER TOTALMENTE PORTÁTIL
'
' RECURSOS:
'   ? Todo código dentro do próprio frmCalendario
'   ? Inicializa SEMPRE no mês/ano atual
'   ? Retorno para:
'       - Planilha
'       - Outro UserForm
'   ? Formatos:
'       - CURTA   -> dd/mm/aaaa
'       - EXTENSA -> dd de mês de aaaa
'   ? Permite bloquear datas passadas
'   ? Destaque visual:
'       - finais de semana
'       - dia atual
'
' OBJETOS NECESSÁRIOS:
'   lblMesAno
'   cmdAnterior
'   cmdProximo
'   cmdDia1 até cmdDia42
'
'====================================================================================

Private DataAtual As Date
Private pFormatoSaida As String
Private pBloquearDatasPassadas As Boolean
Private pCelulaDestino As Range
Private pControleDestino As Object
Private pOrigemPlanilha As Boolean

'====================================================================================
' INICIALIZAÇÃO DO FORMULÁRIO
'====================================================================================
' IMPORTANTE:
'   Sempre inicia no mês/ano atual.
'====================================================================================

Private Sub UserForm_Initialize()
    DataAtual = Date
End Sub

'====================================================================================
' MÉTODO PÚBLICO
'====================================================================================
' ABRE O CALENDÁRIO PARA RETORNAR DATA À PLANILHA
'
' EXEMPLO:
'
' frmCalendario.AbrirParaPlanilha _
'       Range("A1"), _
'       "CURTA", _
'       True
'
'====================================================================================

Public Sub AbrirParaPlanilha( _
                ByVal Destino As Range, _
                Optional ByVal FormatoSaida As String = "CURTA", _
                Optional ByVal BloquearDatasPassadas As Boolean = False)

    Set pCelulaDestino = Destino

    Set pControleDestino = Nothing

    pOrigemPlanilha = True

    pFormatoSaida = UCase(Trim(FormatoSaida))

    pBloquearDatasPassadas = BloquearDatasPassadas

    '======================================================
    ' DEFINE DATA ATUAL
    '======================================================

    DataAtual = Date

    '======================================================
    ' CARREGA CALENDÁRIO
    '======================================================

    CarregarCalendario

    Me.Show

End Sub

'====================================================================================
' MÉTODO PÚBLICO
'====================================================================================
' ABRE O CALENDÁRIO PARA RETORNAR DATA A UM CONTROLE
'
' EXEMPLO:
'
' frmCalendario.AbrirParaFormulario _
'       Me.txtData, _
'       "EXTENSA", _
'       False
'
'====================================================================================

Public Sub AbrirParaFormulario( _
                ByVal ControleDestino As Object, _
                Optional ByVal FormatoSaida As String = "CURTA", _
                Optional ByVal BloquearDatasPassadas As Boolean = False)

    Set pControleDestino = ControleDestino

    Set pCelulaDestino = Nothing

    pOrigemPlanilha = False

    pFormatoSaida = UCase(Trim(FormatoSaida))

    pBloquearDatasPassadas = BloquearDatasPassadas

    '======================================================
    ' DEFINE DATA ATUAL
    '======================================================

    DataAtual = Date

    '======================================================
    ' AGORA O CALENDÁRIO É MONTADO
    ' JÁ COM TODOS OS PARÂMETROS DEFINIDOS
    '======================================================

    CarregarCalendario

    Me.Show

End Sub

'====================================================================================
' AVANÇA MÊS
'====================================================================================

Private Sub cmdProximo_Click()

    DataAtual = DateAdd("m", 1, DataAtual)

    CarregarCalendario

End Sub

'====================================================================================
' RETORNA MÊS
'====================================================================================

Private Sub cmdAnterior_Click()

    DataAtual = DateAdd("m", -1, DataAtual)

    CarregarCalendario

End Sub

'====================================================================================
' CARREGA CALENDÁRIO
'====================================================================================

Private Sub CarregarCalendario()

    Dim PrimeiroDia As Date

    Dim Inicio As Long

    Dim DiasMes As Long

    Dim i As Long

    Dim Botao As MSForms.CommandButton

    Dim DataBotao As Date

    Dim PosicaoBotao As Long

    '==========================================================
    ' GARANTE PRIMEIRO DIA DO MÊS
    '==========================================================

    PrimeiroDia = DateSerial(Year(DataAtual), Month(DataAtual), 1)

    '==========================================================
    ' TÍTULO
    '==========================================================

    lblMesAno.Caption = _
        UCase(Format(PrimeiroDia, "mmmm yyyy"))

    '==========================================================
    ' POSIÇÃO INICIAL
    '==========================================================

    Inicio = Weekday(PrimeiroDia, vbSunday)

    '==========================================================
    ' TOTAL DE DIAS DO MÊS
    '==========================================================

    DiasMes = Day(DateSerial(Year(PrimeiroDia), _
                             Month(PrimeiroDia) + 1, 0))

    '==========================================================
    ' LIMPA BOTÕES
    '==========================================================

    For i = 1 To 42

        Set Botao = Me.Controls("cmdDia" & i)

        With Botao

            .Caption = ""

            .Enabled = False

            .Tag = ""

            .BackColor = RGB(255, 255, 255)

            .ForeColor = RGB(0, 0, 0)

            .Font.Bold = False

        End With

    Next i

    '==========================================================
    ' PREENCHE DIAS
    '==========================================================

    For i = 1 To DiasMes

        PosicaoBotao = Inicio + i - 1

        Set Botao = Me.Controls("cmdDia" & PosicaoBotao)

        DataBotao = DateSerial(Year(DataAtual), Month(DataAtual), i)

        With Botao

            .Caption = i

            .Tag = DataBotao

            .Enabled = True

            '==================================================
            ' BLOQUEIO DE DATAS PASSADAS
            '==================================================

            If pBloquearDatasPassadas Then

                If DataBotao < Date Then

                    .Enabled = False

                End If

            End If

            '==================================================
            ' FINAIS DE SEMANA
            '==================================================

            If Weekday(DataBotao, vbSunday) = 1 _
            Or Weekday(DataBotao, vbSunday) = 7 Then

                .BackColor = RGB(188, 230, 254)

                .ForeColor = RGB(2, 32, 121)

                .Font.Bold = True

            End If

            '==================================================
            ' DIA ATUAL
            '==================================================

            If DataBotao = Date Then

                .BackColor = RGB(190, 248, 211)

                .ForeColor = RGB(242, 30, 30)

                .Font.Bold = True

            End If

        End With

    Next i

End Sub

'====================================================================================
' FORMATA DATA
'====================================================================================

Private Function FormatarData(ByVal DataValor As Date) As String

    Select Case pFormatoSaida

        Case "EXTENSA"

            FormatarData = _
                Day(DataValor) & _
                " de " & _
                LCase(Format(DataValor, "mmmm")) & _
                " de " & _
                Year(DataValor)

        Case Else

            FormatarData = Format(DataValor, "dd/mm/yyyy")

    End Select

End Function

'====================================================================================
' RETORNA DATA AO DESTINO
'====================================================================================

Private Sub RetornarData(ByVal DataEscolhida As Date)

    Dim TextoData As String

    TextoData = FormatarData(DataEscolhida)

    If pOrigemPlanilha Then

        If Not pCelulaDestino Is Nothing Then

            pCelulaDestino.Value = TextoData

        End If

    Else

        If Not pControleDestino Is Nothing Then

            pControleDestino.Value = TextoData

        End If

    End If

    Unload Me

End Sub

'====================================================================================
' ROTINA CENTRAL DE SELEÇÃO
'====================================================================================

Private Sub SelecionarData(btn As MSForms.CommandButton)

    Dim DataEscolhida As Date

    If btn.Caption = "" Then Exit Sub

    If btn.Enabled = False Then Exit Sub

    DataEscolhida = DateSerial( _
                    Year(DataAtual), _
                    Month(DataAtual), _
                    CLng(btn.Caption))

    RetornarData DataEscolhida

End Sub

'====================================================================================
' EVENTOS DOS BOTÕES
'====================================================================================

Private Sub cmdDia1_Click(): SelecionarData cmdDia1: End Sub
Private Sub cmdDia2_Click(): SelecionarData cmdDia2: End Sub
Private Sub cmdDia3_Click(): SelecionarData cmdDia3: End Sub
Private Sub cmdDia4_Click(): SelecionarData cmdDia4: End Sub
Private Sub cmdDia5_Click(): SelecionarData cmdDia5: End Sub
Private Sub cmdDia6_Click(): SelecionarData cmdDia6: End Sub
Private Sub cmdDia7_Click(): SelecionarData cmdDia7: End Sub
Private Sub cmdDia8_Click(): SelecionarData cmdDia8: End Sub
Private Sub cmdDia9_Click(): SelecionarData cmdDia9: End Sub
Private Sub cmdDia10_Click(): SelecionarData cmdDia10: End Sub
Private Sub cmdDia11_Click(): SelecionarData cmdDia11: End Sub
Private Sub cmdDia12_Click(): SelecionarData cmdDia12: End Sub
Private Sub cmdDia13_Click(): SelecionarData cmdDia13: End Sub
Private Sub cmdDia14_Click(): SelecionarData cmdDia14: End Sub
Private Sub cmdDia15_Click(): SelecionarData cmdDia15: End Sub
Private Sub cmdDia16_Click(): SelecionarData cmdDia16: End Sub
Private Sub cmdDia17_Click(): SelecionarData cmdDia17: End Sub
Private Sub cmdDia18_Click(): SelecionarData cmdDia18: End Sub
Private Sub cmdDia19_Click(): SelecionarData cmdDia19: End Sub
Private Sub cmdDia20_Click(): SelecionarData cmdDia20: End Sub
Private Sub cmdDia21_Click(): SelecionarData cmdDia21: End Sub
Private Sub cmdDia22_Click(): SelecionarData cmdDia22: End Sub
Private Sub cmdDia23_Click(): SelecionarData cmdDia23: End Sub
Private Sub cmdDia24_Click(): SelecionarData cmdDia24: End Sub
Private Sub cmdDia25_Click(): SelecionarData cmdDia25: End Sub
Private Sub cmdDia26_Click(): SelecionarData cmdDia26: End Sub
Private Sub cmdDia27_Click(): SelecionarData cmdDia27: End Sub
Private Sub cmdDia28_Click(): SelecionarData cmdDia28: End Sub
Private Sub cmdDia29_Click(): SelecionarData cmdDia29: End Sub
Private Sub cmdDia30_Click(): SelecionarData cmdDia30: End Sub
Private Sub cmdDia31_Click(): SelecionarData cmdDia31: End Sub
Private Sub cmdDia32_Click(): SelecionarData cmdDia32: End Sub
Private Sub cmdDia33_Click(): SelecionarData cmdDia33: End Sub
Private Sub cmdDia34_Click(): SelecionarData cmdDia34: End Sub
Private Sub cmdDia35_Click(): SelecionarData cmdDia35: End Sub
Private Sub cmdDia36_Click(): SelecionarData cmdDia36: End Sub
Private Sub cmdDia37_Click(): SelecionarData cmdDia37: End Sub
Private Sub cmdDia38_Click(): SelecionarData cmdDia38: End Sub
Private Sub cmdDia39_Click(): SelecionarData cmdDia39: End Sub
Private Sub cmdDia40_Click(): SelecionarData cmdDia40: End Sub
Private Sub cmdDia41_Click(): SelecionarData cmdDia41: End Sub
Private Sub cmdDia42_Click(): SelecionarData cmdDia42: End Sub

