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

Private DataAtual As Date

Public FormularioOrigem As Object
Public PlanilhaDestino As Worksheet
Public CelulaDestino As Range
' ==========================================
' INICIALIZAÇÃO DO FORMULÁRIO
' ==========================================
Private Sub UserForm_Initialize()
    DataAtual = Date
    CarregarCalendario
End Sub
' ==========================================
' AVANÇAR MESES
' ==========================================
Private Sub cmdProximo_Click()
    DataAtual = DateAdd("m", 1, DataAtual)
    CarregarCalendario
End Sub
' ==========================================
' RETROCEDER MESES
' ==========================================
Private Sub cmdAnterior_Click()
    DataAtual = DateAdd("m", -1, DataAtual)
    CarregarCalendario
End Sub
' ==========================================
' SELEÇÃO DE DATAS | ROTINA DE SELEÇÃO
' ==========================================
Private Sub SelecionarData(btn As MSForms.CommandButton)

    Dim DataEscolhida As Date

    If btn.Caption = "" Then Exit Sub

    DataEscolhida = DateSerial(Year(DataAtual), Month(DataAtual), CInt(btn.Caption))

    ' =====================================
    ' RETORNO PARA USERFORM
    ' =====================================
    If Not FormularioOrigem Is Nothing Then

        If TextBoxValida Then
            FormularioOrigem.Controls("txt_DatePickerValue").Value = Format(DataEscolhida, "dd/mm/yyyy")
        End If

    End If

    ' =====================================
    ' RETORNO PARA PLANILHA
    ' =====================================
    If Not PlanilhaDestino Is Nothing Then

        If Not CelulaDestino Is Nothing Then

            CelulaDestino.Value = DataEscolhida
            CelulaDestino.NumberFormat = "dd/mm/yyyy"

        End If

    End If

    Unload Me

End Sub
Private Sub cmdDia1_Click()
    SelecionarData cmdDia1
End Sub
Private Sub cmdDia2_Click()
    SelecionarData cmdDia2
End Sub
Private Sub cmdDia3_Click()
    SelecionarData cmdDia3
End Sub
Private Sub cmdDia4_Click()
    SelecionarData cmdDia4
End Sub
Private Sub cmdDia5_Click()
    SelecionarData cmdDia5
End Sub
Private Sub cmdDia6_Click()
    SelecionarData cmdDia6
End Sub
Private Sub cmdDia7_Click()
    SelecionarData cmdDia7
End Sub
Private Sub cmdDia8_Click()
    SelecionarData cmdDia8
End Sub
Private Sub cmdDia9_Click()
    SelecionarData cmdDia9
End Sub
Private Sub cmdDia10_Click()
    SelecionarData cmdDia10
End Sub
Private Sub cmdDia11_Click()
    SelecionarData cmdDia11
End Sub
Private Sub cmdDia12_Click()
    SelecionarData cmdDia12
End Sub
Private Sub cmdDia13_Click()
    SelecionarData cmdDia13
End Sub
Private Sub cmdDia14_Click()
    SelecionarData cmdDia14
End Sub
Private Sub cmdDia15_Click()
    SelecionarData cmdDia15
End Sub
Private Sub cmdDia16_Click()
    SelecionarData cmdDia16
End Sub
Private Sub cmdDia17_Click()
    SelecionarData cmdDia17
End Sub
Private Sub cmdDia18_Click()
    SelecionarData cmdDia18
End Sub
Private Sub cmdDia19_Click()
    SelecionarData cmdDia19
End Sub
Private Sub cmdDia20_Click()
    SelecionarData cmdDia20
End Sub
Private Sub cmdDia21_Click()
    SelecionarData cmdDia21
End Sub
Private Sub cmdDia22_Click()
    SelecionarData cmdDia22
End Sub
Private Sub cmdDia23_Click()
    SelecionarData cmdDia23
End Sub
Private Sub cmdDia24_Click()
    SelecionarData cmdDia24
End Sub
Private Sub cmdDia25_Click()
    SelecionarData cmdDia25
End Sub
Private Sub cmdDia26_Click()
    SelecionarData cmdDia26
End Sub
Private Sub cmdDia27_Click()
    SelecionarData cmdDia27
End Sub
Private Sub cmdDia28_Click()
    SelecionarData cmdDia28
End Sub
Private Sub cmdDia29_Click()
    SelecionarData cmdDia29
End Sub
Private Sub cmdDia30_Click()
    SelecionarData cmdDia30
End Sub
Private Sub cmdDia31_Click()
    SelecionarData cmdDia31
End Sub
Private Sub cmdDia32_Click()
    SelecionarData cmdDia32
End Sub
Private Sub cmdDia33_Click()
    SelecionarData cmdDia33
End Sub
Private Sub cmdDia34_Click()
    SelecionarData cmdDia34
End Sub
Private Sub cmdDia35_Click()
    SelecionarData cmdDia35
End Sub
Private Sub cmdDia36_Click()
    SelecionarData cmdDia36
End Sub
Private Sub cmdDia37_Click()
    SelecionarData cmdDia37
End Sub
Private Sub cmdDia38_Click()
    SelecionarData cmdDia38
End Sub
Private Sub cmdDia39_Click()
    SelecionarData cmdDia39
End Sub
Private Sub cmdDia40_Click()
    SelecionarData cmdDia40
End Sub
Private Sub cmdDia41_Click()
    SelecionarData cmdDia41
End Sub
Private Sub cmdDia42_Click()
    SelecionarData cmdDia42
End Sub
' ==========================================
' CARREGA DADOS INICIAIS NO CALENDÁRIO
' ==========================================
Private Sub CarregarCalendario()

    Dim PrimeiroDia As Date
    Dim Inicio As Integer
    Dim DiasMes As Integer
    Dim i As Integer
    Dim Botao As MSForms.CommandButton
    Dim DataSelecionada As Date

    lblMesAno.Caption = Format(DataAtual, "mmmm yyyy")

    PrimeiroDia = DateSerial(Year(DataAtual), Month(DataAtual), 1)

    Inicio = Weekday(PrimeiroDia, vbSunday)

    DiasMes = Day(DateSerial(Year(DataAtual), Month(DataAtual) + 1, 0))

    ' Limpa os botões
    For i = 1 To 42

        Set Botao = Me.Controls("cmdDia" & i)

        Botao.Caption = ""
        Botao.Enabled = False
        Botao.BackColor = vbButtonFace
        Botao.ForeColor = vbBlack

    Next i

    ' Preenche os dias
    For i = 1 To DiasMes

        Set Botao = Me.Controls("cmdDia" & (Inicio + i - 1))

        Botao.Caption = i
        Botao.Enabled = True

        DataSelecionada = DateSerial(Year(DataAtual), Month(DataAtual), i)

        ' DESTACAR HOJE
        If DataSelecionada = Date Then
            Botao.BackColor = vbYellow
        End If

        ' FINAIS DE SEMANA EM VERMELHO
        If Weekday(DataSelecionada, vbSunday) = 1 _
        Or Weekday(DataSelecionada, vbSunday) = 7 Then

            Botao.ForeColor = vbRed

        End If

        ' BLOQUEAR DATAS PASSADAS
        If DataSelecionada < Date Then
            Botao.Enabled = False
            Botao.ForeColor = RGB(150, 150, 150)
        End If

    Next i

End Sub
' ==========================================
' CRIA VALIDAÇÃO DA TEXTBOX
' ==========================================
Private Function TextBoxValida() As Boolean

    Dim ctrl As Control

    On Error GoTo TrataErro

    Set ctrl = FormularioOrigem.Controls("txt_DatePickerValue")

    If ctrl Is Nothing Then

        MsgBox "O formulário chamador não possui a TextBox obrigatória:" _
                & vbCrLf & _
                "txt_DatePickerValue", vbCritical

        TextBoxValida = False
        Exit Function

    End If

    TextBoxValida = True
    Exit Function

TrataErro:

    MsgBox "Erro ao localizar a TextBox 'txt_datepicked'." & vbCrLf & _
           "Verifique se o controle existe no formulário.", vbCritical

    TextBoxValida = False

End Function


' ==========================================
' MACRO PARA ABRIR CALENDÁRIO NA PLANILHA
' ==========================================
Sub AbrirCalendario()

    Set frmCalendario.CelulaDestino = ActiveCell

    frmCalendario.Show

End Sub
' ==========================================
' MÉTODO PARA ABRIR NO USERFORM
' ==========================================
Public Sub AbrirNoFormulario(frm As Object)

    Set FormularioOrigem = frm

    Set PlanilhaDestino = Nothing
    Set CelulaDestino = Nothing

    Me.Show

End Sub
' ==========================================
' MÉTODO PARA ABRIR NA PLANILHA
' ==========================================
Public Sub AbrirNaPlanilha(ws As Worksheet, cel As Range)

    Set PlanilhaDestino = ws
    Set CelulaDestino = cel

    Set FormularioOrigem = Nothing

    Me.Show

End Sub

' ==========================================
' ==========================================
' ==========================================
'COMO USAR NO USERFORM
'No formulário chamador:
'O formulário precisa possuir uma TextBox com: (Name) => txt_DatePickerValue

'Private Sub txtData_Enter()
'    frmCalendario.AbrirNoFormulario Me
'End Sub

' ==========================================
' ==========================================
' ==========================================
'COMO USAR NA PLANILHA
'Exemplo:
'Sub AbrirCalendarioNaCelula()
'    frmCalendario.AbrirNaPlanilha _
'        ThisWorkbook.Sheets("Agenda"), _
'        ThisWorkbook.Sheets("Agenda").Range("B5")
'End Sub
