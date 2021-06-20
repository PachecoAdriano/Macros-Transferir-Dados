Attribute VB_Name = "Módulo1"
Sub previdencia()

Dim Dia         As Worksheet
Dim Mov         As Worksheet
Dim Socio       As Worksheet
Dim Cliente     As String
Dim Valor       As Currency
Dim Sol         As String
Dim Mes         As String
Dim Ano         As String
Dim Op          As String
Dim Regime      As String
Dim Proposta    As String
Dim Plano       As String


Set Dia = Sheets("Movimentações DIA")
Set Mov = Sheets("MOVIMENTAÇÕES PREVIDÊNCIA")


Application.ScreenUpdating = False

Dia.Select
Range("c1048576").Select
ActiveCell.End(xlUp).Select

Do While ActiveCell.Row >= 3

    Proposta = ActiveCell
    Cliente = Trim(ActiveCell.Offset(0, 28))
    Ano = Left(ActiveCell.Offset(0, 8).Value, 4)
    Mes = Mid(ActiveCell.Offset(0, 8).Value, 5, 2)
    Sol = Right(ActiveCell.Offset(0, 8).Value, 2)
    Regime = ActiveCell.Offset(0, 30)
    Valor = Replace(ActiveCell.Offset(0, 10), ".", ",")
    Op = ActiveCell.Offset(0, -1)
    Plano = ActiveCell.Offset(0, 29)
    
    Mov.Select
    
    Range("A1048576").Select
    ActiveCell.End(xlUp).Offset(1, 0).Select
    
    With ActiveCell
        .Value = Cliente
        .Offset(0, 3).Value = Sol & "/" & Mes & "/" & Ano
        .Offset(0, 4).Value = Valor
        .Offset(0, 5).Value = Plano
        .Offset(0, 6).Value = Regime
        .Offset(0, 7).Value = Proposta
        .Offset(0, 9).Value = Op
    End With

    Dia.Activate
    ActiveCell.Offset(-1, 0).Select

Loop

Application.ScreenUpdating = True

Mov.Select

MsgBox "Você é um fofo!"

Range("a1048576").Select
ActiveCell.End(xlUp).Offset(1, 0).Select


End Sub


