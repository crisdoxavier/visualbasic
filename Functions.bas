Attribute VB_Name = "Functions"
Option Explicit
Option Base 1

Function fUltimaLinhaPlan(PlanRef As String) As Long

fUltimaLinhaPlan = ThisWorkbook.Worksheets(PlanRef).Range("A1048576").End(xlUp).Row

End Function

Public Function fUltimaColunaPlan(PlanRef As String) As Long

fUltimaColunaPlan = ThisWorkbook.Worksheets(PlanRef).Range("XFD1").End(xlToLeft).Column

End Function

Public Function fLinhaAtualPlan(PlanRef As String) As Long

fLinhaAtualPlan = ThisWorkbook.Worksheets(PlanRef).Range(ActiveCell.Address).Rows("1:1").Row

End Function

Public Function fColunaAtualPlan(PlanRef As String) As Long

fColunaAtualPlan = ThisWorkbook.Worksheets(PlanRef).Range(ActiveCell.Address).Columns("A:A").Column

End Function

Function fUltimaLinhaIntervalo(PlanRef As String, Coluna As String) As Long

fUltimaLinhaIntervalo = ThisWorkbook.Worksheets(PlanRef).Range(Coluna & "1048576").End(xlUp).Row

End Function

Sub AtualizaNomes(Nome As String, Planilha As String, CellInicial As String, UltimaColuna As String)

'N�o � fun��o propriamente dita, pois n�o retorna dados, mas est� aqui porque � como se fosse uma fun��o

Dim UltimaLinhaPlanRef As Long

UltimaLinhaPlanRef = fUltimaLinhaPlan(Planilha)
    
    ThisWorkbook.Names(Nome).Delete
    ThisWorkbook.Names.Add Name:=Nome, RefersTo:=Range(Planilha & "!" & CellInicial & ":" & UltimaColuna & UltimaLinhaPlanRef)

End Sub

Sub AtualizaNomesIntervalo(Nome As String, Planilha As String, CellInicial As String, Coluna As String, fColuna As String)
'N�o � fun��o propriamente dita, pois n�o retorna dados, mas est� aqui porque � como se fosse uma fun��o

Dim UltimaLinhaIntervalo As Long

'Caso o intervalo n�o tenha todas as linhas preenchidas, fColuna deve ser uma coluna sem c�lulas vazias
UltimaLinhaIntervalo = fUltimaLinhaIntervalo(Planilha, fColuna)
    
    ThisWorkbook.Names(Nome).Delete
    ThisWorkbook.Names.Add Name:=Nome, RefersTo:=Range(Planilha & "!" & CellInicial & ":" & Coluna & UltimaLinhaIntervalo)

End Sub
