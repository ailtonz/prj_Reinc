Option Compare Database
Option Explicit

Dim WithEvents mformulario As Form


Private Sub cmdFiltrar_Click()

Main 2

End Sub

Private Sub Form_Load()

Main 1

End Sub

Sub Main(Tipo)

Me.Caption = strTitulo
Me.lstCadastro.ColumnWidths = strColunas
Me.lstCadastro.ColumnHeads = True

If strTabela = "tblControleEstoque" Then
   Me.lstCadastro.RowSource = "qryControleEstoque01"
   Me.lstCadastro.ColumnWidths = "0cm;2cm;6cm;6cm;0cm"
ElseIf strTabela = "tblNotasFiscais" Then
   Me.lstCadastro.RowSource = "qryNotasFiscais01"
   Me.lstCadastro.ColumnWidths = "2cm;8cm;3cm;0cm;0cm"
ElseIf strTabela = "tblProdutos" Then
   If Tipo = 1 Then
      Me.lstCadastro.RowSource = "SELECT tblProdutos.codProduto, tblProdutos.pro_Descricao, tblProdutos.codProduto, tblEstoqueEmpresa.Emp_QtdTotal, tblProdutos.pro_VlUnitario, [Emp_QtdTotal]*[pro_VlUnitario] AS Total " & _
                                 "FROM tblProdutos LEFT JOIN tblEstoqueEmpresa ON tblProdutos.codProduto = tblEstoqueEmpresa.codProduto " & _
                                 "ORDER BY tblProdutos.pro_Descricao;"
   ElseIf Tipo = 2 Then
      Me.lstCadastro.RowSource = "SELECT tblProdutos.codProduto, tblProdutos.pro_Descricao, tblProdutos.codProduto, tblEstoqueEmpresa.Emp_QtdTotal, tblProdutos.pro_VlUnitario, [Emp_QtdTotal]*[pro_VlUnitario] AS Total " & _
                                 "FROM tblProdutos LEFT JOIN tblEstoqueEmpresa ON tblProdutos.codProduto = tblEstoqueEmpresa.codProduto " & _
                                 "WHERE pro_descricao like '*" & LCase(Trim(txtProcura)) & "*' " & _
                                 "ORDER BY tblProdutos.pro_Descricao;"
   End If
   Me.lstCadastro.ColumnWidths = "0;4cm;2cm;2cm;2cm;2cm"
   Me.header1.Visible = True
   
ElseIf strTabela = "tblBoletos" Then
   Me.lstCadastro.RowSource = "qryBoletos"
   Me.lstCadastro.ColumnWidths = "0cm;2cm;6cm;2cm;1cm"

Else
   Me.lstCadastro.RowSource = "Select * from " & strTabela & " order by " & strOrdem
End If


End Sub

Private Sub lstCadastro_DblClick(Cancel As Integer)
cmdAlterar_Click
End Sub

Private Sub cmdNovo_Click()
Manipulacao strTabela, "Novo"
End Sub

Private Sub cmdAlterar_Click()
Manipulacao strTabela, "Alterar"
End Sub

Private Sub cmdExcluir_Click()
Manipulacao strTabela, "Excluir"
End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click

    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    MsgBox Err.Description
    Resume Exit_cmdFechar_Click
    
End Sub

Private Sub Manipulacao(tabela As String, Operacao As String)

If IsNull(lstCadastro.Value) And Operacao <> "Novo" Then
   Exit Sub
End If

'Formulario
Select Case tabela
 Case "tblCadastros"
  Set mformulario = New Form_frmCadastro
 
 Case "tblProdutos"
  Set mformulario = New Form_frmProdutos
  
 Case "tblControleEstoque"
  Set mformulario = New Form_frmControleEstoque
 
 Case "tblFiscal"
  Set mformulario = New Form_frmFiscal
 
 Case "tblNotasFiscais"
  Set mformulario = New Form_frmNotasFiscais
  
Case "tblBoletos"
  Set mformulario = New Form_frmCadBoletos

End Select


Select Case Operacao

 Case "Novo"
  With mformulario
   .Caption = "Novo Registro"
   .AllowDeletions = False
   .AllowAdditions = True
   .Visible = True
  End With
  DoCmd.GoToRecord , , acNewRec
  
 Case "Alterar"
  With mformulario
   .Caption = "Alteração de resgistro"
   .Filter = strIdentificacao & " = " & lstCadastro.Value
   .FilterOn = True
   .AllowDeletions = False
   .AllowAdditions = False
   .Visible = True
  End With
  
 Case "Excluir"
 
  If MsgBox("Deseja excluir este registro?", vbInformation + vbOKCancel) = vbOK Then
     DoCmd.SetWarnings False
     DoCmd.RunSQL ("Delete from " & strTabela & " where " & strIdentificacao & " = " & lstCadastro)
     DoCmd.SetWarnings True
  End If
  
End Select
lstCadastro.Requery

Saida:
End Sub

