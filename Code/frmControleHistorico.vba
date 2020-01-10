Option Compare Database

Dim WithEvents mformulario As Form
    
'Private Sub codMed_Pro_Click()

'Dim val

'val = InputBox("Qual a quantidade em: " & codMed_Pro.Column(1), "Quantidade", 1)

'If codMed_Pro.Column(4) = 1 Then
'   ctl_QtdTotal = val * codMed_Pro.Column(3)
'ElseIf codMed_Pro.Column(4) = 2 Then
'   ctl_QtdTotal = val / codMed_Pro.Column(3)
'End If

'ctl_QtdTotal.SetFocus

'End Sub

Private Sub codProduto_Click()

'codMed_Pro.RowSource = "Select * From tblMedidasXProdutos Where codproduto = " & codProduto.Value

End Sub


Private Sub codProduto_DblClick(Cancel As Integer)

If codProduto.Text = "" Then
   Manipulacao "Novo"
   codProduto.RowSource = "SELECT tblProdutos.* FROM tblProdutos ORDER BY tblProdutos.pro_Descricao; "
   'codMed_Pro.RowSource = "Select * From tblMedidasXProdutos Where codproduto = " & codProduto.Value
Else
   Manipulacao "Alterar"
   codProduto.RowSource = "SELECT tblProdutos.* FROM tblProdutos ORDER BY tblProdutos.pro_Descricao; "
   'codMed_Pro.RowSource = "Select * From tblMedidasXProdutos Where codproduto = " & codProduto.Value
End If

End Sub


Private Sub codProduto_GotFocus()
codProduto.Dropdown
End Sub

Private Sub codProduto_NotInList(NewData As String, Response As Integer)

If FindCodigo(codProduto.Text) <> 0 Then
   
   codProduto.Value = FindCodigo(codProduto.Text)

Else

   Manipulacao "Novo"
   'codProduto.Value = ""
   
   codProduto.RowSource = "SELECT tblProdutos.* FROM tblProdutos ORDER BY tblProdutos.pro_Descricao;"
   'codMed_Pro.RowSource = "Select * From tblMedidasXProdutos Where codproduto = " & codProduto.Value
 

End If

End Sub

Private Sub Manipulacao(Operacao As String)
Set mformulario = New Form_frmProdutos

Select Case Operacao
 Case "Novo"
  ProDes = codProduto.Text
  With mformulario
   .Caption = "Novo Registro"
   .AllowDeletions = False
   .AllowAdditions = True
   .Visible = True
  End With
  DoCmd.GoToRecord , , acNewRec
 Case "Alterar"
  With mformulario
   .Caption = "Alteração de registro"
   .Filter = "codProduto = " & codProduto.Column(0)
   .FilterOn = True
   .AllowDeletions = False
   .AllowAdditions = False
   .Visible = True
  End With
End Select

End Sub

Private Sub ctl_QtdTotal_AfterUpdate()

If Not IsNull(codProduto.Column(0)) Then
 
 If Not Form_frmControleEstoque.codOperacao.Column(0) = 1 Then
    If ctl_QtdTotal > ChecarEstoqueEmpresa(codProduto.Column(0)) Then
       MsgBox "Atenção a quantidade informada é superior ao estoque que é: " & ChecarEstoqueEmpresa(codProduto.Column(0)), vbOKOnly + vbInformation, "Estoque baixo."
       'Me.ctl_QtdTotal.Value = ChecarEstoqueEmpresa(codProduto.Column(0))
    End If
 End If
 
End If

End Sub

Private Function ChecarEstoqueEmpresa(codProduto As Integer) As Long
Dim rstProdutos As DAO.Recordset
 Set rstProdutos = CurrentDb.OpenRecordset("Select * from tblEstoqueEmpresa where codproduto = " & codProduto)
 If rstProdutos.RecordCount > 0 Then ChecarEstoqueEmpresa = rstProdutos.Fields("Emp_QtdTotal")
End Function

Private Function FindCodigo(Descr As String) As Long
 Dim rstProdutos As DAO.Recordset
 Set rstProdutos = CurrentDb.OpenRecordset("Select * from tblProdutos Where pro_descricao = '" & Descr & "' Order BY codproduto DESC")
 If rstProdutos.RecordCount > 0 Then
    FindCodigo = rstProdutos.Fields("codproduto")
 Else
    FindCodigo = 0
 End If
End Function

Private Sub ctl_QtdTotal_Enter()

If Not IsNull(codProduto.Column(0)) Then
 
  If Me.ctl_VlUnitario.Value = 0 Then
     Me.ctl_VlUnitario.Value = codProduto.Column(3)
  End If
 
End If

End Sub
