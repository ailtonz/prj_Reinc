Option Compare Database
Option Explicit

Private Sub cmdEfetivar_Click()

If IsNull(codOperacao) Then
   Exit Sub
Else
   If IsNull(codCadastro) And (Me.codOperacao <> 5) Then
      Exit Sub
   End If
End If

DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

Dim rstCadastro As DAO.Recordset
Dim rstControle As DAO.Recordset
Dim rstHistorico As DAO.Recordset
Dim rstHistorico2 As DAO.Recordset
Dim rstEstoqueEmpresa As DAO.Recordset
Dim rstEstoqueCliente As DAO.Recordset
Dim rstProdutos As DAO.Recordset
Dim x As Integer
Dim Fornec As Boolean


Set rstControle = CurrentDb.OpenRecordset("Select * from tblControleEstoque where codControle = " & Me.txtcodControle)
Set rstHistorico = CurrentDb.OpenRecordset("Select * from tblControleHistorico where codControle = " & Me.txtcodControle)


If Me.codOperacao = 5 Then
   Fornec = False
Else
   Set rstCadastro = CurrentDb.OpenRecordset("Select * from tblCadastros where codCadastro = " & Me.codCadastro.Column(0))
   Fornec = (rstCadastro.Fields("codCategoria") = "F")
   rstCadastro.Close
End If


rstControle.Edit
rstControle.Fields("OK") = True
rstControle.Update

rstHistorico.MoveLast
rstHistorico.MoveFirst

If Fornec Or Me.codOperacao = 5 Then

 Do While Not rstHistorico.EOF
 
  If Me.codOperacao = 1 Then
     Set rstHistorico2 = CurrentDb.OpenRecordset("Select * from tblControleHistorico where codProduto = " & rstHistorico.Fields("codProduto") & " and ctl_vlunitario = 0;")
  End If
  
  Set rstProdutos = CurrentDb.OpenRecordset("Select * from tblprodutos where codProduto = " & rstHistorico.Fields("codProduto"))
  If Not rstProdutos.EOF Then
     rstProdutos.Edit
     rstProdutos.Fields("pro_vlunitario") = rstHistorico.Fields("ctl_vlunitario")
     rstProdutos.Update
  End If
  If Me.codOperacao = 1 Then
     Do While Not rstHistorico2.EOF
        rstHistorico2.Edit
        rstHistorico2.Fields("ctl_vlunitario") = rstHistorico.Fields("ctl_vlunitario")
        rstHistorico2.Update
        rstHistorico2.MoveNext
     Loop
  End If
  
  Set rstEstoqueEmpresa = CurrentDb.OpenRecordset("Select * from tblEstoqueEmpresa where codProduto = " & rstHistorico.Fields("codProduto"))
  If rstEstoqueEmpresa.RecordCount > 0 Then
     rstEstoqueEmpresa.Edit
  Else
     rstEstoqueEmpresa.AddNew
  End If
  
  If rstControle.Fields("codOperacao") = 1 Or rstControle.Fields("codOperacao") = 5 Then
     rstEstoqueEmpresa.Fields("codProduto") = rstHistorico.Fields("codProduto")
     rstEstoqueEmpresa.Fields("Emp_QtdTotal") = rstEstoqueEmpresa.Fields("Emp_QtdTotal") + rstHistorico.Fields("ctl_QtdTotal")
  Else
     rstEstoqueEmpresa.Fields("Emp_QtdTotal") = rstEstoqueEmpresa.Fields("Emp_QtdTotal") - rstHistorico.Fields("ctl_QtdTotal")
  End If
  
  rstEstoqueEmpresa.Update
  rstHistorico.MoveNext
  
 Loop
 If Me.codOperacao = 1 Then
    rstHistorico2.Close
 End If
 
 rstEstoqueEmpresa.Close
 rstProdutos.Close
 
Else
 
 If rstControle.Fields("codOperacao") = 4 Then
    Set rstEstoqueCliente = CurrentDb.OpenRecordset("Select * from tblEstoqueClientes where codCadastro = " & Me.codCadastro.Column(0))
    Do While Not rstEstoqueCliente.EOF
        rstEstoqueCliente.Edit
        rstEstoqueCliente.Fields("cli_QtdTotal") = 0
        rstEstoqueCliente.Update
        rstEstoqueCliente.MoveNext
    Loop
    rstEstoqueCliente.Close
 End If
 
 Do While Not rstHistorico.EOF
 
  Set rstEstoqueEmpresa = CurrentDb.OpenRecordset("Select * from tblEstoqueEmpresa where codProduto = " & rstHistorico.Fields("codProduto"))
  Set rstEstoqueCliente = CurrentDb.OpenRecordset("Select * from tblEstoqueClientes where codProduto = " & rstHistorico.Fields("codProduto") & " and codCadastro = " & Me.codCadastro.Column(0))
  If rstEstoqueCliente.RecordCount > 0 Then
     rstEstoqueCliente.Edit
  Else
     rstEstoqueCliente.AddNew
     rstEstoqueCliente.Fields("codProduto") = rstHistorico.Fields("codProduto")
     rstEstoqueCliente.Fields("codCadastro") = Me.codCadastro.Column(0)
  End If
  
  If Not rstEstoqueEmpresa.EOF Then
     rstEstoqueEmpresa.Edit
  Else
     rstEstoqueEmpresa.AddNew
     rstEstoqueEmpresa.Fields("codProduto") = rstHistorico.Fields("codProduto")
  End If
  
  If rstControle.Fields("codOperacao") = 2 Then 'ENTRADA no Cliente SAIDA na Empresa
     
     rstEstoqueCliente.Fields("cli_QtdTotal") = rstEstoqueCliente.Fields("cli_QtdTotal") + rstHistorico.Fields("ctl_QtdTotal")
     rstEstoqueEmpresa.Fields("Emp_QtdTotal") = rstEstoqueEmpresa.Fields("Emp_QtdTotal") - rstHistorico.Fields("ctl_QtdTotal")
  
  ElseIf rstControle.Fields("codOperacao") = 3 Then 'Devolução
     
     rstEstoqueCliente.Fields("cli_QtdTotal") = rstEstoqueCliente.Fields("cli_QtdTotal") - rstHistorico.Fields("ctl_QtdTotal")
     rstEstoqueEmpresa.Fields("Emp_QtdTotal") = rstEstoqueEmpresa.Fields("Emp_QtdTotal") + rstHistorico.Fields("ctl_QtdTotal")
  
  ElseIf rstControle.Fields("codOperacao") = 4 Then 'Fechamento

     'rstEstoqueCliente.Fields("cli_QtdTotal") = rstEstoqueCliente.Fields("cli_QtdTotal") - rstHistorico.Fields("ctl_QtdTotal")
     rstEstoqueCliente.Fields("cli_QtdTotal") = rstEstoqueCliente.Fields("cli_QtdTotal") + rstHistorico.Fields("ctl_QtdTotal")
  
  ElseIf rstControle.Fields("codOperacao") = 6 Then 'Acerto

     rstEstoqueCliente.Fields("cli_QtdTotal") = rstEstoqueCliente.Fields("cli_QtdTotal") - rstHistorico.Fields("ctl_QtdTotal")
     
  End If
  
  rstEstoqueCliente.Update
  rstEstoqueEmpresa.Update
  rstHistorico.MoveNext
  
 Loop
 rstEstoqueEmpresa.Close
 rstEstoqueCliente.Close
 
End If

rstControle.Close
rstHistorico.Close

Me.codOperacao.SetFocus
Me.cmdEfetivar.Enabled = False
Me.cmdVisualizar.Enabled = True
Me.subItens.Locked = True
Me.codCadastro.Locked = True
Me.srfNossoEstoque.Requery
Form_frmPesquisa.lstCadastro.Requery


End Sub

Private Sub codCadastro_Click()
Me.subItens.Locked = False
End Sub

Private Sub codOperacao_Click()
COperacao
End Sub

Sub COperacao()

Me.codCadastro.Locked = False

If Me.codOperacao.Column(0) = 1 Then
 Me.codCadastro.RowSource = "SELECT [codCadastro], [cad_Descricao], [codCategoria] FROM tblCadastros where codCategoria = 'F';"
ElseIf Me.codOperacao.Column(0) = 5 Then
 Me.codCadastro.RowSource = ""
Else
 Me.codCadastro.RowSource = "SELECT [codCadastro], [cad_Descricao], [codCategoria] FROM tblCadastros where codCategoria = 'C';"
End If

End Sub

Private Sub Form_Close()
Form_frmPesquisa.lstCadastro.Requery
End Sub

Private Sub Form_Current()
COperacao
If chkOk = True Then
   Me.cmdEfetivar.Enabled = False
   Me.codCadastro.Locked = True
   Me.codOperacao.Locked = True
   Me.subItens.Locked = True
   Me.cmdVisualizar.Enabled = True
Else
   Me.cmdEfetivar.Enabled = True
   Me.codCadastro.Locked = False
   Me.codOperacao.Locked = False
   Me.subItens.Locked = False
   Me.cmdVisualizar.Enabled = False
End If

End Sub

Private Sub cmdVisualizar_Click()
On Error GoTo Err_cmdVisualizar_Click

    Dim stDocName As String

    stDocName = "rptMovimento"
    DoCmd.OpenReport stDocName, acPreview, , "tblControleEstoque.codControle = " & Me.txtcodControle

Exit_cmdVisualizar_Click:
    Exit Sub

Err_cmdVisualizar_Click:
    MsgBox Err.Description
    Resume Exit_cmdVisualizar_Click
End Sub
