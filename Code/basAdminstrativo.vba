Option Compare Database
Option Explicit

Public strTabela As String
Public strIdentificacao As String
Public strDescricao As String
Public strTitulo As String
Public strCliente As String
Public strOrdem As String
Public strCriterio As String
Public strColunas As String

Public ProCOD As String
Public ProDes As String



Public Function NewCod(tabela, Campo)

Dim rs1 As DAO.Recordset
Set rs1 = CurrentDb.OpenRecordset("SELECT Max([" & Campo & "])+1 AS CodigoNovo FROM " & tabela & ";")
If Not rs1.EOF Then
   NewCod = rs1.Fields("CodigoNovo")
   If IsNull(NewCod) Then
      NewCod = 1
   End If
Else
   NewCod = 1
End If
rs1.Close

End Function

Public Function Cadastro(tabela As String, Identificacao As String, Optional Titulo As String, Optional Ordem As String, Optional Colunas As String)


Dim strLinkCriteria As String
    
'Variaveis publicas
strTabela = tabela
strIdentificacao = Identificacao
'strDescricao = Descricao
strTitulo = Titulo
'strCriterio = Criterio
strColunas = Colunas
strOrdem = Ordem

'Formulario
'strFormulario As String,
DoCmd.OpenForm "frmPesquisa", , , strLinkCriteria


End Function

Public Sub Salvar()
On Error GoTo Err_Salvar


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    Form_frmPesquisa.lstCadastro.Requery
    'Me.cmdCancelar.Enabled = False
    
Exit_Salvar:
    DoCmd.Close
    Exit Sub

Err_Salvar:
    MsgBox Err.Description
    Resume Exit_Salvar
    
End Sub
Public Sub Cancelar()
On Error GoTo Err_Cancelar


    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    Form_frmPesquisa.lstCadastro.Requery

    
Exit_Cancelar:
    DoCmd.Close
    Exit Sub

Err_Cancelar:
    If Err.Number <> 2046 Then
       MsgBox Err.Description
    End If
    Resume Exit_Cancelar
    
End Sub

Public Sub Fechar()
On Error GoTo Err_Fechar

    DoCmd.Close

Exit_Fechar:
    Exit Sub

Err_Fechar:
    MsgBox Err.Description
    Resume Exit_Fechar
    
End Sub

'Sub LimpaCampoVencimento()
'Dim rsNF As DAO.Recordset
'
'
'Set rsNF = CurrentDb.OpenRecordset("Select * from tblNotasFiscais")
'
'BeginTrans
'
'Do While Not rsNF.EOF
'
'    rsNF.Edit
'    rsNF.Fields("nf_dados4") = Mid(rsNF.Fields("nf_dados4"), 12)
'    rsNF.Update
'
'    rsNF.MoveNext
'
'Loop
'
'CommitTrans
'
'End Sub
