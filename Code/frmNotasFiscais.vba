Option Compare Database
Option Explicit
Dim WithEvents mRelatorio As Report

Private Sub cmdCopiar_Click()

Dim nrNota

Dim cnf_data
Dim cnf_natureza
Dim ccodCLI
Dim cnf_tipo
Dim cnf_cfop
Dim cnf_dados1
Dim cnf_dados2
Dim cnf_dados3
Dim cnf_dados4

Dim rs1 As DAO.Recordset
Dim rs2 As DAO.Recordset
Dim rs3 As DAO.Recordset
Set rs1 = CurrentDb.OpenRecordset("Select * from tblNotasfiscais WHERE codNF = " & Form_frmNotasFiscais.codNF)
Set rs2 = CurrentDb.OpenRecordset("Select * from tblNotasfiscaisItens WHERE codNF = " & Form_frmNotasFiscais.codNF)
Set rs3 = CurrentDb.OpenRecordset("Select * from tblNotasfiscaisItens ")

nrNota = InputBox("Digite o número da nota: ", "Nota Fiscal", NewCod(Form.RecordSource, codNF.ControlSource))
If Not nrNota = "" And Not nrNota = "0" Then

   If Not rs1.EOF Then

      cnf_data = rs1.Fields("nf_data")
      cnf_natureza = rs1.Fields("nf_natureza")
      ccodCLI = rs1.Fields("codcli")
      cnf_tipo = rs1.Fields("nf_tipo")
      cnf_cfop = rs1.Fields("nf_cfop")
      cnf_dados1 = rs1.Fields("nf_dados1")
      cnf_dados2 = rs1.Fields("nf_dados2")
      cnf_dados3 = rs1.Fields("nf_dados3")
      cnf_dados4 = rs1.Fields("nf_dados4")

      rs1.AddNew
      rs1.Fields("codnf") = nrNota
      rs1.Fields("nf_data") = Date
      rs1.Fields("nf_natureza") = cnf_natureza
      rs1.Fields("codcli") = ccodCLI
      rs1.Fields("nf_tipo") = cnf_tipo
      rs1.Fields("nf_cfop") = cnf_cfop
      rs1.Fields("nf_dados1") = cnf_dados1
      rs1.Fields("nf_dados2") = cnf_dados2
      rs1.Fields("nf_dados3") = cnf_dados3
      rs1.Fields("nf_dados4") = cnf_dados4
      rs1.Update

      If Not rs2.EOF Then

         Do While Not rs2.EOF

            rs3.AddNew
            rs3.Fields("codnf") = nrNota
            rs3.Fields("nfi_descricao") = rs2.Fields("nfi_descricao")
            rs3.Fields("nfi_qtd") = rs2.Fields("nfi_qtd")
            rs3.Fields("nfi_valor") = rs2.Fields("nfi_valor")
            rs3.Fields("nfi_icms") = rs2.Fields("nfi_icms")
            rs3.Update

            rs2.MoveNext

         Loop

      End If
      
      MsgBox "Copiar efetuado com sucesso! Feche este tela e abra o novo registro!"

   End If

End If

rs1.Close
rs2.Close

End Sub

Private Sub Form_Open(Cancel As Integer)

codNF.DefaultValue = NewCod(Form.RecordSource, codNF.ControlSource)

End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click
    
    Form_frmPesquisa.lstCadastro.Requery
    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    MsgBox Err.Description
    Resume Exit_cmdFechar_Click
    
End Sub
Private Sub cmdSalvar_Click()
On Error GoTo Err_cmdSalvar_Click
    
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    Form_frmPesquisa.lstCadastro.Requery
    DoCmd.Close
    
Exit_cmdSalvar_Click:
    Exit Sub

Err_cmdSalvar_Click:
    MsgBox Err.Description
    Resume Exit_cmdSalvar_Click
    
End Sub

Private Sub cmdVisualizar_Click()

On Error GoTo Err_cmdVisualizar_Click
 
 Set mRelatorio = New Report_rptNotasFiscais
 
  With mRelatorio
   .Caption = "Visualizando: " & codNF.Value
   .Filter = "codnf = " & codNF.Value
   .FilterOn = True
   .Visible = True
  End With

Exit_cmdVisualizar_Click:
    Exit Sub

Err_cmdVisualizar_Click:
    MsgBox Err.Description
    Resume Exit_cmdVisualizar_Click
    
End Sub

Private Sub nf_natureza_GotFocus()

'If IsNull(nf_natureza) Then
   'If Not IsNull(codfis) Then
   '   nf_natureza.Text = codfis.Column(2)
   'End If
'End If

End Sub

Private Sub nf_cfop_GotFocus()

'If IsNull(nf_cfop) Then
   'If Not IsNull(codfis) Then
   '   nf_cfop.Text = codfis.Column(3)
   'End If
'End If

End Sub

Private Sub cmdGerarBoleto_Click()

Dim rsNF As DAO.Recordset
Dim rsBoleto As DAO.Recordset
Dim rsVL_NF As DAO.Recordset
Dim strSQL As String
     
Set rsNF = CurrentDb.OpenRecordset("SELECT * from tblNotasFiscais")
Set rsBoleto = CurrentDb.OpenRecordset("SELECT * from tblBoletos")

Dim Instrucao01 As String
Dim Instrucao02 As String

Instrucao01 = "Sr. Caixa após vencimento cobrar juros mora de R$ "
Instrucao02 = " ao dia."

rsNF.FindFirst "codNF = " & Me.codNF


If Not rsNF.NoMatch Then

    strSQL = "SELECT tblNotasFiscaisItens.codNF, Sum([nfi_Valor]*[nfi_qtd]) AS Valor " & _
             "FROM tblNotasFiscaisItens GROUP BY tblNotasFiscaisItens.codNF  " & _
             "HAVING tblNotasFiscaisItens.codNF = " & Me.codNF
             
    Set rsVL_NF = CurrentDb.OpenRecordset(strSQL)
      
    BeginTrans
    
    rsBoleto.AddNew
    rsBoleto.Fields("codbol") = NewCod("tblBoletos", "codbol")
    rsBoleto.Fields("bol_numerodoc") = rsNF.Fields("codNF")
    rsBoleto.Fields("codCadastro") = rsNF.Fields("codCLI")
    rsBoleto.Fields("bol_datadoc") = rsNF.Fields("nf_data")
    rsBoleto.Fields("bol_dataprocess") = rsNF.Fields("nf_data")
    rsBoleto.Fields("bol_vencimento") = rsNF.Fields("nf_dados4")
    rsBoleto.Fields("bol_valordoc") = rsVL_NF.Fields("Valor")
    rsBoleto.Fields("bol_instrucoes") = Instrucao01 & rsNF.Fields("nf_ValorDaMora") & Instrucao02
    rsBoleto.Update

    CommitTrans

    MsgBox "Atenção: O boleto foi criado com sucesso!", vbOKOnly + vbInformation
    Me.codNF.SetFocus
    Me.cmdGerarBoleto.Enabled = False
    
End If

End Sub



