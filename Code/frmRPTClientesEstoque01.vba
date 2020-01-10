Option Compare Database
Dim WithEvents mRelatorio As Report

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click
    
    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    MsgBox Err.Description
    Resume Exit_cmdFechar_Click
    
End Sub

Private Sub cmdVer2_Click()
On Error GoTo Err_cmdSalvar_Click

 Dim StrAux As String
 Dim StrAux1 As String
 
 Set mRelatorio = New Report_rptClientesEstoque01
 
 strRelTitulo = "Relatório da OS"
 
 If Not IsNull(lstClientes2.Value) Then
    StrAux1 = " and codcadastro = " & lstClientes2.Value
    'strRelTitulo = lstClientes2.Column(1)
 End If
  
  
 With mRelatorio
    .Filter = "[Emissao] >= #" & Format(txtIni2.Value, "mm/dd/yyyy") & "# And [Emissao] <= #" & Format(txtFim2.Value, "mm/dd/yyyy") & "# " & StrAux1
    .FilterOn = True
    .Visible = True
 End With
  
    
Exit_cmdSalvar_Click:
    Exit Sub

Err_cmdSalvar_Click:
    MsgBox Err.Description
    Resume Exit_cmdSalvar_Click
    
End Sub


