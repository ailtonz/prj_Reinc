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
 
 Set mRelatorio = New Report_rptEmpresaEstoque01
 
 If Not IsNull(lstClientes2.Value) Then
    StrAux1 = " and codcadastro = " & lstClientes2.Value
 End If
  
  With mRelatorio
   .Filter = "[dtEmissao] >= #" & Format(txtIni2.Value, "m/d/yyyy") & "# And [dtEmissao] <= #" & Format(txtFim2.Value, "m/d/yyyy") & "# " & StrAux1
   .FilterOn = True
   .Visible = True
  End With
  
    
Exit_cmdSalvar_Click:
    Exit Sub

Err_cmdSalvar_Click:
    MsgBox Err.Description
    Resume Exit_cmdSalvar_Click
    
End Sub


