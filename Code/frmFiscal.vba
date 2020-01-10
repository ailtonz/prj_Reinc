Option Compare Database

Private Sub cmdSalvar_Click()
 Salvar
End Sub
Private Sub cmdCancelar_Click()
 Cancelar
End Sub
Private Sub cmdFechar_Click()
 Fechar
End Sub


Private Sub pro_Descricao_GotFocus()

If ProDes <> "" Then
   pro_Descricao.DefaultValue = "='" & ProDes & "'"
End If
ProDes = ""
ProCOD = ""

End Sub

