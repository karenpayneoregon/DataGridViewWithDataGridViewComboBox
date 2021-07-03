Namespace Extensions
    Module Extensions
        <DebuggerStepThrough()>
        <Runtime.CompilerServices.Extension()>
        Public Function IsComboBoxCell(sender As DataGridViewCell) As Boolean
            Dim result As Boolean = False
            If sender.EditType IsNot Nothing Then
                If sender.EditType Is GetType(DataGridViewComboBoxEditingControl) Then
                    result = True
                End If
            End If
            Return result
        End Function

        <DebuggerStepThrough()>
        <Runtime.CompilerServices.Extension()>
        Public Function NotComboBox(ByVal sender As DataGridViewCell) As Boolean
            Return Not sender.IsComboBoxCell
        End Function


        <DebuggerStepThrough()>
        <Runtime.CompilerServices.Extension()>
        Public Function Identifier(ByVal sender As BindingSource) As Integer
            Return CType(sender.Current, DataRowView).Row.Field(Of Integer)("Id")
        End Function
    End Module
End Namespace