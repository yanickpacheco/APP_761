Imports Dato
Public Class clsListaUsuarioBaseBI
    Dim daLista As New clsListaUsuarioBaseDA

    Public Function ListarUsuarios(ByVal _crm As String) As DataTable
        Return daLista.ListarUsuarios(_crm)
    End Function
End Class
