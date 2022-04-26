Imports Entidad
Imports Dato

Public Class clsAdicionalBI
    Dim clsAdic As New clsAdicionalDA

    Public Sub Insertar(ByVal adicional As eAdicional)
        clsAdic.Insertar(adicional)
    End Sub

    Public Function Carga_Adicional(ByVal _id As String) As List(Of eAdicional)
        Return clsAdic.Carga_Adicional(_id)
    End Function

    Public Sub Eliminar(ByVal _c_id As Int64)
        clsAdic.Eliminar(_c_id)
    End Sub

End Class
