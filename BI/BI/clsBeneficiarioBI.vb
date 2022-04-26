Imports Entidad
Imports Dato

Public Class clsBeneficiarioBI
    Dim clsBene As New clsBeneficiarioDA
    Public Sub Eliminar(ByVal cli As String)
        clsBene.Eliminar(cli)
    End Sub
    Public Sub Insertar(ByVal beneficiario As eBeneficiario)
        clsBene.Insertar(beneficiario)
    End Sub

    Public Function CargaBeneficiarios(ByVal id As String) As List(Of eBeneficiario)
        Return clsBene.Carga_Beneficiarios(id)
    End Function
End Class
