Imports System.Collections.Generic
Imports Dato
Imports Entidad
Public Class clsParentescoCampaniaBI
    Dim daParentescoCampania As New clsParentescoCampaniaDA

    Public Function BuscarParentescoPorId(ByVal _idCRM As Integer, ByVal _tipoPersona As Integer) As List(Of eTipoParentesco)
        Return daParentescoCampania.BuscaParentescoPorId(_idCRM, _tipoPersona)
    End Function

    Public Function BuscaParentescoEdadPorId(ByVal _idCRM As Integer, ByVal _idTipoBeneficiarioCRM As Integer) As List(Of eTipoParentesco)
        Return daParentescoCampania.BuscaParentescoEdadPorId(_idCRM, _idTipoBeneficiarioCRM)
    End Function
End Class
