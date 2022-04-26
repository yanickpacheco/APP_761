Imports Dato
Imports Entidad

Public Class clsTipoContratoBI

    Public Function ListaTipoContratoPorCampania(ByVal _calCodigo As Int32) As List(Of eTipoContrato)
        Dim daTipoContrato As New clsTipoContratoDA
        Return daTipoContrato.ListaTipoContratoPorCampania(_calCodigo)
    End Function

    Public Function BuscarTipoContratoPorIdTipoContrato(ByVal _idTipoContrato As Int16) As eTipoContrato
        Dim daTipoContrato As New clsTipoContratoDA
        Return daTipoContrato.BuscarTipoContratoPorIdTipoContrato(_idTipoContrato)
    End Function

    Public Function ListaTipoContratoPosAdicionales(ByVal _calCodigo As Int16, ByVal _cantAdic As SByte) As List(Of eTipoContrato)
        Dim daTipoContrato As New clsTipoContratoDA
        Return daTipoContrato.ListaTipoContratoPosAdicionales(_calCodigo, _cantAdic)
    End Function


End Class
