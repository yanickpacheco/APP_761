Imports System.Collections.Generic
Imports Dato
Imports Entidad

Public Class clsCampoAdicionalBI

    '********* Busca todos los campos adicionales de la campaña filtrada ***************
    Public Function BuscarCampoAdicionalPorCalCodigo(ByVal _Calcodigo As Int32) As List(Of eCampoAdicional)

        Dim campoAdicional As New clsCampoAdicionalDA
        Return campoAdicional.BuscarCampoAdicionalPorCalCodigo(_Calcodigo)

    End Function


    '********* BUSCA TODOS LOS DATOS CONFIGURADOS EN LA TABLA CAMPOADICIONAL ***************
    Public Function BuscaDatosCampoAdicional(ByVal _Calcodigo As Int32, ByVal _idCliente As Int32) As List(Of eCampoAdicional)

        Dim campoAdicional As New clsCampoAdicionalDA
        Return campoAdicional.BuscaDatosCampoAdicional(_Calcodigo, _idCliente)

    End Function

End Class
