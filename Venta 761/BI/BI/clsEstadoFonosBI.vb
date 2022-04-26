Imports Dato
Imports Entidad

Public Class clsEstadoFonosBI
    Dim daEstadoFono As New clsEstadoFonosDA

    Public Function listarEstadoFono() As List(Of eEstadoFonos)
        Return daEstadoFono.listarEstadoFono()
    End Function

    Public Sub guardarEstadoFono(ByVal callid As String, ByVal id As String, ByVal estadoFono As Int16)
        daEstadoFono.guardaEstadoFono(callid, id, estadoFono)
    End Sub
End Class
