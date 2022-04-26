Imports Dato
Imports Entidad
Public Class clsCiudadBI

    ''' <summary>
    ''' Obtiene una ciudad filtrado por idCiudad
    ''' </summary>
    ''' <param name="_idCiudad"></param>
    ''' <returns>Entidad ciudad</returns>
    ''' <remarks></remarks>
    Public Function BuscaCiudadPorIdCiudad(ByVal _idCiudad As Int32) As eCiudad

        Dim daCiudad As New clsCiudadDA
        Return daCiudad.BuscaCiudadPorIdCiudad(_idCiudad)

    End Function

    Public Function ListaCiudad() As List(Of eCiudad)
        Dim daCiudad As New clsCiudadDA
        Return daCiudad.ListaCiudad()
    End Function

End Class
