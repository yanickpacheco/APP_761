Imports Dato
Imports Entidad

Public Class clsConfiguracionCampaniaBI

    Public Function ObtenerConfiguracionCampania(ByRef _CRM As String) As eProyecto
        Dim daConfiguracionCampania As New clsConfiguracionCampaniaDA
        Return daConfiguracionCampania.ObtenerConfiguracionCampania(_CRM)
    End Function

End Class
