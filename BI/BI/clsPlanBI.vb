Imports Dato
Imports System.Collections.Generic
Imports Entidad

Public Class clsPlanBI

    Public Function BuscarPlanPorTipoContrato(ByVal _idTipoContrato As Int16, ByVal _idCrm As Int32) As List(Of ePlan)
        Dim biplan As New clsPlanDA
        Return biplan.BuscarPlanPorTipoContrato(_idTipoContrato, _idCrm)
    End Function

    Public Function BuscarPlanPorIdPlan(ByVal _idPlan As Int64) As ePlan
        Dim biplan As New clsPlanDA
        Return biplan.BuscarPlanPorIdPlan(_idPlan)
    End Function

End Class
