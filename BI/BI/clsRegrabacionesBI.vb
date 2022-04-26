Imports Dato
Imports Entidad

Public Class clsRegrabacionesBI
    Dim ges As New clsGesRegrabacionDA
    '********* METODO PARA TRAER TODOS LOS DATOS DE LAS VENTAS A SER GESTIONADAS EN REGRABACIONES ***************
    Public Function Regrabaciones(ByVal usuario As String) As DataTable
        Return ges.Regrabaciones(usuario)
    End Function
    '************METODO PARA GUARDAR EN LA TABLA GES_REGRABACION ************************
    Public Sub GuardaClienteGes(ByVal _cli As eCliente, ByVal _rec As eRecuperacion, ByVal _ges As eGestion)
        ges.GuardaClienteGes(_cli, _rec, _ges)
    End Sub
    '************PROCEDIMIENTO PARA ACTUALIZAR LOS CAMPOS EN LA TABLA CLIENTE SIN VENTA  ************************
    Public Sub ActualizaCteSinVta(ByVal _seg_est As String, ByVal _est_obj As String, ByVal _call_id As String, ByVal _cli_id As String)
        ges.ActualizaCteSinVta(_seg_est, _est_obj, _call_id, _cli_id)
    End Sub
    '************PROCEDIMIENTO PARA ACTUALIZAR LOS CAMPOS EN LA TABLA AsignacionCalidad  ************************
    Public Sub GrabaAsignaCalidad(ByVal _cliid As String, ByVal _calcodigo_centro As Integer, ByVal _agent As String, ByVal fecha As String, ByVal opcion As Integer, ByVal _agent_vta As String)
        ges.GrabaAsignaCalidad(_cliid, _calcodigo_centro, _agent, fecha, opcion, _agent_vta)
    End Sub
    '*********************METODO PARA ACTUALIZAR CLIENTE AGENDADO DE REGRABACIONES***********
    Public Sub ActualizaClienteAgen(ByVal _cliid As String, ByVal _est_obj As String, ByVal _call_id As String)
        ges.ActualizaClienteAgen(_cliid, _est_obj, _call_id)
    End Sub
    '*********************METODO PARA ACTUALIZAR CLIENTE ACON VENTA EN REGRABACIONES***********
    Public Sub ActualizaClienteVenta(ByVal _cli As eCliente, ByVal _ges As eGestion)
        ges.ActualizaClienteVenta(_cli, _ges)
    End Sub
End Class
