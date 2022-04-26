Imports System.Collections.Generic
Imports Entidad

Public Class clsParentescoDA
    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    'VARIABLE QUE ALMACENARA LOS PARAMETROS A LA BASE DE DATOS
    Dim dt As New DataTable
    
    '***********METODO PARA BUSCAR UN PARENTESCO POR ID***************************
    Public Function BuscarParentescoPorId() As List(Of eParentesco)
        'tipoConexion = 2
        'vlSqlParam.Clear()
        'query = "dbo.pa_ListarParentesco"

        'dt = con.TraeDatosConP(vlSqlParam, query, tipoConexion)

        'Dim listParentesco As New List(Of eParentesco)
        'Dim row As New eParentesco
        'For x As Int16 = 0 To dt.Rows.Count - 1

        '    row.idParentesco = IIf(IsDBNull(dt.Rows(x)("idParentesco")), Nothing, dt.Rows(x)("idParentesco"))
        '    row.nombreParentesco = IIf(IsDBNull(dt.Rows(x)("nombreParentesco")), Nothing, dt.Rows(x)("nombreParentesco"))
        '    listParentesco.Add(row)
        'Next
        'vlSqlParam.Clear()

        'Return listParentesco

    End Function


End Class
