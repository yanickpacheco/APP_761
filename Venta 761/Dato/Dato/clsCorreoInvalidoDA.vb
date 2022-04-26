Imports System.Collections.Generic
Imports Entidad

Public Class clsCorreoInvalidoDA
    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    'VARIABLE QUE ALMACENARA LOS PARAMETROS A LA BASE DE DATOS
    Dim dt As New DataTable

    '****************Metodo para listar todas las comunas**********************************
    ''' <summary>
    ''' Metodo para listar todos los correos inválidos
    ''' </summary>
    ''' <returns>List(Of clsComunaEnt)</returns>
    ''' <remarks></remarks>
    Public Function listarCorreosInvalidos() As List(Of eCorreoInvalido)

        'tipoConexion = 1
        'vlSqlParam.Clear()
        'query = "dbo.pa_CorreosInvalidos"

        'dt = con.TraeDatosConP(vlSqlParam, query, tipoConexion)

        'Dim listCorreosInvalidos As New List(Of eCorreoInvalido)

        'For x As Int16 = 0 To dt.Rows.Count - 1
        '    Dim row As New eCorreoInvalido

        '    row.idCorreoInvalido = IIf(IsDBNull(dt.Rows(x)("idCorreoInvalido")), Nothing, dt.Rows(x)("idCorreoInvalido"))
        '    row.correo = IIf(IsDBNull(dt.Rows(x)("correo")), Nothing, dt.Rows(x)("correo"))
        '    row.fechaIngreso = IIf(IsDBNull(dt.Rows(x)("fechaIngreso")), Nothing, dt.Rows(x)("fechaIngreso"))
        '    listCorreosInvalidos.Add(row)
        'Next

        'vlSqlParam.Clear()
        'Return listCorreosInvalidos

    End Function

End Class

