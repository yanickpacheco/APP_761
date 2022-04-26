Imports System.Data.SqlClient
Public Class clsConexion

    Public Function CadenaConexion(ByVal IdBase As Integer) As String
        Dim stringConexion As String
        stringConexion = ""

        Select Case IdBase
            Case 1 '"GLOBAL"
                stringConexion = "Data Source = 192.168.1.14; initial catalog=BASE_GENERAL_CORDIAL; User Id= cp; Password=Cordial.passs"
            Case 2 '"BASEACTUAL"
                stringConexion = "Data Source=192.168.1.14;Initial Catalog=BDD_" + codigoProyecto + ";Persist Security Info=True;User ID=cp;Password=Cordial.passs"
                'Case 3 '"SRVCS"
                '    stringConexion = "Data Source=152.231.73.174;Initial Catalog=Global;Persist Security Info=True;User ID=mg;Password=cordial.passs"
        End Select
        Return stringConexion
    End Function

    Private Function conectar(ByVal x As Integer, ByVal con As SqlConnection) As Boolean
        con.Open()
        Return True
    End Function

    Private Sub desconectar(ByVal x As Integer, ByVal con As SqlConnection)
        If con.State = ConnectionState.Open Then
            con.Close()
        End If
    End Sub

    Public Function Ejecutar(ByVal cmd As SqlCommand, ByVal idConexion As String) As Boolean
        Dim valido As Boolean
        Dim con As New SqlConnection(CadenaConexion(idConexion))
        Try
            If conectar(idConexion, con) = True Then
                Dim rows As Integer
                cmd.Connection = con
                rows = cmd.ExecuteNonQuery
                If rows > 0 Then
                    valido = True
                Else
                    valido = False
                End If
            End If
        Catch ex As Exception
            MsgBox("ERROR: nombre procedimiento " + cmd.CommandText + " Detalle: " & ex.Message & vbNewLine, MsgBoxStyle.Critical, "CORDIALPHONE")
        Finally
            desconectar(idConexion, con)
        End Try
        Return valido
    End Function

    Public Function TraeDatos(ByVal cmd As SqlCommand, ByVal idConexion As String) As DataTable
        Dim con As New SqlConnection(CadenaConexion(idConexion))
        Dim DA As SqlDataAdapter
        Dim mitabla As New DataTable
        Try
            If conectar(idConexion, con) Then
                Dim dt As New DataTable
                cmd.Connection = con
                DA = New SqlDataAdapter(cmd)
                DA.Fill(mitabla)
            End If
        Catch ex As Exception
            MsgBox("ERROR: nombre procedimiento " + cmd.CommandText + " Detalle: " & ex.Message & vbNewLine, MsgBoxStyle.Critical, "CORDIALPHONE")
        Finally
            desconectar(idConexion, con)
        End Try
        Return mitabla
    End Function
End Class
