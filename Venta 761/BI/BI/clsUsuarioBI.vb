Imports Dato
Imports Entidad
Public Class clsUsuarioBI
    ''' <summary>
    ''' Obtiene el usuario a conectar
    ''' </summary>
    ''' <param name="user"></param>
    ''' <returns>boolean</returns>
    ''' <remarks></remarks>

    Dim dasUsuario As New clsUsuarioDA
Public Function Validar_user(ByVal user As String) As Boolean
        Return dasUsuario.Validar_user(user)
    End Function


    ''' <summary>
    ''' Este metodo valida si el usuario esta valido para Regrabaciones
    ''' </summary>
    ''' <param name="user"></param>
    ''' <param name="tipo"></param>
    ''' <returns>boolean</returns>
    ''' <remarks></remarks>
    Public Function Validar_Reg(ByVal user As String) As Boolean
        Return dasUsuario.Validar_REG(user)
    End Function
End Class
