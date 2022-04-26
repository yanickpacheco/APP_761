Imports Entidad
Imports Dato

Public Class clsScriptBI


    '********* Busca el script por el calcodigo y el idTipoScript ***************
    Public Function BuscarScriptPorIdTipoScript(ByVal _calCodigo As Int32, ByVal _idTipoScript As Int32) As eScript

        Dim daScript As New clsScriptDA
        Return daScript.BuscarScriptPorIdTipoScript(_calCodigo, _idTipoScript)

    End Function
End Class
