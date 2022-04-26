Imports Dato
Imports Entidad
Imports System.Collections.Generic

Public Class clsParentescoBI

    Public Function BuscarParentescoPorId() As List(Of eParentesco)
        Dim daParentesco As New clsParentescoDA
        Return daParentesco.BuscarParentescoPorId()
    End Function

End Class
