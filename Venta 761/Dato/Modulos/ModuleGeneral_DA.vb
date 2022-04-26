Imports System.Data.SqlClient

Module ModuleGeneral_DA
    Public Const codigoProyecto As String = "761" ' variable que almacenara el numero de la campaña se ocupara para los procedimientos para reemplazar el numero de la campaña
    Public cnn As New SqlConnection("data source = 190.98.236.22; initial catalog = BDD_" & codigoProyecto & "; User Id= cp; Password=Cordial.passs")
    Public cnnG As New SqlConnection("data source = 190.98.236.22; initial catalog = BASE_GENERAL_CORDIAL; User Id= cp; Password=Cordial.passs")

End Module
