Imports System.Data.OleDb

Module mdlcorona
    Public CONN As OleDbConnection
    Public DA As OleDbDataAdapter
    Public DS As DataSet
    Public CMD As OleDbCommand
    Public DR As OleDbDataReader

    Sub konekDB()
        Try
            CONN = New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=corona.mdb")
            CONN.Open()
            ' MsgBox("koneksi DataBase Sukses")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Module
