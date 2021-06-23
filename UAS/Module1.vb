Imports System.Data.OleDb
Module Module1
    Public Conn As OleDbConnection
    Public da As OleDbDataAdapter
    Public ds As DataSet
    Public cmd As OleDbCommand
    Public rd As OleDbDataReader
    Public Str As String
    Public Sub Koneksi()
        'Str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\UAS.accdb"
        Str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=UAS.accdb"
        Conn = New OleDbConnection(Str)
        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
            'MsgBox("Konek")
        Else
            'MsgBox("Koneksi Gagal...!")
        End If
    End Sub
    Sub btn()

    End Sub
End Module
