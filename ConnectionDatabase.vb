Imports System.Data.SqlClient

Module ConnectionDatabase

    Public Function Database() As String
        Dim db As String = "Data Source= 192.168.1.188\SQLEXPRESS;
            initial catalog=TESE;
            Persist Security Info=True;
            User ID=TESE;
            Password=Sanindo123;
            Connect Timeout=15000;
            Max Pool Size=15000;
            MultipleActiveResultSets=true;
            Pooling=True"
        Return db
    End Function

    Public Function readData(query As String) As DataSet
        'Try
        Dim konek As New SqlConnection(Database)
            Dim sc As New SqlCommand(query, konek)
            Dim adapter As New SqlDataAdapter(sc)
            Dim ds As New DataSet

            adapter.Fill(ds)
            'koneksi.Close()

            Return ds
        'Catch ex As Exception
        'MsgBox("Database connection Error!")
        'End Try
    End Function

    Public Function insertData(query As String) As Boolean
        Dim konek As New SqlConnection(Database)
        konek.Open()
        Dim sc As New SqlCommand(query, konek)
        Dim adapter As New SqlDataAdapter(sc)

        If adapter.SelectCommand.ExecuteNonQuery().ToString() = 1 Then
            Return True
            konek.Close()
        Else
            Return False
        End If
    End Function

    Public Function updateData(query As String) As Boolean
        Try
            Dim konek As New SqlConnection(Database)
            konek.Open()
            Dim sc As New SqlCommand(query, konek)
            Dim adapter As New SqlDataAdapter(sc)
            Debug.WriteLine(adapter.SelectCommand.ExecuteNonQuery().ToString())
            If adapter.SelectCommand.ExecuteNonQuery().ToString() = 1 Then
                Return True
                konek.Close()
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex)
        End Try

    End Function

    Public Function deleteData(query As String) As Boolean
        Dim konek As New SqlConnection(Database)
        konek.Open()
        Dim sc As New SqlCommand(query, konek)
        Dim adapter As New SqlDataAdapter(sc)

        If adapter.SelectCommand.ExecuteNonQuery().ToString() = 1 Then
            Return True
            konek.Close()
        Else
            Return False
        End If
    End Function
End Module