Imports System.Data
Imports System.Data.OleDb

Module funcionesdb

    Public conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=database2.mdb;Persist Security Info=False;")

    Public cmd As New OleDb.OleDbCommand
    Public dr As OleDb.OleDbDataReader

    Public Sub conectarse()

        Try
            conn.Open()
            MsgBox("Conexión exitosa")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub consultar(ByRef identificacion As String)
        cmd.Connection = conn
        cmd.CommandType = CommandType.Text

        If identificacion <> "" Then
            cmd.CommandText = "SELECT NOMBRES, APELLIDOS, CORREO, DIRECCION FROM PERSONA WHERE IDPERSONA=" + identificacion
        Else
            cmd.CommandText = "SELECT NOMBRES, APELLIDOS, CORREO, DIRECCION FROM PERSONA "

        End If
        Try
            dr = cmd.ExecuteReader()

            If dr.HasRows Then
                While dr.Read()
                    MsgBox(dr(0).ToString + " " + dr(1).ToString + " " + dr(2).ToString + " " + dr(3).ToString)

                End While

            Else
                MsgBox("No existe registro para la consulta")

            End If
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

End Module
