Imports System.IO 'esta libreria nos va a servir para poder activar el commandialog
Imports Microsoft.Office.Interop
Imports System.Data
Imports System.Data.OleDb
Imports System
Imports Microsoft.VisualBasic
Imports System.Data.SqlClient

Module Importar

    'Permitir conectarnos a nuestra base de datos sqlserver'
    Public cnn As SqlConnection
    Public sqlBC As SqlBulkCopy
    Public comando As SqlCommand

    'Conectar a la base de datos sqlserver'
    Sub abrirConexion()
        Try
            cnn = New SqlConnection("Data Source=10.109.18.81;Initial Catalog=GEANCAR; User Id = Ahora; Password = ")
            'cnn = New SqlConnection("Data Source=.;Initial Catalog=GEANCAR; User Id = sa; Password = fco0667")
            'strConn = "Data Source = IPPUBLICA: 1433; Initial Catalog = TUBD; User ID = USUARIO;  Password = CONTRASEÑA"
            cnn.Open()
            MessageBox.Show("CONECTADO A BASE DE DATOS GEANCAR")
        Catch ex As Exception
            MessageBox.Show("NO SE CONECTO: " + ex.ToString)
        End Try
    End Sub

    Sub importarExcel(ByVal tabla As DataGridView)
        Dim myFileDialog As New OpenFileDialog()
        Dim xSheet As String = ""

        borrar()

        With myFileDialog
            .Filter = "Excel Files |*.xlsx|Excel|*.xls"
            .Title = "Open File"
            .ShowDialog()
        End With
        If myFileDialog.FileName.ToString <> "" Then
            Dim ExcelFile As String = myFileDialog.FileName.ToString

            Dim ds As New DataSet
            Dim da As OleDbDataAdapter
            Dim dt As DataTable
            Dim conn As OleDbConnection

            xSheet = "PriceList"
            'xSheet = InputBox("Digite el nombre de la Hoja que desea importar", "Complete")
            conn = New OleDbConnection(
                              "Provider=Microsoft.ACE.OLEDB.12.0;" &
                              "data source=" & ExcelFile & "; " &
                             "Extended Properties='Excel 12.0 Xml;HDR=Yes'")

            Try
                da = New OleDbDataAdapter("SELECT * FROM  [" & xSheet & "$]", conn)


                conn.Open()
                da.Fill(ds, "MyData")
                dt = ds.Tables("MyData")
                tabla.DataSource = ds
                tabla.DataMember = "MyData"

                sqlBC = New SqlBulkCopy(cnn)
                sqlBC.DestinationTableName = "Pers_TarifaJCB_TMP"
                sqlBC.WriteToServer(ds.Tables(0))

            Catch ex As Exception

                MsgBox("Inserte un nombre valido de la Hoja que desea importar", MsgBoxStyle.Information, "Informacion")
            Finally
                conn.Close()
            End Try
        End If
        MsgBox("Se ha cargado la PriceList correctamente", MsgBoxStyle.Information, "Importado con exito")
    End Sub

    Sub borrar()

        Select Case MsgBox("Eliminar la Litsprice anterior y continuar", MsgBoxStyle.OkCancel, "caption")
            Case MsgBoxResult.Ok

                Try
                    ' cnn.Open()
                    Dim elimina As String = "delete from Pers_TarifaJCB_TMP"
                    comando = New SqlCommand(elimina, cnn)
                    Dim i As String = comando.ExecuteNonQuery()

                Catch ex As Exception
                    ' MsgBox("Inserte un nombre valido de la Hoja que desea importar", MsgBoxStyle.Information, "Informacion")
                    MsgBox("Error: " + ex.ToString, MsgBoxStyle.Information, "Informacion")
                Finally
                    'cnn.Close()
                End Try

                MessageBox.Show("La PriceList anterior se ha eliminado")

            Case MsgBoxResult.Cancel
                MessageBox.Show("El borrado ha sido cancelado")
                End
                Form1.Close()

                ' Case MsgBoxResult.No
                '    MessageBox.Show("NO button")
        End Select




    End Sub

    Sub actualizar()

        Try
            ' cnn.Open()
            Dim actualiza As String = "exec Ppers_JCB"
            comando = New SqlCommand(actualiza, cnn)
            Dim i As String = comando.ExecuteNonQuery()

        Catch ex As Exception
            'MsgBox("Inserte un nombre valido de la Hoja que desea importar", MsgBoxStyle.Information, "Informacion")
            MsgBox("Error: " + ex.ToString, MsgBoxStyle.Information, "Informacion")
        Finally
            'cnn.Close()
        End Try

        MsgBox("Finalizada la Actualizacion del JCB PRiceList")

    End Sub

    Sub Salir()

        Form1.Close()



    End Sub


End Module