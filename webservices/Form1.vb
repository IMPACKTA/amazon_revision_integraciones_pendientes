Imports MySql.Data.MySqlClient

Public Class Form1

    Public conexion As MySqlConnection

    Public dbcomm As MySqlCommand
    Public dbread As MySqlDataReader
    Dim sql As String


    Dim textoFichero As String = ""
    Dim fechaFichero As Date = Now

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim fechaDescarga As Date = Now
        Dim fechaActual As String = fechaDescarga.ToString("dd/MM/yyyy HH:mm:ss")

        textoFichero = "INICIO: " & fechaActual & vbCrLf


        If connect() Then
            revision_integraciones()
        End If

        fechaDescarga = Now
        fechaActual = fechaDescarga.ToString("dd/MM/yyyy HH:mm:ss")

        textoFichero &= "FIN: " + fechaActual & vbCrLf

        Dim exists As Boolean = System.IO.Directory.Exists("LOG")

        If exists = False Then
            System.IO.Directory.CreateDirectory("LOG")
        End If

        Dim fFcihero As String = fechaFichero.ToString("ddMMyyyy")

        Dim ficheroLog As String = Application.StartupPath + "/LOG/" & fFcihero & ".txt"

        Dim sw As New System.IO.StreamWriter(ficheroLog, True, System.Text.Encoding.Default)
        sw.WriteLine(textoFichero)
        sw.Close()

        Me.Close()
    End Sub

    Sub revision_integraciones()
        If conexion.State Then
        Else
            connect()
        End If

        Dim fecha As Date = Now.AddDays(-1)
        Dim f As String = fecha.ToString("yyyy/MM/dd HH:mm:ss")


        Try

            textoFichero &= "Compruebo si hay pendiente integraciones de amazon pasadas 24 horas." + vbCrLf

            sql = "SELECT idEcommerce FROM clientesecommerce WHERE tipo='amazon' AND fechaRegistro <='" & f & "' AND accesstoken='' ORDER BY fechaRegistro, idEcommerce LIMIT 1"


            dbcomm = New MySqlCommand(sql, conexion)
            dbread = dbcomm.ExecuteReader()

            Dim table2 As New DataTable
            table2.Load(dbread)
            dbread.Close()


            If table2.Rows.Count > 0 Then

                Try
                    Dim idEcommerce As String = table2.Rows(0).Item(0).ToString

                    textoFichero &= "Se han encontrado " & table2.Rows.Count & " pendientes, elimino la primera tienda " & idEcommerce & vbCrLf

                    sql = "DELETE FROM clientesecommerce WHERE idEcommerce = '" & idEcommerce & "'"
                    dbcomm = New MySqlCommand(sql, conexion)
                    dbread = dbcomm.ExecuteReader()
                    dbread.Close()

                    textoFichero &= "Tienda " & idEcommerce & " eliminada correctamente" & vbCrLf


                    Try
                        textoFichero &= "Buscamos si hay tiendas por autorizar" & vbCrLf

                        sql = "SELECT idEcommerce FROM clientesecommerce WHERE tipo='amazon' AND accesstoken='' ORDER BY fechaRegistro, idEcommerce"
                        dbcomm = New MySqlCommand(sql, conexion)
                        dbread = dbcomm.ExecuteReader()

                        Dim table1 As New DataTable
                        table1.Load(dbread)
                        dbread.Close()

                        If table1.Rows.Count > 0 Then
                            For i = 0 To table1.Rows.Count - 1
                                Try
                                    idEcommerce = table2.Rows(i).Item(0).ToString
                                    textoFichero &= "Actualizo la fecha de registro de la tienda " & idEcommerce & " para dar 24 horas al cliente para su autorizacion" & vbCrLf

                                    fecha = Now
                                    f = fecha.ToString("yyyy/MM/dd HH:mm:ss")

                                    sql = "UPDATE clientesecommerce SET fechaRegistro='" & f & "'  WHERE idEcommerce = '" & idEcommerce & "'"
                                    dbcomm = New MySqlCommand(sql, conexion)
                                    dbread = dbcomm.ExecuteReader()
                                    dbread.Close()

                                    textoFichero &= "Fecha de registro actualizada correctamente" & vbCrLf
                                Catch ex As Exception
                                    textoFichero &= "ERROR ACTUALIZANDO FECHA REGISTRO " & ex.Message & vbCrLf
                                End Try
                            Next
                        Else
                            textoFichero &= "No hay tiendas por autorizar" & vbCrLf
                        End If



                    Catch ex As Exception
                        textoFichero &= "ERROR ACTUALIZANDO FECHA REGISTRO " & ex.Message & vbCrLf
                    End Try


                Catch ex As Exception
                    textoFichero &= "ERROR ELIMINANDO TIENDA " & ex.Message & vbCrLf
                End Try

            Else
                textoFichero &= "No hay integraciones pendientes" + vbCrLf
            End If


        Catch ex As Exception
            textoFichero &= "ERROR CONSULTA BD MYIMPACKTA " & ex.Message + vbCrLf
        End Try
    End Sub


    Function connect()
        Dim conectado As Boolean = False

        Dim servidorBD As String = ""
        Dim usuarioBD As String = ""
        Dim passwordBD As String = ""
        Dim nombreBD As String = ""

        Dim llave As String = ""

        llave = ficheros_conexion_bd.Class1.leerIni(Application.StartupPath + "/config.ini", "CONFIG", "LLAVE")

        servidorBD = ficheros_conexion_bd.Class1.leerIni(Application.StartupPath + "/config.ini", "CONFIG", "SERVER")
        usuarioBD = ficheros_conexion_bd.Class1.leerIni(Application.StartupPath + "/config.ini", "CONFIG", "USER")
        nombreBD = ficheros_conexion_bd.Class1.leerIni(Application.StartupPath + "/config.ini", "CONFIG", "BD")
        passwordBD = ficheros_conexion_bd.Class1.desencriptar(ficheros_conexion_bd.Class1.leerIni(Application.StartupPath + "/config.ini", "CONFIG", "PASSWORD"), llave)


        Try
            Dim str As String = "server=" & servidorBD & ";" & "user id=" & usuarioBD & ";" & "password=" & passwordBD & "; database=" & nombreBD & ";"
            conexion = New MySqlConnection(str)
            conexion.Open()

            conectado = True
        Catch ex As MySqlException
            textoFichero = textoFichero & "ERROR CONECTANDO CON LA BD DE MYIMPACKTA " + vbCrLf
        End Try

        Return conectado
    End Function
End Class
