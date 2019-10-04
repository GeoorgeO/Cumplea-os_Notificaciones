Imports System.Data.SqlClient

Public Class AvisoCumple
    Dim tVentanaY As Integer = 200
    Dim tVentanaX As Integer = 300
    Dim arreglo() As String
    Dim arreglo2() As String
    Dim m As Integer

    Private Sub AvisoCumple_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'animacionventana()
        'Timer1.Start()
        Dim nombre As String = ""
        Dim fecha As Date
        Dim lector As SqlDataReader
        Dim bandera As Boolean = False



        Dim n As Integer

        Try
            m = 1
            Using cnx = New SqlConnection("Data Source=192.168.3.254;Initial Catalog=Vistas;User ID=sa;Password=inventumc762$")
                cnx.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnx
                    cmd.CommandText = "SET DATEFORMAT dmy   select Nombre, FechaC from RH_Cumpleanios  where FechaC>= getdate()-10 and FechaC<= getdate() + 10  order by FechaC"
                    n = 0

                    lector = cmd.ExecuteReader
                    While lector.Read
                        nombre = CStr(lector(0).ToString)
                        fecha = CDate(lector(1).ToString)
                        Dim numdias As Integer = DateDiff(DateInterval.Day, Today, CDate(fecha))

                        If numdias <= 2 And numdias >= 0 Then

                            Select Case numdias
                                Case 0
                                    Label1.Text = "Hoy cumple años:"
                                Case 1
                                    Label1.Text = "Mañana cumple años:"
                                Case 2
                                    Label1.Text = "Pasado mañana cumple años:"
                            End Select
                            ReDim Preserve arreglo(n)
                            ReDim Preserve arreglo2(n)
                            arreglo2(n) = Label1.Text
                            arreglo(n) = nombre
                            n = n + 1
                            Label2.Text = nombre
                            bandera = True
                            'Dim sonido As System.Media.SoundPlayer
                            'sonido = New System.Media.SoundPlayer(My.Application.Info.DirectoryPath + "\Sound_WAV.wav")
                            'sonido.Play()


                        Else
                            Dim dia As Integer = CDate(fecha).Date.Day
                            Dim mes As Integer = CDate(fecha).Date.Month
                            Dim anios As Integer = CDate(fecha).Date.Year
                            actualizaCumple(dia, mes, anios, nombre)

                        End If
                    End While
                    lector.Close()
                    If bandera = False Then
                        Me.Close()
                    End If
                End Using
                cnx.Close()
            End Using
            If arreglo.Length >= 1 Then
                otravez(0)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "HA ocurrido un error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Function actualizaCumple(ByVal dia, ByVal mes, ByVal anio, ByVal nombre)

        If anio <= CInt(Now.Date.Year) Then
            If (anio = CInt(Now.Date.Year) And CInt(Now.Date.Month) >= mes) Or (anio < CInt(Now.Date.Year)) Then
                If (CInt(Now.Date.Day) > dia And
                    CInt(Now.Date.Month) = mes) Or (CInt(Now.Date.Month) > mes) Or (anio < CInt(Now.Date.Year)) Then
                    Try
                        Using cnx = New SqlConnection("Data Source=192.168.3.254;Initial Catalog=Vistas;User ID=sa;Password=inventumc762$")
                            cnx.Open()
                            Using cmd As New SqlCommand()
                                cmd.Connection = cnx
                                cmd.CommandText = "SET DATEFORMAT dmy   update RH_cumpleanios set FechaC='" & dia & "/" & mes & "/" & anio + 1 & "' where Nombre='" & nombre & "'"

                                cmd.ExecuteNonQuery()


                            End Using
                            cnx.Close()
                        End Using

                    Catch ex As Exception
                        MessageBox.Show(ex.Message, "HA ocurrido un error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try

                End If
            End If
        End If
        Return Nothing
    End Function

    Function animacionventana()
        Dim desktopSize As Size
        desktopSize = System.Windows.Forms.SystemInformation.PrimaryMonitorSize
        Dim height As Integer = desktopSize.Height
        Dim width As Integer = desktopSize.Width
        Me.Location = New Point(width - tVentanaX, height - tVentanaY + 130)

        tVentanaY = tVentanaY + 5
        If tVentanaY >= 354 Then
            Timer1.Stop()
        End If
        Return Nothing
    End Function

    Function otravez(ByVal veces As Integer)
        Dim sonido As System.Media.SoundPlayer
        sonido = New System.Media.SoundPlayer(My.Application.Info.DirectoryPath + "\Sound_WAV.wav")
        tVentanaY = 200
        animacionventana()
        Timer1.Start()
        Label1.Text = arreglo2(veces)
        Label2.Text = arreglo(veces)

        sonido.Play()

        Return Nothing
    End Function

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        animacionventana()
    End Sub

    Private Sub PictureBox2_Click(sender As System.Object, e As System.EventArgs) Handles PictureBox2.Click
        If arreglo.Length >= 1 And m <= arreglo.Length - 1 Then
            Me.Visible = True
            otravez(m)
            m = m + 1
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub Label3_Click(sender As System.Object, e As System.EventArgs) Handles Label3.Click
        
        If arreglo.Length > 1 And m <= arreglo.Length - 1 Then
            Me.Visible = True
            otravez(m)
            m = m + 1
        Else
            Me.Visible = False
        End If
        'Me.Visible = False
    End Sub

    Private Sub Timer2_Tick(sender As System.Object, e As System.EventArgs) Handles Timer2.Tick
        Me.Visible = True
    End Sub

    Private Sub PictureBox3_Click(sender As System.Object, e As System.EventArgs) Handles PictureBox3.Click

    End Sub
End Class