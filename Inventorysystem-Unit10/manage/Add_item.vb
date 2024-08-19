Imports MySql.Data.MySqlClient
Public Class Add_item
    Dim qrcode As String
    Dim partcode As String
    Dim supplier As String
    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click

        Try
            If P_partcode.Text = "" Then
                MessageBox.Show("Scan QR first!")
            Else
                con.Close()
                con.Open()
                Dim cmd As New MySqlCommand("SELECT `partcode` FROM `unit10_tblmaster` WHERE partcode = '" & P_partcode.Text & "' ", con)
                dr = cmd.ExecuteReader
                If dr.Read = False Then
                    con.Close()
                    con.Open()
                    Dim cmdinsert As New MySqlCommand(" INSERT INTO `unit10_tblmaster`
(`partname`, `partcode`, `stockU10`, `section`) 
VALUES ('" & P_Partname.Text & "','" & P_partcode.Text & "','0','" & P_section.Text & "')", con)
                    cmdinsert.ExecuteNonQuery()
                    MessageBox.Show("Partcode Added successfully!")
                    con.Close()
                    P_partcode.Clear()
                    P_Partname.Clear()
                    P_section.Clear()
                    txtqr.Text = ""
                    txtqr.Focus()
                Else
                    MessageBox.Show("Partcode already exists!")
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            con.Close()
        End Try

    End Sub

    Private Sub processQRcode(qrcode As String)
        Try

            Dim parts() As String = qrcode.Split("|")

            'CON 1 : QR SPLITING
            If parts.Length >= 5 AndAlso parts.Length <= 8 Then
                partcode = parts(0).Remove(0, 2).Trim

                supplier = parts(1).Remove(0, 2).Trim

            Else  'CON 1 : QR SPLITING
                MessageBox.Show("Invalid QR!")
                con.Close()
                txtqr.Text = ""
                txtqr.Focus()
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub Guna2Button3_Click(sender As Object, e As EventArgs) Handles Guna2Button3.Click
        Try
            If idno.Text = "" And fname.Text = "" Then
                MessageBox.Show("All Fields are required!")
            Else
                con.Close()
                con.Open()
                Dim cmd As New MySqlCommand("SELECT IDno FROM `scanoperator_is` WHERE IDno = '" & idno.Text & "' ", con)
                dr = cmd.ExecuteReader
                If dr.Read = False Then
                    con.Close()
                    con.Open()

                    Dim cmdinsert As New MySqlCommand(" INSERT INTO `scanoperator_is` (`IDno`, `Fullname`, `status`) VALUES ('" & idno.Text & "','" & fname.Text & "','user10')", con)
                    cmdinsert.ExecuteNonQuery()
                    MessageBox.Show("USER Added successfully!")
                    con.Close()
                    idno.Clear()
                    fname.Clear()
                Else
                    MessageBox.Show("USER already Exist!")
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            con.Close()
        End Try

    End Sub

    Private Sub txtqr_KeyDown(sender As Object, e As KeyEventArgs) Handles txtqr.KeyDown
        If e.KeyCode = Keys.Enter Then
            qrcode = txtqr.Text
            processQRcode(txtqr.Text)
            P_partcode.Text = partcode

        End If
    End Sub

End Class