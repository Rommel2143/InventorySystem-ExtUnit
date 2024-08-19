Imports MySql.Data.MySqlClient
Public Class suggest_improvent
    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click

        insert_to_tblimprovement()

        MessageBox.Show("Thank you for sharing your suggestion. We will consider this one!", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
        txtsuggest.Clear()
    End Sub
    Private Sub insert_to_tblimprovement()
        Try

            con.Close()
            con.Open()
            Dim cmdinsert As New MySqlCommand("INSERT INTO `tblimprovement`(`reported`, `improvement`, `system`, `PClocation`) VALUES ('" & fname & "','" & txtsuggest.Text & "','IS','" & PClocation & "')", con)
            dr = cmdinsert.ExecuteReader


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            con.Close()
        End Try

    End Sub
End Class