Imports MySql.Data.MySqlClient
Public Class Scan_out

    Dim batch As String
    Dim supplier As String

    'duplicate info
    Dim status As String
    Dim located As String
    Dim dateout As String
    Dim partcode As String
    Dim qrcode As String
    Dim lotnumber As String
    Dim remarks As String
    Dim qty As Integer

    'selected item
    Dim itemid As String = ""
    Dim itempartcode As String = ""
    Dim itemqty As Integer = 0

    Private Sub Scan_In_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtdate.Text = date1


    End Sub

    Private Sub Txtqr_KeyDown(sender As Object, e As KeyEventArgs) Handles txtqr.KeyDown

        If e.KeyCode = Keys.Enter Then
            txtboxno.Clear()
            txtboxno.Focus()
        End If
    End Sub
    Private Sub ProcessQRcode(qrcode As String)
        Try

            Dim parts() As String = qrcode.Split("|")

            'CON 1 : QR SPLITING
            If parts.Length >= 5 AndAlso parts.Length <= 8 Then
                partcode = parts(0).Remove(0, 2).Trim
                lotnumber = parts(2).Remove(0, 2).Trim
                qty = parts(3).Remove(0, 2).Trim
                remarks = parts(4).Remove(0, 2).Trim
                supplier = parts(1).Remove(0, 2).Trim

                'CON 2 : DUPLICATE QR or GET location
                con.Close()
                con.Open()
                Dim cmdselect As New MySqlCommand("SELECT `qrcode`,`dateout`,`status` FROM `unit10_tblscan` WHERE `qrcode`='" & qrcode & "'", con)
                dr = cmdselect.ExecuteReader
                If dr.Read = True Then
                    status = dr.GetString("status")


                    Select Case status
                        Case "IN"
                            con.Close()
                            con.Open()
                            Dim cmdpartcode As New MySqlCommand("SELECT `partcode` FROM `unit10_tblmaster` WHERE `partcode`='" & partcode & "'", con)
                            dr = cmdpartcode.ExecuteReader
                            If dr.Read = True Then
                                deduct_to_stock(qty, partcode)
                                update_tblscan()
                                refreshgrid()
                                refreshgrid2()
                                return_ok()

                            Else  'CON 3 : PARTCODE
                                showerror("No Partcode Exists!")
                                return_ng()
                            End If

                        Case "OUT"
                            'duplicate
                            showduplicate()
                            return_ng()
                    End Select

                Else 'CON 2 : no record
                    'no record
                    showerror("No Record Found!")
                    return_ng()
                End If
            Else  'CON 1 : QR SPLITING
                showerror("INVALID QR SCANNED!")
                con.Close()
                return_ng()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub


    Private Sub Guna2TextBox2_TextChanged(sender As Object, e As EventArgs) Handles batchcode.TextChanged
        Try
            batch = batchcode.Text
            If batchcode.Text = "" Then
                txtqr.Enabled = False
                Label4.Visible = False
                Label7.Visible = False

            Else

                viewdata("SELECT `batchout`, `userout`, `dateout` FROM `unit10_tblscan`
                         WHERE `dateout`='" & datedb & "' and `userout`='" & idno & "' and `batchout`= '" & batchcode.Text & "' and `located`='" & PClocation & "'")
                If dr.Read = True Then
                    Label4.Visible = True
                    Label7.Visible = True
                    txtqr.Enabled = False
                Else
                    txtqr.Enabled = True
                    Label4.Visible = False
                    Label7.Visible = False
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub
    Private Sub update_tblscan()
        Try


            con.Close()
            con.Open()
            Dim cmdupdate As New MySqlCommand("UPDATE `unit10_tblscan` SET `status` = 'OUT', `userout`='" & idno & "', `dateout`='" & datedb & "', `batchout`= '" & batchcode.Text & "',`boxno`='" & txtboxno.Text & "' WHERE `qrcode`= '" & qrcode & "' ", con)
            cmdupdate.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs)
        results_IN.Show()
        results_IN.BringToFront()
    End Sub




    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub txtqr_TextChanged(sender As Object, e As EventArgs) Handles txtqr.TextChanged

    End Sub
    Private Sub showduplicate()
        Try
            labelerror.Visible = True
            texterror.Text = "DUPLICATE! Already scanned!"
            soundduplicate()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub showerror(text As String)

        Try
            labelerror.Visible = True
            texterror.Text = text
            sounderror()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub return_ok()
        Try
            labelerror.Visible = False
            txtqr.Clear()
            txtqr.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub return_ng()
        Try
            labelerror.Visible = True
            txtqr.Clear()
            txtqr.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub refreshgrid()
        Try
            con.Close()
            con.Open()
            Dim cmdrefreshgrid As New MySqlCommand("SELECT `id`,`batch`,`qrcode`,`partcode`,  `lotnumber`, `remarks`, `qty` FROM `unit10_tblscan`
                                                    WHERE `dateout`='" & datedb & "' and `userout`='" & idno & "' and `batchout`='" & batch & "' and `status`='OUT' ", con)

            Dim da As New MySqlDataAdapter(cmdrefreshgrid)
            Dim dt As New DataTable
            da.Fill(dt)
            datagrid1.DataSource = dt
            datagrid1.AutoResizeColumns()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally

            con.Close()
        End Try
    End Sub

    Private Sub refreshgrid2()
        Try
            con.Close()
            con.Open()
            Dim cmdrefreshgrid As New MySqlCommand("SELECT `partcode`, SUM(`qty`) FROM `unit10_tblscan`
                                                    WHERE `dateout`='" & datedb & "' and `batchout`='" & batch & "' and `userout`='" & idno & "'
                                                    GROUP BY partcode", con)

            Dim da As New MySqlDataAdapter(cmdrefreshgrid)
            Dim dt As New DataTable
            da.Fill(dt)
            datagrid2.DataSource = dt
            datagrid2.AutoResizeColumns()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally

            con.Close()
        End Try
    End Sub
    Private Sub datagrid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagrid1.CellClick
        Try
            With datagrid1
                itemid = .SelectedCells(0).Value
                itempartcode = .SelectedCells(3).Value.ToString()
                itemqty = .SelectedCells(6).Value()

            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub btndelete_Click(sender As Object, e As EventArgs) Handles btnprint.Click
        print_report.Show()
        print_report.BringToFront()
    End Sub


    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

        txtqr.Enabled = True
        Label4.Visible = False
        Label7.Visible = False

        refreshgrid()
        refreshgrid2()
    End Sub

    Private Sub cmbsearch_TextChanged(sender As Object, e As EventArgs) Handles cmbsearch.TextChanged
        Try
            con.Close()
            con.Open()
            Dim cmdrefreshgrid As New MySqlCommand("SELECT `id`,`batch`,`qrcode`,`partcode`,  `lotnumber`, `remarks`, `qty` FROM `unit10_tblscan`
                                                     WHERE `dateout`='" & datedb & "' and `located`='" & PClocation & "' and `userout`='" & idno & "' and `status`='OUT' and (`qrcode` REGEXP '" & cmbsearch.Text & "' or `batch` REGEXP '" & cmbsearch.Text & "')", con)

            Dim da As New MySqlDataAdapter(cmdrefreshgrid)
            Dim dt As New DataTable
            da.Fill(dt)
            datagrid1.DataSource = dt
            datagrid1.AutoResizeColumns()

            con.Close()
            con.Open()
            Dim cmdrefreshgrid2 As New MySqlCommand("SELECT `partcode`, SUM(`qty`) FROM `unit10_tblscan`
                                                  WHERE `dateout`='" & datedb & "' and `located`='" & PClocation & "' and `userout`='" & idno & "' and `status`='OUT' and (`qrcode` REGEXP '" & cmbsearch.Text & "' or `batch` REGEXP '" & cmbsearch.Text & "')               
                                                  GROUP BY partcode", con)

            Dim da2 As New MySqlDataAdapter(cmdrefreshgrid2)
            Dim dt2 As New DataTable
            da2.Fill(dt2)
            datagrid2.DataSource = dt2
            datagrid2.AutoResizeColumns()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click
        results_IN.Show()
        results_IN.BringToFront()

    End Sub

    Private Sub datagrid1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagrid1.CellContentClick

    End Sub

    Private Sub txtboxno_TextChanged(sender As Object, e As EventArgs) Handles txtboxno.TextChanged

    End Sub

    Private Sub txtboxno_KeyDown(sender As Object, e As KeyEventArgs) Handles txtboxno.KeyDown
        If e.KeyCode = Keys.Enter Then
            qrcode = txtqr.Text
            ProcessQRcode(txtqr.Text)

        End If


    End Sub
End Class