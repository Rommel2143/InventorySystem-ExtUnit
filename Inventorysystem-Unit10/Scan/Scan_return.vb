Imports MySql.Data.MySqlClient
Public Class Scan_return

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

    Private Sub Scan_out_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtdate.Text = date1


    End Sub

    Private Sub txtqr_KeyDown(sender As Object, e As KeyEventArgs) Handles txtqr.KeyDown

        If e.KeyCode = Keys.Enter Then
            qrcode = txtqr.Text
            processQRcode(txtqr.Text)
        End If
    End Sub
    Private Sub processQRcode(qrcode As String)
        Try

            Dim parts() As String = qrcode.Split("|")

            'CON 1 : QR SPLITING
            If parts.Length >= 5 AndAlso parts.Length <= 8 Then
                partcode = parts(0).Remove(0, 2).Trim
                lotnumber = parts(2).Remove(0, 2).Trim
                qty = parts(3).Remove(0, 2).Trim
                remarks = parts(4).Remove(0, 2).Trim
                supplier = parts(1).Remove(0, 2).Trim

                'CON 2: DUPLICATION
                con.Close()
                con.Open()
                Dim cmdselect As New MySqlCommand("SELECT `qrcode`,`status`,`located`,`dateout` FROM `unit10_tblscan` WHERE `qrcode`='" & qrcode & "'", con)
                dr = cmdselect.ExecuteReader
                If dr.Read = True Then
                    status = dr.GetString("status")
                    located = dr.GetString("located")


                    Select Case status
                        Case "OUT"
                            'update out and deduct
                            add_to_stock(qty, partcode)
                            update_unit10_tblscan()
                            add_to_returncount(qrcode)
                            return_ok()
                            con.Close()

                        Case "IN"
                            'duplicate
                            showerror("QR Status IN")

                    End Select

                Else 'CON 2 else: DUPLICATION 
                    showerror("NO RECORD FOUND!")
                    con.Close()
                    txtqr.Text = ""
                    txtqr.Focus()

                End If
            Else  'CON 1 : QR SPLITING
                showerror("INVALID QR SCANNED!")
                con.Close()
                txtqr.Text = ""
                txtqr.Focus()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub


    Private Sub update_unit10_tblscan()
        Try
            con.Close()
            con.Open()
            Dim cmdupdate As New MySqlCommand("UPDATE `unit10_tblscan` SET `status`='IN',
                                                                    `batchout`='',
                                                                    `userout`='',
                                                                    `dateout`= Null,
                                                                    `boxno`='' 
                                                              WHERE `qrcode`='" & qrcode & "'", con)
            cmdupdate.ExecuteNonQuery()



        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            con.Close()
        End Try

    End Sub

    Private Sub showduplicate()
        Try
            labelerror.Visible = True
            texterror.Text = "DUPLICATE! Already scanned"
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

    Private Sub btndelete_Click(sender As Object, e As EventArgs)
        'print_report.Show()
        'print_report.BringToFront()
    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs)
        results_OUT.Show()
        results_OUT.BringToFront()
    End Sub

    Private Sub txtqr_TextChanged(sender As Object, e As EventArgs) Handles txtqr.TextChanged

    End Sub
End Class