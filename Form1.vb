
Imports System.Data.OleDb


Public Class Form1

    Private Sub dgv_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv.CellContentClick

    End Sub
    Dim conn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet

    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader

    Sub koneksi()
        conn = New OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=Db.accdb")
        conn.Open()

    End Sub
    Sub kosongkan()

        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        ComboBox1.Text = ""
        TextBox1.Focus()
        Call tampilstatus()

    End Sub
    Sub databaru()

        TextBox2.Clear()
        TextBox3.Clear()
        ComboBox1.Text = ""
        TextBox2.Focus()

    End Sub
    Sub ketemu()

        'TextBox1.Enabled = False
        TextBox2.Text = dr("nama_user")
        TextBox3.Text = dr("pwd_user")
        ComboBox1.Text = dr("status_user")
        TextBox2.Focus()

    End Sub
    Sub tampilgrid()

        da = New OleDbDataAdapter("select * from Db", conn)
        ds = New DataSet
        da.Fill(ds)
        dgv.DataSource = ds.Tables(0)
        dgv.ReadOnly = True

    End Sub
    Sub carikode()

        cmd = New OleDbCommand("select * from Db where kode_user='" & TextBox1.Text & "'", conn)
        dr = cmd.ExecuteReader
        dr.Read()

    End Sub
    Sub tampilstatus()
        cmd = New OleDbCommand("select distinct status_user from Db", conn)
        dr = cmd.ExecuteReader
        ComboBox1.Items.Clear()

        Do While dr.Read
            ComboBox1.Items.Add(dr("status_user"))
        Loop

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.CenterToScreen()
        Call koneksi()
        Call kosongkan()
        Call tampilgrid()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Data Belum Lengkap")
            Exit Sub
        End If

        Call carikode()
        If Not dr.HasRows Then

            Dim simpan As String = "insert into Db values ('" & TextBox1.Text & "', '" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & ComboBox1.Text & "')"
            cmd = New OleDbCommand(simpan, conn)
            cmd.ExecuteNonQuery()
            Call kosongkan()
            Call tampilgrid()
            MsgBox("Data Berhasil Disimpan")

        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Data Belum Lengkap")
            Exit Sub
        End If

        Call carikode()
        If dr.HasRows Then

            Dim edit As String = "update Db set nama_user='" & TextBox2.Text & "',pwd_user='" & TextBox3.Text & "',status_user='" & ComboBox1.Text & "' where kode_user='" & TextBox1.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call kosongkan()
            Call tampilgrid()
            MsgBox("Data Berhasil Diedit")

        End If

    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            Call carikode()
            If dr.HasRows() Then
                Call ketemu()
            Else
                Call databaru()
            End If
        End If

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("kode harus diisi")
            TextBox1.Focus()
            Exit Sub
        End If

        Call carikode()
        If Not dr.HasRows Then
            MsgBox("kode tidak terdaftar")
            Exit Sub
        End If

        If MessageBox.Show("Yakin akan dihapus...?", "perhatian!!!", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Dim hapus As String = "delete from Db where kode_user='" & TextBox1.Text & "'"
            cmd = New OleDbCommand(hapus, conn)
            cmd.ExecuteNonQuery()
            Call kosongkan()
            Call tampilgrid()
            MsgBox("Data berhasil dihapus")
        Else
            Call kosongkan()
        End If

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call kosongkan()

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()

    End Sub

    Private Sub TextBox4_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        da = New OleDbDataAdapter("select * from Db where nama_user like '%" & TextBox4.Text & "%'", conn)
        ds = New DataSet
        da.Fill(ds)
        dgv.DataSource = ds.Tables(0)
        dgv.ReadOnly = True

    End Sub

    Private Sub dgv_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgv.CellMouseClick
        On Error Resume Next
        TextBox1.Text = dgv.Rows(e.RowIndex).Cells(0).Value

        Call carikode()
        If dr.HasRows Then
            Call ketemu()
        End If

    End Sub
End Class
