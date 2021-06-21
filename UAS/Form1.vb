Imports System.Data
Imports System.Data.OleDb
Public Class Pemakai
    Sub matikan_isian()
        txtKode.Enabled = False
        txtNama.Enabled = True
        cbStatus.Enabled = False
        txtPassword.Enabled = False
    End Sub
    Sub aktifkan_isian()
        txtKode.Enabled = True
        txtNama.Enabled = True
        cbStatus.Enabled = True
        txtPassword.Enabled = True
    End Sub
    Sub awal()
        Call matikan_isian()
        btnNew.Text = "New"
        btnSimpan.Enabled = False
        btnUbah.Enabled = False
        btnHapus.Enabled = False
        btnCari.Enabled = True
        btnKeluar.Enabled = True
    End Sub

    Sub tampilkan()
        da = New OleDbDataAdapter("SELECT * FROM Pemakai", Conn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "Pemakai")
        DataGridView1.DataSource = ds.Tables("Pemakai")
        DataGridView1.Refresh()
    End Sub

    Sub kosongkan()
        txtKode.Clear()
        txtNama.Clear()
        txtPassword.Clear()
        cbStatus.Text = ""
    End Sub

    Private Sub Pemakai_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call awal()
        Call Koneksi()
        Call tampilkan()
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Call aktifkan_isian()
        If btnNew.Text = "New" Then
            btnNew.Text = "Batal"
            btnSimpan.Enabled = True
            btnCari.Enabled = False
            btnKeluar.Enabled = False
        Else
            btnNew.Text = "Batal"
            Call awal()
            Call tampilkan()
            Call kosongkan()
            btnNew.Text = "New"
        End If
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSimpan.Click
        If txtKode.Text = "" Or txtNama.Text = "" Or txtPassword.Text = "" Then
            MsgBox("Data Belum Lengkap...!!")
            txtKode.Focus()
            Exit Sub
        Else
            If MessageBox.Show("Yakin akan simpan ?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call Koneksi()
                cmd = New OleDbCommand("Select * From Pemakai where Kode='" & txtKode.Text & "'", Conn)
                rd = cmd.ExecuteReader
                rd.Read()
                If Not rd.HasRows Then
                    Dim kode, nama, status, password, simpan As String
                    kode = txtKode.Text
                    nama = txtNama.Text
                    status = cbStatus.Text
                    password = txtPassword.Text
                    simpan = "INSERT INTO Pemakai  VALUES" & "('" & kode & "','" & nama & "','" & status & "','" & password & "')"
                    cmd = New OleDbCommand(simpan, Conn)
                    cmd.ExecuteNonQuery()
                    MsgBox("Simpan Sukses...!")
                    Call tampilkan()
                    Call kosongkan()
                    Call awal()
                End If
            End If
            End If
    End Sub

    Private Sub btnUbah_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUbah.Click
        If txtKode.Text = "" Or txtNama.Text = "" Or cbStatus.Text = "" Or txtPassword.Text = "" Then
            MsgBox("Data Belum Lengkap...!")
            txtKode.Focus()
            Exit Sub
        Else
            If MessageBox.Show("Yakin akan diubah ?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then

                Call Koneksi()
                Dim kode, nama, status, password, ubah As String
                kode = txtKode.Text
                nama = txtNama.Text
                status = cbStatus.Text
                password = txtPassword.Text
                ubah = "update Pemakai set Nama='" & nama & "',Status='" & status & "' ,Pass='" & password & "' where Kode='" & kode & "'"
                cmd = New OleDbCommand(ubah, Conn)
                cmd.ExecuteNonQuery()
                MsgBox("data berhasil di ubah...!")
            End If
        End If
        Call tampilkan()
        Call kosongkan()
        Call awal()
        txtKode.Focus()
    End Sub

    Private Sub btnCari_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCari.Click
        Call aktifkan_isian()
        btnNew.Text = "Batal"
        btnUbah.Enabled = True
        btnHapus.Enabled = True
        btnNew.Enabled = True
        btnSimpan.Enabled = False
        btnCari.Enabled = False
        btnKeluar.Enabled = False

        'proses pencarian
        cmd = New OleDbCommand("SELECT * FROM Pemakai WHERE Nama like '%" & txtNama.Text & "%'", Conn)
        rd = cmd.ExecuteReader
        rd.Read()
        If rd.HasRows Then
            da = New OleDbDataAdapter("SELECT * FROM Pemakai WHERE Nama like '%" & txtNama.Text & "%'", Conn)
            ds = New DataSet
            da.Fill(ds, "Dapat")
            DataGridView1.DataSource = ds.Tables("Dapat")
            DataGridView1.ReadOnly = True
        Else
            MsgBox("Data Tidak Ditemukan")
            Call awal()
            Call kosongkan()
            txtNama.Focus()
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHapus.Click
        If txtKode.Text = "" Then
            MsgBox("silahkan pilih data yang ingin dihapus")
        Else
            If MessageBox.Show("Yakin akan dihapus ?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call Koneksi()
                Dim Hapus As String = "delete From Pemakai where Kode='" & txtKode.Text & "'"
                cmd = New OleDbCommand(Hapus, Conn)
                cmd.ExecuteNonQuery()
                MsgBox("data berhasil di hapus")
                Call tampilkan()
                Call kosongkan()
                Call awal()
            End If
        End If
    End Sub

    Private Sub btnKeluar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKeluar.Click
        Me.Close()
    End Sub


    Private Sub DataGridView1_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        On Error Resume Next
        txtKode.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        txtNama.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        cbStatus.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        txtPassword.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
    End Sub
End Class
