Imports System.Data.OleDb
Public Class Form1
    Dim Conn As OleDbConnection
    Dim Da As OleDbDataAdapter
    Dim Ds As DataSet
    Dim LokasiDb As String
    Dim cmd As OleDbCommand
    Dim Rd As OleDbDataReader

    Sub Koneksi()
        LokasiDb = "Provider=Microsoft.ACE.OleDb.12.0;Data Source=DB_APLIKASI.accdb"
        Conn = New OleDbConnection(LokasiDb)
        If Conn.State = ConnectionState.Closed Then Conn.Open()
    End Sub

    Sub KondisiAwal()
        Call Koneksi()
        Da = New OleDbDataAdapter("Select * From TBL_MAHASISWA", Conn)
        Ds = New DataSet()
        Ds.Clear()
        Da.Fill(Ds, "TBL_MAHASISWA")
        DataGridView1.DataSource = (Ds.Tables("TBL_MAHASISWA"))

        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("Pria")
        ComboBox1.Items.Add("Wanita")
        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList


        TextBox1.MaxLength = 3
    End Sub

    Sub DataBersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub

    'code dibawah di hiraukan dulu
    Sub DataBersih_1()
        TextBox2.Text = ""
        ComboBox1.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub

    'code dibawah di hiraukan dulu
    Sub TexMute()
        TextBox1.Enabled = True
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
    End Sub

    'code dibawah di hiraukan dulu
    Sub TextOpen()
        TextBox1.Enabled = False
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call TexMute()
        Call KondisiAwal()
    End Sub
    '===============================================
    'CODE DIBAWAH UNTUK TOMBOL INPUT / CREATE
    '===============================================
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'TAMBAHAN BARU LAGI, KETIKA BUTTON INPUT DIKLIK NANTI AKAN BERUBAH MENJADI SIMPAN DAN AKAN MUNCUL NIM SECARA OTOMATIS
        If Button1.Text = "INPUT" Then
            Call TextOpen()
            Button1.Text = "SIMPAN"
            Call Koneksi()
            Call DataBersih_1()
            'code dibawah untuk menjeneret code secara otomatis
            cmd = New OleDbCommand("Select * From TBL_MAHASISWA where NIM in (select max(NIM) from TBL_MAHASISWA)", Conn)
            Dim UrutanKode As String
            Dim Hitung As Long
            Rd = cmd.ExecuteReader
            Rd.Read()
            If Not Rd.HasRows Then
                UrutanKode = "BRG" + "001"
            Else
                Hitung = Microsoft.VisualBasic.Right(Rd.GetString(0), 3) + 1
                UrutanKode = "BRG" + Microsoft.VisualBasic.Right("000" & Hitung, 3)
            End If
            TextBox1.Text = UrutanKode
            TextBox2.Focus()
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or ComboBox1.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("Data Harus Terisi Semua...")
                Call TexMute()
                Button1.Text = "INPUT"
                Call DataBersih()
            Else
                Call Koneksi()
                Dim SimpanData As String = "insert into TBL_MAHASISWA values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox1.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
                cmd = New OleDbCommand(SimpanData, Conn)
                cmd.ExecuteNonQuery()
                MsgBox("Data Berhasi Ter-Input...")
                Call KondisiAwal()
                Call DataBersih()
                Button1.Text = "INPUT"
                Call TexMute()
            End If

        End If


    End Sub

    '===============================================
    'CODE DIBWAH UNTUK TOMBOL EDIT
    '===============================================
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Button2.Text = "EDIT" Then
            Call TextOpen()
            Button2.Text = "RUBAH"
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or ComboBox1.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("Data Harus Terisi Semua...")
                Call TexMute()
                Button2.Text = "EDIT"
                Button1.Focus()
            Else
                Call Koneksi()
                If MessageBox.Show("Anda Yakin INGIN mengedit KOLOM ini ?", "EDIT KOLOM", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Dim EditData As String = "UPDATE TBL_MAHASISWA set NamaMahasiswa='" & TextBox2.Text & "',JenisKelamin='" & ComboBox1.Text & "',AlamatMahasiswa='" & TextBox3.Text & "',TelpMahasiswa='" & TextBox4.Text & "' where NIM='" & TextBox1.Text & "'"
                    cmd = New OleDbCommand(EditData, Conn)
                    cmd.ExecuteNonQuery()
                    MsgBox("Data Berhasi Ter-Edit...")
                    Call KondisiAwal()
                    Call DataBersih()
                    Button2.Text = "EDIT"
                    Call TexMute()
                Else
                    MsgBox("Data Tidak JADI di EDIT...")
                    Call DataBersih()
                    Button2.Text = "EDIT"
                    Call TexMute()
                End If
            End If
        End If




    End Sub

    '===============================================
    'CODE DIBAWAH UNTUK TOMBOL DELETE
    '===============================================
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("Data Harus Terisi Semua...")
        Else
            Call Koneksi()
            If MessageBox.Show("Anda Yakin INGIN Hapus Data INI ?...", "INFO", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                Dim HapusData As String = "DELETE From TBL_MAHASISWA  where NIM='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(HapusData, Conn)
                cmd.ExecuteNonQuery()
                MsgBox("Data Berhasi Ter-Hapus ...")
                Call KondisiAwal()
                Call DataBersih()
            Else
                MsgBox("Data Anda Masih Aman Cuyy ...")
            End If
        End If
    End Sub

    '===============================================
    'CODE DIBAWAH UNTUK TOMBOL CLOSE PADA FORM
    '===============================================
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    '===============================================
    'CODE DIBAWAH UNTUK MENAMPILKAN DATA BERDASARKAN "NIM"
    '===============================================
    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            cmd = New OleDbCommand("Select * From TBL_MAHASISWA where NIM='" & TextBox1.Text & "'", Conn)
            Rd = cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                TextBox2.Text = Rd.Item("NamaMahasiswa")
                ComboBox1.Text = Rd.Item("JenisKelamin")
                TextBox3.Text = Rd.Item("AlamatMahasiswa")
                TextBox4.Text = Rd.Item("TelpMahasiswa")
            Else
                MsgBox("Data Yang lu Cari , Kagak Ada Jirr ...")
                Call DataBersih()
            End If
        End If
    End Sub

    '===============================================
    'CODE DIBAWAH INTUK MENAMPILKAN WINDOWS/FORM BARU YANG NAMA NYA "DAFTAR MAHASISWA"
    '===============================================
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim NewData As New DaftarMahasiswa()

        NewData.Show()
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index

        On Error Resume Next
        Button1.Text = "INPUT"
        Call TexMute()
        Button2.Text = "EDIT"
        TextBox1.Enabled = False
        TextBox1.Text = DataGridView1.Item(0, i).Value
        TextBox2.Text = DataGridView1.Item(1, i).Value
        ComboBox1.Text = DataGridView1.Item(2, i).Value
        TextBox3.Text = DataGridView1.Item(3, i).Value
        TextBox4.Text = DataGridView1.Item(4, i).Value


    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Call DataBersih()
        Button1.Text = "INPUT"
        Button2.Text = "EDIT"
        Call TexMute()
    End Sub


End Class
