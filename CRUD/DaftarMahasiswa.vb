Imports System.Data.OleDb
Public Class DaftarMahasiswa
    Dim Conn As OleDbConnection
    Dim Da As OleDbDataAdapter
    Dim Ds As DataSet
    Dim LokasiDB As String
    Dim cmd As OleDbCommand
    Dim Rd As OleDbDataReader

    Sub Koneksi()
        LokasiDB = "Provider=Microsoft.ACE.OleDb.12.0;Data Source=DB_APLIKASI.accdb"
        Conn = New OleDbConnection(LokasiDB)
        If Conn.State = ConnectionState.Closed Then Conn.Open()
    End Sub

    Sub TableData()
        Call Koneksi()
        Da = New OleDbDataAdapter("Select * From TBL_MAHASISWA", Conn)
        Ds = New DataSet
        Ds.Clear()
        Da.Fill(Ds, "TBL_MAHASISWA")
        DataGridView1.DataSource = (Ds.Tables("TBL_MAHASISWA"))
    End Sub
    Private Sub DaftarMahasiswa_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call TableData()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    '===============================================
    'DI BAWAH INI KODE KALU KLIK PAKAI BUTTON LANGSUG NUNCUL DATA
    '===============================================
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call Koneksi()
        cmd = New OleDbCommand("Select * From TBL_MAHASISWA where NamaMahasiswa Like '%" & TextBox1.Text & "%'", Conn)
        Rd = cmd.ExecuteReader
        Rd.Read()
        If Rd.HasRows Then
            Call Koneksi()
            Da = New OleDbDataAdapter("Select * From TBL_MAHASISWA where NamaMahasiswa Like '%" & TextBox1.Text & "%'", Conn)
            Ds = New DataSet
            Ds.Clear()
            Da.Fill(Ds, "KetemuData")
            DataGridView1.DataSource = (Ds.Tables("KetemuData"))
            DataGridView1.ReadOnly = True
        Else
            MsgBox("Data tidak ada cuyy...")
        End If
    End Sub

    '===============================================
    'KALAU INI HANYA MENGETIKAN PADA TEXTBOX , UNTUK MEMUNCULKAN DATA DATAGRIDVIEW  
    '===============================================
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Call Koneksi()
        cmd = New OleDbCommand("Select * From TBL_MAHASISWA where NamaMahasiswa Like '%" & TextBox1.Text & "%'", Conn)
        Rd = cmd.ExecuteReader
        Rd.Read()
        If Rd.HasRows Then
            Call Koneksi()
            Da = New OleDbDataAdapter("Select * From TBL_MAHASISWA where NamaMahasiswa Like '%" & TextBox1.Text & "%'", Conn)
            Ds = New DataSet
            Da.Fill(Ds, "KetemuData")
            DataGridView1.DataSource = (Ds.Tables("KetemuData"))
            DataGridView1.ReadOnly = True
        Else
            MsgBox("Data tidak ada cuyy...")
        End If
    End Sub



    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TableData()
    End Sub


End Class