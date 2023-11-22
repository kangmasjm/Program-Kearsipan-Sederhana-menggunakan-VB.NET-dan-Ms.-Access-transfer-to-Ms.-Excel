Kali ini saya ingin sharing tentang aplikasi sederhana ini yang bernama Kearsipan.

Prosedur kerja aplikasi sederhana ini antara lain Input Surat Masuk dan Surat Keluar. Indikator penginputannya adalah:

Surat Masuk (No Surat, Tanggal Surat, Perihal, Pengirim)
Surat Keluar (No Surat, Tanggal Surat, Perihal, Penerima)
Kemudian data yang sudah masuk ke database bisa di export ke Ms. Excel…

Format .xls seperti berikut:
![image](https://github.com/kangmasjm/Program-Kearsipan-Sederhana-menggunakan-VB.NET-dan-Ms.-Access-transfer-to-Ms.-Excel/assets/59429191/35b85547-d8c9-4441-a274-324734c7682a)

Cell yang diwarnai kuning akan digunakan untuk counter saat record bertambah…

Lets Create the Project….

Buatlah database dengan nama DBKearsipan menggunakan Ms. Access yang terdiri dari tbl_surat_masuk dan tbl_surat_keluar.

Tabel Surat Masuk berisi (ID,NoSurat,TanggalSurat,Perihal,Dari)
Tabel Surat Keluar berisi (ID,NoSurat,TanggalSurat,Perihal,Kepada)
next…………

Pertama. buatlah desain form seperti gambar berikut ini:
![image](https://github.com/kangmasjm/Program-Kearsipan-Sederhana-menggunakan-VB.NET-dan-Ms.-Access-transfer-to-Ms.-Excel/assets/59429191/8961b814-8019-40f9-bf2f-576c1b6592b5)

Komponen form antara lain: (TextBox, ComboBox, Button, DataGridView)

Kedua. buatlah script impor oledb

Imports System.Data.OleDb

Imports excel= Microsoft.Office.Interop.Excel

Ketiga. buatlah koneksi ke database

tempatkan script dibawah Public Class Form1
Dim koneksi As New OleDb.OleDbConnection(“Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Juliana Mansur\Documents\DBKearsipan.accdb”)

Keempat. buatlah sub prosedur sesuai keinginan

Sub tampilGridSuratMasuk()
koneksi.Close()
koneksi.Open()

Me.Tbl_surat_masukTableAdapter.Fill(Me.DBKearsipanDataSet.tbl_surat_masuk)
Dim da As New OleDb.OleDbDataAdapter(“Select * From tbl_surat_masuk”, koneksi)
Dim ds As New DataSet
da.Fill(ds, “tbl_surat_masuk”)
Me.DGVM.DataSource = ds.Tables(“tbl_surat_masuk”)
Me.Tbl_surat_masukBindingSource.DataSource = ds.Tables(“tbl_surat_masuk”)
End Sub
Sub tampilGridSuratKeluar()
koneksi.Close()
koneksi.Open()

Me.Tbl_surat_keluarTableAdapter.Fill(Me.DBKearsipanDataSet.tbl_surat_keluar)
Dim da1 As New OleDb.OleDbDataAdapter(“Select * From tbl_surat_keluar”, koneksi)
Dim ds1 As New DataSet
da1.Fill(ds1, “tbl_surat_keluar”)
Me.DGVK.DataSource = ds1.Tables(“tbl_surat_keluar”)
Me.Tbl_surat_keluarBindingSource.DataSource = ds1.Tables(“tbl_surat_keluar”)
End Sub
Sub kosong()
Me.ComboBox1.Text = “”
Me.TextBox1.Text = “”
Me.TextBox2.Text = “”
Me.TextBox3.Text = “”
Me.TextBox4.Text = “”
Me.TextBox5.Text = “”
End Sub
Sub textOff()
Me.TextBox1.Enabled = False
Me.TextBox2.Enabled = False
Me.TextBox3.Enabled = False
Me.TextBox5.Enabled = False
End Sub
Sub textOn()
Me.TextBox1.Enabled = True
Me.TextBox2.Enabled = True
Me.TextBox3.Enabled = True
Me.TextBox5.Enabled = True
End Sub

Kelima. buatlah script di Form_Load

TextBox6.Visible = False
‘TODO: This line of code loads data into the ‘DBKearsipanDataSet.tbl_surat_masuk’ table. You can move, or remove it, as needed.
Me.Tbl_surat_masukTableAdapter.Fill(Me.DBKearsipanDataSet.tbl_surat_masuk)
‘TODO: This line of code loads data into the ‘DBKearsipanDataSet.tbl_surat_keluar’ table. You can move, or remove it, as needed.
Me.Tbl_surat_keluarTableAdapter.Fill(Me.DBKearsipanDataSet.tbl_surat_keluar)

Call tampilGridSuratKeluar()
Call tampilGridSuratMasuk()

Call textOff()
ComboBox1.Items.Add(“Surat Masuk”)
ComboBox1.Items.Add(“Surat Keluar”)
ComboBox5.Items.Add(“NoSurat”)
ComboBox5.Items.Add(“TanggalSurat”)
ComboBox5.Items.Add(“Perihal”)

Button2.Enabled = False
Button3.Enabled = False

Keenam. lengkapi script di masing-masing Command/Button

Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
If Button1.Text = “&ADD” Then
Call kosong()
Call textOn()
Button2.Enabled = True
Button3.Enabled = False
Button4.Enabled = False
ComboBox1.Focus()
Button1.Text = “&CANCEL”

ElseIf Button1.Text = “&CANCEL” Then
Call kosong()
Call textOff()
Button2.Enabled = False
Button3.Enabled = False
Button4.Enabled = False
Button1.Text = “&ADD”

End If
End Sub

Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
If Me.ComboBox1.Text <> “” And TextBox1.Text <> “” And TextBox2.Text <> “” And TextBox3.Text <> “” And TextBox5.Text <> “” Then
koneksi.Close()
koneksi.Open()
If ComboBox1.Text = “Surat Masuk” Then
Dim simpan As New OleDbCommand
simpan.Connection = koneksi
simpan.CommandType = CommandType.Text
simpan.CommandText = “INSERT INTO tbl_surat_masuk (NoSurat,TanggalSurat,Perihal,Kepada) VALUES (‘” & TextBox1.Text & “‘, ‘” & TextBox5.Text & “‘,'” & TextBox2.Text & “‘,'” & TextBox3.Text & “‘)”
simpan.ExecuteNonQuery()
MsgBox(“Data Tersimpan”)
ElseIf ComboBox1.Text = “Surat Keluar” Then
Dim simpan As New OleDbCommand
simpan.Connection = koneksi
simpan.CommandType = CommandType.Text
simpan.CommandText = “INSERT INTO tbl_surat_keluar (NoSurat,TanggalSurat,Perihal,Dari) VALUES (‘” & TextBox1.Text & “‘, ‘” & TextBox5.Text & “‘,'” & TextBox2.Text & “‘,'” & TextBox3.Text & “‘)”
simpan.ExecuteNonQuery()
MsgBox(“Data Tersimpan”)
End If
koneksi.Close()
Call kosong()
Call textOff()
Button1.Enabled = True
Button1.Text = “&ADD”
Button2.Enabled = False
Button3.Enabled = False
Button4.Enabled = False
Call tampilGridSuratKeluar()
Call tampilGridSuratMasuk()
Else
MsgBox(“Data tidak boleh kosong !”)
Call kosong()
End If
End Sub

Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
If ComboBox1.Text = “Surat Masuk” Then
Label6.Text = “Dari”
ComboBox5.Items.Add(“Dari”)
ElseIf ComboBox1.Text = “Surat Keluar” Then
Label6.Text = “Kepada”
ComboBox5.Items.Add(“Kepada”)
End If

End Sub

Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
If Button3.Text = “&EDIT” Then
Button1.Enabled = False
Button3.Text = “&UPDATE”
Call textOn()
ElseIf Button3.Text = “&UPDATE” Then
Dim tblsumber As String
If ComboBox1.Text = “Surat Masuk” Then
tblsumber = “UPDATE tbl_surat_masuk SET NoSurat = ‘” & TextBox1.Text & “‘,TanggalSurat = ‘” & TextBox5.Text & “‘, Perihal = ‘” & TextBox2.Text & “‘, Dari = ‘” & TextBox3.Text & “‘ WHERE ID = ‘” & TextBox6.Text & “‘”
ElseIf ComboBox1.Text = “Surat Keluar” Then
tblsumber = “UPDATE tbl_surat_keluar SET NoSurat = ‘” & TextBox1.Text & “‘,TanggalSurat = ‘” & TextBox5.Text & “‘, Perihal = ‘” & TextBox2.Text & “‘, Kepada = ‘” & TextBox3.Text & “‘ WHERE ID = ‘” & TextBox6.Text & “‘”
Else
tblsumber = “”
End If

koneksi.Close()
koneksi.Open()
Dim edit As New OleDbCommand
edit.Connection = koneksi
edit.CommandType = CommandType.Text
edit.CommandText = tblsumber
edit.ExecuteNonQuery()
MsgBox(“Data Berhasil Diubah !”)
Button4.Text = “&EDIT”
Call textOff()
Call kosong()
Button4.Enabled = False
Button2.Enabled = True
Button5.Enabled = False
End If
End Sub

Private Sub ComboBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged
If ComboBox1.Text = “” Then
MsgBox(“Pilih Jenis Surat”)
End If
End Sub

Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
If ComboBox5.Text <> “” And TextBox4.Text <> “” Then

Dim tabelna, tabelnm, sumber As String
If ComboBox1.Text = “Surat Masuk” Then
tabelnm = “tbl_surat_masuk”
sumber = “Dari”
tabelna = “SELECT * FROM tbl_surat_masuk WHERE ” & ComboBox5.Text & ” LIKE ‘%” & TextBox4.Text & “%'”
ElseIf ComboBox1.Text = “Surat Keluar” Then
tabelna = “SELECT * FROM tbl_surat_keluar WHERE ” & ComboBox5.Text & ” LIKE ‘%” & TextBox4.Text & “%'”
tabelnm = “tbl_surat_keluar”
sumber = “Kepada”
Else
tabelna = “”
tabelnm = “”
sumber = “”
End If
koneksi.Close()
koneksi.Open()
Dim cari As New OleDbCommand
cari.Connection = koneksi
cari.CommandType = CommandType.Text
cari.CommandText = tabelna
Dim dr As OleDbDataReader
dr = cari.ExecuteReader
If dr.HasRows = True Then
dr.Read()
TextBox1.Text = dr(“NoSurat”)
TextBox2.Text = dr(“Perihal”)
TextBox3.Text = dr(sumber)
TextBox5.Text = dr(“TanggalSurat”)
TextBox6.Text = dr(“ID”)
End If
Button3.Enabled = True
Button4.Enabled = True
Else
MsgBox(“Isi data pencariannya !”)
End If
End Sub

Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
koneksi.Close()
koneksi.Open()

Dim tbsumber As String
If ComboBox1.Text = “Surat Masuk” Then
tbsumber = “DELETE FROM tbl_surat_masuk WHERE ID='” & TextBox6.Text & “‘”
ElseIf ComboBox1.Text = “Surat Keluar” Then
tbsumber = “DELETE FROM tbl_surat_keluar WHERE ID='” & TextBox6.Text & “‘”
Else
tbsumber = “”
End If

MsgBox(tbsumber)
Dim hapus As New OleDbCommand
hapus.Connection = koneksi
hapus.CommandType = CommandType.Text
hapus.CommandText = tbsumber
hapus.ExecuteNonQuery()
MsgBox(“Data Berhasil Dihapus !”)

Call textOff()
Call kosong()

Call tampilGridSuratKeluar()
Call tampilGridSuratMasuk()

Button1.Enabled = True
Button2.Enabled = False
Button3.Enabled = False
Button4.Enabled = False
End Sub

terakhir ini script untuk export ke .xls

Dim ObjAppExcel As New excel.Application
Dim ObjDocExcel = ObjAppExcel.Workbooks.Open(“C:\Users\Juliana Mansur\Documents\Kearsipan.xls”)
Dim urutan As Integer
urutan = ObjAppExcel.Range(“F1″).Value
urutan = urutan + 1

ObjAppExcel.Range(“A” & urutan).Insert()
ObjAppExcel.Range(“A” & urutan).Value = ComboBox1.Text
ObjAppExcel.Range(“B” & urutan).Insert()
ObjAppExcel.Range(“B” & urutan).Value = TextBox1.Text
ObjAppExcel.Range(“C” & urutan).Insert()
ObjAppExcel.Range(“C” & urutan).Value = TextBox2.Text
ObjAppExcel.Range(“D” & urutan).Insert()
ObjAppExcel.Range(“D” & urutan).Value = TextBox3.Text
ObjAppExcel.Range(“E” & urutan).Insert()
ObjAppExcel.Range(“E” & urutan).Value = TextBox4.Text
ObjAppExcel.Range(“F1″).Value = urutan
ObjDocExcel.Save()
ObjDocExcel.Close()
ObjAppExcel.Quit()

Selesai…. Silahkan untuk mencoba….
