using System;
using System.Data.SqlClient;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace ExcelToSqlApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                txtLog.AppendText("Se�ilen Dosya: " + filePath + Environment.NewLine);

                // Excel verisini SQL'e aktar
                AktarVerileriSQL(filePath);
            }
        }
        private void AktarVerileriSQL(string filePath)
        {
            try
            {
                // Excel dosyas�n� a�
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(1); // �lk sayfay� se�

                    using (SqlConnection con = new SqlConnection("Server=DESKTOP-4TMVP9R\\SQLSERVER2022;Database=ExcelDataDB;Integrated Security=true;"))
                    {
                        con.Open();

                        // �lk sat�r s�tun isimleri olarak al�nacak
                        bool firstRow = true;
                        foreach (var row in worksheet.RowsUsed()) // Kullan�lan t�m sat�rlar� al�r
                        {
                            if (firstRow)
                            {
                                firstRow = false;
                                continue; // �lk sat�r� atla (s�tun ba�l�klar�)
                            }

                            // Her bir sat�r� i�le
                            using (SqlCommand cmd = new SqlCommand("INSERT INTO Calisanlar (ID, FirstName, LastName) VALUES (@id, @firstName, @lastName)", con))
                            {
                                // ID de�erinin say�sal oldu�unu kontrol et
                                int id;
                                if (int.TryParse(row.Cell(1).GetString(), out id))
                                {
                                    cmd.Parameters.AddWithValue("@id", id); // ID
                                }
                                else
                                {
                                    txtLog.AppendText($"Hata: Ge�ersiz ID de�eri: {row.Cell(1).GetString()}" + Environment.NewLine);
                                    continue; // Ge�ersiz de�er varsa bu sat�r� atla
                                }

                                cmd.Parameters.AddWithValue("@firstName", row.Cell(2).GetString()); // AD
                                cmd.Parameters.AddWithValue("@lastName", row.Cell(3).GetString()); // SOYAD

                                cmd.ExecuteNonQuery();
                            }

                            txtLog.AppendText($"Sat�r i�lendi: {row.RowNumber()}" + Environment.NewLine);
                        }

                        txtLog.AppendText("T�m veriler ba�ar�yla aktar�ld�!" + Environment.NewLine);
                    }
                }
            }
            catch (Exception ex)
            {
                txtLog.AppendText("Hata: " + ex.Message + Environment.NewLine);
            }
        }



    }
}
