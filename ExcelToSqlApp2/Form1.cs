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
                txtLog.AppendText("Seçilen Dosya: " + filePath + Environment.NewLine);

                // Excel verisini SQL'e aktar
                AktarVerileriSQL(filePath);
            }
        }
        private void AktarVerileriSQL(string filePath)
        {
            try
            {
                // Excel dosyasýný aç
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(1); // Ýlk sayfayý seç

                    using (SqlConnection con = new SqlConnection("Server=DESKTOP-4TMVP9R\\SQLSERVER2022;Database=ExcelDataDB;Integrated Security=true;"))
                    {
                        con.Open();

                        // Ýlk satýr sütun isimleri olarak alýnacak
                        bool firstRow = true;
                        foreach (var row in worksheet.RowsUsed()) // Kullanýlan tüm satýrlarý alýr
                        {
                            if (firstRow)
                            {
                                firstRow = false;
                                continue; // Ýlk satýrý atla (sütun baþlýklarý)
                            }

                            // Her bir satýrý iþle
                            using (SqlCommand cmd = new SqlCommand("INSERT INTO Calisanlar (ID, FirstName, LastName) VALUES (@id, @firstName, @lastName)", con))
                            {
                                // ID deðerinin sayýsal olduðunu kontrol et
                                int id;
                                if (int.TryParse(row.Cell(1).GetString(), out id))
                                {
                                    cmd.Parameters.AddWithValue("@id", id); // ID
                                }
                                else
                                {
                                    txtLog.AppendText($"Hata: Geçersiz ID deðeri: {row.Cell(1).GetString()}" + Environment.NewLine);
                                    continue; // Geçersiz deðer varsa bu satýrý atla
                                }

                                cmd.Parameters.AddWithValue("@firstName", row.Cell(2).GetString()); // AD
                                cmd.Parameters.AddWithValue("@lastName", row.Cell(3).GetString()); // SOYAD

                                cmd.ExecuteNonQuery();
                            }

                            txtLog.AppendText($"Satýr iþlendi: {row.RowNumber()}" + Environment.NewLine);
                        }

                        txtLog.AppendText("Tüm veriler baþarýyla aktarýldý!" + Environment.NewLine);
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
