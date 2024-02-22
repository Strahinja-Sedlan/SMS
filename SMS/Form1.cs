using System.Data.OleDb;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualBasic;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace SMS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:/Users/PC/Desktop/DDK/DDK - Indjija.mdb");
            OleDbCommand select = new OleDbCommand();
            select.Connection = con;
            select.CommandText = "Select * From listingIndjija";
            con.Open();
            OleDbDataReader reader = select.ExecuteReader();
            var path = @"C:\users\pc\Desktop";
            using var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Sheet1");
            worksheet.ColumnWidth = 20;
            worksheet.Cell("A1").Value = "+3816*******";
            worksheet.Cell("A2").Value = "+3816*******";
            worksheet.Cell("A3").Value = "+3816*******";
            var selectedDate = dateTimePicker1.Value;
            int count = 4;
            while (reader.Read())
            {
                bool townExists = reader.IsDBNull(8);
                bool isChecked = reader.GetBoolean(1);
                bool phoneExists = reader.IsDBNull(3);
                string phone = "";
                if (phoneExists == false)
                {
                    phone = reader.GetString(3);
                }
                else
                {
                    continue;
                }
                string mobilePhone = "";
                if (phone.StartsWith("+") || phone.StartsWith("3"))
                {
                    mobilePhone = phone;
                }
                else
                {
                    continue;
                }
                bool birthdayExists = reader.IsDBNull(6);
                string date = "";
                if (birthdayExists == false)
                {
                    date = reader.GetDateTime(6).ToString();
                }
                bool genderExists = reader.IsDBNull(5);
                string gender = "";
                if (genderExists == false)
                {
                    gender = reader.GetString(5);
                }
                DateTime birthday;
                DateTime.TryParse(date, out birthday);
                DateTime lastDonation;
                DateTime thisYear = DateTime.Now;
                bool checkLastDonation = reader.IsDBNull(11);
                if (checkLastDonation == true)
                {
                    lastDonation = new DateTime(2008, 5, 1, 8, 30, 52);
                }
                else
                {
                    lastDonation = reader.GetDateTime(11);
                }
                        if ((selectedDate - birthday).Days <= 23741 && isChecked == false)
                        {
                            if (gender == "мушко")
                            {
                                if ((selectedDate - lastDonation).Days >= 84)
                                {
                                    worksheet.Cell("A" + count).Value = mobilePhone;
                                    count++;
                                }
                            }
                            else if (gender == "женско")
                            {
                                if ((selectedDate - lastDonation).Days >= 112)
                                {
                                    worksheet.Cell("A" + count).Value = mobilePhone;
                                    count++;
                                }
                            }
                            else
                            {

                            }
                        }
            }
            string fileName = "/" + "SMS - " + selectedDate.Day.ToString() + "." + selectedDate.Month.ToString() + "." + selectedDate.Year.ToString() + ".xlsx";
            workbook.SaveAs(path + fileName);
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
        }
    }
}
