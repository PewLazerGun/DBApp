using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
using static DBApp.DatabasePreferences;

namespace DBApp
{
    public partial class PrintWindow1 : Form
    {
        public PrintWindow1()
        {
            InitializeComponent();
        }
        //Вывод данных на печать по заданным параметрам
        private void button1_Click(object sender, EventArgs e)
        {
            //string dbPathMyDocs = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
      
            if (label3.Visible)
                label3.Visible = false;

            SqlDataReader sqlReader = null;
            //Если на вывод 1 поле бд
            if ((string.Equals(comboBox2.Text, "1")))
            {
                if (!string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrWhiteSpace(textBox2.Text))
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM[Table] WHERE [Id]=@Id", getDb());

                    command.Parameters.AddWithValue("Id", textBox2.Text);

                    sqlReader = command.ExecuteReader();
                    sqlReader.Read();
                    //Если это Повестка
                    if ((string.Equals(comboBox1.Text, @"Повестка")))
                    {

                        var xlApplication = new Microsoft.Office.Interop.Excel.Application();
                        Workbook xlWorkbook;
                        Worksheet xlWorksheet;
                        string xlsPathMyDocs = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        string xlsPath = Path.Combine(xlsPathMyDocs, @"Повестка.xls");
                        xlWorkbook = xlApplication.Workbooks.Open(xlsPath, Type.Missing, true, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        xlWorksheet = (Worksheet)xlWorkbook.Worksheets.get_Item(1);

                        try
                        {
                            //Повестка, 1\1
                            xlApplication.Cells[1, 6] = Convert.ToString(sqlReader["Command"]);
                            xlApplication.Cells[13, 5] = Convert.ToString(sqlReader["Command"]);
                            xlApplication.Cells[14, 1] = Convert.ToString(sqlReader["MilitaryRank"]);
                            xlApplication.Cells[6, 7] = Convert.ToString(sqlReader["Name"]);
                            xlApplication.Cells[14, 6] = Convert.ToString(sqlReader["Name"]);
                            xlApplication.Cells[40, 5] = Convert.ToString(sqlReader["Name"]);
                            xlApplication.Cells[16, 7] = Convert.ToString(sqlReader["YearOfBirth"]);
                            xlApplication.Cells[2, 4] = Convert.ToString(sqlReader["VUS"]);
                            xlApplication.Cells[13, 7] = Convert.ToString(sqlReader["VUS"]);
                            xlApplication.Cells[17, 8] = Convert.ToString(sqlReader["ResidentialAddress"]);
                            xlApplication.Cells[19, 7] = Convert.ToString(sqlReader["PlaceOfWork"]);
                            xlApplication.Cells[39, 11] = Convert.ToString(sqlReader["PlaceOfWork"]);
                            xlApplication.Cells[1, 22] = Convert.ToString(sqlReader["OfficialPlot"]);
                            xlApplication.Cells[3, 16] = Convert.ToString(sqlReader["HomeRoute"]);
                            xlApplication.Cells[4, 16] = Convert.ToString(sqlReader["OfficialRoute"]);
                            xlApplication.Cells[13, 15] = Convert.ToString(sqlReader["Liter"]);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error: " + ex.Message);
                        }

                        xlApplication.Application.DisplayAlerts = false;
                        try
                        {
                            xlWorkbook.SaveAs();

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error: " + ex.Message);
                        }

                        xlApplication.Visible = true;
                        xlApplication.UserControl = true;

                        Close();
                    }
                    if ((string.Equals(comboBox1.Text, @"Мобилизационное предписание")))
                    {

                        var xlApplication = new Microsoft.Office.Interop.Excel.Application();
                        Workbook xlWorkbook;
                        Worksheet xlWorksheet;
                        string xlsPathMyDocs = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        string xlsPath = Path.Combine(xlsPathMyDocs, @"МобПред.xls");
                        xlWorkbook = xlApplication.Workbooks.Open(xlsPath, Type.Missing, true, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        xlWorksheet = (Worksheet)xlWorkbook.Worksheets.get_Item(1);

                        try
                        {
                            //МобПред 1\1
                            xlApplication.Cells[3, 7] = Convert.ToString(sqlReader["Command"]);
                            xlApplication.Cells[3, 23] = Convert.ToString(sqlReader["Command"]);
                            xlApplication.Cells[3, 40] = Convert.ToString(sqlReader["Command"]);
                            xlApplication.Cells[4, 7] = Convert.ToString(sqlReader["VUS"]);
                            xlApplication.Cells[4, 23] = Convert.ToString(sqlReader["VUS"]);
                            xlApplication.Cells[4, 40] = Convert.ToString(sqlReader["VUS"]);
                            xlApplication.Cells[7, 8] = Convert.ToString(sqlReader["MilitaryRank"]);
                            xlApplication.Cells[7, 24] = Convert.ToString(sqlReader["MilitaryRank"]);
                            xlApplication.Cells[7, 41] = Convert.ToString(sqlReader["MilitaryRank"]);
                            xlApplication.Cells[8, 11] = Convert.ToString(sqlReader["Name"]);
                            xlApplication.Cells[8, 27] = Convert.ToString(sqlReader["Name"]);
                            xlApplication.Cells[8, 44] = Convert.ToString(sqlReader["Name"]);
                            xlApplication.Cells[10, 7] = Convert.ToString(sqlReader["YearOfBirth"]);
                            xlApplication.Cells[10, 23] = Convert.ToString(sqlReader["YearOfBirth"]);
                            xlApplication.Cells[10, 40] = Convert.ToString(sqlReader["YearOfBirth"]);
                            xlApplication.Cells[6, 1] = Convert.ToString(sqlReader["MilitaryPost"]) + ", " + Convert.ToString(sqlReader["JQCode"]);
                            xlApplication.Cells[6, 17] = Convert.ToString(sqlReader["MilitaryPost"]);
                            xlApplication.Cells[6, 34] = Convert.ToString(sqlReader["MilitaryPost"]);
                            xlApplication.Cells[28, 2] = Convert.ToString(sqlReader["Day"]);
                            xlApplication.Cells[28, 5] = Convert.ToString(sqlReader["Month"]);
                            xlApplication.Cells[28, 12] = Convert.ToString(sqlReader["Year"]);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error: " + ex.Message);
                        }

                        xlApplication.Application.DisplayAlerts = false;
                        try
                        {
                            xlWorkbook.SaveAs();

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error: " + ex.Message);
                        }

                        xlApplication.Visible = true;
                        xlApplication.UserControl = true;

                        Close();
                    }
                    if ((string.Equals(comboBox1.Text, @"Приписная карта")))
                    {

                        var xlApplication = new Microsoft.Office.Interop.Excel.Application();
                        Workbook xlWorkbook;
                        Worksheet xlWorksheet;

                        string xlsPathMyDocs = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        string xlsPath = Path.Combine(xlsPathMyDocs, @"ПриписнаяКарта.xls");
                        xlWorkbook = xlApplication.Workbooks.Open(xlsPath, Type.Missing, true, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        xlWorksheet = (Worksheet)xlWorkbook.Worksheets.get_Item(1);

                        try
                        {
                            //ПрипКарта, 1\1
                            xlApplication.Cells[28, 26] = Convert.ToString(sqlReader["Command"]);
                            xlApplication.Cells[14, 11] = Convert.ToString(sqlReader["MilitaryRank"]);
                            xlApplication.Cells[30, 2] = Convert.ToString(sqlReader["Name"]);
                            xlApplication.Cells[30, 26] = Convert.ToString(sqlReader["YearOfBirth"]);
                            xlApplication.Cells[10, 10] = Convert.ToString(sqlReader["MilitaryPost"]);
                            xlApplication.Cells[8, 5] = Convert.ToString(sqlReader["byVUS"]);
                            xlApplication.Cells[10, 1] = Convert.ToString(sqlReader["JobQualification"]);
                            xlApplication.Cells[10, 7] = Convert.ToString(sqlReader["JQCode"]);
                            xlApplication.Cells[14, 1] = Convert.ToString(sqlReader["MRByRegistration"]);
                            xlApplication.Cells[14, 7] = Convert.ToString(sqlReader["MRCode"]);
                            xlApplication.Cells[17, 1] = Convert.ToString(sqlReader["TypeOfWeapon"]);
                            xlApplication.Cells[17, 7] = Convert.ToString(sqlReader["ToWCode"]);
                            xlApplication.Cells[20, 7] = Convert.ToString(sqlReader["RU"]);
                            xlApplication.Cells[23, 7] = Convert.ToString(sqlReader["CategoryOfWorkability"]);
                            xlApplication.Cells[25, 7] = Convert.ToString(sqlReader["ReserveCategory"]);
                            xlApplication.Cells[2, 19] = Convert.ToString(sqlReader["SpecialTalents"]);
                            xlApplication.Cells[8, 15] = Convert.ToString(sqlReader["OVKKKpoVUS"]);
                            xlApplication.Cells[10, 11] = Convert.ToString(sqlReader["OVKKJobQualification"]);
                            xlApplication.Cells[10, 17] = Convert.ToString(sqlReader["OVKKKJQCode"]);
                            xlApplication.Cells[14, 11] = Convert.ToString(sqlReader["OVKKKMR"]);
                            xlApplication.Cells[14, 17] = Convert.ToString(sqlReader["OVKKKMRCode"]);
                            xlApplication.Cells[17, 11] = Convert.ToString(sqlReader["OVKKKToW"]);
                            xlApplication.Cells[17, 17] = Convert.ToString(sqlReader["OVKKKToWCode"]);
                            xlApplication.Cells[32, 6] = Convert.ToString(sqlReader["Education"]);
                            xlApplication.Cells[34, 10] = Convert.ToString(sqlReader["MilitaryEducation"]);
                            xlApplication.Cells[34, 26] = Convert.ToString(sqlReader["PersonalNumber"]);
                            xlApplication.Cells[36, 13] = Convert.ToString(sqlReader["CivilSpecialty"]);
                            //Срочная служба
                            xlApplication.Cells[42, 1] = Convert.ToString(sqlReader["MilitaryUnit1"]);
                            xlApplication.Cells[43, 1] = Convert.ToString(sqlReader["MilitaryUnit2"]);
                            xlApplication.Cells[42, 4] = Convert.ToString(sqlReader["VUSCode1"]);
                            xlApplication.Cells[43, 4] = Convert.ToString(sqlReader["VUSCode2"]);
                            xlApplication.Cells[42, 7] = Convert.ToString(sqlReader["MPCode1"]);
                            xlApplication.Cells[43, 7] = Convert.ToString(sqlReader["MPCode2"]);
                            xlApplication.Cells[42, 10] = Convert.ToString(sqlReader["MRCode1"]);
                            xlApplication.Cells[43, 10] = Convert.ToString(sqlReader["MRCode2"]);
                            xlApplication.Cells[42, 13] = Convert.ToString(sqlReader["MP1"]) + " " + Convert.ToString(sqlReader["OVKKJobQualification"]);
                            xlApplication.Cells[43, 13] = Convert.ToString(sqlReader["MP2"]);
                            xlApplication.Cells[42, 17] = Convert.ToString(sqlReader["ToWCode1"]);
                            xlApplication.Cells[43, 17] = Convert.ToString(sqlReader["ToWCode1"]);
                            xlApplication.Cells[42, 20] = Convert.ToString(sqlReader["ToW1"]);
                            xlApplication.Cells[43, 20] = Convert.ToString(sqlReader["ToW2"]);
                            xlApplication.Cells[42, 24] = Convert.ToString(sqlReader["Start1"]);
                            xlApplication.Cells[43, 24] = Convert.ToString(sqlReader["Start2"]);
                            xlApplication.Cells[42, 27] = Convert.ToString(sqlReader["End1"]);
                            xlApplication.Cells[43, 27] = Convert.ToString(sqlReader["End2"]);
                            //Альтернативная служба
                            xlApplication.Cells[47, 1] = Convert.ToString(sqlReader["AlternateMU1"]);
                            xlApplication.Cells[48, 1] = Convert.ToString(sqlReader["AlternateMU2"]);
                            xlApplication.Cells[47, 9] = Convert.ToString(sqlReader["AlternatePost1"]);
                            xlApplication.Cells[48, 9] = Convert.ToString(sqlReader["AlternatePost2"]);
                            xlApplication.Cells[47, 16] = Convert.ToString(sqlReader["AlternateStart1"]);
                            xlApplication.Cells[48, 16] = Convert.ToString(sqlReader["AlternateStart2"]);
                            xlApplication.Cells[47, 23] = Convert.ToString(sqlReader["AlternateEnd1"]);
                            xlApplication.Cells[48, 23] = Convert.ToString(sqlReader["AlternateEnd2"]);
                            //Военные сборы
                            xlApplication.Cells[52, 1] = Convert.ToString(sqlReader["MCYear"]);
                            xlApplication.Cells[52, 4] = Convert.ToString(sqlReader["MCAmountOfDays"]);
                            xlApplication.Cells[52, 7] = Convert.ToString(sqlReader["MCMU"]);
                            xlApplication.Cells[52, 10] = Convert.ToString(sqlReader["MCVUSCode"]);
                            xlApplication.Cells[52, 13] = Convert.ToString(sqlReader["MCMRCode"]);
                            xlApplication.Cells[52, 16] = Convert.ToString(sqlReader["MCMP"]);
                            xlApplication.Cells[52, 23] = Convert.ToString(sqlReader["MCToWCode"]);
                            xlApplication.Cells[52, 26] = Convert.ToString(sqlReader["MCToW"]);
                            //Страница 2
                            xlApplication.Cells[55, 31] = Convert.ToString(sqlReader["PlaceOfWork"]) + ", " + Convert.ToString(sqlReader["Position"]);
                            xlApplication.Cells[57, 39] = Convert.ToString(sqlReader["ResidentialAddress"]);
                            xlApplication.Cells[60, 43] = Convert.ToString(sqlReader["MaritalStatus"]);
                            xlApplication.Cells[62, 44] = Convert.ToString(sqlReader["DocumentWasDelivered"]);
                            xlApplication.Cells[64, 40] = Convert.ToString(sqlReader["InsteadWho"]);
                            xlApplication.Cells[66, 43] = Convert.ToString(sqlReader["MedicalExamination"]);
                            xlApplication.Cells[68, 47] = Convert.ToString(sqlReader["Bacteriocarrier"]);
                            xlApplication.Cells[70, 40] = Convert.ToString(sqlReader["Hostilites"]);
                            xlApplication.Cells[72, 43] = Convert.ToString(sqlReader["AccessNumber"]);
                            xlApplication.Cells[74, 37] = Convert.ToString(sqlReader["SpecialNotes"]);
                            //Антропометрические данные
                            xlApplication.Cells[83, 31] = Convert.ToString(sqlReader["Height"]);
                            xlApplication.Cells[83, 37] = Convert.ToString(sqlReader["Headdress"]);
                            xlApplication.Cells[83, 42] = Convert.ToString(sqlReader["GasMask"]);
                            xlApplication.Cells[83, 47] = Convert.ToString(sqlReader["Outfit"]);
                            xlApplication.Cells[83, 55] = Convert.ToString(sqlReader["Shoes"]);
                            //МП выдано
                            xlApplication.Cells[89, 32] = Convert.ToString(sqlReader["Day"]);
                            xlApplication.Cells[89, 35] = Convert.ToString(sqlReader["Month"]);
                            xlApplication.Cells[89, 40] = Convert.ToString(sqlReader["Year"]);

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error: " + ex.Message);
                        }

                        xlApplication.Application.DisplayAlerts = false;
                        try
                        {
                            xlWorkbook.SaveAs();

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error: " + ex.Message);
                        }

                        xlApplication.Visible = true;
                        xlApplication.UserControl = true;

                        Close();
                    }
                    else
                    {
                        label3.Visible = true;
                        label3.Text = "Ошибка! Необходимо заполнить все поля!";
                    }
                }
            }
            else
            {
                //Если их 2
                if ((string.Equals(comboBox2.Text, "2")))
                {
                    if ((!string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrWhiteSpace(textBox2.Text))
                        && (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrWhiteSpace(textBox1.Text)))
                    {
                        SqlCommand command = new SqlCommand("SELECT * FROM[Table] WHERE [Id]=@Id", getDb());
                        SqlCommand command2 = new SqlCommand("SELECT * FROM[Table] WHERE [Id]=@Id", getDb());
                        command.Parameters.AddWithValue("Id", textBox2.Text);
                        command2.Parameters.AddWithValue("Id", textBox1.Text);
                        sqlReader = command.ExecuteReader();
                        sqlReader.Read();
                        if ((string.Equals(comboBox1.Text, @"Повестка")))
                        {

                            var xlApplication = new Microsoft.Office.Interop.Excel.Application();
                            Workbook xlWorkbook;
                            Worksheet xlWorksheet;

                            string xlsPathMyDocs = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                            string xlsPath = Path.Combine(xlsPathMyDocs, @"Повестка.xls");
                            xlWorkbook = xlApplication.Workbooks.Open(xlsPath, Type.Missing, true, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                            xlWorksheet = (Worksheet)xlWorkbook.Worksheets.get_Item(1);

                            try
                            {
                                //Повестка 1\2
                                xlApplication.Cells[1, 6] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[13, 5] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[14, 1] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[6, 7] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[14, 6] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[40, 5] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[16, 7] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[2, 4] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[13, 7] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[17, 8] = Convert.ToString(sqlReader["ResidentialAddress"]);
                                xlApplication.Cells[19, 7] = Convert.ToString(sqlReader["PlaceOfWork"]);
                                xlApplication.Cells[39, 11] = Convert.ToString(sqlReader["PlaceOfWork"]);
                                xlApplication.Cells[1, 22] = Convert.ToString(sqlReader["OfficialPlot"]);
                                xlApplication.Cells[3, 16] = Convert.ToString(sqlReader["HomeRoute"]);
                                xlApplication.Cells[4, 16] = Convert.ToString(sqlReader["OfficialRoute"]);
                                xlApplication.Cells[13, 15] = Convert.ToString(sqlReader["Liter"]);

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            sqlReader.Close();
                            sqlReader = command2.ExecuteReader();
                            sqlReader.Read();

                            try
                            {
                                //Повестка 2\2
                                xlApplication.Cells[1, 31] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[13, 30] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[14, 26] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[6, 32] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[14, 31] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[40, 30] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[16, 32] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[2, 29] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[13, 32] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[17, 33] = Convert.ToString(sqlReader["ResidentialAddress"]);
                                xlApplication.Cells[19, 32] = Convert.ToString(sqlReader["PlaceOfWork"]);
                                xlApplication.Cells[39, 36] = Convert.ToString(sqlReader["PlaceOfWork"]);
                                xlApplication.Cells[1, 47] = Convert.ToString(sqlReader["OfficialPlot"]);
                                xlApplication.Cells[3, 41] = Convert.ToString(sqlReader["HomeRoute"]);
                                xlApplication.Cells[4, 41] = Convert.ToString(sqlReader["OfficialRoute"]);
                                xlApplication.Cells[13, 40] = Convert.ToString(sqlReader["Liter"]);

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            xlApplication.Application.DisplayAlerts = false;
                            try
                            {
                                xlWorkbook.SaveAs();

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            xlApplication.Visible = true;
                            xlApplication.UserControl = true;

                            Close();
                        }
                        if ((string.Equals(comboBox1.Text, @"Мобилизационное предписание")))
                        {

                            var xlApplication = new Microsoft.Office.Interop.Excel.Application();
                            Workbook xlWorkbook;
                            Worksheet xlWorksheet;

                            string xlsPathMyDocs = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                            string xlsPath = Path.Combine(xlsPathMyDocs, @"МобПред.xls");
                            xlWorkbook = xlApplication.Workbooks.Open(xlsPath, Type.Missing, true, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                            xlWorksheet = (Worksheet)xlWorkbook.Worksheets.get_Item(1);

                            try
                            {
                                //МобПред 1\2
                                xlApplication.Cells[3, 7] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[3, 23] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[3, 40] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[4, 7] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[4, 23] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[4, 40] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[7, 8] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[7, 24] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[7, 41] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[8, 11] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[8, 27] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[8, 44] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[10, 7] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[10, 23] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[10, 40] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[6, 1] = Convert.ToString(sqlReader["MilitaryPost"]) + ", " + Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[6, 17] = Convert.ToString(sqlReader["MilitaryPost"]) + ", " + Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[6, 34] = Convert.ToString(sqlReader["MilitaryPost"]) + ", " + Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[28, 2] = Convert.ToString(sqlReader["Day"]);
                                xlApplication.Cells[28, 5] = Convert.ToString(sqlReader["Month"]);
                                xlApplication.Cells[28, 12] = Convert.ToString(sqlReader["Year"]);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            sqlReader.Close();
                            sqlReader = command2.ExecuteReader();
                            sqlReader.Read();

                            try
                            {
                                //МобПред 2\2
                                xlApplication.Cells[32, 7] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[32, 23] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[32, 40] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[33, 7] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[33, 23] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[33, 40] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[36, 8] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[36, 24] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[36, 41] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[37, 11] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[37, 27] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[37, 44] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[39, 7] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[39, 23] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[39, 40] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[35, 1] = Convert.ToString(sqlReader["MilitaryPost"]) + ", " + Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[35, 17] = Convert.ToString(sqlReader["MilitaryPost"]) + ", " + Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[35, 34] = Convert.ToString(sqlReader["MilitaryPost"]) + ", " + Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[57, 2] = Convert.ToString(sqlReader["Day"]);
                                xlApplication.Cells[57, 5] = Convert.ToString(sqlReader["Month"]);
                                xlApplication.Cells[57, 12] = Convert.ToString(sqlReader["Year"]);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            xlApplication.Application.DisplayAlerts = false;
                            try
                            {
                                xlWorkbook.SaveAs();

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            xlApplication.Visible = true;
                            xlApplication.UserControl = true;

                            Close();
                        }
                        if ((string.Equals(comboBox1.Text, @"Приписная карта")))
                        {

                            var xlApplication = new Microsoft.Office.Interop.Excel.Application();
                            Workbook xlWorkbook;
                            Worksheet xlWorksheet;

                            string xlsPathMyDocs = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                            string xlsPath = Path.Combine(xlsPathMyDocs, @"ПриписнаяКарта.xls");
                            xlWorkbook = xlApplication.Workbooks.Open(xlsPath, Type.Missing, true, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                            xlWorksheet = (Worksheet)xlWorkbook.Worksheets.get_Item(1);

                            try
                            {
                                //ПрипКарта 1\2
                                xlApplication.Cells[28, 26] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[14, 11] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[30, 2] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[30, 26] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[10, 1] = Convert.ToString(sqlReader["MilitaryPost"]);
                                xlApplication.Cells[10, 10] = Convert.ToString(sqlReader["MilitaryPost"]);
                                xlApplication.Cells[55, 31] = Convert.ToString(sqlReader["PlaceOfWork"]);
                                xlApplication.Cells[58, 31] = Convert.ToString(sqlReader["ResidentialAddress"]);
                                xlApplication.Cells[8, 5] = Convert.ToString(sqlReader["byVUS"]);
                                xlApplication.Cells[10, 1] = Convert.ToString(sqlReader["JobQualification"]);
                                xlApplication.Cells[10, 7] = Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[14, 1] = Convert.ToString(sqlReader["MRByRegistration"]);
                                xlApplication.Cells[14, 7] = Convert.ToString(sqlReader["MRCode"]);
                                xlApplication.Cells[17, 1] = Convert.ToString(sqlReader["TypeOfWeapon"]);
                                xlApplication.Cells[17, 7] = Convert.ToString(sqlReader["ToWCode"]);
                                xlApplication.Cells[20, 7] = Convert.ToString(sqlReader["RU"]);
                                xlApplication.Cells[23, 7] = Convert.ToString(sqlReader["CategoryOfWorkability"]);
                                xlApplication.Cells[25, 7] = Convert.ToString(sqlReader["ReserveCategory"]);
                                xlApplication.Cells[2, 19] = Convert.ToString(sqlReader["SpecialTalents"]);
                                xlApplication.Cells[8, 15] = Convert.ToString(sqlReader["OVKKKpoVUS"]);
                                xlApplication.Cells[10, 11] = Convert.ToString(sqlReader["OVKKJobQualification"]);
                                xlApplication.Cells[10, 17] = Convert.ToString(sqlReader["OVKKKJQCode"]);
                                xlApplication.Cells[14, 11] = Convert.ToString(sqlReader["OVKKKMR"]);
                                xlApplication.Cells[14, 17] = Convert.ToString(sqlReader["OVKKKMRCode"]);
                                xlApplication.Cells[17, 11] = Convert.ToString(sqlReader["OVKKKToW"]);
                                xlApplication.Cells[17, 17] = Convert.ToString(sqlReader["OVKKKToWCode"]);
                                xlApplication.Cells[32, 6] = Convert.ToString(sqlReader["Education"]);
                                xlApplication.Cells[34, 10] = Convert.ToString(sqlReader["MilitaryEducation"]);
                                xlApplication.Cells[34, 26] = Convert.ToString(sqlReader["PersonalNumber"]);
                                xlApplication.Cells[36, 13] = Convert.ToString(sqlReader["CivilSpecialty"]);
                                //Срочная служба
                                xlApplication.Cells[42, 1] = Convert.ToString(sqlReader["MilitaryUnit1"]);
                                xlApplication.Cells[43, 1] = Convert.ToString(sqlReader["MilitaryUnit2"]);
                                xlApplication.Cells[42, 4] = Convert.ToString(sqlReader["VUSCode1"]);
                                xlApplication.Cells[43, 4] = Convert.ToString(sqlReader["VUSCode2"]);
                                xlApplication.Cells[42, 7] = Convert.ToString(sqlReader["MPCode1"]);
                                xlApplication.Cells[43, 7] = Convert.ToString(sqlReader["MPCode2"]);
                                xlApplication.Cells[42, 10] = Convert.ToString(sqlReader["MRCode1"]);
                                xlApplication.Cells[43, 10] = Convert.ToString(sqlReader["MRCode2"]);
                                xlApplication.Cells[42, 13] = Convert.ToString(sqlReader["MP1"]) + " " + Convert.ToString(sqlReader["OVKKJobQualification"]);
                                xlApplication.Cells[43, 13] = Convert.ToString(sqlReader["MP2"]);
                                xlApplication.Cells[42, 17] = Convert.ToString(sqlReader["ToWCode1"]);
                                xlApplication.Cells[43, 17] = Convert.ToString(sqlReader["ToWCode1"]);
                                xlApplication.Cells[42, 20] = Convert.ToString(sqlReader["ToW1"]);
                                xlApplication.Cells[43, 20] = Convert.ToString(sqlReader["ToW2"]);
                                xlApplication.Cells[42, 24] = Convert.ToString(sqlReader["Start1"]);
                                xlApplication.Cells[43, 24] = Convert.ToString(sqlReader["Start2"]);
                                xlApplication.Cells[42, 27] = Convert.ToString(sqlReader["End1"]);
                                xlApplication.Cells[43, 27] = Convert.ToString(sqlReader["End2"]);
                                //Альтернативная служба
                                xlApplication.Cells[47, 1] = Convert.ToString(sqlReader["AlternateMU1"]);
                                xlApplication.Cells[48, 1] = Convert.ToString(sqlReader["AlternateMU2"]);
                                xlApplication.Cells[47, 9] = Convert.ToString(sqlReader["AlternatePost1"]);
                                xlApplication.Cells[48, 9] = Convert.ToString(sqlReader["AlternatePost2"]);
                                xlApplication.Cells[47, 16] = Convert.ToString(sqlReader["AlternateStart1"]);
                                xlApplication.Cells[48, 16] = Convert.ToString(sqlReader["AlternateStart2"]);
                                xlApplication.Cells[47, 23] = Convert.ToString(sqlReader["AlternateEnd1"]);
                                xlApplication.Cells[48, 23] = Convert.ToString(sqlReader["AlternateEnd2"]);
                                //Военные сборы
                                xlApplication.Cells[52, 1] = Convert.ToString(sqlReader["MCYear"]);
                                xlApplication.Cells[52, 4] = Convert.ToString(sqlReader["MCAmountOfDays"]);
                                xlApplication.Cells[52, 7] = Convert.ToString(sqlReader["MCMU"]);
                                xlApplication.Cells[52, 10] = Convert.ToString(sqlReader["MCVUSCode"]);
                                xlApplication.Cells[52, 13] = Convert.ToString(sqlReader["MCMRCode"]);
                                xlApplication.Cells[52, 16] = Convert.ToString(sqlReader["MCMP"]);
                                xlApplication.Cells[52, 23] = Convert.ToString(sqlReader["MCToWCode"]);
                                xlApplication.Cells[52, 26] = Convert.ToString(sqlReader["MCToW"]);
                                //Страница 2
                                xlApplication.Cells[55, 31] = Convert.ToString(sqlReader["PlaceOfWork"]) + ", " + Convert.ToString(sqlReader["Position"]);
                                xlApplication.Cells[57, 39] = Convert.ToString(sqlReader["ResidentialAddress"]);
                                xlApplication.Cells[60, 43] = Convert.ToString(sqlReader["MaritalStatus"]);
                                xlApplication.Cells[62, 44] = Convert.ToString(sqlReader["DocumentWasDelivered"]);
                                xlApplication.Cells[64, 40] = Convert.ToString(sqlReader["InsteadWho"]);
                                xlApplication.Cells[66, 43] = Convert.ToString(sqlReader["MedicalExamination"]);
                                xlApplication.Cells[68, 47] = Convert.ToString(sqlReader["Bacteriocarrier"]);
                                xlApplication.Cells[70, 40] = Convert.ToString(sqlReader["Hostilites"]);
                                xlApplication.Cells[72, 43] = Convert.ToString(sqlReader["AccessNumber"]);
                                xlApplication.Cells[74, 37] = Convert.ToString(sqlReader["SpecialNotes"]);
                                //Антропометрические данные
                                xlApplication.Cells[83, 31] = Convert.ToString(sqlReader["Height"]);
                                xlApplication.Cells[83, 37] = Convert.ToString(sqlReader["Headdress"]);
                                xlApplication.Cells[83, 42] = Convert.ToString(sqlReader["GasMask"]);
                                xlApplication.Cells[83, 47] = Convert.ToString(sqlReader["Outfit"]);
                                xlApplication.Cells[83, 55] = Convert.ToString(sqlReader["Shoes"]);
                                //МП выдано
                                xlApplication.Cells[89, 32] = Convert.ToString(sqlReader["Day"]);
                                xlApplication.Cells[89, 35] = Convert.ToString(sqlReader["Month"]);
                                xlApplication.Cells[89, 40] = Convert.ToString(sqlReader["Year"]);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            sqlReader.Close();
                            sqlReader = command2.ExecuteReader();
                            sqlReader.Read();

                            try
                            {
                                //ПрипКарта 2\2
                                xlApplication.Cells[28, 56] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[14, 41] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[30, 32] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[30, 56] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[10, 31] = Convert.ToString(sqlReader["MilitaryPost"]);
                                xlApplication.Cells[10, 40] = Convert.ToString(sqlReader["MilitaryPost"]);
                                xlApplication.Cells[55, 1] = Convert.ToString(sqlReader["PlaceOfWork"]);
                                xlApplication.Cells[58, 1] = Convert.ToString(sqlReader["ResidentialAddress"]);
                                xlApplication.Cells[8, 35] = Convert.ToString(sqlReader["byVUS"]);
                                xlApplication.Cells[10, 31] = Convert.ToString(sqlReader["JobQualification"]);
                                xlApplication.Cells[10, 37] = Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[14, 31] = Convert.ToString(sqlReader["MRByRegistration"]);
                                xlApplication.Cells[14, 37] = Convert.ToString(sqlReader["MRCode"]);
                                xlApplication.Cells[17, 31] = Convert.ToString(sqlReader["TypeOfWeapon"]);
                                xlApplication.Cells[17, 37] = Convert.ToString(sqlReader["ToWCode"]);
                                xlApplication.Cells[20, 37] = Convert.ToString(sqlReader["RU"]);
                                xlApplication.Cells[23, 37] = Convert.ToString(sqlReader["CategoryOfWorkability"]);
                                xlApplication.Cells[25, 37] = Convert.ToString(sqlReader["ReserveCategory"]);
                                xlApplication.Cells[2, 49] = Convert.ToString(sqlReader["SpecialTalents"]);
                                xlApplication.Cells[8, 45] = Convert.ToString(sqlReader["OVKKKpoVUS"]);
                                xlApplication.Cells[10, 41] = Convert.ToString(sqlReader["OVKKJobQualification"]);
                                xlApplication.Cells[10, 47] = Convert.ToString(sqlReader["OVKKKJQCode"]);
                                xlApplication.Cells[14, 41] = Convert.ToString(sqlReader["OVKKKMR"]);
                                xlApplication.Cells[14, 47] = Convert.ToString(sqlReader["OVKKKMRCode"]);
                                xlApplication.Cells[17, 41] = Convert.ToString(sqlReader["OVKKKToW"]);
                                xlApplication.Cells[17, 47] = Convert.ToString(sqlReader["OVKKKToWCode"]);
                                xlApplication.Cells[32, 36] = Convert.ToString(sqlReader["Education"]);
                                xlApplication.Cells[34, 40] = Convert.ToString(sqlReader["MilitaryEducation"]);
                                xlApplication.Cells[34, 56] = Convert.ToString(sqlReader["PersonalNumber"]);
                                xlApplication.Cells[36, 43] = Convert.ToString(sqlReader["CivilSpecialty"]);
                                //Срочная служба
                                xlApplication.Cells[42, 31] = Convert.ToString(sqlReader["MilitaryUnit1"]);
                                xlApplication.Cells[43, 31] = Convert.ToString(sqlReader["MilitaryUnit2"]);
                                xlApplication.Cells[42, 34] = Convert.ToString(sqlReader["VUSCode1"]);
                                xlApplication.Cells[43, 34] = Convert.ToString(sqlReader["VUSCode2"]);
                                xlApplication.Cells[42, 37] = Convert.ToString(sqlReader["MPCode1"]);
                                xlApplication.Cells[43, 37] = Convert.ToString(sqlReader["MPCode2"]);
                                xlApplication.Cells[42, 40] = Convert.ToString(sqlReader["MRCode1"]);
                                xlApplication.Cells[43, 40] = Convert.ToString(sqlReader["MRCode2"]);
                                xlApplication.Cells[42, 43] = Convert.ToString(sqlReader["MP1"]) + " " + Convert.ToString(sqlReader["OVKKJobQualification"]);
                                xlApplication.Cells[43, 43] = Convert.ToString(sqlReader["MP2"]);
                                xlApplication.Cells[42, 47] = Convert.ToString(sqlReader["ToWCode1"]);
                                xlApplication.Cells[43, 47] = Convert.ToString(sqlReader["ToWCode1"]);
                                xlApplication.Cells[42, 50] = Convert.ToString(sqlReader["ToW1"]);
                                xlApplication.Cells[43, 50] = Convert.ToString(sqlReader["ToW2"]);
                                xlApplication.Cells[42, 54] = Convert.ToString(sqlReader["Start1"]);
                                xlApplication.Cells[43, 54] = Convert.ToString(sqlReader["Start2"]);
                                xlApplication.Cells[42, 57] = Convert.ToString(sqlReader["End1"]);
                                xlApplication.Cells[43, 57] = Convert.ToString(sqlReader["End2"]);
                                //Альтернативная служба
                                xlApplication.Cells[47, 31] = Convert.ToString(sqlReader["AlternateMU1"]);
                                xlApplication.Cells[48, 31] = Convert.ToString(sqlReader["AlternateMU2"]);
                                xlApplication.Cells[47, 38] = Convert.ToString(sqlReader["AlternatePost1"]);
                                xlApplication.Cells[48, 38] = Convert.ToString(sqlReader["AlternatePost2"]);
                                xlApplication.Cells[47, 46] = Convert.ToString(sqlReader["AlternateStart1"]);
                                xlApplication.Cells[48, 46] = Convert.ToString(sqlReader["AlternateStart2"]);
                                xlApplication.Cells[47, 53] = Convert.ToString(sqlReader["AlternateEnd1"]);
                                xlApplication.Cells[48, 53] = Convert.ToString(sqlReader["AlternateEnd2"]);
                                //Военные сборы
                                xlApplication.Cells[52, 31] = Convert.ToString(sqlReader["MCYear"]);
                                xlApplication.Cells[52, 34] = Convert.ToString(sqlReader["MCAmountOfDays"]);
                                xlApplication.Cells[52, 37] = Convert.ToString(sqlReader["MCMU"]);
                                xlApplication.Cells[52, 40] = Convert.ToString(sqlReader["MCVUSCode"]);
                                xlApplication.Cells[52, 43] = Convert.ToString(sqlReader["MCMRCode"]);
                                xlApplication.Cells[52, 46] = Convert.ToString(sqlReader["MCMP"]);
                                xlApplication.Cells[52, 53] = Convert.ToString(sqlReader["MCToWCode"]);
                                xlApplication.Cells[52, 56] = Convert.ToString(sqlReader["MCToW"]);
                                //Страница 2
                                xlApplication.Cells[55, 1] = Convert.ToString(sqlReader["PlaceOfWork"]) + ", " + Convert.ToString(sqlReader["Position"]);
                                xlApplication.Cells[57, 9] = Convert.ToString(sqlReader["ResidentialAddress"]);
                                xlApplication.Cells[60, 13] = Convert.ToString(sqlReader["MaritalStatus"]);
                                xlApplication.Cells[62, 14] = Convert.ToString(sqlReader["DocumentWasDelivered"]);
                                xlApplication.Cells[64, 10] = Convert.ToString(sqlReader["InsteadWho"]);
                                xlApplication.Cells[66, 13] = Convert.ToString(sqlReader["MedicalExamination"]);
                                xlApplication.Cells[68, 17] = Convert.ToString(sqlReader["Bacteriocarrier"]);
                                xlApplication.Cells[70, 10] = Convert.ToString(sqlReader["Hostilites"]);
                                xlApplication.Cells[72, 13] = Convert.ToString(sqlReader["AccessNumber"]);
                                xlApplication.Cells[74, 7] = Convert.ToString(sqlReader["SpecialNotes"]);
                                //Антропометрические данные
                                xlApplication.Cells[83, 1] = Convert.ToString(sqlReader["Height"]);
                                xlApplication.Cells[83, 7] = Convert.ToString(sqlReader["Headdress"]);
                                xlApplication.Cells[83, 12] = Convert.ToString(sqlReader["GasMask"]);
                                xlApplication.Cells[83, 17] = Convert.ToString(sqlReader["Outfit"]);
                                xlApplication.Cells[83, 25] = Convert.ToString(sqlReader["Shoes"]);
                                //МП выдано
                                xlApplication.Cells[89, 2] = Convert.ToString(sqlReader["Day"]);
                                xlApplication.Cells[89, 5] = Convert.ToString(sqlReader["Month"]);
                                xlApplication.Cells[89, 10] = Convert.ToString(sqlReader["Year"]);

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            xlApplication.Application.DisplayAlerts = false;
                            try
                            {
                                xlWorkbook.SaveAs();

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            xlApplication.Visible = true;
                            xlApplication.UserControl = true;

                            Close();
                        }
                        else
                        {
                            label3.Visible = true;
                            label3.Text = "Ошибка! Необходимо заполнить все поля!";
                        }
                    }
                }
                //Иначе (Если их 3)
                else
                {
                    if ((!string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrWhiteSpace(textBox2.Text))
                        && (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrWhiteSpace(textBox1.Text))
                        && (!string.IsNullOrEmpty(textBox3.Text) && !string.IsNullOrWhiteSpace(textBox3.Text)))
                    {
                        SqlCommand command = new SqlCommand("SELECT * FROM[Table] WHERE [Id]=@Id", getDb());
                        SqlCommand command2 = new SqlCommand("SELECT * FROM[Table] WHERE [Id]=@Id", getDb());
                        SqlCommand command3 = new SqlCommand("SELECT * FROM[Table] WHERE [Id]=@Id", getDb());
                        command.Parameters.AddWithValue("Id", textBox2.Text);
                        command2.Parameters.AddWithValue("Id", textBox1.Text);
                        command3.Parameters.AddWithValue("Id", textBox3.Text);
                        sqlReader = command.ExecuteReader();
                        sqlReader.Read();
                        if ((string.Equals(comboBox1.Text, @"Повестка")))
                        {

                            var xlApplication = new Microsoft.Office.Interop.Excel.Application();
                            Workbook xlWorkbook;
                            Worksheet xlWorksheet;

                            string xlsPathMyDocs = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                            string xlsPath = Path.Combine(xlsPathMyDocs, @"Повестка.xls");
                            xlWorkbook = xlApplication.Workbooks.Open(xlsPath, Type.Missing, true, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            xlWorksheet = (Worksheet)xlWorkbook.Worksheets.get_Item(1);

                            try
                            {
                                //Повестка 1\3
                                xlApplication.Cells[1, 6] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[13, 5] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[14, 1] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[6, 7] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[14, 6] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[40, 5] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[16, 7] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[2, 4] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[13, 7] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[17, 8] = Convert.ToString(sqlReader["ResidentialAddress"]);
                                xlApplication.Cells[19, 7] = Convert.ToString(sqlReader["PlaceOfWork"]);
                                xlApplication.Cells[39, 11] = Convert.ToString(sqlReader["PlaceOfWork"]);
                                xlApplication.Cells[1, 22] = Convert.ToString(sqlReader["OfficialPlot"]);
                                xlApplication.Cells[3, 16] = Convert.ToString(sqlReader["HomeRoute"]);
                                xlApplication.Cells[4, 16] = Convert.ToString(sqlReader["OfficialRoute"]);
                                xlApplication.Cells[13, 15] = Convert.ToString(sqlReader["Liter"]);

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            sqlReader.Close();
                            sqlReader = command2.ExecuteReader();
                            sqlReader.Read();

                            try
                            {
                                //Повестка 2\3
                                xlApplication.Cells[1, 31] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[13, 30] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[14, 26] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[6, 32] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[14, 31] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[40, 30] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[16, 32] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[2, 29] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[13, 32] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[17, 33] = Convert.ToString(sqlReader["ResidentialAddress"]);
                                xlApplication.Cells[19, 32] = Convert.ToString(sqlReader["PlaceOfWork"]);
                                xlApplication.Cells[39, 36] = Convert.ToString(sqlReader["PlaceOfWork"]);
                                xlApplication.Cells[1, 47] = Convert.ToString(sqlReader["OfficialPlot"]);
                                xlApplication.Cells[3, 41] = Convert.ToString(sqlReader["HomeRoute"]);
                                xlApplication.Cells[4, 41] = Convert.ToString(sqlReader["OfficialRoute"]);
                                xlApplication.Cells[13, 40] = Convert.ToString(sqlReader["Liter"]);

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            sqlReader.Close();
                            sqlReader = command3.ExecuteReader();
                            sqlReader.Read();

                            try
                            {
                                //Повестка 3\3
                                xlApplication.Cells[1, 56] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[13, 55] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[14, 51] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[6, 57] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[14, 56] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[40, 55] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[16, 57] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[2, 54] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[13, 57] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[17, 58] = Convert.ToString(sqlReader["ResidentialAddress"]);
                                xlApplication.Cells[19, 57] = Convert.ToString(sqlReader["PlaceOfWork"]);
                                xlApplication.Cells[39, 61] = Convert.ToString(sqlReader["PlaceOfWork"]);
                                xlApplication.Cells[1, 72] = Convert.ToString(sqlReader["OfficialPlot"]);
                                xlApplication.Cells[3, 66] = Convert.ToString(sqlReader["HomeRoute"]);
                                xlApplication.Cells[4, 66] = Convert.ToString(sqlReader["OfficialRoute"]);
                                xlApplication.Cells[13, 65] = Convert.ToString(sqlReader["Liter"]);

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            xlApplication.Application.DisplayAlerts = false;
                            try
                            {
                                xlWorkbook.SaveAs();

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            xlApplication.Visible = true;
                            xlApplication.UserControl = true;

                            Close();
                        }
                        if ((string.Equals(comboBox1.Text, @"Мобилизационное предписание")))
                        {

                            var xlApplication = new Microsoft.Office.Interop.Excel.Application();
                            Workbook xlWorkbook;
                            Worksheet xlWorksheet;

                            string xlsPathMyDocs = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                            string xlsPath = Path.Combine(xlsPathMyDocs, @"МобПред.xls");
                            xlWorkbook = xlApplication.Workbooks.Open(xlsPath, Type.Missing, true, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            xlWorksheet = (Worksheet)xlWorkbook.Worksheets.get_Item(1);

                            try
                            {
                                //МобПред 1\3
                                xlApplication.Cells[3, 7] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[3, 23] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[3, 40] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[4, 7] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[4, 23] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[4, 40] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[7, 8] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[7, 24] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[7, 41] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[8, 11] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[8, 27] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[8, 44] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[10, 7] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[10, 23] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[10, 40] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[6, 1] = Convert.ToString(sqlReader["MilitaryPost"]) + ", " + Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[6, 17] = Convert.ToString(sqlReader["MilitaryPost"]) + ", " + Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[6, 34] = Convert.ToString(sqlReader["MilitaryPost"]) + ", " + Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[28, 2] = Convert.ToString(sqlReader["Day"]);
                                xlApplication.Cells[28, 5] = Convert.ToString(sqlReader["Month"]);
                                xlApplication.Cells[28, 12] = Convert.ToString(sqlReader["Year"]);

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            sqlReader.Close();
                            sqlReader = command2.ExecuteReader();
                            sqlReader.Read();

                            try
                            {
                                //МобПред 2\3
                                xlApplication.Cells[32, 7] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[32, 23] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[32, 40] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[33, 7] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[33, 23] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[33, 40] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[36, 8] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[36, 24] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[36, 41] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[37, 11] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[37, 27] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[37, 44] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[39, 7] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[39, 23] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[39, 40] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[35, 1] = Convert.ToString(sqlReader["MilitaryPost"]) + ", " + Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[35, 17] = Convert.ToString(sqlReader["MilitaryPost"]) + ", " + Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[35, 34] = Convert.ToString(sqlReader["MilitaryPost"]) + ", " + Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[57, 2] = Convert.ToString(sqlReader["Day"]);
                                xlApplication.Cells[57, 5] = Convert.ToString(sqlReader["Month"]);
                                xlApplication.Cells[57, 12] = Convert.ToString(sqlReader["Year"]);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            sqlReader.Close();
                            sqlReader = command3.ExecuteReader();
                            sqlReader.Read();

                            try
                            {
                                //МобПред 3\3
                                xlApplication.Cells[61, 7] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[61, 23] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[61, 40] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[62, 7] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[62, 23] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[62, 40] = Convert.ToString(sqlReader["VUS"]);
                                xlApplication.Cells[65, 8] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[65, 24] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[65, 41] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[66, 11] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[66, 27] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[66, 44] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[68, 7] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[68, 23] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[68, 40] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[64, 1] = Convert.ToString(sqlReader["MilitaryPost"]) + ", " + Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[64, 17] = Convert.ToString(sqlReader["MilitaryPost"]) + ", " + Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[64, 34] = Convert.ToString(sqlReader["MilitaryPost"]) + ", " + Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[86, 2] = Convert.ToString(sqlReader["Day"]);
                                xlApplication.Cells[86, 5] = Convert.ToString(sqlReader["Month"]);
                                xlApplication.Cells[86, 12] = Convert.ToString(sqlReader["Year"]);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            xlApplication.Application.DisplayAlerts = false;
                            try
                            {
                                xlWorkbook.SaveAs();

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            xlApplication.Visible = true;
                            xlApplication.UserControl = true;

                            Close();
                        }
                        if ((string.Equals(comboBox1.Text, @"Приписная карта")))
                        {

                            var xlApplication = new Microsoft.Office.Interop.Excel.Application();
                            Workbook xlWorkbook;
                            Worksheet xlWorksheet;

                            string xlsPathMyDocs = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                            string xlsPath = Path.Combine(xlsPathMyDocs, @"ПриписнаяКарта.xls");
                            xlWorkbook = xlApplication.Workbooks.Open(xlsPath, Type.Missing, true, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                            xlWorksheet = (Worksheet)xlWorkbook.Worksheets.get_Item(1);

                            try
                            {
                                //ПрипКарта 1\3 (2)
                                xlApplication.Cells[28, 26] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[14, 11] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[30, 2] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[30, 26] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[10, 1] = Convert.ToString(sqlReader["MilitaryPost"]);
                                xlApplication.Cells[10, 10] = Convert.ToString(sqlReader["MilitaryPost"]);
                                xlApplication.Cells[55, 31] = Convert.ToString(sqlReader["PlaceOfWork"]);
                                xlApplication.Cells[58, 31] = Convert.ToString(sqlReader["ResidentialAddress"]);
                                xlApplication.Cells[8, 5] = Convert.ToString(sqlReader["byVUS"]);
                                xlApplication.Cells[10, 1] = Convert.ToString(sqlReader["JobQualification"]);
                                xlApplication.Cells[10, 7] = Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[14, 1] = Convert.ToString(sqlReader["MRByRegistration"]);
                                xlApplication.Cells[14, 7] = Convert.ToString(sqlReader["MRCode"]);
                                xlApplication.Cells[17, 1] = Convert.ToString(sqlReader["TypeOfWeapon"]);
                                xlApplication.Cells[17, 7] = Convert.ToString(sqlReader["ToWCode"]);
                                xlApplication.Cells[20, 7] = Convert.ToString(sqlReader["RU"]);
                                xlApplication.Cells[23, 7] = Convert.ToString(sqlReader["CategoryOfWorkability"]);
                                xlApplication.Cells[25, 7] = Convert.ToString(sqlReader["ReserveCategory"]);
                                xlApplication.Cells[2, 19] = Convert.ToString(sqlReader["SpecialTalents"]);
                                xlApplication.Cells[8, 15] = Convert.ToString(sqlReader["OVKKKpoVUS"]);
                                xlApplication.Cells[10, 11] = Convert.ToString(sqlReader["OVKKJobQualification"]);
                                xlApplication.Cells[10, 17] = Convert.ToString(sqlReader["OVKKKJQCode"]);
                                xlApplication.Cells[14, 11] = Convert.ToString(sqlReader["OVKKKMR"]);
                                xlApplication.Cells[14, 17] = Convert.ToString(sqlReader["OVKKKMRCode"]);
                                xlApplication.Cells[17, 11] = Convert.ToString(sqlReader["OVKKKToW"]);
                                xlApplication.Cells[17, 17] = Convert.ToString(sqlReader["OVKKKToWCode"]);
                                xlApplication.Cells[32, 6] = Convert.ToString(sqlReader["Education"]);
                                xlApplication.Cells[34, 10] = Convert.ToString(sqlReader["MilitaryEducation"]);
                                xlApplication.Cells[34, 26] = Convert.ToString(sqlReader["PersonalNumber"]);
                                xlApplication.Cells[36, 13] = Convert.ToString(sqlReader["CivilSpecialty"]);
                                //Срочная служба
                                xlApplication.Cells[42, 1] = Convert.ToString(sqlReader["MilitaryUnit1"]);
                                xlApplication.Cells[43, 1] = Convert.ToString(sqlReader["MilitaryUnit2"]);
                                xlApplication.Cells[42, 4] = Convert.ToString(sqlReader["VUSCode1"]);
                                xlApplication.Cells[43, 4] = Convert.ToString(sqlReader["VUSCode2"]);
                                xlApplication.Cells[42, 7] = Convert.ToString(sqlReader["MPCode1"]);
                                xlApplication.Cells[43, 7] = Convert.ToString(sqlReader["MPCode2"]);
                                xlApplication.Cells[42, 10] = Convert.ToString(sqlReader["MRCode1"]);
                                xlApplication.Cells[43, 10] = Convert.ToString(sqlReader["MRCode2"]);
                                xlApplication.Cells[42, 13] = Convert.ToString(sqlReader["MP1"]) + " " + Convert.ToString(sqlReader["OVKKJobQualification"]);
                                xlApplication.Cells[43, 13] = Convert.ToString(sqlReader["MP2"]);
                                xlApplication.Cells[42, 17] = Convert.ToString(sqlReader["ToWCode1"]);
                                xlApplication.Cells[43, 17] = Convert.ToString(sqlReader["ToWCode1"]);
                                xlApplication.Cells[42, 20] = Convert.ToString(sqlReader["ToW1"]);
                                xlApplication.Cells[43, 20] = Convert.ToString(sqlReader["ToW2"]);
                                xlApplication.Cells[42, 24] = Convert.ToString(sqlReader["Start1"]);
                                xlApplication.Cells[43, 24] = Convert.ToString(sqlReader["Start2"]);
                                xlApplication.Cells[42, 27] = Convert.ToString(sqlReader["End1"]);
                                xlApplication.Cells[43, 27] = Convert.ToString(sqlReader["End2"]);
                                //Альтернативная служба
                                xlApplication.Cells[47, 1] = Convert.ToString(sqlReader["AlternateMU1"]);
                                xlApplication.Cells[48, 1] = Convert.ToString(sqlReader["AlternateMU2"]);
                                xlApplication.Cells[47, 9] = Convert.ToString(sqlReader["AlternatePost1"]);
                                xlApplication.Cells[48, 9] = Convert.ToString(sqlReader["AlternatePost2"]);
                                xlApplication.Cells[47, 16] = Convert.ToString(sqlReader["AlternateStart1"]);
                                xlApplication.Cells[48, 16] = Convert.ToString(sqlReader["AlternateStart2"]);
                                xlApplication.Cells[47, 23] = Convert.ToString(sqlReader["AlternateEnd1"]);
                                xlApplication.Cells[48, 23] = Convert.ToString(sqlReader["AlternateEnd2"]);
                                //Военные сборы
                                xlApplication.Cells[52, 1] = Convert.ToString(sqlReader["MCYear"]);
                                xlApplication.Cells[52, 4] = Convert.ToString(sqlReader["MCAmountOfDays"]);
                                xlApplication.Cells[52, 7] = Convert.ToString(sqlReader["MCMU"]);
                                xlApplication.Cells[52, 10] = Convert.ToString(sqlReader["MCVUSCode"]);
                                xlApplication.Cells[52, 13] = Convert.ToString(sqlReader["MCMRCode"]);
                                xlApplication.Cells[52, 16] = Convert.ToString(sqlReader["MCMP"]);
                                xlApplication.Cells[52, 23] = Convert.ToString(sqlReader["MCToWCode"]);
                                xlApplication.Cells[52, 26] = Convert.ToString(sqlReader["MCToW"]);
                                //Страница 2
                                xlApplication.Cells[55, 31] = Convert.ToString(sqlReader["PlaceOfWork"]) + ", " + Convert.ToString(sqlReader["Position"]);
                                xlApplication.Cells[57, 39] = Convert.ToString(sqlReader["ResidentialAddress"]);
                                xlApplication.Cells[60, 43] = Convert.ToString(sqlReader["MaritalStatus"]);
                                xlApplication.Cells[62, 44] = Convert.ToString(sqlReader["DocumentWasDelivered"]);
                                xlApplication.Cells[64, 40] = Convert.ToString(sqlReader["InsteadWho"]);
                                xlApplication.Cells[66, 43] = Convert.ToString(sqlReader["MedicalExamination"]);
                                xlApplication.Cells[68, 47] = Convert.ToString(sqlReader["Bacteriocarrier"]);
                                xlApplication.Cells[70, 40] = Convert.ToString(sqlReader["Hostilites"]);
                                xlApplication.Cells[72, 43] = Convert.ToString(sqlReader["AccessNumber"]);
                                xlApplication.Cells[74, 37] = Convert.ToString(sqlReader["SpecialNotes"]);
                                //Антропометрические данные
                                xlApplication.Cells[83, 31] = Convert.ToString(sqlReader["Height"]);
                                xlApplication.Cells[83, 37] = Convert.ToString(sqlReader["Headdress"]);
                                xlApplication.Cells[83, 42] = Convert.ToString(sqlReader["GasMask"]);
                                xlApplication.Cells[83, 47] = Convert.ToString(sqlReader["Outfit"]);
                                xlApplication.Cells[83, 55] = Convert.ToString(sqlReader["Shoes"]);
                                //МП выдано
                                xlApplication.Cells[89, 32] = Convert.ToString(sqlReader["Day"]);
                                xlApplication.Cells[89, 35] = Convert.ToString(sqlReader["Month"]);
                                xlApplication.Cells[89, 40] = Convert.ToString(sqlReader["Year"]);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            sqlReader.Close();
                            sqlReader = command2.ExecuteReader();
                            sqlReader.Read();

                            try
                            {
                                //ПрипКарта 2\3 (2)
                                xlApplication.Cells[28, 56] = Convert.ToString(sqlReader["Command"]);
                                xlApplication.Cells[14, 41] = Convert.ToString(sqlReader["MilitaryRank"]);
                                xlApplication.Cells[30, 32] = Convert.ToString(sqlReader["Name"]);
                                xlApplication.Cells[30, 56] = Convert.ToString(sqlReader["YearOfBirth"]);
                                xlApplication.Cells[10, 31] = Convert.ToString(sqlReader["MilitaryPost"]);
                                xlApplication.Cells[10, 40] = Convert.ToString(sqlReader["MilitaryPost"]);
                                xlApplication.Cells[55, 1] = Convert.ToString(sqlReader["PlaceOfWork"]);
                                xlApplication.Cells[58, 1] = Convert.ToString(sqlReader["ResidentialAddress"]);
                                xlApplication.Cells[8, 35] = Convert.ToString(sqlReader["byVUS"]);
                                xlApplication.Cells[10, 31] = Convert.ToString(sqlReader["JobQualification"]);
                                xlApplication.Cells[10, 37] = Convert.ToString(sqlReader["JQCode"]);
                                xlApplication.Cells[14, 31] = Convert.ToString(sqlReader["MRByRegistration"]);
                                xlApplication.Cells[14, 37] = Convert.ToString(sqlReader["MRCode"]);
                                xlApplication.Cells[17, 31] = Convert.ToString(sqlReader["TypeOfWeapon"]);
                                xlApplication.Cells[17, 37] = Convert.ToString(sqlReader["Code"]);
                                xlApplication.Cells[20, 37] = Convert.ToString(sqlReader["RU"]);
                                xlApplication.Cells[23, 37] = Convert.ToString(sqlReader["CategoryOfWorkability"]);
                                xlApplication.Cells[25, 37] = Convert.ToString(sqlReader["ReserveCategory"]);
                                xlApplication.Cells[2, 49] = Convert.ToString(sqlReader["SpecialTalents"]);
                                xlApplication.Cells[8, 45] = Convert.ToString(sqlReader["OVKKKpoVUS"]);
                                xlApplication.Cells[10, 41] = Convert.ToString(sqlReader["OVKKJobQualification"]);
                                xlApplication.Cells[10, 47] = Convert.ToString(sqlReader["OVKKKJQCode"]);
                                xlApplication.Cells[14, 41] = Convert.ToString(sqlReader["OVKKKMR"]);
                                xlApplication.Cells[14, 47] = Convert.ToString(sqlReader["OVKKKMRCode"]);
                                xlApplication.Cells[17, 41] = Convert.ToString(sqlReader["OVKKKToW"]);
                                xlApplication.Cells[17, 47] = Convert.ToString(sqlReader["OVKKKToWCode"]);
                                xlApplication.Cells[32, 36] = Convert.ToString(sqlReader["Education"]);
                                xlApplication.Cells[34, 40] = Convert.ToString(sqlReader["MilitaryEducation"]);
                                xlApplication.Cells[34, 56] = Convert.ToString(sqlReader["PersonalNumber"]);
                                xlApplication.Cells[36, 43] = Convert.ToString(sqlReader["CivilSpecialty"]);
                                //Срочная служба
                                xlApplication.Cells[42, 31] = Convert.ToString(sqlReader["MilitaryUnit1"]);
                                xlApplication.Cells[43, 31] = Convert.ToString(sqlReader["MilitaryUnit2"]);
                                xlApplication.Cells[42, 34] = Convert.ToString(sqlReader["VUSCode1"]);
                                xlApplication.Cells[43, 34] = Convert.ToString(sqlReader["VUSCode2"]);
                                xlApplication.Cells[42, 37] = Convert.ToString(sqlReader["MPCode1"]);
                                xlApplication.Cells[43, 37] = Convert.ToString(sqlReader["MPCode2"]);
                                xlApplication.Cells[42, 40] = Convert.ToString(sqlReader["MRCode1"]);
                                xlApplication.Cells[43, 40] = Convert.ToString(sqlReader["MRCode2"]);
                                xlApplication.Cells[42, 43] = Convert.ToString(sqlReader["MP1"]) + " " + Convert.ToString(sqlReader["OVKKJobQualification"]);
                                xlApplication.Cells[43, 43] = Convert.ToString(sqlReader["MP2"]);
                                xlApplication.Cells[42, 47] = Convert.ToString(sqlReader["ToWCode1"]);
                                xlApplication.Cells[43, 47] = Convert.ToString(sqlReader["ToWCode1"]);
                                xlApplication.Cells[42, 50] = Convert.ToString(sqlReader["ToW1"]);
                                xlApplication.Cells[43, 50] = Convert.ToString(sqlReader["ToW2"]);
                                xlApplication.Cells[42, 54] = Convert.ToString(sqlReader["Start1"]);
                                xlApplication.Cells[43, 54] = Convert.ToString(sqlReader["Start2"]);
                                xlApplication.Cells[42, 57] = Convert.ToString(sqlReader["End1"]);
                                xlApplication.Cells[43, 57] = Convert.ToString(sqlReader["End2"]);
                                //Альтернативная служба
                                xlApplication.Cells[47, 31] = Convert.ToString(sqlReader["AlternateMU1"]);
                                xlApplication.Cells[48, 31] = Convert.ToString(sqlReader["AlternateMU2"]);
                                xlApplication.Cells[47, 38] = Convert.ToString(sqlReader["AlternatePost1"]);
                                xlApplication.Cells[48, 38] = Convert.ToString(sqlReader["AlternatePost2"]);
                                xlApplication.Cells[47, 46] = Convert.ToString(sqlReader["AlternateStart1"]);
                                xlApplication.Cells[48, 46] = Convert.ToString(sqlReader["AlternateStart2"]);
                                xlApplication.Cells[47, 53] = Convert.ToString(sqlReader["AlternateEnd1"]);
                                xlApplication.Cells[48, 53] = Convert.ToString(sqlReader["AlternateEnd2"]);
                                //Военные сборы
                                xlApplication.Cells[52, 31] = Convert.ToString(sqlReader["MCYear"]);
                                xlApplication.Cells[52, 34] = Convert.ToString(sqlReader["MCAmountOfDays"]);
                                xlApplication.Cells[52, 37] = Convert.ToString(sqlReader["MCMU"]);
                                xlApplication.Cells[52, 40] = Convert.ToString(sqlReader["MCVUSCode"]);
                                xlApplication.Cells[52, 43] = Convert.ToString(sqlReader["MCMRCode"]);
                                xlApplication.Cells[52, 46] = Convert.ToString(sqlReader["MCMP"]);
                                xlApplication.Cells[52, 53] = Convert.ToString(sqlReader["MCToWCode"]);
                                xlApplication.Cells[52, 56] = Convert.ToString(sqlReader["MCToW"]);
                                //Страница 2
                                xlApplication.Cells[55, 1] = Convert.ToString(sqlReader["PlaceOfWork"]) + ", " + Convert.ToString(sqlReader["Position"]);
                                xlApplication.Cells[57, 9] = Convert.ToString(sqlReader["ResidentialAddress"]);
                                xlApplication.Cells[60, 13] = Convert.ToString(sqlReader["MaritalStatus"]);
                                xlApplication.Cells[62, 14] = Convert.ToString(sqlReader["DocumentWasDelivered"]);
                                xlApplication.Cells[64, 10] = Convert.ToString(sqlReader["InsteadWho"]);
                                xlApplication.Cells[66, 13] = Convert.ToString(sqlReader["MedicalExamination"]);
                                xlApplication.Cells[68, 17] = Convert.ToString(sqlReader["Bacteriocarrier"]);
                                xlApplication.Cells[70, 10] = Convert.ToString(sqlReader["Hostilites"]);
                                xlApplication.Cells[72, 13] = Convert.ToString(sqlReader["AccessNumber"]);
                                xlApplication.Cells[74, 7] = Convert.ToString(sqlReader["SpecialNotes"]);
                                //Антропометрические данные
                                xlApplication.Cells[83, 1] = Convert.ToString(sqlReader["Height"]);
                                xlApplication.Cells[83, 7] = Convert.ToString(sqlReader["Headdress"]);
                                xlApplication.Cells[83, 12] = Convert.ToString(sqlReader["GasMask"]);
                                xlApplication.Cells[83, 17] = Convert.ToString(sqlReader["Outfit"]);
                                xlApplication.Cells[83, 25] = Convert.ToString(sqlReader["Shoes"]);
                                //МП выдано
                                xlApplication.Cells[89, 2] = Convert.ToString(sqlReader["Day"]);
                                xlApplication.Cells[89, 5] = Convert.ToString(sqlReader["Month"]);
                                xlApplication.Cells[89, 10] = Convert.ToString(sqlReader["Year"]);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            xlApplication.Application.DisplayAlerts = false;
                            try
                            {
                                xlWorkbook.SaveAs();

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error: " + ex.Message);
                            }

                            xlApplication.Visible = true;
                            xlApplication.UserControl = true;

                            Close();
                        }
                        else
                        {
                            label3.Visible = true;
                            label3.Text = "Ошибка! Необходимо заполнить все поля!";
                        }
                    }
                }
            }
        }
        private void label2_Click(object sender, EventArgs e)
        {

        }
        //Обработчик выбора количества полей БД на вывод.
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            label5.Visible = false;
            label6.Visible = false;
            textBox1.Visible = false;
            textBox3.Visible = false;

            if (!(string.Equals(comboBox2.Text, "1")))
            {
                if ((string.Equals(comboBox2.Text, "2")))
                {
                    label5.Visible = true;
                    textBox1.Visible = true;
                }
                else
                {
                    label5.Visible = true;
                    textBox1.Visible = true;
                    label6.Visible = true;
                    textBox3.Visible = true;
                }
            }
        }
    }
}