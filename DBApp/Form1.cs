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
    public partial class Form1 : Form 
    {
        BindingSource bind;

        private DBApp.DatabaseDataSet.TableDataTable dTable = new DBApp.DatabaseDataSet.TableDataTable();

        public Form1()
        {
            InitializeComponent();
        }
        //Запуск приложения. Подключение и загрузка данных БД.
        private void Form1_Load(object sender, EventArgs e)
        {

            bind = new BindingSource();
            // TODO: данная строка кода позволяет загрузить данные в таблицу "databaseDataSet.Table". При необходимости она может быть перемещена или удалена.
            //this.tableTableAdapter.Fill(this.databaseDataSet.Table);
            SqlDataAdapter da = new SqlDataAdapter("SELECT * from [Table]", getDb());
            dTable.Clear();
            da.Fill(dTable);
            this.tableBindingSource.DataSource = dTable;
        }
        public void RefreshData()
        {
            SqlDataAdapter da = new SqlDataAdapter("SELECT * from [Table]",getDb());
            dTable.Clear();
            da.Fill(dTable);
            this.tableBindingSource.DataSource = dTable;
        }
        //Обновление
        private void Form1_Activated(object sender, EventArgs e)
        {
            this.RefreshData();
        }
        //Кнопка обновления
        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.RefreshData();
        }
        
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }
        //Вызов окна для печати 
        private void печатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintWindow1 f5 = new PrintWindow1();
            f5.Show();
        }
        //Вызов окна добавления
        private void добавитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddWindow f2 = new AddWindow(this);
            f2.Show();
        }
        //Вызов окна изменения
        private void изменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateWindow f3 = new UpdateWindow(this);
            f3.Show();
        }
        //Вызов окна удаления
        /*private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //DeleteWindow f4 = new DeleteWindow(this);
            //f4.Show();
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.Remove(row);
            }
        }*/

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //this.dataGridView1.Rows.Remove(this.dataGridView1.CurrentRow);
            //foreach (DataGridViewRow r in dataGridView1.SelectedRows)
            DeleteCab();
        }
        void DeleteCab()
        {

            if (MessageBox.Show("Удалить?\n\n", "Вопрос",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                SqlDataAdapter da = new SqlDataAdapter("DELETE FROM [Table] WHERE [Id]=" + dataGridView1.CurrentRow.Cells[0].Value, getDb());
                dTable.Clear();
                da.Fill(dTable);
                this.tableBindingSource.DataSource = dTable;
                this.RefreshData();
            }
        }
        //Кнопка сортировки
        private void button5_Click(object sender, EventArgs e)
        {
            if (label3.Visible)
                label3.Visible = false;

            if ((string.Equals(comboBox1.Text, @"ВУС")) &&
                (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrWhiteSpace(textBox1.Text)))
            {
                SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM [Table] WHERE [byVUS] LIKE @byVUS", getDb());
                da.SelectCommand.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@byVUS",
                    Value = "%" + textBox1.Text + "%",
                    SqlDbType = SqlDbType.NVarChar,
                    Size = 2000  // Assuming a 2000 char size of the field annotation (-1 for MAX)
                });
                dTable.Clear();
                da.Fill(dTable);
                this.tableBindingSource.DataSource = dTable;
            }

            if ((string.Equals(comboBox1.Text, @"ФИО")) &&
                (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrWhiteSpace(textBox1.Text)))
            {
                SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM [Table] WHERE [Name] LIKE @Name", getDb());
                da.SelectCommand.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@Name",
                    Value = "%" + textBox1.Text + "%",
                    SqlDbType = SqlDbType.NVarChar,
                    Size = 2000  // Assuming a 2000 char size of the field annotation (-1 for MAX)
                });
                dTable.Clear();
                da.Fill(dTable);
                this.tableBindingSource.DataSource = dTable;

            }

            if ((string.Equals(comboBox1.Text, @"Код должностной квалификации")) &&
                (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrWhiteSpace(textBox1.Text)))
            {
                SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM [Table] WHERE [JQCode] LIKE @JQCode", getDb());
                da.SelectCommand.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@JQCode",
                    Value = "%" + textBox1.Text + "%",
                    SqlDbType = SqlDbType.NVarChar,
                    Size = 2000  // Assuming a 2000 char size of the field annotation (-1 for MAX)
                });
                dTable.Clear();
                da.Fill(dTable);
                this.tableBindingSource.DataSource = dTable;
            }

            if ((string.Equals(comboBox1.Text, @"Год рождения")) &&
                (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrWhiteSpace(textBox1.Text)))
            {
                SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM [Table] WHERE [YearOfBirth] LIKE @YearOfBirth", getDb());
                da.SelectCommand.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@YearOfBirth",
                    Value = "%" + textBox1.Text + "%",
                    SqlDbType = SqlDbType.NVarChar,
                    Size = 2000  // Assuming a 2000 char size of the field annotation (-1 for MAX)
                });
                dTable.Clear();
                da.Fill(dTable);
                this.tableBindingSource.DataSource = dTable;
            }
        }

        //Кнопка справки
        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InfoWindow f5 = new InfoWindow();
            f5.Show();
        }
    }
}