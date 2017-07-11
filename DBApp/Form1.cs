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

namespace DBApp
{
    public partial class Form1 : Form
    {
        SqlConnection sqlConnection;

        public Form1()
        {
            InitializeComponent();
        }
        //Запуск приложения. Подключение и загрузка данных БД.
        private async void Form1_Load(object sender, EventArgs e)
        {
            string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\users\user\documents\visual studio 2015\Projects\DBApp\DBApp\Database.mdf;Integrated Security=True";
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();

            SqlDataReader sqlReader = null;
            SqlCommand command = new SqlCommand("SELECT * FROM [Table]", sqlConnection);

            try
            {
                sqlReader = await command.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    listBox1.Items.Add(Convert.ToString(sqlReader["Id"]) + "      " + Convert.ToString(sqlReader["Name"]) + "     " + Convert.ToString(sqlReader["MilitaryRank"]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }
        }
        //Вызов окна добавления строки в бд
        private void button1_Click(object sender, EventArgs e)
        {
            AddWindow f2 = new AddWindow();
            f2.Show();
        }
        //Кнопка обновления
        private async void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();

            SqlDataReader sqlReader = null;
            SqlCommand command = new SqlCommand("SELECT * FROM [Table]", sqlConnection);

            try
            {
                sqlReader = await command.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    listBox1.Items.Add(Convert.ToString(sqlReader["Id"]) + "      " + Convert.ToString(sqlReader["Name"]) + "     " + Convert.ToString(sqlReader["MilitaryRank"]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }
        }
        //Вызов окна изменения строки бд.
        private void button2_Click(object sender, EventArgs e)
        {
            UpdateWindow f3 = new UpdateWindow();
            f3.Show();
        }
        //Вызов окна удаления строки из бд.
        private void button3_Click(object sender, EventArgs e)
        {
            DeleteWindow f4 = new DeleteWindow();
            f4.Show();
        }
        //Обработчики выхода из приложения. Завершение сеанса с бд.
        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }
        
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }
        //Кнопка сортировки
        private async void button5_Click_1(object sender, EventArgs e)
        {
            if (label3.Visible)
                label3.Visible = false;


            SqlDataReader sqlReader = null;

            if ((string.Equals(comboBox1.Text, @"ФИО")) &&
                (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrWhiteSpace(textBox1.Text)))
            {

                listBox1.Items.Clear();

                SqlCommand command = new SqlCommand("SELECT * FROM [Table] WHERE [Name] LIKE @Name", sqlConnection);
                command.Parameters.AddWithValue("Name", "%" + textBox1.Text + "%");
                try
                {
                    sqlReader = await command.ExecuteReaderAsync();

                    while (await sqlReader.ReadAsync())
                    {
                        listBox1.Items.Add(Convert.ToString(sqlReader["Id"]) + "      " + Convert.ToString(sqlReader["Name"]) + "     " + Convert.ToString(sqlReader["MilitaryRank"]));
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    if (sqlReader != null)
                        sqlReader.Close();
                }
            }

            if ((string.Equals(comboBox1.Text, @"Воинское звание")) &&
                (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrWhiteSpace(textBox1.Text)))
            {

                listBox1.Items.Clear();

                SqlCommand command = new SqlCommand("SELECT * FROM [Table] WHERE [MilitaryRank] LIKE @MilitaryRank", sqlConnection);
                command.Parameters.AddWithValue("MilitaryRank", "%" + textBox1.Text + "%");
                try
                {
                    sqlReader = await command.ExecuteReaderAsync();

                    while (await sqlReader.ReadAsync())
                    {
                        listBox1.Items.Add(Convert.ToString(sqlReader["Id"]) + "      " + Convert.ToString(sqlReader["Name"]) + "     " + Convert.ToString(sqlReader["MilitaryRank"]));
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    if (sqlReader != null)
                        sqlReader.Close();
                }
            }
            /*if ((!string.Equals(comboBox1.Text, @"ФИО")) ||
                (!string.Equals(comboBox1.Text, @"Воинское звание")) ||
                (string.IsNullOrEmpty(textBox1.Text)) || (string.IsNullOrWhiteSpace(textBox1.Text)))
            {
                label3.Visible = true;
                label3.Text = "Ошибка! Проверьте правильность заполнения полей!";
            }*/
        }
    }
}
/*using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MDB
{
    public partial class MDB : Form
    {
        SqlConnection sqlConnection;
        public MDB()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private async void MDB_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "databaseDataSet.Table". При необходимости она может быть перемещена или удалена.
            this.tableTableAdapter.Fill(this.databaseDataSet.Table);
            string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\users\user\documents\visual studio 2015\Projects\MDB\MDB\Database.mdf;Integrated Security=True";
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();

            SqlDataReader sqlReader = null;
            SqlCommand command = new SqlCommand("SELECT * FROM [Table]", sqlConnection);

            try
            {
                sqlReader = await command.ExecuteReaderAsync();
                
                while(await sqlReader.ReadAsync())
                {
                    listBox1.Items.Add(Convert.ToString(sqlReader["Id"]) + "      " + Convert.ToString(sqlReader["Name"]) + "     " + Convert.ToString(sqlReader["MilitaryRank"]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }

        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }

        private void MDB_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            if (label7.Visible)
                label7.Visible = false;

            if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrWhiteSpace(textBox1.Text) &&
                !string.IsNullOrEmpty(comboBox1.Text) && !string.IsNullOrWhiteSpace(comboBox1.Text))
            {
                SqlCommand command = new SqlCommand("INSERT INTO [Table] (Name, MilitaryRank)VALUES(@Name, @MilitaryRank)", sqlConnection);
                command.Parameters.AddWithValue("Name", textBox1.Text);
                command.Parameters.AddWithValue("MilitaryRank", comboBox1.Text);

                await command.ExecuteNonQueryAsync();

            }
            else
            {
                label7.Visible = true;
                label7.Text = "Ошибка! Необходимо заполнить все поля!";
            }
        }

        private async void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();

            SqlDataReader sqlReader = null;
            SqlCommand command = new SqlCommand("SELECT * FROM [Table]", sqlConnection);

            try
            {
                sqlReader = await command.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    listBox1.Items.Add(Convert.ToString(sqlReader["Id"]) + "      " + Convert.ToString(sqlReader["Name"]) + "     " + Convert.ToString(sqlReader["MilitaryRank"]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            if (label8.Visible)
                label8.Visible = false;

            if (!string.IsNullOrEmpty(textBox5.Text) && !string.IsNullOrWhiteSpace(textBox5.Text) &&
                !string.IsNullOrEmpty(textBox4.Text) && !string.IsNullOrWhiteSpace(textBox4.Text) &&
                !string.IsNullOrEmpty(comboBox2.Text) && !string.IsNullOrWhiteSpace(comboBox2.Text))
            {
                SqlCommand command = new SqlCommand("UPDATE [Table] SET [Name]=@Name, [MilitaryRank]=@MilitaryRank WHERE [Id]=@Id", sqlConnection);

                command.Parameters.AddWithValue("Name", textBox4.Text);
                command.Parameters.AddWithValue("Id", textBox5.Text);
                command.Parameters.AddWithValue("MilitaryRank", comboBox2.Text);

                await command.ExecuteNonQueryAsync();

            }
            else
            {
                label8.Visible = true;
                label8.Text = "Ошибка! Необходимо заполнить все поля!";
            }

        }

        private async void button3_Click(object sender, EventArgs e)
        {
            if (label9.Visible)
                label9.Visible = false;

            if (!string.IsNullOrEmpty(textBox6.Text) && !string.IsNullOrWhiteSpace(textBox6.Text))
            {
                SqlCommand command = new SqlCommand("DELETE FROM [Table] WHERE [Id]=@Id", sqlConnection);

                command.Parameters.AddWithValue("Id", textBox6.Text);

                await command.ExecuteNonQueryAsync();
            }
            else
            {
                label9.Visible = true;
                label9.Text = "Ошибка! Поле 'Идентификатор' не заполнено!";
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
*/