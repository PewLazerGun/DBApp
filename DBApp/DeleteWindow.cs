using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DBApp
{
    public partial class DeleteWindow : Form
    {
        SqlConnection sqlConnection;

        public DeleteWindow()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\users\user\documents\visual studio 2015\Projects\DBApp\DBApp\Database.mdf;Integrated Security=True";
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();

            if (label3.Visible)
                label3.Visible = false;

            if (!string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrWhiteSpace(textBox2.Text))
            {
                SqlCommand command = new SqlCommand("DELETE FROM [Table] WHERE [Id]=@Id", sqlConnection);

                command.Parameters.AddWithValue("Id", textBox2.Text);

                await command.ExecuteNonQueryAsync();

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
