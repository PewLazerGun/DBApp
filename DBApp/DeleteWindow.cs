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
using System.IO;
using static DBApp.DatabasePreferences;

namespace DBApp
{
    public partial class DeleteWindow : Form
    {
        private Form1 formActivity;
        public DeleteWindow(Form1 forma)
        {
            InitializeComponent();
            this.formActivity = forma;
        }
        public DeleteWindow()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (label3.Visible)
                label3.Visible = false;

            if (!string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrWhiteSpace(textBox2.Text))
            {
                SqlCommand command = new SqlCommand("DELETE FROM [Table] WHERE [Id]=@Id", getDb());

                command.Parameters.AddWithValue("Id", textBox2.Text);

                command.ExecuteNonQuery();
                formActivity.RefreshData();
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
