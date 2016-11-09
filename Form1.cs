using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace prakse1._2
{
    public partial class LoginForm : Form
    {
        public LoginForm()
        {
            InitializeComponent();
            textBox2.PasswordChar ='*';

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string connection = prakse1._2.Properties.Settings.Default.PrakseDBConnectionString;
            SqlConnection con = new SqlConnection(connection);
            con.Open();
            SqlCommand query = new SqlCommand("SELECT * FROM [Membership] WHERE username ='" + textBox1.Text + "'AND password='"
                 + textBox2.Text + "'AND isApproved =1", con);
            SqlDataReader reader = query.ExecuteReader();
            if (reader.HasRows)
            {
                MessageBox.Show("SUCCESS");
                con.Close();
                this.Hide();
                Form FormName = new MainForm();
                FormName.Show();
            }
            else
            {
                MessageBox.Show("Nepareizs lietotājvārds vai parole, mēģiniet vēlreiz!", "ERROR");
            }

            this.Hide();
            Form FormName = new MainForm();
            FormName.Show();
        }
    }
}
