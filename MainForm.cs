using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace prakse1._2
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'prakseDBDataSet10.Comission' table. You can move, or remove it, as needed.
            this.comissionTableAdapter5.Fill(this.prakseDBDataSet10.Comission);
            // TODO: This line of code loads data into the 'prakseDBDataSet9.Comission' table. You can move, or remove it, as needed.
            this.comissionTableAdapter4.Fill(this.prakseDBDataSet9.Comission);
            // TODO: This line of code loads data into the 'prakseDBDataSet8.Comission' table. You can move, or remove it, as needed.
            this.comissionTableAdapter3.Fill(this.prakseDBDataSet8.Comission);
            // TODO: This line of code loads data into the 'prakseDBDataSet7.Objects' table. You can move, or remove it, as needed.
            this.objectsTableAdapter1.Fill(this.prakseDBDataSet7.Objects);
            // TODO: This line of code loads data into the 'prakseDBDataSet4.Signifiers' table. You can move, or remove it, as needed.
            this.signifiersTableAdapter1.Fill(this.prakseDBDataSet4.Signifiers);            
            // TODO: This line of code loads data into the 'prakseDBDataSet1.Signifiers' table. You can move, or remove it, as needed.
            this.signifiersTableAdapter.Fill(this.prakseDBDataSet1.Signifiers);
            // TODO: This line of code loads data into the 'prakseDBDataSet.Owners' table. You can move, or remove it, as needed.
            this.ownersTableAdapter.Fill(this.prakseDBDataSet.Owners);

        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            comboParakst1.Visible = true;
            button4.Visible = true;
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            comboParakst2.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            comboKomis1.Visible = true;
            button2.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            comboKomis2.Visible = true;
            button3.Visible = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            comboKomis3.Visible = true;
        }

        private void button6_Click(object sender, EventArgs e) //SKATĪT ŠEIT
        {
            List<string> regNr = new List<string>();
            List<string> OccupationParakst1 = new List<string>();
            List<string> OccupationParakst2 = new List<string>();
            List<string> Komis1 = new List<string>();
            List<string> Komis2 = new List<string>();
            List<string> Komis3 = new List<string>();
            List<string> objektList = new List<string>();

            string connection = prakse1._2.Properties.Settings.Default.PrakseDBConnectionString;
            try
            {
                using (SqlConnection con = new SqlConnection(connection))
                {
                    con.Open();
                    //pieprasijumu definesana
                    SqlCommand queryOwner = new SqlCommand("SELECT * FROM [Owners] WHERE Name ='" + comboOwner.Text + "'", con);
                    SqlCommand queryParakst1 = new SqlCommand("SELECT * FROM [Signifiers] WHERE Name ='" + comboParakst1.Text + "'", con);
                    SqlCommand queryParakst2 = new SqlCommand("SELECT * FROM [Signifiers] WHERE Name ='" + comboParakst2.Text + "'", con);
                    SqlCommand queryKomis1 = new SqlCommand("SELECT * FROM [Comission] WHERE Name =N'" + comboKomis1.Text + "'", con);
                    SqlCommand queryKomis2 = new SqlCommand("SELECT * FROM [Comission] WHERE Name =N'" + comboKomis2.Text + "'", con);
                    SqlCommand queryKomis3 = new SqlCommand("SELECT * FROM [Comission] WHERE Name =N'" + comboKomis3.Text + "'", con);
                    SqlCommand queryObject = new SqlCommand("SELECT * FROM [Objects] WHERE Name ='" + comboName.Text + "'", con);
                    //1.piepr. izp.
                    SqlDataReader reader2 = queryKomis1.ExecuteReader(); 
                    while (reader2.Read())
                    {
                        MessageBox.Show(reader2.GetString(2));
                        Komis1.Add(reader2.GetString(2)); //pers kods
                        Komis1.Add(reader2.GetString(3)); //occupation
                        Komis1.Add(reader2.GetString(4)); //phone nr
                        
                    }
                    reader2.Close();

                    if (comboKomis2.Visible == true)
                    {
                        reader2 = queryKomis2.ExecuteReader();
                        while (reader2.Read())
                        {
                            MessageBox.Show(reader2.GetString(2));
                            Komis2.Add(reader2.GetString(2)); //pers kods
                            Komis2.Add(reader2.GetString(3)); //occupation
                            Komis2.Add(reader2.GetString(4)); //phone nr

                        }
                        reader2.Close();
                    }

                    if (comboKomis3.Visible == true)
                    {
                        reader2 = queryKomis3.ExecuteReader();
                        while (reader2.Read())
                        {
                            MessageBox.Show(reader2.GetString(2));
                            Komis3.Add(reader2.GetString(2)); //pers kods
                            Komis3.Add(reader2.GetString(3)); //occupation
                            Komis3.Add(reader2.GetString(4)); //phone nr

                        }
                    }
                    reader2.Close();
                    //SqlDataReader reader = queryOwner.ExecuteReader();
                    reader2 = queryOwner.ExecuteReader();
                    while (reader2.Read())
                    {
                        
                        MessageBox.Show(reader2.GetString(2));
                        regNr.Add(reader2.GetString(2));//regNr = regNr[0]
                    }
                    reader2.Close();
                    //2.piepr.
                    SqlDataReader reader = queryParakst1.ExecuteReader();
                    while (reader.Read())
                    {
                        MessageBox.Show("occup" + reader.GetString(2));
                        MessageBox.Show("per code" + reader.GetString(3));
                        OccupationParakst1.Add(reader.GetString(2));//occupation = OccupationParakst1[0]
                        OccupationParakst1.Add(reader.GetString(3));//persCode = OccupationParakst1[1]
                    }
                    reader.Close();

                    if (comboParakst2.Visible == true) //3
                    {
                        reader = queryParakst2.ExecuteReader();
                        while (reader.Read())
                        {
                            OccupationParakst2.Add(reader.GetString(2));
                            OccupationParakst2.Add(reader.GetString(3));
                        }
                        reader.Close();
                    }

                    reader = queryObject.ExecuteReader();//STRĀDĀ
                    while (reader.Read())
                    {
                        objektList.Add(reader.GetString(2)); //number = objekts[0]
                        objektList.Add(reader.GetString(3)); // cena = objekts[1]
                    }
                    reader.Close();
                    con.Close();
                }
                
            }
            catch (Exception)
            {
                MessageBox.Show("ERROR! CAN NOT CONNECT TO DATABASE");
            }
            
            { // aizpildisana excel faila 
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("EXCEL could not be started. Check that your office installation and project references are correct.");
                    return;
                }
                xlApp.Visible = true;
                Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Worksheet ws = (Worksheet)wb.Worksheets[1];
                if (ws == null)
                {
                    MessageBox.Show("Worksheet could not be created. Check that your office installation and project references are correct.");
                }
                //ws.cells[row,column] = ...;
                ws.Columns.AutoFit();
                ws.Cells[3, 3] = "Rēzeknes novada pāsvaldība";
                ws.Cells[4, 3] = "Rēzekne, Atbrīvošanas aleja 95a, LV-4601";
                ws.Cells[7, 4] = "Norastīšanas akts Nr. _______";
                ws.Cells[10, 3] = "APSTIPRINU";
                ws.Cells[11, 3] = "____________";
                ws.Cells[12, 3] = comboParakst1.Text + " (" + OccupationParakst1[0] + ") Personas kods: " + OccupationParakst1[1];
                if (comboParakst2.Visible == true)
                {
                    ws.Cells[14, 3] = "____________";
                    ws.Cells[15, 3] = comboParakst2.Text + " (" + OccupationParakst2[0] + ") Personas kods: " + OccupationParakst2[1]; ;
                }
                ws.Cells[17, 3] = "20__.gada__.___________";
                int n = 20;
                ws.Cells[n++, 2] = "Komisija šādā sastāvā: ";
                ws.Cells[n++, 2] = comboKomis1.Text
                     + " - " + Komis1[1] + " ar p.k.: " + Komis1[0] +
                     " tel. nr.: " + Komis1[2] + ",";//21
                if (comboKomis2.Visible == true)
                {
                    ws.Cells[n++, 2] = comboKomis2.Text
                     + " - " + Komis2[1] + " ar p.k.: " + Komis2[0] +
                     " tel. nr.: " + Komis2[2] + ",";
                }
                if (comboKomis3.Visible == true)
                {
                    ws.Cells[n++, 2] = comboKomis3.Text
                     + " - " + Komis3[1] + " ar p.k.: " + Komis3[0] +
                     " tel. nr.: " + Komis3[2] + ",";//23

                }
                ws.Cells[n++, 2] = "sastādīja aktu par inventāra: " + comboName.Text + //24
                    " ar ser. nr.: " + objektList[0] + " un cenu: " + objektList[1] + " Eur,";

                ws.Cells[n++, 3] = "kas ir " + comboOwner.Text + " (reģistrācijas numurs: " + regNr[0] + ") īpašums";
                ws.Cells[n+2, 2] = "Norakstīšanas iemesls: " + textBoxIemesls.Text;
                ws.Cells[n++, 2] = "Norakstīšanas datums: " + dateTimePicker1.Text;
                ws.Cells[34, 6] = "Komisijas locekļi:";
                ws.Cells[35, 7] = comboKomis1.Text; ws.Cells[36, 7] = "____________";
                ws.Cells[37, 7] = "(paraksts)";
                int ind = 38;
                if (comboKomis2.Visible == true)
                {
                    ws.Cells[38, 7] = comboKomis2.Text; ws.Cells[39, 7] = "____________";
                    ws.Cells[40, 7] = "(paraksts)";
                    ind = 41;
                }
                if (comboKomis3.Visible == true)
                {
                    ws.Cells[41, 7] = comboKomis3.Text; ws.Cells[42, 7] = "____________";
                    ws.Cells[43, 7] = "(paraksts)";
                    ind = 44;
                }
                ws.Cells[ind, 7] = textBoxParakstisanasVieta.Text;


                wb.SaveAs("Norakstisanas akts.xls");
            }
            
        }
    }
}
