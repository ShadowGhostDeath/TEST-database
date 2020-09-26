using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using exportWord = Microsoft.Office.Interop.Word;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb;";

        private OleDbConnection myConnection;

        public Form1()
        {
            InitializeComponent();

            myConnection = new OleDbConnection(connectString);
            
            myConnection.Open();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            myConnection.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string query = "SELECT Фамилия FROM Сотрудники WHERE № = 1";
            OleDbCommand command = new OleDbCommand(query, myConnection);
            textBox1.Text = command.ExecuteScalar().ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string query = "SELECT Фамилия, Имя, Отчество, Должность, Телефон FROM Сотрудники ORDER BY №";
            OleDbCommand command = new OleDbCommand(query, myConnection);
            OleDbDataReader reader = command.ExecuteReader();
            listBox1.Items.Clear();
            while (reader.Read())
            {
                listBox1.Items.Add(reader[0].ToString() + " " + reader[1].ToString() + " " + reader[2].ToString() + " " + reader[3].ToString() + " " + reader[4].ToString() + " ");
            }
            reader.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string query = "SELECT №, Должности FROM Должности ORDER BY №";
            OleDbCommand command = new OleDbCommand(query, myConnection);
            OleDbDataReader reader = command.ExecuteReader();
            listBox1.Items.Clear();
            while (reader.Read())
            {
                listBox1.Items.Add(reader[0].ToString() + " " + reader[1].ToString() + " ");
            }
            reader.Close();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string query = "SELECT №, Телефоны FROM Номера ORDER BY №";
            OleDbCommand command = new OleDbCommand(query, myConnection);
            OleDbDataReader reader = command.ExecuteReader();
            listBox1.Items.Clear();
            while (reader.Read())
            {
                listBox1.Items.Add(reader[0].ToString() + " " + reader[1].ToString() + " ");
            }
            reader.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string query = "INSERT INTO Сотрудники (Фамилия, Имя, Отчество) VALUES ('Трутов', 'Олег', 'Семенович')";
            OleDbCommand command = new OleDbCommand(query, myConnection);
            command.ExecuteNonQuery();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string query = "DELETE FROM Сотрудники WHERE № > 10";
            OleDbCommand command = new OleDbCommand(query, myConnection);
            command.ExecuteNonQuery();
        }
    }
}
