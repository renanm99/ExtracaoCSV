using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Extracoes
{
    public partial class ExtracaoCSV : Form
    {
        public ExtracaoCSV()
        {
            InitializeComponent();
        }

        private int count { get; set; }
        private string div { get; set; }

        public void GerarCSV()
        {
            SqlConnection sqlConn = new SqlConnection("data source=srvlogp02; user id=prconsul;password=consulta; initial catalog=BDFICH; application name=SS");
            try
            {
                sqlConn.Open();

                SqlCommand sqlComm = new SqlCommand();
                sqlComm.CommandText = richTextBox1.Text;
                sqlComm.Connection = sqlConn;
                sqlComm.CommandTimeout = 0;

                using (SqlDataReader reader = sqlComm.ExecuteReader())
                {
                    int arquivo = 0;
                    if (reader.HasRows)
                    {
                        arquivo += 1;
                        string filename = textBox1.Text + " - ";
                        string filecsv = textBox3.Text + "\\" + filename + arquivo + ".csv";

                        while (reader.Read())
                        {
                            //if (count == numericUpDown1.Value)
                            //{
                            //    this.count = 0;
                            //    arquivo += 1;
                            //    filecsv = filename + arquivo + ".csv";

                            //    if (this.count == 0)
                            //    {
                            //        gerarCabecalho(filecsv, reader);
                            //    }

                            //    gerarCorpo(filecsv, reader);
                            //    continue;
                            //}
                            //else
                            //{
                            //    if (this.count == 0)
                            //    {
                            //        gerarCabecalho(filecsv, reader);
                            //    }

                            //    gerarCorpo(filecsv, reader);
                            //}

                            if (count == numericUpDown1.Value)
                            {
                                this.count = 0;
                                arquivo += 1;
                                filecsv = textBox3.Text + "\\" + filename + arquivo + ".csv";
                            }

                            if (this.count == 0)
                            {
                                gerarCabecalho(filecsv, reader);
                            }
                            gerarCorpo(filecsv, reader);
                        }
                    }
                    count = 0;
                    MessageBox.Show("Processo foi finalizado.\nForam gerados " + arquivo + " arquivos.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro");
            }
            finally
            {
                sqlConn.Close();
                sqlConn.Dispose();
            }
        }

        public bool verificaCampos()
        {
            if (richTextBox1.Text != "" || textBox1.Text != "" || numericUpDown1.Value > 0)
            {
                return true;
            }
            return false;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (verificaCampos())
            {
                this.count = 0;
                this.div = ";";
                if (checkBox1.Checked) { this.div = ","; }
                GerarCSV();
            }
            else
            {
                MessageBox.Show("Corrija os campos!");
            }
        }

        private void gerarCabecalho(string filecsv, SqlDataReader reader)
        {
            string header = "";

            for (int i = 0; i < reader.FieldCount; i++)
            {
                header += String.Format("{0}", reader.GetName(i));
                if (i == reader.FieldCount - 1)
                {
                    header += Environment.NewLine;
                    File.WriteAllText(filecsv, header);
                    continue;
                }
                header += this.div;
            }
        }

        private void gerarCorpo(string filecsv, SqlDataReader reader)
        {
            string body = "";

            for (int i = 0; i < reader.FieldCount; i++)
            {
                body += String.Format("{0}", reader.GetValue(i));
                if (i == reader.FieldCount - 1)
                {
                    body += Environment.NewLine;
                    File.AppendAllText(filecsv, body);
                    body = "";
                    this.count++;
                    continue;
                }
                body += this.div;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = folderBrowserDialog1.SelectedPath;
                button1.Enabled = true;
            }
        }
    }

}
