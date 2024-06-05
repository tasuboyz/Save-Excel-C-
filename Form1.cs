using Interfaccia_Database_sql_C_;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace db_movimenti
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CustomizeDataGrid1();
        }

        #region Personalizzazione
        private void CustomizeDataGrid1()
        {
            CustomizePanel();
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.DefaultCellStyle.BackColor = Color.FromArgb(50, 49, 69);
            dataGridView1.DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(50, 49, 69);
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(110, 182, 65);
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(50, 49, 69);
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font(dataGridView1.Font, FontStyle.Bold);
            dataGridView1.BackgroundColor = Color.FromArgb(44, 43, 60);
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
        }

        private void CustomizePanel()
        {
            panel1.BackColor = Color.FromArgb(50, 49, 69);
        }
        #endregion

        private void Tabella()
        {
            try 
            { 
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "File Excel|*.xls;*.xlsx;*.csv";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string selectedFilePath = openFileDialog.FileName;
                        DirectoryPaths paths = new DirectoryPaths();
                        string iniFilePath = paths.GetExcelIniPath();
                        IniFile iniFile = new IniFile(iniFilePath);
                        iniFile.Write("Settings", "ExcelPath", selectedFilePath);

                        if (File.Exists(iniFilePath))
                        {
                            var excelFilePath = iniFile.Read("Settings", "ExcelPath");
                            // Fai qualcosa con il percorso del file Excel, ad esempio, mostra un messaggio
                            MessageBox.Show($"Percorso del file Excel selezionato: {excelFilePath}");
                            UpdateExcelLabel(excelFilePath);
                            Table dt = new Table();
                            dataGridView1.DataSource = dt.ReadExcel();
                        }
                        else
                        {
                            // Gestisci il caso in cui il file ini non esiste
                        }
                    }
                    else
                    {
                        // Non aggiornare il file INI se l'utente annulla la selezione
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateExcelLabel(string excelFilePath)
        {
            string nomeFile = Path.GetFileName(excelFilePath);
            if (!string.IsNullOrEmpty(excelFilePath))
            {
                Excel_Label.Text = nomeFile;
            }
            else
            {
                Excel_Label.Text = "EXCEL";
            }
        }
    }
}
