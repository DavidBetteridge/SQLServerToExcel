using ClosedXML.Excel;
using Microsoft.Data.ConnectionUI;
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

namespace SqlToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            this.txtConnectionString.Text = "Data Source=.;Initial Catalog=master;Integrated Security=True";
            this.txtFilename.Text = @"c:\temp\extract.xlsx";
            this.txtSQL.Text = "SELECT * FROM sys.databases";
        }

        /// <summary>
        /// The user has clicked cancel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmdCancel_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void cmdCS_Click(object sender, EventArgs e)
        {
            var dataConnectionDialog = new DataConnectionDialog();
            DataSource.AddStandardDataSources(dataConnectionDialog);

            if (DataConnectionDialog.Show(dataConnectionDialog) == DialogResult.OK)
            {
                this.txtConnectionString.Text = dataConnectionDialog.ConnectionString;
            }

        }

        private async void cmdOK_Click(object sender, EventArgs e)
        {
            var cs = this.txtConnectionString.Text.Trim();
            if (string.IsNullOrWhiteSpace(cs))
            {
                MessageBox.Show("Please enter the connection string", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.txtConnectionString.Focus();
                return;
            }

            var sql = this.txtSQL.Text.Trim();
            if (string.IsNullOrWhiteSpace(sql))
            {
                MessageBox.Show("Please enter the sql query", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.txtSQL.Focus();
                return;
            }

            var filename = this.txtFilename.Text.Trim();
            if (string.IsNullOrWhiteSpace(filename))
            {
                MessageBox.Show("Please enter the name of the file you need to generate", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.txtFilename.Focus();
                return;
            }

            try
            {
                using (var cn = new SqlConnection(cs))
                {
                    await cn.OpenAsync();

                    using (var cmd = new SqlCommand())
                    {
                        cmd.Connection = cn;
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = sql;

                        using (var dr = await cmd.ExecuteReaderAsync())
                        {
                            var rowCount = 1;
                            var workbook = new XLWorkbook();
                            var worksheet = workbook.Worksheets.Add("Sheet1");

                            // Get the column names
                            for (int i = 0; i < dr.FieldCount; i++)
                            {
                                var name = dr.GetName(i);
                                worksheet.Cell(GetColumnLetter(i + 1) + rowCount.ToString()).Value = name;
                            }


                            while (await dr.ReadAsync())
                            {
                                rowCount++;

                                for (int i = 0; i < dr.FieldCount; i++)
                                {
                                    var value = dr.GetValue(i).ToString();
                                    worksheet.Cell(GetColumnLetter(i + 1) + rowCount.ToString()).Value = value;
                                }

                            }

                            workbook.SaveAs(filename);
                            MessageBox.Show("Complete");
                        }

                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private static string GetColumnLetter(int column)
        {
            if (column < 1) return String.Empty;
            return GetColumnLetter((column - 1) / 26) + (char)('A' + (column - 1) % 26);
        }

        private void cmdBrowse_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.FileName = this.txtFilename.Text;
            if (openFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                this.txtFilename.Text = openFileDialog1.FileName;
            }
        }
    }
}
