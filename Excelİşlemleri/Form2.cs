using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using exceleekle = Microsoft.Office.Interop.Excel;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace Excelİşlemleri
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            exceleekle.Application dosya = new exceleekle.Application();
            dosya.Visible = true;
            object missing = Type.Missing;
            Workbook kitap = dosya.Workbooks.Add(missing);
            Worksheet sayfa = (Worksheet)kitap.Sheets[1];

            int sutun = 1;
            int satir = 1;

            for(int i=0;i<dataGridView1.Columns.Count;i++) //sütun oluşturuyoruz.
            {
                Range myrange = (Range)sayfa.Cells[satir, sutun + i]; //alan oluşturuyor.
                myrange.Value2 = dataGridView1.Columns[i].HeaderText;
            }
            satir++;

            for(int j=0;j<dataGridView1.Rows.Count;j++) //satır oluşturuyoruz.
            {
                for(int k=0;k<dataGridView1.Columns.Count;k++) //satırlara denk gelebilecek sütunlar.
                {
                    Range myrange = (Range)sayfa.Cells[satir + j, sutun + k];
                    myrange.Value2 = dataGridView1[k,j].Value == null ? " " : dataGridView1[k,j].Value;
                    myrange.Select();

                }
            }



        }
    }
}
