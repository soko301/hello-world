using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class SkewPopUP : Form
    {
        public SkewPopUP(DataTable dataSource)
        {
            InitializeComponent();
            dataGridView1.DataSource = dataSource;

            formatSkewView();
        }

        private void formatSkewView()
        {

            for (int i = 0; i <= dataGridView1.Columns.Count - 1; i++)
            {
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.White;
                dataGridView1.Columns[i].Width = 55;

            }


            dataGridView1.Columns[0].DefaultCellStyle.BackColor = Color.DarkGray;
            dataGridView1.Columns[10].DefaultCellStyle.BackColor = Color.Cyan;
            


            for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView1.Rows[i].Height = 18;
            }







        }



    }
}
