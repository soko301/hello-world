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
    public partial class tradeSim : Form
    {
          public tradeSim(DataTable dataSource)
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
                  dataGridView1.Columns[i].Width = 75;

              }


              dataGridView1.Columns[0].DefaultCellStyle.BackColor = Color.DarkGray;
              dataGridView1.Columns[11].DefaultCellStyle.BackColor = Color.Cyan;



              for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
              {
                  dataGridView1.Rows[i].Height = 18;
              }







          }

          private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
          {
              if (e.KeyCode == Keys.Delete)
              {
                  this.Close();
              }
          }
    }
}
