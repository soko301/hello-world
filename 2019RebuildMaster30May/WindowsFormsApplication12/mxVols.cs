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
    public partial class mxVols : Form
    {
          public mxVols(DataTable dataSource)
        {
            InitializeComponent();
            dataGridView1.DataSource = dataSource;
            format();

        }

        private void format()
        {

         for (int i = 0; i <= dataGridView1.Columns.Count - 1; i++)
              {
                  dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                  dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.White;
                  dataGridView1.Columns[i].Width = 60;

              }

         dataGridView1.Columns[0].DefaultCellStyle.BackColor = Color.Cyan;
         dataGridView1.Columns[6].DefaultCellStyle.BackColor = Color.Cyan;
         dataGridView1.Columns[12].DefaultCellStyle.BackColor = Color.Cyan;

         for (int i = 13; i <= 17; i++)
         {
             dataGridView1.Columns[i].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
             dataGridView1.Columns[i].DefaultCellStyle.ForeColor = Color.Black;
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
