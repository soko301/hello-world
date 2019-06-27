using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace pricer
{
    public partial class risk : Form
    {
        //public risk(DataTable dataSource, DataTable dataSource1)
        //{
        //    InitializeComponent();
        //    dataGridView1.DataSource = dataSource;
        //    dataGridView2.DataSource = dataSource1;
        //    formatSkewView();
        //}

        public risk(DataSet ds)
        {
            InitializeComponent();


            TabControl dynamicTabControl = new TabControl();
            dynamicTabControl.Name = "DynamicTabControl";
            dynamicTabControl.BackColor = Color.White;
            dynamicTabControl.ForeColor = Color.Black;
            dynamicTabControl.Font = new Font("Arial", 10);
            dynamicTabControl.Width = 600;
            dynamicTabControl.Height = 200;
            dynamicTabControl.Dock = DockStyle.Fill;
          

            Controls.Add(dynamicTabControl); 

            foreach (DataTable dt in ds.Tables)
            {
                string tname = dt.TableName;
                TabPage tp = new TabPage(tname);
                dynamicTabControl.TabPages.Add(tp);

                DataGridView dg = new DataGridView();
                dg.Dock = DockStyle.Fill;
                tp.Controls.Add(dg);

                dg.DataSource = dt;
            }
        }

      


        //private void formatSkewView()
        //{

        //    for (int i = 0; i <= dataGridView1.Columns.Count - 1; i++)
        //    {
        //        dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        //        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.White;
        //        dataGridView1.Columns[i].Width = 70;
        //        //dataGridView1.Columns[i].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
        //       // dataGridView1.Columns[i].DefaultCellStyle.ForeColor = Color.Black;

        //    }


        //    //for (int i = 0; i <= dataGridView2.Columns.Count - 1; i++)
        //    //{
        //    //    dataGridView2.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        //    //    dataGridView2.Columns[i].DefaultCellStyle.BackColor = Color.White;
        //    //    dataGridView2.Columns[i].Width = 70;
        //    //    //dataGridView1.Columns[i].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
        //    //    // dataGridView1.Columns[i].DefaultCellStyle.ForeColor = Color.Black;

        //    //}

        //}
    }
}
