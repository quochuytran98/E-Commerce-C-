﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using app = Microsoft.Office.Interop.Excel.Application;

namespace QL_CUAHANGADIDAS
{
    public partial class frmBestSeller : Form
    {
        public frmBestSeller()
        {
            InitializeComponent();
            load();
        }
        void load()
        {
            dtgvBestSeller.DataSource =DAO.BillDAO.Instance.GetBestSeller();
        }
       
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn xuất file EXCEL?", "Thông báo", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
            {
                DAO.Excel.Instance.export2Excel(dtgvBestSeller, @"D:\", "BestSeller_ADIDAS");
                MessageBox.Show("Đã xuất file BestSeller_ADIDAS vào Ổ D ");
            }
        }
    }
}
