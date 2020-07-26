using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EverConsumerUtilities
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void buttonSalesInvoice_Click(object sender, EventArgs e)
        {
            Forms.SalesInvoiceForm salesInvoiceForm = new Forms.SalesInvoiceForm();
            salesInvoiceForm.Show();

            Hide();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }
    }
}
