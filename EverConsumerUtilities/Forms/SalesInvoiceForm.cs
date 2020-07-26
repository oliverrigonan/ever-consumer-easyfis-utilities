using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace EverConsumerUtilities.Forms
{
    public partial class SalesInvoiceForm : Form
    {
        public SalesInvoiceForm()
        {
            InitializeComponent();
            GetBranchData();
        }

        public void GetBranchData()
        {
            try
            {
                List<Entities.MstBranch> branches = new List<Entities.MstBranch>();

                var httpWebRequest = (HttpWebRequest)WebRequest.Create("http://everconsumer.easyfis.com/api/everConsumerIntegration/salesInvoice/dropdown/list/branch");
                httpWebRequest.Method = "GET";
                httpWebRequest.Accept = "application/json";

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
                    JavaScriptSerializer js = new JavaScriptSerializer();
                    List<Entities.MstBranch> deserializedBranches = (List<Entities.MstBranch>)js.Deserialize(result, typeof(List<Entities.MstBranch>));

                    if (deserializedBranches != null)
                    {
                        if (deserializedBranches.Any())
                        {
                            var branchData = from d in deserializedBranches
                                             select new Entities.MstBranch
                                             {
                                                 BranchCode = d.BranchCode,
                                                 Branch = d.Branch
                                             };

                            branches = branchData.ToList();
                        }
                    }
                }

                comboBoxBranch.DataSource = branches;
                comboBoxBranch.ValueMember = "BranchCode";
                comboBoxBranch.DisplayMember = "Branch";
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Ever Consumer Easyfis Utilities", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void GetSalesInvoiceData()
        {
            try
            {
                dataGridViewSalesInvoice.Rows.Clear();
                dataGridViewSalesInvoice.Refresh();

                dataGridViewSalesInvoice.Columns[0].DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#01A6F0");
                dataGridViewSalesInvoice.Columns[0].DefaultCellStyle.SelectionBackColor = ColorTranslator.FromHtml("#01A6F0");
                dataGridViewSalesInvoice.Columns[0].DefaultCellStyle.ForeColor = Color.White;

                String startDate = Convert.ToDateTime(dateTimePickerStartDate.Value).ToString("MM-dd-yyyy", CultureInfo.InvariantCulture);
                String endDate = Convert.ToDateTime(dateTimePickerEndDate.Value).ToString("MM-dd-yyyy", CultureInfo.InvariantCulture);
                String branchCode = comboBoxBranch.SelectedValue.ToString();
                String filter = textBoxFilter.Text;

                var httpWebRequest = (HttpWebRequest)WebRequest.Create("http://everconsumer.easyfis.com/api/everConsumerIntegration/salesInvoice/list/" + startDate + "/" + endDate + "/" + branchCode);
                httpWebRequest.Method = "GET";
                httpWebRequest.Accept = "application/json";

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
                    JavaScriptSerializer js = new JavaScriptSerializer();
                    List<Entities.TrnSalesInvoice> deserializedSales = (List<Entities.TrnSalesInvoice>)js.Deserialize(result, typeof(List<Entities.TrnSalesInvoice>));

                    if (deserializedSales != null)
                    {
                        if (deserializedSales.Any())
                        {
                            foreach (var salesData in deserializedSales)
                            {
                                if (salesData.SINumber.Contains(filter) ||
                                    salesData.Customer.Contains(filter) ||
                                    salesData.Remarks.Contains(filter) ||
                                    salesData.DocumentReference.Contains(filter))
                                {
                                    dataGridViewSalesInvoice.Rows.Add(
                                        "Print",
                                        salesData.Id,
                                        salesData.SINumber,
                                        salesData.SIDate,
                                        salesData.ManualSINumber,
                                        salesData.Customer,
                                        salesData.Remarks,
                                        salesData.DocumentReference,
                                        salesData.Amount.ToString("#,##.00"),
                                        salesData.IsLocked
                                    );
                                }
                            }
                        }
                    }

                    buttonGet.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ever Consumer Easyfis Utilities", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonGet_Click(object sender, EventArgs e)
        {
            buttonGet.Enabled = false;
            GetSalesInvoiceData();
        }

        private void SalesInvoiceForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            MainForm mainForm = new MainForm();
            mainForm.Show();
        }

        private void dataGridViewSalesInvoice_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex > -1 && dataGridViewSalesInvoice.CurrentCell.ColumnIndex == dataGridViewSalesInvoice.Columns["ColumnButtonPrint"].Index)
                {
                    Int32 SIId = Convert.ToInt32(dataGridViewSalesInvoice.Rows[dataGridViewSalesInvoice.CurrentCell.RowIndex].Cells[dataGridViewSalesInvoice.Columns["ColumnId"].Index].Value);
                    new Print.PrintTrnSalesInvoiceDetailPDF(SIId);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ever Consumer Easyfis Utilities", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
