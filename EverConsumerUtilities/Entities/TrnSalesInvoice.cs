using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EverConsumerUtilities.Entities
{
    class TrnSalesInvoice
    {
        public Int32 Id { get; set; }
        public Int32 BranchId { get; set; }
        public String SINumber { get; set; }
        public String SIDate { get; set; }
        public Int32 CustomerId { get; set; }
        public String Customer { get; set; }
        public String ContactNumber { get; set; }
        public String Address { get; set; }
        public String CustomerGroup { get; set; }
        public String TIN { get; set; }
        public Int32 TermId { get; set; }
        public String Term { get; set; }
        public String DocumentReference { get; set; }
        public String ManualSINumber { get; set; }
        public Int32? SOId { get; set; }
        public String ManualSONumber { get; set; }
        public String Remarks { get; set; }
        public Decimal Amount { get; set; }
        public Decimal PaidAmount { get; set; }
        public Decimal AdjustmentAmount { get; set; }
        public Decimal BalanceAmount { get; set; }
        public Int32 SoldById { get; set; }
        public String SoldBy { get; set; }
        public Int32 PreparedById { get; set; }
        public Int32 CheckedById { get; set; }
        public Int32 ApprovedById { get; set; }
        public String Status { get; set; }
        public Boolean IsCancelled { get; set; }
        public Boolean IsPrinted { get; set; }
        public Boolean IsLocked { get; set; }
        public Int32 CreatedById { get; set; }
        public String CreatedBy { get; set; }
        public String CreatedDateTime { get; set; }
        public Int32 UpdatedById { get; set; }
        public String UpdatedBy { get; set; }
        public String UpdatedDateTime { get; set; }
    }

    class TrnSalesInvoiceVATAnalysis
    {
        public Decimal VATSales { get; set; }
        public Decimal VATZeroRatedSales { get; set; }
        public Decimal VATExemptSales { get; set; }
        public Decimal LessDiscount { get; set; }
        public Decimal TotalSales { get; set; }
        public Decimal VAT { get; set; }
        public Decimal TotalAmountDue { get; set; }
    }
}
