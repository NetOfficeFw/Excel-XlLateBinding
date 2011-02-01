using System;
using System.Data;
using System.Collections.Generic;
using System.Text;

namespace Example10
{
    internal enum ProductSortMode
    { 
        ByCountOfSalesMonth =1,
        ByMaximumOutComeMonth = 2,

        ByCountOfSalesYear =3,
        ByMaximumOutComeYear = 4,
         
        ByCountOfSalesTotal =5,
        ByMaximumOutComeTotal =6
    }

    class SalesReportReportEntity
    {
        #region MyRegion
        
        public int CountOfSales { get; set; }
        public double ManufactorPriceSummary { get; set; }
        public double SalesPricesSummary { get; set; }
        public double OutcomeSummary { get; set; }        

        #endregion
    }

    class SalesReportProduct : IComparable
    {
        #region Fields

        internal SalesReport _parent;

        internal SalesReportReportEntity _Year = new SalesReportReportEntity();
        internal SalesReportReportEntity _PrevYear = new SalesReportReportEntity();
        internal SalesReportReportEntity _Total = new SalesReportReportEntity();
        internal SalesReportReportEntity _Month = new SalesReportReportEntity();
        internal List<SalesReportReportEntity> _PrevMonths = new List<SalesReportReportEntity>();
        internal SalesReportReportEntity _OpenOrders = new SalesReportReportEntity();
        
        #endregion

        #region Properties

        internal int ProductId { get; set; }
        internal string ProductName { get; set; }
        internal int StorageCount { get; set; }
        internal int SalesRankMonth { get; set; }
        internal int SalesRankYear { get; set; }
        internal int SalesRankTotal { get; set; }

        public SalesReportReportEntity Year
        {
            get
            {
                return _Year;
            }
        }

        public SalesReportReportEntity PrevYear
        {
            get
            {
                return _PrevYear;
            }
        }

        public SalesReportReportEntity Total
        {
            get
            {
                return _Total;
            }
        }

        public SalesReportReportEntity Month
        {
            get
            {
                return _Month;
            }
        }

        public List<SalesReportReportEntity> PrevMonths
        {
            get
            {
                return _PrevMonths;
            }
        }

        public SalesReportReportEntity OpenOrders
        {
            get
            {
                return _OpenOrders;
            }
        }

        #endregion

        #region Construction

        public SalesReportProduct(SalesReport parent)
        {
            _parent = parent;
        }
        
        #endregion
        
        #region IComparable
        
        public int CompareTo(Object o)
        {
            switch (_parent.SortMode)
            {
                case ProductSortMode.ByCountOfSalesMonth:
                    return this._Month.CountOfSales - ((SalesReportProduct)o)._Month.CountOfSales;
                case ProductSortMode.ByMaximumOutComeMonth:
                    return (int)Math.Round(((SalesReportProduct)o)._Month.OutcomeSummary - this._Month.OutcomeSummary);
                case ProductSortMode.ByCountOfSalesYear:
                    return this._Year.CountOfSales - ((SalesReportProduct)o)._Year.CountOfSales;
                case ProductSortMode.ByMaximumOutComeYear:
                    return (int)Math.Round(((SalesReportProduct)o)._Year.OutcomeSummary - this._Year.OutcomeSummary);
                case ProductSortMode.ByCountOfSalesTotal:
                    return this._Total.CountOfSales - ((SalesReportProduct)o)._Total.CountOfSales;
                default:
                    return (int)Math.Round(((SalesReportProduct)o)._Total.OutcomeSummary - this._Total.OutcomeSummary);
            }
        }
        
        #endregion

    }

    class SalesReport
    {
        #region Fields

        private int _yearToReport;
        private int _monthToReport;

        private DataSet _dataProducts = new DataSet("dataProducts");
        private DataSet _dataSales = new DataSet("dataSales");
        private DataSet _dataOrders = new DataSet("dataOrders");

        private List<SalesReportProduct> _listProducts = new List<SalesReportProduct>();

        #endregion

        #region Construction

        public SalesReport(int yearToReport, int monthToReport)
        {
            _yearToReport = yearToReport;
            _monthToReport = monthToReport;
            CreateSampleData();
        }

        #endregion
        
        #region Properties

        public SalesReportProduct[] Products
        {
            get
            {
                return _listProducts.ToArray();
            }
        }

        public ProductSortMode SortMode { get; set; }

        #endregion

        #region Public Methods

        public void Sort()
        {            
            _listProducts.Sort();

        }
        // Get data and calculate the report
        public void Proceed()
        {
            // perfom report for any product
            foreach (DataRow row in _dataProducts.Tables[0].Rows)
            {
                SalesReportProduct newProduct = new SalesReportProduct(this);
                int ProductId = Convert.ToInt32(row["iProductId"]);
                string ProductName = Convert.ToString(row["iName"]);
                int StorageCount = Convert.ToInt32(row["iStorageCount"]);
                newProduct.ProductId = ProductId;
                newProduct.ProductName = ProductName;
                newProduct.StorageCount = StorageCount;
                GetSalesAmounts(ProductId, ref newProduct._Year,  ref newProduct._PrevYear,  ref newProduct._Total, ref newProduct._Month, ref newProduct._PrevMonths);
                _listProducts.Add(newProduct);
                GetOrders(ProductId, ref newProduct._OpenOrders); 
            }

            SortMode = ProductSortMode.ByCountOfSalesMonth; 
            _listProducts.Sort();
            for (int i = 0; i < _listProducts.Count; i++)
            {
                _listProducts[i].SalesRankMonth = _listProducts.Count - i;
            }

            SortMode = ProductSortMode.ByCountOfSalesYear;
            _listProducts.Sort();
            for (int i = 0; i < _listProducts.Count; i++)
            {
                _listProducts[i].SalesRankYear = _listProducts.Count - i;
            }

            SortMode = ProductSortMode.ByCountOfSalesTotal;
            _listProducts.Sort();
            for (int i = 0; i < _listProducts.Count; i++)
            {
                _listProducts[i].SalesRankTotal = _listProducts.Count - i;
            }

            SortMode = ProductSortMode.ByMaximumOutComeMonth;
            _listProducts.Sort();
        }

        #endregion

        #region Private Methods

        private double GetManufactorPrice(int prodId)
        {
            string filterExpression = "iProductId = " + prodId;
            DataRow[] rows = _dataProducts.Tables[0].Select(filterExpression);
            if (rows.Length == 0)
                throw (new IndexOutOfRangeException("product not found."));

            return Convert.ToDouble(rows[0]["iManufactorAmount"]);
        }

        private void GetOrders(int prodId, ref SalesReportReportEntity orderTotals)
        {
            double ManufactorPrice = GetManufactorPrice(prodId);
            string filterExpression = "iProductId = " + prodId;
            DataRow[] rows = _dataOrders.Tables[0].Select(filterExpression);
            foreach (var row in rows)
            {
                orderTotals.ManufactorPriceSummary += ManufactorPrice;
                orderTotals.SalesPricesSummary += Convert.ToDouble(row["iOrderAmount"]);
                orderTotals.CountOfSales++;
            }
            orderTotals.OutcomeSummary = orderTotals.SalesPricesSummary - orderTotals.ManufactorPriceSummary;
        }

        private void GetSalesAmounts(int prodId, ref SalesReportReportEntity yearAmounts,
            ref SalesReportReportEntity yearPrevAmounts, ref SalesReportReportEntity yearTotals,
            ref SalesReportReportEntity monthTotals, ref List<SalesReportReportEntity> prevMonthsTotals)
        {
            double ManufactorPrice = GetManufactorPrice(prodId);
            string filterExpression = "iProductId = " + prodId;
            DataRow[] rows = _dataSales.Tables[0].Select(filterExpression);

            for (int i = 1; i <= _monthToReport - 1; i++)
            {
                prevMonthsTotals.Add(new SalesReportReportEntity());
            }

            foreach (var row in rows)
            {
                DateTime dtSalesDate = Convert.ToDateTime(row["iSalesDate"]);
                if (dtSalesDate.Year == _yearToReport)
                {
                    yearAmounts.ManufactorPriceSummary += ManufactorPrice;
                    yearAmounts.SalesPricesSummary += Convert.ToDouble(row["iSalesAmount"]);
                    yearAmounts.CountOfSales++;
                }

                if ((dtSalesDate.Year == _yearToReport) && (dtSalesDate.Month == _monthToReport))
                {
                    monthTotals.ManufactorPriceSummary += ManufactorPrice;
                    monthTotals.SalesPricesSummary += Convert.ToDouble(row["iSalesAmount"]);
                    monthTotals.CountOfSales++;
                } 

                if ((dtSalesDate.Year == _yearToReport) && (dtSalesDate.Month < _monthToReport))
                {
                    prevMonthsTotals[dtSalesDate.Month - 1].ManufactorPriceSummary += ManufactorPrice;
                    prevMonthsTotals[dtSalesDate.Month - 1].SalesPricesSummary += Convert.ToDouble(row["iSalesAmount"]);
                    prevMonthsTotals[dtSalesDate.Month - 1].CountOfSales++;
                } 
                if ((dtSalesDate.Year == _yearToReport) && (dtSalesDate.Month > _monthToReport))
                {

                }

                if (dtSalesDate.Year == _yearToReport-1)
                {
                    yearPrevAmounts.ManufactorPriceSummary += ManufactorPrice;
                    yearPrevAmounts.SalesPricesSummary += Convert.ToDouble(row["iSalesAmount"]);
                    yearPrevAmounts.CountOfSales++;
                }

                yearTotals.ManufactorPriceSummary += ManufactorPrice;
                yearTotals.SalesPricesSummary += Convert.ToDouble(row["iSalesAmount"]);
                yearTotals.CountOfSales++;
            }

            yearAmounts.OutcomeSummary = yearAmounts.SalesPricesSummary - yearAmounts.ManufactorPriceSummary;
            yearPrevAmounts.OutcomeSummary = yearPrevAmounts.SalesPricesSummary - yearPrevAmounts.ManufactorPriceSummary;
            monthTotals.OutcomeSummary = monthTotals.SalesPricesSummary - monthTotals.ManufactorPriceSummary;
            foreach (SalesReportReportEntity itemMonth in prevMonthsTotals)
            {
                itemMonth.OutcomeSummary = itemMonth.SalesPricesSummary - itemMonth.ManufactorPriceSummary;
            }
            yearTotals.OutcomeSummary = yearTotals.SalesPricesSummary - yearTotals.ManufactorPriceSummary;
        }
        
        private void CreateSampleData()
        {
           
            System.IO.Stream schemaStream = GetRessourceStream(this.GetType(), "Example10.XmlData.products.xsd");
            System.IO.Stream dataStream = GetRessourceStream(this.GetType(), "Example10.XmlData.products.xml");

            _dataProducts.ReadXmlSchema(schemaStream);
            _dataProducts.ReadXml(dataStream);

            schemaStream.Close();
            dataStream.Close();

            // the xml files contains sales data from 01/2008 - 06/2010
            schemaStream = GetRessourceStream(this.GetType(), "Example10.XmlData.sales.xsd");
            dataStream = GetRessourceStream(this.GetType(), "Example10.XmlData.sales.xml");

            _dataSales.ReadXmlSchema(schemaStream);
            _dataSales.ReadXml(dataStream);

            schemaStream.Close();
            dataStream.Close();

            schemaStream = GetRessourceStream(this.GetType(), "Example10.XmlData.orders.xsd");
            dataStream = GetRessourceStream(this.GetType(), "Example10.XmlData.orders.xml");

            _dataOrders.ReadXmlSchema(schemaStream);
            _dataOrders.ReadXml(dataStream);
            
            schemaStream.Close();
            dataStream.Close();
        }

        private System.IO.Stream GetRessourceStream(Type type, string fileName)
        {
            try
            {
                System.IO.Stream returnStream = type.Module.Assembly.GetManifestResourceStream(fileName);
                return returnStream;
            }
            catch
            {
                throw (new System.IO.IOException("Error accessing resource File."));
            }
        }

        #endregion
    }
}
