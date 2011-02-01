using System;
using System.Xml;
using System.Data;
using System.Windows.Forms;
using System.Reflection;
using System.Globalization;

using LateBindingApi.Excel;
using LateBindingApi.Excel.Office;
using LateBindingApi.Excel.Charts;
using LateBindingApi.Excel.Styles;
using LateBindingApi.Excel.Enums;

namespace Example10
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// month and year to report, no selection dialog in this example
        /// the sample data contains datasets from 01/2008 - 06/2010
        /// </summary>
        private readonly int _yearToReport = 2009;
        private readonly int _monthToReport = 5;

        #region Fields

        XlApplication _excelApplication;
        SalesReport _report;
        
        #endregion

        #region Construction

        public Form1()
        {
            InitializeComponent();       
        }
        
        #endregion

        #region Gui Trigger

        private void button1_Click(object sender, EventArgs e)
        {
         
            // start excel and turn off msg boxes
            _excelApplication = new XlApplication();
            _excelApplication.DisplayAlerts = false;
            _excelApplication.ScreenUpdating = false;
           
            // add a new workbook
            XlWorkbook workBook = _excelApplication.Workbooks.Add();

            // we use the first sheet as summary sheet and remove the 2 last sheets
            XlWorksheet summarySheet = workBook.Worksheets[1];
            workBook.Worksheets[3].Delete();
            workBook.Worksheets[2].Delete();

      
            // we get the data & perform the report
            _report = new SalesReport(_yearToReport, _monthToReport);
            _report.Proceed();

            // we create named styles for the range.Style property
            CreateStorageAndRankingStyle(workBook, "StorageAndRanking");
            CreateMonthStyle(workBook, "MonthInfos");
            CreateMonthStyle(workBook, "YearTotalInfos");

            // write product sheets
            XlWorksheet productSheet = null;
            foreach (SalesReportProduct itemProduct in _report.Products)
            {
                productSheet = workBook.Worksheets.Add();
                ProceedProductWorksheet(productSheet, itemProduct);
                productSheet.Move(null, workBook.Worksheets[workBook.Worksheets.Count]);
            }

            // write summary sheet
            ProceedSummaryWorksheet(_report, workBook,summarySheet, productSheet);
            summarySheet.Range("$A2").Select();

            // save the book 
            string fileExtension = XlConverter.GetDefaultExtension(_excelApplication);
            string workbookFile = string.Format("{0}\\Example10{1}", Environment.CurrentDirectory, fileExtension);
            workBook.SaveAs(workbookFile);

            // close excel and dispose reference
            _excelApplication.Quit();
            _excelApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Workbook saved.", workbookFile);
            fDialog.ShowDialog(this);
        }
        
        #endregion

        #region Write Summary to Worksheet Functions

        private void ProceedSummaryMatrix(SalesReport report, XlWorksheet summarySheet, XlStyle matrixStyle)
        {

            // table columns
            summarySheet.Range("B2").Value = "Count";
            summarySheet.Range("C2").Value = "Revenue";
            summarySheet.Range("D2").Value = "%";          
            summarySheet.Range("E2").Value = "Storage";

            string leftBottomCellAdress = XlConverter.ToCellAdress(1, 3 + report.Products.Length);
            string rightBottomCellAdress = XlConverter.ToCellAdress(5, 3 + report.Products.Length);

            summarySheet.Range("$A2:$" + rightBottomCellAdress).Style = matrixStyle;

            int rowIndex = 3;
            int columnIndex = 1;

            int i = 0;
            foreach (SalesReportProduct itemProduct in report.Products)
            {
                string prodName = itemProduct.ProductName;
                int prodId = itemProduct.ProductId;
                summarySheet.Cells(rowIndex, columnIndex).Value = prodName;

                string formula = string.Format("='{0}-{1}'!{2}", itemProduct.ProductName, itemProduct.ProductId, XlConverter.ToCellAdress(_monthToReport + 1, 13));
                summarySheet.Cells(rowIndex, columnIndex + 1).Value = formula;

                formula = string.Format("='{0}-{1}'!{2}", itemProduct.ProductName, itemProduct.ProductId, XlConverter.ToCellAdress(_monthToReport + 1, 12));
                summarySheet.Cells(rowIndex, columnIndex + 2).Value = formula;

                formula = string.Format("={0}*100/{1}", XlConverter.ToCellAdress(3, rowIndex), XlConverter.ToCellAdress(3, 3 + report.Products.Length));
                summarySheet.Cells(rowIndex, columnIndex + 3).Formula = formula;
              
                formula = string.Format("='{0}-{1}'!{2}", itemProduct.ProductName, itemProduct.ProductId, "B6");
                summarySheet.Cells(rowIndex, columnIndex + 4).Value = formula;
                int storeCount = Convert.ToInt16(summarySheet.Cells(rowIndex, columnIndex + 4).Value);
                
                if( (i%2) == 0)
                    summarySheet.Range("$A" + (i + 3).ToString() + ":$E" + (i + 3).ToString()).Interior.Color = XlConverter.ToDouble(System.Drawing.Color.Gainsboro);
                
                rowIndex++;
                i++;
            }

            string sumFormula = string.Format("=Sum({0}:{1})", "C3", "C" + (report.Products.Length + 3 - 1).ToString());
            summarySheet.Cells(rowIndex, columnIndex + 2).Value = sumFormula;
         
            summarySheet.Range("$C3:$C" + (report.Products.Length+3).ToString()).NumberFormat = "#,##0.00 €";
            summarySheet.Range("$D3:$D" + (report.Products.Length + 3).ToString()).NumberFormat = "0\"%\"";    
            summarySheet.Cells(3 + report.Products.Length,1).Value = "Total:";
            summarySheet.Range("D2").HorizontalAlignment = XlHAlign.xlHAlignCenter;
            summarySheet.Range("$B2:$E2").Font.Bold = true;
            summarySheet.Range(leftBottomCellAdress + ":" + rightBottomCellAdress).Font.Bold = true;
            summarySheet.Range(leftBottomCellAdress + ":" + rightBottomCellAdress).BorderAround(XlLineStyle.xlDouble, XlBorderWeight.xlMedium);  
        }

        private void ProceedSummaryWorksheet(SalesReport report, XlWorkbook workBook, XlWorksheet summarySheet, XlWorksheet afterSheet)
        {
            summarySheet.Name = "Summary";

            XlStyle matrixStyle = CreateSummaryStyle(workBook, "MatrixStyle");
            ProceedSummaryMatrix(report, summarySheet, matrixStyle);
            ProceedSummaryWorksheetCharts(summarySheet, report.Products.Length+1);
            ProceedSummaryPrintSettings(summarySheet);
            summarySheet.Columns.AutoFit();// proceed AutoFit before header
            ProceedSummaryWorksheetHeader(summarySheet);
           
            summarySheet.Select();
        }

        private void ProceedSummaryWorksheetCharts(XlWorksheet summarySheet, int countOfProducts)
        {

            string captionRangeAdress = "$A2:$" + XlConverter.ToCellAdress(1, 1 + countOfProducts);
            string fieldRangeAdress = "$C2:$" + XlConverter.ToCellAdress(3, 1 + countOfProducts);
           
            double chartTopPosition = summarySheet.Rows[countOfProducts+5].Top;
            double chartWidth = summarySheet.Columns[13].Left;

            XlChartObject chartSummary = summarySheet.ChartObjects.Add(1, chartTopPosition, chartWidth, 260);
            chartSummary.Chart.SetSourceData(summarySheet.Range(captionRangeAdress + ";" + fieldRangeAdress));
            
            fieldRangeAdress = "$D2:$" + XlConverter.ToCellAdress(4, 1 + countOfProducts);
            chartTopPosition = summarySheet.Rows[2].Top;
            double chartLeftPosition = summarySheet.Columns[8].Left;
            double chartHeight = summarySheet.Rows[countOfProducts + 3].Top - chartTopPosition;
            chartWidth = summarySheet.Columns[13].Left - summarySheet.Columns[8].Left;

            XlChartObject chartPercentOutcome = summarySheet.ChartObjects.Add(chartLeftPosition, chartTopPosition, chartWidth, chartHeight);
            chartPercentOutcome.Chart.ChartType = XlChartType.xlPie;
            chartPercentOutcome.Chart.SetSourceData(summarySheet.Range(captionRangeAdress + ";" + fieldRangeAdress));
            
        }

        private void ProceedSummaryWorksheetHeader(XlWorksheet summarySheet)
        {
            summarySheet.PageSetup.LeftHeader = "&D created";
            summarySheet.PageSetup.CenterHeader = "Vintage Digital Inc.";
            summarySheet.PageSetup.RightHeader = string.Format("Monthly Sales Report {1:00}/{0}", _yearToReport, _monthToReport);        
        }

        private void ProceedSummaryPrintSettings(XlWorksheet summarySheet)
        {
            summarySheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
            summarySheet.PageSetup.Zoom = false;
            summarySheet.PageSetup.FitToPagesTall = 1;
            summarySheet.PageSetup.FitToPagesWide = 1;
        }

        private XlStyle CreateSummaryStyle(XlWorkbook workBook, string styleName)
        {
            /*
             * borders in styles doesnt realy working, very simple using is possible with the index trick. thats all
            */
            XlStyle newStyle = workBook.Styles.Add(styleName);
            newStyle.Font.Size = 12;
            newStyle.Font.Name = "Courier New";
            newStyle.Borders[(XlBordersIndex)Constants.xlTop].LineStyle = XlLineStyle.xlContinuous;
            newStyle.Borders[(XlBordersIndex)Constants.xlTop].Color = 0;
            newStyle.Borders[(XlBordersIndex)Constants.xlTop].Weight = 2;

            newStyle.Borders[(XlBordersIndex)Constants.xlBottom].LineStyle = LateBindingApi.Excel.Enums.XlLineStyle.xlContinuous;
            newStyle.Borders[(XlBordersIndex)Constants.xlBottom].Color = 0;
            newStyle.Borders[(XlBordersIndex)Constants.xlBottom].Weight = 2;

            newStyle.Borders[(XlBordersIndex)Constants.xlLeft].LineStyle = LateBindingApi.Excel.Enums.XlLineStyle.xlDouble;
            newStyle.Borders[(XlBordersIndex)Constants.xlLeft].Color = 0;
            newStyle.Borders[(XlBordersIndex)Constants.xlLeft].Weight = 2;

            newStyle.Borders[(XlBordersIndex)Constants.xlRight].LineStyle = LateBindingApi.Excel.Enums.XlLineStyle.xlLineStyleNone;
            newStyle.Borders[(XlBordersIndex)Constants.xlRight].Color = 0;
            newStyle.Borders[(XlBordersIndex)Constants.xlRight].Weight = 2;

            return newStyle;

        }

        #endregion

        #region Write Product to Worksheet Functions

        private void ProceedProductWorksheet(XlWorksheet productSheet, SalesReportProduct itemProduct)
        {    
            string sheetName = string.Format("{0}-{1}", itemProduct.ProductName, itemProduct.ProductId.ToString());
            productSheet.Name = sheetName;
            
            // its not a random chain, write data first and create charts second
            ProceedProductStorageInfo(productSheet, itemProduct);
            ProceedProductMonthInfo(productSheet, itemProduct);
            ProceedProductYearTotalInfo(productSheet, itemProduct);
            ProceedProductMonthCharts(productSheet, itemProduct);
            ProceedProductPrintSettings(productSheet);
            productSheet.Columns.AutoFit(); // proceed AutoFit before header & ranking
            ProceedProductWorksheetHeader(productSheet, itemProduct);
            ProceedProductRankingInfo(productSheet, itemProduct);
             
            productSheet.Columns[14].ColumnWidth  = 2.14;
            productSheet.Columns[17].ColumnWidth = 5.14;
        }

        private void ProceedProductStorageInfo(XlWorksheet productSheet, SalesReportProduct itemProduct)
        {
            int rowIndex = 3;
            int columnIndex = 1;

            productSheet.Range("$A3:$B6").StyleAsString = "StorageAndRanking";

            productSheet.Cells(rowIndex, columnIndex).Value = "Storage Info";
            productSheet.Cells(rowIndex, columnIndex).Font.Bold = true;
            productSheet.Cells(rowIndex + 1, columnIndex).Value = "Storage Count";
            productSheet.Cells(rowIndex + 2, columnIndex).Value = "Sales in Progress";
            productSheet.Cells(rowIndex + 3, columnIndex).Value = "Recalc Storage Count ";

            productSheet.Cells(rowIndex + 1, columnIndex + 1).Value = itemProduct.StorageCount;
            productSheet.Cells(rowIndex + 2, columnIndex + 1).Value = itemProduct.OpenOrders.CountOfSales;
            productSheet.Cells(rowIndex + 3, columnIndex + 1).Value = itemProduct.StorageCount - itemProduct.OpenOrders.CountOfSales;

            SetProductStorageCountColor(itemProduct.StorageCount, productSheet.Cells(rowIndex + 1, columnIndex + 1));
            SetProductStorageCountColor(itemProduct.StorageCount - itemProduct.OpenOrders.CountOfSales, productSheet.Cells(rowIndex + 3, columnIndex + 1));
        }
       
        private void ProceedProductRankingInfo(XlWorksheet productSheet, SalesReportProduct itemProduct)
        {
            int rowIndex = 3;
            int columnIndex = 4;

            productSheet.Range("$D3:$F6").StyleAsString = "StorageAndRanking";

            productSheet.Cells(rowIndex, columnIndex).Value = "Count Ranking";
            productSheet.Cells(rowIndex, columnIndex).Font.Bold = true;
            productSheet.Cells(rowIndex + 1, columnIndex).Value = "Month";
            productSheet.Cells(rowIndex + 2, columnIndex).Value = "Year";
            productSheet.Cells(rowIndex + 3, columnIndex).Value = "Total";

            productSheet.Cells(rowIndex + 1, columnIndex + 2).Value = itemProduct.SalesRankMonth;
            productSheet.Cells(rowIndex + 2, columnIndex + 2).Value = itemProduct.SalesRankYear;
            productSheet.Cells(rowIndex + 3, columnIndex + 2).Value = itemProduct.SalesRankTotal;

            productSheet.Range("$D4:$E6").Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlLineStyleNone;
        }
        
        private void ProceedProductMonthInfo(XlWorksheet productSheet, SalesReportProduct itemProduct)
        {
            int rowIndex = 9;
            int iMonthCellIndex = 1;

            productSheet.Range("$A9:$M13").StyleAsString = "MonthInfos";

            productSheet.Cells(rowIndex + 1, iMonthCellIndex).Value = "ManufacturerPriceSummary";
            productSheet.Cells(rowIndex + 2, iMonthCellIndex).Value = "SalesPricesSummary";
            productSheet.Cells(rowIndex + 3, iMonthCellIndex).Value = "TotalRevenue";
            productSheet.Cells(rowIndex + 4, iMonthCellIndex).Value = "CountOfSales";

            iMonthCellIndex = 2; ;
            foreach (SalesReportReportEntity itemMonth in itemProduct.PrevMonths)
            {
                productSheet.Cells(rowIndex, iMonthCellIndex).Value = GetMonthName(iMonthCellIndex - 2);
                productSheet.Cells(rowIndex + 1, iMonthCellIndex).Value = itemMonth.ManufactorPriceSummary;
                productSheet.Cells(rowIndex + 2, iMonthCellIndex).Value = itemMonth.SalesPricesSummary;
                productSheet.Cells(rowIndex + 3, iMonthCellIndex).Value = itemMonth.OutcomeSummary;
                productSheet.Cells(rowIndex + 4, iMonthCellIndex).Value = itemMonth.CountOfSales;
                iMonthCellIndex++;
            }
            string cellAdress1 = XlConverter.ToCellAdress(itemProduct.PrevMonths.Count + 2, 10);
            string cellAdress2 = XlConverter.ToCellAdress(itemProduct.PrevMonths.Count + 2, 12);

            productSheet.Range("$B10:$" + cellAdress1).Interior.Color = XlConverter.ToDouble(System.Drawing.Color.Gainsboro);
            productSheet.Range("$B12:$" + cellAdress2).Interior.Color = XlConverter.ToDouble(System.Drawing.Color.Gainsboro);

            productSheet.Cells(rowIndex, iMonthCellIndex).Value = GetMonthName(_monthToReport - 1);
            productSheet.Cells(rowIndex + 1, iMonthCellIndex).Value = itemProduct.Month.ManufactorPriceSummary;
            productSheet.Cells(rowIndex + 2, iMonthCellIndex).Value = itemProduct.Month.SalesPricesSummary;
            productSheet.Cells(rowIndex + 3, iMonthCellIndex).Value = itemProduct.Month.OutcomeSummary;
            productSheet.Cells(rowIndex + 4, iMonthCellIndex).Value = itemProduct.Month.CountOfSales;

            for (int i = itemProduct.PrevMonths.Count + 2; i <= 12; i++)
            {
                iMonthCellIndex++;
                productSheet.Cells(rowIndex, iMonthCellIndex).Value = GetMonthName(i - 1);
            }

            productSheet.Range("$B9:$M9").NumberFormat = "";
            productSheet.Range("$B9:$M9").Font.Bold = true;

            productSheet.Range("$B13:$M13").NumberFormat = "";
            productSheet.Range("$B13:$M13").HorizontalAlignment = XlHAlign.xlHAlignCenter;

            if (itemProduct.PrevMonths.Count < 11)
            { 
                string topLeftMergeCellAdress = "$" +XlConverter.ToCellAdress(itemProduct.PrevMonths.Count + 3,10);
                productSheet.Range(topLeftMergeCellAdress + ":$M13").MergeCells = true;
            }
        }

        private void ProceedProductPrintSettings(XlWorksheet productSheet)
        {
            productSheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
            productSheet.PageSetup.Zoom = false;
            productSheet.PageSetup.FitToPagesTall = 1;
            productSheet.PageSetup.FitToPagesWide = 1;
            productSheet.PageSetup.PrintArea = "$A$1:$R$39";
        }

        private void ProceedProductMonthCharts(XlWorksheet productSheet, SalesReportProduct itemProduct)
        {
           
            double chartTop = productSheet.Rows[15].Top;
            double chartWidth = productSheet.Columns[14].Left;
            double chartHeight = productSheet.Rows[30].Top -productSheet.Rows[15].Top;

            XlChartObject chartMonths = productSheet.ChartObjects.Add(1, chartTop, chartWidth, chartHeight);
            chartMonths.Chart.SetSourceData(productSheet.Range("$A9:$M12"));

            chartTop = productSheet.Rows[31].Top;
            chartWidth = productSheet.Columns[14].Left;
            chartHeight = productSheet.Rows[40].Top - productSheet.Rows[33].Top;
            XlChartObject chartCountMonths = productSheet.ChartObjects.Add(1, chartTop, chartWidth, chartHeight);
            chartCountMonths.Chart.ChartType = XlChartType.xlLine;
            chartCountMonths.Chart.SetSourceData(productSheet.Range("$A13:$M13"));

            double chartLeft = productSheet.Columns[15].Left;
            chartTop = productSheet.Rows[15].Top;
            chartWidth = productSheet.Columns[19].Left - productSheet.Columns[15].Left;
            chartHeight = productSheet.Rows[30].Top - productSheet.Rows[15].Top;
            XlChartObject chartCountYears = productSheet.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight);
            chartCountYears.Chart.ChartType = XlChartType.xlCylinderColClustered;
            chartCountYears.Chart.SetSourceData(productSheet.Range("$O9:$P12"));

        }

        private void ProceedProductWorksheetHeader(XlWorksheet productSheet, SalesReportProduct itemProduct)
        {
            int rowIndex = 1;
            int columnIndex = 1;

            productSheet.PageSetup.LeftHeader = "&D created";
            productSheet.PageSetup.CenterHeader = "Vintage Digital Inc.";
            productSheet.PageSetup.RightHeader = string.Format("Monthly Sales Report {1:00}/{0}", _yearToReport, _monthToReport);

            productSheet.Cells(rowIndex, columnIndex).Value = itemProduct.ProductName;
            productSheet.Cells(rowIndex, columnIndex).Font.Bold = true;
            productSheet.Cells(rowIndex, columnIndex).Font.Underline = true;
            productSheet.Cells(rowIndex, columnIndex).Font.Size = 14;
            productSheet.Cells(rowIndex, columnIndex).Font.Name = "Verdana";

        }
 
        private void ProceedProductYearTotalInfo(XlWorksheet productSheet, SalesReportProduct itemProduct)
        {
            int ColumnIndex = 15;
            int RowIndex = 9;

            productSheet.Range("$O9:$R13").StyleAsString = "YearTotalInfos";
      
            productSheet.Cells(RowIndex, ColumnIndex).Value = "Year " + _yearToReport.ToString();
            productSheet.Cells(RowIndex + 1, ColumnIndex).Value = itemProduct.Year.ManufactorPriceSummary;
            productSheet.Cells(RowIndex + 2, ColumnIndex).Value = itemProduct.Year.SalesPricesSummary;
            productSheet.Cells(RowIndex + 3, ColumnIndex).Value = itemProduct.Year.OutcomeSummary;
            productSheet.Cells(RowIndex + 4, ColumnIndex).Value = itemProduct.Year.CountOfSales;

            productSheet.Cells(RowIndex, ColumnIndex + 1).Value = "Year " + (_yearToReport - 1).ToString();
            productSheet.Cells(RowIndex + 1, ColumnIndex + 1).Value = itemProduct.PrevYear.ManufactorPriceSummary;
            productSheet.Cells(RowIndex + 2, ColumnIndex + 1).Value = itemProduct.PrevYear.SalesPricesSummary;
            productSheet.Cells(RowIndex + 3, ColumnIndex + 1).Value = itemProduct.PrevYear.OutcomeSummary;
            productSheet.Cells(RowIndex + 4, ColumnIndex + 1).Value = itemProduct.PrevYear.CountOfSales;

            productSheet.Range("$O10:$P10").Interior.Color = XlConverter.ToDouble(System.Drawing.Color.Gainsboro);
            productSheet.Range("$O12:$P12").Interior.Color = XlConverter.ToDouble(System.Drawing.Color.Gainsboro);

            ColumnIndex = 18;
            RowIndex = 9;

            productSheet.Cells(RowIndex, ColumnIndex).Value = "Total";
            productSheet.Cells(RowIndex + 1, ColumnIndex).Value = itemProduct.Total.ManufactorPriceSummary;
            productSheet.Cells(RowIndex + 2, ColumnIndex).Value = itemProduct.Total.SalesPricesSummary;
            productSheet.Cells(RowIndex + 3, ColumnIndex).Value = itemProduct.Total.OutcomeSummary;
            productSheet.Cells(RowIndex + 4, ColumnIndex).Value = itemProduct.Total.CountOfSales;

            productSheet.Range("$R10").Interior.Color = XlConverter.ToDouble(System.Drawing.Color.Gainsboro);
            productSheet.Range("$R12").Interior.Color = XlConverter.ToDouble(System.Drawing.Color.Gainsboro);
            
            productSheet.Range("$O9:$R9").NumberFormat = "";
            productSheet.Range("$O9:$R9").Font.Bold = true;

            productSheet.Range("$O13:$R13").NumberFormat = "";
            productSheet.Range("$O13:$R13").HorizontalAlignment = XlHAlign.xlHAlignCenter;

            productSheet.Range("$Q9:$Q13").MergeCells = true;

        }
    
        private XlStyle CreateStorageAndRankingStyle(XlWorkbook workBook, string styleName)
        {
            /*
             * borders in styles doesnt realy working, very simple using is possible with the index trick. thats all
            */
            XlStyle newStyle = workBook.Styles.Add(styleName);
            newStyle.Font.Size = 12;
            newStyle.Font.Name = "Courier New";
            newStyle.Borders[(XlBordersIndex)Constants.xlTop].LineStyle = XlLineStyle.xlContinuous;
            newStyle.Borders[(XlBordersIndex)Constants.xlTop].Color = 0;
            newStyle.Borders[(XlBordersIndex)Constants.xlTop].Weight = 2;

            newStyle.Borders[(XlBordersIndex)Constants.xlBottom].LineStyle = LateBindingApi.Excel.Enums.XlLineStyle.xlContinuous;
            newStyle.Borders[(XlBordersIndex)Constants.xlBottom].Color = 0;
            newStyle.Borders[(XlBordersIndex)Constants.xlBottom].Weight = 2;

            newStyle.Borders[(XlBordersIndex)Constants.xlLeft].LineStyle = LateBindingApi.Excel.Enums.XlLineStyle.xlDouble;
            newStyle.Borders[(XlBordersIndex)Constants.xlLeft].Color = 0;
            newStyle.Borders[(XlBordersIndex)Constants.xlLeft].Weight = 2;

            newStyle.Borders[(XlBordersIndex)Constants.xlRight].LineStyle = LateBindingApi.Excel.Enums.XlLineStyle.xlLineStyleNone;
            newStyle.Borders[(XlBordersIndex)Constants.xlRight].Color = 0;
            newStyle.Borders[(XlBordersIndex)Constants.xlRight].Weight = 2;

            return newStyle;

        }

        private XlStyle CreateMonthStyle(XlWorkbook workBook, string styleName)
        {
            /*
             * borders in styles doesnt realy working, very simple using is possible with the index trick. thats all
            */

            XlStyle newStyle = workBook.Styles.Add(styleName);
            newStyle.Font.Size = 12;
            newStyle.Borders[(XlBordersIndex)Constants.xlTop].LineStyle = XlLineStyle.xlContinuous;
            newStyle.Borders[(XlBordersIndex)Constants.xlTop].Color = 0;
            newStyle.Borders[(XlBordersIndex)Constants.xlTop].Weight = 2;

            newStyle.Borders[(XlBordersIndex)Constants.xlBottom].LineStyle = LateBindingApi.Excel.Enums.XlLineStyle.xlContinuous;
            newStyle.Borders[(XlBordersIndex)Constants.xlBottom].Color = 0;
            newStyle.Borders[(XlBordersIndex)Constants.xlBottom].Weight = 2;

            newStyle.Borders[(XlBordersIndex)Constants.xlLeft].LineStyle = LateBindingApi.Excel.Enums.XlLineStyle.xlDouble;
            newStyle.Borders[(XlBordersIndex)Constants.xlLeft].Color = 0;
            newStyle.Borders[(XlBordersIndex)Constants.xlLeft].Weight = 2;

            newStyle.Borders[(XlBordersIndex)Constants.xlRight].LineStyle = LateBindingApi.Excel.Enums.XlLineStyle.xlLineStyleNone;
            newStyle.Borders[(XlBordersIndex)Constants.xlRight].Color = 0;
            newStyle.Borders[(XlBordersIndex)Constants.xlRight].Weight = 2;

            newStyle.NumberFormat = "#,##0.00 €";

            return newStyle;

        }

        private XlStyle CreateYearTotalStyle(XlWorkbook workBook, string styleName)
        {
            /*
             * borders in styles doesnt realy working, very simple using is possible with the index trick. thats all
            */

            XlStyle newStyle = workBook.Styles.Add(styleName);
            newStyle.Font.Size = 12;
            newStyle.Borders[(XlBordersIndex)Constants.xlTop].LineStyle = XlLineStyle.xlContinuous;
            newStyle.Borders[(XlBordersIndex)Constants.xlTop].Color = 0;
            newStyle.Borders[(XlBordersIndex)Constants.xlTop].Weight = 2;

            newStyle.Borders[(XlBordersIndex)Constants.xlBottom].LineStyle = LateBindingApi.Excel.Enums.XlLineStyle.xlContinuous;
            newStyle.Borders[(XlBordersIndex)Constants.xlBottom].Color = 0;
            newStyle.Borders[(XlBordersIndex)Constants.xlBottom].Weight = 2;

            newStyle.Borders[(XlBordersIndex)Constants.xlLeft].LineStyle = LateBindingApi.Excel.Enums.XlLineStyle.xlDouble;
            newStyle.Borders[(XlBordersIndex)Constants.xlLeft].Color = 0;
            newStyle.Borders[(XlBordersIndex)Constants.xlLeft].Weight = 2;

            newStyle.Borders[(XlBordersIndex)Constants.xlRight].LineStyle = LateBindingApi.Excel.Enums.XlLineStyle.xlLineStyleNone;
            newStyle.Borders[(XlBordersIndex)Constants.xlRight].Color = 0;
            newStyle.Borders[(XlBordersIndex)Constants.xlRight].Weight = 2;

            newStyle.NumberFormat = "#,##0.00 €";

            return newStyle;

        }

        private string GetMonthName(int index)
        {
            return LateBindingApi.Excel.XlLateBindingApiSettings.XlThreadCulture.DateTimeFormat.MonthNames[index].Substring(0, 3);
        }

        private void SetProductStorageCountColor(int storageCount, XlRange range)
        {
            if (storageCount <= 50)
                range.Interior.Color = XlConverter.ToDouble(System.Drawing.Color.Red);
            else if (storageCount <= 100)
                range.Interior.Color = XlConverter.ToDouble(System.Drawing.Color.Yellow);
            else
                range.Interior.Color = XlConverter.ToDouble(System.Drawing.Color.Green);
        }

        #endregion
    }
}
