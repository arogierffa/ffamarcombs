using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Google.Apis.Analytics.v3.Data;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.Style;

namespace FFAAnalyticsCenter.Controllers
{ 
    public class FFAAnalytics_Data
    {
        public string Metrics { get; set; }
        public string Dimensions { get; set; }
        public string Sort { get; set; }
        public string Filters { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string ReportTitle { get; set; }
        public bool bCompare { get; set; }
        public DateTime CompareStartDate { get; set; }
        public DateTime CompareEndDate { get; set; }
        public int MaxResults { get; set; }
    }

    public class AnalyticsController : Controller
    {
        GoogleAnalyticsController _api =
            new GoogleAnalyticsController(AppDomain.CurrentDomain.BaseDirectory + "/FFA Google Analytics-be18f6d8619e.p12",
                "ffa-google-analytics@ffa-ga-119121.iam.gserviceaccount.com");

        string _gaid = "";

        public AnalyticsController() { }

        //
        // GET: /Analytics/
        public FileStreamResult NewApiCall(string profileid, string metrics, string dimensions, string sort, string filters,
            string startdate, string enddate, string reporttitle,
            string maxresults, string bcompare, string comparestartdate = null, string compareenddate = null)
        {
            // prepare data for processing
            _gaid = profileid;
            DateTime dtStartTime = Convert.ToDateTime(startdate);
            DateTime dtEndTime = Convert.ToDateTime(enddate);
            DateTime dtCStartTime = new DateTime();
            DateTime dtCEndTime = new DateTime();
            Int32 MaxResults = Convert.ToInt32(maxresults);
            bool bCompareReport = (bcompare == "true");
            if (bCompareReport)
            {
                dtCStartTime = Convert.ToDateTime(comparestartdate);
                dtCEndTime = Convert.ToDateTime(compareenddate);
            }

            MemoryStream result = ProcessApiCall(metrics, dimensions, sort, filters, dtStartTime, dtEndTime, reporttitle,
                MaxResults, bCompareReport, dtCStartTime, dtCEndTime);

            var download = new FileStreamResult(result,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") // content type
            { FileDownloadName = reporttitle + "_" + startdate + "-" + enddate + (bCompareReport ? "_" + dtCStartTime + "-" + dtCEndTime : "") + ".xlsx" };

            return download;
        }

        public MemoryStream ProcessApiCall(string metrics, string dimensions, string sort, string filters,
            DateTime startdate, DateTime enddate, string reporttitle,
            Int32 maxresults, bool bcompare, DateTime comparestartdate, DateTime compareenddate)
        {
            FFAAnalytics_Data reportOverview = new FFAAnalytics_Data()
            {
                Metrics = metrics,
                Dimensions = dimensions,
                Sort = sort,
                Filters = filters,
                StartDate = startdate,
                EndDate = enddate,
                ReportTitle = reporttitle,
                bCompare = bcompare,
                CompareStartDate = comparestartdate,
                CompareEndDate = compareenddate,
                MaxResults = maxresults
            };

            GoogleAnalyticsController.AnalyticDataPoint data = _api.GetAnalyticsData(_gaid, dimensions, metrics, sort,
                filters, startdate, enddate, maxresults);
            GoogleAnalyticsController.AnalyticDataPoint comparedata = null;
            if (reportOverview.bCompare)
            {
                // display data as a comparison report
                comparedata = _api.GetAnalyticsData(_gaid, dimensions, metrics, sort, filters, comparestartdate,
                    compareenddate, maxresults);
            }

            // generate excel file for download
            MemoryStream excel = GenerateExcelSheet(reportOverview, data, comparedata);

            return excel;
        }

        public MemoryStream GenerateExcelSheet(FFAAnalytics_Data reportOverview,
            GoogleAnalyticsController.AnalyticDataPoint data,
            GoogleAnalyticsController.AnalyticDataPoint compareData = null)
        {
            //http://localhost:17375/FFAAnalytics/NewApiCall?metrics=ga:totalEvents,ga:uniqueEvents&dimensions=ga:eventLabel&sort=-ga:totalEvents&filters=ga:eventCategory==PDF-Files-ORG&startdate=2016-06-01&enddate=2016-06-17&reporttitle=test&maxresults=10&bcompare=false
            ExcelPackage xls = new ExcelPackage();
            var ws = xls.Workbook.Worksheets.Add(reportOverview.ReportTitle);

            var reportDesc = reportOverview.ReportTitle + " from " + reportOverview.StartDate.ToShortDateString() + "-" +
                             reportOverview.EndDate.ToShortDateString() +
                             (reportOverview.bCompare
                                 ? " comparing to " + reportOverview.CompareStartDate.ToShortDateString() + "-" +
                                   reportOverview.CompareEndDate.ToShortDateString()
                                 : "");

            ws.Cells[1, 1].Value = reportDesc;

            // column headers
            ws.Column(1).Width = 70;
            for (var x = 0; x < data.ColumnHeaders.Count; x++)
            {
                var colName = data.ColumnHeaders[x].Name;
                // translate technical name into comprehensive name
                if (colName == "ga:totalEvents") colName = "Total Events";
                if (colName == "ga:uniqueEvents") colName = "Unique Events";
                if (colName == "ga:eventLabel") colName = "Event Label";
                if (colName == "ga:eventCategory") colName = "Event Category";
                if (colName == "ga:sessions") colName = "Sessions";
                if (colName == "ga:percentNewSessions") colName = "New Sessions %";
                if (colName == "ga:newUsers") colName = "New Users";
                if (colName == "ga:bounceRate") colName = "Bounce Rate";
                if (colName == "ga:pageviewsPerSession") colName = "Page Views Per Session";
                if (colName == "ga:source") colName = "Source";
                if (colName == "ga:channelGrouping") colName = "Channel Grouping";
                if (colName == "ga:pageviews") colName = "Page Views";
                if (colName == "ga:uniquePageViews") colName = "Unique Page Views";
                if (colName == "ga:avgTimeOnPage") colName = "Average Time On Page";
                if (colName == "ga:entrances") colName = "Entrances";
                if (colName == "ga:bounceRate") colName = "Bounce Rate";
                if (colName == "ga:exitRate") colName = "Exit Rate";
                if (colName == "ga:pagePath") colName = "URL";
                if (colName == "ga:avgSessionDuration") colName = "Average Session Duration";
                if (colName == "ga:userType") colName = "User Type";
                if (colName == "ga:pageTitle") colName = "Page Title";

                // record column name
                ws.Cells[3, x + 1].Value = colName;
                ws.Cells[3, x + 1].Style.Font.Bold = true;
            }

            var currRow = 4;
            // populate grand total row
            ws.Cells[currRow, 1].Value = "TOTAL";
            // formatting
            ws.Cells[currRow, 1].Style.Font.Italic = true;
            ws.Cells[currRow, 1].Style.Font.Bold = true;
            ws.Cells[currRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            ws.Cells[currRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[currRow, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#B4C6E7"));

            if (reportOverview.bCompare)
            {
                currRow += 2;
                // populate data
                for (var x = 1; x < data.ColumnHeaders.Count; x++)
                {
                    // format total row
                    ws.Cells[currRow - 2, x + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[currRow - 2, x + 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#B4C6E7"));
                    // initial date range
                    ws.Cells[currRow, x + 1].Value = ConvertData(data.Rows.Sum(t => (data.ColumnHeaders[x].DataType == "INTEGER" ? Convert.ToInt32(t[x]) : Convert.ToDecimal(t[x]))),
                        data.ColumnHeaders[x].DataType);
                    FormatCell(ws, currRow, x + 1, data.ColumnHeaders[x].DataType);
                    ws.Cells[currRow, x + 1].Style.Font.Bold = true;
                    ws.Cells[currRow, x + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[currRow, x + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[currRow, x + 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));
                    // compare date range
                    ws.Cells[currRow + 1, x + 1].Value = ConvertData(compareData.Rows.Sum(t => (compareData.ColumnHeaders[x].DataType == "INTEGER" ? Convert.ToInt32(t[x]) : Convert.ToDecimal(t[x]))),
                        compareData.ColumnHeaders[x].DataType);
                    FormatCell(ws, currRow + 1, x + 1, compareData.ColumnHeaders[x].DataType);
                    ws.Cells[currRow + 1, x + 1].Style.Font.Bold = true;
                    ws.Cells[currRow + 1, x + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[currRow + 1, x + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[currRow + 1, x + 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));
                    // change %
                    var a = Convert.ToDecimal(ws.Cells[currRow, x + 1].Value);
                    var b = Convert.ToDecimal(ws.Cells[currRow + 1, x + 1].Value);
                    ws.Cells[currRow - 1, x + 1].Value = (a == 0 || b == 0 ? 0 : ConvertData(((a - b)/b), "DECIMAL")); // percentage difference
                    ws.Cells[currRow - 1, x + 1].Style.Numberformat.Format = "0.00%";
                    ws.Cells[currRow - 1, x + 1].Style.Font.Bold = true;
                    ws.Cells[currRow - 1, x + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[currRow - 1, x + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[currRow - 1, x + 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));

                    // conditional format
                    if (Convert.ToInt32(ws.Cells[currRow, x + 1].Value) >=
                        Convert.ToInt32(ws.Cells[currRow + 1, x + 1].Value))
                    {
                        ws.Cells[currRow - 1, x + 1].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#70AD47"));
                    }
                    else
                    {
                        ws.Cells[currRow - 1, x + 1].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FF0000"));
                    }
                }

                // formatting
                ws.Cells[currRow - 1, 1].Value = "Change %";
                ws.Cells[currRow - 1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells[currRow - 1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[currRow - 1, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));

                ws.Cells[currRow, 1].Value = reportOverview.StartDate.ToShortDateString() + " - " +
                                             reportOverview.EndDate.ToShortDateString();
                ws.Cells[currRow, 1].Style.Font.Italic = true;
                ws.Cells[currRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells[currRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[currRow, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));

                ws.Cells[currRow + 1, 1].Value = reportOverview.CompareStartDate.ToShortDateString() + " - " +
                                                 reportOverview.CompareStartDate.ToShortDateString();
                ws.Cells[currRow + 1, 1].Style.Font.Italic = true;
                ws.Cells[currRow + 1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells[currRow + 1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[currRow + 1, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"));
                currRow++;
            }
            else
            {
                // list grand total normal data
                for (var x = 1; x < data.ColumnHeaders.Count; x++)
                {
                    // unique case exceptions to convert data format
                    if (data.ColumnHeaders[x].DataType == "PERCENT" || 
                        data.ColumnHeaders[x].DataType == "FLOAT" || 
                        data.ColumnHeaders[x].DataType == "TIME")
                    {
                        ws.Cells[currRow, x + 1].Value = ConvertData(data.Rows.Sum(t => Convert.ToDecimal(t[x])), "decimal");
                        FormatCell(ws, currRow, x + 1, data.ColumnHeaders[x].DataType);
                    }
                    else
                    {
                        ws.Cells[currRow, x + 1].Value = ConvertData(data.Rows.Sum(t => Convert.ToInt32(t[x])),
                        data.ColumnHeaders[x].DataType);
                        if (data.ColumnHeaders[x].DataType == "INTEGER") FormatCell(ws, currRow, x + 1, data.ColumnHeaders[x].DataType);
                    }
                                       
                    // formatting
                    ws.Cells[currRow, x + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[currRow, x + 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#B4C6E7"));
                }
            }
            currRow++;

            // populate data rows
            for (var x = 0; x < data.Rows.Count; x++)
            {
                if (reportOverview.bCompare)
                {
                    // list compare report data
                    ws.Cells[currRow, 1].Value = data.Rows[x][0]; // item
                    ws.Cells[currRow + 1, 1].Value = "Change %" + "    "; // percentage difference
                    ws.Cells[currRow + 1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[currRow + 2, 1].Value = reportOverview.StartDate.ToShortDateString() + " - " +
                                                     reportOverview.EndDate.ToShortDateString() + "    ";
                        // initial date range
                    ws.Cells[currRow + 2, 1].Style.Font.Italic = true;
                    ws.Cells[currRow + 2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[currRow + 3, 1].Value = reportOverview.CompareStartDate.ToShortDateString() + " - " +
                                                     reportOverview.CompareEndDate.ToShortDateString() + "    ";
                        // compare date range
                    ws.Cells[currRow + 3, 1].Style.Font.Italic = true;
                    ws.Cells[currRow + 3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    for (var y = 1; y < data.ColumnHeaders.Count; y++)
                    {
                        var a = Convert.ToDecimal(data.Rows[x][y]);
                        decimal b = 0;
                        if (compareData.Rows.Count > x)
                        {
                            b = Convert.ToDecimal(compareData.Rows[x][y]);
                        }
                        ws.Cells[currRow + 1, y + 1].Value = (a == 0 || b == 0 ? 0 : ConvertData((a - b)/b, "DECIMAL")); // percentage difference
                        FormatCell(ws, currRow + 1, y + 1, "DECIMAL");
                        ws.Cells[currRow + 1, y + 1].Style.Numberformat.Format = "0.00%";
                        ws.Cells[currRow + 1, y + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[currRow + 2, y + 1].Value = ConvertData(a, data.ColumnHeaders[y].DataType); // initial date range
                        FormatCell(ws, currRow + 2, y + 1, data.ColumnHeaders[y].DataType);
                        //ws.Cells[currRow + 2, y + 1].Style.Numberformat.Format = "###,###";
                        ws.Cells[currRow + 2, y + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[currRow + 3, y + 1].Value = ConvertData(b,
                            compareData.ColumnHeaders[y].DataType); // compare date range
                        //ws.Cells[currRow + 3, y + 1].Style.Numberformat.Format = "###,###";
                        FormatCell(ws, currRow + 3, y + 1, data.ColumnHeaders[y].DataType);
                        ws.Cells[currRow + 3, y + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }

                    // row formatting
                    for (var y = 1; y <= data.ColumnHeaders.Count; y++)
                    {
                        ws.Cells[currRow, y].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[currRow, y].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9d9d9")); // light gray
                    }

                    currRow += 4;
                }
                else
                {
                    // list normal report data
                    for (var y = 0; y < data.ColumnHeaders.Count; y++)
                    {
                        ws.Cells[currRow, y + 1].Value = ConvertData(data.Rows[x][y], data.ColumnHeaders[y].DataType);
                        ws.Cells[currRow, y + 1].Style.HorizontalAlignment = (y == 0 ? ExcelHorizontalAlignment.Left : ExcelHorizontalAlignment.Right);
                        FormatCell(ws, currRow, y + 1, data.ColumnHeaders[y].DataType);

                        // row formatting
                        if (currRow%2 == 1)
                        {
                            ws.Cells[currRow, y + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[currRow, y + 1].Style.Fill.BackgroundColor.SetColor(
                                ColorTranslator.FromHtml("#d9d9d9")); // light gray
                        }
                    }
                    currRow++;
                }
            }

            // column type formatting
            for (var x = 2; x < 10; x++)
            {
                ws.Column(x).Width = 15;
            }

            // pane freezing
            ws.View.FreezePanes(4, 10);

            // create chart
            if (!reportOverview.bCompare) CreateChart(xls, reportOverview, data);

            // send download data
            MemoryStream memorystream = new MemoryStream();
            xls.SaveAs(memorystream);
            memorystream.Position = 0;

            return memorystream;
        }

        public void FormatCell(ExcelWorksheet ws, int row, int col, string dataType)
        {
            switch(dataType){
                case "PERCENT":
                case "FLOAT":
                case "TIME":
                case "DECIMAL":
                    ws.Cells[row, col].Style.Numberformat.Format = "#,##0.00";
                    break;
                case "INTEGER":
                default:
                    ws.Cells[row, col].Style.Numberformat.Format = "#,##0";
                    break;
            }
        }

        public object ConvertData(object val, string dataType)
        {
            var returnVal = val;
            switch (dataType)
            {
                case "INTEGER":
                    returnVal = Convert.ToInt32(returnVal);
                    break;
                case "PERCENT":
                case "FLOAT":
                case "TIME":
                case "DECIMAL":
                    returnVal = Convert.ToDecimal(returnVal);
                    break;
                case "STRING":
                    returnVal = returnVal.ToString();
                    break;
            }

            return returnVal;
        }

        public void CreateChart(ExcelPackage xls, FFAAnalytics_Data reportOverview,
            GoogleAnalyticsController.AnalyticDataPoint data,
            GoogleAnalyticsController.AnalyticDataPoint compareData = null)
        {
            reportOverview.Dimensions = "ga:year,ga:month,ga:day";
            reportOverview.Sort = "-ga:year,-ga:month,-ga:day";
            reportOverview.MaxResults = 1000;
            GoogleAnalyticsController.AnalyticDataPoint chartData = _api.GetAnalyticsData(_gaid,
                reportOverview.Dimensions, reportOverview.Metrics, reportOverview.Sort,
                reportOverview.Filters, reportOverview.StartDate, reportOverview.EndDate, reportOverview.MaxResults);

            var chartws = xls.Workbook.Worksheets.Add("ChartData");
            var firstRow = 16;

            // populate column headers
            var dateIndex = new int[3] {-1, -1, -1};
            for (var x = 0; x < chartData.ColumnHeaders.Count; x++)
            {
                var col = chartData.ColumnHeaders[x];
                chartws.Cells[firstRow, x + 1].Value = col.Name;

                // find index for date related column headers
                if (col.Name == "ga:year") dateIndex[0] = x;
                if (col.Name == "ga:month") dateIndex[1] = x;
                if (col.Name == "ga:day") dateIndex[2] = x;
            }

            // populate data row
            for (var x = 0; x < chartData.Rows.Count; x++)
            {
                for (var y = 0; y < chartData.ColumnHeaders.Count; y++)
                {
                    chartws.Cells[x + 1 + firstRow, y + 1].Value = ConvertData(chartData.Rows[x][y],
                        chartData.ColumnHeaders[y].DataType);
                }
            }

            // populate concatinated date data
            chartws.Cells[firstRow, chartData.ColumnHeaders.Count + 1].Value = "concantinateddate";
            chartws.Cells[firstRow, chartData.ColumnHeaders.Count + 2].Value = "datevalue";
            for (var x = 0; x < chartData.Rows.Count; x++)
            {
                chartws.Cells[x + 1 + firstRow, chartData.ColumnHeaders.Count + 1].Style.Numberformat.Format =
                    "mm/dd/yyyy";
                chartws.Cells[x + 1 + firstRow, chartData.ColumnHeaders.Count + 1].Formula = "=DATE(" +
                                                                                             Convert.ToInt32(
                                                                                                 chartData.Rows[x][
                                                                                                     dateIndex[0]]) +
                                                                                             ","
                                                                                             +
                                                                                             Convert.ToInt32(
                                                                                                 chartData.Rows[x][
                                                                                                     dateIndex[1]]) +
                                                                                             ","
                                                                                             +
                                                                                             Convert.ToInt32(
                                                                                                 chartData.Rows[x][
                                                                                                     dateIndex[2]]) +
                                                                                             ")";
                chartws.Cells[x + 1 + firstRow, chartData.ColumnHeaders.Count + 2].Formula = "=DATEVALUE(\"" +
                                                                                             Convert.ToInt32(
                                                                                                 chartData.Rows[x][
                                                                                                     dateIndex[1]]) +
                                                                                             "/"
                                                                                             +
                                                                                             Convert.ToInt32(
                                                                                                 chartData.Rows[x][
                                                                                                     dateIndex[2]]) +
                                                                                             "/"
                                                                                             +
                                                                                             Convert.ToInt32(
                                                                                                 chartData.Rows[x][
                                                                                                     dateIndex[0]]) +
                                                                                             "\")";
            }

            var chart = chartws.Drawings.AddChart("Timeline", eChartType.XYScatterLines);
            chart.SetPosition(0, 0, 0, 0);
            chart.SetSize(800, 300);
            chart.Title.Text = "Metrics Over Time";

            var Yrange = chartws.Cells[1 + firstRow, 1, chartData.Rows.Count + 1 + firstRow, 1];
            var Xrange =
                chartws.Cells[
                    1 + firstRow, chartData.ColumnHeaders.Count - 2, chartData.Rows.Count + 1 + firstRow,
                    chartData.ColumnHeaders.Count - 2];
            var series = chart.Series.Add(Yrange, Xrange);
            series.Header = reportOverview.StartDate.ToShortDateString() + "-" +
                            reportOverview.EndDate.ToShortDateString();
            chart.XAxis.MajorUnit = Math.Floor(chartData.Rows.Count*0.4) - 1;
            chart.XAxis.MinorUnit = Math.Floor(chartData.Rows.Count*0.06);

            // calculate bounds
            chartws.Cells[1 + firstRow, chartData.ColumnHeaders.Count + 2].Calculate();
            chartws.Cells[chartData.Rows.Count + 1 + firstRow, chartData.ColumnHeaders.Count + 2].Calculate();
            var minDate = new DateTime(Convert.ToInt32(chartData.Rows[chartData.Rows.Count - 1][dateIndex[0]]),
                Convert.ToInt32(chartData.Rows[chartData.Rows.Count - 1][dateIndex[1]]),
                Convert.ToInt32(chartData.Rows[chartData.Rows.Count - 1][dateIndex[2]]));
            var maxDate = new DateTime(Convert.ToInt32(chartData.Rows[0][dateIndex[0]]),
                Convert.ToInt32(chartData.Rows[0][dateIndex[1]]),
                Convert.ToInt32(chartData.Rows[0][dateIndex[2]]));
            chart.XAxis.MinValue = minDate.ToOADate();
            chart.XAxis.MaxValue = maxDate.ToOADate();

            // delete temporary columns
            chartws.DeleteColumn(chartData.ColumnHeaders.Count + 2); // date value
            chartws.DeleteColumn(dateIndex[2] + 1); // day
            chartws.DeleteColumn(dateIndex[1] + 1); // month
            chartws.DeleteColumn(dateIndex[0] + 1); // year
        }
    }
}