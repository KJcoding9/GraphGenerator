using System;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace GraphGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            Program p = new Program();
            PathData pathData = new PathData();

            Console.WriteLine("xlsxデータのパスを入力してください");

            pathData.ExcelPath = Console.ReadLine() ?? " ";

            using (Application excelApp = new Application())
            {
                var book = p.FileRead(pathData.ExcelPath, excelApp);

                if(book != null && book.Worksheets.Count >= 1)
                {
                    Worksheet? worksheet = book.Worksheets[1] as Worksheet;

                    if (worksheet != null)
                    {
                        p.GraphGenerate(worksheet);
                    }
                    book.Save();
                }
                excelApp.Quit();
            }
        }
        Workbook? FileRead(string path, Application excelApp)
        {
            Workbook? book = null;
            try
            {
                book = excelApp.Workbooks.Open(@path);
            }
            catch (Exception e)
            {
                book = null;

                Console.WriteLine(e + "エクセルファイルを開けませんでした。");

                excelApp.Quit();
            }
            return book;
        }
        void GraphGenerate(Worksheet sheet)
        {
            if (sheet != null)
            {
                var chartObjects = sheet.ChartObjects() as ChartObjects;

                if (chartObjects != null)
                {
                    //グラフの大きさ調整
                    var chartObject = chartObjects.Add(125, 25, 1000, 350);

                    var chart = chartObject.Chart;

                    //グラフのタイプを指定する(今は散布図の直線マーカーなし)
                    chart.ChartType = XlChartType.xlXYScatterLinesNoMarkers;

                    Console.WriteLine("グラフにしたいシートの最終行数を入力してください");

                    var setRow = Console.ReadLine();

                    if (sheet.Cells[2, 2].Value == null)
                    {
                        for (int j = 3; j <= int.Parse(setRow); j++)
                        {
                            if (sheet.Cells[j, 2].Value != null)
                            {
                                sheet.Cells[2, 2].Value = sheet.Cells[j, 2].Value;

                                break;
                            }
                        }
                    }
                    chart.SetSourceData(sheet.Range("D21:D" + setRow));

                    Series series = (Series)chart.SeriesCollection(1);

                    series.XValues = sheet.get_Range("C2:C" + setRow);
                    series.Values = sheet.get_Range("D21:D" + setRow);
                }
            }
        }
    }
    class PathData
    {
        public string ExcelPath { get; set; } = "";
    }
}