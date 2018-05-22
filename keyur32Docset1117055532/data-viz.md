# Create data visualizations

Help users understand their data by helping them create data visualizations like native Excel charts, or even use web technologies to bring in your own custom data visualizations.


![People Graph](https://az158878.vo.msecnd.net/marketing/Partner_21474836617/Product_42949674809/Asset_e208e0a1-cffd-44fa-9a34-c673b52d84b4/App0212.png)


# [JavaScript/TypeScript](#tab/js)

### Custom Data Visualizations

You can use the Office Javascript API to create your own custom visualizations.  Start with the either the Charting APIs available in Excel 1.7

```javascript

    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem("Sample");
        var dataRange = sheet.getRange("A1:B13");
        var chart = sheet.charts.add("Line", dataRange, "auto");

        chart.title.text = "Sales Data";
        chart.legend.position = "right"
        chart.legend.format.fill.setSolidColor("white");
        chart.dataLabels.format.font.size = 15;
        chart.dataLabels.format.font.color = "black";

        return context.sync();
    }).catch(errorHandlerFunction);
```

> [!div class="nextstepaction"]
> [Try it out](http://dev.office.com)

#### Learning Path
1. Check out the [Charting APIs](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-charts)
2. Learn about [Office add-ins developement](https://dev.office.com)
3. Explore custom data visualizations on [AppSource](https://appsource.microsoft.com/en-us/marketplace/apps?product=power-bi-visuals)

### PowerBI Custom Visuals

You can quickly create custom visuals in Excel and Power BI.

![Visuals](https://powerbicdn.azureedge.net/mediahandler/blog/media/PowerBI/blog/2051210e-e17b-4320-b2bb-cb4bbf391563.jpg)

> [!div class="nextstepaction"]
> [Try it out custom visuals](http://dev.office.com)

#### Learning Path
1. Check out the [Power BI Custom Visualizations tutorial](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-charts)
2. Explore the various [Power BI Custom Visualization Samples](https://github.com/Microsoft/PowerBI-visuals)
3. Explore custom Power BI visualizations on [AppSource](https://appsource.microsoft.com/en-us/marketplace/apps?product=power-bi-visuals)


# [C#](#tab/csharp)

### Create custom visuals using the .NET Chart APIs

The following code snippet will show you how you can create a simple chart by using the Excel Client Library:

```csharp
    using System;
    using System.Windows.Forms;
    using Excel = Microsoft.Office.Interop.Excel; 

    namespace WindowsApplication1
    {
        public partial class Form1 : Form
        {
            public Form1()
            {
                InitializeComponent();
            }

            private void button1_Click(object sender, EventArgs e)
            {

                //start Excel
                Excel.Application xlApp ;
                Excel.Workbook xlWorkBook ;
                Excel.Worksheet xlWorkSheet ;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                //add data 
                xlWorkSheet.Cells[1, 1] = "";
                xlWorkSheet.Cells[1, 2] = "Student1";
                xlWorkSheet.Cells[1, 3] = "Student2";
                xlWorkSheet.Cells[1, 4] = "Student3";

                xlWorkSheet.Cells[2,t 1] = "Term1";
                xlWorkSheet.Cells[2, 2] = "80";
                xlWorkSheet.Cells[2, 3] = "65";
                xlWorkSheet.Cells[2, 4] = "45";

                xlWorkSheet.Cells[3, 1] = "Term2";
                xlWorkSheet.Cells[3, 2] = "78";
                xlWorkSheet.Cells[3, 3] = "72";
                xlWorkSheet.Cells[3, 4] = "60";

                xlWorkSheet.Cells[4, 1] = "Term3";
                xlWorkSheet.Cells[4, 2] = "82";
                xlWorkSheet.Cells[4, 3] = "80";
                xlWorkSheet.Cells[4, 4] = "65";

                xlWorkSheet.Cells[5, 1] = "Term4";
                xlWorkSheet.Cells[5, 2] = "75";
                xlWorkSheet.Cells[5, 3] = "82";
                xlWorkSheet.Cells[5, 4] = "68";

                //Create a chart
                Excel.Range chartRange ; 
                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
                Excel.Chart chartPage = myChart.Chart;

                chartRange = xlWorkSheet.get_Range("A1", "d5");
                chartPage.SetSourceData(chartRange, misValue);
                chartPage.ChartType = Excel.XlChartType.xlColumnClustered; 

                //Save the workbook and close excel
                xlWorkBook.SaveAs("chart.xlsx", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                MessageBox.Show("Excel file created!");
            }

            private void releaseObject(object obj)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    obj = null;
                }
                catch (Exception ex)
                {
                    obj = null;
                    MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
                }
                finally
                {
                    GC.Collect();
                }
            }

        }
    }
```

> [!div class="nextstepaction"]
> [Learn More](http://dev.office.com)
