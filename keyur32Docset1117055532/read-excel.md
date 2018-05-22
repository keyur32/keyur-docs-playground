#Read and analyze data in Excel files

One of the most common scenarios you can do in Excel today is to read and analyze your data. 

This typically involves:
a) Iterating over a number of rows and columns that make up a  Table or a often times, a simple set of values in the data.
b) Then depending on your scenario, you can do operations like custom calculations, or batch/send that data to an external service for additional processing (i.e. save my excel data to Sql Azure)


Pick a language below to see how you can make custom functions in Excel.

# [JavaScript](#tab/js)

You can make a custom function in an Excel Add-in using the Office Javascript API.  To use custom functions, you're users will need Office 365 or Office Online.

### Example 
Here's a simple script that will get the range of of values.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sheet1");]

    //try sheet.getUsedRange() to get the set of used cells in the worksheet in Excel
    var range = sheet.getRange("A1:D10000");
    range.load("address");

    return context.sync()
        .then(function () {

            //this is where you can have your for loop and iterate over the range[i][j]
            console.log(`The address of the range B2:C5 is "${range.address}"`);

            //properties to try
            //range[i][j].value <- underlying value
            //range[i][j].formuala <- formula
            //range[i][j].address <- address of current cell in the range
        });
}).catch(errorHandlerFunction);
```

> [!div class="nextstepaction"]
> [Try it out!](http://dev.office.com)



### Learning Path

1. [Join the developer program to get Office 365](https://aka.ms/o365devprogram)
2. [Learn about Excel add-ins](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview)
3. [Learn about working with Ranges](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-ranges)
5. [Learn about how you can deploy your add-in to your users](https://docs.microsoft.com/en-us/office/dev/add-ins/publish/publish)


# [C#](#tab/c#)


The following code snippet will show you how you can read a simple range by using the Excel Client Library:

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
                Excel.Range xlRange = xlWorksheet.UsedRange;

                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!
                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        //new line
                        if (j == 1)
                            Console.Write("\r\n");

                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");

                        //add useful things here!   
                    }
                }


                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

            }
        }
    }

```

> [!div class="nextstepaction"]
> [Learn More](http://dev.office.com)
