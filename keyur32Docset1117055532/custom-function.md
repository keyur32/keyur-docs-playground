#Create a custom function

Custom functions (similar to user-defined functions, or UDFs), enable developers to add any JavaScript function to Excel using an add-in. Users can then access custom functions like any other native function in Excel (such as =SUM()). This article explains how to create custom functions in Excel.

The following illustration shows you how an end user would insert a custom function into a cell. The function that adds 42 to a pair of numbers.

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />


Pick a language below to see how you can make custom functions in Excel.

# [JavaScript](#tab/js)

You can make a custom function in an Excel Add-in using the Office Javascript API.  To use custom functions, you're users will need Office 365 or Office Online.


Here's a simple function that you can call directly from Excel.

```javascript
function ADD42(a, b) {
    return a + b + 42;
}
```

> [!div class="nextstepaction"]
> [Try it out!](http://dev.office.com)


## Learning Path
1. [Join the developer program to get Office 365](https://aka.ms/o365devprogram)
2. [Learn about Excel add-ins](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview)
3. [Try out custom function calls in ScriptLab](https://appsource.microsoft.com/en-us/product/office/WA104380862?tab=Overview)
4. [Watch the 2018 Build video by Michael Saunders](https://channel9.msdn.com/events/Build/2018/BRK2419?term=excel%20)
5. [Learn about how you can deploy your add-in to your users](https://docs.microsoft.com/en-us/office/dev/add-ins/publish/publish)


## Partners

# [TypeScript](#tab/ts)

Here's a simple function that you can call directly from Excel.

```typescript
function ADD42(a, b) {
    return a + b + 42;
}
```

> [!div class="nextstepaction"]
> [Try it out!](http://dev.office.com)


## Learning Path
1. [Join the developer program to get Office 365](https://aka.ms/o365devprogram)
2. [Learn about Excel add-ins](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview)
3. [Try out custom function calls in ScriptLab](https://appsource.microsoft.com/en-us/product/office/WA104380862?tab=Overview)
4. [Watch the 2018 Build video by Michael Saunders](https://channel9.msdn.com/events/Build/2018/BRK2419?term=excel%20)
5. [Learn about how you can deploy your add-in to your users](https://docs.microsoft.com/en-us/office/dev/add-ins/publish/publish)


# [C#](#tab/csharp)

While there isn't direct support for C# user defefined functions. are a few options you can use to create custom functions in Excel, and will work for users of Office for Windows on 2007+.


1. Create an [Automation Add-in](https://support.microsoft.com/en-us/help/291392/excel-com-add-ins-and-automation-add-ins). Excel 'Automation add-ins' are essential COM Add-ins that also add Excel custom function capabilities. 
2. Create an XLL add-in and wrap that in a .NET com DLL.  Or leverage, an open source library such as [ExcelDNA.net](https://excel-dna.net/) that abstracts much of this for you. 

```csharp
 public double Sum42(double a, double b)
 {
      return a + b + 42; 
 }
```


# [VBA](#tab/vba)

Content for Windows...

# [C](#tab/c)


You can build custom functions in C and C++ by calling the Excel Native APIs.

```c++
double ADD42(LPXLOPER12 a, LPXLOPER12 b)
{   
    return a.val + b.val + 42;
}
```

> [!div class="nextstepaction"]
> [See XLL SDK documentation](https://msdn.microsoft.com/en-us/library/office/bb687883.aspx)



---
