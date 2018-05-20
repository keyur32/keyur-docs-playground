#Create a custom function

Custom functions (similar to user-defined functions, or UDFs), enable developers to add any JavaScript function to Excel using an add-in. Users can then access custom functions like any other native function in Excel (such as =SUM()). This article explains how to create custom functions in Excel.

The following illustration shows you how an end user would insert a custom function into a cell. The function that adds 42 to a pair of numbers.

# [JavaScript](#tab/js)

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

There are a few options you can use to create custom functions in Excel.
1. Create an Automation Add-in. This leverages windows COM.

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
