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

# [TypeScript](#tab/ts)

Content for Windows...

# [C#](#tab/csharp)

Content for Linux...

# [VBA](#tab/vba)

Content for Windows...

---
