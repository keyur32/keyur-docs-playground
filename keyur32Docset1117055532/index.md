# Custom Functions tutorial

```typescript
//Call Translate
function TRANSLATE(text, locale) {
    return new OfficeExtension.Promise(function (resolve) {

        //Use Fetch API to call Microsoft Translation API
        fetch(`https://sdx-demo1.azurewebsites.net/api/translate-api?name=${text}&locale=${locale}`)
            .then((response) => response.json())
            .then((responseJson) => {
                // return result
                resolve(responseJson);
            })
            .catch((error) => {
                debugger;
            });
    });
}

```
