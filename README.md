This is a library for **Google Apps Script** projects. It provides methods to perform an [Exponential backoff](https://developers.google.com/drive/v3/web/handle-errors#exponential-backoff) logic whenever it is needed and rewrite error objects before sending them to Stackdriver Logging.


## Methods


### expBackoff(func)

In some cases, Google APIs can return errors on which it make sense to retry, like '_We're sorry, a server error occurred. Please wait a bit and try again._'.

In such cases, it make sense to wrap the call to the API in an Exponential backoff logic to retry multiple times.

Note that most Google Apps Script services, like GmailApp and SpreadsheetApp already have an Exponential backoff logic implemented by Google. Thus it does not make sense to wrap every call to those services in an Exponential backoff logic.

Things are different for Google Apps Script [advanced services](https://developers.google.com/apps-script/guides/services/advanced#enabling_advanced_services) where such logic is not implemented. Thus this method is mostly useful for calls linked to Google Apps Script advanced services.


#### Example

```JS

// Calls an anonymous function that gets the subject of the vacation responder in Gmail for the currently authenticated user.

 var responseSubject = ErrorHandler.expBackoff(function(){return Gmail.Users.Settings.getVacation("me").responseSubject;});

```


#### Parameters


<table>
  <tr>
   <td>Name
   </td>
   <td>Type
   </td>
   <td>Description
   </td>
  </tr>
  <tr>
   <td>Func
   </td>
   <td>Function
   </td>
   <td>The anonymous or named function to call.
   </td>
  </tr>
</table>



#### Return 

The value returned by the called function or nothing if the called function isn't returning anything.


### urlFetchWithExpBackOff(url, params)

This method works exactly like the [fetch()](https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app#fetchurl-params) method of Apps Script UrlFetchApp service.

It simply wraps the fetch() call in an Exponential backoff logic.

We advise to replace all existing calls to `UrlFetchApp.fetch()` by this new method.

`UrlFetchApp.fetch(url, params)` => `ErrorHandler.urlFetchWithExpBackOff(url, params)`


### logError(e, additionalParams)

When exception logging is enabled, unhandled exceptions are automatically sent to Stackdriver Logging with a stack trace ([see official documentation](https://developers.google.com/apps-script/guides/logging#exception_logging)).

But if you wrap your code in a try...catch statement and use `console.error(error)` to send the exception to Stackdriver Logging, only the error message will be logged and not the stack trace.

We advise to replace all existing calls to `console.error(error)` by this new method to correctly log the stack trace, in the exact same format used for unhandled exceptions.

`console.error(error)` => `ErrorHandler.logError(error)`


#### Parameters


<table>
  <tr>
   <td>Name
   </td>
   <td>Type
   </td>
   <td>Description
   </td>
  </tr>
  <tr>
   <td>e
   </td>
   <td>String || Error || {lineNumber: number, fileName: string, responseCode: string}
   </td>
   <td>A standard error object retrieved in a try...catch statement or created with new Error() or a simple error message or an object
   </td>
  </tr>
</table>



## Setup

You can copy the code of this library in your own Google Apps Script project or reuse it as a [standard library](https://developers.google.com/apps-script/guides/libraries). In both cases, methods are called using the ErrorHandler class / namespace, meaning you will use them exactly in the same way.

To install it as a library, use the following script ID and select the latest version:

`1mpYgNprGHnJW0BSPclqNIIErVaTyXQ7AjdmLaaE5O9s9bAOrsY14PMCy`

To copy the code in your project, simply copy-past the content of this file in a new script file in your project:

`https://github.com/RomainVialard/ErrorHandler/blob/master/src/ErrorHandler.gs.js`

NPM install:

To be documented


## Warning

This library contains 3 methods directly available as functions and callable without using the ErrorHandler class / namespace:



*   expBackoff()
*   urlFetchWithExpBackOff()
*   logError()

For this reason, if you copy the code in your project, make sure you don't have any other function with the exact same name.
