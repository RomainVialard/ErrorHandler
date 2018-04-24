This is a library for **Google Apps Script** projects. It provides methods to perform an [Exponential backoff](https://developers.google.com/drive/v3/web/handle-errors#exponential-backoff) logic whenever it is needed and rewrite error objects before sending them to Stackdriver Logging.


## Methods


### expBackoff(func, options)

In some cases, Google APIs can return errors on which it make sense to retry, like '_We're sorry, a server error occurred. Please wait a bit and try again._'.

In such cases, it make sense to wrap the call to the API in an Exponential backoff logic to retry multiple times.

Note that most Google Apps Script services, like GmailApp and SpreadsheetApp already have an Exponential backoff logic implemented by Google. Thus it does not make sense to wrap every call to those services in an Exponential backoff logic.

Things are different for Google Apps Script [advanced services](https://developers.google.com/apps-script/guides/services/advanced#enabling_advanced_services) where such logic is not implemented. Thus this method is mostly useful for calls linked to Google Apps Script advanced services.


#### Example

```javascript

// Calls an anonymous function that gets the subject of the vacation responder in Gmail for the currently authenticated user.

 var responseSubject = ErrorHandler.expBackoff(function(){return Gmail.Users.Settings.getVacation("me").responseSubject;});

```


#### Parameters

Name    | Type     | Default value | Description
--------|----------|---------------|------------
func    | Function |               | The anonymous or named function to call.
options | Object   | {}            | OPTIONAL, options for exponential backoff
options.throwOnFailure      | boolean | false | If true, throw a CustomError on failure
options.doNotLogKnownErrors | boolean | false | If true, will not log known errors to stackdriver
options.verbose             | boolean | false | If true, will log a warning on a successful call that failed at least once
options.retryNumber         | number  | 5     | Maximum number of retry on error

#### Return 

The value returned by the called function, or a CustomError on failure if options.throwOnFailure == false

CustomError is an instance of Error


### urlFetchWithExpBackOff(url, params)

This method works exactly like the [fetch()](https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app#fetchurl-params) method of Apps Script UrlFetchApp service.

It simply wraps the fetch() call in an Exponential backoff logic.

We advise to replace all existing calls to `UrlFetchApp.fetch()` by this new method.

`UrlFetchApp.fetch(url, params)` => `ErrorHandler.urlFetchWithExpBackOff(url, params)`


### logError(error, additionalParams)

When exception logging is enabled, unhandled exceptions are automatically sent to Stackdriver Logging with a stack trace ([see official documentation](https://developers.google.com/apps-script/guides/logging#exception_logging)).

But if you wrap your code in a try...catch statement and use `console.error(error)` to send the exception to Stackdriver Logging, only the error message will be logged and not the stack trace.
This method will also leverage the list of known errors and their translation to better aggregate on Stackdriver Logging: the message will be the English version if available.

We advise to replace all existing calls to `console.error(error)` by this new method to correctly log the stack trace, in the exact same format used for unhandled exceptions.

`console.error(error)` => `ErrorHandler.logError(error)`


#### Parameters

Name    | Type     | Default value | Description
--------|----------|---------------|------------
error   | String, Error, or { lineNumber: number, fileName: string, responseCode: string} | | A standard error object retrieved in a try...catch statement or created with new Error() or a simple error message or an object
additionalParams | Object, {addonName: string} | {} | Add custom values to the logged Error, the 'addonName' property will pass the AddonName to remove from error fileName
options | Object   | {}            | Options for logError
options.asWarning | boolean | false | If true, use console.warn instead console.error
options.doNotLogKnownErrors | booleab | false | if true, will not log known errors to stackdriver


#### Return

Return the CustomError, which is an Error, with a property 'context', as defined below:
```javascript
/**
 * @typedef {Error} CustomError
 *
 * @property {{
 *   locale: string,
 *   originalMessage: string,
 *   knownError: boolean,
 *   variables: Array<{}>,
 *   errorName: string,
 *   reportLocation: {
 *     lineNumber: number,
 *     filePath: string,
 *     directLink: string,
 *   },
 * }} context
 */
```


### getNormalizedError(localizedErrorMessage, partialMatches)

Try to get the corresponding english version of the error if listed in this library.

There are Error match exactly, and Errors that can be partially matched, for example when it contains a variable part (eg: a document ID, and email)

To get variable part, use the following pattern:
```javascript
var variables = []; // The empty array will be filled if necessary in the function
var normalizedError = ErrorHandler.getNormalizedError('Documento 1234567890azerty mancante (forse è stato eliminato?)', variables);

// The normalized Error, with its message in English, or '' of no match are found:
// normalizedError = 'Document is missing (perhaps it was deleted?)'

// the variable part:
// variables = [{variable: 'docId', value: '1234567890azerty'}]
```


#### Parameters

Name                  | Type     | Default value | Description
----------------------|----------|---------------|------------
localizedErrorMessage | String   |               | The Error message in the user's locale
partialMatches        | Array<{ variable: string, value: string }>    | [] | OPTIONAL, Pass an empty array, getNormalizedError() will populate it with found extracted variables in case of a partial match


The value returned is the error in English or '' if no matching error was found


### getErrorLocale(localizedErrorMessage)

Try to find the locale of the localized thrown error


#### Parameters

Name                  | Type     | Default value | Description
----------------------|----------|---------------|------------
localizedErrorMessage | String   |               | The Error message in the user's locale


#### Return

The locale ('en', 'it', ...) or '' if no matching error found

### NORMALIZED_ERRORS
constant

List all known Errors in a fixed explicit English message.
Serves as reference for locallizing Errors.

Use this to check which error a normalized error is:
```javascript
var normalizedError = ErrorHandler.getNormalizedError('Documento 1234567890azerty mancante (forse è stato eliminato?)');

if (normalizedError === ErrorHandler.NORMALIZED_ERRORS.DOCUMENT_MISSING) {
  // Do something on document missing error
}
```

### NORETRY_ERRORS
constant

List all Errors for which there are no benefit in re-trying

## Setup

You can copy the code of this library in your own Google Apps Script project or reuse it as a [standard library](https://developers.google.com/apps-script/guides/libraries). In both cases, methods are called using the ErrorHandler class / namespace, meaning you will use them exactly in the same way.

To install it as a library, use the following script ID and select the latest version:

`1mpYgNprGHnJW0BSPclqNIIErVaTyXQ7AjdmLaaE5O9s9bAOrsY14PMCy`

To copy the code in your project, simply copy-past the content of this file in a new script file in your project:

`https://github.com/RomainVialard/ErrorHandler/blob/master/src/ErrorHandler.gs.js`

NPM install:

To be documented


## Warning

This library contains 5 methods directly available as functions and callable without using the ErrorHandler class / namespace:

*   expBackoff()
*   urlFetchWithExpBackOff()
*   logError()
*   getNormalizedError()
*   getErrorLocale()

and 2 Object constant

* NORMALIZED_ERRORS
* NORETRY_ERRORS


For this reason, if you copy the code in your project, make sure you don't have any other function with the exact same name.
