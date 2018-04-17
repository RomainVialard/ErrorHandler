/****************************************************************
 * ErrorHandler library
 * https://github.com/RomainVialard/ErrorHandler
 *
 * Performs exponential backoff when needed
 * And makes sure that caught errors are correctly logged in Stackdriver
 *
 * expBackoff()
 * urlFetchWithExpBackOff()
 * logError()
 *
 * _convertErrorStack()
 *****************************************************************/


/**
 * Invokes a function, performing up to 5 retries with exponential backoff.
 * Retries with delays of approximately 1, 2, 4, 8 then 16 seconds for a total of
 * about 32 seconds before it gives up and rethrows the last error.
 * See: https://developers.google.com/google-apps/documents-list/#implementing_exponential_backoff
 * Original author: peter.herrmann@gmail.com (Peter Herrmann)
 *
 * @example
 * // Calls an anonymous function that concatenates a greeting with the current Apps user's email
 * ErrorHandler.expBackoff(function(){return "Hello, " + Session.getActiveUser().getEmail();});
 *
 * @example
 * // Calls an existing function
 * ErrorHandler.expBackoff(myFunction);
 *
 * @param {Function} func - The anonymous or named function to call.
 * 
 * @return {*} - The value returned by the called function.
 */
function expBackoff(func) {
  
  // execute func() then retry 5 times at most if errors
  for (var n = 0; n < 6; n++) {
    // actual exponential backoff
    n && Utilities.sleep((Math.pow(2, n-1) * 1000) + (Math.round(Math.random() * 1000)));
    
    var response = undefined;
    var error = undefined;
    
    var noError = true;
    var isUrlFetchResponse = false;
    
    // Try / catch func()
    try { response = func() }
    catch(err) {
      error = err;
      noError = false;
    }
    
    
    // Handle retries on UrlFetch calls with muteHttpExceptions
    if (noError && response && typeof response.getResponseCode === "function") {
      isUrlFetchResponse = true;
      
      var responseCode = response.getResponseCode();
      
      // Only perform retries on error 500 for now
      if (responseCode === 500) {
        error = response;
        noError = false;
      }
    }
    
    // Return result that is not an error
    if (noError) return response;
    
    
    // Process error retry
    if (!isUrlFetchResponse && error.message) {
      // Check for errors thrown by Google APIs on which there's no need to retry
      // eg: "Access denied by a security policy established by the administrator of your organization. 
      //      Please contact your administrator for further assistance."
      if (error.message.indexOf('Invalid requests') !== -1
          || error.message.indexOf('Access denied') !== -1
          || error.message.indexOf('Mail service not enabled') !== -1) {
        throw error;
      }
      
      // TODO: YAMM specific ?: move to YAMM code
      else if (error.message.indexOf('response too large') !== -1) {
        // Thrown after calling Gmail.Users.Threads.get()
        // maybe because a specific thread contains too many messages
        // best to skip the thread
        return null;
      }
    }
    
  }
  
  
  // Action after last re-try
  if (isUrlFetchResponse) {
    
    ErrorHandler.logError(new Error(response.getContentText()), {
      shouldInvestigate: true,
      failedAfter5Retries: true,
      urlFetchWithMuteHttpExceptions: true,
      context: "Exponential Backoff"
    });
    
    return response;
  }
  
  else {
    // 'User-rate limit exceeded' is always followed by 'Retry after' + timestamp
    // Maybe we should parse the timestamp to check how long we need to wait 
    // and if we should abort or not
    // 'User Rate Limit Exceeded' (without '-') isn't followed by 'Retry after' and it makes sense to retry
    if (error.message && error.message.indexOf('User-rate limit exceeded') !== -1) {
      ErrorHandler.logError(error, {
        shouldInvestigate: true,
        failedAfter5Retries: true,
        context: "Exponential Backoff"
      });
      
      return null;
    }
    
    // Investigate on errors that are still happening after 5 retries
    // Especially error "Not Found" - does it make sense to retry on it?
    ErrorHandler.logError(error, {
      failedAfter5Retries: true,
      context: "Exponential Backoff"
    });
    
    throw error;
  }
}

/**
 * Helper function to automatically handles exponential backoff on UrlFetch use
 *
 * @param {string} url
 * @param {Object} params
 * 
 * @return {UrlFetchApp.HTTPResponse}  - fetch response
 */
function urlFetchWithExpBackOff(url, params) {
  params = params || {};
  
  params.muteHttpExceptions = true;
  
  return ErrorHandler.expBackoff(function(){
    return UrlFetchApp.fetch(url, params);
  });
}

/**
 * If we simply log the error object, only the error message will be submitted to Stackdriver Logging
 * Best to re-write the error as a new object to get lineNumber & stack trace
 * 
 * @param {String || Error || {lineNumber: number, fileName: string, responseCode: string}} e
 * @param {Object || {addonName: string}} additionalParams
 */
function logError(e, additionalParams) {
  e = (typeof e === 'string') ? new Error(e) : e;
  
  var log = {
    context: {}
  };
  
  if (e.name) {
    // examples of error name: Error, ReferenceError, Exception, GoogleJsonResponseException
    // would be nice to categorize
    log.context.errorName = e.name;
    e.message = e.name +": "+ e.message;
  }
  log.message = e.message;
  
  // Manage error Stack
  if (e.lineNumber && e.fileName && e.stack) {
    log.context.reportLocation = {
      lineNumber: e.lineNumber,
      filePath: e.fileName
    };
    
    var addonName = additionalParams && additionalParams.addonName || undefined;
    
    var res = ErrorHandler_._convertErrorStack(e.stack, addonName);
    log.context.reportLocation.functionName = res.lastFunctionName;
    log.message+= '\n    '+ res.stack;
  }
  
  if (e.responseCode) {
    log.context.responseCode = e.responseCode;
  }
  
  // Add custom information
  if (additionalParams) {
    log.customParams = {};
    
    for (var i in additionalParams) {
      log.customParams[i] = additionalParams[i];
    }
  }
  
  // Send error to stackdriver log
  console.error(log);
}

var ErrorHandler_ = {};
ErrorHandler_._err1 = "Conditional format rule cannot reference a different sheet.";
ErrorHandler_._err2 = "Invalid email";

var errorMessageTranslations = {
  // "Conditional format rule cannot reference a different sheet."
  "Quy tắc định dạng có điều kiện không thể tham chiếu một trang tính khác.": ErrorHandler_._err1,
  "La regla de formato condicional no puede hacer referencia a una hoja diferente.": ErrorHandler_._err1,
  "La regola di formattazione condizionale non può contenere un riferimento a un altro foglio.": ErrorHandler_._err1,
  "La règle de mise en forme conditionnelle ne doit pas faire référence à une autre feuille.": ErrorHandler_._err1,
  "Die Regel für eine bedingte Formatierung darf sich nicht auf ein anderes Tabellenblatt beziehen.": ErrorHandler_._err1,
  "Правило условного форматирования не может ссылаться на другой лист.": ErrorHandler_._err1,
  // "Invalid email"
  "E-mail incorrect": ErrorHandler_._err2,
  "E-mail inválido": ErrorHandler_._err2
};

// noinspection JSUnusedGlobalSymbols, ThisExpressionReferencesGlobalObjectJS
this['ErrorHandler'] = {
  // Add local alias to run the library as normal code
  expBackoff: expBackoff,
  urlFetchWithExpBackOff: urlFetchWithExpBackOff,
  logError: logError,
  errorMessageTranslations: errorMessageTranslations
};

//<editor-fold desc="# Private methods">

// Get GAS global object: top-level this
// noinspection ThisExpressionReferencesGlobalObjectJS
ErrorHandler_._this = this;

/**
 * Format stack:
 * "at File Name:lineNumber (myFunction)" => "at myFunction(File Name:lineNumber)"
 *
 * @param {string} stack - Stack given by GAS with console.stack
 * @param {string} [addonName] - Optional Add-on name added by GAS to the stacks: will remove it from output stack
 *
 * @return {{
 *   stack: string,
 *   lastFunctionName: string
 * }} - formatted stack and last functionName executed
 */
ErrorHandler_._convertErrorStack = function (stack, addonName) {
  // allow to use a global variable instead of passing the addonName in each call
  // noinspection JSUnresolvedVariable
  addonName = addonName || ErrorHandler_._this['SCRIPT_PROJECT_TITLE'] || '';
  
  var formattedStack = [];
  var lastFunctionName = '';
  var res;
  var regex = new RegExp('at\\s([^:]+?)'+ (addonName ? '(?:\\s\\('+ addonName +'\\))?' : '') +':(\\d+)(?:\\s\\(([^)]+)\\))?', 'gm');
  
  while (res = regex.exec(stack)) {
    var [/* total match */, fileName, lineNumber, functionName] = res;
    
    if (!lastFunctionName) lastFunctionName = functionName || '';
    
    formattedStack.push('at '+ (functionName || '[unknown function]') +'('+ fileName +':'+ lineNumber +')');
  }
  
  return {
    stack: formattedStack.join('\n    '),
    lastFunctionName: lastFunctionName
  };
};

//</editor-fold>
