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
 * @param {Object || {addonName: string}} [additionalParams]
 */
function logError(e, additionalParams) {
  e = (typeof e === 'string') ? new Error(e) : e;
  
  // Localize error message
  var partialMatches = [];
  var normalizedMessage = ErrorHandler.getNormalizedError(e.message, partialMatches);
  var message = normalizedMessage || e.message;
  
  var locale;
  try {
    locale = Session.getActiveUserLocale();
  }
  catch(err) {
    // Try to add the locale
    locale = ErrorHandler.getErrorLocale(e.message);
  }
  
  var log = {
    context: {
      locale: locale || '',
      originalMessage: e.message,
      translated: !!normalizedMessage,
    }
  };
  
  
  // Add partialMatches if any
  if (partialMatches.length) {
    log.context.variables = {};
    
    partialMatches.forEach(function (match) {
      log.context.variables[match.variable] = match.value;
    });
  }
  
  if (e.name) {
    // examples of error name: Error, ReferenceError, Exception, GoogleJsonResponseException
    // would be nice to categorize
    log.context.errorName = e.name;
    message = e.name +": "+ message;
  }
  log.message = message;
  
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


/**
 * Return the english version of the error if listed in this library
 * 
 * @type {string} localizedErrorMessage
 * @type {Array<{
 *   variable: string,
 *   value: string
 * }>} partialMatches - Pass an empty array, getNormalizedError() will populate it with found extracted variables in case of a partial match
 * 
 * @return {ErrorHandler_.NORMALIZED_ERROR | ''} the error in English or '' if no matching error was found
 */
function getNormalizedError(localizedErrorMessage, partialMatches) {
  /**
   * @type {ErrorHandler_.ErrorMatcher}
   */
  var error = ErrorHandler_._ERROR_MESSAGE_TRANSLATIONS[localizedErrorMessage];
  
  if (error) return error.ref;
  if (typeof localizedErrorMessage !== 'string') return '';
  
  // No exact match, try to execute a partial match
  var match;
  
  /**
   * @type {ErrorHandler_.PartialMatcher}
   */
  var matcher;
  
  for (var i = 0; matcher = ErrorHandler_._ERROR_PARTIAL_MATCH[i]; i++) {
    // search for a match
    match = localizedErrorMessage.match(matcher.regex);
    if (match) break;
  }
  
  // No match found
  if (!match) return '';
  
  // Extract partial match variables
  for (var j = 0, variable; variable = matcher.variables[j] ; j++) {
    partialMatches.push({
      variable: variable,
      value: match[j+1] !== undefined && match[j+1] || ''
    });
  }
  
  return matcher.ref;
}

/**
 * Try to find the locale of the localized thrown error
 * 
 * @type {string} localizedErrorMessage
 * 
 * @return {string | ''} return the locale or '' if no matching error found
 */
function getErrorLocale(localizedErrorMessage) {
  /**
   * @type {ErrorHandler_.ErrorMatcher}
   */
  var error = ErrorHandler_._ERROR_MESSAGE_TRANSLATIONS[localizedErrorMessage];
  
  if (error) return error.locale;
  if (typeof localizedErrorMessage !== 'string') return '';
  
  // No exact match, try to execute a partial match
  var match;
  
  /**
   * @type {ErrorHandler_.PartialMatcher}
   */
  var matcher;
  
  for (var i = 0; matcher = ErrorHandler_._ERROR_PARTIAL_MATCH[i]; i++) {
    // search for a match
    match = localizedErrorMessage.match(matcher.regex);
    if (match) break;
  }
  
  // No match found
  if (!match) return '';
  
  return matcher.locale;
}

/**
 * @typedef {string} ErrorHandler_.NORMALIZED_ERROR
 */
/**
 * @type {Object<ErrorHandler_.NORMALIZED_ERROR>}
 */
NORMALIZED_ERROR = {
  CONDITIONNAL_RULE_REFERENCE_DIF_SHEET: "Conditional format rule cannot reference a different sheet.",
  SERVER_ERROR_RETRY_LATER: "We're sorry, a server error occurred. Please wait a bit and try again.",
  EMPTY_RESPONSE: "Empty response",
  LIMIT_EXCEEDED: "Limit Exceeded: .",
  SERVICE_INVOKED_TOO_MANY_TIMES_EMAIL: "Service invoked too many times for one day: email.",
  
  // Partial match error
  INVALID_EMAIL: 'Invalid email',
};


// noinspection JSUnusedGlobalSymbols, ThisExpressionReferencesGlobalObjectJS
this['ErrorHandler'] = {
  // Add local alias to run the library as normal code
  expBackoff: expBackoff,
  urlFetchWithExpBackOff: urlFetchWithExpBackOff,
  logError: logError,
  
  getNormalizedError: getNormalizedError,
  getErrorLocale: getErrorLocale,
  NORMALIZED_ERROR: NORMALIZED_ERROR,
};

//<editor-fold desc="# Private methods">

var ErrorHandler_ = {};


// noinspection ThisExpressionReferencesGlobalObjectJS
/**
 * Get GAS global object: top-level this
 */
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


// noinspection NonAsciiCharacters, JSNonASCIINames
/**
 * Map all different error translation to their english counterpart,
 * Thanks to Google AppsScript throwing localized errors, it's impossible to easily catch them and actually do something to fix it for our users.
 * 
 * @type {Object<ErrorHandler_.ErrorMatcher>}
 */
ErrorHandler_._ERROR_MESSAGE_TRANSLATIONS = {
  // "Conditional format rule cannot reference a different sheet."
  "Conditional format rule cannot reference a different sheet.": { ref: NORMALIZED_ERROR.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'en'},
  "Quy tắc định dạng có điều kiện không thể tham chiếu một trang tính khác.": { ref: NORMALIZED_ERROR.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'vi'},
  "La regla de formato condicional no puede hacer referencia a una hoja diferente.": { ref: NORMALIZED_ERROR.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'es'},
  "La regola di formattazione condizionale non può contenere un riferimento a un altro foglio.": { ref: NORMALIZED_ERROR.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'it'},
  "La règle de mise en forme conditionnelle ne doit pas faire référence à une autre feuille.": { ref: NORMALIZED_ERROR.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'fr'},
  "Die Regel für eine bedingte Formatierung darf sich nicht auf ein anderes Tabellenblatt beziehen.": { ref: NORMALIZED_ERROR.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'de'},
  "Правило условного форматирования не может ссылаться на другой лист.": { ref: NORMALIZED_ERROR.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'ru'},
  
  // "We're sorry, a server error occurred. Please wait a bit and try again."
  "We're sorry, a server error occurred. Please wait a bit and try again.": { ref: NORMALIZED_ERROR.SERVER_ERROR_RETRY_LATER, locale: 'en'},
  "Spiacenti. Si è verificato un errore del server. Attendi e riprova.": { ref: NORMALIZED_ERROR.SERVER_ERROR_RETRY_LATER, locale: 'it'},
  "Une erreur est survenue sur le serveur. Nous vous prions de nous en excuser et vous invitons à réessayer ultérieurement.": { ref: NORMALIZED_ERROR.SERVER_ERROR_RETRY_LATER, locale: 'fr'},
  "Xin lỗi bạn, máy chủ đã gặp lỗi. Vui lòng chờ một lát và thử lại.": { ref: NORMALIZED_ERROR.SERVER_ERROR_RETRY_LATER, locale: 'vi'},
  "Lo sentimos, se ha producido un error en el servidor. Espera un momento y vuelve a intentarlo.": { ref: NORMALIZED_ERROR.SERVER_ERROR_RETRY_LATER, locale: 'es'},
  "ขออภัย มีข้อผิดพลาดของเซิร์ฟเวอร์เกิดขึ้น โปรดรอสักครู่แล้วลองอีกครั้ง": { ref: NORMALIZED_ERROR.SERVER_ERROR_RETRY_LATER, locale: 'th'},
  "很抱歉，伺服器發生錯誤，請稍後再試。": { ref: NORMALIZED_ERROR.SERVER_ERROR_RETRY_LATER, locale: 'zh_TW'},
  "Infelizmente ocorreu um erro do servidor. Espere um momento e tente novamente.": { ref: NORMALIZED_ERROR.SERVER_ERROR_RETRY_LATER, locale: 'pt'},
  
  // "Empty response"
  "Empty response": { ref: NORMALIZED_ERROR.EMPTY_RESPONSE, locale: 'en'},
  "Respuesta vacía": { ref: NORMALIZED_ERROR.EMPTY_RESPONSE, locale: 'es'},
  "Réponse vierge": { ref: NORMALIZED_ERROR.EMPTY_RESPONSE, locale: 'fr'},
  "Câu trả lời trống": { ref: NORMALIZED_ERROR.EMPTY_RESPONSE, locale: 'vi'},
  "Resposta vazia": { ref: NORMALIZED_ERROR.EMPTY_RESPONSE, locale: 'pt'},
  "Prázdná odpověď": { ref: NORMALIZED_ERROR.EMPTY_RESPONSE, locale: 'cs'},
  
  // "Limit Exceeded: ."
  "Limit Exceeded: .": { ref: NORMALIZED_ERROR.LIMIT_EXCEEDED, locale: 'en'},
  "Límite excedido: .": { ref: NORMALIZED_ERROR.LIMIT_EXCEEDED, locale: 'es'},
  "Limite dépassée : .": { ref: NORMALIZED_ERROR.LIMIT_EXCEEDED, locale: 'fr'},
  
  // "Service invoked too many times for one day: email."
  "Service invoked too many times for one day: email.": { ref: NORMALIZED_ERROR.SERVICE_INVOKED_TOO_MANY_TIMES_EMAIL, locale: 'en'},
  "Trop d'appels pour ce service aujourd'hui : email.": { ref: NORMALIZED_ERROR.SERVICE_INVOKED_TOO_MANY_TIMES_EMAIL, locale: 'fr'},
  "Servicio solicitado demasiadas veces en un mismo día: gmail.": { ref: NORMALIZED_ERROR.SERVICE_INVOKED_TOO_MANY_TIMES_EMAIL, locale: 'es'},
  "Serviço chamado muitas vezes no mesmo dia: email.": { ref: NORMALIZED_ERROR.SERVICE_INVOKED_TOO_MANY_TIMES_EMAIL, locale: 'pt'},
  
  
};

/**
 * @type {Array<ErrorHandler_.PartialMatcher>}
 */
ErrorHandler_._ERROR_PARTIAL_MATCH = [
  {regex: /^Invalid email: (.*)$/,
    variables: ['email'],
    ref: NORMALIZED_ERROR.INVALID_EMAIL,
    locale: 'en'},
];

/**
 * @typedef {{}} ErrorHandler_.PartialMatcher
 * 
 * @property {RegExp} regex - Regex describing the error
 * @property {Array<string>} variables - Ordered list naming the successive extracted value by the regex groups
 * @property {ErrorHandler_.NORMALIZED_ERROR} ref - Error reference
 * @property {string} locale - Error locale
 */
/**
 * @typedef {{}} ErrorHandler_.ErrorMatcher
 * 
 * @property {ErrorHandler_.NORMALIZED_ERROR} ref - Error reference
 * @property {string} locale - Error locale
 */


//</editor-fold>
