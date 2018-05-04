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
 * getNormalizedError()
 * getErrorLocale()
 * 
 * NORMALIZED_ERRORS
 * NORETRY_ERRORS
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
 * @param {{}} [options] - options for exponential backoff
 * @param {boolean} options.throwOnFailure - default to FALSE, if true, throw the ErrorHandler_.CustomError on failure
 * @param {boolean} options.doNotLogKnownErrors - default to FALSE, if true, will not log known errors to stackdriver
 * @param {boolean} options.verbose - default to FALSE, if true, will log a warning on a successful call that failed at least once
 * @param {number} options.retryNumber - default to 5, maximum number of retry on error
 *
 * @return {* | ErrorHandler_.CustomError} - The value returned by the called function, or ErrorHandler_.CustomError on failure if throwOnFailure == false
 */
function expBackoff(func, options) {
  
  // enforce defaults
  options = options || {};
  
  var retry = options.retryNumber || 5;
  if (retry < 1 || retry > 6) retry = 5;
  
  var previousError = null;
  var retryDelay = null;
  var oldRetryDelay = null;
  var customError;
  
  // execute func() then retry <retry> times at most if errors
  for (var n = 0; n <= retry; n++) {
    // actual exponential backoff
    n && Utilities.sleep(retryDelay || (Math.pow(2, n-1) * 1000) + (Math.round(Math.random() * 1000)));
    
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
    if (noError) {
      if (n && options.verbose){
        var info = {
          context: "Exponential Backoff",
          successful: true,
          numberRetry: n,
        };
        
        retryDelay && (info.retryDelay = retryDelay);
        
        ErrorHandler.logError(previousError, info, {
          asWarning: true,
          doNotLogKnownErrors: options.doNotLogKnownErrors,
        });
      }
      
      return response;
    }
    previousError = error;
    oldRetryDelay = retryDelay;
    retryDelay = null;
    
    
    // Process error retry
    if (!isUrlFetchResponse && error.message) {
      var variables = [];
      var normalizedError = ErrorHandler.getNormalizedError(error.message, variables);
      
      // If specific error that explicitly give the retry time
      if (normalizedError === ErrorHandler.NORMALIZED_ERRORS.USER_RATE_LIMIT_EXCEEDED_RETRY_AFTER_SPECIFIED_TIME && variables[0] && variables[0].value) {
        retryDelay = (new Date(variables[0].value) - new Date()) + 1000;
        
        oldRetryDelay && ErrorHandler.logError(error, {
          failReason: 'Failed after waiting '+ oldRetryDelay +'ms',
          context: "Exponential Backoff",
          numberRetry: n,
          retryDelay: retryDelay,
        }, {
          asWarning: true,
          doNotLogKnownErrors: options.doNotLogKnownErrors,
        });
        
        // Do not wait too long
        if (retryDelay < 32000) continue;
        
        customError = ErrorHandler.logError(error, {
          failReason: 'Retry delay > 31s',
          context: "Exponential Backoff",
          numberRetry: n,
          retryDelay: retryDelay,
        }, {doNotLogKnownErrors: options.doNotLogKnownErrors});
        
        if (options.throwOnFailure) throw customError;
        return customError;
      }
      
      // Check for errors thrown by Google APIs on which there's no need to retry
      // eg: "Access denied by a security policy established by the administrator of your organization. 
      //      Please contact your administrator for further assistance."
      if (!ErrorHandler.NORETRY_ERRORS[normalizedError]) continue;
      
      customError = ErrorHandler.logError(error, {
        failReason: 'No retry needed',
        numberRetry: n,
        context: "Exponential Backoff"
      }, {doNotLogKnownErrors: options.doNotLogKnownErrors});
      
      if (options.throwOnFailure) throw customError;
      return customError;
    }
    
  }
  
  
  // Action after last re-try
  if (isUrlFetchResponse) {
    ErrorHandler.logError(new Error(response.getContentText()), {
      failReason: 'Max retries reached',
      urlFetchWithMuteHttpExceptions: true,
      context: "Exponential Backoff"
    }, {doNotLogKnownErrors: options.doNotLogKnownErrors});
    
    return response;
  }
  
  
  // Investigate on errors that are still happening after 5 retries
  // Especially error "Not Found" - does it make sense to retry on it?
  customError = ErrorHandler.logError(error, {
    failReason: 'Max retries reached',
    context: "Exponential Backoff"
  }, {doNotLogKnownErrors: options.doNotLogKnownErrors});
  
  if (options.throwOnFailure) throw customError;
  return customError;
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
 * @typedef {Error} ErrorHandler_.CustomError
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

/**
 * If we simply log the error object, only the error message will be submitted to Stackdriver Logging
 * Best to re-write the error as a new object to get lineNumber & stack trace
 *
 * @param {String || Error || {lineNumber: number, fileName: string, responseCode: string}} error
 * @param {Object || {addonName: string, versionNumber: number}} [additionalParams]
 *
 * @param {{}} [options] - Options for logError
 * @param {boolean} options.asWarning - default to FALSE, use console.warn instead console.error
 * @param {boolean} options.doNotLogKnownErrors - default to FALSE, if true, will not log known errors to stackdriver
 *
 * @return {ErrorHandler_.CustomError}
 */
function logError(error, additionalParams, options) {
  options = options || {};
  
  error = (typeof error === 'string') ? new Error(error) : error;
  
  // Localize error message
  var partialMatches = [];
  var normalizedMessage = ErrorHandler.getNormalizedError(error.message, partialMatches);
  var message = normalizedMessage || error.message;
  
  var locale;
  var scriptId;
  try {
    locale = Session.getActiveUserLocale();
    scriptId = ScriptApp.getScriptId();
  }
  catch(err) {
    // Try to add the locale
    locale = ErrorHandler.getErrorLocale(error.message);
  }
  
  var log = {
    context: {
      locale: locale || '',
      originalMessage: error.message,
      knownError: !!normalizedMessage,
    }
  };
  
  
  // Add partialMatches if any
  if (partialMatches.length) {
    log.context.variables = {};
    
    partialMatches.forEach(function (match) {
      log.context.variables[match.variable] = match.value;
    });
  }
  
  if (error.name) {
    // examples of error name: Error, ReferenceError, Exception, GoogleJsonResponseException
    // would be nice to categorize
    log.context.errorName = error.name;
    message = error.name +": "+ message;
  }
  log.message = message;
  
  // allow to use a global variable instead of passing the addonName in each call
  // noinspection JSUnresolvedVariable
  var addonName = additionalParams && additionalParams.addonName || ErrorHandler_._this['SCRIPT_PROJECT_TITLE'] || '';
  
  // Manage error Stack
  if (error.lineNumber && error.fileName && error.stack) {
    var fileName = addonName && error.fileName.replace(' ('+ addonName +')', '') || error.fileName;
    
    log.context.reportLocation = {
      lineNumber: error.lineNumber,
      filePath: fileName,
      directLink: 'https://script.google.com/macros/d/'+ scriptId +'/edit?f='+ fileName +'&s='+ error.lineNumber
    };
    
    var res = ErrorHandler_._convertErrorStack(error.stack, addonName);
    log.context.reportLocation.functionName = res.lastFunctionName;
    log.message+= '\n    '+ res.stack;
  }
  
  if (error.responseCode) {
    log.context.responseCode = error.responseCode;
  }
  
  // allow to use a global variable instead of passing the addonName in each call
  // noinspection JSUnresolvedVariable
  var versionNumber = additionalParams && additionalParams.versionNumber || ErrorHandler_._this['SCRIPT_VERSION_DEPLOYED'] || '';
  if (versionNumber) {
    log.serviceContext = {
      version: versionNumber
    };
  }
  
  // Add custom information
  if (additionalParams) {
    log.customParams = {};
    
    for (var i in additionalParams) {
      log.customParams[i] = additionalParams[i];
    }
  }
  
  // Send error to stackdriver log
  if (!options.doNotLogKnownErrors || !normalizedMessage) {
    if (options.asWarning) console.warn(log);
    else console.error(log);
  }
  
  // Return an error, with context
  var customError = new Error(normalizedMessage || error.message);
  customError.context = log.context;
  
  return customError;
}


/**
 * Return the english version of the error if listed in this library
 *
 * @type {string} localizedErrorMessage
 * @type {Array<{
 *   variable: string,
 *   value: string
 * }>} [partialMatches] - Pass an empty array, getNormalizedError() will populate it with found extracted variables in case of a partial match
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
  if (partialMatches && Array.isArray(partialMatches)) {
    for (var j = 0, variable; variable = matcher.variables[j]; j++) {
      partialMatches.push({
        variable: variable, value: match[j + 1] !== undefined && match[j + 1] || ''
      });
    }
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
 * List all known Errors
 */
NORMALIZED_ERRORS = {
  CONDITIONNAL_RULE_REFERENCE_DIF_SHEET: "Conditional format rule cannot reference a different sheet.",
  SERVER_ERROR_RETRY_LATER: "We're sorry, a server error occurred. Please wait a bit and try again.",
  AUTHORIZATION_REQUIRED: "Authorization is required to perform that action. Please run the script again to authorize it.",
  EMPTY_RESPONSE: "Empty response",
  LIMIT_EXCEEDED: "Limit Exceeded: .",
  USER_RATE_LIMIT_EXCEEDED: "User Rate Limit Exceeded",
  RATE_LIMIT_EXCEEDED: "Rate Limit Exceeded",
  NOT_FOUND: "Not Found",
  BACKEND_ERROR: "Backend Error",
  SERVICE_INVOKED_TOO_MANY_TIMES_EMAIL: "Service invoked too many times for one day: email.",
  TRYING_TO_EDIT_PROTECTED_CELL: "You are trying to edit a protected cell or object. Please contact the spreadsheet owner to remove protection if you need to edit.",
  NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND: "No item with the given ID could be found, or you do not have permission to access it.",
  NO_PERMISSION_TO_ACCESS_THE_REQUESTED_DOCUMENT: "You do not have permissions to access the requested document.",
  UNABLE_TO_TALK_TO_TRIGGER_SERVICE: "Unable to talk to trigger service",
  MAIL_SERVICE_NOT_ENABLED: "Mail service not enabled",
  INVALID_THREAD_ID_VALUE: "Invalid thread_id value",
  LABEL_ID_NOT_FOUND: "labelId not found",
  LABEL_NAME_EXISTS_OR_CONFLICTS: "Label name exists or conflicts",
  NO_RECIPIENT: "Failed to send email: no recipient",
  
  // Partial match error
  INVALID_EMAIL: 'Invalid email',
  DOCUMENT_MISSING: 'Document is missing (perhaps it was deleted?)',
  USER_RATE_LIMIT_EXCEEDED_RETRY_AFTER_SPECIFIED_TIME: 'User-rate limit exceeded. Retry after specified time.',
  INVALID_ARGUMENT: 'Invalid argument',
  SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME: 'A sheet with this name already exists. Please enter another name.',
};

/**
 * List all error for which retrying will not make the call succeed
 */
NORETRY_ERRORS = {};
NORETRY_ERRORS[NORMALIZED_ERRORS.INVALID_EMAIL] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.MAIL_SERVICE_NOT_ENABLED] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.NO_RECIPIENT] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.NOT_FOUND] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.SERVICE_INVOKED_TOO_MANY_TIMES_EMAIL] = true;

NORETRY_ERRORS[NORMALIZED_ERRORS.NO_PERMISSION_TO_ACCESS_THE_REQUESTED_DOCUMENT] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.TRYING_TO_EDIT_PROTECTED_CELL] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME] = true;

NORETRY_ERRORS[NORMALIZED_ERRORS.AUTHORIZATION_REQUIRED] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.INVALID_ARGUMENT] = true;


// noinspection JSUnusedGlobalSymbols, ThisExpressionReferencesGlobalObjectJS
this['ErrorHandler'] = {
  // Add local alias to run the library as normal code
  expBackoff: expBackoff,
  urlFetchWithExpBackOff: urlFetchWithExpBackOff,
  logError: logError,
  
  getNormalizedError: getNormalizedError,
  getErrorLocale: getErrorLocale,
  NORMALIZED_ERRORS: NORMALIZED_ERRORS,
  NORETRY_ERRORS: NORETRY_ERRORS,
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
  "Conditional format rule cannot reference a different sheet.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'en'},
  "Quy tắc định dạng có điều kiện không thể tham chiếu một trang tính khác.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'vi'},
  "La regla de formato condicional no puede hacer referencia a una hoja diferente.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'es'},
  "La regola di formattazione condizionale non può contenere un riferimento a un altro foglio.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'it'},
  "La règle de mise en forme conditionnelle ne doit pas faire référence à une autre feuille.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'fr'},
  "Une règle de mise en forme conditionnelle ne peut pas faire référence à une autre feuille.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'fr_ca'},
  "Die Regel für eine bedingte Formatierung darf sich nicht auf ein anderes Tabellenblatt beziehen.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'de'},
  "Правило условного форматирования не может ссылаться на другой лист.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'ru'},
  "조건부 서식 규칙은 다른 시트를 참조할 수 없습니다.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'ko'},
  "條件式格式規則無法參照其他工作表。": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'zh_tw'},
  "条件格式规则无法引用其他工作表。": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'zh_cn'},
  "條件格式規則無法參照其他工作表。": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'zh_hk'},
  "条件付き書式ルールで別のシートを参照することはできません。": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'ja'},
  "Pravidlo podmíněného formátu nemůže odkazovat na jiný list.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'cs'},
  "Nosacījumformāta kārtulai nevar būt atsauce uz citu lapu.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'lv'},
  "Pravidlo podmieneného formátovania nemôže odkazovať na iný hárok.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'sk'},
  "Conditionele opmaakregel kan niet verwijzen naar een ander blad.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'nl'},
  "Ehdollinen muotoilusääntö ei voi viitata toiseen taulukkoon.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'fi'},
  "กฎการจัดรูปแบบตามเงื่อนไขอ้างอิงแผ่นงานอื่นไม่ได้": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'th'},
  "Reguła formatowania warunkowego nie może odwoływać się do innego arkusza.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'pl'},
  "Aturan format bersyarat tidak dapat merujuk ke sheet yang berbeda.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'in'},
  "Villkorsstyrd formateringsregel får inte referera till ett annat arbetsblad.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'sv'},
  "La regla de format condicional no pot fer referència a un altre full.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'ca'},
  "A feltételes formázási szabály nem tud másik munkalapot meghívni.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'hu'},
  "A regra de formatação condicional não pode fazer referência a uma página diferente.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'pt'},
  "Правило умовного форматування не може посилатися на інший аркуш.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'uk'},
  "لا يمكن أن تشير الصيغة الشرطية إلى ورقة مختلفة.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'ar_sa'},
  "Ο κανόνας μορφής υπό συνθήκες δεν μπορεί να αναφέρεται σε διαφορετικό φύλλο.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'el'},
  "En betinget formateringsregel kan ikke referere til et annet ark.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'no'},
  "Koşullu biçimlendirme kuralı farklı bir sayfaya başvuramaz.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'tr'},
  "Pravilo pogojnega oblikovanja se ne more sklicevati na drug list.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'sl'},
  "Hindi maaaring mag-reference ng ibang sheet ang conditional format rule.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'fil'},
  "En betinget formatregel kan ikke henvise til et andet ark.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'da'},
  "כלל של פורמט מותנה לא יכול לכלול הפניה לגיליון אחר.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'iw'},
  "Formatu baldintzatuaren arauak ezin dio egin erreferentzia beste orri bati.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'eu'},
  "Sąlyginio formato taisyklė negali nurodyti kito lapo.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'lt'},
  "Regula cu format condiționat nu poate face referire la altă foaie.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'ro'},
  "Tingimusvormingu reegel ei saa viidata teisele lehele.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'et'},
  
  // "We're sorry, a server error occurred. Please wait a bit and try again."
  "We're sorry, a server error occurred. Please wait a bit and try again.": { ref: NORMALIZED_ERRORS.SERVER_ERROR_RETRY_LATER, locale: 'en'},
  "Spiacenti. Si è verificato un errore del server. Attendi e riprova.": { ref: NORMALIZED_ERRORS.SERVER_ERROR_RETRY_LATER, locale: 'it'},
  "Une erreur est survenue sur le serveur. Nous vous prions de nous en excuser et vous invitons à réessayer ultérieurement.": { ref: NORMALIZED_ERRORS.SERVER_ERROR_RETRY_LATER, locale: 'fr'},
  "Xin lỗi bạn, máy chủ đã gặp lỗi. Vui lòng chờ một lát và thử lại.": { ref: NORMALIZED_ERRORS.SERVER_ERROR_RETRY_LATER, locale: 'vi'},
  "Lo sentimos, se ha producido un error en el servidor. Espera un momento y vuelve a intentarlo.": { ref: NORMALIZED_ERRORS.SERVER_ERROR_RETRY_LATER, locale: 'es'},
  "Lo sentimos, se produjo un error en el servidor. Aguarde un momento e inténtelo de nuevo.": { ref: NORMALIZED_ERRORS.SERVER_ERROR_RETRY_LATER, locale: 'es_419'},
  "ขออภัย มีข้อผิดพลาดของเซิร์ฟเวอร์เกิดขึ้น โปรดรอสักครู่แล้วลองอีกครั้ง": { ref: NORMALIZED_ERRORS.SERVER_ERROR_RETRY_LATER, locale: 'th'},
  "很抱歉，伺服器發生錯誤，請稍後再試。": { ref: NORMALIZED_ERRORS.SERVER_ERROR_RETRY_LATER, locale: 'zh_tw'},
  "Infelizmente ocorreu um erro do servidor. Espere um momento e tente novamente.": { ref: NORMALIZED_ERRORS.SERVER_ERROR_RETRY_LATER, locale: 'pt'},
  "Sajnáljuk, szerverhiba történt. Kérjük, várjon egy kicsit, majd próbálkozzon újra.": { ref: NORMALIZED_ERRORS.SERVER_ERROR_RETRY_LATER, locale: 'hu'},
  "Ett serverfel uppstod. Vänta lite och försök igen.": { ref: NORMALIZED_ERRORS.SERVER_ERROR_RETRY_LATER, locale: 'sv'},
  "A apărut o eroare de server. Așteptați puțin și încercați din nou.": { ref: NORMALIZED_ERRORS.SERVER_ERROR_RETRY_LATER, locale: 'ro'},
  "Ein Serverfehler ist aufgetreten. Bitte versuchen Sie es später erneut.": { ref: NORMALIZED_ERRORS.SERVER_ERROR_RETRY_LATER, locale: 'de'},
  
  // "Authorization is required to perform that action. Please run the script again to authorize it."
  "Authorization is required to perform that action. Please run the script again to authorize it.": { ref: NORMALIZED_ERRORS.AUTHORIZATION_REQUIRED, locale: 'en'},
  "Autorisation requise pour exécuter cette action. Exécutez à nouveau le script pour autoriser cette action.": { ref: NORMALIZED_ERRORS.AUTHORIZATION_REQUIRED, locale: 'fr'},
  "Cần được cho phép để thực hiện tác vụ đó. Hãy chạy lại tập lệnh để cho phép tác vụ.": { ref: NORMALIZED_ERRORS.AUTHORIZATION_REQUIRED, locale: 'vi'},
  
  // "Empty response"
  "Empty response": { ref: NORMALIZED_ERRORS.EMPTY_RESPONSE, locale: 'en'},
  "Respuesta vacía": { ref: NORMALIZED_ERRORS.EMPTY_RESPONSE, locale: 'es'},
  "Réponse vierge": { ref: NORMALIZED_ERRORS.EMPTY_RESPONSE, locale: 'fr'},
  "Câu trả lời trống": { ref: NORMALIZED_ERRORS.EMPTY_RESPONSE, locale: 'vi'},
  "Resposta vazia": { ref: NORMALIZED_ERRORS.EMPTY_RESPONSE, locale: 'pt'},
  "Prázdná odpověď": { ref: NORMALIZED_ERRORS.EMPTY_RESPONSE, locale: 'cs'},
  
  // "Limit Exceeded: ."
  "Limit Exceeded: .": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED, locale: 'en'},
  "Límite excedido: .": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED, locale: 'es'},
  "Limite dépassée : .": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED, locale: 'fr'},
  
  // "User Rate Limit Exceeded" - eg: Gmail.Users.Threads.get
  "User Rate Limit Exceeded": { ref: NORMALIZED_ERRORS.USER_RATE_LIMIT_EXCEEDED, locale: 'en'},
  
  // "Rate Limit Exceeded" - eg: Gmail.Users.Messages.send
  "Rate Limit Exceeded": { ref: NORMALIZED_ERRORS.RATE_LIMIT_EXCEEDED, locale: 'en'},
  
  // "Not Found"
  // with uppercase "f" when calling Gmail.Users.Messages or Gmail.Users.Drafts endpoints
  "Not Found": { ref: NORMALIZED_ERRORS.NOT_FOUND, locale: 'en'},
  // with lowercase "f" when calling Gmail.Users.Threads endpoint
  "Not found": { ref: NORMALIZED_ERRORS.NOT_FOUND, locale: 'en'},
  
  // "Backend Error"
  "Backend Error": { ref: NORMALIZED_ERRORS.BACKEND_ERROR, locale: 'en'},
  
  // "Service invoked too many times for one day: email."
  "Service invoked too many times for one day: email.": { ref: NORMALIZED_ERRORS.SERVICE_INVOKED_TOO_MANY_TIMES_EMAIL, locale: 'en'},
  "Trop d'appels pour ce service aujourd'hui : email.": { ref: NORMALIZED_ERRORS.SERVICE_INVOKED_TOO_MANY_TIMES_EMAIL, locale: 'fr'},
  "Servicio solicitado demasiadas veces en un mismo día: gmail.": { ref: NORMALIZED_ERRORS.SERVICE_INVOKED_TOO_MANY_TIMES_EMAIL, locale: 'es'},
  "Servicio solicitado demasiadas veces en un mismo día: email.": { ref: NORMALIZED_ERRORS.SERVICE_INVOKED_TOO_MANY_TIMES_EMAIL, locale: 'es'},
  "Serviço chamado muitas vezes no mesmo dia: email.": { ref: NORMALIZED_ERRORS.SERVICE_INVOKED_TOO_MANY_TIMES_EMAIL, locale: 'pt'},  
  
  // "You are trying to edit a protected cell or object. Please contact the spreadsheet owner to remove protection if you need to edit."
  "You are trying to edit a protected cell or object. Please contact the spreadsheet owner to remove protection if you need to edit.": { ref: NORMALIZED_ERRORS.TRYING_TO_EDIT_PROTECTED_CELL, locale: 'en'},
  "保護されているセルやオブジェクトを編集しようとしています。編集する必要がある場合は、スプレッドシートのオーナーに連絡して保護を解除してもらってください。": { ref: NORMALIZED_ERRORS.TRYING_TO_EDIT_PROTECTED_CELL, locale: 'ja'},
  "Estás intentando editar una celda o un objeto protegidos. Ponte en contacto con el propietario de la hoja de cálculo para desprotegerla si es necesario modificarla.": { ref: NORMALIZED_ERRORS.TRYING_TO_EDIT_PROTECTED_CELL, locale: 'es'},
  "Vous tentez de modifier une cellule ou un objet protégés. Si vous avez besoin d'effectuer cette modification, demandez au propriétaire de la feuille de calcul de supprimer la protection.": { ref: NORMALIZED_ERRORS.TRYING_TO_EDIT_PROTECTED_CELL, locale: 'fr'},
  
  // "No item with the given ID could be found, or you do not have permission to access it."
  "No item with the given ID could be found, or you do not have permission to access it.": { ref: NORMALIZED_ERRORS.NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND, locale: 'en'},
  "Không tìm thấy mục nào có ID đã cung cấp hoặc bạn không có quyền truy cập vào mục đó.": { ref: NORMALIZED_ERRORS.NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND, locale: 'vi'},
  "No se ha encontrado ningún elemento con el ID proporcionado o no tienes permiso para acceder a él.": { ref: NORMALIZED_ERRORS.NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND, locale: 'es'},
  "No se ha encontrado ningún elemento con la ID proporcionada o no tienes permiso para acceder a él.": { ref: NORMALIZED_ERRORS.NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND, locale: 'es_419'},
  "Nessun elemento trovato con l'ID specificato o non disponi di autorizzazioni per accedervi.": { ref: NORMALIZED_ERRORS.NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND, locale: 'it'},
  "Det gick inte att hitta någon post med angivet ID eller så saknar du behörighet för att få åtkomst till den.": { ref: NORMALIZED_ERRORS.NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND, locale: 'sv'},
  "Er is geen item met de opgegeven id gevonden of je hebt geen toestemming om het item te openen.": { ref: NORMALIZED_ERRORS.NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND, locale: 'nl'},
  "Nenhum item com o ID fornecido foi encontrado ou você não tem permissão para acessá-lo.": { ref: NORMALIZED_ERRORS.NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND, locale: 'pt'},
  "Impossible de trouver l'élément correspondant à cet identifiant. Vous n'êtes peut-être pas autorisé à y accéder.": { ref: NORMALIZED_ERRORS.NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND, locale: 'fr'},
  "No s'ha trobat cap element amb aquest identificador o no teniu permís per accedir-hi.": { ref: NORMALIZED_ERRORS.NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND, locale: 'ca'},
  "Элемент с заданным кодом не найден или у вас нет прав доступа к нему.": { ref: NORMALIZED_ERRORS.NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND, locale: 'ru'},
  "Nebyly nalezeny žádné položky se zadaným ID nebo nemáte oprávnění k nim přistupovat.": { ref: NORMALIZED_ERRORS.NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND, locale: 'cs'},
  "Item dengan ID yang diberikan tidak dapat ditemukan atau Anda tidak memiliki izin untuk mengaksesnya.": { ref: NORMALIZED_ERRORS.NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND, locale: 'in'},
  "指定された ID のアイテムは見つからなかったか、アクセスする権限がありません。": { ref: NORMALIZED_ERRORS.NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND, locale: 'ja'},
  
  // "You do not have permissions to access the requested document."
  "You do not have permissions to access the requested document.": { ref: NORMALIZED_ERRORS.NO_PERMISSION_TO_ACCESS_THE_REQUESTED_DOCUMENT, locale: 'en'},
  
  // "Unable to talk to trigger service"
  "Unable to talk to trigger service": { ref: NORMALIZED_ERRORS.UNABLE_TO_TALK_TO_TRIGGER_SERVICE, locale: 'en'},
  "Impossible de communiquer pour déclencher le service": { ref: NORMALIZED_ERRORS.UNABLE_TO_TALK_TO_TRIGGER_SERVICE, locale: 'fr'},
  
  // "Mail service not enabled"
  "Mail service not enabled": { ref: NORMALIZED_ERRORS.MAIL_SERVICE_NOT_ENABLED, locale: 'en'},
  
  // "Invalid thread_id value"
  "Invalid thread_id value": { ref: NORMALIZED_ERRORS.INVALID_THREAD_ID_VALUE, locale: 'en'},
  
  // "labelId not found"
  "labelId not found": { ref: NORMALIZED_ERRORS.LABEL_ID_NOT_FOUND, locale: 'en'},
  
  // "Label name exists or conflicts"
  "Label name exists or conflicts": { ref: NORMALIZED_ERRORS.LABEL_NAME_EXISTS_OR_CONFLICTS, locale: 'en'},
  
  // "Invalid to header" - eg: Gmail.Users.Messages.send
  "Invalid to header": { ref: NORMALIZED_ERRORS.INVALID_EMAIL, locale: 'en'},
  // "Invalid cc header" - eg: Gmail.Users.Messages.send
  "Invalid cc header": { ref: NORMALIZED_ERRORS.INVALID_EMAIL, locale: 'en'},  
  
  // "Failed to send email: no recipient" - eg: GmailApp.sendEmail()
  "Failed to send email: no recipient": { ref: NORMALIZED_ERRORS.NO_RECIPIENT, locale: 'en'},
  // "Recipient address required" - eg: Gmail.Users.Messages.send()
  "Recipient address required": { ref: NORMALIZED_ERRORS.NO_RECIPIENT, locale: 'en'},
  
};

/**
 * @type {Array<ErrorHandler_.PartialMatcher>}
 */
ErrorHandler_._ERROR_PARTIAL_MATCH = [
  // Invalid email: XXX
  {regex: /^Invalid email: (.*)$/,
    variables: ['email'],
    ref: NORMALIZED_ERRORS.INVALID_EMAIL,
    locale: 'en'},
  {regex: /^El correo electrónico no es válido: (.*)$/,
    variables: ['email'],
    ref: NORMALIZED_ERRORS.INVALID_EMAIL,
    locale: 'es'},
  {regex: /^無效的電子郵件：(.*)$/,
    variables: ['email'],
    ref: NORMALIZED_ERRORS.INVALID_EMAIL,
    locale: 'zh_TW'},
  {regex: /^E-mail incorrect : (.*)$/,
    variables: ['email'],
    ref: NORMALIZED_ERRORS.INVALID_EMAIL,
    locale: 'fr'},
    
  // Document XXX is missing (perhaps it was deleted?)
  {regex: /^Document (\S*) is missing \(perhaps it was deleted\?\)$/,
    variables: ['docId'],
    ref: NORMALIZED_ERRORS.DOCUMENT_MISSING,
    locale: 'en'},
  {regex: /^Documento (\S*) mancante \(forse è stato eliminato\?\)$/,
    variables: ['docId'],
    ref: NORMALIZED_ERRORS.DOCUMENT_MISSING,
    locale: 'it'},
  {regex: /^Tài liệu (\S*) bị thiếu \(có thể tài liệu đã bị xóa\?\)$/,
    variables: ['docId'],
    ref: NORMALIZED_ERRORS.DOCUMENT_MISSING,
    locale: 'vi'},
    
  // User-rate limit exceeded. Retry after XXX
  {regex: /^(?:Limit Exceeded: : )?User-rate limit exceeded\.\s+Retry after (.*Z)/,
    variables: ['timestamp'],
    ref: NORMALIZED_ERRORS.USER_RATE_LIMIT_EXCEEDED_RETRY_AFTER_SPECIFIED_TIME,
    locale: 'en'},
  
  // "Invalid argument: XXX" - wrong email alias used - eg: GmailApp.sendEmail()
  {regex: /^Invalid argument: (.*)$/,
    variables: ['email'],
    ref: NORMALIZED_ERRORS.INVALID_ARGUMENT,
    locale: 'en'},
  
  // "A sheet with the name "XXX" already exists. Please enter another name." - eg: [Sheet].setName()
  {regex: /^A sheet with the name "([^"]*)" already exists\. Please enter another name\.$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'en'},
  {regex: /^Ya existe una hoja con el nombre "([^"]*)"\. Ingresa otro\.$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'es_419'},
  
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
  
