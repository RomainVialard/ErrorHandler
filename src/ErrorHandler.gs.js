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
    // examples of error name: Error, ReferenceError, Exception, GoogleJsonResponseException, HttpResponseException
    // would be nice to categorize
    log.context.errorName = error.name;
    if (error.name === "HttpResponseException") {
      // In this case message is usually very long as it contains the HTML of the error response page
      // eg: 'Response Code: 502. Message: <!DOCTYPE html> <html lang=en>'
      // for now, shorten and only retrieve response code
      message = message.split('.')[0];
    }
    message = error.name + ": " + message;
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
  if (matcher.variables && partialMatches && Array.isArray(partialMatches)) {
    for (var index = 0, variable; variable = matcher.variables[index]; index++) {
      partialMatches.push({
        variable: variable, value: match[index + 1] !== undefined && match[index + 1] || ''
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
  // Google Sheets
  CONDITIONNAL_RULE_REFERENCE_DIF_SHEET: "Conditional format rule cannot reference a different sheet.",
  TRYING_TO_EDIT_PROTECTED_CELL: "You are trying to edit a protected cell or object. Please contact the spreadsheet owner to remove protection if you need to edit.",
  RANGE_NOT_FOUND: "Range not found",
  RANGE_COORDINATES_ARE_OUTSIDE_SHEET_DIMENSIONS: "The coordinates of the range are outside the dimensions of the sheet.",
  RANGE_COORDINATES_INVALID: "The coordinates or dimensions of the range are invalid.",

  // Google Drive
  NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND: "No item with the given ID could be found, or you do not have permission to access it.",
  NO_PERMISSION_TO_ACCESS_THE_REQUESTED_DOCUMENT: "You do not have permissions to access the requested document.",
  LIMIT_EXCEEDED_DRIVEAPP: "Limit Exceeded: DriveApp.",

  // Gmail / email service
  MAIL_SERVICE_NOT_ENABLED: "Mail service not enabled",
  INVALID_THREAD_ID_VALUE: "Invalid thread_id value",
  LABEL_ID_NOT_FOUND: "labelId not found",
  LABEL_NAME_EXISTS_OR_CONFLICTS: "Label name exists or conflicts",
  INVALID_LABEL_NAME: "Invalid label name",
  NO_RECIPIENT: "Failed to send email: no recipient",
  IMAP_FEATURES_DISABLED_BY_ADMIN: "IMAP features disabled by administrator",
  LIMIT_EXCEEDED_MAX_RECIPIENTS_PER_MESSAGE: "Limit Exceeded: Email Recipients Per Message.",
  LIMIT_EXCEEDED_EMAIL_BODY_SIZE: "Limit Exceeded: Email Body Size.",
  LIMIT_EXCEEDED_EMAIL_TOTAL_ATTACHMENTS_SIZE: "Limit Exceeded: Email Total Attachments Size.",
  LIMIT_EXCEEDED_EMAIL_SUBJECT_LENGTH: "Argument too large: subject",
  GMAIL_NOT_DEFINED: "\"Gmail\" is not defined.",
  GMAIL_OPERATION_NOT_ALLOWED: "Gmail operation not allowed.",

  // Google Calendar
  CALENDAR_SERVICE_NOT_ENABLED: "Calendar service not enabled",

  // miscellaneous
  SERVER_ERROR_RETRY_LATER: "We're sorry, a server error occurred. Please wait a bit and try again.",
  AUTHORIZATION_REQUIRED: "Authorization is required to perform that action. Please run the script again to authorize it.",
  EMPTY_RESPONSE: "Empty response",
  BAD_VALUE: "Bad value",
  LIMIT_EXCEEDED: "Limit Exceeded: .",
  USER_RATE_LIMIT_EXCEEDED: "User Rate Limit Exceeded",
  RATE_LIMIT_EXCEEDED: "Rate Limit Exceeded",
  NOT_FOUND: "Not Found",
  BAD_REQUEST: "Bad Request",
  BACKEND_ERROR: "Backend Error",
  UNABLE_TO_TALK_TO_TRIGGER_SERVICE: "Unable to talk to trigger service",
  ACTION_NOT_ALLOWED_THROUGH_EXEC_API: "Script has attempted to perform an action that is not allowed when invoked through the Google Apps Script Execution API.",
  TOO_MANY_LOCK_OPERATIONS: "There are too many LockService operations against the same script.",

  // Partial match error
  INVALID_EMAIL: 'Invalid email',
  DOCUMENT_MISSING: 'Document is missing (perhaps it was deleted?)',
  USER_RATE_LIMIT_EXCEEDED_RETRY_AFTER_SPECIFIED_TIME: 'User-rate limit exceeded. Retry after specified time.',
  DAILY_LIMIT_EXCEEDED: "Daily Limit Exceeded",
  SERVICE_INVOKED_TOO_MANY_TIMES_FOR_ONE_DAY: "Service invoked too many times for one day.",
  SERVICE_UNAVAILABLE: "Service unavailable",
  SERVICE_ERROR: "Service error",
  INVALID_ARGUMENT: 'Invalid argument',
  SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME: 'A sheet with this name already exists. Please enter another name.',
  API_NOT_ENABLED: 'Project is not found and cannot be used for API calls',
};

/**
 * List all error for which retrying will not make the call succeed
 */
NORETRY_ERRORS = {};
NORETRY_ERRORS[NORMALIZED_ERRORS.INVALID_EMAIL] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.MAIL_SERVICE_NOT_ENABLED] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.GMAIL_NOT_DEFINED] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.NO_RECIPIENT] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.LIMIT_EXCEEDED_MAX_RECIPIENTS_PER_MESSAGE] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.LIMIT_EXCEEDED_EMAIL_BODY_SIZE] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.LIMIT_EXCEEDED_EMAIL_TOTAL_ATTACHMENTS_SIZE] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.LIMIT_EXCEEDED_EMAIL_SUBJECT_LENGTH] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.NOT_FOUND] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.BAD_REQUEST] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.SERVICE_INVOKED_TOO_MANY_TIMES_FOR_ONE_DAY] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.IMAP_FEATURES_DISABLED_BY_ADMIN] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.LABEL_NAME_EXISTS_OR_CONFLICTS] = true;

NORETRY_ERRORS[NORMALIZED_ERRORS.NO_PERMISSION_TO_ACCESS_THE_REQUESTED_DOCUMENT] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.TRYING_TO_EDIT_PROTECTED_CELL] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.RANGE_NOT_FOUND] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.RANGE_COORDINATES_ARE_OUTSIDE_SHEET_DIMENSIONS] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.RANGE_COORDINATES_INVALID] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME] = true;

NORETRY_ERRORS[NORMALIZED_ERRORS.CALENDAR_SERVICE_NOT_ENABLED] = true;

NORETRY_ERRORS[NORMALIZED_ERRORS.AUTHORIZATION_REQUIRED] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.INVALID_ARGUMENT] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.ACTION_NOT_ALLOWED_THROUGH_EXEC_API] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.DAILY_LIMIT_EXCEEDED] = true;
NORETRY_ERRORS[NORMALIZED_ERRORS.API_NOT_ENABLED] = true;

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
  "Правило за условно обликовање не може да указује на другу табелу.": { ref: NORMALIZED_ERRORS.CONDITIONNAL_RULE_REFERENCE_DIF_SHEET, locale: 'sr'},

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
  "É necessária autorização para executar esta ação. Execute o script novamente para autorizar a ação.": { ref: NORMALIZED_ERRORS.AUTHORIZATION_REQUIRED, locale: 'pt'},

  // "Empty response"
  "Empty response": { ref: NORMALIZED_ERRORS.EMPTY_RESPONSE, locale: 'en'},
  "Respuesta vacía": { ref: NORMALIZED_ERRORS.EMPTY_RESPONSE, locale: 'es'},
  "Réponse vierge": { ref: NORMALIZED_ERRORS.EMPTY_RESPONSE, locale: 'fr'},
  "Câu trả lời trống": { ref: NORMALIZED_ERRORS.EMPTY_RESPONSE, locale: 'vi'},
  "Resposta vazia": { ref: NORMALIZED_ERRORS.EMPTY_RESPONSE, locale: 'pt'},
  "Prázdná odpověď": { ref: NORMALIZED_ERRORS.EMPTY_RESPONSE, locale: 'cs'},
  "Răspuns gol": { ref: NORMALIZED_ERRORS.EMPTY_RESPONSE, locale: 'ro_MD'},

  // "Bad value"
  "Bad value": { ref: NORMALIZED_ERRORS.BAD_VALUE, locale: 'en'},
  "Helytelen érték": { ref: NORMALIZED_ERRORS.BAD_VALUE, locale: 'hu'},
  "Valor incorrecto": { ref: NORMALIZED_ERRORS.BAD_VALUE, locale: 'es'},
  "Giá trị không hợp lệ": { ref: NORMALIZED_ERRORS.BAD_VALUE, locale: 'vi'},
  "Valeur incorrecte": { ref: NORMALIZED_ERRORS.BAD_VALUE, locale: 'fr'},
  "Valor inválido": { ref: NORMALIZED_ERRORS.BAD_VALUE, locale: 'pt'},

  // "Limit Exceeded: ." - eg: Gmail App.sendEmail()
  "Limit Exceeded: .": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED, locale: 'en'},
  "Límite excedido: .": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED, locale: 'es'},
  "Limite dépassée : .": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED, locale: 'fr'},
  "超過上限：。": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED, locale: 'zh_TW'},

  // "Limit Exceeded: Email Recipients Per Message." - eg: Gmail App.sendEmail()
  "Limit Exceeded: Email Recipients Per Message.": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED_MAX_RECIPIENTS_PER_MESSAGE, locale: 'en'},
  "Sınır Aşıldı: İleti Başına E-posta Alıcısı.": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED_MAX_RECIPIENTS_PER_MESSAGE, locale: 'tr'},
  "Đã vượt quá giới hạn: Người nhận email trên mỗi thư.": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED_MAX_RECIPIENTS_PER_MESSAGE, locale: 'vi'},
  "Límite excedido: Destinatarios de correo electrónico por mensaje.": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED_MAX_RECIPIENTS_PER_MESSAGE, locale: 'es'},

  // "Limit Exceeded: Email Body Size." - eg: Gmail App.sendEmail()
  "Limit Exceeded: Email Body Size.": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED_EMAIL_BODY_SIZE, locale: 'en'},
  "Límite Excedido: Tamaño del cuerpo del mensaje.": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED_EMAIL_BODY_SIZE, locale: 'es_PE'},

  // "Limit Exceeded: Email Total Attachments Size." - eg: Gmail App.sendEmail()
  "Limit Exceeded: Email Total Attachments Size.": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED_EMAIL_TOTAL_ATTACHMENTS_SIZE, locale: 'en'},
  "Límite excedido: Tamaño total de los archivos adjuntos del correo electrónico.": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED_EMAIL_TOTAL_ATTACHMENTS_SIZE, locale: 'es'},

  // "Argument too large: subject" - eg: Gmail App.sendEmail()
  "Argument too large: subject": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED_EMAIL_SUBJECT_LENGTH, locale: 'en'},
  "Argument trop grand : subject": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED_EMAIL_SUBJECT_LENGTH, locale: 'fr'},
  "Argumento demasiado grande: subject": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED_EMAIL_SUBJECT_LENGTH, locale: 'es'},

  // "User Rate Limit Exceeded" - eg: Gmail.Users.Threads.get
  "User Rate Limit Exceeded": { ref: NORMALIZED_ERRORS.USER_RATE_LIMIT_EXCEEDED, locale: 'en'},

  // "Rate Limit Exceeded" - eg: Gmail.Users.Messages.send
  "Rate Limit Exceeded": { ref: NORMALIZED_ERRORS.RATE_LIMIT_EXCEEDED, locale: 'en'},

  // "Not Found"
  // with uppercase "f" when calling Gmail.Users.Messages or Gmail.Users.Drafts endpoints
  "Not Found": { ref: NORMALIZED_ERRORS.NOT_FOUND, locale: 'en'},
  // with lowercase "f" when calling Gmail.Users.Threads endpoint
  "Not found": { ref: NORMALIZED_ERRORS.NOT_FOUND, locale: 'en'},
  "Não encontrado": { ref: NORMALIZED_ERRORS.NOT_FOUND, locale: 'pt_PT'},
  "No se ha encontrado.": { ref: NORMALIZED_ERRORS.NOT_FOUND, locale: 'es'},

  // "Bad Request" - eg: all 'list' requests from Gmail advanced service, maybe if there are 0 messages in Gmail (new account)
  "Bad Request": { ref: NORMALIZED_ERRORS.BAD_REQUEST, locale: 'en'},

  // "Backend Error"
  "Backend Error": { ref: NORMALIZED_ERRORS.BACKEND_ERROR, locale: 'en'},

  // "You are trying to edit a protected cell or object. Please contact the spreadsheet owner to remove protection if you need to edit."
  "You are trying to edit a protected cell or object. Please contact the spreadsheet owner to remove protection if you need to edit.": { ref: NORMALIZED_ERRORS.TRYING_TO_EDIT_PROTECTED_CELL, locale: 'en'},
  "保護されているセルやオブジェクトを編集しようとしています。編集する必要がある場合は、スプレッドシートのオーナーに連絡して保護を解除してもらってください。": { ref: NORMALIZED_ERRORS.TRYING_TO_EDIT_PROTECTED_CELL, locale: 'ja'},
  "Estás intentando editar una celda o un objeto protegidos. Ponte en contacto con el propietario de la hoja de cálculo para desprotegerla si es necesario modificarla.": { ref: NORMALIZED_ERRORS.TRYING_TO_EDIT_PROTECTED_CELL, locale: 'es'},
  "Vous tentez de modifier une cellule ou un objet protégés. Si vous avez besoin d'effectuer cette modification, demandez au propriétaire de la feuille de calcul de supprimer la protection.": { ref: NORMALIZED_ERRORS.TRYING_TO_EDIT_PROTECTED_CELL, locale: 'fr'},
  "Эта область защищена. Чтобы изменить ее, обратитесь к владельцу таблицы.": { ref: NORMALIZED_ERRORS.TRYING_TO_EDIT_PROTECTED_CELL, locale: 'ru'},
  "Estás intentando modificar una celda o un objeto protegido. Si necesitas realizar cambios, comunícate con el propietario de la hoja de cálculo para que quite la protección.": { ref: NORMALIZED_ERRORS.TRYING_TO_EDIT_PROTECTED_CELL, locale: 'es_MX'},
  "Você está tentando editar uma célula ou um objeto protegido. Se precisar editar, entre em contato com o proprietário da planilha para remover a proteção.": { ref: NORMALIZED_ERRORS.TRYING_TO_EDIT_PROTECTED_CELL, locale: 'pt'},
  "Покушавате да измените заштићену ћелију или објекат. Контактирајте власника табеле да уклони заштиту ако треба да унесете измене.": { ref: NORMALIZED_ERRORS.TRYING_TO_EDIT_PROTECTED_CELL, locale: 'sr'},

  // "Range not found" - eg: Range.getValue()
  "Range not found": { ref: NORMALIZED_ERRORS.RANGE_NOT_FOUND, locale: 'en'},
  "Range  not found": { ref: NORMALIZED_ERRORS.RANGE_NOT_FOUND, locale: 'en_GB'},
  "No se ha encontrado el intervalo.": { ref: NORMALIZED_ERRORS.RANGE_NOT_FOUND, locale: 'es'},
  "Intervalo não encontrado": { ref: NORMALIZED_ERRORS.RANGE_NOT_FOUND, locale: 'pt'},
  "Không tìm thấy dải ô": { ref: NORMALIZED_ERRORS.RANGE_NOT_FOUND, locale: 'vi'},
  "Plage introuvable": { ref: NORMALIZED_ERRORS.RANGE_NOT_FOUND, locale: 'fr'},
  "Vahemikku ei leitud": { ref: NORMALIZED_ERRORS.RANGE_NOT_FOUND, locale: 'et'},

  // "The coordinates of the range are outside the dimensions of the sheet."
  "The coordinates of the range are outside the dimensions of the sheet.": { ref: NORMALIZED_ERRORS.RANGE_COORDINATES_ARE_OUTSIDE_SHEET_DIMENSIONS, locale: 'en'},
  "As coordenadas do intervalo estão fora das dimensões da página.": { ref: NORMALIZED_ERRORS.RANGE_COORDINATES_ARE_OUTSIDE_SHEET_DIMENSIONS, locale: 'pt'},
  "Tọa độ của dải ô nằm ngoài kích thước của trang tính.": { ref: NORMALIZED_ERRORS.RANGE_COORDINATES_ARE_OUTSIDE_SHEET_DIMENSIONS, locale: 'vi'},

  // "The coordinates or dimensions of the range are invalid."
  "The coordinates or dimensions of the range are invalid.": { ref: NORMALIZED_ERRORS.RANGE_COORDINATES_INVALID, locale: 'en'},

  // "No item with the given ID could be found, or you do not have permission to access it." - eg:Drive App.getFileById
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
  "Не вдалося знайти елемент із зазначеним ідентифікатором. Або у вас немає дозволу на доступ до нього.": { ref: NORMALIZED_ERRORS.NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND, locale: 'uk'},
  "Verilen kimliğe sahip öğe bulunamadı veya bu öğeye erişme iznine sahip değilsiniz.": { ref: NORMALIZED_ERRORS.NO_ITEM_WITH_GIVEN_ID_COULD_BE_FOUND, locale: 'tr'},

  // "You do not have permissions to access the requested document."
  "You do not have permissions to access the requested document.": { ref: NORMALIZED_ERRORS.NO_PERMISSION_TO_ACCESS_THE_REQUESTED_DOCUMENT, locale: 'en'},
  "Bạn không có quyền truy cập tài liệu yêu cầu.": { ref: NORMALIZED_ERRORS.NO_PERMISSION_TO_ACCESS_THE_REQUESTED_DOCUMENT, locale: 'vi'},
  "No dispones del permiso necesario para acceder al documento solicitado.": { ref: NORMALIZED_ERRORS.NO_PERMISSION_TO_ACCESS_THE_REQUESTED_DOCUMENT, locale: 'es'},
  "Vous n'avez pas l'autorisation d'accéder au document demandé.": { ref: NORMALIZED_ERRORS.NO_PERMISSION_TO_ACCESS_THE_REQUESTED_DOCUMENT, locale: 'fr'},
  "Non disponi dell'autorizzazione necessaria per accedere al documento richiesto.": { ref: NORMALIZED_ERRORS.NO_PERMISSION_TO_ACCESS_THE_REQUESTED_DOCUMENT, locale: 'it'},
  "No cuenta con los permisos necesarios para acceder al documento solicitado.": { ref: NORMALIZED_ERRORS.NO_PERMISSION_TO_ACCESS_THE_REQUESTED_DOCUMENT, locale: 'es_CO'},

  // "Limit Exceeded: DriveApp."
  "Limit Exceeded: DriveApp.": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED_DRIVEAPP, locale: 'en'},
  "Límite Excedido: DriveApp.": { ref: NORMALIZED_ERRORS.LIMIT_EXCEEDED_DRIVEAPP, locale: 'es_419'},
  
  // "Unable to talk to trigger service"
  "Unable to talk to trigger service": { ref: NORMALIZED_ERRORS.UNABLE_TO_TALK_TO_TRIGGER_SERVICE, locale: 'en'},
  "Impossible de communiquer pour déclencher le service": { ref: NORMALIZED_ERRORS.UNABLE_TO_TALK_TO_TRIGGER_SERVICE, locale: 'fr'},
  "Không thể trao đổi với người môi giới để kích hoạt dịch vụ": { ref: NORMALIZED_ERRORS.UNABLE_TO_TALK_TO_TRIGGER_SERVICE, locale: 'vi'},
  "No es posible ponerse en contacto con el servicio de activación.": { ref: NORMALIZED_ERRORS.UNABLE_TO_TALK_TO_TRIGGER_SERVICE, locale: 'es'},
  "無法與觸發服務聯絡": { ref: NORMALIZED_ERRORS.UNABLE_TO_TALK_TO_TRIGGER_SERVICE, locale: 'zh_TW'},

  // "Script has attempted to perform an action that is not allowed when invoked through the Google Apps Script Execution API."
  "Script has attempted to perform an action that is not allowed when invoked through the Google Apps Script Execution API.": { ref: NORMALIZED_ERRORS.ACTION_NOT_ALLOWED_THROUGH_EXEC_API, locale: 'en'},

  // "Mail service not enabled"
  "Mail service not enabled": { ref: NORMALIZED_ERRORS.MAIL_SERVICE_NOT_ENABLED, locale: 'en'},
  "Gmail operation not allowed. : Mail service not enabled": { ref: NORMALIZED_ERRORS.MAIL_SERVICE_NOT_ENABLED, locale: 'en'},

  // This error happens because the Gmail advanced service was not properly loaded during this Apps Script process execution
  // In this case, we need to start a new process execution, ie restart exec from client side - no need to retry multiple times
  "\"Gmail\" is not defined.": { ref: NORMALIZED_ERRORS.GMAIL_NOT_DEFINED, locale: 'en'},

  // "Gmail operation not allowed." - eg: Gmail App.sendEmail()
  "Gmail operation not allowed.": { ref: NORMALIZED_ERRORS.GMAIL_OPERATION_NOT_ALLOWED, locale: 'en'},
  "Gmail operation not allowed. ": { ref: NORMALIZED_ERRORS.GMAIL_OPERATION_NOT_ALLOWED, locale: 'en'},

  // "Invalid thread_id value"
  "Invalid thread_id value": { ref: NORMALIZED_ERRORS.INVALID_THREAD_ID_VALUE, locale: 'en'},

  // "labelId not found"
  "labelId not found": { ref: NORMALIZED_ERRORS.LABEL_ID_NOT_FOUND, locale: 'en'},

  // "Label name exists or conflicts"
  "Label name exists or conflicts": { ref: NORMALIZED_ERRORS.LABEL_NAME_EXISTS_OR_CONFLICTS, locale: 'en'},

  // "Invalid label name"
  "Invalid label name": { ref: NORMALIZED_ERRORS.INVALID_LABEL_NAME, locale: 'en'},

  // "Invalid to header" - eg: Gmail.Users.Messages.send
  "Invalid to header": { ref: NORMALIZED_ERRORS.INVALID_EMAIL, locale: 'en'},
  // "Invalid cc header" - eg: Gmail.Users.Messages.send
  "Invalid cc header": { ref: NORMALIZED_ERRORS.INVALID_EMAIL, locale: 'en'},

  // "Failed to send email: no recipient" - eg: Gmail App.sendEmail()
  "Failed to send email: no recipient": { ref: NORMALIZED_ERRORS.NO_RECIPIENT, locale: 'en'},
  // "Recipient address required" - eg: Gmail.Users.Messages.send()
  "Recipient address required": { ref: NORMALIZED_ERRORS.NO_RECIPIENT, locale: 'en'},

  // "IMAP features disabled by administrator"
  "IMAP features disabled by administrator": { ref: NORMALIZED_ERRORS.IMAP_FEATURES_DISABLED_BY_ADMIN, locale: 'en'},

  // "There are too many LockService operations against the same script." - eg: Lock.tryLock()
  "There are too many LockService operations against the same script.": { ref: NORMALIZED_ERRORS.TOO_MANY_LOCK_OPERATIONS, locale: 'en'},
  "Có quá nhiều thao tác LockService trên cùng một tập lệnh.": { ref: NORMALIZED_ERRORS.TOO_MANY_LOCK_OPERATIONS, locale: 'vi'},

  // "The Google Calendar is not enabled for the user." - eg: CalendarApp.getDefaultCalendar()
  "The Google Calendar is not enabled for the user.": { ref: NORMALIZED_ERRORS.CALENDAR_SERVICE_NOT_ENABLED, locale: 'en'},

  // If you enabled this API recently, wait a few minutes for the action to propagate to our systems and retry.
  "If you enabled this API recently, wait a few minutes for the action to propagate to our systems and retry.": { ref: NORMALIZED_ERRORS.API_NOT_ENABLED, locale: 'en'}
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
  {regex: /^Email không hợp lệ: (.*)$/,
    variables: ['email'],
    ref: NORMALIZED_ERRORS.INVALID_EMAIL,
    locale: 'vi'},
  {regex: /^Ongeldige e-mail: (.*)$/,
    variables: ['email'],
    ref: NORMALIZED_ERRORS.INVALID_EMAIL,
    locale: 'nl'},

  // Document XXX is missing (perhaps it was deleted?) - eg: Spreadsheet App.openById()
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
  {regex: /^Falta el documento (\S*) \(puede que se haya eliminado\)\.$/,
    variables: ['docId'],
    ref: NORMALIZED_ERRORS.DOCUMENT_MISSING,
    locale: 'es'},
  {regex: /^找不到文件「([^」]*)」\(可能已遭刪除\)$/,
    variables: ['docId'],
    ref: NORMALIZED_ERRORS.DOCUMENT_MISSING,
    locale: 'zh_TW'},
  {regex: /^Документ (\S*) отсутствует \(возможно, он был удален\)$/,
    variables: ['docId'],
    ref: NORMALIZED_ERRORS.DOCUMENT_MISSING,
    locale: 'ru'},
  {regex: /^Le document (\S*) est manquant \(peut-être a-t-il été supprimé \?\)$/,
    variables: ['docId'],
    ref: NORMALIZED_ERRORS.DOCUMENT_MISSING,
    locale: 'fr'},
  {regex: /^Dokumen (\S*) hilang \(mungkin dihapus\?\)$/,
    variables: ['docId'],
    ref: NORMALIZED_ERRORS.DOCUMENT_MISSING,
    locale: 'in'},
  {regex: /^Das Dokument (\S*) fehlt\. \(Vielleicht wurde es gelöscht\?\)$/,
    variables: ['docId'],
    ref: NORMALIZED_ERRORS.DOCUMENT_MISSING,
    locale: 'de'},
  {regex: /^O documento (\S*) está ausente \(será que foi excluído\?\)$/,
    variables: ['docId'],
    ref: NORMALIZED_ERRORS.DOCUMENT_MISSING,
    locale: 'pt'},
  {regex: /^Chybí dokument (\S*) \(je možné, že byl smazán\)\.$/,
    variables: ['docId'],
    ref: NORMALIZED_ERRORS.DOCUMENT_MISSING,
    locale: 'cs'},
  {regex: /^Nema dokumenta (\S*) \(možda je izbrisan\?\)$/,
    variables: ['docId'],
    ref: NORMALIZED_ERRORS.DOCUMENT_MISSING,
    locale: 'hr'},
  {regex: /^A\(z\) (\S*) dokumentum hiányzik \(talán törölték\?\)$/,
    variables: ['docId'],
    ref: NORMALIZED_ERRORS.DOCUMENT_MISSING,
    locale: 'hu'},

  // User-rate limit exceeded. Retry after XXX - this error can be prefixed with a translated version of 'Limit Exceeded'
  {regex: /User-rate limit exceeded\.\s+Retry after (.*Z)/,
    variables: ['timestamp'],
    ref: NORMALIZED_ERRORS.USER_RATE_LIMIT_EXCEEDED_RETRY_AFTER_SPECIFIED_TIME,
    locale: 'en'},

  // User Rate Limit Exceeded. Rate of requests for user exceed configured project quota.
  // You may consider re-evaluating expected per-user traffic to the API and adjust project quota limits accordingly.
  // You may monitor aggregate quota usage and adjust limits in the API Console: https://console.developers.google.com/XXX
  {regex: /User Rate Limit Exceeded\. Rate of requests for user exceed configured project quota\./,
    ref: NORMALIZED_ERRORS.USER_RATE_LIMIT_EXCEEDED,
    locale: 'en'},
  
  // Daily Limit Exceeded. The quota will be reset at midnight Pacific Time (PT). 
  // You may monitor your quota usage and adjust limits in the API Console: https://console.developers.google.com/XXX
  {regex: /Daily Limit Exceeded\. The quota will be reset at midnight Pacific Time/,
    ref: NORMALIZED_ERRORS.DAILY_LIMIT_EXCEEDED,
    locale: 'en'},

  // Service invoked too many times for one day: XXX. (XXX: urlFetch, email)
  {regex: /^Service invoked too many times for one day: ([^.]*)\.$/,
    variables: ['service'],
    ref: NORMALIZED_ERRORS.SERVICE_INVOKED_TOO_MANY_TIMES_FOR_ONE_DAY,
    locale: 'en'},
  {regex: /^Trop d'appels pour ce service aujourd'hui : ([^.]*)\.$/,
    variables: ['service'],
    ref: NORMALIZED_ERRORS.SERVICE_INVOKED_TOO_MANY_TIMES_FOR_ONE_DAY,
    locale: 'fr'},
  {regex: /^Servicio solicitado demasiadas veces en un mismo día: ([^.]*)\.$/,
    variables: ['service'],
    ref: NORMALIZED_ERRORS.SERVICE_INVOKED_TOO_MANY_TIMES_FOR_ONE_DAY,
    locale: 'es'},
  {regex: /^Serviço chamado muitas vezes no mesmo dia: ([^.]*)\.$/,
    variables: ['service'],
    ref: NORMALIZED_ERRORS.SERVICE_INVOKED_TOO_MANY_TIMES_FOR_ONE_DAY,
    locale: 'pt'},
  {regex: /^Dịch vụ bị gọi quá nhiều lần trong một ngày: ([^.]*)\.$/,
    variables: ['service'],
    ref: NORMALIZED_ERRORS.SERVICE_INVOKED_TOO_MANY_TIMES_FOR_ONE_DAY,
    locale: 'vi'},

  // Service unavailable: XXX (XXX: Docs)
  {regex: /^Service unavailable: (.*)$/,
    variables: ['service'],
    ref: NORMALIZED_ERRORS.SERVICE_UNAVAILABLE,
    locale: 'en'},

  // Service error: XXX (XXX: Spreadsheets)
  {regex: /^Service error: (.*)$/,
    variables: ['service'],
    ref: NORMALIZED_ERRORS.SERVICE_ERROR,
    locale: 'en'},
  {regex: /^Erro de serviço: (.*)$/,
    variables: ['service'],
    ref: NORMALIZED_ERRORS.SERVICE_ERROR,
    locale: 'pt'},

  // "Invalid argument: XXX" - wrong email alias used - eg: Gmail App.sendEmail()
  {regex: /^Invalid argument: (.*)$/,
    variables: ['email'],
    ref: NORMALIZED_ERRORS.INVALID_ARGUMENT,
    locale: 'en'},

  // "A sheet with the name "XXX" already exists. Please enter another name." - eg: [Sheet].setName()
  {regex: /^A sheet with the name "([^"]*)" already exists\. Please enter another name\.$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'en'},
  {regex: /^A sheet with the name ‘([^’]*)’ already exists\. Please enter another name\.$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'en_NZ'},
  {regex: /^Ya existe una hoja con el nombre "([^"]*)"\. Ingresa otro\.$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'es_419'},
  {regex: /^Ya existe una hoja con el nombre "([^"]*)"\. Introduce un nombre distinto\.$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'es'},
  {regex: /^Đã tồn tại một trang tính có tên "([^"]*)"\. Vui lòng nhập tên khác\.$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'vi'},
  {regex: /^Une feuille nommée "([^"]*)" existe déjà\. Veuillez saisir un autre nom\.$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'fr'},
  {regex: /^Une feuille nommée « ([^»]*)» existe déjà\. Veuillez saisir un autre nom\.$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'fr_CA'},
  {regex: /^Esiste già un foglio con il nome "([^"]*)"\. Inserisci un altro nome\.$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'it'},
  {regex: /^已有工作表使用「([^」]*)」這個名稱。請輸入其他名稱。$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'zh_TW'},
  {regex: /^Es ist bereits ein Tabellenblatt mit dem Namen "([^"]*)" vorhanden\. Geben Sie einen anderen Namen ein\.$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'de'},
  {regex: /^이름이 ‘([^’]*)’인 시트가 이미 있습니다\. 다른 이름을 입력해 주세요\.$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'ko'},
  {regex: /^シート名「([^」]*)」はすでに存在しています。別の名前を入力してください。$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'ja'},
  {regex: /^Er is al een blad met de naam ([^.]*)\. Geef een andere naam op\.$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'nl'},
  {regex: /^Já existe uma página chamada "([^"]*)"\. Insira outro nome\.$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'pt'},
  {regex: /^Лист "([^"]*)" уже существует\. Введите другое название\.$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'ru'},
  {regex: /^Hárok s názvom ([^ž]*) už existuje\. Zadajte iný názov\.$/,
    variables: ['sheetName'],
    ref: NORMALIZED_ERRORS.SHEET_ALREADY_EXISTS_PLEASE_ENTER_ANOTHER_NAME,
    locale: 'sk'},
  {regex: /^Project ([0-9]+) is not found and cannot be used for API calls\. If it is recently created, enable (.*) by visiting (.*) then retry\. If you enabled this API recently, wait a few minutes for the action to propagate to our systems and retry\.$/,
    variables: ['projectId', 'apiName', 'consoleUrl'],
    ref: NORMALIZED_ERRORS.API_NOT_ENABLED,
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
  
