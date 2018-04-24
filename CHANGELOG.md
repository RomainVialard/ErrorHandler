# Change Log

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## Unreleased
<!-- Add new, unreleased changes here. -->
* none

## [2.0.0] - 2018-04-24
* Breaking changes: expBackoff(): use options to specify its behavior (throw on fail, verbose, retryNumber, doNotLogKnownErrors)
* Leverage Normalized GAS error in logError()
* Leverage Normalized GAS error in expBackoff()
* new function: getNormalizedError(): Return the english version of the error if listed in this library
* new function: getErrorLocale(): Try to find the locale of the localized thrown error
* Normalize GAS error messages across language and variable messages (eg: when containing a document ID or a variable part)
* List known GAS error message

## [1.1.2] - 2018-04-05
* Fix issue for global Add-on name reference error

## [1.1.1] - 2018-04-05
* Fix issue where response is undefined

## [1.1.0] - 2018-04-04
* Add Exponential backoff for UrlFetch call
* Add urlFetchWithExpBackOff() method

## [1.0.1] - 2018-03-30
* Fix destructuring not accepted by GAS

## [1.0.0] - 2018-03-30
* Initial npm version
