/**
 * @file manages the error log file and content.
 * @version 0.0.1
 */

/**
 * @typedef {object} ErrorLogHash
 * @property {string} Path is the randomly generated error log file path.
 */

/** @module errorLog */
(function() {
  var fs = new ActiveXObject('Scripting.FileSystemObject');
  var wshell = new ActiveXObject('WScript.Shell');
  /** @type {ErrorLogHash} */
  var errorLog = {
    Path: fs.BuildPath(wshell.ExpandEnvironmentStrings('%TEMP%'), WSH.CreateObject('Scriptlet.TypeLib').Guid.substr(1, 36).toLowerCase() + '.tmp.log'),
    /**
     * Display the content of the error log file in a message box if it is not empty.
     */
    Read: function () {
      try {
        var FOR_READING = 1;
        var txtStream = fs.OpenTextFile(this.Path, FOR_READING);
        // Read the error message and remove the ANSI escaped character for red coloring.
        var errorMessage = txtStream.ReadAll().replace(/(\x1B\[31;1m)|(\x1B\[0m)/g, '');
        if (errorMessage.length) {
          var OKONLY_BUTTON = 0;
          var ERROR_ICON = 16;
          var NO_TIMEOUT = 0;
          wshell.Popup(errorMessage, NO_TIMEOUT, 'Convert to HTML', OKONLY_BUTTON + ERROR_ICON);
        }
      } catch (error) { }
      if (txtStream) {
        txtStream.Close();
      }
    },
    /**
     * Delete the error log file.
     */
    Delete: function () {
      try {
        fs.DeleteFile(this.Path);
      } catch (error) { }
    }
  }
  return errorLog;
})();