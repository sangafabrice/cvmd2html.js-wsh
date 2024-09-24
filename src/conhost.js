/**
 * @file returns the shortcut target script runner watcher.
 * @version 0.0.1
 */

/**
 * Represents the shortcut target script runner host.
 * @typedef {object} ConsoleHost
 */

/** @module conhost */
(function() {
  /** @type {MessageBox} */
  var MessageBox = include('src/msgbox.js');

  /**
   * @param {string} pwshExePath is the powershell core path.
   * @param {string} pwshScriptPath is the target powershell script path.
   * @returns {object} the ConsoleHost type.
   */
  return function (pwshExePath, pwshScriptPath) {
    /** @class @constructs ConsoleHost */
    function ConsoleHost() { }

    /**
     * Execute the runner of the shortcut target script and wait for its exit.
     * @public @static @memberof ConsoleHost
     * @param {string} markdownPath is the input markdown file path.
     */
    ConsoleHost.StartWith = function(markdownPath) {
      WaitForExit(wshell.Exec(format('"{0}" -nop -ep Bypass -w Hidden -cwa "try { & $args[0] -MarkdownPath $args[1] } catch { Write-Error $_.Exception.Message }" "{1}" "{2}"', pwshExePath, pwshScriptPath, markdownPath)));
    }

    /**
     * Represents the data obtained from the console host.
     * @private @class @constructs ConsoleData
     */
    function ConsoleData() {
      /** @private @type {string} is the expected prompt from the console host. */
      this.OverwritePromptText = '';
    }

    /**
     * Show the overwrite prompt that the child process sends. Handle the event when the
     * PowerShell Core (child) process redirects output to the parent Standard Output stream.
     * @public
     * @param {object} pwshExe it the sender child process.
     * @param {string} outData the output text line sent.
     */
    ConsoleData.prototype.HandleOutputDataReceived = function (pwshExe, outData) {
      if (outData.length) {
        // Show the message box when the text line is a question.
        // Otherwise, append the text line to the overall message text variable.
        if (outData.match(/\?\s*$/)) {
          this.OverwritePromptText += '\n' + outData;
          // Write the user's choice to the child process console host.
          pwshExe.StdIn.WriteLine(MessageBox.Show(this.OverwritePromptText, MessageBox.WARNING));
          this.OverwritePromptText = '';
        } else {
          this.OverwritePromptText += outData + '\n';
        }
      }
    }

    /**
     * Observe when the child process exits with or without an error.
     * Call the appropriate handler for each outcome.
     * @private
     * @param {object} pwshExe is the PowerShell Core process or child process.
     */
    function WaitForExit(pwshExe) {
      var conhostData = new ConsoleData();
      // Wait for the process to complete.
      while (!pwshExe.Status && !pwshExe.ExitCode) {
        conhostData.HandleOutputDataReceived(pwshExe, pwshExe.StdOut.ReadLine());
      }
      // When the process terminated with an error.
      if (pwshExe.ExitCode) {
        HandleErrorDataReceived(pwshExe.StdErr.ReadAll());
      }
    }

    /**
     * Show the error message that the child process writes on the console host.
     * @private
     * @param {string} errData the error message text.
     */
    function HandleErrorDataReceived(errData) {
      if (errData.length) {
        // Remove the ANSI color tag characters from the error message data text.
        errData = errData.replace(/(\x1B\[31;1m)|(\x1B\[0m)/g, '');
        MessageBox.Show(errData.substring(errData.indexOf(':') + 2));
      }
    }

    return ConsoleHost;
  }
})();