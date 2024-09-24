/**
 * @file returns the markdown conversion message box.
 * @version 0.0.1.1
 */

/**
 * Represents the markdown conversion message box.
 * @typedef {object} MessageBox
 * @property {number} WARNING specifies that the dialog shows a warning message.
 */

/** @module msgbox */
(function() {
  /** @private @constant {string} */
  var MESSAGE_BOX_TITLE = 'Convert to HTML';
  /** @private @constant {number} */
  var NO_MESSAGE_TIMEOUT = 0;
  /** @private @constant {number} */
  var YESNO_BUTTON = 4;
  /** @private @constant {number} */
  var OK_BUTTON = 0;
  /** @private @constant {number} */
  var OK_POPUPRESULT = 1;
  /** @private @constant {number} */
  var NO_POPUPRESULT = 7;
  /** @private @constant {number} */
  var ERROR_MESSAGE = 16;
  /** @private @constant {number} */
  var WARNING_MESSAGE = 48;

  /** @typedef MessageBox */
  var MessageBox = { };

  /** @public @static @readonly @property {number} */
  MessageBox.WARNING = WARNING_MESSAGE;
  // Object.defineProperty() method does not work in WSH.
  // It is not possible in this implementation to make the
  // property non-writable.

  /**
   * Show a warning message or an error message box.
   * The function does not return anything when the message box is an error.
   * @public @static @memberof MessageBox
   * @param {string} message is the message text.
   * @param {number} [messageType = ERROR_MESSAGE] message box type (Warning/Error).
   * @returns {string|void} "Yes" or "No" depending on the user's click when the message box is a warning.
   */
  MessageBox.Show = function(message, messageType) {
    if (messageType != ERROR_MESSAGE && messageType != WARNING_MESSAGE) {
      messageType = ERROR_MESSAGE;
    }
    // The error message box shows the OK button alone.
    // The warning message box shows the alternative Yes or No buttons.
    messageType += messageType == ERROR_MESSAGE ? OK_BUTTON:YESNO_BUTTON;
    switch (wshell.Popup(message, NO_MESSAGE_TIMEOUT, MESSAGE_BOX_TITLE, messageType)) {
      case OK_POPUPRESULT:
      case NO_POPUPRESULT:
        WSH.Quit();
    }
  }

  return MessageBox;
})();