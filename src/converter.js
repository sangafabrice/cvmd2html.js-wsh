/**
 * @file returns the method to convert from markdown to html.
 * @version 0.0.1.3
 */

/**
 * Represents the markdown to html converter.
 * @typedef {object} MarkdownToHtml
 */

/** @module converter */
(function() {
  /** @type {MessageBox} */
  var MessageBox = include('src/msgbox.js');

  // Append 1 to the left of the binary representation of Exception HResult to force
  // 64-bit computation. The value of the HResults here should be read as a 64-bit 
  // negative integer instead of 32-bit unsigned in JScript. This is for clarity.
  /** @constant {int64} */
  var PERMISSION_DENIED = 0x800A0046 | 0x100000000;
  /** @constant {int64} */
  var FILE_NOT_FOUND = 0x800A0035 | 0x100000000;

  /** @constant {regexp} */
  var MARKDOWN_REGEX = /\.md$/i;

  /**
   * @param {string} htmlLibraryPath is the path string of the html loading the library.
   * @param {string} jsLibraryPath is the javascript library path.
   * @returns {object} the Converter type.
   */
  return function (htmlLibraryPath, jsLibraryPath) {
    /** @class @constructs MarkdownToHtml */
    function MarkdownToHtml() { }

    /**
     * Show a warning message or an error message box.
     * The function does not return anything when the message box is an error.
     * @public @static @memberof MessageBox
     * @param {string} message is the message text.
     * @param {number} [messageType = ERROR_MESSAGE] message box type (Warning/Error).
     * @returns {string|void} "Yes" or "No" depending on the user's click when the message box is a warning.
     */
    MarkdownToHtml.ConvertFrom = function (markdownPath) {
      // Validate the input markdown path string.
      if (!MARKDOWN_REGEX.test(markdownPath)) {
        MessageBox.Show(format('"{0}" is not a markdown (.md) file.', markdownPath));
      }
      SetHtmlContent(GetHtmlPath(markdownPath), ConvertToHtml(GetContent(markdownPath)));
    }

    /**
     * Convert a markdown content to an html document.
     * @param {string} mardownContent is the content to convert.
     * @returns {string} the output html document content. 
     */
    function ConvertToHtml(markdownContent) {
      // Build the HTML document that will load the showdown.js library.
      var document = new ActiveXObject('htmlFile');
      document.open();
      document.write(format(GetContent(htmlLibraryPath), GetContent(jsLibraryPath)));
      document.close();
      return document.parentWindow.convertMarkdown(markdownContent);
    }

    /**
     * This function returns the output path when it is unique without prompts or when
     * the user accepts to overwrite an existing HTML file. Otherwise, it exits the script.
     * @returns {string} the output html path.
     */
    function GetHtmlPath(markdownPath) {
      var htmlPath = markdownPath.replace(MARKDOWN_REGEX, '.html');
      if (fs.FileExists(htmlPath)) {
        MessageBox.Show(format('The file "{0}" already exists.\n\nDo you want to overwrite it?', htmlPath), MessageBox.WARNING);
      } else if (fs.FolderExists(htmlPath)) {
        MessageBox.Show(format('"{0}" cannot be overwritten because it is a directory.', htmlPath));
      }
      return htmlPath;
    }

    /**
     * Get the content of a file.
     * @param {string} filePath is path that is read.
     * @returns {string} the content of the file.
     */
    function GetContent(filePath) {
      var FOR_READING = 1;
      try {
        with (fs.OpenTextFile(filePath, FOR_READING)) {
          var content = ReadAll();
          Close();
        }
        return content;
      } catch (error) {
        switch (error.number) {
          case PERMISSION_DENIED:
            if (!fs.FolderExists(filePath)) {
              MessageBox.Show(format('Access to the path "{0}" is denied.', filePath));
            }
          case FILE_NOT_FOUND:
            MessageBox.Show(format('File "{0}" is not found.', filePath));
          default:
            MessageBox.Show(format('Unspecified error trying to read from "{0}".', filePath));
        }
      } 
    }

    /**
     * Write the html text to the output HTML file.
     * It notifies the user when the operation did not complete with success.
     * @param {string} htmlPath is the output html path.
     * @param {string} htmlContent is the content of the html file.
     */
    function SetHtmlContent(htmlPath, htmlContent) {
      var FOR_WRITING = 2;
      try {
        var txtStream = fs.OpenTextFile(htmlPath, FOR_WRITING, true);
        txtStream.Write(htmlContent);
      } catch (error) {
        if (error.number == PERMISSION_DENIED) {
          MessageBox.Show(format('Access to the path "{0}" is denied.', htmlPath));
        } else {
          MessageBox.Show(format('Unspecified error trying to write to "{0}".', htmlPath));
        }
      } finally {
        if (txtStream != undefined) {
          txtStream.Close();
        }
      }
    }

    return MarkdownToHtml;
  }
})();