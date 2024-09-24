/**
 * @file Launches the shortcut target PowerShell script with the selected markdown as an argument.
 * It aims to eliminate the flashing console window when the user clicks on the shortcut menu.
 * @version 0.0.1.3
 */

// Imports.
var fs = new ActiveXObject('Scripting.FileSystemObject');
var wshell = new ActiveXObject('WScript.Shell');
var typeLib = new ActiveXObject('Scriptlet.TypeLib');
var registry = GetObject('winmgmts:StdRegProv');

/** @type {ParamHash} */
var param = include('src/parameters.js');
/** @type {PackageHash} */
var package = include('src/package.js');

/** The application execution. */
if (param.Markdown) {
  var WINDOW_STYLE_HIDDEN = 0;
  var WAIT_ON_RETURN = true;
  /** @type {ErrorLogHash} */
  var errorLog = include('src/errorLog.js');
  package.IconLink.Create(param.Markdown);
  if (wshell.Run(format('C:\\Windows\\System32\\cmd.exe /d /c ""{0}" 2> "{1}""', package.IconLink.Path, errorLog.Path), WINDOW_STYLE_HIDDEN, WAIT_ON_RETURN)) {
    errorLog.Read();
    errorLog.Delete();
  }
  package.IconLink.Delete();
  WSH.Quit();
}

/** Configuration and settings. */
if (param.Set || param.Unset) {
  var setup = include('src/setup.js');
  if (param.Set) {
    setup.Set(param.NoIcon, package.MenuIconPath);
  } else if (param.Unset) {
    setup.Unset();
  }
}

/**
 * Replace the format item "{n}" by the nth input in a list of arguments.
 * @param {string} formatStr the pattern format.
 * @param {...string} args the replacement texts.
 * @returns {string} a copy of format with the format items replaced by args.
 */
function format(formatStr, args) {
  args = Array.prototype.slice.call(arguments).slice(1);
  while (args.length > 0) {
    formatStr = formatStr.replace(new RegExp('\\{' + (args.length - 1) + '\\}', 'g'), args.pop());
  }
  return formatStr;
}

/**
 * Import the specified jscript source file.
 * @param {string} libraryPath is the source file path.
 * @returns {object} the object returned by the library.
 */
function include(libraryPath) {
  var FOR_READING = 1;
  try {
    with(fs.OpenTextFile(fs.BuildPath(fs.GetParentFolderName(WSH.ScriptFullName), libraryPath), FOR_READING)) {
      var content = ReadAll();
      Close();
    }
    return eval(content);
  } catch (error) { }
}