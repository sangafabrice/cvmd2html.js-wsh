/**
 * @file Launches the shortcut target PowerShell script with the selected markdown as an argument.
 * It aims to eliminate the flashing console window when the user clicks on the shortcut menu.
 * @version 0.0.1.9
 */

// Imports.
var fs = new ActiveXObject('Scripting.FileSystemObject');
var wshell = new ActiveXObject('WScript.Shell');

/** @type {ParamHash} */
var param = include('src/parameters.js');
/** @type {PackageHash} */
var package = include('src/package.js');

/** The application execution. */
if (param.RunLink) {
  var WINDOW_STYLE_HIDDEN = 0;
  var WAIT_ON_RETURN = true;
  package.IconLink.Create(param.Markdown);
  wshell.Run(format('"{0}"', package.IconLink.Path), WINDOW_STYLE_HIDDEN, WAIT_ON_RETURN)
  package.IconLink.Delete();
  WSH.Quit();
}

if (param.Markdown) {
  include('src/converter.js')(package.JsLibraryPath).ConvertFrom(param.Markdown);
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
 * Get the WSH runtime in GUI mode (wscript.exe).
 * @returns {string} the WScript.Exe path.
 */
function getDefaultCustomIconLinkTarget() {
  return fs.BuildPath(fs.GetParentFolderName(WSH.FullName), 'wscript.exe');
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