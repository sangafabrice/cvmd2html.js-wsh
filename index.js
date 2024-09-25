/**
 * @file Launches the shortcut target PowerShell script with the selected markdown as an argument.
 * It aims to eliminate the flashing console window when the user clicks on the shortcut menu.
 * @version 0.0.1.1
 */

/** @type {ParamHash} */
var param = include('src/parameters.js');
/** @type {PackageHash} */
var package = include('src/package.js');

/** The application execution. */
if (param.Markdown) {
  if (!package.IconLink.IsValid()) {
    WScript.Quit();
  }
  var WINDOW_STYLE_HIDDEN = 0xC;
  var startInfo = GetObject('winmgmts:Win32_ProcessStartup').SpawnInstance_();
  startInfo.ShowWindow = WINDOW_STYLE_HIDDEN;
  GetObject('winmgmts:Win32_Process').Create(format('C:\\Windows\\System32\\cmd.exe /d /c ""{0}" "{1}""', package.IconLink.Path, param.Markdown), null, startInfo);
  WScript.Quit();
}

/** Configuration and settings. */
if (param.Set || param.Unset) {
  var setup = include('src/setup.js');
  if (param.Set) {
    package.IconLink.Create();
    setup.Set(param.NoIcon, package.MenuIconPath);
  } else if (param.Unset) {
    setup.Unset();
    package.IconLink.Delete();
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
  var fs = new ActiveXObject('Scripting.FileSystemObject');
  var FOR_READING = 1;
  try {
    with(fs.OpenTextFile(fs.BuildPath(fs.GetParentFolderName(WScript.ScriptFullName), libraryPath), FOR_READING)) {
      var content = ReadAll();
      Close();
    }
    return eval(content);
  } catch (error) { }
}