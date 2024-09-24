/**
 * @file Launches the shortcut target PowerShell script with the selected markdown as an argument.
 * It aims to eliminate the flashing console window when the user clicks on the shortcut menu.
 * @version 0.0.1.4
 */

/** @type {ParamHash} */
var param = include('src/parameters.js');
/** @type {PackageHash} */
var package = include('src/package.js');

/** The application execution. */
if (param.Markdown) {
  var WINDOW_STYLE_HIDDEN = 0xC;
  var startInfo = GetObject('winmgmts:Win32_ProcessStartup').SpawnInstance_();
  startInfo.ShowWindow = WINDOW_STYLE_HIDDEN;
  /** @type {ErrorLogHash} */
  var errorLog = include('src/errorLog.js');
  var processService = GetObject('winmgmts:Win32_Process');
  var createMethod = processService.Methods_('Create');
  var inParam = createMethod.InParameters.SpawnInstance_();
  inParam.CommandLine = format('C:\\Windows\\System32\\cmd.exe /d /c ""{0}" 2> "{1}""', package.IconLink.Path, errorLog.Path);
  inParam.ProcessStartupInformation = startInfo;
  package.IconLink.Create(param.Markdown);
  waitForChildExit(processService.ExecMethod_(createMethod.Name, inParam).ProcessId);
  package.IconLink.Delete();
  errorLog.Read();
  errorLog.Delete();
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
 * Wait for the child process exit.
 * @param {number} parentProcessId is the parent process identifier.
 */
function waitForChildExit(parentProcessId) {
  // The process termination event query.
  // Select the process whose parent is the intermediate process used for executing the link.
  var wmiQuery = 'SELECT * FROM __InstanceDeletionEvent WITHIN 0.1 WHERE TargetInstance ISA "Win32_Process" AND TargetInstance.Name="pwsh.exe" AND TargetInstance.ParentProcessId=' + parentProcessId;
  // Wait for the process to exit.
  GetObject('winmgmts:').ExecNotificationQuery(wmiQuery).NextEvent();
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
    with(fs.OpenTextFile(fs.BuildPath(fs.GetParentFolderName(WSH.ScriptFullName), libraryPath), FOR_READING)) {
      var content = ReadAll();
      Close();
    }
    return eval(content);
  } catch (error) { }
}