/**
 * @file Launches the shortcut target PowerShell script with the selected markdown as an argument.
 * It aims to eliminate the flashing console window when the user clicks on the shortcut menu.
 * @version 0.0.1.3
 */

RequestAdminPrivileges();

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
  if (waitForExit(processService.ExecMethod_(createMethod.Name, inParam).ProcessId)) {
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
 * Wait for the process exit.
 * @param {number} processId is the process identifier.
 * @return {number} the process exit code.
 */
function waitForExit(processId) {
  // The process termination event query.
  var wmiQuery = 'SELECT * FROM Win32_ProcessStopTrace WHERE ProcessName="cmd.exe" AND ProcessId=' + processId;
  // Wait for the process to exit.
  return GetObject('winmgmts:').ExecNotificationQuery(wmiQuery).NextEvent().ExitStatus;
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

/**
 * Request administrator privileges if standard user.
 */
function RequestAdminPrivileges() {
  var registry = GetObject('winmgmts:StdRegProv');
  var checkAccessMethod = registry.Methods_('CheckAccess');
  var inParam = checkAccessMethod.InParameters.SpawnInstance_();
  with (inParam) {
    hDefKey = 0x80000003; // HKU
    sSubKeyName = 'S-1-5-19\\Environment';
  }
  if (registry.ExecMethod_(checkAccessMethod.Name, inParam).bGranted) {
    return;
  }
  var inputCommand = format('"{0}"', WSH.ScriptFullName);
  for (var index = 0; index < WSH.Arguments.Count(); index++) {
    inputCommand += format(' "{0}"', WSH.Arguments(index));
  }
  var WINDOW_STYLE_HIDDEN = 0;
  WSH.CreateObject('Shell.Application').ShellExecute(WSH.FullName, inputCommand, null, 'runas', WINDOW_STYLE_HIDDEN);
  WSH.Quit();
}