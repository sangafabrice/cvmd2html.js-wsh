/**
 * @file returns information about the resource files used by the project.
 * It also provides a way to manage the custom icon link that can be installed and uninstalled.
 * @version 0.0.1.3
 */

/**
 * @typedef {object} PackageHash
 * @property {string} Root is the project root path.
 * @property {string} ResourcePath is the project resources directory path.
 * @property {string} MenuIconPath is the shortcut menu icon path.
 * @property {string} PwshExePath is the powershell core runtime path.
 * @property {string} PwshScriptPath is the shortcut target powershell script path.
 * @property {object} IconLink represents an adapted link object.
 * @property {string} IconLink.DirName is the parent directory path of the custom icon link.
 * @property {string} IconLink.Name is the custom icon link file name.
 * @property {string} IconLink.Path is the custom icon link full path.
 */

/** @module package */
(function() {
  /** @type {PackageHash} */
  var package = {
    Root: fs.GetParentFolderName(WSH.ScriptFullName)
  };
  package.ResourcePath = fs.BuildPath(package.Root, 'rsc');
  package.PwshScriptPath = fs.BuildPath(package.ResourcePath, 'cvmd2html.ps1');
  package.MenuIconPath = fs.BuildPath(package.ResourcePath, 'menu.ico');
  package.PwshExePath = (function() {
    var getStringValueMethod = registry.Methods_('GetStringValue');
    var inParam = getStringValueMethod.InParameters.SpawnInstance_();
    // The HKLM registry subkey stores the PowerShell Core application path.
    inParam.sSubKeyName = 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe';
    return registry.ExecMethod_(getStringValueMethod.Name, inParam).sValue;
  })();
  package.IconLink = {
    DirName: wshell.ExpandEnvironmentStrings('%TEMP%'),
    Name: typeLib.Guid.substr(1, 36).toLowerCase() + '.tmp.lnk',
    /**
     * Create the custom icon link file.
     * @method @memberof package.IconLink
     * @param {string} markdownPath is the input markdown file path.
     */
    Create: function (markdownPath) {
      fs.CreateTextFile(this.Path).Close();
      var link = this.GetLink();
      link.Path = package.PwshExePath;
      link.Arguments = format('-ep Bypass -nop -w Hidden -f "{0}" -Markdown "{1}"', package.PwshScriptPath, markdownPath);
      link.SetIconLocation(package.MenuIconPath, 0);
      link.Save();
    },
    /**
     * Delete the custom icon link file.
     * @method @memberof package.IconLink
     */
    Delete: function () {
      try {
        fs.DeleteFile(this.Path);
      } catch (error) { }
    },
    /**
     * Get the link.
     * @returns {object} the link object.
     */
    GetLink: function () {
      return shell.NameSpace(this.DirName).ParseName(this.Name).GetLink;
    }
  }
  package.IconLink.Path = fs.BuildPath(package.IconLink.DirName, package.IconLink.Name);
  return package;
})();