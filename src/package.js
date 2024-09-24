/**
 * @file returns information about the resource files used by the project.
 * It also provides a way to manage the custom icon link that can be installed and uninstalled.
 * @version 0.0.1
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
  var fs = new ActiveXObject('Scripting.FileSystemObject');
  var wshell = new ActiveXObject('WScript.Shell');
  /** @type {PackageHash} */
  var package = {
    Root: fs.GetParentFolderName(WSH.ScriptFullName)
  };
  package.ResourcePath = fs.BuildPath(package.Root, 'rsc');
  package.PwshScriptPath = fs.BuildPath(package.ResourcePath, 'cvmd2html.ps1');
  /**
   * Get the partial "arguments" property string of the custom icon link.
   * The command is partial because it does not include the markdown file path string.
   * The markdown file path string will be input when calling the shortcut link.
   */
  var customIconLinkArguments = format('-ep Bypass -nop -w Hidden -f "{0}" -Markdown', package.PwshScriptPath);
  package.MenuIconPath = fs.BuildPath(package.ResourcePath, 'menu.ico');
  // The HKLM registry subkey stores the PowerShell Core application path.
  package.PwshExePath = wshell.RegRead('HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe\\');
  package.IconLink = {
    DirName: wshell.SpecialFolders('StartMenu'),
    Name: 'cvmd2html.lnk',
    /**
     * Create the custom icon link file.
     * @method @memberof package.IconLink
     */
    Create: function () {
      var link = this.GetLink();
      link.TargetPath = package.PwshExePath;
      link.Arguments = customIconLinkArguments;
      link.IconLocation = package.MenuIconPath;
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
     * Validate the link properties.
     * @method @memberof package.IconLink
     * @returns {boolean} true if the link properties are as expected, false otherwise. 
     */
    IsValid: function () {
      var linkItem = this.GetLink();
      var targetCommand = '{0} {1}';
      return format(targetCommand, linkItem.TargetPath, linkItem.Arguments).toLowerCase() == format(targetCommand, package.PwshExePath, customIconLinkArguments).toLowerCase();
    },
    /**
     * Get the link.
     * @returns {object} the link object.
     */
    GetLink: function () {
      return wshell.CreateShortcut(this.Path);
    }
  }
  package.IconLink.Path = fs.BuildPath(package.IconLink.DirName, package.IconLink.Name);
  return package;
})();