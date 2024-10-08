/**
 * @file returns information about the resource files used by the project.
 * It also provides a way to manage the custom icon link that can be installed and uninstalled.
 * @version 0.0.1.12
 */

/**
 * @typedef {object} PackageHash
 * @property {string} Root is the project root path.
 * @property {string} ResourcePath is the project resources directory path.
 * @property {string} MenuIconPath is the shortcut menu icon path.
 * @property {string} HtmlLibraryPath is the html file that embeds the JS library path.
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
  package.HtmlLibraryPath = fs.BuildPath(package.ResourcePath, 'showdown.html');
  package.MenuIconPath = fs.BuildPath(package.ResourcePath, 'menu.ico');
  package.IconLink = {
    DirName: wshell.ExpandEnvironmentStrings('%TEMP%'),
    Name: WSH.CreateObject('Scriptlet.TypeLib').Guid.substr(1, 36).toLowerCase() + '.tmp.lnk',
    /**
     * Create the custom icon link file.
     * @method @memberof package.IconLink
     * @param {string} markdownPath is the input markdown file path.
     */
    Create: function (markdownPath) {
      var link = this.GetLink();
      link.TargetPath =  getDefaultCustomIconLinkTarget();
      link.Arguments = format('"{0}" /Markdown:"{1}"', WSH.ScriptFullName, markdownPath);
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