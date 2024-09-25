/**
 * @file returns the methods for managing the shortcut menu option: install and uninstall.
 * @version 0.0.1.1
 */

/** @module setup */
(function() {
  var HKCU = 0x80000001;
  var VERB_KEY = 'SOFTWARE\\Classes\\SystemFileAssociations\\.md\\shell\\cthtml';
  var registry = GetObject('winmgmts:StdRegProv');
  var setup = {
    /**
     * Configure the shortcut menu in the registry.
     * @param {boolean} paramNoIcon specifies that the menu icon should not be set.
     * @param {string} menuIconPath is the shortcut menu icon file path.
     */
    Set: function (paramNoIcon, menuIconPath) {
      var COMMAND_KEY = VERB_KEY + '\\command';
      var command = format('{0} "{1}" /Markdown:"%1"', WScript.FullName.replace(/\\cscript\.exe$/i, '\\wscript.exe'), WScript.ScriptFullName);
      registry.CreateKey(HKCU, COMMAND_KEY);
      registry.SetStringValue(HKCU, COMMAND_KEY, null, command);
      registry.SetStringValue(HKCU, VERB_KEY, null, 'Convert to &HTML');
      var iconValueName = 'Icon';
      if (paramNoIcon) {
        registry.DeleteValue(HKCU, VERB_KEY, iconValueName);
      } else {
        registry.SetStringValue(HKCU, VERB_KEY, iconValueName, menuIconPath);
      }
    },
    /** Remove the shortcut menu by removing the verb key and subkeys. */
    Unset: function () {
      var enumKeyMethod = registry.Methods_('EnumKey');
      var inParam = enumKeyMethod.InParameters.SpawnInstance_();
      inParam.hDefKey = HKCU;
      // Recursion is used because a key with subkeys cannot be deleted.
      // Recursion helps removing the leaf keys first.
      (function(key) {
        inParam.sSubKeyName = key;
        var sNames = registry.ExecMethod_(enumKeyMethod.Name, inParam).sNames;
        if (sNames != null) {
          var sNamesArray = sNames.toArray();
          for (var index = 0; index < sNamesArray.length; index++) {
            arguments.callee(format('{0}\\{1}', key, sNamesArray[index]));
          }
        }
        registry.DeleteKey(HKCU, key);
      })(VERB_KEY);
    }
  }
  return setup;
})();