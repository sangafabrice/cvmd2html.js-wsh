/**
 * @file returns the methods for managing the shortcut menu option: install and uninstall.
 * @version 0.0.1.1
 */

/** @module setup */
(function() {
  var VERB_KEY = 'SOFTWARE\\Classes\\SystemFileAssociations\\.md\\shell\\cthtml';
  var KEY_FORMAT = 'HKCU\\{0}\\';
  var setup = {
    /** Configure the shortcut menu in the registry. */
    Set: function (paramNoIcon, menuIconPath) {
      VERB_KEY = format(KEY_FORMAT, VERB_KEY);
      var COMMAND_KEY = VERB_KEY + 'command\\';
      var VERBICON_VALUENAME = VERB_KEY + 'Icon';
      var command = format('{0} //E:jscript "{1}" /Markdown:"%1"', WSH.FullName.replace(/\\cscript\.exe$/i, '\\wscript.exe'), WSH.ScriptFullName);
      wshell.RegWrite(COMMAND_KEY, command);
      wshell.RegWrite(VERB_KEY, 'Convert to &HTML');
      if (paramNoIcon) {
        try {
          wshell.RegDelete(VERBICON_VALUENAME);
        } catch (error) { }
      } else {
        wshell.RegWrite(VERBICON_VALUENAME, menuIconPath);
      }
    },
    /** Remove the shortcut menu by removing the verb key and subkeys. */
    Unset: function () {
      var HKCU = 0x80000001;
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
        try {
          wshell.RegDelete(format(KEY_FORMAT, key));
        } catch (error) { }
      })(VERB_KEY);
    }
  }
  return setup;
})();