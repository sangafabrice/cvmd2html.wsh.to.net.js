/**
 * @file Launches the shortcut target PowerShell
 * script with the selected markdown as an argument.
 * It aims to eliminate the flashing console window
 * when the user clicks on the shortcut menu.
 * @version 0.0.1
 */

/**
 * The parameters and arguments.
 * @typedef {object} ParamHash
 * @property {string} Markdown is the selected markdown file path.
 * @property {boolean} Set installs the shortcut menu.
 * @property {boolean} NoIcon installs the shortcut menu without icon.
 * @property {boolean} Unset uninstalls the shortcut menu.
 * @property {boolean} Help shows help.
 */

/** @type {ParamHash} */
var param = GetParameters();

if (param.Markdown) {
  StartWith(param.Markdown);
  WSH.Quit();
}

if (param.Help) {
  ShowHelp();
}

if (param.Set || param.Unset) {
  var VERB_KEY = 'SOFTWARE\\Classes\\SystemFileAssociations\\.md\\shell\\cthtml';
  var KEY_FORMAT = 'HKCU\\{0}\\';
  if (param.Set) {
    var VERB_KEY = Format(KEY_FORMAT, VERB_KEY);
    var COMMAND_KEY = VERB_KEY + 'command\\';
    var VERBICON_VALUENAME = VERB_KEY + 'Icon';
    var registry = new ActiveXObject('WScript.Shell');
    var shortcutIconPath = ChangeScriptExtension('.ico');
    // Create the link with the partial "arguments" string.
    var link = GetCustomIconLink();
    link.TargetPath = GetPwshPath();
    link.Arguments = GetCustomIconLinkArguments();
    link.IconLocation = shortcutIconPath;
    link.Save();
    // Configure the shortcut menu in the registry.
    var command = Format(
      '{0} //E:jscript "{1}" /Markdown:"%1"',
      WSH.FullName.replace(/\\cscript\.exe$/i, '\\wscript.exe'),
      WSH.ScriptFullName
    );
    registry.RegWrite(COMMAND_KEY, command);
    registry.RegWrite(VERB_KEY, 'Convert to &HTML');
    if (param.NoIcon) {
      registry.RegDelete(VERBICON_VALUENAME);
    } else {
      registry.RegWrite(VERBICON_VALUENAME, shortcutIconPath);
    }
  } else if (param.Unset) {
    var HKCU = 0x80000001;
    var registry = GetObject('winmgmts:StdRegProv');
    // Remove the shortcut menu.
    // Remove the verb key and subkeys.
    var enumKeyMethod = registry.Methods_('EnumKey');
    var inParam = enumKeyMethod.InParameters.SpawnInstance_();
    inParam.hDefKey = HKCU;
    // Recursion is used because a key with subkeys cannot be deleted.
    // Recursion helps removing the leaf keys first.
    (function(key) {
      inParam.sSubKeyName = key;
      var sNames = registry.ExecMethod_(enumKeyMethod.Name, inParam).sNames;
      if (sNames != null) {
        sNames = sNames.toArray();
        for (var index = 0; index < sNames.length; index++) {
          arguments.callee(key + '\\' + sNames[index]);
        }
      }
      WSH.CreateObject('WScript.Shell').RegDelete(Format(KEY_FORMAT, key));
    })(VERB_KEY);
  }
  WSH.Quit();
}

ShowHelp();

/**
 * Start the shortcut target PowerShell script with
 * the path of the selected markdown file as an argument.
 * @param {string} markdown is the input markdown path argument.
 */
function StartWith(markdown) {
  var link = GetCustomIconLink();
  if (!IsLinkReady(link)) {
    return;
  }
  var WINDOW_STYLE_HIDDEN = 0;
  WSH.CreateObject('WScript.Shell').Run(
    Format('"{0}" "{1}"', link.FullName, markdown), WINDOW_STYLE_HIDDEN
  );
}

/**
 * Change the launcher script path extension.
 * This change implies that the launcher script and the resulting
 * path file reside in the same directory and have the same name.
 * @param {string} extension is the new extension.
 * @returns {string} a file path with the new extension.
 */
function ChangeScriptExtension(extension) {
  return WSH.ScriptFullName.replace(/\.js$/i, extension);
}

/**
 * Get the PowerShell Core application path from the registry.
 * @returns {string} the pwsh.exe full path or an empty string.
 */
function GetPwshPath() {
  // The HKLM registry subkey stores the PowerShell Core application path.
  var PWSH_KEY = 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe\\';
  return WSH.CreateObject('WScript.Shell').RegRead(PWSH_KEY);
}

/**
 * Check the link target command.
 * @param {object} link is the shortcut link.
 * @param {string} link.TargetPath is the path to the runner.
 * @param {string} link.Arguments is the target command line list of arguments.
 * @returns {boolean} True if the target command is as expected, false otherwise.
 */
function IsLinkReady(link) {
  var format = '{0} {1}';
  return Format(format, link.TargetPath, link.Arguments).toLowerCase() ==
    Format(format, GetPwshPath(), GetCustomIconLinkArguments()).toLowerCase();
}

/**
 * Get the custom icon link object from its path.
 * @returns {object} the specified link file object.
 */
function GetCustomIconLink() {
  return WSH.CreateObject('WScript.Shell').CreateShortcut(ChangeScriptExtension('.lnk'));
}

/**
 * Get the partial "arguments" property string of the custom icon link.
 * The command is partial because it does not include the markdown file path string.
 * The markdown file path string will be input when calling the shortcut link.
 * @returns {string} the "arguments" property of the custom icon link.
 */
function GetCustomIconLinkArguments() {
  return Format('-ep Bypass -nop -w Hidden -f "{0}" -Markdown', ChangeScriptExtension('.ps1'));
}

/**
 * Replace the format item "{n}" by the nth input in a list of arguments.
 * @param {string} format the pattern format.
 * @param {...string} args the replacement texts.
 * @returns {string} a copy of format with the format items replaced by args.
 */
function Format(format, args) {
  args = Array.prototype.slice.call(arguments).slice(1);
  while (args.length > 0) {
    format = format.replace(new RegExp('\\{' + (args.length - 1) + '\\}', 'g'), args.pop());
  }
  return format;
}

/**
 * Get the input arguments and parameters.
 * @returns {ParamHash|undefined} a value-name pair of arguments.
 */
function GetParameters() {
  var WshArguments = WSH.Arguments;
  var WshNamed = WshArguments.Named;
  var paramCount = WshArguments.Count();
  if (paramCount == 1) {
    var paramMarkdown = WshArguments.Named('Markdown');
    if (WshNamed.Exists('Markdown') && paramMarkdown != undefined && paramMarkdown.length) {
      return { Markdown: paramMarkdown };
    }
    var paramHelp = WshNamed.Exists('Help');
    if (paramHelp) {
      return { Help: paramHelp };
    }
    var param = { Set: WshNamed.Exists('Set') };
    if (param.Set) {
      var noIconParam = WshArguments.Named('Set');
      var isNoIconParam = false;
      param.NoIcon = noIconParam != undefined && (isNoIconParam = /^NoIcon$/i.test(noIconParam));
      if (noIconParam == undefined || isNoIconParam) {
        return param;
      }
    }
    param = { Unset: WshNamed.Exists('Unset') };
    if (param.Unset && WshArguments.Named('Unset') == undefined) {
      return param;
    }
  } else if (paramCount == 0) {
    return {
      Set: true,
      NoIcon: false
    }
  }
  ShowHelp();
}

/**
 * Show help and quit.
 */
function ShowHelp() {
  var helpText = '';
  helpText += 'The MarkdownToHtml shortcut launcher.\n';
  helpText += 'It starts the shortcut menu target script in a hidden window.\n\n';
  helpText += 'Syntax:\n';
  helpText += '  Convert-MarkdownToHtml.js /Markdown:<markdown file path>\n';
  helpText += '  Convert-MarkdownToHtml.js [/Set[:NoIcon]]\n';
  helpText += '  Convert-MarkdownToHtml.js /Unset\n';
  helpText += '  Convert-MarkdownToHtml.js /Help\n\n';
  helpText += "<markdown file path>  The selected markdown's file path.\n";
  helpText += '                 Set  Configure the shortcut menu in the registry.\n';
  helpText += '              NoIcon  Specifies that the icon is not configured.\n';
  helpText += '               Unset  Removes the shortcut menu.\n';
  helpText += '                Help  Show the help doc.\n';
  with (WSH) {
    Echo(helpText);
    Quit();
  }
}