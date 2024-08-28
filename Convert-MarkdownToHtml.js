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
  var HKCU = 0x80000001;
  var VERB_KEY = 'SOFTWARE\\Classes\\SystemFileAssociations\\.md\\shell\\cthtml';
  var registry = GetObject('winmgmts:StdRegProv');
  if (param.Set) {
    var shortcutIconPath = ChangeScriptExtension('.ico');
    // Configure the shortcut menu in the registry.
    var COMMAND_KEY = VERB_KEY + '\\command';
    var command = Format(
      '{0} //E:jscript "{1}" /Markdown:"%1"',
      WSH.FullName.replace(/\\cscript\.exe$/i, '\\wscript.exe'),
      WSH.ScriptFullName
    );
    registry.CreateKey(HKCU, COMMAND_KEY);
    registry.SetStringValue(HKCU, COMMAND_KEY, null, command);
    registry.SetStringValue(HKCU, VERB_KEY, null, 'Convert to &HTML');
    var iconValueName = 'Icon';
    if (param.NoIcon) {
      registry.DeleteValue(HKCU, VERB_KEY, iconValueName);
    } else {
      registry.SetStringValue(HKCU, VERB_KEY, iconValueName, shortcutIconPath);
    }
  } else if (param.Unset) {
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
      registry.DeleteKey(HKCU, key);
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
  // Create the intermediate link.
  var linkPath = GetDynamicLinkPathWith(markdown);
  //Start the link.
  var WINDOW_STYLE_HIDDEN = 0xC;
  var startInfo = GetObject('winmgmts:Win32_ProcessStartup').SpawnInstance_();
  startInfo.ShowWindow = WINDOW_STYLE_HIDDEN;
  var processService = GetObject('winmgmts:Win32_Process');
  var createMethod = processService.Methods_('Create');
  var inParam = createMethod.InParameters.SpawnInstance_();
  inParam.CommandLine = Format('C:\\Windows\\System32\\cmd.exe /d /c "{0}"', linkPath);
  inParam.ProcessStartupInformation = startInfo;
  WaitForExit(processService.ExecMethod_(createMethod.Name, inParam).ProcessId);
  // Delete the link.
  WSH.CreateObject('Scripting.FileSystemObject').DeleteFile(linkPath, true);
}

/**
 * Save and get the dynamic link.
 * @param {string} markdown is the input markdown path argument.
 * @returns {string} the link path.
 */
function GetDynamicLinkPathWith(markdown) {
  var shell = WSH.CreateObject('WScript.Shell');
  var link = shell.CreateShortcut(shell.ExpandEnvironmentStrings(Format(
    '%TEMP%\\{0}.tmp.lnk',
    WSH.CreateObject('Scriptlet.TypeLib').Guid.substr(1, 36).toLowerCase()
  )));
  link.TargetPath = GetPwshPath();
  link.Arguments = Format('-f "{0}" -Markdown "{1}"', ChangeScriptExtension('.ps1'), markdown);
  link.IconLocation = ChangeScriptExtension('.ico');
  link.Save();
  return link.FullName;
}

/**
 * Wait for the process exit.
 * @param {number} processId is the process identifier.
 */
function WaitForExit(processId) {
  try {
    var moniker = 'winmgmts:Win32_Process=' + processId;
    while (GetObject(moniker).Name == 'cmd.exe') { }
  } catch (error) { }
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
  var registry = GetObject('winmgmts:StdRegProv');
  var getStringValueMethod = registry.Methods_('GetStringValue');
  var inParam = getStringValueMethod.InParameters.SpawnInstance_();
  // The HKLM registry subkey stores the PowerShell Core application path.
  inParam.sSubKeyName = 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe';
  var outParam = registry.ExecMethod_(getStringValueMethod.Name, inParam);
  if (!outParam.ReturnValue) {
    return outParam.sValue;
  }
  return '';
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