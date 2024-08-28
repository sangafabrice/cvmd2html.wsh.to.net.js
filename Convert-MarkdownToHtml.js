/**
 * @file Launches the shortcut target PowerShell
 * script with the selected markdown as an argument.
 * It aims to eliminate the flashing console window
 * when the user clicks on the shortcut menu.
 * @version 0.0.1
 */
@set @MAJOR = 0
@set @MINOR = 0
@set @BUILD = 1
@set @REVISION = 0

import System;
import System.Runtime.InteropServices;
import System.Reflection;

[assembly: AssemblyProduct('MarkdownToHtml Shortcut')]
[assembly: AssemblyInformationalVersion(@MAJOR + '.' + @MINOR + '.' + @BUILD + '.' + @REVISION)]
[assembly: AssemblyCopyright('\u00A9 2024 sangafabrice')]
[assembly: AssemblyCompany('sangafabrice')]
[assembly: AssemblyVersion(@MAJOR + '.' + @MINOR + '.' + @BUILD + '.' + @REVISION)]
[assembly: AssemblyTitle('CvMd2Html Launcher')]

/**
 * The parameters and arguments.
 * @typedef {object} ParamHash
 * @property {string} ApplicationPath is the assembly path.
 * @property {string} Markdown is the selected markdown file path.
 * @property {boolean} Set installs the shortcut menu.
 * @property {boolean} NoIcon installs the shortcut menu without icon.
 * @property {boolean} Unset uninstalls the shortcut menu.
 * @property {boolean} Help shows help.
 */

/** @type {ParamHash} */
var param = GetParameters(Environment.GetCommandLineArgs());

if (param.Markdown) {
  StartWith(param.Markdown);
  Quit(0);
}

if (param.Help) {
  ShowHelp();
}

if (param.Set || param.Unset) {
  var HKCU = param.Set ? 0x80000001:-2147483647;
  var VERB_KEY = 'SOFTWARE\\Classes\\SystemFileAssociations\\.md\\shell\\cthtml';
  var registry = GetObject('winmgmts:StdRegProv');
  if (param.Set) {
    var shortcutIconPath = ChangeScriptExtension('.ico');
    // Configure the shortcut menu in the registry.
    var COMMAND_KEY = VERB_KEY + '\\command';
    var command = Format('"{0}" /Markdown:"%1"', param.ApplicationPath);
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
    var enumKeyMethod = registry.Methods_.Item('EnumKey');
    var inParam = enumKeyMethod.InParameters.SpawnInstance_();
    inParam.hDefKey = HKCU;
    // Recursion is used because a key with subkeys cannot be deleted.
    // Recursion helps removing the leaf keys first.
    var deleteVerbKey = function(key) {
      var recursive = function func(key) {
        inParam.sSubKeyName = key;
        var sNames = registry.ExecMethod_(enumKeyMethod.Name, inParam).sNames;
        if (sNames != null) {
          for (var index = 0; index < sNames.length; index++) {
            func(key + '\\' + sNames[index]);
          }
        }
        registry.DeleteKey(HKCU, key);
      };
      recursive(key);
    }
    deleteVerbKey(VERB_KEY);
    deleteVerbKey = null;
    Marshal.FinalReleaseComObject(enumKeyMethod);
    Marshal.FinalReleaseComObject(inParam);
    enumKeyMethod = null;
    inParam = null;
  }
  Marshal.FinalReleaseComObject(registry);
  registry = null;
  Quit(0);
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
  var createMethod = processService.Methods_.Item('Create');
  var inParam = createMethod.InParameters.SpawnInstance_();
  inParam.CommandLine = Format('C:\\Windows\\System32\\cmd.exe /d /c "{0}"', linkPath);
  inParam.ProcessStartupInformation = startInfo;
  WaitForExit(processService.ExecMethod_(createMethod.Name, inParam).ProcessId);
  // Delete the link.
  (new ActiveXObject('Scripting.FileSystemObject')).DeleteFile(linkPath, true);
  Marshal.FinalReleaseComObject(startInfo);
  Marshal.FinalReleaseComObject(processService);
  Marshal.FinalReleaseComObject(createMethod);
  Marshal.FinalReleaseComObject(inParam);
  startInfo = null;
  processService = null;
  createMethod = null;
  inParam = null;
}

/**
 * Save and get the dynamic link.
 * @param {string} markdown is the input markdown path argument.
 * @returns {string} the link path.
 */
function GetDynamicLinkPathWith(markdown) {
  var shell = new ActiveXObject('WScript.Shell');
  var link = shell.CreateShortcut(shell.ExpandEnvironmentStrings(Format(
    '%TEMP%\\{0}.tmp.lnk',
    (new ActiveXObject('Scriptlet.TypeLib')).Guid.substr(1, 36).toLowerCase()
  )));
  link.TargetPath = GetPwshPath();
  link.Arguments = Format('-f "{0}" -Markdown "{1}"', [ChangeScriptExtension('.ps1'), markdown]);
  link.IconLocation = ChangeScriptExtension('.ico');
  link.Save();
  try {
    return link.FullName;
  } finally {
    Marshal.FinalReleaseComObject(shell);
    Marshal.FinalReleaseComObject(link);
    shell = null;
    link = null;
  }
}

/**
 * Wait for the process exit.
 * @param {number} parentProcessId is the parent process identifier.
 */
function WaitForExit(parentProcessId) {
  // The process termination event query.
  // Select the process whose parent is the intermediate process used for executing the link.
  var wmiQuery = 'SELECT * FROM __InstanceDeletionEvent WITHIN 1 WHERE TargetInstance ISA "Win32_Process" AND ' +
    'TargetInstance.Name="pwsh.exe" AND TargetInstance.ParentProcessId=' + parentProcessId;
  // Wait for the process to exit.
  (new ActiveXObject('WbemScripting.SWbemLocator')).ConnectServer().ExecNotificationQuery(wmiQuery).NextEvent();
}

/**
 * Change the launcher assembly path extension.
 * This change implies that the launcher and the resulting
 * path file reside in the same directory and have the same name.
 * @param {string} extension is the new extension.
 * @returns {string} a file path with the new extension.
 */
function ChangeScriptExtension(extension) {
  return param.ApplicationPath.replace(/\.exe$/i, extension);
}

/**
 * Get the PowerShell Core application path from the registry.
 * @returns {string} the pwsh.exe full path or an empty string.
 */
function GetPwshPath() {
  var registry = GetObject('winmgmts:StdRegProv');
  var getStringValueMethod = registry.Methods_.Item('GetStringValue');
  var inParam = getStringValueMethod.InParameters.SpawnInstance_();
  // The HKLM registry subkey stores the PowerShell Core application path.
  inParam.sSubKeyName = 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe';
  var outParam = registry.ExecMethod_(getStringValueMethod.Name, inParam);
  try {
    if (!outParam.ReturnValue) {
      return outParam.sValue;
    }
    return '';
  } finally {
    Marshal.FinalReleaseComObject(registry);
    Marshal.FinalReleaseComObject(getStringValueMethod);
    Marshal.FinalReleaseComObject(inParam);
    Marshal.FinalReleaseComObject(outParam);
    registry = null;
    getStringValueMethod = null;
    inParam = null;
    outParam = null;
  }
}

/**
 * Replace the format item "{n}" by the nth input in a list of arguments.
 * @param {string} format the pattern format.
 * @param {...string} args the replacement texts.
 * @returns {string} a copy of format with the format items replaced by args.
 */
function Format(format, args) {
  if (args.constructor !== Array) {
    return format.replace(/\{0\}/g, args);
  }
  while (args.length > 0) {
    format = format.replace(new RegExp('\\{' + (args.length - 1) + '\\}', 'g'), args.pop());
  }
  return format;
}

/**
 * Get the input arguments and parameters.
 * @param {string[]} args is the list of command line arguments including the command path.
 * @returns {ParamHash|undefined} a value-name pair of arguments.
 */
function GetParameters(args) {
  var applicationPath = args[0];
  if (args.length == 2) {
    var arg = args[1];
    var param = { ApplicationPath: applicationPath };
    var paramName = arg.split(':', 1)[0].toLowerCase();
    if (paramName == '/markdown') {
      param.Markdown = arg.replace(new RegExp('^' + paramName + ':?', 'i'), '')
      if (param.Markdown.length > 0) {
        return param;
      }
    }
    switch (arg.toLowerCase()) {
      case '/set':
        param.Set = true;
        param.NoIcon = false;
        return param;
      case '/set:noicon':
        param.Set = true;
        param.NoIcon = true;
        return param;
      case '/unset':
        param.Unset = true;
        return param;
      case '/help':
        param.Help = true;
        return param;
    }
  } else if (args.length == 1) {
    return {
      Set: true,
      NoIcon: false,
      ApplicationPath: applicationPath
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
  helpText += '  Convert-MarkdownToHtml /Markdown:<markdown file path>\n';
  helpText += '  Convert-MarkdownToHtml [/Set[:NoIcon]]\n';
  helpText += '  Convert-MarkdownToHtml /Unset\n';
  helpText += '  Convert-MarkdownToHtml /Help\n\n';
  helpText += "<markdown file path>  The selected markdown's file path.\n";
  helpText += '                 Set  Configure the shortcut menu in the registry.\n';
  helpText += '              NoIcon  Specifies that the icon is not configured.\n';
  helpText += '               Unset  Removes the shortcut menu.\n';
  helpText += '                Help  Show the help doc.\n';
  (new ActiveXObject('WScript.Shell')).Popup(helpText, 0);
  Quit(1);
}

/**
 * Clean up and quit.
 * @param {number} exitCode .
 */
function Quit(exitCode) {
  GC.Collect();
  Environment.Exit(exitCode);
}