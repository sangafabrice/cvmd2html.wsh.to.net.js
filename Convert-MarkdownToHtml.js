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
import Microsoft.VisualBasic;
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
  var VERB_KEY = 'SOFTWARE\\Classes\\SystemFileAssociations\\.md\\shell\\cthtml';
  var KEY_FORMAT = 'HKCU\\{0}\\';
  var registry;
  if (param.Set) {
    var VERB_KEY = Format(KEY_FORMAT, VERB_KEY);
    var COMMAND_KEY = VERB_KEY + 'command\\';
    var VERBICON_VALUENAME = VERB_KEY + 'Icon';
    registry = new ActiveXObject('WScript.Shell');
    var shortcutIconPath = ChangeScriptExtension('.ico');
    // Create the link with the partial "arguments" string.
    var link = registry.CreateShortcut(ChangeScriptExtension('.lnk'));;
    link.TargetPath = GetPwshPath();
    link.Arguments = GetCustomIconLinkArguments();
    link.IconLocation = shortcutIconPath;
    link.Save();
    Marshal.FinalReleaseComObject(link);
    link = null;
    // Configure the shortcut menu in the registry.
    var command = Format('"{0}" /Markdown:"%1"', param.ApplicationPath);
    registry.RegWrite(COMMAND_KEY, command);
    registry.RegWrite(VERB_KEY, 'Convert to &HTML');
    if (param.NoIcon) {
      registry.RegDelete(VERBICON_VALUENAME);
    } else {
      registry.RegWrite(VERBICON_VALUENAME, shortcutIconPath);
    }
  } else if (param.Unset) {
    var HKCU = -2147483647;
    registry = Interaction.GetObject('winmgmts:StdRegProv');
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
        Interaction.CreateObject('WScript.Shell').RegDelete(Format(KEY_FORMAT, key));
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
  var linkPath = ChangeScriptExtension('.lnk');
  if (!IsLinkReady(linkPath)) {
    return;
  }
  Interaction.Shell(
    Format('C:\\Windows\\System32\\cmd.exe /d /c ""{0}" "{1}""', [linkPath, markdown]),
    AppWinStyle.Hide
  );
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
  // The HKLM registry subkey stores the PowerShell Core application path.
  var PWSH_KEY = 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe\\';
  return Interaction.CreateObject('WScript.Shell').RegRead(PWSH_KEY);
}

/**
 * Check the link target command.
 * @param {object} linkPath is the shortcut link path.
 * @returns {boolean} True if the target command is as expected, false otherwise.
 */
function IsLinkReady(linkPath) {
  var link = Interaction.CreateObject('WScript.Shell').CreateShortcut(linkPath);
  var format = '{0} {1}';
  try {
    return Strings.StrComp(
      Format(format, [link.TargetPath, link.Arguments]),
      Format(format, [GetPwshPath(), GetCustomIconLinkArguments()]),
      CompareMethod.Text
    ) == 0;
  } finally {
    if (link != undefined) {
      Marshal.FinalReleaseComObject(link);
      link = null;
    }
  }
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
  if (args.constructor !== Array) {
    return Strings.Replace(format, '{0}', args);
  }
  while (args.length > 0) {
    format = Strings.Replace(format, '{' + (args.length - 1) + '}', args.pop());
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
  Interaction.MsgBox(String(helpText));
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