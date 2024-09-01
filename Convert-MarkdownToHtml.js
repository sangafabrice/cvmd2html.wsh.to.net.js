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
import IWshRuntimeLibrary;
import WbemScripting;
import Scriptlet;

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
  /**
   * @constant {int64} 0x80000001
   * The real value is not returned because
   * of overflow when casting int64 to int32.
   */
  var HKCU = -2147483647;
  var VERB_KEY = 'SOFTWARE\\Classes\\SystemFileAssociations\\.md\\shell\\cthtml';
  var registry: SWbemObject = (new SWbemLocatorClass()).ConnectServer().Get('StdRegProv');
  var getMethod = registry.Methods_.Item;
  var execMethod = registry.ExecMethod_;
  var inParam;
  /**
   * Set the input parameter of the StdRegProv methods.
   * @param {string} parameter
   * @param {any} value
   */
  var setInParam = function (parameter, value) {
    inParam.Properties_.Item(parameter).Value = value;
  }
  if (param.Set) {
    // Configure the shortcut menu in the registry.
    var COMMAND_KEY = VERB_KEY + '\\command';
    var command = Format('"{0}" /Markdown:"%1"', param.ApplicationPath);
    var setStringValueMethod: SWbemMethod = getMethod('SetStringValue');
    inParam = setStringValueMethod.InParameters.SpawnInstance_();
    // Create the command key.
    setInParam('hDefKey', HKCU);
    setInParam('sSubKeyName', COMMAND_KEY);
    // Reuse inParam spawned by setStringValueMethod in CreateKey signature.
    execMethod('CreateKey', inParam);
    // Set the command key default value.
    setInParam('sValue', command);
    execMethod(setStringValueMethod.Name, inParam);
    // Set the verb key default value.
    setInParam('sSubKeyName', VERB_KEY);
    setInParam('sValue', 'Convert to &HTML');
    execMethod(setStringValueMethod.Name, inParam);
    setInParam('sValueName', 'Icon');
    if (param.NoIcon) {
      // Remove the shortcut icon.
      // Reuse inParam spawned by setStringValueMethod in DeleteValue signature.
      execMethod('DeleteValue', inParam);
    } else {
      // Set the shortcut icon to the assembly main icon resource.
      setInParam('sValue', ChangeScriptExtension('.ico'));
      execMethod(setStringValueMethod.Name, inParam);
    }
    Marshal.FinalReleaseComObject(setStringValueMethod);
    setStringValueMethod = null;
  } else if (param.Unset) {
    // Remove the shortcut menu.
    // Remove the verb key and subkeys.
    var enumKeyMethod: SWbemMethod = getMethod('EnumKey');
    inParam = enumKeyMethod.InParameters.SpawnInstance_();
    setInParam('hDefKey', HKCU);
    // Recursion is used because a key with subkeys cannot be deleted.
    // Recursion helps removing the leaf keys first.
    var deleteVerbKey = function(key) {
      var recursive = function func(key) {
        setInParam('sSubKeyName', key);
        var sNames = execMethod(enumKeyMethod.Name, inParam).Properties_.Item('sNames').Value;
        if (sNames != null) {
          for (var index = 0; index < sNames.length; index++) {
            func(key + '\\' + sNames[index]);
          }
        }
        // Reuse inParam spawned by enumKeyMethod in DeleteKey signature.
        setInParam('sSubKeyName', key);
        execMethod('DeleteKey', inParam)
      };
      recursive(key);
    }
    deleteVerbKey(VERB_KEY);
    deleteVerbKey = null;
    Marshal.FinalReleaseComObject(enumKeyMethod);
    enumKeyMethod = null;
  }
  Marshal.FinalReleaseComObject(registry);
  Marshal.FinalReleaseComObject(inParam);
  registry = null;
  inParam = null;
  Quit(0);
}

ShowHelp();

/**
 * Start the shortcut target PowerShell script with
 * the path of the selected markdown file as an argument.
 * @param {string} markdown is the input markdown path argument.
 */
function StartWith(markdown) {
  var linkPath = GetDynamicLinkPathWith(markdown);
  var WINDOW_STYLE_HIDDEN = Object(0);
  var WAIT_ON_RETURN = Object(true);
  (new WshShellClass()).Run(Format('"{0}"', linkPath), WINDOW_STYLE_HIDDEN, WAIT_ON_RETURN);
  (new FileSystemObjectClass()).DeleteFile(linkPath, true);
}

/**
 * Save and get the dynamic link.
 * @param {string} markdown is the input markdown path argument.
 * @returns {string} the link path.
 */
function GetDynamicLinkPathWith(markdown) {
  var shell: WshShell = new WshShellClass();
  var link = shell.CreateShortcut(shell.ExpandEnvironmentStrings(Format(
    '%TEMP%\\{0}.tmp.lnk',
    IGenScriptletTLib.GUID
  )));
  // The HKLM registry subkey stores the PowerShell Core application path.
  link.TargetPath = shell.RegRead('HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe\\');
  link.Arguments = Format(
    '-ep Bypass -nop -w Hidden -f "{0}" -Markdown "{1}"',
    [ChangeScriptExtension('.ps1'), markdown]
  );
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
  (new WshShellClass()).Popup(helpText, 0);
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