/**
 * @file Launches the shortcut target PowerShell
 * script with the selected markdown as an argument.
 * It aims to eliminate the flashing console window
 * when the user clicks on the shortcut menu.
 * @version 0.0.1
 */
import System;
import System.IO;
import System.Text;
import System.Windows.Forms;
import System.Runtime.InteropServices;
import IWshRuntimeLibrary;
import ROOT.CIMV2.WIN32;
import Microsoft.Win32;

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
  var HKCU = Registry.CurrentUser;
  var SHELL_SUBKEY = 'SOFTWARE\\Classes\\SystemFileAssociations\\.md\\shell';
  var VERB = 'cthtml';
  if (param.Set) {
    var VERB_SUBKEY = String.Format('{0}\\{1}', SHELL_SUBKEY, VERB);
    var VERB_KEY = String.Format('{0}\\{1}', HKCU, VERB_SUBKEY);
    // Configure the shortcut menu in the registry.
    var COMMAND_KEY = VERB_KEY + '\\command';
    var command = String.Format('"{0}" /Markdown:"%1"', param.ApplicationPath);
    Registry.SetValue(COMMAND_KEY, '', command);
    Registry.SetValue(VERB_KEY, '', 'Convert to &HTML');
    var iconValueName = 'Icon';
    if (param.NoIcon) {
      var VERB_KEY_OBJ = HKCU.CreateSubKey(VERB_SUBKEY);
      if (VERB_KEY_OBJ) {
        VERB_KEY_OBJ.DeleteValue(iconValueName);
        VERB_KEY_OBJ.Close();
      }
    } else {
      Registry.SetValue(VERB_KEY, iconValueName, param.ApplicationPath);
    }
  } else if (param.Unset) {
    // Remove the shortcut menu.
    var SHELL_KEY_OBJ = HKCU.CreateSubKey(SHELL_SUBKEY);
    if (SHELL_KEY_OBJ) {
      SHELL_KEY_OBJ.DeleteSubKeyTree(VERB);
      SHELL_KEY_OBJ.Close();
    }
  }
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
  Process.WaitForExit(Process.Create(String.Format('C:\\Windows\\System32\\cmd.exe /d /c "{0}"', linkPath)));
  File.Delete(linkPath);
}

/**
 * Save and get the dynamic link.
 * @param {string} markdown is the input markdown path argument.
 * @returns {string} the link path.
 */
function GetDynamicLinkPathWith(markdown) {
  var link = (new WshShellClass()).CreateShortcut(
    String.Format('{1}\\{0}.tmp.lnk', Guid.NewGuid(), Environment.ExpandEnvironmentVariables('%TEMP%'))
  );
  link.TargetPath = GetPwshPath();
  link.Arguments = String.Format('-f "{0}" -Markdown "{1}"', ChangeScriptExtension('.ps1'), markdown);
  link.IconLocation = ChangeScriptExtension('.exe');
  link.Save();
  try {
    return link.FullName;
  } finally {
    Marshal.FinalReleaseComObject(link);
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
  return Path.ChangeExtension(param.ApplicationPath, extension);
}

/**
 * Get the PowerShell Core application path from the registry.
 * @returns {string} the pwsh.exe full path or an empty string.
 */
function GetPwshPath() {
  // The HKLM registry subkey stores the PowerShell Core application path.
  return Registry.GetValue('HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe', '', null);
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
    if (/^\/markdown:[^:]/i.test(arg)) {
      param.Markdown = arg.Split(char[]([':']), 2)[1];
      return param;
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
  var helpTextBuilder = new StringBuilder();
  with (helpTextBuilder) {
    AppendLine('The MarkdownToHtml shortcut launcher.');
    AppendLine('It starts the shortcut menu target script in a hidden window.');
    AppendLine();
    AppendLine('Syntax:');
    AppendLine('  Convert-MarkdownToHtml /Markdown:<markdown file path>');
    AppendLine('  Convert-MarkdownToHtml [/Set[:NoIcon]]');
    AppendLine('  Convert-MarkdownToHtml /Unset');
    AppendLine('  Convert-MarkdownToHtml /Help');
    AppendLine();
    AppendLine("<markdown file path>  The selected markdown's file path.");
    AppendLine('                 Set  Configure the shortcut menu in the registry.');
    AppendLine('              NoIcon  Specifies that the icon is not configured.');
    AppendLine('               Unset  Removes the shortcut menu.');
    AppendLine('                Help  Show the help doc.');
  }
  MessageBox.Show(helpTextBuilder);
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