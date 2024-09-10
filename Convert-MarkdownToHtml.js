/**
 * @file Converts markdown file content to html document.
 * @version 1.0.0.2
 */
import System;
import System.IO;
import System.Text;
import System.Windows.Forms;
import System.Diagnostics;
import Microsoft.Win32;
import Markdig;

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
(function() {
/* The app module */

/** @class */
var MessageBox = GetMessageBoxType();
SetHtmlContent(GetHtmlPath(), Markdown.ToHtml(File.ReadAllText(CheckMarkdown(param.Markdown))));
Quit(0);

/**
 * Validate the input markdown path string.
 * @param {string} markdown the input markdown file path.
 * @returns {string} the markdown file path if it is valid.
*/
function CheckMarkdown(markdown) {
  if (String.Compare(String(Path.GetExtension(markdown)), '.md', true)) {
    MessageBox.Show(String.Format('"{0}" is not a markdown (.md) file.', markdown));
  }
  if (!File.Exists(markdown)) {
    MessageBox.Show(String.Format('"{0}" cannot be found.', markdown));
  }
  return markdown;
}

/**
 * Write the html text to the output HTML file.
 * It notifies the user when the operation did not complete with success.
 * @param {string} htmlPath is the output html path.
 * @param {string} htmlContent is the content of the html file.
 */
function SetHtmlContent(htmlPath, htmlContent) {
  try {
    File.WriteAllText(htmlPath, htmlContent);
  } catch (error) {
    if (error.number == -2146823266) {
      MessageBox.Show(String.Format('Access to the path "{0}" is denied.', htmlPath));
    } else {
      MessageBox.Show(String.Format('Unspecified error trying to write to "{0}".', htmlPath));
    }
  }
}

/**
 * This function returns the output path when it is unique
 * without prompts or when the user accepts to overwrite an
 * existing HTML file. Otherwise, it exits the script.
 * @returns {string} the output html path.
 */
function GetHtmlPath() {
  var htmlPath = Path.ChangeExtension(param.Markdown, '.html');
  if (File.Exists(htmlPath)) {
    MessageBox.Show(
      String.Format('The file "{0}" already exists.\n\nDo you want to overwrite it?', htmlPath),
      MessageBox.WARNING
    );
  } else if (Directory.Exists(htmlPath)) {
    MessageBox.Show(String.Format('"{0}" cannot be overwritten because it is a directory.', htmlPath));
  }
  return htmlPath;
}

/**
 * @returns the MessageBox type.
 */
function GetMessageBoxType() {
  /** @private @constant {string} */
  var MESSAGE_BOX_TITLE = 'Convert to HTML';
  /** 
   * @private @constant
   * Do not remove repetition. It is there to solve a bug.
   */
  var EXPECTED_DIALOGRESULT: System.Array = [DialogResult.OK, DialogResult.No, DialogResult.No];

  /**
   * Represents the markdown conversion message box.
   * @typedef {object} MessageBox
   * @property {number} WARNING specifies that the dialog shows a warning message.
   */
  var MessageBox = { };

  /** @public @static @readonly @property {number} */
  MessageBox.WARNING = MessageBoxIcon.Exclamation;
  // Object.defineProperty() method does not work in WSH.
  // It is not possible in this implementation to make the
  // property non-writable.

  /**
   * Show a warning message or an error message box.
   * The function does not return anything when the message box is an error.
   * @public @static @method Show @memberof MessageBox
   * @param {string} message is the message text.
   * @param {number} [messageType = ERROR_MESSAGE] message box type (Warning/Error).
   * @returns {string|void} "Yes" or "No" depending on the user's click when the message box is a warning.
   */
  MessageBox.Show = function(message, messageType: MessageBoxIcon) {
    if (messageType != MessageBoxIcon.Error && messageType != MessageBoxIcon.Exclamation) {
      messageType = MessageBoxIcon.Error;
    }
    // The error message box shows the OK button alone.
    // The warning message box shows the alternative Yes or No buttons.
    var messageButton: MessageBoxButtons = messageType == MessageBoxIcon.Error ? MessageBoxButtons.OK:MessageBoxButtons.YesNo;
    if (
      System.Array.BinarySearch(
        EXPECTED_DIALOGRESULT,
        System.Windows.Forms.MessageBox.Show(message, MESSAGE_BOX_TITLE, messageButton, messageType)
      ) >= 0
    ) {
      Quit(1);
    }
  }

  return MessageBox;
}
})();
}

if (param.Help) {
  ShowHelp();
}

/* Configuration and settings */
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