/**
 * @file Watches the shortcut target PowerShell script runner
 * and redirect the console output and errors to a message box.
 * It aims to separate the window and console user interfaces.
 * @version 0.1.0
 */
import System;
import System.IO;
import System.Text;
import System.Windows.Forms;
import System.Diagnostics;
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
  (new ConversionWatcher(param.Markdown)).Start();
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
 * Represent the ConversionWatcher type.
 */
class ConversionWatcher {
  /** @class */
  static var MessageBox = GetMessageBoxType();

  /**
   * The specified Markdown path argument.
   * @private @type {string}
   */
  var MarkdownPath;
  /**
   * The overwrite prompt text as read from the powershell core console host.
   * @private @type {string}
   */
  var OverwritePromptText = new StringBuilder();

  /**
   * @class @constructs ConversionWatcher
   * @param {string} markdown is the specified Markdown path argument.
   */
  function ConversionWatcher(markdown) {
    MarkdownPath = markdown;
  }
  
  /**
   * Show the overwrite prompt that the child process sends.
   * Subsequently, wait for the user's response.
   * Handle the event when the PowerShell Core (child) process
   * redirects output to the parent Standard Output stream.
   * @param pwshExe it the sender child process.
   * @param {object} pwshExe.StandardInput  the standard input stream.
   * @param outEvtArgs the received output text line.
   */
  function HandleOutputDataReceived(pwshExe: Object, outEvtArgs: DataReceivedEventArgs) {
    var outData = outEvtArgs.Data;
    if (!String.IsNullOrEmpty(outData)) {
      // Show the message box when the text line is a question.
      // Otherwise, append the text line to the overall message text variable.
      if (outData.match(/\?\s*$/)) {
        OverwritePromptText.AppendLine();
        OverwritePromptText.AppendLine(outData);
        // Write the user's choice to the child process console host.
        pwshExe.StandardInput.WriteLine(MessageBox.Show(String(OverwritePromptText), MessageBox.WARNING));
        // Optional
        OverwritePromptText.Clear();
      } else {
        OverwritePromptText.AppendLine(outData);
      }
    }
  }

  /**
   * Show the error message that the child process writes on the console host.
   * It handles the event when the child process redirects errors to the parent Standard
   * Error stream. Raised exceptions are terminating errors. Thus, this handler only notifies
   * the user of an error and displays the error message. For this reason, this subroutine
   * does not define the sender objPwshExe object parameter in its signature.
   * @private
   * @param {string} errData the error message text.
   */
  static function HandleErrorDataReceived(errData) {
    if (errData) {
      // Remove the ANSI escaped characters from the error message data text.
      MessageBox.Show(errData.replace(/\x1B\[[0-9;]*[a-zA-Z]/g, '').replace(/^[^:]*:\s+/, ''));
    }
  }

  /**
   * Start a PowerShell Core process that runs the shortcut menu target
   * script with the markdown path as the argument.
   * The Try-Catch handles the errors thrown by the process. The Standard Error
   * Stream encoding is not utf-8. For this reason, it surrounds the message with
   * unwanted characters. The error message delimiter constant string separates
   * the informative message from noisy characters.
   */
  public function Start() {
    OverwritePromptText.Clear();
    var pwshExe = new Process();
    var pwshStartInfo = new ProcessStartInfo(
      // The HKLM registry subkey stores the PowerShell Core application path.
      Registry.GetValue('HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe', '', null),
      String.Format(
        '-nop -ep Bypass -w Hidden -cwa "$ErrorView = ""ConciseView""; Import-Module $args[0]; {2} -MarkdownPath $args[1]" "{0}" "{1}"',
        ChangeScriptExtension('.psm1'), MarkdownPath, Path.GetFileNameWithoutExtension(param.ApplicationPath)
      )
    );
    // Redirect streams to the launcher process.
    with (pwshStartInfo) {
      RedirectStandardOutput = true;
      RedirectStandardInput = true;
      RedirectStandardError = true;
      CreateNoWindow = true;
      UseShellExecute = false;
    }
    pwshExe.StartInfo = pwshStartInfo;
    // Register the event handlers.
    pwshExe.EnableRaisingEvents = true;
    pwshExe.add_OutputDataReceived(HandleOutputDataReceived);
    // Start the child process.
    with (pwshExe) {
      Start();
      BeginOutputReadLine();
      WaitForExit();
    }
    // When the process terminated with an error.
    HandleErrorDataReceived(pwshExe.StandardError.ReadToEnd());
    with (pwshExe) {
      Close();
      Dispose();
    }
  }

  /**
   * @returns the MessageBox type.
   */
  static function GetMessageBoxType() {
    /** @private @constant {string} */
    var MESSAGE_BOX_TITLE = 'Convert to HTML';
    /** 
     * @private @constant
     * Do not remove repetition. It is there to solve a bug.
     */
    var EXPECTED_DIALOGRESULT: System.Array = [DialogResult.Yes, DialogResult.No, DialogResult.No];
  
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
      try {
        return EXPECTED_DIALOGRESULT.GetValue(
          System.Array.BinarySearch(
            EXPECTED_DIALOGRESULT,
            System.Windows.Forms.MessageBox.Show(message, MESSAGE_BOX_TITLE, messageButton, messageType)
          )
        );
      } catch (error) { }
    }
  
    return MessageBox;
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