/**
 * @file Watches the shortcut target PowerShell script runner
 * and redirect the console output and errors to a message box.
 * It aims to separate the window and console user interfaces.
 * @version 0.1.0
 */
import System;
import System.Runtime.InteropServices;
import IWshRuntimeLibrary;
import WbemScripting;

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
/* The watcher module */

/** @class */
var ConversionWatcher = GetConversionWatcherType();

(new ConversionWatcher(param.Markdown)).Start();
Quit(0);

/**
 * @returns the ConversionWatcher type.
 */
function GetConversionWatcherType() {
  /** @class */
  var MessageBox = GetMessageBoxType();

  /**
   * The specified Markdown path argument.
   * @private @type {string}
   */
  var MarkdownPath;
  /**
   * The overwrite prompt text as read from the powershell core console host.
   * @private @type {string}
   */
  var OverwritePromptText;

  /**
   * Separate the polluted characters from the informative message.
   * @private @constant {string}
   */
  var ERROR_MESSAGE_DELIM = '--';

  /**
   * @class @constructs ConversionWatcher
   * @param {string} markdown is the specified Markdown path argument.
   */
  function ConversionWatcher(markdown) {
    MarkdownPath = markdown;
    OverwritePromptText = '';
  }

  /**
   * Execute the runner of the shortcut target script and wait for its exit.
   * @public @memberof ConversionWatcher @instance
   */
  ConversionWatcher.prototype.Start = function() {
    WaitForExit(StartPwshExeWithMarkdown());
    // This method will run only once.
    delete ConversionWatcher.prototype.Start;
  }

  /**
   * Start a PowerShell Core process that runs the shortcut menu target
   * script with the markdown path as the argument.
   * The Try-Catch handles the errors thrown by the process. The Standard Error
   * Stream encoding is not utf-8. For this reason, it surrounds the message with
   * unwanted characters. The error message delimiter constant string separates
   * the informative message from noisy characters.
   * @private
   * @returns {object} the process started by the built command.
   */
  function StartPwshExeWithMarkdown() {
    var shell: WshShell = new WshShellClass();
    try {
      return shell.Exec(Format(
        '"{0}" -nop -ep Bypass -w Hidden -cwa "try{ Import-Module $args[0]; {4} -MarkdownPath $args[1] }' +
        // Get uniform error messages format by handling them in a catch statement.
        'catch { Write-Error (""{1}"" + $_.Exception.Message + ""{1}"") }" "{2}" "{3}"',
        [
          // The HKLM registry subkey stores the PowerShell Core application path.
          shell.RegRead('HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe\\'),
          ERROR_MESSAGE_DELIM,
          ChangeScriptExtension('.psm1'),
          MarkdownPath,
          (new FileSystemObjectClass()).GetBaseName(param.ApplicationPath)
        ]
      ));
    } finally {
      Marshal.FinalReleaseComObject(shell);
      shell = null;
    }
  }

  /**
   * Observe when the child process exits with or without an error.
   * Call the appropriate handler for each outcome.
   * @private
   * @param {object} pwshExe is the PowerShell Core process or child process.
   * @param {number} pwshExe.Status specifies whether the process is running (=0) or has completed (=1).
   * @param {number} pwshExe.ExitCode specifies that the process terminated with an error (!=0) or not (=0).
   * @param {object} pwshExe.StdOut the standard output stream.
   * @param {object} pwshExe.StdErr the standard error stream.
   */
  function WaitForExit(pwshExe) {
    // Wait for the process to complete.
    while (!pwshExe.Status && !pwshExe.ExitCode) {
      HandleOutputDataReceived(pwshExe, pwshExe.StdOut.ReadLine());
    }
    // When the process terminated with an error.
    if (pwshExe.ExitCode) {
      HandleErrorDataReceived(pwshExe.StdErr.ReadAll());
    }
    Marshal.FinalReleaseComObject(pwshExe);
    pwshExe = null;
  }

  /**
   * Show the overwrite prompt that the child process sends.
   * Subsequently, wait for the user's response.
   * Handle the event when the PowerShell Core (child) process
   * redirects output to the parent Standard Output stream.
   * @private
   * @param {object} pwshExe it the sender child process.
   * @param {object} pwshExe.StdIn the standard input stream.
   * @param {string} outData the output text line sent.
   */
  function HandleOutputDataReceived(pwshExe, outData) {
    if (outData.length > 0) {
      // Show the message box when the text line is a question.
      // Otherwise, append the text line to the overall message text variable.
      if (outData.match(/\?\s*$/)) {
        OverwritePromptText += '\n' + outData;
        // Write the user's choice to the child process console host.
        pwshExe.StdIn.WriteLine(MessageBox.Show(OverwritePromptText, MessageBox.WARNING));
        // Optional
        OverwritePromptText = '';
      } else {
        OverwritePromptText += outData + '\n';
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
  function HandleErrorDataReceived(errData) {
    if (errData.length > 0) {
      // Remove the polluted characters from the error message data text.
      var delimIndex = errData.indexOf(ERROR_MESSAGE_DELIM);
      var delimLastIndex = errData.lastIndexOf(ERROR_MESSAGE_DELIM);
      MessageBox.Show(errData.substring(delimIndex+2, delimLastIndex));
    }
  }

  /**
   * Represents the shortcut target script runner watcher.
   * @typedef {object} ConversionWatcher
   */
  return ConversionWatcher;
}

/**
 * @returns the MessageBox type.
 */
function GetMessageBoxType() {
  /** @private @constant {number} */
  var MESSAGE_BOX_TITLE = 'Convert to HTML';
  /** @private @constant {number} */
  var NO_MESSAGE_TIMEOUT = 0;
  /** @private @constant {number} */
  var YESNO_BUTTON = 4;
  /** @private @constant {number} */
  var OK_BUTTON = 0;
  /** @private @constant {number} */
  var YES_POPUPRESULT = 6;
  /** @private @constant {number} */
  var NO_POPUPRESULT = 7;
  /** @private @constant {number} */
  var ERROR_MESSAGE = 16;
  /** @private @constant {number} */
  var WARNING_MESSAGE = 48;

  /**
   * Represents the markdown conversion message box.
   * @typedef {object} MessageBox
   * @property {number} WARNING specifies that the dialog shows a warning message.
   */
  var MessageBox = { };

  /** @public @static @readonly @property {number} */
  MessageBox.WARNING = WARNING_MESSAGE;
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
  MessageBox.Show = function(message, messageType) {
    if (messageType != ERROR_MESSAGE && messageType != WARNING_MESSAGE) {
      messageType = ERROR_MESSAGE;
    }
    // The error message box shows the OK button alone.
    // The warning message box shows the alternative Yes or No buttons.
    messageType += messageType == ERROR_MESSAGE ? OK_BUTTON:YESNO_BUTTON;
    switch ((new WshShellClass()).Popup(message, Object(NO_MESSAGE_TIMEOUT), Object(MESSAGE_BOX_TITLE), Object(messageType))) {
      case YES_POPUPRESULT:
        return 'Yes';
      case NO_POPUPRESULT:
        return 'No';
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
      setInParam('sValue', param.ApplicationPath);
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