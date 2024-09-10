/**
 * @file Converts markdown file content to html document.
 * @version 1.0.0.2
 */
import System;
import System.Runtime.InteropServices;
import System.Resources;
import System.Reflection;
import Microsoft.VisualBasic.FileIO;
import Microsoft.VisualBasic;
import mshtml;
import ROOT.CIMV2;

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
/** @constant {regexp} */
var MARKDOWN_REGEX = /\.md$/i;
CheckMarkdown();
ConvertTo(GetHtmlPath());
Quit(0);

/**
 * Validate the input markdown path string.
 */
function CheckMarkdown() {
  if (!MARKDOWN_REGEX.test(param.Markdown)) {
    MessageBox.Show(Format('"{0}" is not a markdown (.md) file.', param.Markdown));
  }
  if (!FileSystem.FileExists(param.Markdown)) {
    MessageBox.Show(Format('"{0}" cannot be found.', param.Markdown));
  }
}

/**
 * Convert the content of the markdown file to html.
 * @param {string} htmlPath is the output html path.
 */
function ConvertTo(htmlPath) {
  SetHtmlContent(htmlPath, ConvertFrom(FileSystem.ReadAllText(param.Markdown)));
}

/**
 * Write the html text to the output HTML file.
 * It notifies the user when the operation did not complete with success.
 * @param {string} htmlPath is the output html path.
 * @param {string} htmlContent is the content of the html file.
 */
function SetHtmlContent(htmlPath, htmlContent) {
  try {
    FileSystem.WriteAllText(htmlPath, htmlContent, false);
  } catch (error) {
    if (error.number == -2146823266) {
      MessageBox.Show(Format('Access to the path "{0}" is denied.', htmlPath));
    } else {
      MessageBox.Show(Format('Unspecified error trying to write to "{0}".', htmlPath));
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
  var htmlPath = param.Markdown.replace(MARKDOWN_REGEX, '.html');
  if (FileSystem.FileExists(htmlPath)) {
    MessageBox.Show(
      Format('The file "{0}" already exists.\n\nDo you want to overwrite it?', htmlPath),
      MessageBox.WARNING
    );
  } else if (FileSystem.DirectoryExists(htmlPath)) {
    MessageBox.Show(Format('"{0}" cannot be overwritten because it is a directory.', htmlPath));
  }
  return htmlPath;
}

/**
 * Convert a markdown content to an html document.
 * @param {string} mardownContent is the content to convert.
 * @returns {string} the output html document content. 
 */
function ConvertFrom(markdownContent) {
  // Build the HTML document that will load the showdown.js library.
  var document: HTMLDocumentClass = new HTMLDocumentClass();
  document.open();
  document.IHTMLDocument2_write(
    new ResourceManager('Resource', Assembly.GetExecutingAssembly())
    .GetString('LoaderHtml')
  );
  document.body.innerHTML = markdownContent;
  document.parentWindow.execScript('convertMarkdown()', 'javascript');
  try {
    return document.body.innerHTML;
  } finally {
    if (document != undefined) {
      document.close();
      Marshal.FinalReleaseComObject(document);
      document = null;
    }
  }
}

/**
 * @returns the MessageBox type.
 */
function GetMessageBoxType() {
  /** @private @constant {number} */
  var MESSAGE_BOX_TITLE = 'Convert to HTML';

  /**
   * Represents the markdown conversion message box.
   * @typedef {object} MessageBox
   * @property {number} WARNING specifies that the dialog shows a warning message.
   */
  var MessageBox = { };

  /** @public @static @readonly @property {number} */
  MessageBox.WARNING = MsgBoxStyle.Exclamation;
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
    if (messageType != MsgBoxStyle.Critical && messageType != MsgBoxStyle.Exclamation) {
      messageType = MsgBoxStyle.Critical;
    }
    // The error message box shows the OK button alone.
    // The warning message box shows the alternative Yes or No buttons.
    messageType += messageType == MsgBoxStyle.Critical ? MsgBoxStyle.OkOnly:MsgBoxStyle.YesNo;
    if (
      Strings.Filter(
        String[]([' OK ', ' No ']),
        Format(' {0} ', Interaction.MsgBox(String(message), messageType, MESSAGE_BOX_TITLE)),
        true,
        CompareMethod.Text
      ).length > 0
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
  var HKCU: uint = 0x80000001;
  var VERB_KEY = 'SOFTWARE\\Classes\\SystemFileAssociations\\.md\\shell\\cthtml';
  if (param.Set) {
    // Configure the shortcut menu in the registry.
    var COMMAND_KEY = VERB_KEY + '\\command';
    var command = Format('"{0}" /Markdown:"%1"', param.ApplicationPath);
    StdRegProv.CreateKey(HKCU, COMMAND_KEY);
    StdRegProv.SetStringValue(HKCU, COMMAND_KEY, null, command);
    StdRegProv.SetStringValue(HKCU, VERB_KEY, null, 'Convert to &HTML');
    var iconValueName = 'Icon';
    if (param.NoIcon) {
      StdRegProv.DeleteValue(HKCU, VERB_KEY, iconValueName);
    } else {
      StdRegProv.SetStringValue(HKCU, VERB_KEY, iconValueName, param.ApplicationPath);
    }
  } else if (param.Unset) {
    // Remove the shortcut menu.
    // Remove the verb key and subkeys.
    // Recursion is used because a key with subkeys cannot be deleted.
    // Recursion helps removing the leaf keys first.
    var deleteVerbKey = function(key) {
      var recursive = function func(key) {
        var sNames = StdRegProv.EnumKey(HKCU, key);
        if (sNames != null) {
          for (var index = 0; index < sNames.length; index++) {
            func(key + '\\' + sNames[index]);
          }
        }
        StdRegProv.DeleteKey(HKCU, key);
      };
      recursive(key);
    }
    deleteVerbKey(VERB_KEY);
    deleteVerbKey = null;
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