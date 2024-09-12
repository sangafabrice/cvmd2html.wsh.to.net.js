import System;
import System.IO;
import System.Text;

package MarkdownToHtml.Shortcut {

  /// <summary>Represents the command line parameters.</summary>
  internal abstract class Param {

    /// <summary>Get the input markdown path.</summary>
    /// <remarks>It is an object as an alternative to nullable string.</remarks>
    public static function get Markdown(): Object {
      return param.Markdown;
    }

    /// <summary>Get the expected output html path.</summary>
    public static function get Html(): String {
      return param.Html;
    }

    /// <summary>Specify to configure the shortcut in the registry.</summary>
    public static function get Set(): Boolean {
      return param.Set;
    }

    /// <summary>Specify to remove the shortcut menu.</summary>
    public static function get Unset(): Boolean {
      return param.Unset;
    }

    /// <summary>Specify to configure the shortcut without the icon.</summary>
    public static function get NoIcon(): Boolean {
      return param.NoIcon;
    }

    /// <summary>Get the input arguments and parameters.</summary>
    /// <param name="args">The command line arguments including the command path.</param>
    public static function Parse(args: String[]): void {
      param.ApplicationPath = (new FileInfo(args[0])).FullName;
      if (args.length == 2) {
        var arg = args[1];
        if (/^\/markdown:[^:]/i.test(arg)) {
          param.Html = GetHtmlPath(param.Markdown = CheckMarkdown(arg.Split(char[]([':']), 2)[1]));
          return;
        }
        switch (arg.toLowerCase()) {
          case '/set':
            param.Set = true;
            param.NoIcon = false;
            return;
          case '/set:noicon':
            param.Set = true;
            param.NoIcon = true;
            return;
          case '/unset':
            param.Unset = true;
            return;
        }
      } else if (args.length == 1) {
        param.Set = true;
        param.NoIcon = false;
        return;
      }
      ShowHelp();
    }

    /// <summary>Represents the application information from the command line.</summary>
    public static abstract class App {
      
      /// <summary>Get the application path.</summary>
      public static function get Path(): String {
        return param.ApplicationPath;
      }

      /// <summary>Get the application path with a different extension.</summary>
      /// <param name="extension">The new extension.</param>
      /// <returns>The application path with the new extension.</returns>
      public static function ChangePathExtension(extension: String): String {
        return System.IO.Path.ChangeExtension(param.ApplicationPath, extension);
      }
    }

    /// <summary>The parameter object.</summary>
    static var param: Object = { };

    /// <summary>Validate the input markdown path string.</summary>
    /// <param name="markdown">The input markdown file path.</param>
    /// <returns>The markdown file path if it is valid, or quit otherwise.</returns>
    static function CheckMarkdown(markdown: String): String {
      try {
        if (String.Compare(String(Path.GetExtension(markdown)), '.md', true)) {
          throw new Exception(String.Format('"{0}" is not a markdown (.md) file.', markdown));
        }
        if (!File.Exists(markdown)) {
          throw new FileNotFoundException(null, markdown);
        }
      } catch (error: Exception) {
        MessageBox.Show(error.Message);
      }
      return markdown;
    }

    /// <summary>
    /// Get the output html path when it is unique
    /// without prompts or when the user accepts to overwrite an
    /// existing HTML file. Otherwise, it exits the application.
    /// </summary>
    /// <param name="markdown">The input markdown file path.</param>
    /// <returns>The output html path.</returns>
    static function GetHtmlPath(markdown: String): String {
      var htmlPath = Path.ChangeExtension(markdown, '.html');
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

    /// <summary>Show help and quit.</summary>
    static function ShowHelp(): void {
      var helpTextBuilder: StringBuilder = new StringBuilder();
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
      MessageBox.Show(helpTextBuilder, MessageBox.HELP);
      Command.Quit(1);
    }
  }
}