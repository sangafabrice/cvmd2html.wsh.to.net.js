import System;
import System.IO;
import Microsoft.Win32;
import Markdig;

package MarkdownToHtml.Shortcut {

  /// <summary>Represents the application actions.</summary>
  internal abstract class Command {

    /// <summary>Write the html text to the output HTML file.</summary>
    /// <remarks>It notifies the user when the operation did not complete with success.</remarks>
    public static function ConvertToHtml(): void {
      try {
        File.WriteAllText(Param.Html, Markdown.ToHtml(Markdown.Normalize(File.ReadAllText(Param.Markdown))));
      } catch (error: Exception) {
        MessageBox.Show(error.Message);
      }
    }

    public static abstract class Configuration {

      public static function Set(): void {
        var COMMAND_KEY = VERB_KEY + '\\command';
        var command = String.Format('"{0}" /Markdown:"%1"', Param.App.Path);
        Registry.SetValue(COMMAND_KEY, '', command);
        Registry.SetValue(VERB_KEY, '', 'Convert to &HTML');
      }

      public static function Unset(): void {
        var SHELL_KEY_OBJ: RegistryKey = HKCU.CreateSubKey(SHELL_SUBKEY);
        if (SHELL_KEY_OBJ) {
          SHELL_KEY_OBJ.DeleteSubKeyTree(VERB);
          SHELL_KEY_OBJ.Close();
        }
      }

      public static function AddIcon(): void {
        Registry.SetValue(VERB_KEY, iconValueName, Param.App.Path);
      }

      public static function RemoveIcon(): void {
        var VERB_KEY_OBJ: RegistryKey = HKCU.CreateSubKey(VERB_SUBKEY);
        if (VERB_KEY_OBJ) {
          VERB_KEY_OBJ.DeleteValue(iconValueName);
          VERB_KEY_OBJ.Close();
        }
      }

      static var HKCU: RegistryKey = Registry.CurrentUser;
      static var SHELL_SUBKEY: String = 'SOFTWARE\\Classes\\SystemFileAssociations\\.md\\shell';
      static var VERB: String = 'cthtml';
      static var iconValueName: String = 'Icon';
      static var VERB_SUBKEY = String.Format('{0}\\{1}', SHELL_SUBKEY, VERB);
      static var VERB_KEY = String.Format('{0}\\{1}', HKCU, VERB_SUBKEY);
    }

    /// <summary>Clean up and quit.</summary>
    public static function Quit(exitCode: int): void {
      GC.Collect();
      Environment.Exit(exitCode);
    }
  }
}