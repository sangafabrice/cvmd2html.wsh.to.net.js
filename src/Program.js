import System;

package MarkdownToHtml.Shortcut {

  abstract class Program {
    
    static function Main(args: String[]): void {
      Param.Parse(args);
      if (Param.Markdown) {
        Command.ConvertToHtml();
      } else if (Param.Set) {
        Command.Configuration.Set();
        if (Param.NoIcon) {
          Command.Configuration.RemoveIcon();
        } else {
          Command.Configuration.AddIcon();
        }
      } else if (Param.Unset) {
        Command.Configuration.Unset();
      }
      Command.Quit(0);
    }
  }
}