/**
 * @file Launches the shortcut target PowerShell script with the selected markdown as an argument.
 * @version 0.0.1.9
 */

// #region: header of utils.js
// Constants and variables.

/** @constant */
var BUTTONS_OKONLY = 0;
/** @constant */
var BUTTONS_YESNO = 4;
/** @constant */
var POPUP_ERROR = 16;
/** @constant */
var POPUP_NORMAL = 0;
/** @constant */
var POPUP_WARNING = 48;

/** @typedef */
var FileSystemObject = new ActiveXObject('Scripting.FileSystemObject');
/** @typedef */
var WshShell = new ActiveXObject('WScript.Shell');

var ScriptRoot = FileSystemObject.GetParentFolderName(WSH.ScriptFullName)

// #endregion

// #region: main
// The main part of the program.

/** @typedef */
var Package = getPackage();
/** @typedef */
var Param = getParameters();

/** The application execution. */
if (Param.Markdown) {
  // #region: process.js
  // Process type definition.

  var Process = (function() {
    /** @constructor */
    function ConsoleHost() { }

    /**
     * Execute the runner of the shortcut target script and wait for its exit.
     * @param {string} commandLine is the command to start.
     */
    ConsoleHost.Start = function(commandLine) {
      WaitForExit(WshShell.Exec(commandLine));
    }

    /**
     * Represents the data obtained from the console host.
     * @private @constructs ConsoleData
     */
    function ConsoleData() {
      /** The expected prompt from the console host. */
      this.OverwritePromptText = '';
    }

    /**
     * Show the overwrite prompt that the child process sends. Handle the event when the
     * PowerShell Core (child) process redirects output to the parent Standard Output stream.
     * @param {WshScriptExec} pwshExe it the sender child process.
     * @param {string} outData the output text line sent.
     */
    ConsoleData.prototype.HandleOutputDataReceived = function (pwshExe, outData) {
      if (outData.length) {
        // Show the message box when the text line is a question.
        // Otherwise, append the text line to the overall message text variable.
        if (outData.match(/\?\s*$/)) {
          this.OverwritePromptText += '\n' + outData;
          // Write the user's choice to the child process console host.
          pwshExe.StdIn.WriteLine(popup(this.OverwritePromptText, POPUP_WARNING, BUTTONS_YESNO));
          this.OverwritePromptText = '';
        } else {
          this.OverwritePromptText += outData + '\n';
        }
      }
    }

    /**
     * Observe when the child process exits with or without an error.
     * Call the appropriate handler for each outcome.
     * @private
     * @param {WshScriptExec} pwshExe is the PowerShell Core process or child process.
     */
    function WaitForExit(pwshExe) {
      var conhostData = new ConsoleData();
      // Wait for the process to complete.
      while (!pwshExe.Status && !pwshExe.ExitCode) {
        conhostData.HandleOutputDataReceived(pwshExe, pwshExe.StdOut.ReadLine());
      }
      // When the process terminated with an error.
      if (pwshExe.ExitCode) {
        HandleErrorDataReceived(pwshExe.StdErr.ReadAll());
      }
    }

    /**
     * Show the error message that the child process writes on the console host.
     * @private
     * @param {string} errData the error message text.
     */
    function HandleErrorDataReceived(errData) {
      if (errData.length) {
        // Remove the ANSI color tag characters from the error message data text.
        errData = errData.replace(/(\x1B\[31;1m)|(\x1B\[0m)/g, '');
        popup(errData.substring(errData.indexOf(':') + 2), POPUP_ERROR);
      }
    }

    return ConsoleHost;
  })();

  // #endregion

  /** @constant */
  var CMD_LINE_FORMAT = '"{0}" -nop -ep Bypass -w Hidden -cwa "' +
    'try {' +
      'Import-Module $args[0];' +
      'cvmd2html -MarkdownPath $args[1]' +
    '} catch {' +
      'Write-Error $_.Exception.Message' +
    '}" "{1}" "{2}"';
  Process.Start(format(CMD_LINE_FORMAT, Package.PwshExePath, Package.PwshScriptPath, Param.Markdown));
  quit(0);
}

/** Configuration and settings. */
if (Param.Set ^ Param.Unset) {
  // #region: setup.js
  // Methods for managing the shortcut menu option: install and uninstall.

  /** @typedef */
  var Setup = (function() {
    var HKCU = 0x80000001;
    var VERB_KEY = 'SOFTWARE\\Classes\\SystemFileAssociations\\.md\\shell\\cthtml';
    var ICON_VALUENAME = 'Icon';

    return {
      /** Configure the shortcut menu in the registry. */
      Set: function () {
        var COMMAND_KEY = VERB_KEY + '\\command';
        var command = format('{0} "{1}" /Markdown:"%1"', WSH.FullName.replace(/\\cscript\.exe$/i, '\\wscript.exe'), WSH.ScriptFullName);
        StdRegProv.CreateKey(HKCU, COMMAND_KEY);
        StdRegProv.SetStringValue(HKCU, COMMAND_KEY, null, command);
        StdRegProv.SetStringValue(HKCU, VERB_KEY, null, 'Convert to &HTML');
      },

      /**
       * Add an icon to the shortcut menu in the registry.
       * @param {string} menuIconPath is the shortcut menu icon file path.
       */
      AddIcon: function (menuIconPath) {
        StdRegProv.SetStringValue(HKCU, VERB_KEY, ICON_VALUENAME, menuIconPath);
      },

      /** Remove the shortcut icon menu. */
      RemoveIcon: function () {
        StdRegProv.DeleteValue(HKCU, VERB_KEY, ICON_VALUENAME);
      },

      /** Remove the shortcut menu by removing the verb key and subkeys. */
      Unset: function () {
        var stdRegProvMethods = StdRegProv.Methods_;
        var enumKeyMethod = stdRegProvMethods('EnumKey');
        var enumKeyMethodParams = enumKeyMethod.InParameters;
        var inParams = enumKeyMethodParams.SpawnInstance_();
        inParams.hDefKey = HKCU;
        // Recursion is used because a key with subkeys cannot be deleted.
        // Recursion helps removing the leaf keys first.
        (function(key) {
          inParams.sSubKeyName = key;
          var outParams = StdRegProv.ExecMethod_(enumKeyMethod.Name, inParams);
          var sNames = outParams.sNames;
          outParams = null;
          if (sNames != null) {
            var sNamesArray = sNames.toArray();
            for (var index = 0; index < sNamesArray.length; index++) {
              arguments.callee(format('{0}\\{1}', key, sNamesArray[index]));
            }
          }
          StdRegProv.DeleteKey(HKCU, key);
        })(VERB_KEY);
        inParams = null;
        enumKeyMethodParams = null;
        enumKeyMethod = null;
        stdRegProvMethods = null;
      }
    }
  })();

  // #endregion

  /** @typedef */
  var StdRegProv = GetObject('winmgmts:StdRegProv');

  if (Param.Set) {
    Setup.Set();
    if (Param.NoIcon) {
      Setup.RemoveIcon();
    } else {
      Setup.AddIcon(Package.MenuIconPath);
    }
  } else if (Param.Unset) {
    Setup.Unset();
  }

  StdRegProv = null;

  quit(0);
}

quit(1);

// #endregion

// #region: utils.js
// Utility functions.

/**
 * Generate a random file path.
 * @param {string} extension is the file extension.
 * @returns {string} a random file path.
 */
function generateRandomPath(extension) {
  var typeLib = new ActiveXObject('Scriptlet.TypeLib');
  try {
    return FileSystemObject.BuildPath(WshShell.ExpandEnvironmentStrings('%TEMP%'), typeLib.Guid.substr(1, 36).toLowerCase() + '.tmp' + extension);
  } finally {
    typeLib = null;
  }
}

/**
 * Delete the specified file.
 * @param {string} filePath is the file path.
 */
function deleteFile(filePath) {
  try {
    FileSystemObject.DeleteFile(filePath);
  } catch (e) { }
}

/**
 * Show the application message box.
 * @param {string} messageText is the message text to show.
 * @param {number} popupType[POPUP_NORMAL] is the type of popup box.
 * @param {number} popupButtons[BUTTONS_OKONLY] are the buttons of the message box.
 */
function popup(messageText, popupType, popupButtons) {
  /** @constant */
  var WINDOW_STYLE_HIDDEN = 0;
  /** @constant */
  var WAIT_ON_RETURN = true;
  if (!popupType) {
    popupType = POPUP_NORMAL;
  }
  if (!popupButtons) {
    popupButtons = BUTTONS_OKONLY;
  }

  // #region: answerLog.js
  // Manage the answer log file and content.

  /** @typedef */
  var AnswerLog = {
    /** The answer log file path. */
    Path: generateRandomPath('.log'),

    /** Return the content of the answer log file. */
    Read: function () {
      /** @constant */
      var FOR_READING = 1;
      try {
        var txtStream = FileSystemObject.OpenTextFile(this.Path, FOR_READING);
        return txtStream.ReadLine();
      } catch (e) { }
      finally {
        if (txtStream) {
          txtStream.Close();
          txtStream = null;
        }
      }
    },

    /** Delete the answer log file. */
    Delete: function () {
      deleteFile(this.Path)
    }
  }

  // #endregion

  /** @constant */
  var CMD_LINE_FORMAT = 'C:\\Windows\\System32\\cmd.exe /d /c ""{0}" """{1}""" {2} {3} > "{4}""';
  Package.MessageBoxLink.Create();
  WshShell.Run(format(CMD_LINE_FORMAT, Package.MessageBoxLink.Path, messageText.replace(/"/g, "*").replace(/\n/g, "^"), popupButtons, popupType, AnswerLog.Path), WINDOW_STYLE_HIDDEN, WAIT_ON_RETURN);
  Package.MessageBoxLink.Delete();
  try {
    return AnswerLog.Read();
  } finally {
    AnswerLog.Delete();
  }
}

/**
 * Replace the format item "{n}" by the nth input in a list of arguments.
 * @param {string} formatStr the pattern format.
 * @param {...string} args the replacement texts.
 * @returns {string} a copy of format with the format items replaced by args.
 */
function format(formatStr, args) {
  args = Array.prototype.slice.call(arguments).slice(1);
  while (args.length > 0) {
    formatStr = formatStr.replace(new RegExp('\\{' + (args.length - 1) + '\\}', 'g'), args.pop());
  }
  return formatStr;
}

/** Destroy the COM objects. */
function dispose() {
  WshShell = null;
  FileSystemObject = null;
}

/**
 * Clean up and quit.
 * @param {number} exitCode .
 */
function quit(exitCode) {
  dispose();
  WSH.Quit(exitCode);
}

// #endregion

// #region: package.js
// Information about the resource files used by the project.

/** Get the package type. */
function getPackage() {
  /** @constant */
  var POWERSHELL_SUBKEY = 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe\\';
  /** The project resources directory path. */
  var resourcePath = FileSystemObject.BuildPath(ScriptRoot, 'rsc');

  var pwshExePath = WshShell.RegRead(POWERSHELL_SUBKEY);
  var pwshScriptPath = FileSystemObject.BuildPath(ScriptRoot, 'cvmd2html.psd1');
  var messageBoxScriptPath = FileSystemObject.BuildPath(resourcePath, 'messageBox.ps1');
  var menuIconPath = FileSystemObject.BuildPath(resourcePath, 'menu.ico');

  /** @constructor @abstract */
  function ShortcutLink() {
    /** The shortcut link path. */
    this.Path = generateRandomPath('.lnk');
    /** Abstract method for creating a shortcut link. */
    this.Create = function () { };
  }

  /** Delete the custom icon link file. */
  ShortcutLink.prototype.Delete = function() {
    deleteFile(this.Path);
  }

  /**
   * Factory method for creating a ShortcutLink.
   * @param {Function} func
   * @returns {ShortcutLink}
   */
  var createShortcutLink = function (func) {
    var link = new ShortcutLink();
    link.Create = func;
    return link;
  }

  /**
   * Set the Create link method.
   * @param {string} linkArguments
   * @returns {function(): void}
   */
  var createLinkFunction = function(linkArguments) {
    return function() {
      var link = WshShell.CreateShortcut(this.Path);
      link.TargetPath = pwshExePath;
      link.Arguments = linkArguments;
      link.IconLocation = menuIconPath;
      link.Save();
      link = null;
    }
  }

  return {
    /** The powershell core runtime path. */
    PwshExePath: pwshExePath,
    /** The shortcut target powershell script path. */
    PwshScriptPath: pwshScriptPath,
    /** The shortcut menu icon path. */
    MenuIconPath: menuIconPath,

    /** Represents an adapted link object. */
    MessageBoxLink: createShortcutLink(createLinkFunction(format('-nol -noni -nop -NoProfileLoadTime -f "{0}"', messageBoxScriptPath)))
  }
}

// #endregion

// #region: parameters.js
// Parsed parameters.

/**
 * @typedef {object} ParamHash
 * @property {string} Markdown is the selected markdown file path.
 * @property {boolean} Set installs the shortcut menu.
 * @property {boolean} NoIcon installs the shortcut menu without icon.
 * @property {boolean} Unset uninstalls the shortcut menu.
 * @property {boolean} RunLink runs the icon custom shortcut link.
 * @property {boolean} Help shows help.
 */

/** @returns {ParamHash} */
function getParameters() {
  var WshArguments = WSH.Arguments;
  var WshNamed = WshArguments.Named;
  var paramCount = WshArguments.Count();
  if (paramCount == 1) {
    var paramMarkdown = WshNamed('Markdown');
    if (WshNamed.Exists('Markdown') && paramMarkdown && paramMarkdown.length) {
      return {
        Markdown: paramMarkdown
      }
    }
    var param = { Set: WshNamed.Exists('Set') };
    if (param.Set) {
      var noIconParam = WshNamed('Set');
      var isNoIconParam = false;
      param.NoIcon = noIconParam && (isNoIconParam = /^NoIcon$/i.test(noIconParam));
      if (noIconParam == undefined || isNoIconParam) {
        return param;
      }
    }
    param = { Unset: WshNamed.Exists('Unset') };
    if (param.Unset && WshNamed('Unset') == undefined) {
      return param;
    }
    return {
      Markdown: WshArguments(0)
    }
  } else if (paramCount == 0) {
    return {
      Set: true,
      NoIcon: false
    }
  }
  var helpText = '';
  helpText += 'The MarkdownToHtml shortcut launcher.\n';
  helpText += 'It starts the shortcut menu target script in a hidden window.\n\n';
  helpText += 'Syntax:\n';
  helpText += '  Convert-MarkdownToHtml.js /Markdown:<markdown file path>\n';
  helpText += '  Convert-MarkdownToHtml.js [/Set[:NoIcon]]\n';
  helpText += '  Convert-MarkdownToHtml.js /Unset\n';
  helpText += '  Convert-MarkdownToHtml.js /Help\n\n';
  helpText += "<markdown file path>  The selected markdown's file path.\n";
  helpText += '                 Set  Configure the shortcut menu in the registry.\n';
  helpText += '              NoIcon  Specifies that the icon is not configured.\n';
  helpText += '               Unset  Removes the shortcut menu.\n';
  helpText += '                Help  Show the help doc.\n';
  popup(helpText);
  quit(1);
}

// #endregion