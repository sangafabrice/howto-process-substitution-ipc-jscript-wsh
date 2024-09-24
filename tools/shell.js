/**
 * @file library for jsdoc.
 * @version 0.0.1.0
 */

class WshShortcut {
  constructor() {
    this.TargetPath = "";
    this.Arguments = "";
    this.IconLocation = "";
  }

  Save() { throw new Error('Not implemented yet.') }
}

class WshScriptExec {
  constructor() {
    this.ExitCode = 0;
    this.ProcessID = 0;
    this.Status = 0;
    this.StdErr = new TextStream();
    this.StdIn = new TextStream();
    this.StdOut = new TextStream();
  }

  Terminate() { throw new Error('Not implemented yet.') }
}


class WshShell {
  constructor() { }

  /**
   * @param {string} regPath
   * @returns {string}
   */
  static RegRead(regPath) { throw new Error('Not implemented yet.') }

  /**
   * @param {string} regPath
   * @param {string} value
   */
  static RegWrite(regPath, value) { throw new Error('Not implemented yet.') }

  /**
   * @param {string} regPath
   */
  static RegDelete(regPath) { throw new Error('Not implemented yet.') }

  /**
   * @param {string} prompt
   * @param {number} secondsToWait
   * @param {string} title
   * @param {number} type
   * @returns {number}
   */
  static Popup(prompt, secondsToWait, title, type) { throw new Error('Not implemented yet.') }

  /**
   * @param {string} pathName
   * @returns {WshShortcut}
   */
  static CreateShortcut(pathName) { throw new Error('Not implemented yet.') }

  /**
   * @param {string} envVariable
   * @returns {string}
   */
  static ExpandEnvironmentStrings(envVariable) { throw new Error('Not implemented yet.') }

  /**
   * @param {string} command
   * @returns {WshScriptExec}
   */
  static Exec(command) { throw new Error('Not implemented yet.') }

  /**
   * @param {string} command
   * @param {number} windowStyle
   * @param {boolean} waitOnReturn
   * @returns {number}
   */
  static Run(command, windowStyle, waitOnReturn) { throw new Error('Not implemented yet.') }

  /**
   * @param {string} name
   * @returns {string}
   */
  static SpecialFolders(name) { throw new Error('Not implemented yet.') }
}

class FileSystemObject {
  constructor() { }

  /**
   * @param {string} path
   * @param {string} name
   * @returns {string}
   */
  static BuildPath(path, name) { throw new Error('Not implemented yet.') }

  /**
   * @param {string} path
   * @returns {string}
   */
  static GetParentFolderName(path) { throw new Error('Not implemented yet.') }

  /**
   * @param {string} path
   */
  static DeleteFile(path) { throw new Error('Not implemented yet.') }

  /**
   * @param {string} filename
   * @param {string} iomode
   * @returns {TextStream}
   */
  static OpenTextFile(filename, iomode) { throw new Error('Not implemented yet.') }

  /**
   * @param {string} filename
   * @returns {TextStream}
   */
  static CreateTextFile(filename) { throw new Error('Not implemented yet.') }
}

class TextStream {
  constructor() { }

  /**
   * @returns {string}
   */
  ReadAll() { throw new Error('Not implemented yet.') }

  /**
   * @returns {string}
   */
  ReadLine() { throw new Error('Not implemented yet.') }

  /**
   * @param {string} text
   */
  Write(text) { throw new Error('Not implemented yet.') }

  /**
   * @param {string} text
   */
  WriteLine(text) { throw new Error('Not implemented yet.') }

  Close() { throw new Error('Not implemented yet.') }
}