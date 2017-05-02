
  if (!String.prototype.format) {
    String.prototype.format = function() {
      var args = arguments;
      return this.replace(/{(\d+)}/g, function(match, number) { 
        return typeof args[number] != 'undefined'
          ? args[number]
          : match
        ;
      });
    };
  }

  if (!Array.prototype.contains) {
    Array.prototype.contains = function(x) {
      for (var i=0; i<this.length; i++) {
        if (this[i]==x) return true;
      }
      return false;
    };
  }

  if (!Array.prototype.remove) {
    Array.prototype.remove = function(x) {
      var result = [];
      for (var i=0; i<this.length; i++) {
        if (this[i]==x) continue;
        result.push(this[i]);
      }
      return result;
    };
  }

  if (!Array.prototype.removeAll) {
    Array.prototype.removeAll = function(a) {
      var result = [];
      for (var i=0; i<this.length; i++) {
        if (a.contains(this[i])) continue;
        result.push(this[i]);
      }
      return result;
    };
  }

  function getScriptCurrentDirPath() {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    return fso.getParentFolderName(WScript.ScriptFullName);
  }

  function echo(msg) {
    WScript.Echo(msg);
  }

  function msgbox(msg) {
    var shell = new ActiveXObject("WScript.Shell");
    shell.Popup(msg, 0, "Windows Script Host", 0);
  }

  function make2digits(n) {
    return ("0" + n).slice(-2);
  }

  function formatMonth(dt, sep) {
    var year  = dt.getFullYear();
    var month = make2digits(dt.getMonth() + 1);
    return "" + year + sep + month;
  }

  function readTextFromFile_Utf8(path) {
    var StreamTypeEnum    = { adTypeBinary: 1, adTypeText: 2 };
    var SaveOptionsEnum   = { adSaveCreateNotExist: 1, adSaveCreateOverWrite: 2 };
    var LineSeparatorEnum = { adLF: 10, adCR: 13, adCRLF: -1 };
    var StreamReadEnum    = { adReadAll: -1, adReadLine: -2 };
    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type = StreamTypeEnum.adTypeText;
    stream.Charset = "utf-8";
    stream.LineSeparator = LineSeparatorEnum.adLF;
    stream.Open();
    stream.LoadFromFile(path);
    var result = stream.ReadText(StreamReadEnum.adReadAll);
    stream.Close();
    return result;
  }

  function readLinesFromFile_Utf8(path) {
    var StreamTypeEnum    = { adTypeBinary: 1, adTypeText: 2 };
    var SaveOptionsEnum   = { adSaveCreateNotExist: 1, adSaveCreateOverWrite: 2 };
    var LineSeparatorEnum = { adLF: 10, adCR: 13, adCRLF: -1 };
    var StreamReadEnum    = { adReadAll: -1, adReadLine: -2 };
    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type = StreamTypeEnum.adTypeText;
    stream.Charset = "utf-8";
    stream.LineSeparator = LineSeparatorEnum.adLF;
    var result = "";
    stream.Open();
    stream.LoadFromFile(path);
    var count = 0;
    while (!stream.EOS) {
      result += stream.ReadText(StreamReadEnum.adReadLine) + "\n";
      count++;
      if (count > 0 && (count % 100000) == 0) {
        WScript.Echo("readLinesFromFile_Utf8({0}): {1}".format(path, count));
      }
    }
    if (count >= 100000) WScript.Echo("readLinesFromFile_Utf8({0}): {1}...Done!".format(path, count));

    stream.Close();
    return result;
  }

  function writeTextToFile_Utf8_NoBOM(path, text) {
    var StreamTypeEnum  = { adTypeBinary: 1, adTypeText: 2 };
    var SaveOptionsEnum = { adSaveCreateNotExist: 1, adSaveCreateOverWrite: 2 };
    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type = StreamTypeEnum.adTypeText;
    stream.Charset = "utf-8";
    stream.Open();
    stream.WriteText(text);
    stream.Position = 0
    stream.Type = StreamTypeEnum.adTypeBinary;
    stream.Position = 3
    var buf = stream.Read();
    stream.Position = 0
    stream.Write(buf);
    stream.SetEOS();
    stream.SaveToFile(path, SaveOptionsEnum.adSaveCreateOverWrite);
    stream.Close();
  }

  function readVectorFromFile_Utf8(path) {
    var json = readLinesFromFile_Utf8(path);
    var vec = JSON.parse(json);
    var result = [];
    for (var i=0; i<vec.length; i++) {
      result.push(vec[i]);
    }
    return result;
  }

  function writeVectorToFile_Utf8(path, vec) {
    var dcopy = JSON.parse(JSON.stringify(vec));
    for (var propName in dcopy) {
      if (propName != "length" && !/^\d+$/.exec(propName)) {
        delete dcopy[propName];
      }
    }
    var json = JSON.stringify(dcopy, null, 2);
    writeTextToFile_Utf8_NoBOM(path, json);
  }

