var NScript = (function () {
  var shell = new ActiveXObject("WScript.Shell");
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  var wmi = null;
  try { wmi = GetObject("winmgmts:\\\\.\\root\\cimv2"); } catch (e) { wmi = null; }

  function hasOwn(obj, key) { return Object.prototype.hasOwnProperty.call(obj, key); }

  function scriptFile() { try { return String(WScript.ScriptFullName); } catch (e) { return ""; } }
  function scriptDir() { var p = scriptFile(); if (!p) return cwd(); try { return fso.GetParentFolderName(p); } catch (e) { return cwd(); } }

  function cwd() { return shell.CurrentDirectory; }
  function chdir(path) { shell.CurrentDirectory = normalizePath(path); }

  function isAbs(path) {
    if (!path) return false;
    var p = String(path);
    if (p.length >= 2 && p.charAt(1) === ":") return true;
    if (p.length >= 2 && p.charAt(0) === "\\" && p.charAt(1) === "\\") return true;
    return false;
  }

  function normalizePath(path) {
    if (path == null) return "";
    var p = String(path).replace(/\//g, "\\");
    p = p.replace(/\\{2,}/g, "\\");
    return p;
  }

  function trimSlashEnd(p) {
    while (p.length > 1 && (p.charAt(p.length - 1) === "\\" || p.charAt(p.length - 1) === "/")) p = p.substring(0, p.length - 1);
    return p;
  }

  function joinPath(a, b) {
    a = normalizePath(a || "");
    b = normalizePath(b || "");
    if (!a) return b;
    if (!b) return a;
    a = trimSlashEnd(a);
    if (b.charAt(0) === "\\" || b.charAt(0) === "/") return a + b;
    return a + "\\" + b;
  }

  function dirname(path) {
    path = normalizePath(path || "");
    if (!path) return "";
    try { return fso.GetParentFolderName(path); } catch (e) {}
    var i = path.lastIndexOf("\\");
    if (i <= 0) return "";
    return path.substring(0, i);
  }

  function basename(path, ext) {
    path = normalizePath(path || "");
    var i = path.lastIndexOf("\\");
    var name = i >= 0 ? path.substring(i + 1) : path;
    if (ext && name.length >= ext.length && name.substring(name.length - ext.length) === ext) return name.substring(0, name.length - ext.length);
    return name;
  }

  function extname(path) {
    path = normalizePath(path || "");
    var i = path.lastIndexOf(".");
    var j = path.lastIndexOf("\\");
    if (i < 0 || i < j) return "";
    return path.substring(i);
  }

  function resolvePath(baseDir, target) {
    var t = normalizePath(target || "");
    if (!t) return normalizePath(baseDir || "");
    if (isAbs(t)) return t;
    var b = normalizePath(baseDir || "");
    if (!b) b = cwd();
    return normalizePath(joinPath(b, t));
  }

  function isLikelyPath(request) {
    if (request == null) return false;
    var r = String(request);
    if (!r) return false;
    if (r.indexOf("\\") >= 0 || r.indexOf("/") >= 0) return true;
    if (r.length >= 2 && r.charAt(1) === ":") return true;
    if (r.charAt(0) === ".") return true;
    return false;
  }

  function fileExists(path) { try { return fso.FileExists(normalizePath(path)); } catch (e) { return false; } }
  function dirExists(path) { try { return fso.FolderExists(normalizePath(path)); } catch (e) { return false; } }

  function ensureDir(path) {
    var p = normalizePath(path);
    if (!p) return;
    if (dirExists(p)) return;
    var parts = p.split("\\");
    var cur = "";
    for (var i = 0; i < parts.length; i++) {
      var part = parts[i];
      if (part === "" && i === 0) { cur = "\\"; continue; }
      if (!part || part === ".") continue;
      if (cur === "\\" || cur === "") cur = cur + part;
      else cur = cur + "\\" + part;
      if (cur.length === 2 && cur.charAt(1) === ":") continue;
      if (!dirExists(cur)) {
        try { fso.CreateFolder(cur); } catch (e) {}
      }
    }
  }

  function readFile(path) {
    var p = normalizePath(path);
    var stream = fso.OpenTextFile(p, 1);
    var text = stream.ReadAll();
    stream.Close();
    return text;
  }

  function writeFile(path, data) {
    var p = normalizePath(path);
    ensureDir(dirname(p));
    var stream = fso.OpenTextFile(p, 2, true);
    stream.Write(String(data));
    stream.Close();
  }

  function appendFile(path, data) {
    var p = normalizePath(path);
    ensureDir(dirname(p));
    var stream = fso.OpenTextFile(p, 8, true);
    stream.Write(String(data));
    stream.Close();
  }

  function copyFile(src, dst, overwrite) {
    src = normalizePath(src);
    dst = normalizePath(dst);
    ensureDir(dirname(dst));
    try { fso.CopyFile(src, dst, overwrite == null ? true : !!overwrite); return true; } catch (e) { return false; }
  }

  function moveFile(src, dst) {
    src = normalizePath(src);
    dst = normalizePath(dst);
    ensureDir(dirname(dst));
    try { fso.MoveFile(src, dst); return true; } catch (e) { return false; }
  }

  function unlink(path) {
    path = normalizePath(path);
    try { if (fileExists(path)) fso.DeleteFile(path, true); return true; } catch (e) { return false; }
  }

  function rmdir(path) {
    path = normalizePath(path);
    try { if (dirExists(path)) fso.DeleteFolder(path, true); return true; } catch (e) { return false; }
  }

  function readdir(path) {
    var p = normalizePath(path);
    var folder = fso.GetFolder(p);
    var out = [];
    var e = new Enumerator(folder.SubFolders);
    for (; !e.atEnd(); e.moveNext()) out.push(String(e.item().Name));
    e = new Enumerator(folder.Files);
    for (; !e.atEnd(); e.moveNext()) out.push(String(e.item().Name));
    return out;
  }

  function stat(path) {
    var p = normalizePath(path);
    if (fileExists(p)) {
      var file = fso.GetFile(p);
      return {
        isFile: function () { return true; },
        isDirectory: function () { return false; },
        size: file.Size,
        mtime: file.DateLastModified,
        ctime: file.DateCreated,
        atime: file.DateLastAccessed,
        fullName: String(file.Path)
      };
    }
    if (dirExists(p)) {
      var folder = fso.GetFolder(p);
      return {
        isFile: function () { return false; },
        isDirectory: function () { return true; },
        size: 0,
        mtime: folder.DateLastModified,
        ctime: folder.DateCreated,
        atime: folder.DateLastAccessed,
        fullName: String(folder.Path)
      };
    }
    throw new Error("Path not found: " + p);
  }

  function sleep(ms) { WScript.Sleep(ms); }

  function exec(command, wait, windowStyle) {
    var w = windowStyle == null ? 0 : windowStyle;
    var shouldWait = wait == null ? true : !!wait;
    var code = shell.Run(command, w, shouldWait);
    return code;
  }

  function execCapture(command) {
    var proc = shell.Exec(command);
    while (proc.Status === 0) WScript.Sleep(10);
    var out = "";
    var err = "";
    try { out = proc.StdOut.ReadAll(); } catch (e) {}
    try { err = proc.StdErr.ReadAll(); } catch (e2) {}
    return { code: proc.ExitCode, stdout: out, stderr: err };
  }

  function env(key, value) {
    var envProc = shell.Environment("PROCESS");
    if (value === undefined) return String(envProc(key));
    envProc(key) = String(value);
    return String(envProc(key));
  }

  function argv() {
    var a = [];
    for (var i = 0; i < WScript.Arguments.length; i++) a.push(String(WScript.Arguments.Item(i)));
    return a;
  }

  function now() { return (new Date()).getTime(); }

  function createConsole() {
    function formatArgs(args) {
      var parts = [];
      for (var i = 0; i < args.length; i++) {
        var v = args[i];
        if (v == null) parts.push(String(v));
        else if (typeof v === "string") parts.push(v);
        else {
          try { parts.push(json.stringify(v)); } catch (e) { parts.push(String(v)); }
        }
      }
      return parts.join(" ");
    }
    return {
      log: function () { WScript.Echo(formatArgs(arguments)); },
      error: function () { WScript.Echo(formatArgs(arguments)); },
      warn: function () { WScript.Echo(formatArgs(arguments)); },
      info: function () { WScript.Echo(formatArgs(arguments)); }
    };
  }

  function createEvents() {
    function EventEmitter() {
      this._events = {};
      this._maxListeners = 50;
    }
    EventEmitter.prototype.setMaxListeners = function (n) { this._maxListeners = Number(n); return this; };
    EventEmitter.prototype.on = function (name, fn) {
      if (!this._events[name]) this._events[name] = [];
      this._events[name].push(fn);
      return this;
    };
    EventEmitter.prototype.addListener = EventEmitter.prototype.on;
    EventEmitter.prototype.once = function (name, fn) {
      var self = this;
      function wrapper() {
        self.off(name, wrapper);
        return fn.apply(null, arguments);
      }
      wrapper._original = fn;
      return this.on(name, wrapper);
    };
    EventEmitter.prototype.off = function (name, fn) {
      var list = this._events[name];
      if (!list) return this;
      for (var i = list.length - 1; i >= 0; i--) {
        if (list[i] === fn || list[i]._original === fn) list.splice(i, 1);
      }
      if (list.length === 0) delete this._events[name];
      return this;
    };
    EventEmitter.prototype.removeListener = EventEmitter.prototype.off;
    EventEmitter.prototype.removeAllListeners = function (name) {
      if (name) delete this._events[name];
      else this._events = {};
      return this;
    };
    EventEmitter.prototype.emit = function (name) {
      var list = this._events[name];
      if (!list || list.length === 0) return false;
      var args = [];
      for (var i = 1; i < arguments.length; i++) args.push(arguments[i]);
      list = list.slice(0);
      for (var j = 0; j < list.length; j++) {
        try { list[j].apply(null, args); } catch (e) {}
      }
      return true;
    };
    EventEmitter.prototype.listeners = function (name) { return (this._events[name] || []).slice(0); };
    return { EventEmitter: EventEmitter };
  }

  function createJson() {
    if (typeof JSON !== "undefined" && JSON && JSON.parse && JSON.stringify) return JSON;
    function esc(s) {
      return String(s)
        .replace(/\\/g, "\\\\")
        .replace(/"/g, '\\"')
        .replace(/\r/g, "\\r")
        .replace(/\n/g, "\\n")
        .replace(/\t/g, "\\t");
    }
    function stringify(val) {
      if (val === null) return "null";
      var t = typeof val;
      if (t === "number") return isFinite(val) ? String(val) : "null";
      if (t === "boolean") return val ? "true" : "false";
      if (t === "string") return '"' + esc(val) + '"';
      if (t === "undefined" || t === "function") return "null";
      if (val && typeof val.length === "number" && typeof val !== "string") {
        var arr = [];
        for (var i = 0; i < val.length; i++) arr.push(stringify(val[i]));
        return "[" + arr.join(",") + "]";
      }
      var obj = [];
      for (var k in val) if (hasOwn(val, k)) obj.push(stringify(k) + ":" + stringify(val[k]));
      return "{" + obj.join(",") + "}";
    }
    function parse(s) { return eval("(" + s + ")"); }
    return { stringify: stringify, parse: parse };
  }

  var json = createJson();
  var consoleObj = createConsole();
  var events = createEvents();

  function readJson(path) { return json.parse(readFile(path)); }
  function writeJson(path, obj, pretty) {
    var s = json.stringify(obj);
    if (pretty) {
      try {
        if (typeof JSON !== "undefined" && JSON && JSON.stringify) s = JSON.stringify(obj, null, 2);
      } catch (e) {}
    }
    writeFile(path, s);
  }

  function encodeUriComponent(s) { return encodeURIComponent(String(s)); }
  function decodeUriComponent(s) { return decodeURIComponent(String(s)); }

  function parseQueryString(qs) {
    var out = {};
    if (qs == null) return out;
    var s = String(qs);
    if (s.charAt(0) === "?") s = s.substring(1);
    if (!s) return out;
    var pairs = s.split("&");
    for (var i = 0; i < pairs.length; i++) {
      var p = pairs[i];
      if (!p) continue;
      var eq = p.indexOf("=");
      var k = eq >= 0 ? p.substring(0, eq) : p;
      var v = eq >= 0 ? p.substring(eq + 1) : "";
      k = decodeUriComponent(k.replace(/\+/g, "%20"));
      v = decodeUriComponent(v.replace(/\+/g, "%20"));
      if (hasOwn(out, k)) {
        if (typeof out[k].length === "number") out[k].push(v);
        else out[k] = [ out[k], v ];
      } else out[k] = v;
    }
    return out;
  }

  function stringifyQueryString(obj) {
    var parts = [];
    for (var k in obj) if (hasOwn(obj, k)) {
      var v = obj[k];
      if (v && typeof v.length === "number" && typeof v !== "string") {
        for (var i = 0; i < v.length; i++) parts.push(encodeUriComponent(k) + "=" + encodeUriComponent(v[i]));
      } else parts.push(encodeUriComponent(k) + "=" + encodeUriComponent(v));
    }
    return parts.join("&");
  }

  function parseUrl(url) {
    var s = String(url || "");
    var out = { href: s, protocol: "", host: "", hostname: "", port: "", pathname: "", search: "", query: {}, hash: "" };
    var hashPos = s.indexOf("#");
    if (hashPos >= 0) { out.hash = s.substring(hashPos); s = s.substring(0, hashPos); }
    var qPos = s.indexOf("?");
    if (qPos >= 0) { out.search = s.substring(qPos); out.query = parseQueryString(out.search); s = s.substring(0, qPos); }
    var protoPos = s.indexOf("://");
    if (protoPos >= 0) { out.protocol = s.substring(0, protoPos + 1); s = s.substring(protoPos + 3); }
    var slashPos = s.indexOf("/");
    var authority = slashPos >= 0 ? s.substring(0, slashPos) : s;
    out.pathname = slashPos >= 0 ? s.substring(slashPos) : "/";
    out.host = authority;
    var at = authority.lastIndexOf("@");
    if (at >= 0) authority = authority.substring(at + 1);
    var colon = authority.lastIndexOf(":");
    if (colon >= 0 && authority.indexOf("]") < 0) {
      out.hostname = authority.substring(0, colon);
      out.port = authority.substring(colon + 1);
    } else out.hostname = authority;
    return out;
  }

  function formatUrl(u) {
    var proto = u.protocol ? (u.protocol.indexOf(":") >= 0 ? u.protocol : u.protocol + ":") : "";
    var host = u.host || (u.hostname ? (u.port ? (u.hostname + ":" + u.port) : u.hostname) : "");
    var path = u.pathname || "/";
    var search = u.search || "";
    if (!search && u.query) {
      var qs = stringifyQueryString(u.query);
      if (qs) search = "?" + qs;
    }
    var hash = u.hash || "";
    var auth = host ? (proto ? proto + "//" : "") + host : "";
    return auth + path + search + hash;
  }

  function httpRequest(opts) {
    opts = opts || {};
    var method = (opts.method || "GET").toUpperCase();
    var url = String(opts.url || "");
    var async = !!opts.async;
    var headers = opts.headers || {};
    var body = opts.body == null ? null : String(opts.body);

    var xhr;
    try { xhr = new ActiveXObject("MSXML2.ServerXMLHTTP.6.0"); }
    catch (e) { xhr = new ActiveXObject("MSXML2.XMLHTTP"); }

    var done = false;
    var result = null;
    var error = null;

    function buildResult() {
      var status = 0;
      var statusText = "";
      var responseText = "";
      var responseHeaders = "";
      try { status = xhr.status; } catch (e1) {}
      try { statusText = xhr.statusText; } catch (e2) {}
      try { responseText = xhr.responseText; } catch (e3) {}
      try { responseHeaders = xhr.getAllResponseHeaders(); } catch (e4) {}
      return { status: status, statusText: statusText, text: responseText, headersRaw: responseHeaders };
    }

    if (async) {
      xhr.onreadystatechange = function () {
        if (xhr.readyState === 4) {
          try { result = buildResult(); } catch (e5) { error = e5; }
          done = true;
          if (opts.onDone) opts.onDone(error, result);
        }
      };
    }

    xhr.open(method, url, async);
    for (var k in headers) if (hasOwn(headers, k)) {
      try { xhr.setRequestHeader(String(k), String(headers[k])); } catch (e6) {}
    }
    try { xhr.send(body); } catch (e7) { if (async && opts.onDone) { opts.onDone(e7, null); return; } throw e7; }

    if (!async) return buildResult();
    return { isDone: function () { return done; }, getResult: function () { return result; }, getError: function () { return error; } };
  }

  function isWin() { return true; }

  function osInfo() {
    var out = { platform: "win32", arch: env("PROCESSOR_ARCHITECTURE") || "", hostname: "", user: "", tmpdir: env("TEMP") || env("TMP") || "", homedir: env("USERPROFILE") || "" };
    try { out.user = env("USERNAME") || ""; } catch (e) {}
    try {
      if (wmi) {
        var e1 = new Enumerator(wmi.ExecQuery("SELECT Name FROM Win32_ComputerSystem"));
        for (; !e1.atEnd(); e1.moveNext()) { out.hostname = String(e1.item().Name); break; }
      }
    } catch (e2) {}
    if (!out.hostname) {
      try { out.hostname = env("COMPUTERNAME") || ""; } catch (e3) {}
    }
    return out;
  }

  var moduleCache = {};
  var modulePaths = [];
  modulePaths.push(joinPath(scriptDir(), "modules"));
  modulePaths.push(joinPath(cwd(), "modules"));

  function addModulePath(p) {
    var np = normalizePath(p);
    for (var i = 0; i < modulePaths.length; i++) if (modulePaths[i] === np) return;
    modulePaths.unshift(np);
  }

  function readPackageMain(folder) {
    var pkg = joinPath(folder, "package.json");
    if (!fileExists(pkg)) return null;
    try {
      var obj = json.parse(readFile(pkg));
      if (obj && obj.main) return String(obj.main);
    } catch (e) {}
    return null;
  }

  function tryResolveFile(full) {
    if (fileExists(full)) return full;
    if (fileExists(full + ".js")) return full + ".js";
    return null;
  }

  function tryResolveFolder(full) {
    if (!dirExists(full)) return null;
    var mainRel = readPackageMain(full);
    if (mainRel) {
      var m = tryResolveFile(joinPath(full, mainRel));
      if (m) return m;
    }
    var idx = joinPath(full, "index.js");
    if (fileExists(idx)) return idx;
    return null;
  }

  function resolveAsPath(request, baseDir) {
    var full = resolvePath(baseDir, request);
    var f = tryResolveFile(full);
    if (f) return f;
    var d = tryResolveFolder(full);
    if (d) return d;
    return null;
  }

  function resolveAsModuleName(name) {
    var moduleFile = name + ".js";
    for (var i = 0; i < modulePaths.length; i++) {
      var cand = joinPath(modulePaths[i], moduleFile);
      if (fileExists(cand)) return cand;
      var folder = joinPath(modulePaths[i], name);
      var d = tryResolveFolder(folder);
      if (d) return d;
      var f = tryResolveFile(folder);
      if (f) return f;
    }
    return null;
  }

  function wrapAndEval(code, filename, baseDir) {
    var module = { exports: {} };
    var exports = module.exports;
    var wrapped = "(function(require, exports, module, __filename, __dirname, NScript){\n" + code + "\n})";
    var factory = eval(wrapped);
    function localRequire(p) { return require(p, baseDir); }
    factory(localRequire, exports, module, filename, baseDir, api);
    return module;
  }

  function loadModule(fullPath) {
    if (moduleCache[fullPath]) return moduleCache[fullPath].exports;
    var code = readFile(fullPath);
    var dir = dirname(fullPath);
    var module = wrapAndEval(code, fullPath, dir);
    moduleCache[fullPath] = module;
    return module.exports;
  }

  function require(request, baseDir) {
    baseDir = baseDir || cwd();
    var r = String(request || "");
    if (!r) throw new Error("Cannot find module: " + request);
    var resolved = null;
    if (isLikelyPath(r)) resolved = resolveAsPath(r, baseDir);
    else resolved = resolveAsModuleName(r);
    if (!resolved) throw new Error("Cannot find module: " + r);
    return loadModule(resolved);
  }

  function setTimeout(fn, ms) { sleep(ms); fn(); return 0; }

  function setInterval(fn, ms, times) {
    var count = times == null ? -1 : Number(times);
    while (count !== 0) {
      sleep(ms);
      fn();
      if (count > 0) count--;
    }
    return 0;
  }

  function exit(code) { WScript.Quit(code == null ? 0 : code); }

  function createBuffer() {
    function fromString(s) { return String(s == null ? "" : s); }
    function toBase64(s) {
      var dom = new ActiveXObject("MSXML2.DOMDocument.6.0");
      var el = dom.createElement("b64");
      el.dataType = "bin.base64";
      var stm = new ActiveXObject("ADODB.Stream");
      stm.Type = 2;
      stm.Charset = "utf-8";
      stm.Open();
      stm.WriteText(fromString(s));
      stm.Position = 0;
      stm.Type = 1;
      el.nodeTypedValue = stm.Read();
      stm.Close();
      return String(el.text);
    }
    function fromBase64(b64) {
      var dom = new ActiveXObject("MSXML2.DOMDocument.6.0");
      var el = dom.createElement("b64");
      el.dataType = "bin.base64";
      el.text = String(b64 || "");
      var bytes = el.nodeTypedValue;
      var stm = new ActiveXObject("ADODB.Stream");
      stm.Type = 1;
      stm.Open();
      stm.Write(bytes);
      stm.Position = 0;
      stm.Type = 2;
      stm.Charset = "utf-8";
      var text = stm.ReadText();
      stm.Close();
      return String(text);
    }
    return { fromString: fromString, toBase64: toBase64, fromBase64: fromBase64 };
  }

  var buffer = createBuffer();

  function delay(ms) {
    var emitter = new events.EventEmitter();
    setTimeout(function () { emitter.emit("done"); }, ms);
    return emitter;
  }

  var api = {
    version: "0.3.0",
    console: consoleObj,
    events: events,
    cwd: cwd,
    chdir: chdir,
    path: {
      normalize: normalizePath,
      join: joinPath,
      resolve: resolvePath,
      dirname: dirname,
      basename: basename,
      extname: extname,
      isAbsolute: isAbs
    },
    fs: {
      existsFile: fileExists,
      existsDir: dirExists,
      ensureDir: ensureDir,
      readFile: readFile,
      writeFile: writeFile,
      appendFile: appendFile,
      copyFile: copyFile,
      moveFile: moveFile,
      unlink: unlink,
      rmdir: rmdir,
      readdir: readdir,
      stat: stat,
      readJson: readJson,
      writeJson: writeJson
    },
    process: {
      argv: argv,
      env: env,
      now: now,
      exit: exit,
      platform: "win32",
      arch: function () { return env("PROCESSOR_ARCHITECTURE") || ""; }
    },
    childProcess: {
      exec: exec,
      execCapture: execCapture
    },
    http: {
      request: httpRequest
    },
    querystring: {
      parse: parseQueryString,
      stringify: stringifyQueryString
    },
    url: {
      parse: parseUrl,
      format: formatUrl
    },
    os: {
      info: osInfo
    },
    buffer: buffer,
    timers: {
      setTimeout: setTimeout,
      setInterval: setInterval,
      sleep: sleep,
      delay: delay
    },
    module: {
      require: require,
      addPath: addModulePath,
      paths: function () { return modulePaths.slice(0); }
    },
    require: require,
    json: json
  };

  return api;
})();
