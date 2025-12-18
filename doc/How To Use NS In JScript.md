# How To Use NS In JScript

This guide shows how to use **NS** (your WSH/JScript runtime object) from a plain **Windows Script Host JScript** file.

The examples assume you run scripts with `cscript.exe`.

---

## Requirements

* Windows Script Host enabled
* `cscript.exe` available
* `NScript.js` (the file that defines `NS`) available on disk

Run scripts like this:

```bat
cscript //nologo your_script.js
```

---

## Project Layout (Recommended)

```
project\
  NScript.js
  your_script.js
  modules\
    nhttp.js
```

---

## 1) Loading NS in a Script

WSH JScript does not have a native `import` mechanism. You load NS by reading `NScript.js` and evaluating it.

**your_script.js**

```js
var fso = new ActiveXObject("Scripting.FileSystemObject");
var f = fso.OpenTextFile("NScript.js", 1);
eval(f.ReadAll());
f.Close();

NS.console.log("NS loaded");
```

---

## 2) Console Output

NS provides a `console` object.

```js
NS.console.log("hello");
NS.console.warn("warn");
NS.console.error("error");
```

---

## 3) Working With Paths

Use `NS.path` for common path operations.

```js
var p = NS.path.join(NS.cwd(), "data", "file.txt");
NS.console.log(p);

NS.console.log(NS.path.dirname(p));
NS.console.log(NS.path.extname(p));
```

---

## 4) Working With Files

Use `NS.fs` for file operations.

```js
var filePath = NS.path.join(NS.cwd(), "out.txt");

NS.fs.writeFile(filePath, "test");
NS.fs.appendFile(filePath, "\nmore");
var content = NS.fs.readFile(filePath);

NS.console.log(content);
```

Check existence:

```js
if (NS.fs.existsFile(filePath)) {
  NS.console.log("file exists");
}
```

List directory:

```js
var items = NS.fs.readdir(NS.cwd());
NS.console.log(items);
```

---

## 5) Running External Commands

Use `NS.childProcess`.

Blocking run (returns exit code):

```js
var code = NS.childProcess.exec("whoami", true);
NS.console.log(code);
```

Capture stdout/stderr:

```js
var res = NS.childProcess.execCapture("whoami");
NS.console.log(res.code);
NS.console.log(res.stdout);
NS.console.log(res.stderr);
```

---

## 6) Environment Variables and Arguments

Arguments:

```js
var args = NS.process.argv();
NS.console.log(args);
```

Read environment variables:

```js
NS.console.log(NS.process.env("USERNAME"));
NS.console.log(NS.process.env("TEMP"));
```

Set environment variables (process-level):

```js
NS.process.env("MY_VAR", "123");
NS.console.log(NS.process.env("MY_VAR"));
```

---

## 7) Using Modules

If you have a module in `modules\`, you can load it by name.

Example:

```
modules\nhttp.js
```

Load and use:

```js
var nhttp = NS.require("nhttp");
var res = NS.http.request({ method: "GET", url: "https://example.com", async: false });
NS.console.log(res.status);
```

NS typically resolves module names like:

* `modules\<name>.js`
* `modules\<name>\index.js`
* `modules\<name>\package.json` (`main`)

---

## 8) Making HTTP Requests

Synchronous request:

```js
var res = NS.http.request({
  method: "GET",
  url: "https://example.com",
  async: false
});

NS.console.log(res.status);
NS.console.log(res.text);
```

Asynchronous request (callback style):

```js
NS.http.request({
  method: "GET",
  url: "https://example.com",
  async: true,
  onDone: function (err, res) {
    if (err) {
      NS.console.error(err.message);
      return;
    }
    NS.console.log(res.status);
  }
});
```

---

## 9) Timers

WSH is not event-loop based like Node.js, but NS exposes basic timer helpers.

Delay (blocking):

```js
NS.timers.sleep(250);
NS.console.log("done");
```

Timeout (blocking implementation):

```js
NS.timers.setTimeout(function () {
  NS.console.log("later");
}, 250);
```

Interval (blocking loop):

```js
var count = 0;
NS.timers.setInterval(function () {
  count++;
  NS.console.log("tick " + count);
}, 200, 5);
```

---

## 10) OS Information

```js
var info = NS.os.info();
NS.console.log(info.platform);
NS.console.log(info.arch);
NS.console.log(info.hostname);
NS.console.log(info.homedir);
NS.console.log(info.tmpdir);
```

---

## 11) JSON Helpers

If the host provides `JSON`, NS uses it; otherwise it provides a fallback.

```js
var obj = { a: 1, b: "x" };
var s = NS.json.stringify(obj);
NS.console.log(s);

var back = NS.json.parse(s);
NS.console.log(back.a);
```

---

## Common Notes

### No `async/await`

WSH JScript does not support `async/await`. Use:

* synchronous calls, or
* callback-style asynchronous APIs when available.

### Avoid modern JavaScript syntax

Use ES3-style syntax:

* `var`, `function`
* no arrow functions
* no template strings
* no `class`

---

## Minimal Starter Script

Copy/paste template:

```js
var fso = new ActiveXObject("Scripting.FileSystemObject");
var f = fso.OpenTextFile("NScript.js", 1);
eval(f.ReadAll());
f.Close();

NS.console.log("ready");
```

Run:

```bat
cscript //nologo starter.js
```
