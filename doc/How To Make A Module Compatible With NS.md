# How To Make A Module Compatible With NS

This document describes how to build a JavaScript module that can be loaded by **NS** (the module loader included with your WSH/JScript runtime), in a way that feels similar to Node.js modules while remaining fully compatible with **WSH JScript**.

The goal is to let users write:

```js
var nhttp = NS.require("nhttp");
```

…and have NS resolve that to:

* `.\modules\nhttp.js`, or
* `.\modules\nhttp\index.js`, or
* `.\modules\nhttp\package.json` → `main`

---

## Compatibility Rules

A module is compatible with NS if it follows these rules:

1. It is plain **WSH JScript** (no `import`, no `class`, no `async/await`, no modern syntax).
2. It exports its public API using `module.exports`.
3. It does not rely on Node-only globals such as `process`, `Buffer`, or `require` unless NS provides them.
4. If it has sub-files, it uses NS’s `require` mechanism (the injected `require` argument), not Node’s.

NS loads modules by wrapping the file with a factory function, similar to Node:

```js
(function(require, exports, module, __filename, __dirname, NS){ ... })
```

So your module automatically receives:

* `require`
* `exports`
* `module`
* `__filename`
* `__dirname`
* `NS` (the runtime object)

---

## Folder Layout Options

### Option A: Single-file module

```
project\
  NScript.js
  modules\
    nhttp.js
```

Load:

```js
var nhttp = NS.require("nhttp");
```

Resolution target:

* `modules\nhttp.js`

---

### Option B: Module folder with index.js

```
project\
  NScript.js
  modules\
    nhttp\
      index.js
```

Load:

```js
var nhttp = NS.require("nhttp");
```

Resolution targets (in order):

* `modules\nhttp.js`
* `modules\nhttp\package.json` → `main`
* `modules\nhttp\index.js`

---

### Option C: Module folder with package.json

```
project\
  NScript.js
  modules\
    nhttp\
      package.json
      main.js
```

**package.json**

```json
{
  "name": "nhttp",
  "main": "main.js"
}
```

Load:

```js
var nhttp = NS.require("nhttp");
```

Resolution target:

* `modules\nhttp\main.js` (from `main`)

---

## Writing an NS-Compatible Module

### Minimal module example (nhttp.js)

**modules\nhttp.js**

```js
function get(url) {
  var res = NS.http.request({ method: "GET", url: url, async: false });
  return res;
}

module.exports = {
  get: get
};
```

Usage:

```js
var nhttp = NS.require("nhttp");
var res = nhttp.get("https://example.com");
NS.console.log(res.status, res.text);
```

---

## Using Local Files Inside a Module

If your module has multiple files, use the injected `require` argument to load relative modules.

Example layout:

```
modules\
  nhttp\
    index.js
    headers.js
```

**modules\nhttp\headers.js**

```js
function normalizeHeaderName(name) {
  return String(name || "").toLowerCase();
}

module.exports = {
  normalizeHeaderName: normalizeHeaderName
};
```

**modules\nhttp\index.js**

```js
var headers = require("./headers");

function get(url) {
  var res = NS.http.request({ method: "GET", url: url, async: false });
  return res;
}

module.exports = {
  get: get,
  normalizeHeaderName: headers.normalizeHeaderName
};
```

Usage:

```js
var nhttp = NS.require("nhttp");
NS.console.log(nhttp.normalizeHeaderName("Content-Type"));
```

---

## Export Patterns

### Pattern 1: Export an object

```js
module.exports = {
  a: function () { return 1; },
  b: function () { return 2; }
};
```

### Pattern 2: Export a single function

```js
function main() { return "ok"; }
module.exports = main;
```

Usage:

```js
var main = NS.require("mymodule");
NS.console.log(main());
```

---

## What to Avoid

### Do not use modern syntax that WSH JScript cannot parse

Avoid:

* `import` / `export`
* `class`
* `async` / `await`
* arrow functions `() => {}`
* template strings `` `x${y}` ``

Stick to:

* `function () { ... }`
* `var x = ...`
* basic objects and arrays

### Do not assume Node built-ins exist

Avoid calling:

* `require("fs")`, `require("http")`, etc.

Instead, use NS-provided APIs, for example:

* `NS.fs`
* `NS.http`
* `NS.childProcess`
* `NS.url`
* `NS.querystring`

---

## Recommended Module Design

### Keep the public API small and stable

* Export only what users should call.
* Hide internal helpers in separate files.

### Prefer string inputs/outputs for portability

For structured output, return JSON strings or plain objects (if they only contain primitive fields).

### Validate arguments

Fail fast if inputs are invalid.

---

## Testing a Module

Create a test script:

**test_nhttp.js**

```js
var fso = new ActiveXObject("Scripting.FileSystemObject");
var f = fso.OpenTextFile("NScript.js", 1);
eval(f.ReadAll());
f.Close();

var nhttp = NS.require("nhttp");
var res = nhttp.get("https://example.com");
NS.console.log(res.status);
```

Run:

```bat
cscript //nologo test_nhttp.js
```

---

## Summary Checklist

A module is NS-compatible if:

* It is valid WSH JScript (ES3-level syntax).
* It exports via `module.exports`.
* It uses the injected `require` for relative imports.
* It lives in `modules\` as either:

  * `modules\<name>.js`, or
  * `modules\<name>\index.js`, or
  * `modules\<name>\package.json` with `main`.

That is all you need for NS to load it reliably.
