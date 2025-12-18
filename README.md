<h1 style="display: flex; align-items: center; gap: 12px;">
  <img src="images/ns.jpg" alt="NScript" width="40" height="40" />
  <span>NScript</span>
</h1>

NScript is a lightweight Node.js–inspired module system and standard library for **Windows Script Host (WSH) JScript**.

It is not a full Node.js clone, but it brings familiar concepts such as `require`, modules, and common utilities into the classic JScript environment, making it easier to build structured, reusable scripts on Windows.

---

## What NScript Is

- A **single-file runtime** written in JScript
- Designed for **WSH (cscript / wscript)**
- Inspired by Node.js module and API patterns
- Focused on practicality, not full compatibility

NScript exists to modernize WSH scripting **without replacing it**.

---

## What NScript Is Not

- Not Node.js
- Not V8
- Not an event-loop–based runtime
- Not compatible with modern JavaScript syntax

NScript works within the limits of **classic JScript (ES3-level)**.

---

## Key Features

- `require()`-style module system
- `modules/` folder resolution
- `module.exports` support
- Built-in helpers inspired by Node.js:
  - `console`
  - `fs`
  - `path`
  - `http`
  - `childProcess`
  - `url`
  - `querystring`
  - `os`
  - `timers`
- Works with native Windows technologies:
  - COM / ActiveX
  - .NET (COM-visible libraries)
- Single-file deployment (`NScript.js`)

---

## Requirements

- Windows
- Windows Script Host enabled
- `cscript.exe` or `wscript.exe`

Recommended:
```bat
cscript //nologo
````

---

## Project Layout

Recommended structure:

```
project\
  NScript.js
  main.js
  modules\
    mymodule.js
```

---

## Loading NScript

WSH does not support `import` or `require` natively.
You load NScript by evaluating the file.

```js
var fso = new ActiveXObject("Scripting.FileSystemObject");
var f = fso.OpenTextFile("NScript.js", 1);
eval(f.ReadAll());
f.Close();
```

After this, the global `NS` object is available.

---

## Using Modules

Create a module:

**modules\math.js**

```js
function add(a, b) {
  return a + b;
}

module.exports = {
  add: add
};
```

Use it:

```js
var math = NS.require("math");
NS.console.log(math.add(2, 3));
```

---

## Example Script

**main.js**

```js
var fso = new ActiveXObject("Scripting.FileSystemObject");
var f = fso.OpenTextFile("NScript.js", 1);
eval(f.ReadAll());
f.Close();

var nhttp = NS.require("nhttp");
var res = NS.http.request({
  method: "GET",
  url: "https://example.com",
  async: false
});

NS.console.log(res.status);
```

Run:

```bat
cscript //nologo main.js
```

---

## Native Extensions (C# / C++)

NScript can call native code through **COM**.

* Write a COM-visible C# or C++ library
* Register it with Windows
* Call it from JScript using `ActiveXObject`

This provides a Node.js–like “native addon” capability for WSH.

---

## Limitations

* No `async / await`
* No Promises
* No event loop
* No modern JavaScript syntax
* Windows-only

NScript prioritizes **compatibility and stability** over language features.

---

## Why NScript Exists

* Many Windows environments still rely on WSH
* WSH is powerful but unstructured
* 
NScript fills the gap by providing a **familiar development model** on top of a **trusted Windows runtime**.

---

## Philosophy

* Minimal
* Explicit
* Windows-native
* No magic
* Easy to audit

---
```
```
