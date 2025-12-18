# Using C# Functions With NSAPI

This document explains how to call C# functions from **Windows Script Host (WSH) JScript** using **NSAPI** (`nsapi.js`). The approach is based on **COM-visible .NET class libraries**, which makes it simple for others to build and expose their own native functionality.

---

## Overview

* **You write C# code** as a **COM-visible** class library (DLL).
* You **register** the DLL on Windows (COM registration).
* In JScript, you load and call it using:

  * `nsapi.create("ProgId")`, or
  * `nsapi.call("ProgId", "MethodName", ...)`

This provides a stable “native extension” pattern for WSH JScript.

---

## Requirements

* Windows with **WSH** enabled
* **.NET Framework** (for `regasm`)
* A C# compiler (Visual Studio / MSBuild)

To run scripts, use `cscript.exe`:

```bat
cscript //nologo your_script.js
```

---

## Files

* `nsapi.js` — the NSAPI bridge helper (JScript)
* `NsApi.dll` — your compiled C# COM-visible library
* `your_script.js` — your script that calls the C# library

---

## C# Side

### 1) Create a Class Library

Create a **.NET Framework Class Library** project and add the following code.

**NsApi.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace NsApi
{
    [ComVisible(true)]
    [Guid("9C6E8F8E-6E42-4B5B-9A2C-0F2B6C6E2A11")]
    [ProgId("NsApi.Native")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Native
    {
        public int Add(int a, int b)
        {
            return a + b;
        }

        public string ToUpper(string value)
        {
            return (value ?? "").ToUpperInvariant();
        }
    }
}
```

### 2) Build the DLL

Build the project to produce:

* `NsApi.dll`

### 3) Register the DLL for COM

Run (as Administrator if required):

```bat
regasm NsApi.dll /codebase
```

If you need to unregister:

```bat
regasm NsApi.dll /unregister
```

### 4) 32-bit vs 64-bit

WSH and your DLL must match architecture:

* If you run 64-bit `cscript.exe`, register a 64-bit DLL.
* If you run 32-bit `cscript.exe`, register a 32-bit DLL.

Common paths:

* 64-bit: `C:\Windows\System32\cscript.exe`
* 32-bit: `C:\Windows\SysWOW64\cscript.exe`

---

## JScript Side

### 1) Load NSAPI (`nsapi.js`)

Place `nsapi.js` next to your script (or adjust the path as needed), then load it.

Example loader:

```js
var fso = new ActiveXObject("Scripting.FileSystemObject");
var f = fso.OpenTextFile("nsapi.js", 1);
eval(f.ReadAll());
f.Close();
```

After this, you will have `nsapi` available.

---

## Calling C# Methods

### Option A: Create an instance and call methods

```js
var native = nsapi.create("NsApi.Native");

WScript.Echo(native.Add(3, 5));
WScript.Echo(native.ToUpper("hello"));
```

### Option B: One-liner calls (no stored instance)

```js
WScript.Echo(nsapi.call("NsApi.Native", "Add", 10, 20));
WScript.Echo(nsapi.call("NsApi.Native", "ToUpper", "wsh"));
```

---

## Guidelines for Library Authors

If you want others to easily expose their own C# functions to JScript:

1. Use a stable ProgID pattern:

* `Company.Product` or `Project.Component`

2. Mark the class as COM-visible:

* `[ComVisible(true)]`

3. Ensure methods are:

* `public`
* have COM-friendly argument and return types:

  * `int`, `double`, `bool`, `string`
  * arrays of these types (advanced)

4. Keep APIs small and explicit. If you need complex data, prefer passing strings (often JSON).

---

## Troubleshooting

### “ActiveX component can’t create object”

* The DLL is not registered, or ProgID does not match.
* Run:

  ```bat
  regasm NsApi.dll /codebase
  ```
* Confirm the ProgID in C# matches what you use in JScript.

### “Type mismatch” or unexpected values

* COM prefers simple types.
* Avoid complex .NET objects as return values.
* If you need structured output, return JSON strings.

### Script runs but methods are missing

* Ensure the methods are `public`.
* Ensure the class is `[ComVisible(true)]`.

---

## Minimal Example Layout

```
project\
  nsapi.js
  call_native.js
  bin\
    NsApi.dll
```

Example run:

```bat
cscript //nologo call_native.js
```

---

## Example: call_native.js

```js
var fso = new ActiveXObject("Scripting.FileSystemObject");
var f = fso.OpenTextFile("nsapi.js", 1);
eval(f.ReadAll());
f.Close();

var native = nsapi.create("NsApi.Native");
WScript.Echo(native.Add(1, 2));
WScript.Echo(native.ToUpper("ok"));
```
