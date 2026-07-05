# Third-Party Notices

VS IDE Bridge is distributed under the MIT License (see [LICENSE](LICENSE)). The installer,
the Windows service, and the Visual Studio extension redistribute the third-party components
listed below. Where a component ships with its own license file, that file is included in the
installed tree at the path noted.

## Python (CPython)

The installer can optionally provision a bundled Python runtime (CPython 3.13.x).

- Website: https://www.python.org/
- License: PSF License Version 2.0
- Full license text as installed: `python\managed-runtime\LICENSE.txt`

Copyright (c) 2001 Python Software Foundation; All Rights Reserved.

## pip

The bundled Python runtime includes pip (25.x).

- Website: https://pip.pypa.io/
- License: MIT

pip vendors a number of third-party packages (including certifi, distro, pygments, requests,
rich, tomli, and others). Each vendored package's license file is redistributed unmodified
inside the runtime under `python\managed-runtime\Lib\site-packages\pip\_vendor\<package>\`.

## LibGit2Sharp

The Windows service uses LibGit2Sharp for in-process Git repository access (read-only
inspection flows such as branch, diff, log, remote, and status reads).

- Package: LibGit2Sharp
- Website: https://github.com/libgit2/libgit2sharp
- License: MIT

Copyright (c) LibGit2Sharp contributors.

## libgit2 (native library)

LibGit2Sharp redistributes the native libgit2 library (`git2-*.dll`), which ships in the
service directory alongside `LibGit2Sharp.dll`.

- Website: https://libgit2.org/
- License: GNU General Public License v2 **with Linking Exception**

The linking exception permits linking libgit2 into this application and distributing the
combined work under this project's own terms without the GPL extending to the rest of the
program. libgit2 itself additionally incorporates code under permissive licenses (see the
COPYING file in the libgit2 source distribution).

## Newtonsoft.Json (Json.NET)

Shipped with the Windows service.

- Website: https://www.newtonsoft.com/json
- License: MIT

Copyright (c) 2007 James Newton-King.

## Microsoft .NET libraries

The service and the Visual Studio extension redistribute Microsoft-published .NET libraries,
including: `Microsoft.Build.Locator`, `Microsoft.Extensions.FileSystemGlobbing`,
`Microsoft.Bcl.AsyncInterfaces`, and `System.*` compatibility/runtime packages
(`System.Text.Json`, `System.IO.Pipelines`, `System.Memory`, `System.Buffers`,
`System.Threading.Tasks.Extensions`, `System.ValueTuple`, and related assemblies).

- License: MIT
- Copyright (c) .NET Foundation and Contributors / Microsoft Corporation.

## Microsoft Visual Studio SDK

The Visual Studio extension is built against Microsoft Visual Studio SDK packages
(`Microsoft.VisualStudio.*`, EnvDTE, VSLangProj, and related interop assemblies). These
assemblies are provided by the user's Visual Studio installation at runtime and are not
redistributed by this project. Their use is governed by the Microsoft Visual Studio SDK
license terms.

## Inno Setup

The Windows installer (`vs-ide-bridge-setup-<version>.exe`) is built with Inno Setup, whose
license permits free distribution of installers it creates.

- Website: https://jrsoftware.org/isinfo.php
- License: Inno Setup License (as distributed with Inno Setup)

Copyright (C) 1997-2026 Jordan Russell. All rights reserved.
Portions Copyright (C) 2000-2026 Martijn Laan. All rights reserved.

Inno Setup binaries and source are not redistributed by this repository; Inno Setup is used
as a build-time tool via a local installation of `ISCC.exe`.

## Development-only dependencies

Test and build tooling (xUnit, Microsoft.NET.Test.Sdk, coverlet, and similar packages) is not
redistributed with any release artifact and is therefore not listed above.
