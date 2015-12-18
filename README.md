vscmd
=====

This utility controls Visual Studio from command line.

## Open Files from Command Line

<pre>vscmd <i>path-to-file</i></pre>

This syntax opens the specified file in Visual Studio.

## Debug Arguments

<pre>vsmcd args <i>arguments</i></pre>

This syntax sets the command line arguments
when you debug your project.

* If _arguments_ is omitted, vscmd displays the current settings.
* File paths are expanded to the full paths.
* C++ and C# projects are supported.
