vscmd
=====

This utility controls Visual Studio from command line.

## Open Files from Command Line

<pre>vscmd <i>path-to-file</i></pre>

This syntax opens the specified file in Visual Studio.

## Set Start Program to Debug

<pre>vsmcd start <i>path-to-program</i> <i>arguments...</i></pre>

This syntax sets the start program of the startup project.

* Paths are expanded to the full paths.
* If _arguments_ is specified, they are set to the command line arguments.

## Set Start Arguments to Debug

<pre>vsmcd arg <i>arguments...</i></pre>

This syntax sets the command line arguments of the startup project.

* Paths are expanded to the full paths.
