vscmd
=====

This utility controls Visual Studio from command line.

## Open Files from Command Line

<pre>vscmd <i>path-to-file</i></pre>

This syntax opens the specified file in Visual Studio.

## Debug Start Program

<pre>vsmcd start <i>path-to-program</i> <i>arguments...</i></pre>

This syntax sets the path of start program of the startup project.

* File paths are expanded to the full paths.
* If _arguments_ is specified, they are set to the command line arguments.
  To set the command line arguments without changing the start program,
  please refer to the `arg` command.
* If arguments are omitted, vscmd displays the current settings.

## Debug Arguments

<pre>vsmcd arg <i>arguments...</i></pre>

This syntax sets the command line arguments of the startup project.

* File paths are expanded to the full paths.
* If _arguments_ is omitted, vscmd displays the current settings.
