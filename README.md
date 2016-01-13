vscmd
=====

A command line utility to control Visual Studio.

## Open Files from Command Line

<pre>vscmd <i>path-to-file...</i></pre>

This syntax opens the specified file in Visual Studio.

## Start Debug the Specified Program

<pre>vsmcd start <i>path-to-program</i> <i>arguments...</i></pre>

This syntax sets the start program of the startup project,
and start debug.

Paths are expanded to the full paths.

## Set Start Arguments to Debug

<pre>vsmcd arg <i>arguments...</i></pre>

This syntax sets the command line arguments of the startup project.

Paths are expanded to the full paths.

## Attach to Process

<pre>vscmd attach <i>arguments...</i></pre>

This syntax attaches the Visual Studio debugger to the specified processes.

vscmd attaches the debugger to all process where any _arguments_ match to the process ID or the name.

vscmd displays a list of process ID and name if no _arguments_ are given.