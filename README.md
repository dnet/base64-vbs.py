EXE to VBS converter in Python
==============================

The purpose of this script is to convert an EXE (or any other binary file)
into such a format that

 - the original binary file can be restored on a vanilla Windows box and
 - the resulting file contains printable characters only.

It can be useful especially for penetration testing purposes, when executable
code needs to be entered into a system either via some dumb bindshell or
using emulated keyboards (USB HID, PS/2).

I solved the problem using VBS (Visual Basic Script) that makes it possible
to perform base64 decoding without any dependencies after a fresh install.
Many prefer using `debug` for this task, but that's limited for 64kbytes,
and requires hex characters, which requires 2 or 3 bytes for every input
byte (with or without spaces), whereas base64 requires only 4 bytes for
every 3 input byte (100% vs. 33% overhead).

Usage
-----

	$ python base64.py input.exe output.vbs

Dependencies
------------

 - Python 2.6 (tested on 2.7)

License
-------

The script is under MIT license.

Compression
-----------

The script tries to compress the output as best as it can, using the
`compress_vbs` function. This simple compressor collects all symbols from the
input file, renames variables to one character length identifiers and resolves
all constant expressions. It's really primitive, was only tested on the
`base64.vbs` template provided in the repository, so don't rely on it as a
general VBS minifier.
