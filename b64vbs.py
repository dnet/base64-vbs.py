#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# bas64.py - converts EXE (or other binary) files to VBS scripts
#
# Copyright (c) 2012 András Veres-Szentkirályi
#
# Permission is hereby granted, free of charge, to any person
# obtaining a copy of this software and associated documentation
# files (the "Software"), to deal in the Software without
# restriction, including without limitation the rights to use,
# copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the
# Software is furnished to do so, subject to the following
# conditions:
#
# The above copyright notice and this permission notice shall be
# included in all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
# NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
# HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
# WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
# FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
# OTHER DEALINGS IN THE SOFTWARE.

from __future__ import with_statement, print_function
from base64 import b64encode
import sys

TEMPLATE = 'base64.vbs'
BUFSIZE = 4095 # divisible by 3

def main(args):
	try:
		exe_fn, vbs_fn = args[1:3]
	except ValueError:
		print('Usage: {0} input.exe output.vbs'.format(args[0]), file=sys.stderr)
		sys.exit(1)
	with file(vbs_fn, 'w') as vbs_file:
		with file(TEMPLATE, 'rb') as tpl_file:
			tpl_contents = tpl_file.read()
		tpl_contents = '\n'.join(compress_vbs(tpl_contents.split('\n')))
		before_b64, after_b64 = tpl_contents.split('%%DATA%%', 1)
		vbs_file.write(before_b64)
		with file(exe_fn, 'rb') as exe_file:
			while True:
				buf = exe_file.read(BUFSIZE)
				if not buf:
					break
				vbs_file.write(b64encode(buf))
		vbs_file.write(after_b64)

def compress_vbs(lines):
	dims = set()
	symbols = dict()
	for line in lines:
		if line.startswith('Dim '):
			for symbol in line[4:].split(','):
				value = chr(ord('a') + len(dims))
				dims.add(value)
				symbols[symbol.strip()] = value
		elif line.startswith('Const '):
			symbol, value = line[6:].split('=')
			symbols[symbol.strip()] = value.strip()
	yield 'Dim ' + ','.join(dims)
	for line in lines:
		if line == '' or line.startswith("'") or line.startswith('Dim ') or line.startswith('Const '):
			continue
		for symbol, value in symbols.iteritems():
			line = line.replace(symbol, value)
		yield line.replace(' = ', '=').replace(', ', ',')

if __name__ == '__main__':
	main(sys.argv)
