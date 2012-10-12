#!/usr/bin/env python

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
