#!/usr/bin/python
# -*- coding: utf-8 -*-
"""Script to extract Visual Basic for Applications (VBA)."""

import argparse
import logging
import sys

from olecfrc import vba


class StdoutWriter(object):
  """Class that defines a stdout output writer."""

  def Close(self):
    """Closes the output writer object."""
    return

  def Open(self):
    """Opens the output writer object.

    Returns:
      bool: True if successful or False if not.
    """
    return True

  def WriteText(self, text):
    """Writes text to the output.

    Args:
      text (str): text.
    """
    print(text)


def Main():
  """The main program function.

  Returns:
    bool: True if successful or False if not.
  """
  argument_parser = argparse.ArgumentParser(description=(
      'Extracts VBA from an OLE Compound File.'))

  argument_parser.add_argument(
      '-d', '--debug', dest='debug', action='store_true', default=False,
      help='enable debug output.')

  argument_parser.add_argument(
      'source', nargs='?', action='store', metavar='PATH', default=None,
      help='path of the OLE Compound File.')

  options = argument_parser.parse_args()

  if not options.source:
    print('Source value is missing.')
    print('')
    argument_parser.print_help()
    print('')
    return False

  logging.basicConfig(
      level=logging.INFO, format='[%(levelname)s] %(message)s')

  output_writer = StdoutWriter()

  if not output_writer.Open():
    print('Unable to open output writer.')
    print('')
    return False

  collector_object = vba.VBACollector(debug=options.debug)
  collector_object.Collect(options.source, output_writer)
  output_writer.Close()

  if not collector_object.stream_found:
    print('No VBA stream found.')

  return True


if __name__ == '__main__':
  if not Main():
    sys.exit(1)
  else:
    sys.exit(0)
