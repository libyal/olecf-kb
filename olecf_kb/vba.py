# -*- coding: utf-8 -*-
"""Visual Basic for Applications (VBA) collector."""

from __future__ import print_function
import uuid

import construct
import pyolecf

from olecf_kb import hexdump


if pyolecf.get_version() < '20160814':
  raise ImportWarning('vba.py requires pyolecf 20160814 or later.')


class FStream(object):
  """Class that defines a f stream."""

  _HEADER = construct.Struct(
      'header',
      construct.ULInt32('unknown1'),
      construct.ULInt32('unknown2'),
      construct.ULInt32('unknown3'),
      construct.ULInt32('unknown4'),
      construct.ULInt32('unknown5'),
      construct.ULInt32('unknown6'),
      construct.ULInt32('unknown7'),
      construct.ULInt32('unknown8'),
      construct.ULInt32('unknown9'),
      construct.ULInt32('unknown10'),
      construct.ULInt32('unknown11'),
      construct.Bytes('unknown12', 16),
      construct.Bytes('unknown13', 23))

  _ENTRY = construct.Struct(
      'entry',
      construct.ULInt32('unknown7'),
      construct.ULInt32('unknown8'),
      construct.ULInt16('unknown9'),
      # Does not include the 2 bytes of the size value.
      construct.ULInt16('size'),
      construct.ULInt32('unknown1'),
      construct.ULInt32('unknown2'),
      construct.ULInt32('unknown3'),
      construct.ULInt32('o_stream_entry_size'),
      construct.ULInt16('o_stream_entry_index'),
      construct.ULInt16('unknown6'),
      construct.Bytes('variable_name', lambda ctx: ctx.size - 28))

  # TODO: - 28 does not hold for the last entry but size does.

  def __init__(self, debug=False):
    """Initializes a stream.

    Args:
      debug (Optional[bool]): True if debug information should be printed.
    """
    super(FStream, self).__init__()
    self._debug = debug

  def Read(self, olecf_item):
    """Reads the stream from the OLECF item.

    Args:
      olecf_item (pyolecf.item): OLECF item.

    Returns:
      bool: True if the stream was successfully read.
    """
    stream_data = olecf_item.read()

    header_struct = self._HEADER.parse(stream_data)

    stream_offset = self._HEADER.sizeof()

    if self._debug:
      print('f stream header data:')
      print(hexdump.Hexdump(stream_data[:stream_offset]))

    if self._debug:
      print('Unknown1\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown1))
      print('Unknown2\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown2))
      print('Unknown3\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown3))
      print('Unknown4\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown4))
      print('Unknown5\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown5))
      print('Unknown6\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown6))
      print('Unknown7\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown7))
      print('Unknown8\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown8))
      print('Unknown9\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown9))
      print('Unknown10\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown10))
      print('Unknown11\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown11))

      # CLSID of StdFont: 0be35203-8f91-11ce-9de3-00aa004bb851
      uuid_value = uuid.UUID(bytes_le=header_struct.unknown12)
      print('Unknown12\t\t\t\t\t\t\t: {0:s}'.format(uuid_value))

      print('')

    while stream_offset < olecf_item.size:
      try:
        entry_struct = self._ENTRY.parse(stream_data[stream_offset:])
      except construct.core.FieldError:
        break

      next_stream_offset = stream_offset + 2 + entry_struct.size + 2

      if self._debug:
        print('f stream entry data:')
        print(hexdump.Hexdump(stream_data[stream_offset:next_stream_offset]))

      if self._debug:
        print('Unknown7\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_struct.unknown7))
        print('Unknown8\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_struct.unknown8))
        print('Unknown9\t\t\t\t\t\t\t: 0x{0:04x}'.format(
            entry_struct.unknown9))

        print('Size\t\t\t\t\t\t\t\t: {0:d}'.format(entry_struct.size))

        print('Unknown1\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_struct.unknown1))
        print('Unknown2\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_struct.unknown2))
        print('Unknown3\t\t\t\t\t\t\t: {0:d}'.format(
            entry_struct.unknown3))
        print('O stream entry size\t\t\t\t\t\t: {0:0d}'.format(
            entry_struct.o_stream_entry_size))
        print('O stream entry index\t\t\t\t\t\t: {0:d}'.format(
            entry_struct.o_stream_entry_index))
        print('Unknown6\t\t\t\t\t\t\t: 0x{0:04x}'.format(
            entry_struct.unknown6))

        # TODO: fix this.
        try:
          variable_name = entry_struct.variable_name.decode('cp1252')
          print('Variable name\t\t\t\t\t\t\t: {0:s}'.format(variable_name))
        except UnicodeEncodeError:
          pass

        print('')

      stream_offset = next_stream_offset

    return True


class OStream(object):
  """Class that defines an o stream."""

  _ENTRY_PART1 = construct.Struct(
      'entry_part1',
      construct.ULInt32('unknown1'),
      construct.ULInt32('unknown2'),
      construct.ULInt32('unknown3'),
      construct.ULInt32('unknown4'),
      construct.ULInt32('data_size'),
      construct.ULInt32('unknown6'),
      construct.ULInt32('unknown7'),
      construct.CString('data'))

  _ENTRY_PART2 = construct.Struct(
      'entry_part2',
      construct.ULInt32('unknown7'),
      construct.ULInt32('unknown8'),
      construct.ULInt32('unknown9'),
      construct.ULInt32('unknown10'),
      construct.ULInt32('unknown11'),
      construct.CString('font_name'))

  def __init__(self, debug=False):
    """Initializes a stream.

    Args:
      debug (Optional[bool]): True if debug information should be printed.
    """
    super(OStream, self).__init__()
    self._debug = debug

  def Read(self, olecf_item):
    """Reads the stream from the OLECF item.

    Args:
      olecf_item (pyolecf.item): OLECF item.

    Returns:
      bool: True if the stream was successfully read.
    """
    stream_data = olecf_item.read()

    stream_offset = 0
    while stream_offset < olecf_item.size:
      try:
        entry_part1_struct = self._ENTRY_PART1.parse(
            stream_data[stream_offset:])
      except construct.core.FieldError:
        break

      entry_part_size = (7 * 4) + len(entry_part1_struct.data) + 1
      padding_size = entry_part_size % 4
      if padding_size != 0:
        padding_size = 4 - padding_size

      next_stream_offset = stream_offset + entry_part_size + padding_size

      try:
        entry_part2_struct = self._ENTRY_PART2.parse(
            stream_data[next_stream_offset:])
      except construct.core.FieldError:
        break

      entry_part_size = (5 * 4) + len(entry_part2_struct.font_name) + 1
      padding_size = entry_part_size % 4
      if padding_size != 0:
        padding_size = 4 - padding_size

      next_stream_offset += entry_part_size + padding_size

      if self._debug:
        print('o stream entry data:')
        print(hexdump.Hexdump(stream_data[stream_offset:next_stream_offset]))

      # TODO: add debug info.
      if self._debug:
        print('Unknown1\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part1_struct.unknown1))
        print('Unknown2\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part1_struct.unknown2))
        print('Unknown3\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part1_struct.unknown3))
        print('Unknown4\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part1_struct.unknown4))
        print('Data size\t\t\t\t\t\t\t: {0:d} (0x{1:08x})'.format(
            entry_part1_struct.data_size & 0x7fffffff,
            entry_part1_struct.data_size))
        print('Unknown6\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part1_struct.unknown6))
        print('Data\t\t\t\t\t\t\t\t: {0:s}'.format(entry_part1_struct.data))
        # TODO: alignment padding.
        print('Unknown7\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part2_struct.unknown7))
        print('Unknown8\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part2_struct.unknown8))
        print('Unknown9\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part2_struct.unknown9))
        print('Unknown10\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part2_struct.unknown10))
        print('Unknown11\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part2_struct.unknown11))
        print('Font name\t\t\t\t\t\t\t: {0:s}'.format(
            entry_part2_struct.font_name))
        # TODO: alignment padding.
        print('')

      stream_offset = next_stream_offset

    return True


class VBAProjectStream(object):
  """Class that defines a _VBA_PROJECT (Preformance Cache) stream."""

  _HEADER = construct.Struct(
      'header',
      construct.ULInt32('unknown1'),
      construct.ULInt16('unknown2'),
      construct.ULInt16('unknown3'),
      construct.ULInt32('unknown4'),
      construct.ULInt32('unknown5'),
      construct.ULInt32('unknown6'),
      construct.ULInt32('unknown7'),
      construct.ULInt32('unknown8'),
      construct.ULInt16('unknown9'),
      construct.ULInt16('number_of_strings'),
      construct.ULInt16('unknown11'))

  _STRING = construct.Struct(
      'string',
      # Does not include the end-of-string character.
      construct.ULInt16('string_size'),
      construct.Bytes(
          'string',
          lambda ctx: ctx.string_size),
      construct.ULInt32('unknown1'),
      construct.ULInt32('unknown2'),
      construct.ULInt32('unknown3'))

  def __init__(self, debug=False):
    """Initializes a stream.

    Args:
      debug (Optional[bool]): True if debug information should be printed.
    """
    super(VBAProjectStream, self).__init__()
    self._debug = debug

  def Read(self, olecf_item):
    """Reads the stream from the OLECF item.

    Args:
      olecf_item (pyolecf.item): OLECF item.

    Returns:
      bool: True if the stream was successfully read.
    """
    stream_data = olecf_item.read()

    if self._debug:
      print('_VBA_PROJECT stream data:')
      print(hexdump.Hexdump(stream_data))

    header_struct = self._HEADER.parse(stream_data)

    if self._debug:
      print('Unknown1\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown1))
      print('Unknown2\t\t\t\t\t\t\t: 0x{0:04x}'.format(
          header_struct.unknown2))
      print('Unknown3\t\t\t\t\t\t\t: 0x{0:04x}'.format(
          header_struct.unknown3))
      print('Unknown4\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown4))
      print('Unknown5\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown5))
      print('Unknown6\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown6))
      print('Unknown7\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown7))
      print('Unknown8\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown8))
      print('Unknown9\t\t\t\t\t\t\t: {0:d}'.format(
          header_struct.unknown9))
      print('Number of strings\t\t\t\t\t\t: {0:d}'.format(
          header_struct.number_of_strings))
      print('Unknown11\t\t\t\t\t\t\t: {0:d}'.format(
          header_struct.unknown11))
      print('')

    stream_data_offset = self._HEADER.sizeof()
    for string_index in range(header_struct.number_of_strings):
      string_struct = self._STRING.parse(stream_data[stream_data_offset:])
      value_string = string_struct.string.decode('utf-16-le')

      if self._debug:
        print('String: {0:d} size\t\t\t\t\t\t\t: {1:d}'.format(
            string_index, string_struct.string_size))
        print('String: {0:d}\t\t\t\t\t\t\t: {1:s}'.format(
            string_index, value_string))
        print('Unknown1\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            string_struct.unknown1))
        print('Unknown2\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            string_struct.unknown2))
        print('Unknown3\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            string_struct.unknown3))

      stream_data_offset += 14 + string_struct.string_size

    if self._debug:
      print('')

    return True


class VBACollector(object):
  """Class that defines a Visual Basic for Applications (VBA) collector.

  Attributes:
    steam_found (bool): True if a stream containing VBA was found.
  """

  def __init__(self, debug=False):
    """Initializes a collector.

    Args:
      debug (Optional[bool]): True if debug information should be printed.
    """
    super(VBACollector, self).__init__()
    self._debug = debug

    self.stream_found = False

  def Collect(self, source, output_writer):
    """Collects VBA.

    Args:
      source (str): path of the OLE compound file.
      output_writer (OutputWriter): output writer.
    """
    self.stream_found = False

    olecf_file = pyolecf.file()
    olecf_file.open(source)

    try:
      olecf_macros_project_item = olecf_file.get_item_by_path(
          '\\Macros\\PROJECT')
      if not olecf_macros_project_item:
        return

      stream_data = olecf_macros_project_item.read(
          olecf_macros_project_item.size)
      if self._debug:
        # ID="{%GUID%}"
        # Document=ThisDocument/&H00000000
        # Package={%GUID%}
        # BaseClass=%IDENTIFIER%
        # HelpFile=""
        # Name="Project"
        # HelpContextID="0"
        # VersionCompatible32="393222000"
        # CMG="%IDENTIFIER%"
        # DPB="%IDENTIFIER%"
        # GC="%IDENTIFIER%"

        print('PROJECT stream data:')
        print(stream_data)

      base_class = None
      for line in stream_data.split(b'\n'):
        line = line.strip()
        if line.startswith(b'BaseClass='):
          _, _, base_class = line.rpartition(b'=')

      if base_class:
        olecf_path = '\\Macros\\{0:s}\\f'.format(base_class)
        olecf_f_item = olecf_file.get_item_by_path(olecf_path)
        if olecf_f_item:
          f_stream = FStream(debug=self._debug)
          f_stream.Read(olecf_f_item)

        olecf_path = '\\Macros\\{0:s}\\o'.format(base_class)
        olecf_o_item = olecf_file.get_item_by_path(olecf_path)
        if olecf_o_item:
          o_stream = OStream(debug=self._debug)
          o_stream.Read(olecf_o_item)

      olecf_vba_project_item = olecf_file.get_item_by_path(
          '\\Macros\\VBA\\_VBA_PROJECT')
      if olecf_vba_project_item:
        self.stream_found = True

        vba_project_stream = VBAProjectStream(debug=self._debug)
        vba_project_stream.Read(olecf_vba_project_item)

      # olecf_vba_item.get_sub_item_by_name('dir')
      # MS-OVBA: 2.4.1 Compression and Decompression

    finally:
      olecf_file.close()
