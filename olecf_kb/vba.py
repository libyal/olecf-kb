# -*- coding: utf-8 -*-
"""Visual Basic for Applications (VBA) collector."""

from __future__ import print_function

import construct
import pyolecf

from olecf_kb import hexdump


class FStream(object):
  """Class that defines a f stream."""

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
    # TODO: add support for read with optional size argument.
    stream_data = olecf_item.read(olecf_item.size)

    if self._debug:
      print(u'f stream data:')
      print(hexdump.Hexdump(stream_data))


class OStream(object):
  """Class that defines an o stream."""

  _ENTRY_PART1 = construct.Struct(
      u'entry_part1',
      construct.ULInt32(u'unknown1'),
      construct.ULInt32(u'unknown2'),
      construct.ULInt32(u'unknown3'),
      construct.ULInt32(u'unknown4'),
      construct.ULInt32(u'unknown5'),
      construct.ULInt32(u'unknown6'),
      construct.ULInt32(u'unknown7'),
      construct.CString(u'data'))

  _ENTRY_PART2 = construct.Struct(
      u'entry_part2',
      construct.ULInt32(u'unknown7'),
      construct.ULInt32(u'unknown8'),
      construct.ULInt32(u'unknown9'),
      construct.ULInt32(u'unknown10'),
      construct.ULInt32(u'unknown11'),
      construct.CString(u'font_name'))

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
    # TODO: add support for read with optional size argument.
    stream_data = olecf_item.read(olecf_item.size)

    stream_offset = 0
    while stream_offset < olecf_item.size:
      entry_part1_struct = self._ENTRY_PART1.parse(stream_data[stream_offset:])

      entry_part_size = (7 * 4) + len(entry_part1_struct.data)
      padding_size = entry_part_size % 4
      if padding_size != 0:
        padding_size = 4 - padding_size

      next_stream_offset = stream_offset + entry_part_size + padding_size

      entry_part2_struct = self._ENTRY_PART2.parse(stream_data[next_stream_offset:])

      entry_part_size = (5 * 4) + len(entry_part2_struct.font_name)
      padding_size = entry_part_size % 4
      if padding_size != 0:
        padding_size = 4 - padding_size

      next_stream_offset += entry_part_size + padding_size

      if self._debug:
        print(u'o stream entry data:')
        print(hexdump.Hexdump(stream_data[stream_offset:next_stream_offset]))

      # TODO: add debug info.
      if self._debug:
        print(u'Unknown1\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part1_struct.unknown1))
        print(u'Unknown2\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part1_struct.unknown2))
        print(u'Unknown3\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part1_struct.unknown3))
        print(u'Unknown4\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part1_struct.unknown4))
        print(u'Unknown5\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part1_struct.unknown5))
        print(u'Unknown6\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part1_struct.unknown6))
        print(u'Data\t\t\t\t\t\t\t\t: {0:s}'.format(entry_part1_struct.data))
        # TODO: alignment padding.
        print(u'Unknown7\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part2_struct.unknown7))
        print(u'Unknown8\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part2_struct.unknown8))
        print(u'Unknown9\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part2_struct.unknown9))
        print(u'Unknown10\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part2_struct.unknown10))
        print(u'Unknown11\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            entry_part2_struct.unknown11))
        print(u'Font name\t\t\t\t\t\t\t: {0:s}'.format(
            entry_part2_struct.font_name))
        # TODO: alignment padding.
        print(u'')

      stream_offset = next_stream_offset


class VBAProjectStream(object):
  """Class that defines a _VBA_PROJECT (Preformance Cache) stream."""

  _HEADER = construct.Struct(
      u'header',
      construct.ULInt32(u'unknown1'),
      construct.ULInt16(u'unknown2'),
      construct.ULInt16(u'unknown3'),
      construct.ULInt32(u'unknown4'),
      construct.ULInt32(u'unknown5'),
      construct.ULInt32(u'unknown6'),
      construct.ULInt32(u'unknown7'),
      construct.ULInt32(u'unknown8'),
      construct.ULInt16(u'unknown9'),
      construct.ULInt16(u'number_of_strings'),
      construct.ULInt16(u'unknown11'))

  _STRING = construct.Struct(
      u'string',
      construct.ULInt16(u'string_size'),  # Does not include the end-of-string character.
      construct.Bytes(
          u'string',
          lambda ctx: ctx.string_size),
      construct.ULInt32(u'unknown1'),
      construct.ULInt32(u'unknown2'),
      construct.ULInt32(u'unknown3'))

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
    # TODO: add support for read with optional size argument.
    stream_data = olecf_item.read(olecf_item.size)

    if self._debug:
      print(u'_VBA_PROJECT stream data:')
      print(hexdump.Hexdump(stream_data))

    header_struct = self._HEADER.parse(stream_data)

    if self._debug:
      print(u'Unknown1\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown1))
      print(u'Unknown2\t\t\t\t\t\t\t: 0x{0:04x}'.format(
          header_struct.unknown2))
      print(u'Unknown3\t\t\t\t\t\t\t: 0x{0:04x}'.format(
          header_struct.unknown3))
      print(u'Unknown4\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown4))
      print(u'Unknown5\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown5))
      print(u'Unknown6\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown6))
      print(u'Unknown7\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown7))
      print(u'Unknown8\t\t\t\t\t\t\t: 0x{0:08x}'.format(
          header_struct.unknown8))
      print(u'Unknown9\t\t\t\t\t\t\t: {0:d}'.format(
          header_struct.unknown9))
      print(u'Number of strings\t\t\t\t\t\t: {0:d}'.format(
          header_struct.number_of_strings))
      print(u'Unknown11\t\t\t\t\t\t\t: {0:d}'.format(
          header_struct.unknown11))
      print(u'')

    stream_data_offset = self._HEADER.sizeof()
    for string_index in range(header_struct.number_of_strings):
      string_struct = self._STRING.parse(stream_data[stream_data_offset:])
      value_string = string_struct.string.decode(u'utf-16-le')

      if self._debug:
        print(u'String: {0:d} size\t\t\t\t\t\t\t: {1:d}'.format(
            string_index, string_struct.string_size))
        print(u'String: {0:d}\t\t\t\t\t\t\t: {1:s}'.format(
            string_index, value_string))
        print(u'Unknown1\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            string_struct.unknown1))
        print(u'Unknown2\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            string_struct.unknown2))
        print(u'Unknown3\t\t\t\t\t\t\t: 0x{0:08x}'.format(
            string_struct.unknown3))

      stream_data_offset += 14 + string_struct.string_size

    if self._debug:
      print(u'')


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
      # TODO: replace by:
      # olecf_file.get_item_by_path(u'\\Root Entry\\Macros\\VBA\\_VBA_PROJECT')

      olecf_root_item = olecf_file.get_root_item()
      if not olecf_root_item:
        return

      olecf_macros_item = olecf_root_item.get_sub_item_by_name(u'Macros')
      if not olecf_macros_item:
        return

      olecf_macros_project_item = olecf_macros_item.get_sub_item_by_name(u'PROJECT')
      if not olecf_macros_project_item:
        return

      stream_data = olecf_macros_project_item.read(olecf_macros_project_item.size)
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

        print(u'PROJECT stream data:')
        print(stream_data)

      base_clase = None
      for line in stream_data.split(b'\n'):
        line = line.strip()
        if line.startswith(b'BaseClass='):
          _, _, base_clase = line.rpartition(b'=')

      olecf_vba_item = olecf_macros_item.get_sub_item_by_name(u'VBA')
      if not olecf_vba_item:
        return

      olecf_vba_project_item = olecf_vba_item.get_sub_item_by_name(
          u'_VBA_PROJECT')
      if not olecf_vba_project_item:
        return

      if base_clase:
        # olecf_file.get_item_by_path('\\Root Entry\\Macros\\{0:s}\\f'.format(base_class))

        olecf_base_class_item = olecf_macros_item.get_sub_item_by_name(base_clase)
        if olecf_base_class_item:
          olecf_f_item = olecf_base_class_item.get_sub_item_by_name(u'f')
          if olecf_f_item:
            f_stream = FStream(debug=self._debug)
            f_stream.Read(olecf_f_item)

          olecf_o_item = olecf_base_class_item.get_sub_item_by_name(u'o')
          if olecf_o_item:
            o_stream = OStream(debug=self._debug)
            o_stream.Read(olecf_o_item)

      self.stream_found = True

      vba_project_stream = VBAProjectStream(debug=self._debug)
      vba_project_stream.Read(olecf_vba_project_item)

    finally:
      olecf_file.close()
