# -*- coding: utf-8 -*-
"""Visual Basic for Applications (VBA) collector."""

import uuid

import pyolecf

from dtfabric import errors as dtfabric_errors

from olecfrc import data_format
from olecfrc import errors
from olecfrc import hexdump


class FStream(data_format.BinaryDataFormat):
  """Class that defines a f stream."""

  _DEFINITION_FILE = 'vba.yaml'

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

    Raises:
      ParseError: if the stream data could not be parsed.
    """
    stream_data = olecf_item.read()

    data_type_map = self._GetDataTypeMap('f_stream_header')

    try:
      header_struct = data_type_map.MapByteStream(stream_data)
    except (
        dtfabric_errors.ByteStreamTooSmallError,
        dtfabric_errors.MappingError) as exception:
      raise errors.ParseError(exception)

    stream_offset = data_type_map.GetByteSize()

    if self._debug:
      print('f stream header data:')
      print(hexdump.Hexdump(stream_data[:stream_offset]))

    if self._debug:
      print(f'Unknown1\t\t\t\t\t\t\t: 0x{header_struct.unknown1:08x}')
      print(f'Unknown2\t\t\t\t\t\t\t: 0x{header_struct.unknown2:08x}')
      print(f'Unknown3\t\t\t\t\t\t\t: 0x{header_struct.unknown3:08x}')
      print(f'Unknown4\t\t\t\t\t\t\t: 0x{header_struct.unknown4:08x}')
      print(f'Unknown5\t\t\t\t\t\t\t: 0x{header_struct.unknown5:08x}')
      print(f'Unknown6\t\t\t\t\t\t\t: 0x{header_struct.unknown6:08x}')
      print(f'Unknown7\t\t\t\t\t\t\t: 0x{header_struct.unknown7:08x}')
      print(f'Unknown8\t\t\t\t\t\t\t: 0x{header_struct.unknown8:08x}')
      print(f'Unknown9\t\t\t\t\t\t\t: 0x{header_struct.unknown9:08x}')
      print(f'Unknown10\t\t\t\t\t\t\t: 0x{header_struct.unknown10:08x}')
      print(f'Unknown11\t\t\t\t\t\t\t: 0x{header_struct.unknown11:08x}')

      # CLSID of StdFont: 0be35203-8f91-11ce-9de3-00aa004bb851
      uuid_value = uuid.UUID(bytes_le=header_struct.unknown12)
      print(f'Unknown12\t\t\t\t\t\t\t: {uuid_value!s}')

      print('')

    data_type_map = self._GetDataTypeMap('f_stream_entry')

    while stream_offset < olecf_item.size:
      try:
        entry_struct = data_type_map.MapByteStream(stream_data[stream_offset:])
      except (
          dtfabric_errors.ByteStreamTooSmallError,
          dtfabric_errors.MappingError) as exception:
        raise errors.ParseError(exception)

      next_stream_offset = stream_offset + 2 + entry_struct.size + 2

      if self._debug:
        print('f stream entry data:')
        print(hexdump.Hexdump(stream_data[stream_offset:next_stream_offset]))

      if self._debug:
        print(f'Unknown7\t\t\t\t\t\t\t: 0x{entry_struct.unknown7:08x}')
        print(f'Unknown8\t\t\t\t\t\t\t: 0x{entry_struct.unknown8:08x}')
        print(f'Unknown9\t\t\t\t\t\t\t: 0x{entry_struct.unknown9:04x}')

        print(f'Size\t\t\t\t\t\t\t\t: {entry_struct.size:d}')

        print(f'Unknown1\t\t\t\t\t\t\t: 0x{entry_struct.unknown1:08x}')
        print(f'Unknown2\t\t\t\t\t\t\t: 0x{entry_struct.unknown2:08x}')
        print(f'Unknown3\t\t\t\t\t\t\t: {entry_struct.unknown3:d}')
        print((f'O stream entry size\t\t\t\t\t\t: '
               f'{entry_struct.o_stream_entry_size:0d}'))
        print((f'O stream entry index\t\t\t\t\t\t: '
               f'{entry_struct.o_stream_entry_index:d}'))
        print(f'Unknown6\t\t\t\t\t\t\t: 0x{entry_struct.unknown6:04x}')

        # TODO: fix this.
        try:
          variable_name = entry_struct.variable_name.decode('cp1252')
          print(f'Variable name\t\t\t\t\t\t\t: {variable_name:s}')
        except UnicodeEncodeError:
          pass

        print('')

      stream_offset = next_stream_offset

    return True


class OStream(data_format.BinaryDataFormat):
  """Class that defines an o stream."""

  _DEFINITION_FILE = 'vba.yaml'

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

    Raises:
      ParseError: if the stream data could not be parsed.
    """
    stream_data = olecf_item.read()

    data_type_map1 = self._GetDataTypeMap('o_entry_part1')
    data_type_map2 = self._GetDataTypeMap('o_entry_part2')

    stream_offset = 0
    while stream_offset < olecf_item.size:
      try:
        entry_part1_struct = data_type_map1.MapByteStream(
            stream_data[stream_offset:])
      except (
          dtfabric_errors.ByteStreamTooSmallError,
          dtfabric_errors.MappingError) as exception:
        raise errors.ParseError(exception)

      entry_part_size = (7 * 4) + len(entry_part1_struct.data) + 1
      padding_size = entry_part_size % 4
      if padding_size != 0:
        padding_size = 4 - padding_size

      next_stream_offset = stream_offset + entry_part_size + padding_size

      try:
        entry_part2_struct = data_type_map2.MapByteStream(
            stream_data[next_stream_offset:])
      except (
          dtfabric_errors.ByteStreamTooSmallError,
          dtfabric_errors.MappingError) as exception:
        raise errors.ParseError(exception)

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
        print(f'Unknown1\t\t\t\t\t\t\t: 0x{entry_part1_struct.unknown1:08x}')
        print(f'Unknown2\t\t\t\t\t\t\t: 0x{entry_part1_struct.unknown2:08x}')
        print(f'Unknown3\t\t\t\t\t\t\t: 0x{entry_part1_struct.unknown3:08x}')
        print(f'Unknown4\t\t\t\t\t\t\t: 0x{entry_part1_struct.unknown4:08x}')
        data_size = entry_part1_struct.data_size & 0x7fffffff
        print((f'Data size\t\t\t\t\t\t\t: {data_size:d} '
               f'(0x{entry_part1_struct.data_size:08x})'))
        print(f'Unknown6\t\t\t\t\t\t\t: 0x{entry_part1_struct.unknown6:08x}')
        print(f'Data\t\t\t\t\t\t\t\t: {entry_part1_struct.data:s}')
        # TODO: alignment padding.
        print(f'Unknown7\t\t\t\t\t\t\t: 0x{entry_part2_struct.unknown7:08x}')
        print(f'Unknown8\t\t\t\t\t\t\t: 0x{entry_part2_struct.unknown8:08x}')
        print(f'Unknown9\t\t\t\t\t\t\t: 0x{entry_part2_struct.unknown9:08x}')
        print(f'Unknown10\t\t\t\t\t\t\t: 0x{entry_part2_struct.unknown10:08x}')
        print(f'Unknown11\t\t\t\t\t\t\t: 0x{entry_part2_struct.unknown11:08x}')
        print(f'Font name\t\t\t\t\t\t\t: {entry_part2_struct.font_name:s}')
        # TODO: alignment padding.
        print('')

      stream_offset = next_stream_offset

    return True


class VBAProjectStream(data_format.BinaryDataFormat):
  """Class that defines a _VBA_PROJECT (Performance Cache) stream."""

  _DEFINITION_FILE = 'vba.yaml'

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

    Raises:
      ParseError: if the stream data could not be parsed.
    """
    stream_data = olecf_item.read()

    if self._debug:
      print('_VBA_PROJECT stream data:')
      print(hexdump.Hexdump(stream_data))

    data_type_map = self._GetDataTypeMap('project_stream_header')

    try:
      header_struct = data_type_map.MapByteStream(stream_data)
    except (dtfabric_errors.ByteStreamTooSmallError,
            dtfabric_errors.MappingError) as exception:
      raise errors.ParseError(exception)

    stream_data_offset = data_type_map.GetByteSize()

    if self._debug:
      print(f'Unknown1\t\t\t\t\t\t\t: 0x{header_struct.unknown1:08x}')
      print(f'Unknown2\t\t\t\t\t\t\t: 0x{header_struct.unknown2:04x}')
      print(f'Unknown3\t\t\t\t\t\t\t: 0x{header_struct.unknown3:04x}')
      print(f'Unknown4\t\t\t\t\t\t\t: 0x{header_struct.unknown4:08x}')
      print(f'Unknown5\t\t\t\t\t\t\t: 0x{header_struct.unknown5:08x}')
      print(f'Unknown6\t\t\t\t\t\t\t: 0x{header_struct.unknown6:08x}')
      print(f'Unknown7\t\t\t\t\t\t\t: 0x{header_struct.unknown7:08x}')
      print(f'Unknown8\t\t\t\t\t\t\t: 0x{header_struct.unknown8:08x}')
      print(f'Unknown9\t\t\t\t\t\t\t: {header_struct.unknown9:d}')
      print((f'Number of strings\t\t\t\t\t\t: '
             f'{header_struct.number_of_strings:d}'))
      print(f'Unknown11\t\t\t\t\t\t\t: {header_struct.unknown11:d}')
      print('')

    for string_index in range(header_struct.number_of_strings):
      data_type_map = self._GetDataTypeMap('project_stream_string')

      try:
        string_struct = data_type_map.MapByteStream(
            stream_data[stream_data_offset:])
      except (
          dtfabric_errors.ByteStreamTooSmallError,
          dtfabric_errors.MappingError) as exception:
        raise errors.ParseError(exception)

      value_string = string_struct.string.decode('utf-16-le')

      if self._debug:
        print((f'String: {string_index:d} size\t\t\t\t\t\t\t: '
               f'{string_struct.string_size:d}'))
        print(f'String: {string_index:d}\t\t\t\t\t\t\t: {value_string:s}')
        print(f'Unknown1\t\t\t\t\t\t\t: 0x{string_struct.unknown1:08x}')
        print(f'Unknown2\t\t\t\t\t\t\t: 0x{string_struct.unknown2:08x}')
        print(f'Unknown3\t\t\t\t\t\t\t: 0x{string_struct.unknown3:08x}')

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
    # TODO: remove this once output_writer is used.
    _ = output_writer

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
        olecf_f_item = olecf_file.get_item_by_path(
            f'\\Macros\\{base_class:s}\\f')
        if olecf_f_item:
          f_stream = FStream(debug=self._debug)
          f_stream.Read(olecf_f_item)

        olecf_o_item = olecf_file.get_item_by_path(
            f'\\Macros\\{base_class:s}\\o')
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
