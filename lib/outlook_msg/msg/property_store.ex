defmodule OutlookMsg.Msg.PropertyStore do
  @moduledoc """
  Core MSG property parsing engine.

  Reads OLE directory entries and extracts MAPI properties from an Outlook MSG
  file. This module handles three layers of property storage:

  1. **Named property mapping** (`__nameid_version1.0`) - Maps pseudo-property
     codes (0x8000+) to actual MAPI named property identifiers (code + GUID).

  2. **Inline properties** (`__properties_version1.0`) - Fixed-size property
     values stored in 16-byte records within a single stream.

  3. **Large properties** (`__substg1.0_XXXXYYYY`) - Variable-size property
     values (strings, binaries, embedded objects) stored as separate streams.

  Based on the `PropertyStore` from the ruby-msg `msg.rb` reference implementation.
  """

  import Bitwise

  alias OutlookMsg.Ole.Storage
  alias OutlookMsg.Ole.Dirent
  alias OutlookMsg.Mapi.{Key, Types, Guids}

  # MAPI property type constants
  @pt_short    0x0002
  @pt_long     0x0003
  @pt_float    0x0004
  @pt_double   0x0005
  @pt_currency 0x0006
  @pt_apptime  0x0007
  @pt_error    0x000A
  @pt_boolean  0x000B
  @pt_object   0x000D
  @pt_int64    0x0014
  @pt_string8  0x001E
  @pt_unicode  0x001F
  @pt_systime  0x0040
  @pt_binary   0x0102

  @mv_flag 0x1000

  # Regex for parsing substg stream names
  @substg_regex ~r/^__substg1\.0_([0-9A-Fa-f]{4})([0-9A-Fa-f]{4})(?:-([0-9A-Fa-f]{8}))?$/

  # -------------------------------------------------------------------
  # Public API
  # -------------------------------------------------------------------

  @doc """
  Parses an OLE directory entry tree into a property map.

  Takes an OLE `Storage` and a `Dirent` (directory entry) representing either
  the root message, an attachment, or a recipient. Returns a map of
  `%Key{} => value` pairs containing all decoded MAPI properties.

  Steps:
  1. Parse named property mapping from `__nameid_version1.0` (if present)
  2. Parse inline properties from `__properties_version1.0`
  3. Parse large properties from `__substg1.0_XXXXYYYY` streams
  4. Remap pseudo-properties (code >= 0x8000) using the nameid mapping
  """
  @spec load(Storage.t(), Dirent.t()) :: %{Key.t() => term()}
  def load(%Storage{} = storage, %Dirent{} = dirent) do
    # Step 1: Parse named property mapping
    nameid =
      case Storage.find(storage, dirent, "__nameid_version1.0") do
        nil -> %{}
        nameid_dirent -> parse_nameid(storage, nameid_dirent)
      end

    # Step 2: Determine prefix size based on whether this is root or sub-object
    # Root message entries have a 32-byte header; attachments/recipients have 8 bytes.
    # We detect root by checking if the dirent is the storage root or has type :root,
    # or by checking if the properties stream is large enough and has the nameid child.
    prefix_size =
      if dirent.type == :root or Storage.find(storage, dirent, "__nameid_version1.0") != nil do
        32
      else
        8
      end

    load_with_prefix(storage, dirent, prefix_size, nameid)
  end

  @doc """
  Parses an OLE directory entry tree into a property map using a pre-computed
  nameid mapping and explicit prefix size.

  This is the primary entry point used when the caller has already parsed
  the nameid mapping (from the root `__nameid_version1.0` storage) and knows
  the property stream header size:
  - `prefix_size` of 32 for root message entries
  - `prefix_size` of 8 for attachments and recipients

  The `nameid` map is the result of `parse_nameid/2` and maps pseudo-property
  codes (0x8000+) to their actual `%Key{}` identifiers.
  """
  @spec load(Storage.t(), Dirent.t(), %{non_neg_integer() => Key.t()}, non_neg_integer()) ::
          %{Key.t() => term()}
  def load(%Storage{} = storage, %Dirent{} = dirent, nameid, prefix_size)
      when is_map(nameid) and is_integer(prefix_size) do
    load_with_prefix(storage, dirent, prefix_size, nameid)
  end

  @doc """
  Loads properties with a specific prefix skip size.

  Use `prefix_size` of 32 for root message entries and 8 for attachments
  and recipients. The nameid mapping is looked up from the dirent's children.
  """
  @spec load_with_prefix(Storage.t(), Dirent.t(), non_neg_integer()) :: %{Key.t() => term()}
  def load_with_prefix(%Storage{} = storage, %Dirent{} = dirent, prefix_size) do
    nameid =
      case Storage.find(storage, dirent, "__nameid_version1.0") do
        nil -> %{}
        nameid_dirent -> parse_nameid(storage, nameid_dirent)
      end

    load_with_prefix(storage, dirent, prefix_size, nameid)
  end

  # -------------------------------------------------------------------
  # Named property mapping
  # -------------------------------------------------------------------

  @doc """
  Parses the named property mapping from the `__nameid_version1.0` storage.

  The nameid storage contains three streams that together define a mapping
  from pseudo-property codes (0x8000+) to actual named property identifiers:

  - `__substg1.0_00020102` - GUID table (16-byte entries, index 2+)
  - `__substg1.0_00030102` - Property entry table (8-byte records)
  - `__substg1.0_00040102` - String name table

  Returns a map of `%{pseudo_code => %Key{code: actual_code, guid: guid}}`.
  """
  @spec parse_nameid(Storage.t(), Dirent.t()) :: %{non_neg_integer() => Key.t()}
  def parse_nameid(%Storage{} = storage, %Dirent{} = nameid_dirent) do
    # Read the three sub-streams
    guid_data = read_nameid_stream(storage, nameid_dirent, "__substg1.0_00020102")
    entry_data = read_nameid_stream(storage, nameid_dirent, "__substg1.0_00030102")
    string_data = read_nameid_stream(storage, nameid_dirent, "__substg1.0_00040102")

    # Pre-defined GUIDs at index 0 and 1
    predefined_guids = %{
      0 => Guids.ps_mapi(),
      1 => Guids.ps_public_strings()
    }

    # Parse additional GUIDs from the GUID stream (index 2+)
    guids = parse_guid_table(guid_data, predefined_guids)

    # Parse 8-byte property entry records
    parse_entry_records(entry_data, string_data, guids)
  end

  # -------------------------------------------------------------------
  # Inline properties
  # -------------------------------------------------------------------

  @doc """
  Parses inline properties from the `__properties_version1.0` stream.

  The stream starts with a header (32 bytes for root messages, 8 bytes for
  attachments/recipients), followed by 16-byte property records.

  Each record contains:
  - 2 bytes: property type
  - 2 bytes: property code
  - 4 bytes: flags
  - 8 bytes: value (inline for fixed-size types, size for variable-size)

  Only fixed-size property types are decoded here. Variable-size properties
  (strings, binaries, objects) are loaded from substg streams.

  Returns a map of `%{Key.t() => decoded_value}`.
  """
  @spec parse_properties(binary(), non_neg_integer(), %{non_neg_integer() => Key.t()}) ::
          %{Key.t() => term()}
  def parse_properties(stream_data, prefix_size, nameid) do
    if byte_size(stream_data) < prefix_size do
      %{}
    else
      body = binary_part(stream_data, prefix_size, byte_size(stream_data) - prefix_size)
      parse_property_records(body, nameid, %{})
    end
  end

  # -------------------------------------------------------------------
  # Substg (large property) parsing
  # -------------------------------------------------------------------

  @doc """
  Parses large properties from `__substg1.0_XXXXYYYY` streams.

  Finds all children of the parent dirent matching the substg naming pattern,
  reads their stream data, and decodes the values based on the property type.

  Handles multi-value properties (type with 0x1000 flag) by collecting indexed
  entries into lists.

  Returns a map of `%{Key.t() => decoded_value}`.
  """
  @spec parse_substg(Storage.t(), Dirent.t(), %{non_neg_integer() => Key.t()}) ::
          %{Key.t() => term()}
  def parse_substg(%Storage{} = storage, %Dirent{} = parent, nameid) do
    children = Storage.children(storage, parent)

    # First pass: collect all substg entries
    entries =
      children
      |> Enum.filter(fn child -> child.type == :stream end)
      |> Enum.map(fn child -> {child, decode_substg_name(child.name)} end)
      |> Enum.filter(fn {_child, result} -> result != nil end)

    # Separate multi-value and single-value properties
    {mv_entries, sv_entries} =
      Enum.split_with(entries, fn {_child, {_code, type, _index}} ->
        (type &&& @mv_flag) != 0
      end)

    # Process single-value properties
    props =
      Enum.reduce(sv_entries, %{}, fn {child, {code, type, _index}}, acc ->
        data = Storage.stream(storage, child)
        key = resolve_key(code, nameid)
        value = decode_substg_value(type, data)
        Map.put(acc, key, value)
      end)

    # Process multi-value properties
    # Group by {code, base_type}, collect indexed values
    mv_groups =
      mv_entries
      |> Enum.filter(fn {_child, {_code, _type, index}} -> index != nil end)
      |> Enum.group_by(
        fn {_child, {code, type, _index}} -> {code, Types.base_type(type)} end,
        fn {child, {_code, type, index}} -> {index, child, type} end
      )

    Enum.reduce(mv_groups, props, fn {{code, base_type}, indexed_entries}, acc ->
      key = resolve_key(code, nameid)

      values =
        indexed_entries
        |> Enum.sort_by(fn {index, _child, _type} -> index end)
        |> Enum.map(fn {_index, child, _type} ->
          data = Storage.stream(storage, child)
          decode_substg_value(base_type, data)
        end)

      Map.put(acc, key, values)
    end)
  end

  # -------------------------------------------------------------------
  # Substg name parsing
  # -------------------------------------------------------------------

  @doc """
  Parses a substg stream name into its components.

  Given a name like `"__substg1.0_0037001F"`, returns `{code, type, index}`
  where:
  - `code` is the property code as an integer (e.g., 0x0037)
  - `type` is the property type as an integer (e.g., 0x001F)
  - `index` is the multi-value index as an integer, or `nil` if absent

  Returns `nil` if the name does not match the expected pattern.
  """
  @spec decode_substg_name(String.t()) :: {non_neg_integer(), non_neg_integer(), non_neg_integer() | nil} | nil
  def decode_substg_name(name) do
    case Regex.run(@substg_regex, name) do
      [_full, code_hex, type_hex] ->
        {String.to_integer(code_hex, 16), String.to_integer(type_hex, 16), nil}

      [_full, code_hex, type_hex, index_hex] ->
        {String.to_integer(code_hex, 16), String.to_integer(type_hex, 16),
         String.to_integer(index_hex, 16)}

      _ ->
        nil
    end
  end

  # -------------------------------------------------------------------
  # Private implementation
  # -------------------------------------------------------------------

  # Load properties combining inline and substg with a given prefix size and nameid map.
  defp load_with_prefix(storage, dirent, prefix_size, nameid) do
    # Step 2: Parse inline properties from __properties_version1.0
    inline_props =
      case Storage.stream_by_name(storage, dirent, "__properties_version1.0") do
        {:ok, prop_data} -> parse_properties(prop_data, prefix_size, nameid)
        {:error, :not_found} -> %{}
      end

    # Step 3: Parse large properties from substg streams
    substg_props = parse_substg(storage, dirent, nameid)

    # Merge: substg values override inline values (they are the actual data
    # for variable-size properties whose inline record only contains the size)
    Map.merge(inline_props, substg_props)
  end

  # Read a named sub-stream from the nameid storage, returning empty binary if not found.
  defp read_nameid_stream(storage, nameid_dirent, name) do
    case Storage.stream_by_name(storage, nameid_dirent, name) do
      {:ok, data} -> data
      {:error, :not_found} -> <<>>
    end
  end

  # Parse the GUID table from the __substg1.0_00020102 stream.
  # Each entry is 16 bytes. Index 0 and 1 are predefined; stream entries start at index 2.
  defp parse_guid_table(<<>>, predefined), do: predefined

  defp parse_guid_table(guid_data, predefined) do
    guid_data
    |> chunk_binary(16)
    |> Enum.with_index(2)
    |> Enum.reduce(predefined, fn {guid, index}, acc ->
      Map.put(acc, index, guid)
    end)
  end

  # Parse 8-byte property entry records from the __substg1.0_00030102 stream.
  defp parse_entry_records(<<>>, _string_data, _guids), do: %{}

  defp parse_entry_records(entry_data, string_data, guids) do
    entry_data
    |> chunk_binary(8)
    |> Enum.with_index()
    |> Enum.reduce(%{}, fn {record, prop_index}, acc ->
      case record do
        <<name_or_id::32-little, flags_and_guid::32-little>> ->
          guid_index = (flags_and_guid >>> 1) &&& 0x7FFF
          is_string = (flags_and_guid &&& 0x0001) == 1
          pseudo_code = 0x8000 + prop_index

          guid = Map.get(guids, guid_index, Guids.ps_mapi())

          code =
            if is_string do
              read_nameid_string(string_data, name_or_id)
            else
              name_or_id
            end

          key = %Key{code: code, guid: guid}
          Map.put(acc, pseudo_code, key)

        _ ->
          # Incomplete record, skip
          acc
      end
    end)
  end

  # Read a string name from the string name table.
  # At the given offset: 4-byte LE length, then that many bytes of UTF-16LE string.
  defp read_nameid_string(string_data, offset) do
    if byte_size(string_data) >= offset + 4 do
      <<_skip::binary-size(offset), length::32-little, rest::binary>> = string_data

      if byte_size(rest) >= length do
        utf16_bytes = binary_part(rest, 0, length)

        case :unicode.characters_to_binary(utf16_bytes, {:utf16, :little}, :utf8) do
          utf8 when is_binary(utf8) -> utf8
          _ -> offset
        end
      else
        offset
      end
    else
      offset
    end
  end

  # Parse 16-byte property records from the inline properties stream body.
  defp parse_property_records(<<>>, _nameid, acc), do: acc

  defp parse_property_records(
         <<type::16-little, code::16-little, _flags::32-little, value::binary-size(8),
           rest::binary>>,
         nameid,
         acc
       ) do
    if fixed_size_type?(type) do
      key = resolve_key(code, nameid)
      decoded = decode_inline_value(type, value)
      parse_property_records(rest, nameid, Map.put(acc, key, decoded))
    else
      # Variable-size type: skip (will be loaded from substg)
      parse_property_records(rest, nameid, acc)
    end
  end

  # If we have fewer than 16 bytes remaining, stop parsing
  defp parse_property_records(_incomplete, _nameid, acc), do: acc

  # Determine if a property type is fixed-size (value stored inline).
  defp fixed_size_type?(type) do
    base = Types.base_type(type)

    base in [
      @pt_short,
      @pt_long,
      @pt_float,
      @pt_double,
      @pt_currency,
      @pt_apptime,
      @pt_error,
      @pt_boolean,
      @pt_int64,
      @pt_systime
    ]
  end

  # Decode an inline fixed-size value from the 8-byte value field.
  defp decode_inline_value(type, <<value::binary-size(8)>>) do
    base = Types.base_type(type)

    case base do
      @pt_short ->
        <<v::16-little, _::binary>> = value
        v

      @pt_long ->
        <<v::32-little-signed, _::binary>> = value
        v

      @pt_float ->
        <<v::32-little-float, _::binary>> = value
        v

      @pt_double ->
        <<v::64-little-float>> = value
        v

      @pt_currency ->
        <<v::64-little-signed>> = value
        v / 10_000.0

      @pt_apptime ->
        <<v::64-little-float>> = value
        v

      @pt_error ->
        <<v::32-little, _::binary>> = value
        v

      @pt_boolean ->
        <<v::16-little, _::binary>> = value
        v != 0

      @pt_int64 ->
        <<v::64-little-signed>> = value
        v

      @pt_systime ->
        Types.decode_value(@pt_systime, value)

      _ ->
        value
    end
  end

  # Decode a substg stream value based on its property type.
  defp decode_substg_value(type, data) do
    base = Types.base_type(type)

    case base do
      @pt_unicode ->
        case :unicode.characters_to_binary(data, {:utf16, :little}, :utf8) do
          utf8 when is_binary(utf8) -> strip_trailing_nulls(utf8)
          _ -> data
        end

      @pt_string8 ->
        strip_trailing_nulls(data)

      @pt_binary ->
        data

      @pt_object ->
        data

      _ ->
        # For other types stored in substg, try to decode them
        Types.decode_value(base, data)
    end
  end

  # Resolve a property code to a Key. If code >= 0x8000 and a nameid mapping
  # exists, use the mapped key. Otherwise, create a standard PS_MAPI key.
  defp resolve_key(code, nameid) when code >= 0x8000 do
    case Map.get(nameid, code) do
      nil -> Key.new(code)
      %Key{} = key -> key
    end
  end

  defp resolve_key(code, _nameid) do
    Key.new(code)
  end

  # Split a binary into fixed-size chunks.
  defp chunk_binary(binary, chunk_size) do
    do_chunk_binary(binary, chunk_size, [])
  end

  defp do_chunk_binary(binary, chunk_size, acc) when byte_size(binary) >= chunk_size do
    <<chunk::binary-size(chunk_size), rest::binary>> = binary
    do_chunk_binary(rest, chunk_size, [chunk | acc])
  end

  defp do_chunk_binary(_binary, _chunk_size, acc) do
    Enum.reverse(acc)
  end

  # Strip trailing null bytes from a UTF-8 string.
  defp strip_trailing_nulls(string) do
    String.trim_trailing(string, <<0>>)
  end
end
