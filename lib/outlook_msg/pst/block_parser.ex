defmodule OutlookMsg.Pst.BlockParser do
  @moduledoc "Parse property blocks from PST data blocks"

  import Bitwise

  alias OutlookMsg.Pst.Id2
  alias OutlookMsg.Mapi.Key

  @type1_sig 0xBC  # bcec - RawPropertyStore (variable-length)
  @type2_sig 0x7C  # 7cec - RawPropertyStoreTable (fixed-column table)

  @doc """
  Parse a property block into a map of Key => value.
  Takes raw block data, id2 map, full file data, index map, and encryption type.
  """
  def parse(block_data, id2_map, data, index_map, encryption_type) do
    if byte_size(block_data) < 4 do
      %{}
    else
      <<sig::8, _::binary>> = block_data
      case sig do
        @type1_sig -> parse_type1(block_data, id2_map, data, index_map, encryption_type)
        @type2_sig -> parse_type2(block_data, id2_map, data, index_map, encryption_type)
        _ -> %{}
      end
    end
  end

  # Type 1 (0xBC/bcec): RawPropertyStore - variable length property records
  # Header: <<sig::8, _::8, offset_table_offset::16-little>>
  # Then at offset_table_offset: a list of 2-byte LE offsets
  defp parse_type1(block_data, id2_map, data, index_map, enc) do
    <<_sig::8, _::8, offset_start::16-little, _::binary>> = block_data

    # Read the offset table - starts with count
    if offset_start < 4 or offset_start >= byte_size(block_data),
      do: %{},
      else: parse_type1_props(block_data, offset_start, id2_map, data, index_map, enc)
  end

  defp parse_type1_props(block_data, offset_start, id2_map, data, index_map, enc) do
    # Property records are 8 bytes each: <<type::16-little, code::16-little, value_or_offset::32-little>>
    # They start right after the 4-byte header
    prop_start = 4
    prop_data = binary_part(block_data, prop_start, offset_start - prop_start)
    count = div(byte_size(prop_data), 8)

    if count <= 0 do
      %{}
    else
      Enum.reduce(0..(count - 1), %{}, fn i, acc ->
        offset = i * 8

        if offset + 8 <= byte_size(prop_data) do
          <<_::binary-size(offset), type::16-little, code::16-little, value_raw::binary-size(4), _::binary>> = prop_data
          key = Key.new(code)

          case decode_property_value(type, value_raw, block_data, id2_map, data, index_map, enc) do
            nil -> acc
            value -> Map.put(acc, key, value)
          end
        else
          acc
        end
      end)
    end
  end

  # Type 2 (0x7C/7cec): RawPropertyStoreTable - fixed-column table
  # Used for recipient and attachment tables
  defp parse_type2(block_data, id2_map, data, index_map, enc) do
    # Parse table header to get column definitions
    <<_sig::8, _::8, offset_start::16-little, _::binary>> = block_data

    if offset_start < 4 or offset_start >= byte_size(block_data) do
      %{}
    else
      parse_type1_props(block_data, offset_start, id2_map, data, index_map, enc)
    end
  end

  @doc "Parse a property store table and return a list of property maps (one per row)"
  def parse_table(block_data, id2_map, data, index_map, encryption_type) do
    # For tables, we parse as a single row for simplicity
    [parse(block_data, id2_map, data, index_map, encryption_type)]
  end

  # Decode a property value based on its type
  defp decode_property_value(type, value_raw, block_data, id2_map, data, index_map, enc) do
    base_type = type &&& 0x0FFF

    case base_type do
      0x0002 -> # PT_SHORT
        <<val::16-little, _::binary>> = value_raw
        val
      0x0003 -> # PT_LONG
        <<val::32-little-signed>> = value_raw
        val
      0x000B -> # PT_BOOLEAN
        <<val::16-little, _::binary>> = value_raw
        val != 0
      0x001E -> # PT_STRING8
        <<offset_or_val::32-little>> = value_raw
        read_indirect_string(offset_or_val, block_data, id2_map, data, index_map, enc)
      0x001F -> # PT_UNICODE
        <<offset_or_val::32-little>> = value_raw
        case read_indirect_data(offset_or_val, block_data, id2_map, data, index_map, enc) do
          nil -> nil
          bin ->
            case :unicode.characters_to_binary(bin, {:utf16, :little}) do
              {:error, _, _} -> bin
              {:incomplete, result, _} -> result
              result when is_binary(result) -> String.trim_trailing(result, <<0>>)
            end
        end
      0x0040 -> # PT_SYSTIME
        case read_indirect_data_or_inline(value_raw, 8, block_data, id2_map, data, index_map, enc) do
          <<val::64-little>> -> OutlookMsg.Ole.Types.parse_filetime(<<val::64-little>>)
          _ -> nil
        end
      0x0102 -> # PT_BINARY
        <<offset_or_val::32-little>> = value_raw
        read_indirect_data(offset_or_val, block_data, id2_map, data, index_map, enc)
      _ ->
        <<val::32-little>> = value_raw
        val
    end
  end

  # Read indirect data from block or ID2
  defp read_indirect_data(offset, block_data, id2_map, data, index_map, enc) do
    cond do
      offset == 0 -> nil
      offset < byte_size(block_data) ->
        # Read from within the block using offset table
        # Simplified: treat as inline or look up in id2
        nil
      true ->
        Id2.read_data(id2_map, offset, data, index_map, enc)
    end
  end

  defp read_indirect_string(offset, block_data, id2_map, data, index_map, enc) do
    case read_indirect_data(offset, block_data, id2_map, data, index_map, enc) do
      nil -> nil
      bin -> String.trim_trailing(bin, <<0>>)
    end
  end

  defp read_indirect_data_or_inline(value_raw, expected_size, _block, _id2, _data, _idx, _enc) when byte_size(value_raw) >= expected_size do
    binary_part(value_raw, 0, expected_size)
  end
  defp read_indirect_data_or_inline(value_raw, _expected_size, block_data, id2_map, data, index_map, enc) do
    <<offset::32-little>> = value_raw
    read_indirect_data(offset, block_data, id2_map, data, index_map, enc)
  end
end
