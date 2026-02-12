defmodule OutlookMsg.Pst.Id2 do
  @moduledoc "ID2 association - maps sub-node IDs to index record IDs in PST files"

  defstruct [:id2, :idx_id, :table2_id]

  @doc "Load ID2 associations from a data block"
  def load(_data, _index_map, _encryption_type) do
    %{}
  end

  @doc "Load ID2 map from an index record ID"
  def load_from_id(_data, _index_map, idx2_id, _encryption_type) when idx2_id == 0, do: %{}
  def load_from_id(data, index_map, idx2_id, encryption_type) do
    case Map.get(index_map, idx2_id) do
      nil -> %{}
      index_record ->
        block = read_block(data, index_record)
        block = OutlookMsg.Pst.Encryption.maybe_decrypt(block, encryption_type)
        parse_id2_block(block, data, index_map, encryption_type)
    end
  end

  @doc "Look up an ID2 association"
  def lookup(id2_map, id) do
    Map.get(id2_map, id)
  end

  # Read raw block data from file at index record's offset
  defp read_block(data, %{offset: offset, size: size}) do
    binary_part(data, offset, size)
  end

  # Parse ID2 block - contains 8-byte records for PST97, 16-byte for PST2003
  # Each record: <<id2::32-little, idx_id::32-little, table2_id::32-little>>
  # The block may also be a B-tree page (check signature)
  defp parse_id2_block(block, data, index_map, encryption_type) do
    size = byte_size(block)
    cond do
      size < 8 -> %{}
      true ->
        # Try to parse as flat list of 8-byte records (PST97)
        # or 16-byte records
        record_size = if size >= 16 and rem(size, 16) == 0, do: 16, else: 8
        count = div(size, record_size)
        parse_id2_records(block, 0, count, record_size, %{}, data, index_map, encryption_type)
    end
  end

  defp parse_id2_records(_block, _pos, 0, _rec_size, acc, _data, _index_map, _enc), do: acc
  defp parse_id2_records(block, pos, remaining, 8, acc, data, index_map, enc) do
    <<_::binary-size(pos), id2::32-little, idx_id::32-little, _::binary>> = block
    acc = if id2 != 0, do: Map.put(acc, id2, idx_id), else: acc
    parse_id2_records(block, pos + 8, remaining - 1, 8, acc, data, index_map, enc)
  end
  defp parse_id2_records(block, pos, remaining, 16, acc, data, index_map, enc) do
    <<_::binary-size(pos), id2::32-little, _::32, idx_id::32-little, _::32, _::binary>> = block
    acc = if id2 != 0, do: Map.put(acc, id2, idx_id), else: acc
    parse_id2_records(block, pos + 16, remaining - 1, 16, acc, data, index_map, enc)
  end

  @doc "Read data for an ID2 reference, resolving through the index"
  def read_data(id2_map, id2_key, data, index_map, encryption_type) do
    case lookup(id2_map, id2_key) do
      nil -> nil
      idx_id ->
        case Map.get(index_map, idx_id) do
          nil -> nil
          index_record ->
            block = read_block(data, index_record)
            OutlookMsg.Pst.Encryption.maybe_decrypt(block, encryption_type)
        end
    end
  end
end
