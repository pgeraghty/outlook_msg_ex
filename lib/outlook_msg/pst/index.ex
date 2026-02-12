defmodule OutlookMsg.Pst.Index do
  @moduledoc """
  B-tree index records for PST files.

  Index records map IDs to file offsets/sizes. The B-tree is stored in 512-byte blocks.
  """

  alias OutlookMsg.Warning

  defstruct [:id, :offset, :size, :u1]

  @doc "Load the complete index B-tree from a PST file"
  def load(data, header) do
    {index, _warnings} = load_with_warnings(data, header)
    index
  end

  @doc "Load the index B-tree and return structured warnings for recovery events"
  def load_with_warnings(data, header) do
    case header.version do
      :pst97 -> load_tree(data, header.index1_offset, :pst97, header, MapSet.new())
      :pst2003 -> load_tree(data, header.index1_offset, :pst2003, header, MapSet.new())
    end
  end

  @doc "Look up an index record by ID"
  def lookup(index_map, id) do
    # IDs in PST have lower bits used as flags; mask to get the base ID
    # For lookup, use id as-is since it should already be normalized
    Map.get(index_map, id)
  end

  # Load B-tree traversal
  # Each B-tree page is 512 bytes
  # At offset 0x1F0 (496) within the page: item_count (1 byte)
  # At offset 0x1F1 (497): max_count (1 byte)
  # At offset 0x1F3 (499): level (1 byte) - 0 = leaf, >0 = branch
  defp load_tree(_data, 0, _version, _header, _visited), do: {%{}, []}

  defp load_tree(data, offset, version, header, visited) do
    if MapSet.member?(visited, offset) do
      {%{}, [Warning.new(:pst_branch_loop_detected, "detected recursive PST index branch reference", context: "index_offset=#{offset}")]}
    else
    visited = MapSet.put(visited, offset)
    block =
      if offset < 0 or offset + 512 > byte_size(data) do
        nil
      else
        binary_part(data, offset, 512)
      end

    if is_nil(block) do
      {%{}, []}
    else
      # Read metadata at end of block
      <<_::binary-size(496), item_count::8, max_count::8, _entry_size::8, level::8,
        _::binary>> = block

      _ = max_count

      if level == 0 do
        # Leaf node - parse index records
        {parse_leaf_records(block, item_count, version), []}
      else
        # Branch node - parse child pointers and recurse
        parse_branch_records(data, block, item_count, version, header, visited)
      end
    end
    end
  end

  # PST97 leaf record: 12 bytes each
  # <<id::32-little, offset::32-little, size::16-little, u1::16-little>>
  defp parse_leaf_records(block, count, :pst97) do
    parse_leaf_97(block, 0, count, %{})
  end

  # PST2003 leaf record: 24 bytes each
  # <<id::64-little, offset::64-little, size::16-little, u1::16-little, _pad::32>>
  defp parse_leaf_records(block, count, :pst2003) do
    parse_leaf_2003(block, 0, count, %{})
  end

  defp parse_leaf_97(_block, _pos, 0, acc), do: acc

  defp parse_leaf_97(block, pos, remaining, acc) do
    if pos + 12 > byte_size(block) do
      acc
    else
      <<_::binary-size(pos), id::32-little, offset::32-little, size::16-little, u1::16-little,
        _::binary>> = block

      record = %__MODULE__{id: id, offset: offset, size: size, u1: u1}
      parse_leaf_97(block, pos + 12, remaining - 1, Map.put(acc, id, record))
    end
  end

  defp parse_leaf_2003(_block, _pos, 0, acc), do: acc

  defp parse_leaf_2003(block, pos, remaining, acc) do
    if pos + 24 > byte_size(block) do
      acc
    else
      <<_::binary-size(pos), id::64-little, offset::64-little, size::16-little, u1::16-little,
        _pad::32, _::binary>> = block

      record = %__MODULE__{id: id, offset: offset, size: size, u1: u1}
      parse_leaf_2003(block, pos + 24, remaining - 1, Map.put(acc, id, record))
    end
  end

  # Branch records for PST97: 12 bytes - <<id::32-little, child_offset::32-little, _::32>>
  defp parse_branch_records(data, block, count, :pst97, header, visited) do
    parse_branch_97(data, block, 0, count, %{}, [], header, visited)
  end

  # Branch records for PST2003: 24 bytes - <<id::64-little, child_offset::64-little, _::64>>
  defp parse_branch_records(data, block, count, :pst2003, header, visited) do
    parse_branch_2003(data, block, 0, count, %{}, [], header, visited)
  end

  defp parse_branch_97(_data, _block, _pos, 0, acc, warnings, _header, _visited), do: {acc, warnings}

  defp parse_branch_97(data, block, pos, remaining, acc, warnings, header, visited) do
    if pos + 12 > byte_size(block) do
      {acc, warnings}
    else
      <<_::binary-size(pos), _id::32-little, child_offset::32-little, _::32, _rest::binary>> =
        block

      {child_records, child_warnings} = load_tree(data, child_offset, :pst97, header, visited)
      parse_branch_97(data, block, pos + 12, remaining - 1, Map.merge(acc, child_records), warnings ++ child_warnings, header, visited)
    end
  end

  defp parse_branch_2003(_data, _block, _pos, 0, acc, warnings, _header, _visited), do: {acc, warnings}

  defp parse_branch_2003(data, block, pos, remaining, acc, warnings, header, visited) do
    if pos + 24 > byte_size(block) do
      {acc, warnings}
    else
      <<_::binary-size(pos), _id::64-little, child_offset::64-little, _::64, _rest::binary>> =
        block

      {child_records, child_warnings} = load_tree(data, child_offset, :pst2003, header, visited)

      parse_branch_2003(
        data,
        block,
        pos + 24,
        remaining - 1,
        Map.merge(acc, child_records),
        warnings ++ child_warnings,
        header,
        visited
      )
    end
  end
end
