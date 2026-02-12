defmodule OutlookMsg.Pst.Descriptor do
  @moduledoc """
  Descriptor B-tree records for PST hierarchy.

  Descriptors represent the folder/message hierarchy in a PST file. Each descriptor
  maps to an index entry and has parent-child relationships.
  """

  alias OutlookMsg.Warning

  defstruct [:desc_id, :idx_id, :idx2_id, :parent_desc_id, :children]

  @doc "Load the complete descriptor B-tree from a PST file"
  def load(data, header) do
    {desc, _warnings} = load_with_warnings(data, header)
    desc
  end

  @doc "Load descriptor tree and return structured warnings for recovery events"
  def load_with_warnings(data, header) do
    records =
      case header.version do
        :pst97 -> load_tree(data, header.index2_offset, :pst97, header, MapSet.new())
        :pst2003 -> load_tree(data, header.index2_offset, :pst2003, header, MapSet.new())
      end

    case records do
      {rows, warnings} -> {build_hierarchy(rows), warnings}
      rows when is_list(rows) -> {build_hierarchy(rows), []}
    end
  end

  @doc "Build parent-child hierarchy from flat descriptor list"
  def build_hierarchy(records) do
    # Create a map of desc_id -> descriptor
    by_id = Map.new(records, fn r -> {r.desc_id, %{r | children: []}} end)

    # Build children lists
    by_id =
      Enum.reduce(records, by_id, fn record, acc ->
        parent_id = record.parent_desc_id

        if parent_id != 0 and Map.has_key?(acc, parent_id) do
          parent = acc[parent_id]
          Map.put(acc, parent_id, %{parent | children: [record.desc_id | parent.children]})
        else
          acc
        end
      end)

    # Reverse children lists to maintain order
    Map.new(by_id, fn {id, desc} ->
      {id, %{desc | children: Enum.reverse(desc.children)}}
    end)
  end

  # B-tree traversal, same structure as Index
  defp load_tree(_data, 0, _version, _header, _visited), do: {[], []}

  defp load_tree(data, offset, version, header, visited) do
    if MapSet.member?(visited, offset) do
      {[], [Warning.new(:pst_branch_loop_detected, "detected recursive PST descriptor branch reference", context: "descriptor_offset=#{offset}")]}
    else
      visited = MapSet.put(visited, offset)
    block =
      if offset < 0 or offset + 512 > byte_size(data) do
        nil
      else
        binary_part(data, offset, 512)
      end

    if is_nil(block) do
      {[], []}
    else
      <<_::binary-size(496), item_count::8, _max_count::8, _entry_size::8, level::8,
        _::binary>> = block

      if level == 0 do
        {parse_leaf_records(block, item_count, version), []}
      else
        parse_branch_records(data, block, item_count, version, header, visited)
      end
    end
    end
  end

  # PST97 descriptor leaf: 16 bytes
  # <<desc_id::32-little, idx_id::32-little, idx2_id::32-little, parent_desc_id::32-little>>
  defp parse_leaf_records(block, count, :pst97) do
    if count <= 0 do
      []
    else
      Enum.reduce(0..(count - 1), [], fn i, acc ->
        pos = i * 16

        if pos + 16 > byte_size(block) do
          acc
        else
          <<_::binary-size(pos), desc_id::32-little, idx_id::32-little, idx2_id::32-little,
            parent_desc_id::32-little, _::binary>> = block

          acc ++
            [
              %__MODULE__{
                desc_id: desc_id,
                idx_id: idx_id,
                idx2_id: idx2_id,
                parent_desc_id: parent_desc_id,
                children: []
              }
            ]
        end
      end)
    end
  end

  # PST2003 descriptor leaf: 32 bytes
  # <<desc_id::64-little, idx_id::64-little, idx2_id::64-little, parent_desc_id::32-little, _pad::32>>
  defp parse_leaf_records(block, count, :pst2003) do
    if count <= 0 do
      []
    else
      Enum.reduce(0..(count - 1), [], fn i, acc ->
        pos = i * 32

        if pos + 32 > byte_size(block) do
          acc
        else
          <<_::binary-size(pos), desc_id::64-little, idx_id::64-little, idx2_id::64-little,
            parent_desc_id::32-little, _pad::32, _::binary>> = block

          acc ++
            [
              %__MODULE__{
                desc_id: desc_id,
                idx_id: idx_id,
                idx2_id: idx2_id,
                parent_desc_id: parent_desc_id,
                children: []
              }
            ]
        end
      end)
    end
  end

  # Branch records - same size, contain child offsets
  defp parse_branch_records(data, block, count, :pst97, header, visited) do
    if count <= 0 do
      {[], []}
    else
      Enum.reduce(0..(count - 1), {[], []}, fn i, {rows, warns} ->
        pos = i * 12

        if pos + 12 > byte_size(block) do
          {rows, warns}
        else
          <<_::binary-size(pos), _id::32-little, child_offset::32-little, _::32, _rest::binary>> =
            block

          {child_rows, child_warnings} = load_tree(data, child_offset, :pst97, header, visited)
          {rows ++ child_rows, warns ++ child_warnings}
        end
      end)
    end
  end

  defp parse_branch_records(data, block, count, :pst2003, header, visited) do
    if count <= 0 do
      {[], []}
    else
      Enum.reduce(0..(count - 1), {[], []}, fn i, {rows, warns} ->
        pos = i * 24

        if pos + 24 > byte_size(block) do
          {rows, warns}
        else
          <<_::binary-size(pos), _id::64-little, child_offset::64-little, _::64, _rest::binary>> =
            block

          {child_rows, child_warnings} = load_tree(data, child_offset, :pst2003, header, visited)
          {rows ++ child_rows, warns ++ child_warnings}
        end
      end)
    end
  end
end
