defmodule OutlookMsg.Pst do
  @moduledoc "PST (Personal Storage Table) file parser for Outlook .pst files"

  alias OutlookMsg.Pst.{Header, Encryption, Index, Descriptor, BlockParser, Id2, Item}
  alias OutlookMsg.Mapi.PropertySet
  alias OutlookMsg.Warning

  defstruct [
    :data,              # Raw file binary
    :header,            # Parsed PST header
    :index,             # Index B-tree (id -> index record)
    :descriptors,       # Descriptor hierarchy (desc_id -> descriptor)
    :encryption_type,   # 0 = none, 1 = compressible
    :root_desc_id,      # Root descriptor ID (usually 0x21 = 33)
    warnings: []
  ]

  @root_desc_id 0x21  # Standard root descriptor ID

  @doc "Open a PST file from path or binary"
  def open(path_or_binary) do
    data =
      case path_or_binary do
        <<0x21, 0x42, 0x44, 0x4E, _::binary>> ->
          path_or_binary

        bin when is_binary(bin) ->
          if File.regular?(bin) do
            case File.read(bin) do
              {:ok, d} -> d
              {:error, reason} -> {:error, reason}
            end
          else
            bin
          end
      end

    case data do
      {:error, _} = err -> err
      data when is_binary(data) ->
        with {:ok, header} <- Header.parse(data) do
          {index, warnings1} =
            try do
              Index.load_with_warnings(data, header)
            rescue
              e -> {%{}, [Warning.new(:pst_index_parse_failed, "continuing with empty index", context: Exception.message(e))]}
            end

          warnings1 =
            if warnings1 == [] and header.index1_offset != 0 and map_size(index) == 0 do
              [Warning.new(:pst_index_parse_failed, "index tree was empty despite non-zero root offset", context: "offset=#{header.index1_offset}")]
            else
              warnings1
            end

          {descriptors, warnings2} =
            try do
              Descriptor.load_with_warnings(data, header)
            rescue
              e -> {%{}, [Warning.new(:pst_descriptor_parse_failed, "continuing with empty descriptor map", context: Exception.message(e))]}
            end

          warnings2 =
            if warnings2 == [] and header.index2_offset != 0 and map_size(descriptors) == 0 do
              [Warning.new(:pst_descriptor_parse_failed, "descriptor tree was empty despite non-zero root offset", context: "offset=#{header.index2_offset}")]
            else
              warnings2
            end

          {:ok, %__MODULE__{
            data: data,
            header: header,
            index: index,
            descriptors: descriptors,
            encryption_type: header.encryption_type,
            root_desc_id: @root_desc_id,
            warnings: warnings1 ++ warnings2
          }}
        end
    end
  end

  @doc "Get the root item (top-level folder)"
  def root(%__MODULE__{} = pst) do
    load_item(pst, @root_desc_id)
  end

  @doc "Load a specific item by descriptor ID"
  def load_item(%__MODULE__{} = pst, desc_id) do
    case Map.get(pst.descriptors, desc_id) do
      nil -> nil
      desc ->
        properties = load_properties(pst, desc)
        property_set = PropertySet.new(properties)
        type = Item.detect_type(property_set)
        item = Item.new(desc, property_set, type)
        %{item | pst_ref: pst}
    end
  end

  @doc "Get children items of a folder item"
  def children(%__MODULE__{} = pst, %Item{} = item) do
    desc = item.desc
    child_ids = desc.children || []

    child_ids
    |> Enum.map(fn child_id -> load_item(pst, child_id) end)
    |> Enum.reject(&is_nil/1)
  end

  @doc "Stream all items in the PST file (depth-first traversal)"
  def items(%__MODULE__{} = pst) do
    Stream.resource(
      fn -> [root(pst)] end,
      fn
        [] -> {:halt, nil}
        [nil | rest] -> {[], rest}
        [item | rest] ->
          child_items = children(pst, item)
          {[item], child_items ++ rest}
      end,
      fn _ -> :ok end
    )
  end

  @doc "Stream only message items (not folders)"
  def messages(%__MODULE__{} = pst) do
    items(pst)
    |> Stream.filter(fn item -> Item.message?(item) end)
  end

  @doc "Stream only folder items"
  def folders(%__MODULE__{} = pst) do
    items(pst)
    |> Stream.filter(fn item -> Item.folder?(item) end)
  end

  @doc "Walk the folder tree, calling function for each item with depth"
  def walk(%__MODULE__{} = pst, fun) do
    case root(pst) do
      nil -> :ok
      root_item -> do_walk(pst, root_item, 0, fun)
    end
  end

  defp do_walk(pst, item, depth, fun) do
    fun.(item, depth)
    children(pst, item)
    |> Enum.each(fn child -> do_walk(pst, child, depth + 1, fun) end)
  end

  # Load properties for a descriptor
  defp load_properties(%__MODULE__{} = pst, desc) do
    # Get the main data block through the index
    case Map.get(pst.index, desc.idx_id) do
      nil -> %{}
      index_record ->
        # Read the data block
        block_data = read_block(pst, index_record)
        block_data = Encryption.maybe_decrypt(block_data, pst.encryption_type)

        # Load ID2 associations for this descriptor
        id2_map = Id2.load_from_id(pst.data, pst.index, desc.idx2_id, pst.encryption_type)

        # Parse the property block
        BlockParser.parse(block_data, id2_map, pst.data, pst.index, pst.encryption_type)
    end
  end

  defp read_block(%__MODULE__{data: data}, %{offset: offset, size: size}) do
    if offset + size <= byte_size(data) do
      binary_part(data, offset, size)
    else
      <<>>
    end
  end
end
