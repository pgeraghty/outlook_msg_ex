defmodule OutlookMsg.Ole.Dirent do
  @moduledoc """
  Parses and reconstructs OLE/CFB directory entries.

  Each directory entry is a 128-byte record that describes either a storage
  (folder), stream (file), or the root entry in the compound file. Entries
  reference each other through a red-black tree structure using sibling and
  child SIDs (stream IDs), which this module flattens into an ordinary
  parent-children tree.

  Directory entry layout (128 bytes):

      0-63    name (UTF-16LE, up to 64 bytes)
      64-65   name_size (2 bytes LE, byte count including null terminator)
      66      type (1 byte)
      67      color (1 byte, 0 = red, 1 = black)
      68-71   left_sid (4 bytes LE)
      72-75   right_sid (4 bytes LE)
      76-79   child_sid (4 bytes LE)
      80-95   clsid (16 bytes)
      96-99   state_bits (4 bytes LE)
      100-107 create_time (8 bytes FILETIME)
      108-115 modify_time (8 bytes FILETIME)
      116-119 start_sector (4 bytes LE)
      120-127 size (8 bytes LE; for v3 non-root streams only the low 4 bytes matter)
  """

  @nostream 0xFFFFFFFF
  @entry_size 128

  defstruct [
    :name,
    :type,
    :color,
    :left_sid,
    :right_sid,
    :child_sid,
    :clsid,
    :state_bits,
    :create_time,
    :modify_time,
    :start_sector,
    :size,
    :sid,
    children: []
  ]

  @type t :: %__MODULE__{
          name: String.t(),
          type: :empty | :storage | :stream | :lock_bytes | :property | :root,
          color: :red | :black,
          left_sid: non_neg_integer(),
          right_sid: non_neg_integer(),
          child_sid: non_neg_integer(),
          clsid: binary(),
          state_bits: non_neg_integer(),
          create_time: binary(),
          modify_time: binary(),
          start_sector: non_neg_integer(),
          size: non_neg_integer(),
          sid: non_neg_integer(),
          children: [t()]
        }

  # -------------------------------------------------------------------
  # Public API
  # -------------------------------------------------------------------

  @doc """
  Parses a single 128-byte directory entry binary into a `%Dirent{}` struct.

  `sid` is the zero-based index of this entry within the directory stream.
  """
  @spec parse(binary(), non_neg_integer()) :: t()
  def parse(<<
              name_raw::binary-size(64),
              name_size::little-unsigned-16,
              type_byte::unsigned-8,
              color_byte::unsigned-8,
              left_sid::little-unsigned-32,
              right_sid::little-unsigned-32,
              child_sid::little-unsigned-32,
              clsid::binary-size(16),
              state_bits::little-unsigned-32,
              create_time::binary-size(8),
              modify_time::binary-size(8),
              start_sector::little-unsigned-32,
              size::little-unsigned-64
            >>,
            sid) do
    # name_size is the byte count of the name including the null terminator.
    # Clamp to 64 bytes (the maximum that fits in the field).
    usable_name_size = min(name_size, 64)

    # Grab only the relevant portion of the 64-byte field, then strip the
    # trailing UTF-16LE null terminator(s) and decode to UTF-8.
    name_bytes = binary_part(name_raw, 0, usable_name_size)
    name = decode_utf16le_name(name_bytes)

    %__MODULE__{
      name: name,
      type: decode_type(type_byte),
      color: decode_color(color_byte),
      left_sid: left_sid,
      right_sid: right_sid,
      child_sid: child_sid,
      clsid: clsid,
      state_bits: state_bits,
      create_time: create_time,
      modify_time: modify_time,
      start_sector: start_sector,
      size: size,
      sid: sid,
      children: []
    }
  end

  @doc """
  Parses the complete directory stream binary into a list of `%Dirent{}` structs.

  The binary is split into 128-byte chunks; each chunk becomes one entry.
  Entries whose type is `:empty` are excluded from the result.
  """
  @spec parse_all(binary()) :: [t()]
  def parse_all(data) when is_binary(data) do
    data
    |> split_entries(0, [])
    |> Enum.reverse()
    |> Enum.reject(fn d -> d.type == :empty end)
  end

  @doc """
  Reconstructs the parent-children tree from the flat list of directory entries.

  Starting from the root entry (SID 0), each entry's `child_sid` points into a
  red-black binary search tree whose nodes are linked via `left_sid` and
  `right_sid`. An in-order traversal of that tree produces the ordered list of
  children for the parent entry.

  Returns the root `%Dirent{}` with the `children` field recursively populated.
  """
  @spec build_tree([t()]) :: t()
  def build_tree(dirents) when is_list(dirents) do
    index = Map.new(dirents, fn d -> {d.sid, d} end)
    root = Map.fetch!(index, 0)
    populate_children(root, index)
  end

  @doc """
  Finds a direct child of `dirent` whose name matches `name` (case-insensitive).

  Returns `nil` if no matching child is found.
  """
  @spec find_child(t(), String.t()) :: t() | nil
  def find_child(%__MODULE__{children: children}, name) when is_binary(name) do
    target = String.downcase(name)

    Enum.find(children, fn child ->
      String.downcase(child.name) == target
    end)
  end

  @doc """
  Same as `find_child/2` but raises an error when the child is not found.
  """
  @spec find_child!(t(), String.t()) :: t()
  def find_child!(%__MODULE__{} = dirent, name) when is_binary(name) do
    case find_child(dirent, name) do
      nil ->
        raise ArgumentError,
              "child #{inspect(name)} not found under #{inspect(dirent.name)}"

      child ->
        child
    end
  end

  # -------------------------------------------------------------------
  # Private helpers
  # -------------------------------------------------------------------

  # Split the raw binary into 128-byte chunks and parse each one.
  defp split_entries(<<chunk::binary-size(@entry_size), rest::binary>>, sid, acc) do
    dirent = parse(chunk, sid)
    split_entries(rest, sid + 1, [dirent | acc])
  end

  defp split_entries(_rest, _sid, acc), do: acc

  # Recursively populate the children field for storages and the root.
  defp populate_children(%__MODULE__{child_sid: @nostream} = dirent, _index) do
    dirent
  end

  defp populate_children(%__MODULE__{child_sid: child_sid} = dirent, index) do
    children =
      child_sid
      |> collect_inorder(index)
      |> Enum.map(fn child -> populate_children(child, index) end)

    %{dirent | children: children}
  end

  # In-order traversal of the red-black tree rooted at `sid`.
  defp collect_inorder(@nostream, _index), do: []

  defp collect_inorder(sid, index) do
    case Map.get(index, sid) do
      nil ->
        []

      node ->
        collect_inorder(node.left_sid, index) ++
          [node] ++
          collect_inorder(node.right_sid, index)
    end
  end

  # Decode the entry type byte to an atom.
  defp decode_type(0), do: :empty
  defp decode_type(1), do: :storage
  defp decode_type(2), do: :stream
  defp decode_type(3), do: :lock_bytes
  defp decode_type(4), do: :property
  defp decode_type(5), do: :root
  defp decode_type(_), do: :empty

  # Decode the color byte to an atom.
  defp decode_color(0), do: :red
  defp decode_color(1), do: :black
  defp decode_color(_), do: :black

  # Decode a UTF-16LE name field, stripping trailing null characters.
  defp decode_utf16le_name(<<>>) do
    ""
  end

  defp decode_utf16le_name(name_bytes) do
    stripped = strip_utf16le_nulls(name_bytes)

    case :unicode.characters_to_binary(stripped, {:utf16, :little}, :utf8) do
      utf8 when is_binary(utf8) -> utf8
      _ -> ""
    end
  end

  # Strip trailing UTF-16LE null code units (0x00 0x00 pairs).
  defp strip_utf16le_nulls(binary) do
    size = byte_size(binary)
    aligned = size - rem(size, 2)
    do_strip_nulls(binary, aligned)
  end

  defp do_strip_nulls(binary, size) when size >= 2 do
    offset = size - 2

    case binary do
      <<_::binary-size(offset), 0x00, 0x00, _::binary>> ->
        do_strip_nulls(binary, offset)

      _ ->
        binary_part(binary, 0, size)
    end
  end

  defp do_strip_nulls(_binary, _size), do: <<>>
end
