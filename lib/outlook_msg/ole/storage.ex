defmodule OutlookMsg.Ole.Storage do
  @moduledoc """
  Main entry point for OLE/CFB (Compound File Binary) container parsing.

  Orchestrates `Header`, `Fat`, and `Dirent` modules to provide a high-level
  interface for opening OLE files and reading streams from them.

  ## Usage

      {:ok, storage} = OutlookMsg.Ole.Storage.open("message.msg")
      root = OutlookMsg.Ole.Storage.root(storage)
      children = OutlookMsg.Ole.Storage.children(storage, root)

      # Read a named stream under the root
      {:ok, data} = OutlookMsg.Ole.Storage.stream_by_name(storage, root, "__properties_version1.0")
  """

  alias OutlookMsg.Ole.{Header, Fat, Dirent}

  @end_of_chain 0xFFFFFFFE

  defstruct [:data, :header, :fat, :mini_fat, :dirents, :root, :mini_stream]

  @type t :: %__MODULE__{
          data: binary(),
          header: %Header{},
          fat: %{non_neg_integer() => non_neg_integer()},
          mini_fat: %{non_neg_integer() => non_neg_integer()},
          dirents: [Dirent.t()],
          root: Dirent.t(),
          mini_stream: binary()
        }

  # -------------------------------------------------------------------
  # Public API
  # -------------------------------------------------------------------

  @doc """
  Opens an OLE/CFB file from a file path or raw binary data.

  If the binary begins with the OLE magic bytes (`D0CF11E0`), it is treated
  as raw file data. Otherwise it is interpreted as a file path and read from
  disk.

  Returns `{:ok, %Storage{}}` on success or `{:error, reason}` on failure.
  """
  @spec open(binary()) :: {:ok, t()} | {:error, term()}
  def open(path) when is_binary(path) do
    case path do
      <<0xD0, 0xCF, 0x11, 0xE0, _::binary>> ->
        open_binary(path)

      _ ->
        if File.regular?(path) do
          case File.read(path) do
            {:ok, data} -> open_binary(data)
            {:error, reason} -> {:error, reason}
          end
        else
          # Treat non-file binaries as raw data so corrupted payloads do not
          # route through filesystem calls and raise path-related errors.
          open_binary(path)
        end
    end
  end

  @doc """
  Parses raw OLE/CFB binary data into a `%Storage{}` struct.

  Steps performed:

  1. Parse the 512-byte header
  2. Build the FAT from DIFAT entries
  3. Read the directory stream by following the chain from `header.dir_start_sector`
  4. Parse all directory entries and build the tree
  5. Build the MiniFAT (if present)
  6. Read the mini stream (the root entry's regular stream data)

  Returns `{:ok, %Storage{}}` on success or `{:error, reason}` on failure.
  """
  @spec open_binary(binary()) :: {:ok, t()} | {:error, term()}
  def open_binary(data) when is_binary(data) do
    with {:ok, header} <- Header.parse(data) do
      fat = Fat.build_fat(data, header)

      # Read the directory stream by following the chain from dir_start_sector.
      dir_data = Fat.read_stream(data, header, fat, header.dir_start_sector)
      dirents = Dirent.parse_all(dir_data)
      root = Dirent.build_tree(dirents)

      # Build MiniFAT if one exists.
      mini_fat =
        if header.mini_fat_start != @end_of_chain do
          Fat.build_mini_fat(data, header, fat)
        else
          %{}
        end

      # Read the mini stream: the root entry's stream data via the regular FAT.
      # Streams smaller than mini_cutoff are stored in the mini stream, which is
      # the root directory entry's own stream data.
      mini_stream =
        if root.start_sector != @end_of_chain and root.size > 0 do
          stream_data = Fat.read_stream(data, header, fat, root.start_sector)
          binary_part(stream_data, 0, min(root.size, byte_size(stream_data)))
        else
          <<>>
        end

      {:ok,
       %__MODULE__{
         data: data,
         header: header,
         fat: fat,
         mini_fat: mini_fat,
         dirents: dirents,
         root: root,
         mini_stream: mini_stream
       }}
    end
  end

  @doc """
  Reads the stream data for a given directory entry.

  Streams smaller than `header.mini_cutoff` (typically 4096 bytes) are read
  from the mini stream using the MiniFAT. Larger streams are read directly
  from the file using the regular FAT.

  Returns the stream data as a binary.
  """
  @spec stream(t(), Dirent.t()) :: binary()
  def stream(%__MODULE__{} = storage, %Dirent{} = dirent) do
    if dirent.size < storage.header.mini_cutoff and dirent.type != :root do
      # Small stream: read from mini stream using MiniFAT
      Fat.read_mini_stream(
        storage.mini_stream,
        storage.header,
        storage.mini_fat,
        dirent.start_sector,
        dirent.size
      )
    else
      # Regular stream: read from file using FAT
      data = Fat.read_stream(storage.data, storage.header, storage.fat, dirent.start_sector)
      binary_part(data, 0, min(dirent.size, byte_size(data)))
    end
  end

  @doc """
  Returns the list of children for a given directory entry.
  """
  @spec children(t(), Dirent.t()) :: [Dirent.t()]
  def children(%__MODULE__{}, %Dirent{children: children}), do: children

  @doc """
  Finds a direct child of the given directory entry by name (case-insensitive).

  Returns the matching `%Dirent{}` or `nil` if no child with that name exists.
  """
  @spec find(t(), Dirent.t(), String.t()) :: Dirent.t() | nil
  def find(%__MODULE__{}, %Dirent{children: children}, name) do
    name_down = String.downcase(name)
    Enum.find(children, fn d -> String.downcase(d.name) == name_down end)
  end

  @doc """
  Same as `find/3` but raises an `ArgumentError` if the child is not found.
  """
  @spec find!(t(), Dirent.t(), String.t()) :: Dirent.t()
  def find!(%__MODULE__{} = storage, %Dirent{} = parent, name) do
    case find(storage, parent, name) do
      nil ->
        raise ArgumentError,
              "child #{inspect(name)} not found under #{inspect(parent.name)}"

      child ->
        child
    end
  end

  @doc """
  Convenience function: finds a child by name and reads its stream data.

  Returns `{:ok, binary}` on success or `{:error, :not_found}` if no child
  with the given name exists.
  """
  @spec stream_by_name(t(), Dirent.t(), String.t()) :: {:ok, binary()} | {:error, :not_found}
  def stream_by_name(%__MODULE__{} = storage, %Dirent{} = parent, name) do
    case find(storage, parent, name) do
      nil -> {:error, :not_found}
      child -> {:ok, stream(storage, child)}
    end
  end

  @doc """
  Returns the root directory entry of the storage.
  """
  @spec root(t()) :: Dirent.t()
  def root(%__MODULE__{root: root}), do: root

  @doc """
  Recursively traverses all children of the given directory entry,
  invoking `fun` for each child. For children that are storages or root
  entries, the traversal descends into their children as well.
  """
  @spec each_child(t(), Dirent.t(), (Dirent.t() -> any())) :: :ok
  def each_child(%__MODULE__{} = storage, %Dirent{} = dirent, fun) do
    Enum.each(dirent.children, fn child ->
      fun.(child)

      if child.type in [:storage, :root] do
        each_child(storage, child, fun)
      end
    end)
  end
end
