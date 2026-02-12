defmodule OutlookMsg.Ole.Fat do
  @moduledoc """
  Handles FAT (File Allocation Table) and MiniFAT sector chain logic for
  OLE/CFB (Compound Binary Format) files.

  The FAT is essentially a linked list stored as an array: each entry at
  index *i* contains the index of the next sector in whatever chain sector
  *i* belongs to, or one of the special sentinel values defined below.
  """

  # ── Special sector constants ──────────────────────────────────────────
  @endofchain 0xFFFFFFFE
  @freesect   0xFFFFFFFF
  @fatsect    0xFFFFFFFD
  @difsect    0xFFFFFFFC

  # Safety limit to prevent infinite loops when following chains.
  @max_chain_length 1_000_000

  # ── Public API ────────────────────────────────────────────────────────

  @doc """
  Builds the FAT (a map of `%{sector_index => next_sector}`) by reading
  every FAT sector referenced by the header's DIFAT entries and, when the
  header contains more than 109 DIFAT entries, by following the DIFAT
  chain on disk.

  ## Parameters

    * `bin`    – the complete OLE file as a binary
    * `header` – a parsed `OutlookMsg.Ole.Header` struct

  ## Returns

  A map `%{non_neg_integer() => non_neg_integer()}` mapping each sector
  index to the next sector in its chain (or a sentinel value).
  """
  @spec build_fat(binary(), struct()) :: %{non_neg_integer() => non_neg_integer()}
  def build_fat(bin, header) do
    sector_size = header.sector_size

    # 1. Collect all FAT sector locations from DIFAT entries.
    #    The header embeds up to 109 DIFAT entries directly.
    initial_fat_sectors =
      header.difat_entries
      |> Enum.reject(&(&1 == @freesect))

    # 2. If there are more FAT sectors than fit in the header, follow the
    #    DIFAT chain to collect the remaining FAT sector locations.
    fat_sector_locations =
      if header.difat_count > 0 do
        additional = read_difat_chain(bin, header)
        initial_fat_sectors ++ additional
      else
        initial_fat_sectors
      end

    # 3. Read and parse every FAT sector into a flat map.
    entries_per_sector = div(sector_size, 4)

    fat_sector_locations
    |> Enum.with_index()
    |> Enum.reduce(%{}, fn {sector_num, sector_idx}, acc ->
      offset = sector_offset(sector_num, sector_size)
      <<_skip::binary-size(offset), sector_data::binary-size(sector_size), _rest::binary>> = bin

      parse_sector_entries(sector_data, entries_per_sector)
      |> Enum.with_index()
      |> Enum.reduce(acc, fn {value, entry_idx}, inner_acc ->
        global_index = sector_idx * entries_per_sector + entry_idx
        Map.put(inner_acc, global_index, value)
      end)
    end)
  end

  @doc """
  Builds the MiniFAT by following the MiniFAT chain (starting at
  `header.mini_fat_start`) through the regular FAT.

  ## Parameters

    * `bin`    – the complete OLE file as a binary
    * `header` – a parsed header struct
    * `fat`    – the FAT map as returned by `build_fat/2`

  ## Returns

  A map with the same shape as the FAT: `%{mini_sector_index => next}`.
  """
  @spec build_mini_fat(binary(), struct(), map()) :: %{non_neg_integer() => non_neg_integer()}
  def build_mini_fat(_bin, header, _fat) when header.mini_fat_start == @endofchain do
    %{}
  end

  def build_mini_fat(bin, header, fat) do
    sector_size = header.sector_size
    entries_per_sector = div(sector_size, 4)

    mini_fat_sectors = chain(fat, header.mini_fat_start)

    mini_fat_sectors
    |> Enum.with_index()
    |> Enum.reduce(%{}, fn {sector_num, sector_idx}, acc ->
      offset = sector_offset(sector_num, sector_size)
      <<_skip::binary-size(offset), sector_data::binary-size(sector_size), _rest::binary>> = bin

      parse_sector_entries(sector_data, entries_per_sector)
      |> Enum.with_index()
      |> Enum.reduce(acc, fn {value, entry_idx}, inner_acc ->
        global_index = sector_idx * entries_per_sector + entry_idx
        Map.put(inner_acc, global_index, value)
      end)
    end)
  end

  @doc """
  Follows a sector chain through the given FAT (or MiniFAT) starting at
  `start_sector`, collecting sector numbers in order until `ENDOFCHAIN`
  (0xFFFFFFFE) is reached.

  Returns an ordered list of sector numbers.

  ## Examples

      iex> fat = %{0 => 3, 3 => 5, 5 => 7, 7 => 0xFFFFFFFE}
      iex> OutlookMsg.Ole.Fat.chain(fat, 0)
      [0, 3, 5, 7]

  """
  @spec chain(map(), non_neg_integer()) :: [non_neg_integer()]
  def chain(_fat, @endofchain), do: []
  def chain(_fat, @freesect), do: []

  def chain(fat, start_sector) do
    do_chain(fat, start_sector, [], MapSet.new(), 0)
    |> Enum.reverse()
  end

  # ── Stream reading ────────────────────────────────────────────────────

  @doc """
  Reads a complete regular-sized stream by following the sector chain in
  the FAT and concatenating each sector's raw bytes.

  ## Parameters

    * `bin`          – the complete OLE file binary
    * `header`       – parsed header struct (needs `:sector_size`)
    * `fat`          – the FAT map
    * `start_sector` – first sector of the stream

  ## Returns

  A single binary containing the concatenated stream data.
  """
  @spec read_stream(binary(), struct(), map(), non_neg_integer()) :: binary()
  def read_stream(bin, header, fat, start_sector) do
    sector_size = header.sector_size

    chain(fat, start_sector)
    |> Enum.map(fn sector_num ->
      offset = sector_offset(sector_num, sector_size)
      binary_part(bin, offset, sector_size)
    end)
    |> IO.iodata_to_binary()
  end

  @doc """
  Reads data from the mini stream using the MiniFAT.

  The mini stream is itself a contiguous binary that has already been
  assembled from the root directory entry's regular stream.  Mini sectors
  are laid out sequentially inside that binary.

  ## Parameters

    * `mini_stream`  – the assembled mini stream binary
    * `header`       – parsed header struct (needs `:mini_sector_size`)
    * `mini_fat`     – the MiniFAT map
    * `start_sector` – first mini sector of the stream
    * `size`         – the actual byte length of the stream (the result
      is truncated to this size since the last mini sector may be only
      partially used)

  ## Returns

  A binary of exactly `size` bytes.
  """
  @spec read_mini_stream(binary(), struct(), map(), non_neg_integer(), non_neg_integer()) ::
          binary()
  def read_mini_stream(mini_stream, header, mini_fat, start_sector, size) do
    mini_sector_size = header.mini_sector_size

    data =
      chain(mini_fat, start_sector)
      |> Enum.map(fn sector_num ->
        offset = sector_num * mini_sector_size
        binary_part(mini_stream, offset, mini_sector_size)
      end)
      |> IO.iodata_to_binary()

    binary_part(data, 0, size)
  end

  # ── Helpers ───────────────────────────────────────────────────────────

  @doc """
  Calculates the absolute byte offset in the file for a given sector
  number.  Sector 0 starts immediately after the 512-byte header, so
  the formula is `(sector_num + 1) * sector_size`.
  """
  @spec sector_offset(non_neg_integer(), non_neg_integer()) :: non_neg_integer()
  def sector_offset(sector_num, sector_size) do
    (sector_num + 1) * sector_size
  end

  # ── Private ───────────────────────────────────────────────────────────

  # Follows the chain, guarding against cycles and runaway lengths.
  defp do_chain(_fat, @endofchain, acc, _seen, _count), do: acc
  defp do_chain(_fat, @freesect, acc, _seen, _count), do: acc
  defp do_chain(_fat, @fatsect, acc, _seen, _count), do: acc
  defp do_chain(_fat, @difsect, acc, _seen, _count), do: acc

  defp do_chain(_fat, _sector, acc, _seen, count) when count >= @max_chain_length do
    acc
  end

  defp do_chain(fat, sector, acc, seen, count) do
    if MapSet.member?(seen, sector) do
      # Cycle detected – stop.
      acc
    else
      next = Map.get(fat, sector, @endofchain)
      do_chain(fat, next, [sector | acc], MapSet.put(seen, sector), count + 1)
    end
  end

  # Reads additional FAT sector locations by following the DIFAT chain
  # that starts at `header.difat_start`.  Each DIFAT sector contains
  # `(sector_size / 4) - 1` FAT sector pointers followed by a single
  # uint32 that is the next DIFAT sector number (or ENDOFCHAIN).
  defp read_difat_chain(bin, header) do
    sector_size = header.sector_size
    entries_per_difat_sector = div(sector_size, 4) - 1

    do_read_difat_chain(bin, header.difat_start, sector_size, entries_per_difat_sector, [])
    |> List.flatten()
    |> Enum.reject(&(&1 == @freesect))
  end

  defp do_read_difat_chain(_bin, @endofchain, _sector_size, _entries_per, acc) do
    Enum.reverse(acc)
  end

  defp do_read_difat_chain(_bin, @freesect, _sector_size, _entries_per, acc) do
    Enum.reverse(acc)
  end

  defp do_read_difat_chain(bin, sector_num, sector_size, entries_per, acc) do
    offset = sector_offset(sector_num, sector_size)
    <<_skip::binary-size(offset), sector_data::binary-size(sector_size), _rest::binary>> = bin

    {entries, next_sector} = parse_difat_sector(sector_data, entries_per)
    do_read_difat_chain(bin, next_sector, sector_size, entries_per, [entries | acc])
  end

  # Parses a single DIFAT sector.  The last 4 bytes are the pointer to
  # the next DIFAT sector; all preceding uint32 values are FAT sector
  # locations.
  defp parse_difat_sector(sector_data, entries_count) do
    entries =
      for i <- 0..(entries_count - 1) do
        binary_part(sector_data, i * 4, 4)
        |> :binary.decode_unsigned(:little)
      end

    next_offset = entries_count * 4
    next_sector =
      binary_part(sector_data, next_offset, 4)
      |> :binary.decode_unsigned(:little)

    {entries, next_sector}
  end

  # Parses a sector's raw bytes into a list of uint32 entries (LE).
  defp parse_sector_entries(sector_data, count) do
    for i <- 0..(count - 1) do
      binary_part(sector_data, i * 4, 4)
      |> :binary.decode_unsigned(:little)
    end
  end
end
