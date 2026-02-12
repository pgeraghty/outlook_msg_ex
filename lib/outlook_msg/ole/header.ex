defmodule OutlookMsg.Ole.Header do
  @moduledoc """
  Parses the 512-byte OLE/CFB (Compound File Binary) header
  according to the MS-CFB specification.

  Header layout (offsets in bytes):

      0-7     magic (must be D0CF11E0A1B11AE1)
      8-23    CLSID (ignored)
      24-25   minor version
      26-27   major version
      28-29   byte order (0xFFFE = little-endian)
      30-31   sector shift
      32-33   mini sector shift
      34-39   reserved (6 bytes, must be 0)
      40-43   directory sector count (0 for v3)
      44-47   FAT sector count
      48-51   directory start sector
      52-55   transaction signature (ignored)
      56-59   mini stream cutoff size
      60-63   MiniFAT start sector
      64-67   MiniFAT sector count
      68-71   DIFAT start sector
      72-75   DIFAT sector count
      76-511  first 109 DIFAT entries (each 4 bytes LE)
  """

  @magic <<0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1>>
  @header_size 512
  @free_sector 0xFFFFFFFF
  @end_of_chain 0xFFFFFFFE

  defstruct [
    :magic,
    :minor_version,
    :major_version,
    :byte_order,
    :sector_shift,
    :mini_shift,
    :dir_sector_count,
    :fat_sector_count,
    :dir_start_sector,
    :mini_cutoff,
    :mini_fat_start,
    :mini_fat_count,
    :difat_start,
    :difat_count,
    :difat_entries,
    :sector_size,
    :mini_sector_size
  ]

  @doc """
  Parses the first 512 bytes of OLE/CFB binary data into a `%Header{}` struct.

  Returns `{:ok, %Header{}}` on success or `{:error, reason}` on failure.
  """
  @spec parse(binary()) :: {:ok, %__MODULE__{}} | {:error, String.t()}
  def parse(<<
        @magic,
        _clsid::binary-size(16),
        minor_version::little-unsigned-16,
        major_version::little-unsigned-16,
        byte_order::little-unsigned-16,
        sector_shift::little-unsigned-16,
        mini_shift::little-unsigned-16,
        _reserved::binary-size(6),
        dir_sector_count::little-unsigned-32,
        fat_sector_count::little-unsigned-32,
        dir_start_sector::little-unsigned-32,
        _transaction_sig::little-unsigned-32,
        mini_cutoff::little-unsigned-32,
        mini_fat_start::little-unsigned-32,
        mini_fat_count::little-unsigned-32,
        difat_start::little-unsigned-32,
        difat_count::little-unsigned-32,
        difat_raw::binary-size(436),
        _rest::binary
      >>) do
    difat_entries = parse_difat_entries(difat_raw)

    header = %__MODULE__{
      magic: @magic,
      minor_version: minor_version,
      major_version: major_version,
      byte_order: byte_order,
      sector_shift: sector_shift,
      mini_shift: mini_shift,
      dir_sector_count: dir_sector_count,
      fat_sector_count: fat_sector_count,
      dir_start_sector: dir_start_sector,
      mini_cutoff: mini_cutoff,
      mini_fat_start: mini_fat_start,
      mini_fat_count: mini_fat_count,
      difat_start: difat_start,
      difat_count: difat_count,
      difat_entries: difat_entries,
      sector_size: Bitwise.bsl(1, sector_shift),
      mini_sector_size: Bitwise.bsl(1, mini_shift)
    }

    validate_header(header)
  end

  def parse(<<@magic, rest::binary>>) when byte_size(rest) + 8 < @header_size do
    {:error, "header too short: expected at least #{@header_size} bytes"}
  end

  def parse(<<_other_magic::binary-size(8), _rest::binary>>) do
    {:error, "invalid magic number: not an OLE/CFB file"}
  end

  def parse(data) when is_binary(data) and byte_size(data) < 8 do
    {:error, "header too short: expected at least #{@header_size} bytes"}
  end

  def parse(_data) do
    {:error, "invalid input"}
  end

  # -- Private helpers -------------------------------------------------------

  defp validate_header(%__MODULE__{} = header) do
    cond do
      header.byte_order != 0xFFFE ->
        {:error, "invalid byte order: expected 0xFFFE, got 0x#{Integer.to_string(header.byte_order, 16)}"}

      header.major_version not in [3, 4] ->
        {:error, "unsupported major version: #{header.major_version}"}

      header.major_version == 3 and header.sector_shift != 9 ->
        {:error, "invalid sector shift for v3: expected 9, got #{header.sector_shift}"}

      header.major_version == 4 and header.sector_shift != 12 ->
        {:error, "invalid sector shift for v4: expected 12, got #{header.sector_shift}"}

      header.mini_shift != 6 ->
        {:error, "invalid mini sector shift: expected 6, got #{header.mini_shift}"}

      header.mini_cutoff != 4096 ->
        {:error, "invalid mini stream cutoff: expected 4096, got #{header.mini_cutoff}"}

      true ->
        {:ok, header}
    end
  end

  defp parse_difat_entries(raw) do
    parse_difat_entries(raw, [])
    |> Enum.reverse()
  end

  defp parse_difat_entries(<<>>, acc), do: acc

  defp parse_difat_entries(<<entry::little-unsigned-32, rest::binary>>, acc) do
    case entry do
      @free_sector -> parse_difat_entries(rest, acc)
      @end_of_chain -> parse_difat_entries(rest, acc)
      sector -> parse_difat_entries(rest, [sector | acc])
    end
  end
end
