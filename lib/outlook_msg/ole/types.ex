defmodule OutlookMsg.Ole.Types do
  @moduledoc """
  OLE type helpers for GUID, LPWSTR, and FILETIME parsing.

  Provides functions to parse and encode common OLE/COM data types found
  in Outlook MSG files and other OLE structured storage formats.
  """

  # The number of seconds between the FILETIME epoch (1601-01-01 00:00:00 UTC)
  # and the Unix epoch (1970-01-01 00:00:00 UTC).
  @filetime_epoch_offset 11_644_473_600

  # FILETIME uses 100-nanosecond intervals.
  @filetime_ticks_per_second 10_000_000

  # -------------------------------------------------------------------
  # GUID parsing and formatting
  # -------------------------------------------------------------------

  @doc """
  Parses a 16-byte binary GUID into its standard string representation.

  GUIDs are stored in mixed-endian format:
    - data1: 4 bytes, little-endian
    - data2: 2 bytes, little-endian
    - data3: 2 bytes, little-endian
    - data4: 2 bytes, big-endian
    - data5: 6 bytes, big-endian

  Returns `{:ok, guid_string}` where `guid_string` is in the format
  `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`, or `{:error, reason}`.

  ## Examples

      iex> binary = <<0x33, 0x22, 0x11, 0x00, 0x55, 0x44, 0x77, 0x66,
      ...>            0x88, 0x99, 0xAA, 0xBB, 0xCC, 0xDD, 0xEE, 0xFF>>
      iex> OutlookMsg.Ole.Types.parse_guid(binary)
      {:ok, "{00112233-4455-6677-8899-AABBCCDDEEFF}"}
  """
  @spec parse_guid(binary()) :: {:ok, String.t()} | {:error, String.t()}
  def parse_guid(<<
        data1::little-unsigned-integer-size(32),
        data2::little-unsigned-integer-size(16),
        data3::little-unsigned-integer-size(16),
        data4::big-unsigned-integer-size(16),
        data5::big-unsigned-integer-size(48)
      >>) do
    guid_string =
      "{" <>
        hex_pad(data1, 8) <>
        "-" <>
        hex_pad(data2, 4) <>
        "-" <>
        hex_pad(data3, 4) <>
        "-" <>
        hex_pad(data4, 4) <>
        "-" <>
        hex_pad(data5, 12) <>
        "}"

    {:ok, guid_string}
  end

  def parse_guid(<<_::binary-size(16), _rest::binary>>) do
    {:error, "extra bytes after 16-byte GUID"}
  end

  def parse_guid(_) do
    {:error, "expected exactly 16 bytes for GUID"}
  end

  @doc """
  Formats a GUID struct, tuple, or binary to its string representation.

  Accepts:
    - A 16-byte binary (delegates to `parse_guid/1`)
    - A 5-element tuple `{data1, data2, data3, data4, data5}` with integer values
    - A string already in GUID format (returned as-is)

  Returns the formatted GUID string directly.
  """
  @spec format_guid(binary() | tuple() | String.t()) :: String.t()
  def format_guid(<<_::binary-size(16)>> = binary) do
    {:ok, result} = parse_guid(binary)
    result
  end

  def format_guid({data1, data2, data3, data4, data5})
      when is_integer(data1) and is_integer(data2) and is_integer(data3) and
             is_integer(data4) and is_integer(data5) do
    "{" <>
      hex_pad(data1, 8) <>
      "-" <>
      hex_pad(data2, 4) <>
      "-" <>
      hex_pad(data3, 4) <>
      "-" <>
      hex_pad(data4, 4) <>
      "-" <>
      hex_pad(data5, 12) <>
      "}"
  end

  def format_guid("{" <> _ = guid_string) when is_binary(guid_string) do
    guid_string
  end

  # -------------------------------------------------------------------
  # GUID encoding
  # -------------------------------------------------------------------

  @doc """
  Converts a GUID string like `"{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}"` back
  to a 16-byte mixed-endian binary.

  Returns `{:ok, binary}` or `{:error, reason}`.

  ## Examples

      iex> OutlookMsg.Ole.Types.encode_guid("{00112233-4455-6677-8899-AABBCCDDEEFF}")
      {:ok, <<0x33, 0x22, 0x11, 0x00, 0x55, 0x44, 0x77, 0x66,
              0x88, 0x99, 0xAA, 0xBB, 0xCC, 0xDD, 0xEE, 0xFF>>}
  """
  @spec encode_guid(String.t()) :: {:ok, binary()} | {:error, String.t()}
  def encode_guid(guid_string) when is_binary(guid_string) do
    stripped = String.trim(guid_string)

    with "{" <> rest <- stripped,
         true <- String.ends_with?(rest, "}"),
         hex_body = String.trim_trailing(rest, "}"),
         [d1, d2, d3, d4, d5] <- String.split(hex_body, "-"),
         {:ok, data1} <- hex_to_int(d1),
         {:ok, data2} <- hex_to_int(d2),
         {:ok, data3} <- hex_to_int(d3),
         {:ok, data4} <- hex_to_int(d4),
         {:ok, data5} <- hex_to_int(d5) do
      {:ok,
       <<
         data1::little-unsigned-integer-size(32),
         data2::little-unsigned-integer-size(16),
         data3::little-unsigned-integer-size(16),
         data4::big-unsigned-integer-size(16),
         data5::big-unsigned-integer-size(48)
       >>}
    else
      _ -> {:error, "invalid GUID string format: #{inspect(guid_string)}"}
    end
  end

  def encode_guid(_), do: {:error, "expected a binary GUID string"}

  # -------------------------------------------------------------------
  # LPWSTR parsing
  # -------------------------------------------------------------------

  @doc """
  Parses a null-terminated UTF-16LE encoded string to a UTF-8 Elixir string.

  Strips any trailing null characters (U+0000) from the UTF-16LE data before
  converting to UTF-8.

  Returns `{:ok, string}` or `{:error, reason}`.

  ## Examples

      iex> binary = <<0x48, 0x00, 0x69, 0x00, 0x00, 0x00>>
      iex> OutlookMsg.Ole.Types.parse_lpwstr(binary)
      {:ok, "Hi"}
  """
  @spec parse_lpwstr(binary()) :: {:ok, String.t()} | {:error, String.t()}
  def parse_lpwstr(binary) when is_binary(binary) do
    stripped = strip_utf16le_nulls(binary)

    case :unicode.characters_to_binary(stripped, {:utf16, :little}, :utf8) do
      utf8 when is_binary(utf8) ->
        {:ok, utf8}

      {:error, _, _} ->
        {:error, "invalid UTF-16LE data"}

      {:incomplete, _, _} ->
        {:error, "incomplete UTF-16LE data"}
    end
  end

  def parse_lpwstr(_), do: {:error, "expected a binary"}

  # -------------------------------------------------------------------
  # FILETIME parsing
  # -------------------------------------------------------------------

  @doc """
  Parses an 8-byte Windows FILETIME to an Elixir `DateTime`.

  FILETIME is a little-endian 64-bit unsigned integer representing the number
  of 100-nanosecond intervals since January 1, 1601, 00:00:00 UTC.

  A FILETIME value of 0 is treated as unset and returns `{:error, "zero filetime"}`.

  Returns `{:ok, datetime}` or `{:error, reason}`.

  ## Examples

      iex> # 2023-01-01 00:00:00 UTC
      iex> OutlookMsg.Ole.Types.parse_filetime(<<0x00, 0x00, 0xC3, 0xFD, 0x73, 0x1D, 0xD9, 0x01>>)
      {:ok, ~U[2023-01-01 00:00:00Z]}
  """
  @spec parse_filetime(binary()) :: {:ok, DateTime.t()} | {:error, String.t()}
  def parse_filetime(<<0::little-unsigned-integer-size(64)>>) do
    {:error, "zero filetime"}
  end

  def parse_filetime(<<ticks::little-unsigned-integer-size(64)>>) do
    unix_seconds = div(ticks, @filetime_ticks_per_second) - @filetime_epoch_offset
    remainder_ticks = rem(ticks, @filetime_ticks_per_second)
    microseconds = div(remainder_ticks, 10)

    case DateTime.from_unix(unix_seconds, :second) do
      {:ok, dt} ->
        # Add sub-second precision if present
        if microseconds > 0 do
          {:ok, DateTime.add(dt, microseconds, :microsecond)}
        else
          {:ok, dt}
        end

      {:error, reason} ->
        {:error, "failed to convert filetime to DateTime: #{inspect(reason)}"}
    end
  end

  def parse_filetime(<<_::binary-size(8), _rest::binary>>) do
    {:error, "extra bytes after 8-byte FILETIME"}
  end

  def parse_filetime(_) do
    {:error, "expected exactly 8 bytes for FILETIME"}
  end

  # -------------------------------------------------------------------
  # Private helpers
  # -------------------------------------------------------------------

  # Converts an integer to a zero-padded uppercase hex string of the given width.
  @spec hex_pad(non_neg_integer(), pos_integer()) :: String.t()
  defp hex_pad(value, width) do
    value
    |> Integer.to_string(16)
    |> String.upcase()
    |> String.pad_leading(width, "0")
  end

  # Parses a hex string into an integer.
  @spec hex_to_int(String.t()) :: {:ok, non_neg_integer()} | :error
  defp hex_to_int(hex_string) do
    case Integer.parse(hex_string, 16) do
      {value, ""} -> {:ok, value}
      _ -> :error
    end
  end

  # Strips trailing UTF-16LE null characters (0x00 0x00 pairs) from a binary.
  @spec strip_utf16le_nulls(binary()) :: binary()
  defp strip_utf16le_nulls(binary) do
    byte_size = byte_size(binary)

    # Ensure we have an even number of bytes for valid UTF-16LE
    aligned_size =
      if rem(byte_size, 2) == 1 do
        byte_size - 1
      else
        byte_size
      end

    strip_trailing_utf16le_nulls(binary, aligned_size)
  end

  defp strip_trailing_utf16le_nulls(binary, size) when size >= 2 do
    # Check the last two bytes
    offset = size - 2

    case binary do
      <<_::binary-size(offset), 0x00, 0x00, _::binary>> ->
        strip_trailing_utf16le_nulls(binary, offset)

      _ ->
        binary_part(binary, 0, size)
    end
  end

  defp strip_trailing_utf16le_nulls(_binary, _size), do: <<>>
end
