defmodule OutlookMsg.Mapi.Types do
  @moduledoc """
  MAPI property type constants and encoding/decoding functions.

  Defines the standard MAPI property type codes used in Outlook MSG files
  and provides functions to decode raw binary property values into Elixir terms.
  """

  alias OutlookMsg.Ole.Types, as: OleTypes

  # -------------------------------------------------------------------
  # MAPI property type constants
  # -------------------------------------------------------------------

  @pt_unspecified 0x0000
  @pt_null        0x0001
  @pt_short       0x0002
  @pt_long        0x0003
  @pt_float       0x0004
  @pt_double      0x0005
  @pt_currency    0x0006
  @pt_apptime     0x0007
  @pt_error       0x000A
  @pt_boolean     0x000B
  @pt_object      0x000D
  @pt_int64       0x0014
  @pt_string8     0x001E
  @pt_unicode     0x001F
  @pt_systime     0x0040
  @pt_clsid       0x0048
  @pt_binary      0x0102

  @mv_flag        0x1000

  # -------------------------------------------------------------------
  # Public constant accessor functions
  # -------------------------------------------------------------------

  @doc "Returns the PT_UNSPECIFIED type code (0x0000)."
  @spec pt_unspecified() :: non_neg_integer()
  def pt_unspecified, do: @pt_unspecified

  @doc "Returns the PT_NULL type code (0x0001)."
  @spec pt_null() :: non_neg_integer()
  def pt_null, do: @pt_null

  @doc "Returns the PT_SHORT type code (0x0002) for 16-bit integers."
  @spec pt_short() :: non_neg_integer()
  def pt_short, do: @pt_short

  @doc "Returns the PT_LONG type code (0x0003) for 32-bit integers."
  @spec pt_long() :: non_neg_integer()
  def pt_long, do: @pt_long

  @doc "Returns the PT_FLOAT type code (0x0004) for 32-bit floats."
  @spec pt_float() :: non_neg_integer()
  def pt_float, do: @pt_float

  @doc "Returns the PT_DOUBLE type code (0x0005) for 64-bit floats."
  @spec pt_double() :: non_neg_integer()
  def pt_double, do: @pt_double

  @doc "Returns the PT_CURRENCY type code (0x0006) for 64-bit integer scaled by 10000."
  @spec pt_currency() :: non_neg_integer()
  def pt_currency, do: @pt_currency

  @doc "Returns the PT_APPTIME type code (0x0007) for application time (double)."
  @spec pt_apptime() :: non_neg_integer()
  def pt_apptime, do: @pt_apptime

  @doc "Returns the PT_ERROR type code (0x000A) for 32-bit error codes."
  @spec pt_error() :: non_neg_integer()
  def pt_error, do: @pt_error

  @doc "Returns the PT_BOOLEAN type code (0x000B) for 16-bit booleans."
  @spec pt_boolean() :: non_neg_integer()
  def pt_boolean, do: @pt_boolean

  @doc "Returns the PT_OBJECT type code (0x000D) for embedded objects."
  @spec pt_object() :: non_neg_integer()
  def pt_object, do: @pt_object

  @doc "Returns the PT_INT64 type code (0x0014) for 64-bit integers."
  @spec pt_int64() :: non_neg_integer()
  def pt_int64, do: @pt_int64

  @doc "Returns the PT_STRING8 type code (0x001E) for ANSI strings."
  @spec pt_string8() :: non_neg_integer()
  def pt_string8, do: @pt_string8

  @doc "Returns the PT_UNICODE type code (0x001F) for Unicode (UTF-16LE) strings."
  @spec pt_unicode() :: non_neg_integer()
  def pt_unicode, do: @pt_unicode

  @doc "Returns the PT_SYSTIME type code (0x0040) for FILETIME values."
  @spec pt_systime() :: non_neg_integer()
  def pt_systime, do: @pt_systime

  @doc "Returns the PT_CLSID type code (0x0048) for GUID values."
  @spec pt_clsid() :: non_neg_integer()
  def pt_clsid, do: @pt_clsid

  @doc "Returns the PT_BINARY type code (0x0102) for binary blobs."
  @spec pt_binary() :: non_neg_integer()
  def pt_binary, do: @pt_binary

  @doc "Returns the multi-value flag (0x1000)."
  @spec mv_flag() :: non_neg_integer()
  def mv_flag, do: @mv_flag

  # -------------------------------------------------------------------
  # Type code to name mapping
  # -------------------------------------------------------------------

  @type_names %{
    @pt_unspecified => :pt_unspecified,
    @pt_null        => :pt_null,
    @pt_short       => :pt_short,
    @pt_long        => :pt_long,
    @pt_float       => :pt_float,
    @pt_double      => :pt_double,
    @pt_currency    => :pt_currency,
    @pt_apptime     => :pt_apptime,
    @pt_error       => :pt_error,
    @pt_boolean     => :pt_boolean,
    @pt_object      => :pt_object,
    @pt_int64       => :pt_int64,
    @pt_string8     => :pt_string8,
    @pt_unicode     => :pt_unicode,
    @pt_systime     => :pt_systime,
    @pt_clsid       => :pt_clsid,
    @pt_binary      => :pt_binary
  }

  @doc """
  Converts a MAPI property type code to its atom name.

  For multi-valued types (those with the 0x1000 flag set), the name is
  prefixed with `pt_mv_`. For example, 0x101F becomes `:pt_mv_unicode`.

  Returns `:unknown` for unrecognized type codes.

  ## Examples

      iex> OutlookMsg.Mapi.Types.type_name(0x001F)
      :pt_unicode

      iex> OutlookMsg.Mapi.Types.type_name(0x0003)
      :pt_long

      iex> OutlookMsg.Mapi.Types.type_name(0x101F)
      :pt_mv_unicode
  """
  @spec type_name(non_neg_integer()) :: atom()
  def type_name(type_code) when is_integer(type_code) do
    if multi_value?(type_code) do
      base = base_type(type_code)

      case Map.get(@type_names, base) do
        nil -> :unknown
        name -> :"pt_mv_#{name |> Atom.to_string() |> String.trim_leading("pt_")}"
      end
    else
      Map.get(@type_names, type_code, :unknown)
    end
  end

  # -------------------------------------------------------------------
  # Multi-value helpers
  # -------------------------------------------------------------------

  @doc """
  Strips the multi-value flag from a type code, returning the base type.

  ## Examples

      iex> OutlookMsg.Mapi.Types.base_type(0x101F)
      0x001F

      iex> OutlookMsg.Mapi.Types.base_type(0x0003)
      0x0003
  """
  @spec base_type(non_neg_integer()) :: non_neg_integer()
  def base_type(type_code) when is_integer(type_code) do
    Bitwise.band(type_code, 0x0FFF)
  end

  @doc """
  Checks whether a type code has the multi-value flag (0x1000) set.

  ## Examples

      iex> OutlookMsg.Mapi.Types.multi_value?(0x101F)
      true

      iex> OutlookMsg.Mapi.Types.multi_value?(0x001F)
      false
  """
  @spec multi_value?(non_neg_integer()) :: boolean()
  def multi_value?(type_code) when is_integer(type_code) do
    Bitwise.band(type_code, @mv_flag) != 0
  end

  # -------------------------------------------------------------------
  # Value decoding
  # -------------------------------------------------------------------

  @doc """
  Decodes a raw binary property value according to its MAPI type code.

  Returns the decoded Elixir value appropriate for the type:

    - `PT_SHORT` (0x0002): 16-bit little-endian integer
    - `PT_LONG` (0x0003): 32-bit little-endian signed integer
    - `PT_FLOAT` (0x0004): 32-bit little-endian IEEE 754 float
    - `PT_DOUBLE` (0x0005): 64-bit little-endian IEEE 754 float
    - `PT_CURRENCY` (0x0006): 64-bit signed integer divided by 10,000
    - `PT_APPTIME` (0x0007): 64-bit little-endian float (application time)
    - `PT_ERROR` (0x000A): 32-bit little-endian unsigned integer (error code)
    - `PT_BOOLEAN` (0x000B): 16-bit value, `true` if non-zero
    - `PT_INT64` (0x0014): 64-bit little-endian signed integer
    - `PT_STRING8` (0x001E): ANSI string with trailing nulls stripped
    - `PT_UNICODE` (0x001F): UTF-16LE string decoded to UTF-8, trailing nulls stripped
    - `PT_SYSTIME` (0x0040): 8-byte FILETIME decoded to `DateTime`
    - `PT_CLSID` (0x0048): 16-byte GUID decoded to string representation
    - `PT_BINARY` (0x0102): returned as-is

  For unrecognized type codes, returns the binary as-is.

  ## Examples

      iex> OutlookMsg.Mapi.Types.decode_value(0x0003, <<42, 0, 0, 0>>)
      42

      iex> OutlookMsg.Mapi.Types.decode_value(0x000B, <<1, 0, 0, 0>>)
      true
  """
  @spec decode_value(non_neg_integer(), binary()) :: term()
  def decode_value(type_code, data) when is_integer(type_code) and is_binary(data) do
    do_decode(base_type(type_code), data)
  end

  # PT_SHORT: 16-bit integer
  defp do_decode(@pt_short, <<val::16-little>>) do
    val
  end

  # PT_LONG: 32-bit signed integer
  defp do_decode(@pt_long, <<val::32-little-signed>>) do
    val
  end

  # PT_FLOAT: 32-bit float
  defp do_decode(@pt_float, <<val::32-little-float>>) do
    val
  end

  # PT_DOUBLE: 64-bit float
  defp do_decode(@pt_double, <<val::64-little-float>>) do
    val
  end

  # PT_CURRENCY: 64-bit signed integer, scaled by 10000
  defp do_decode(@pt_currency, <<val::64-little-signed>>) do
    val / 10_000.0
  end

  # PT_APPTIME: 64-bit float (application time)
  defp do_decode(@pt_apptime, <<val::64-little-float>>) do
    val
  end

  # PT_ERROR: 32-bit error code
  defp do_decode(@pt_error, <<val::32-little>>) do
    val
  end

  # PT_BOOLEAN: 16-bit boolean (may have trailing padding bytes)
  defp do_decode(@pt_boolean, <<val::16-little, _::binary>>) do
    val != 0
  end

  # PT_INT64: 64-bit signed integer
  defp do_decode(@pt_int64, <<val::64-little-signed>>) do
    val
  end

  # PT_SYSTIME: 8-byte FILETIME -> DateTime
  defp do_decode(@pt_systime, <<_::binary-size(8)>> = data) do
    case OleTypes.parse_filetime(data) do
      {:ok, datetime} -> datetime
      {:error, _reason} -> data
    end
  end

  # PT_UNICODE: UTF-16LE string -> UTF-8 string
  defp do_decode(@pt_unicode, data) do
    case OleTypes.parse_lpwstr(data) do
      {:ok, string} -> string
      {:error, _reason} -> data
    end
  end

  # PT_STRING8: ANSI string, strip trailing nulls
  defp do_decode(@pt_string8, data) do
    data
    |> strip_trailing_nulls()
  end

  # PT_BINARY: return as-is
  defp do_decode(@pt_binary, data) do
    data
  end

  # PT_CLSID: 16-byte GUID -> string
  defp do_decode(@pt_clsid, <<_::binary-size(16)>> = data) do
    case OleTypes.parse_guid(data) do
      {:ok, guid_string} -> guid_string
      {:error, _reason} -> data
    end
  end

  # Unknown or unhandled types: return binary as-is
  defp do_decode(_type, data) do
    data
  end

  # -------------------------------------------------------------------
  # Private helpers
  # -------------------------------------------------------------------

  # Strips trailing null bytes (0x00) from a binary string.
  @spec strip_trailing_nulls(binary()) :: binary()
  defp strip_trailing_nulls(binary) do
    binary
    |> :binary.bin_to_list()
    |> Enum.reverse()
    |> Enum.drop_while(&(&1 == 0))
    |> Enum.reverse()
    |> :binary.list_to_bin()
  end
end
