defmodule OutlookMsg.Mapi.Key do
  @moduledoc """
  A MAPI property key combining a property code with a GUID.

  A `Key` uniquely identifies a MAPI property by pairing an integer property
  code (e.g., `0x0037` for `PR_SUBJECT`) with a 16-byte property set GUID.
  When no GUID is specified, `PS_MAPI` is used as the default, which covers
  the standard MAPI property range.

  ## Examples

      iex> key = OutlookMsg.Mapi.Key.new(0x0037)
      iex> key.code
      0x0037

      iex> key = OutlookMsg.Mapi.Key.new(0x0037)
      iex> key.guid == OutlookMsg.Mapi.Guids.ps_mapi()
      true

      iex> key = OutlookMsg.Mapi.Key.new(0x0037, <<0::128>>)
      iex> key.guid
      <<0::128>>
  """

  alias OutlookMsg.Mapi.Guids

  defstruct [:code, :guid]

  @type t :: %__MODULE__{
          code: non_neg_integer() | String.t(),
          guid: <<_::128>>
        }

  # -------------------------------------------------------------------
  # Constructors
  # -------------------------------------------------------------------

  @doc """
  Creates a new key from a property code, using `PS_MAPI` as the default GUID.

  ## Examples

      iex> key = OutlookMsg.Mapi.Key.new(0x0037)
      iex> key.code
      0x0037
  """
  @spec new(non_neg_integer()) :: t()
  def new(code) when is_integer(code) do
    %__MODULE__{code: code, guid: Guids.ps_mapi()}
  end

  @doc """
  Creates a new key from a property code and a 16-byte binary GUID.

  ## Examples

      iex> guid = OutlookMsg.Mapi.Guids.psetid_common()
      iex> key = OutlookMsg.Mapi.Key.new(0x8501, guid)
      iex> key.code
      0x8501
  """
  @spec new(non_neg_integer() | String.t(), <<_::128>>) :: t()
  def new(code, guid) when is_integer(code) and is_binary(guid) do
    %__MODULE__{code: code, guid: guid}
  end

  def new(code, guid) when is_binary(code) and is_binary(guid) do
    %__MODULE__{code: code, guid: guid}
  end

  # -------------------------------------------------------------------
  # Symbolic resolution
  # -------------------------------------------------------------------

  @doc """
  Resolves the key to a symbolic atom name, or `nil` if no name is known.

  For keys in the `PS_MAPI` property set, the name is looked up via
  `OutlookMsg.Mapi.Tags.name/1`. For keys in other property sets, the name
  is looked up via `OutlookMsg.Mapi.NamedMap.lookup/2`.

  ## Examples

      iex> key = OutlookMsg.Mapi.Key.new(0x0037)
      iex> OutlookMsg.Mapi.Key.to_sym(key)  # returns e.g. :pr_subject or nil
  """
  @spec to_sym(t()) :: atom() | nil
  def to_sym(%__MODULE__{code: code, guid: guid}) do
    if is_integer(code) do
      if guid == Guids.ps_mapi() do
        case OutlookMsg.Mapi.Tags.name(code) do
          nil -> nil
          :unknown -> nil
          name -> name
        end
      else
        case OutlookMsg.Mapi.NamedMap.lookup(code, guid) do
          nil -> nil
          :unknown -> nil
          name -> name
        end
      end
    else
      nil
    end
  end

  # -------------------------------------------------------------------
  # String conversion
  # -------------------------------------------------------------------

  @doc """
  Returns a human-readable string representation of the key.

  If the key resolves to a symbolic name, that name is returned as a string.
  Otherwise, the property code is formatted as a zero-padded 4-digit
  hexadecimal string. When the GUID is not `PS_MAPI`, the GUID name (or
  raw GUID info) is appended.

  ## Examples

      iex> key = OutlookMsg.Mapi.Key.new(0x0037)
      iex> OutlookMsg.Mapi.Key.to_string(key)
      # "pr_subject" if known, or "0037" if not
  """
  @spec to_string(t()) :: String.t()
  def to_string(%__MODULE__{} = key) do
    case to_sym(key) do
      nil ->
        if is_integer(key.code) do
          hex = key.code |> Integer.to_string(16) |> String.pad_leading(4, "0") |> String.downcase()

          if key.guid == Guids.ps_mapi() do
            hex
          else
            guid_label = format_guid(key.guid)
            "#{hex}@#{guid_label}"
          end
        else
          "#{inspect(key.code)}@#{format_guid(key.guid)}"
        end

      name ->
        Atom.to_string(name)
    end
  end

  @doc """
  Returns `true` if the key resolves to a symbolic name, `false` otherwise.

  ## Examples

      iex> key = OutlookMsg.Mapi.Key.new(0x0037)
      iex> OutlookMsg.Mapi.Key.symbolic?(key)  # depends on Tags module
  """
  @spec symbolic?(t()) :: boolean()
  def symbolic?(%__MODULE__{} = key) do
    to_sym(key) != nil
  end

  # -------------------------------------------------------------------
  # Private helpers
  # -------------------------------------------------------------------

  # Formats a GUID for display. If the GUID is a well-known property set,
  # returns the property set name; otherwise returns a hex representation.
  @spec format_guid(binary()) :: String.t()
  defp format_guid(guid) do
    case Guids.name(guid) do
      :unknown ->
        Base.encode16(guid, case: :lower)

      name ->
        Atom.to_string(name)
    end
  end
end

# -------------------------------------------------------------------
# Protocol implementations
# -------------------------------------------------------------------

defimpl Inspect, for: OutlookMsg.Mapi.Key do
  import Inspect.Algebra

  def inspect(%OutlookMsg.Mapi.Key{} = key, _opts) do
    alias OutlookMsg.Mapi.Key
    alias OutlookMsg.Mapi.Guids

    hex_code =
      if is_integer(key.code) do
        key.code
        |> Integer.to_string(16)
        |> String.pad_leading(4, "0")
        |> String.downcase()
      else
        inspect(key.code)
      end

    sym =
      try do
        Key.to_sym(key)
      rescue
        _e -> nil
      end

    label =
      case sym do
        nil ->
          if is_integer(key.code), do: "0x#{hex_code}", else: "key=#{hex_code}"
        name -> "0x#{hex_code} #{name}"
      end

    guid_part =
      if key.guid == Guids.ps_mapi() do
        ""
      else
        case Guids.name(key.guid) do
          :unknown -> ", guid: #{Base.encode16(key.guid, case: :lower)}"
          name -> ", guid: #{name}"
        end
      end

    concat(["#Mapi.Key<#{label}#{guid_part}>"])
  end
end

defimpl String.Chars, for: OutlookMsg.Mapi.Key do
  def to_string(%OutlookMsg.Mapi.Key{} = key) do
    OutlookMsg.Mapi.Key.to_string(key)
  end
end
