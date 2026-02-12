defmodule OutlookMsg.Mapi.Guids do
  @moduledoc """
  Well-known MAPI property set GUIDs as 16-byte mixed-endian binaries.

  These GUIDs come from the Microsoft MAPI documentation and the ruby-msg
  `property_set.rb` reference implementation. Each GUID is stored as a
  16-byte binary in mixed-endian format (first three components little-endian,
  last two components big-endian), which matches the on-disk representation
  in MSG files.

  ## Usage

      iex> OutlookMsg.Mapi.Guids.ps_mapi()
      <<0x28, 0x03, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
        0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46>>

      iex> OutlookMsg.Mapi.Guids.name(OutlookMsg.Mapi.Guids.ps_mapi())
      :ps_mapi

      iex> OutlookMsg.Mapi.Guids.all() |> map_size()
      16
  """

  # -------------------------------------------------------------------
  # Compile-time GUID encoding helper
  # -------------------------------------------------------------------

  # Converts a GUID string like "00020328-0000-0000-C000-000000000046"
  # into a 16-byte mixed-endian binary.
  #
  # Layout:
  #   - data1 (4 bytes): little-endian
  #   - data2 (2 bytes): little-endian
  #   - data3 (2 bytes): little-endian
  #   - data4 (2 bytes): big-endian
  #   - data5 (6 bytes): big-endian
  @doc false
  @spec encode_guid(String.t()) :: <<_::128>>
  def encode_guid(guid_string) do
    [d1, d2, d3, d4, d5] = String.split(guid_string, "-")

    {data1, ""} = Integer.parse(d1, 16)
    {data2, ""} = Integer.parse(d2, 16)
    {data3, ""} = Integer.parse(d3, 16)
    {data4, ""} = Integer.parse(d4, 16)
    {data5, ""} = Integer.parse(d5, 16)

    <<
      data1::little-unsigned-integer-size(32),
      data2::little-unsigned-integer-size(16),
      data3::little-unsigned-integer-size(16),
      data4::big-unsigned-integer-size(16),
      data5::big-unsigned-integer-size(48)
    >>
  end

  # -------------------------------------------------------------------
  # GUID definitions — {atom_name, guid_string} tuples
  # -------------------------------------------------------------------

  @guid_definitions [
    # Standard MAPI property sets
    {:ps_mapi, "00020328-0000-0000-C000-000000000046"},
    {:ps_public_strings, "00020329-0000-0000-C000-000000000046"},

    # Named property set GUIDs
    {:psetid_common, "00062008-0000-0000-C000-000000000046"},
    {:psetid_address, "00062004-0000-0000-C000-000000000046"},
    {:psetid_appointment, "00062002-0000-0000-C000-000000000046"},
    {:psetid_meeting, "6ED8DA90-450B-101B-98DA-00AA003F1305"},
    {:psetid_log, "0006200A-0000-0000-C000-000000000046"},
    {:psetid_messaging, "41F28F13-83F4-4114-A584-EEDB5A6B0BFF"},
    {:psetid_note, "0006200E-0000-0000-C000-000000000046"},
    {:psetid_post_rss, "00062041-0000-0000-C000-000000000046"},
    {:psetid_task, "00062003-0000-0000-C000-000000000046"},
    {:psetid_unified_messaging, "4442858E-A9E3-4E80-B900-317A210CC15B"},
    {:psetid_air_sync, "71035549-0739-4DCB-9163-00F0580DBBDF"},
    {:psetid_sharing, "00062040-0000-0000-C000-000000000046"},
    {:psetid_attachment, "96357F7F-59E1-47D0-99A7-46515C183B54"},

    # Internet headers
    {:ps_internet_headers, "00020386-0000-0000-C000-000000000046"}
  ]

  # -------------------------------------------------------------------
  # Compile-time: encode each GUID string to a 16-byte binary
  # -------------------------------------------------------------------

  @encoded_guids (for {name, guid_string} <- @guid_definitions do
    [d1, d2, d3, d4, d5] = String.split(guid_string, "-")
    {data1, ""} = Integer.parse(d1, 16)
    {data2, ""} = Integer.parse(d2, 16)
    {data3, ""} = Integer.parse(d3, 16)
    {data4, ""} = Integer.parse(d4, 16)
    {data5, ""} = Integer.parse(d5, 16)

    binary =
      <<
        data1::little-unsigned-integer-size(32),
        data2::little-unsigned-integer-size(16),
        data3::little-unsigned-integer-size(16),
        data4::big-unsigned-integer-size(16),
        data5::big-unsigned-integer-size(48)
      >>

    {name, guid_string, binary}
  end)

  # -------------------------------------------------------------------
  # Build the lookup map: binary GUID => atom name
  # -------------------------------------------------------------------

  @guid_map (for {name, _guid_string, binary} <- @encoded_guids, into: %{} do
    {binary, name}
  end)

  # -------------------------------------------------------------------
  # Generate public accessor functions (one per GUID)
  # -------------------------------------------------------------------

  for {name, guid_string, binary} <- @encoded_guids do
    @guid_binary binary

    @doc "Returns the 16-byte binary for `#{name |> Atom.to_string() |> String.upcase()}` `{#{guid_string}}`."
    @spec unquote(name)() :: <<_::128>>
    def unquote(name)(), do: @guid_binary
  end

  # -------------------------------------------------------------------
  # Public API — lookup helpers
  # -------------------------------------------------------------------

  @doc """
  Returns the atom name for a known 16-byte binary GUID.

  Returns `:unknown` if the GUID is not recognized.

  ## Examples

      iex> OutlookMsg.Mapi.Guids.name(OutlookMsg.Mapi.Guids.ps_mapi())
      :ps_mapi

      iex> OutlookMsg.Mapi.Guids.name(<<0::128>>)
      :unknown
  """
  @spec name(binary()) :: atom()
  def name(<<_::binary-size(16)>> = guid) do
    Map.get(@guid_map, guid, :unknown)
  end

  def name(_), do: :unknown

  @doc """
  Returns a map of all known GUIDs: `%{binary_guid => atom_name}`.

  ## Examples

      iex> map = OutlookMsg.Mapi.Guids.all()
      iex> Map.get(map, OutlookMsg.Mapi.Guids.ps_mapi())
      :ps_mapi
  """
  @spec all() :: %{binary() => atom()}
  def all, do: @guid_map
end
