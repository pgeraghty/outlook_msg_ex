defmodule OutlookMsg.Pst.Header do
  defstruct [
    :magic,           # 4 bytes, should be 0x2142444E ("!BDN")
    :index_type,      # 1 byte at offset 10: 0x0E = 97 format, 0x17 = 2003/64-bit format
    :encryption_type, # 1 byte at offset 513 (97) or 0x0201 (2003): 0=none, 1=compressible
    :index1_offset,   # Root of index B-tree (node index)
    :index2_offset,   # Root of descriptor B-tree
    :version,         # :pst97 or :pst2003
    :file_size
  ]

  @magic <<0x21, 0x42, 0x44, 0x4E>>  # "!BDN"

  def parse(data) when byte_size(data) >= 512 do
    <<magic::binary-size(4), _::binary-size(6), index_type::8, _rest::binary>> = data

    if magic != @magic do
      {:error, :invalid_pst_magic}
    else
      case index_type do
        0x0E -> parse_97(data)
        0x17 -> parse_2003(data)
        _ -> {:error, {:unknown_index_type, index_type}}
      end
    end
  end
  def parse(_), do: {:error, :data_too_short}

  # PST 97 (ANSI) format - 32-bit offsets
  defp parse_97(data) do
    # Encryption type at offset 0x1CD
    <<_::binary-size(0x1CD), encryption_type::8, _::binary>> = data
    # Index offsets at specific positions for 97 format
    # index1 backpointer at offset 0xA0 (4 bytes)
    # index2 backpointer at offset 0xA8 (4 bytes)
    <<_::binary-size(0xA0), index1::32-little, _::binary-size(4), index2::32-little, _::binary>> = data
    # File size at offset 0xA4... actually let's just use byte_size
    {:ok, %__MODULE__{
      magic: @magic,
      index_type: 0x0E,
      encryption_type: encryption_type,
      index1_offset: index1,
      index2_offset: index2,
      version: :pst97,
      file_size: byte_size(data)
    }}
  end

  # PST 2003 (Unicode) format - 64-bit offsets
  defp parse_2003(data) do
    if byte_size(data) < 0x0202 do
      {:error, :data_too_short}
    else
    # Encryption type at offset 0x0201
      <<_::binary-size(0x0201), encryption_type::8, _::binary>> = data
      # index1 at offset 0xB8 (8 bytes), index2 at offset 0xC0 (8 bytes)
      <<_::binary-size(0xB8), index1::64-little, index2::64-little, _::binary>> = data

      {:ok, %__MODULE__{
        magic: @magic,
        index_type: 0x17,
        encryption_type: encryption_type,
        index1_offset: index1,
        index2_offset: index2,
        version: :pst2003,
        file_size: byte_size(data)
      }}
    end
  end
end
