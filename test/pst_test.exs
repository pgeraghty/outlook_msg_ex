defmodule OutlookMsg.PstTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Pst.{Header, Encryption}

  describe "Header.parse/1" do
    test "rejects non-PST data" do
      assert {:error, :invalid_pst_magic} = Header.parse(:binary.copy(<<0>>, 512))
    end

    test "rejects short data" do
      assert {:error, :data_too_short} = Header.parse(<<1, 2, 3>>)
    end

    test "parses valid PST 97 header" do
      # Build a minimal PST 97 header (at least 512 bytes)
      # Magic: "!BDN" at offset 0
      magic = <<0x21, 0x42, 0x44, 0x4E>>
      # Bytes 4-9: padding (6 bytes)
      padding1 = <<0, 0, 0, 0, 0, 0>>
      # Byte 10: index_type 0x0E for PST 97
      index_type = <<0x0E>>
      # Remaining bytes up to offset 0xA0 (160)
      padding2 = :binary.copy(<<0>>, 160 - 11)
      # Offset 0xA0: index1 (4 bytes LE)
      index1 = <<0x10, 0x00, 0x00, 0x00>>  # 16
      # Offset 0xA4: 4 bytes padding
      padding3 = <<0, 0, 0, 0>>
      # Offset 0xA8: index2 (4 bytes LE)
      index2 = <<0x20, 0x00, 0x00, 0x00>>  # 32
      # Pad up to offset 0x1CD (461)
      padding4 = :binary.copy(<<0>>, 461 - 172)
      # Offset 0x1CD: encryption_type
      encryption = <<1>>  # compressible
      # Pad to at least 512 bytes
      remaining = :binary.copy(<<0>>, 512 - 462)

      data = magic <> padding1 <> index_type <> padding2 <>
        index1 <> padding3 <> index2 <> padding4 <>
        encryption <> remaining

      assert byte_size(data) >= 512

      assert {:ok, header} = Header.parse(data)
      assert header.version == :pst97
      assert header.encryption_type == 1
      assert header.index1_offset == 16
      assert header.index2_offset == 32
    end

    test "parses valid PST 2003 header" do
      # Build a minimal PST 2003 header (at least 514 bytes for offset 0x0201)
      magic = <<0x21, 0x42, 0x44, 0x4E>>
      padding1 = <<0, 0, 0, 0, 0, 0>>
      index_type = <<0x17>>  # PST 2003
      # Pad up to offset 0xB8 (184)
      padding2 = :binary.copy(<<0>>, 184 - 11)
      # Offset 0xB8: index1 (8 bytes LE)
      index1 = <<0x30, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00>>  # 48
      # Offset 0xC0: index2 (8 bytes LE)
      index2 = <<0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00>>  # 64
      # Pad up to offset 0x0201 (513)
      padding3 = :binary.copy(<<0>>, 513 - 200)
      # Offset 0x0201: encryption_type
      encryption = <<0>>  # none
      # Pad to reasonable size
      remaining = :binary.copy(<<0>>, 100)

      data = magic <> padding1 <> index_type <> padding2 <>
        index1 <> index2 <> padding3 <>
        encryption <> remaining

      assert byte_size(data) >= 512

      assert {:ok, header} = Header.parse(data)
      assert header.version == :pst2003
      assert header.encryption_type == 0
      assert header.index1_offset == 48
      assert header.index2_offset == 64
    end
  end

  describe "Encryption" do
    test "decrypt/1 applies substitution" do
      # The decryption table maps bytes
      data = <<0, 1, 2, 3>>
      result = Encryption.decrypt(data)
      assert byte_size(result) == 4
      assert result != data  # Should be different after substitution
    end

    test "decrypt/1 produces correct substitution for known bytes" do
      # From the decrypt table: byte 0x00 => 0x47
      data = <<0>>
      result = Encryption.decrypt(data)
      assert result == <<0x47>>
    end

    test "decrypt/1 handles empty binary" do
      assert Encryption.decrypt(<<>>) == <<>>
    end

    test "maybe_decrypt/2 with type 0 returns unchanged" do
      data = <<1, 2, 3, 4>>
      assert Encryption.maybe_decrypt(data, 0) == data
    end

    test "maybe_decrypt/2 with type 1 decrypts" do
      data = <<0, 1, 2>>
      result = Encryption.maybe_decrypt(data, 1)
      assert result == Encryption.decrypt(data)
    end

    test "maybe_decrypt/2 with unknown type returns unchanged" do
      data = <<5, 6, 7>>
      assert Encryption.maybe_decrypt(data, 99) == data
    end

    test "decrypt/1 preserves length" do
      data = :binary.copy(<<0xAB>>, 256)
      result = Encryption.decrypt(data)
      assert byte_size(result) == 256
    end
  end
end
