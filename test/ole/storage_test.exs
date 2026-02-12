defmodule OutlookMsg.Ole.StorageTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Ole.{Header, Fat, Dirent, Types}

  describe "Header.parse/1" do
    test "parses valid OLE header" do
      # Build a minimal valid OLE header (512 bytes)
      magic = <<0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1>>
      clsid = <<0::128>>
      minor_version = <<62, 0>>  # 0x003E
      major_version = <<3, 0>>   # v3
      byte_order = <<0xFE, 0xFF>>
      sector_shift = <<9, 0>>    # 2^9 = 512
      mini_shift = <<6, 0>>      # 2^6 = 64
      reserved = <<0::48>>       # 6 bytes
      dir_sector_count = <<0::32>>
      fat_count = <<1, 0, 0, 0>>
      dir_start = <<0, 0, 0, 0>>
      txn_sig = <<0::32>>
      mini_cutoff = <<0, 16, 0, 0>>  # 4096
      mini_fat_start = <<0xFE, 0xFF, 0xFF, 0xFF>>
      mini_fat_count = <<0::32>>
      difat_start = <<0xFE, 0xFF, 0xFF, 0xFF>>
      difat_count = <<0::32>>
      # DIFAT array: 109 entries, first one is sector 0, rest are 0xFFFFFFFF
      difat = <<0, 0, 0, 0>> <> :binary.copy(<<0xFF, 0xFF, 0xFF, 0xFF>>, 108)

      header_data = magic <> clsid <> minor_version <> major_version <>
        byte_order <> sector_shift <> mini_shift <> reserved <>
        dir_sector_count <> fat_count <> dir_start <> txn_sig <>
        mini_cutoff <> mini_fat_start <> mini_fat_count <>
        difat_start <> difat_count <> difat

      assert byte_size(header_data) == 512

      assert {:ok, header} = Header.parse(header_data)
      assert header.sector_size == 512
      assert header.mini_sector_size == 64
      assert header.fat_sector_count == 1
      assert header.dir_start_sector == 0
      assert header.mini_cutoff == 4096
    end

    test "rejects invalid magic" do
      bad_header = <<0, 0, 0, 0, 0, 0, 0, 0>> <> :binary.copy(<<0>>, 504)
      assert {:error, _} = Header.parse(bad_header)
    end

    test "rejects too-short data" do
      assert {:error, _} = Header.parse(<<1, 2, 3>>)
    end
  end

  describe "Types.parse_guid/1" do
    test "parses 16-byte GUID" do
      # PS_MAPI GUID: {00020328-0000-0000-C000-000000000046}
      # Mixed-endian binary
      guid_binary = <<0x28, 0x03, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
                      0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46>>
      assert {:ok, result} = Types.parse_guid(guid_binary)
      assert is_binary(result)
      assert String.contains?(result, "00020328")
    end

    test "rejects non-16-byte input" do
      assert {:error, _} = Types.parse_guid(<<1, 2, 3>>)
    end
  end

  describe "Types.parse_lpwstr/1" do
    test "converts UTF-16LE to UTF-8" do
      # "Hello" in UTF-16LE with null terminator
      utf16 = <<0x48, 0x00, 0x65, 0x00, 0x6C, 0x00, 0x6C, 0x00, 0x6F, 0x00, 0x00, 0x00>>
      assert {:ok, "Hello"} = Types.parse_lpwstr(utf16)
    end

    test "handles empty string" do
      assert {:ok, ""} = Types.parse_lpwstr(<<0x00, 0x00>>)
      assert {:ok, ""} = Types.parse_lpwstr(<<>>)
    end
  end

  describe "Types.parse_filetime/1" do
    test "converts Windows FILETIME to DateTime" do
      # Known FILETIME value: 2024-01-01 00:00:00 UTC
      # FILETIME = (unix_seconds + 11644473600) * 10_000_000
      ticks = (1_704_067_200 + 11_644_473_600) * 10_000_000
      filetime_binary = <<ticks::64-little>>
      assert {:ok, result} = Types.parse_filetime(filetime_binary)
      assert %DateTime{} = result
      assert result.year == 2024
      assert result.month == 1
      assert result.day == 1
    end

    test "handles zero filetime" do
      assert {:error, "zero filetime"} = Types.parse_filetime(<<0::64>>)
    end
  end

  describe "Types.encode_guid/1" do
    test "round-trips a GUID string" do
      guid_str = "{00020328-0000-0000-C000-000000000046}"
      assert {:ok, binary} = Types.encode_guid(guid_str)
      assert byte_size(binary) == 16
      assert {:ok, ^guid_str} = Types.parse_guid(binary)
    end
  end

  describe "Types.format_guid/1" do
    test "formats a 16-byte binary" do
      guid_binary = <<0x28, 0x03, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
                      0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46>>
      result = Types.format_guid(guid_binary)
      assert String.starts_with?(result, "{")
      assert String.ends_with?(result, "}")
      assert String.contains?(result, "00020328")
    end

    test "passes through already-formatted GUID string" do
      guid_str = "{00020328-0000-0000-C000-000000000046}"
      assert Types.format_guid(guid_str) == guid_str
    end
  end

  describe "Fat.chain/2" do
    test "follows sector chain" do
      # FAT: sector 0 -> 1 -> 2 -> ENDOFCHAIN
      fat = %{0 => 1, 1 => 2, 2 => 0xFFFFFFFE}
      assert Fat.chain(fat, 0) == [0, 1, 2]
    end

    test "handles single sector" do
      fat = %{5 => 0xFFFFFFFE}
      assert Fat.chain(fat, 5) == [5]
    end

    test "handles ENDOFCHAIN start" do
      fat = %{}
      assert Fat.chain(fat, 0xFFFFFFFE) == []
    end

    test "handles FREESECT start" do
      fat = %{}
      assert Fat.chain(fat, 0xFFFFFFFF) == []
    end

    test "detects cycles" do
      # Circular chain: 0 -> 1 -> 0 (cycle)
      fat = %{0 => 1, 1 => 0}
      result = Fat.chain(fat, 0)
      # Should terminate despite cycle
      assert is_list(result)
      assert length(result) <= 2
    end
  end

  describe "Fat.sector_offset/2" do
    test "calculates correct offset" do
      # Sector 0 starts after the 512-byte header
      assert Fat.sector_offset(0, 512) == 512
      assert Fat.sector_offset(1, 512) == 1024
      assert Fat.sector_offset(2, 512) == 1536
    end
  end

  describe "Dirent.parse/2" do
    test "parses a 128-byte directory entry" do
      # Build a root entry
      name = "Root Entry"
      name_utf16 = :unicode.characters_to_binary(name, :utf8, {:utf16, :little})
      name_padded = name_utf16 <> :binary.copy(<<0>>, 64 - byte_size(name_utf16))
      name_size = <<byte_size(name_utf16) + 2, 0>>  # includes null terminator
      type = <<5>>  # root
      color = <<1>>  # black
      left_sid = <<0xFF, 0xFF, 0xFF, 0xFF>>
      right_sid = <<0xFF, 0xFF, 0xFF, 0xFF>>
      child_sid = <<1, 0, 0, 0>>
      clsid = <<0::128>>
      state_bits = <<0::32>>
      create_time = <<0::64>>
      modify_time = <<0::64>>
      start_sector = <<0, 0, 0, 0>>
      size = <<0, 0, 0, 0, 0, 0, 0, 0>>

      entry = name_padded <> name_size <> type <> color <>
        left_sid <> right_sid <> child_sid <>
        clsid <> state_bits <> create_time <> modify_time <>
        start_sector <> size

      assert byte_size(entry) == 128
      dirent = Dirent.parse(entry, 0)
      assert dirent.name == "Root Entry"
      assert dirent.type == :root
      assert dirent.child_sid == 1
      assert dirent.sid == 0
    end

    test "parses a stream entry" do
      name = "TestStream"
      name_utf16 = :unicode.characters_to_binary(name, :utf8, {:utf16, :little})
      name_padded = name_utf16 <> :binary.copy(<<0>>, 64 - byte_size(name_utf16))
      name_size = <<byte_size(name_utf16) + 2, 0>>
      type = <<2>>  # stream
      color = <<0>>  # red
      left_sid = <<0xFF, 0xFF, 0xFF, 0xFF>>
      right_sid = <<0xFF, 0xFF, 0xFF, 0xFF>>
      child_sid = <<0xFF, 0xFF, 0xFF, 0xFF>>
      clsid = <<0::128>>
      state_bits = <<0::32>>
      create_time = <<0::64>>
      modify_time = <<0::64>>
      start_sector = <<5, 0, 0, 0>>
      size = <<100, 0, 0, 0, 0, 0, 0, 0>>

      entry = name_padded <> name_size <> type <> color <>
        left_sid <> right_sid <> child_sid <>
        clsid <> state_bits <> create_time <> modify_time <>
        start_sector <> size

      assert byte_size(entry) == 128
      dirent = Dirent.parse(entry, 3)
      assert dirent.name == "TestStream"
      assert dirent.type == :stream
      assert dirent.color == :red
      assert dirent.sid == 3
      assert dirent.start_sector == 5
      assert dirent.size == 100
    end
  end

  describe "Dirent.parse_all/1" do
    test "parses multiple entries and skips empty ones" do
      root = build_dirent("Root Entry", 5, 0xFF_FF_FF_FF, 0xFF_FF_FF_FF, 0xFF_FF_FF_FF)
      empty = :binary.copy(<<0>>, 128)  # type 0 = empty

      entries = Dirent.parse_all(root <> empty)
      assert length(entries) == 1
      assert hd(entries).name == "Root Entry"
    end
  end

  describe "Dirent.find_child/2" do
    test "finds child by name (case-insensitive)" do
      child = %Dirent{name: "TestChild", type: :stream, children: [],
                      color: :black, left_sid: 0xFFFFFFFF, right_sid: 0xFFFFFFFF,
                      child_sid: 0xFFFFFFFF, clsid: <<0::128>>, state_bits: 0,
                      create_time: <<0::64>>, modify_time: <<0::64>>,
                      start_sector: 0, size: 0, sid: 1}
      parent = %Dirent{name: "Root", type: :root, children: [child],
                       color: :black, left_sid: 0xFFFFFFFF, right_sid: 0xFFFFFFFF,
                       child_sid: 1, clsid: <<0::128>>, state_bits: 0,
                       create_time: <<0::64>>, modify_time: <<0::64>>,
                       start_sector: 0, size: 0, sid: 0}

      assert Dirent.find_child(parent, "testchild") == child
      assert Dirent.find_child(parent, "TESTCHILD") == child
      assert Dirent.find_child(parent, "missing") == nil
    end
  end

  # Helper to build a 128-byte directory entry binary
  defp build_dirent(name, type, left_sid, right_sid, child_sid) do
    name_utf16 = :unicode.characters_to_binary(name, :utf8, {:utf16, :little})
    name_padded = name_utf16 <> :binary.copy(<<0>>, 64 - byte_size(name_utf16))
    name_size = <<byte_size(name_utf16) + 2, 0>>
    type_byte = <<type>>
    color = <<1>>
    left = <<left_sid::32-little>>
    right = <<right_sid::32-little>>
    child = <<child_sid::32-little>>
    clsid = <<0::128>>
    state_bits = <<0::32>>
    create_time = <<0::64>>
    modify_time = <<0::64>>
    start_sector = <<0, 0, 0, 0>>
    size = <<0, 0, 0, 0, 0, 0, 0, 0>>

    name_padded <> name_size <> type_byte <> color <>
      left <> right <> child <>
      clsid <> state_bits <> create_time <> modify_time <>
      start_sector <> size
  end
end
