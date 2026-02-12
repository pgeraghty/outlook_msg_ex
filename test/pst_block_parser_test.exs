defmodule OutlookMsg.PstBlockParserTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Mapi.Key
  alias OutlookMsg.Pst.BlockParser

  describe "parse/5" do
    test "parses a minimal type1 property block with inline PT_LONG" do
      block =
        <<0xBC, 0x00, 13::16-little, 0x03, 0x00, 0x37, 0x00, 42::32-little-signed, 0x00, 0x00>>

      parsed = BlockParser.parse(block, %{}, <<>>, %{}, 0)

      assert parsed[Key.new(0x0037)] == 42
    end

    test "returns empty map for unknown signature" do
      assert BlockParser.parse(<<0x01, 0x00, 0x00, 0x00>>, %{}, <<>>, %{}, 0) == %{}
    end

    test "returns empty map for invalid type1 offset before property area" do
      block = <<0xBC, 0x00, 2::16-little, 0, 0, 0, 0>>
      assert BlockParser.parse(block, %{}, <<>>, %{}, 0) == %{}
    end

    test "returns empty map when type1 has no records" do
      block = <<0xBC, 0x00, 4::16-little>>
      assert BlockParser.parse(block, %{}, <<>>, %{}, 0) == %{}
    end

    test "returns empty map for invalid type2 offset before property area" do
      block = <<0x7C, 0x00, 3::16-little, 0, 0, 0, 0>>
      assert BlockParser.parse(block, %{}, <<>>, %{}, 0) == %{}
    end
  end

  describe "parse_table/5" do
    test "wraps parsed map in a single-row list" do
      block =
        <<0xBC, 0x00, 13::16-little, 0x03, 0x00, 0x0E, 0x00, 7::32-little-signed, 0x00, 0x00>>

      assert [row] = BlockParser.parse_table(block, %{}, <<>>, %{}, 0)
      assert row[Key.new(0x000E)] == 7
    end
  end
end
