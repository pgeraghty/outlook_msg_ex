defmodule OutlookMsg.MapiTagsTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Mapi.Tags

  @moduletag :spec_conformance

  describe "lookup/name/code conformance" do
    test "round-trips known standard properties" do
      known = [
        0x001A,
        0x0037,
        0x0C1F,
        0x0E04,
        0x1000,
        0x1009,
        0x1013,
        0x3705,
        0x3712,
        0x3713,
        0x3716,
        0x39FE,
        0x403E
      ]

      Enum.each(known, fn code ->
        assert {name, _type} = Tags.lookup(code)
        assert Tags.name(code) == name
        assert is_integer(Tags.code(name))
      end)
    end

    test "uses canonical code for alias names with multiple codes" do
      assert Tags.code(:pr_smtp_address) == 0x39FE
      assert Tags.name(0x39FE) == :pr_smtp_address
      assert Tags.name(0x5909) == :pr_smtp_address
    end

    test "returns nil for unknown properties" do
      assert Tags.lookup(0xFFFF) == nil
      assert Tags.name(0xFFFF) == nil
      assert Tags.code(:pr_nonexistent_property) == nil
    end
  end
end
