defmodule OutlookMsg.MapiTypesTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Mapi.Types

  @moduletag :spec_conformance

  describe "type metadata" do
    test "returns expected type names for scalar and multi-value types" do
      assert Types.type_name(0x0003) == :pt_long
      assert Types.type_name(0x001F) == :pt_unicode
      assert Types.type_name(0x101F) == :pt_mv_unicode
      assert Types.type_name(0x1102) == :pt_mv_binary
      assert Types.type_name(0x0FFF) == :unknown
    end

    test "computes base_type and multi_value flags" do
      assert Types.base_type(0x101F) == 0x001F
      assert Types.base_type(0x000B) == 0x000B
      assert Types.multi_value?(0x101F)
      refute Types.multi_value?(0x001F)
    end
  end

  describe "decode_value/2" do
    test "decodes integers, booleans, and floating point values" do
      assert Types.decode_value(0x0002, <<0x34, 0x12>>) == 0x1234
      assert Types.decode_value(0x0003, <<0xFF, 0xFF, 0xFF, 0xFF>>) == -1
      assert Types.decode_value(0x000B, <<1, 0, 0, 0>>) == true
      assert Types.decode_value(0x000B, <<0, 0, 0, 0>>) == false

      f = Types.decode_value(0x0004, <<0, 0, 0x80, 0x3F>>)
      assert_in_delta f, 1.0, 0.00001

      d = Types.decode_value(0x0005, <<0, 0, 0, 0, 0, 0, 0xF0, 0x3F>>)
      assert_in_delta d, 1.0, 0.000000001
    end

    test "decodes currency, int64, and binary/string values" do
      assert Types.decode_value(0x0006, <<0x10, 0x27, 0, 0, 0, 0, 0, 0>>) == 1.0
      assert Types.decode_value(0x0014, <<0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x7F>>) == 9_223_372_036_854_775_807
      assert Types.decode_value(0x001E, "abc\0\0") == "abc"
      assert Types.decode_value(0x0102, <<1, 2, 3>>) == <<1, 2, 3>>
    end

    test "decodes utf16 strings and GUIDs" do
      utf16 = <<0x48, 0x00, 0x69, 0x00, 0x00, 0x00>>
      assert Types.decode_value(0x001F, utf16) == "Hi"

      guid = <<0x28, 0x03, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46>>
      assert Types.decode_value(0x0048, guid) == "{00020328-0000-0000-C000-000000000046}"
    end

    test "decodes systime as DateTime when valid and returns binary when invalid" do
      ticks = (1_704_067_200 + 11_644_473_600) * 10_000_000
      ft = <<ticks::64-little>>
      assert %DateTime{} = Types.decode_value(0x0040, ft)

      assert Types.decode_value(0x0040, <<0::64>>) == <<0::64>>
    end

    test "returns raw binary for unknown types" do
      blob = <<1, 2, 3, 4>>
      assert Types.decode_value(0x7FFF, blob) == blob
    end
  end
end
