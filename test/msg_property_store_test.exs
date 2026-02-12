defmodule OutlookMsg.MsgPropertyStoreTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Msg.PropertyStore

  @moduletag :spec_conformance

  describe "decode_substg_name/1" do
    test "parses standard property stream name" do
      assert PropertyStore.decode_substg_name("__substg1.0_0037001F") == {0x0037, 0x001F, nil}
    end

    test "parses multi-value indexed stream name" do
      assert PropertyStore.decode_substg_name("__substg1.0_1000001F-00000002") ==
               {0x1000, 0x001F, 0x00000002}
    end

    test "rejects malformed names" do
      assert PropertyStore.decode_substg_name("__substg1.0_0037001") == nil
      assert PropertyStore.decode_substg_name("__substg1.0_ZZZZ001F") == nil
      assert PropertyStore.decode_substg_name("random_stream_name") == nil
    end
  end
end
