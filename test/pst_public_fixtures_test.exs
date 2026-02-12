defmodule OutlookMsg.PstPublicFixturesTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Warning

  @fixtures_dir Path.expand("fixtures/public_pst", __DIR__)

  test "public pst fixtures parse with expected recovery behavior" do
    minimal = Path.join(@fixtures_dir, "minimal_pst97.pst")
    corrupt = Path.join(@fixtures_dir, "corrupt_offsets_pst97.pst")
    looped = Path.join(@fixtures_dir, "loop_branch_index_pst97.pst")

    assert {:ok, pst, warnings} = OutlookMsg.open_pst_with_report(minimal)
    assert pst.header.version == :pst97
    assert warnings == []

    assert {:ok, pst2, warnings2} = OutlookMsg.open_pst_with_report(corrupt)
    assert pst2.header.version == :pst97
    assert is_list(warnings2)
    assert Enum.any?(warnings2, fn
      %Warning{code: :pst_index_parse_failed} -> true
      %Warning{code: :pst_descriptor_parse_failed} -> true
      _ -> false
    end)

    assert {:ok, pst3, warnings3} = OutlookMsg.open_pst_with_report(looped)
    assert pst3.header.version == :pst97
    assert Enum.any?(warnings3, fn
      %Warning{code: :pst_branch_loop_detected} -> true
      _ -> false
    end)
  end
end
