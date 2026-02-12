defmodule OutlookMsg.WarningsApiTest do
  use ExUnit.Case, async: true
  alias OutlookMsg.Warning

  test "open_with_warnings/1 returns message and warnings list" do
    fixture = Path.expand("fixtures/public_msg/note.msg", __DIR__)
    assert {:ok, msg, warnings} = OutlookMsg.open_with_warnings(fixture)
    assert is_list(warnings)
    assert is_struct(msg, OutlookMsg.Msg)
  end

  test "open_pst_with_warnings/1 preserves existing error behavior for invalid PST" do
    assert {:error, :invalid_pst_magic} = OutlookMsg.open_pst_with_warnings(:binary.copy(<<0>>, 512))
  end

  test "open_pst_with_warnings/1 handles short binary as PST data (no path badarg)" do
    assert {:error, :data_too_short} = OutlookMsg.open_pst_with_warnings(<<1, 2, 3>>)
  end

  test "open_eml_with_report/1 returns structured warnings" do
    raw = [
      "BadHeaderWithoutColon",
      "Subject: Hi",
      "",
      "Body"
    ] |> Enum.join("\r\n")

    assert {:ok, _mime, warnings} = OutlookMsg.open_eml_with_report(raw)
    assert Enum.any?(warnings, &match?(%Warning{}, &1))
    assert Enum.any?(warnings, &(to_string(&1) =~ "malformed header line"))
  end
end
