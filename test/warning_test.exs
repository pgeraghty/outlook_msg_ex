defmodule OutlookMsg.WarningTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Warning

  @moduletag :spec_conformance

  test "new/3 sets default fields" do
    w = Warning.new(:malformed_header_line, "ignored malformed line")
    assert w.code == :malformed_header_line
    assert w.severity == :warn
    assert w.recoverable == true
    assert w.context == nil
  end

  test "format/1 includes severity/code/context" do
    w = Warning.new(:property_parse_failed, "property parsing failed", severity: :error, context: "root")
    formatted = Warning.format(w)
    assert formatted =~ "[error:property_parse_failed]"
    assert formatted =~ "root"
    assert formatted =~ "property parsing failed"
  end

  test "format_all/1 handles mixed structured and plain warnings" do
    list = [Warning.new(:a, "x"), "legacy warning text"]
    out = Warning.format_all(list)
    assert length(out) == 2
    assert Enum.at(out, 0) =~ "[warn:a]"
    assert Enum.at(out, 1) == "legacy warning text"
  end
end
