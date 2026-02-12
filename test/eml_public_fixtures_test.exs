defmodule OutlookMsg.EmlPublicFixturesTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Mime

  @fixtures_dir Path.expand("fixtures/public_eml", __DIR__)

  @expectations %{
    "simple_plain.eml" => %{subject: "RFC-like plain sample", multipart: false},
    "folded_headers.eml" => %{subject: "This is a folded header subject", multipart: false},
    "multipart_alternative.eml" => %{subject: "Multipart alternative sample", multipart: true, parts: 2},
    "multipart_mixed_attachment.eml" => %{subject: "Multipart mixed attachment sample", multipart: true, parts: 2},
    "encoded_word_headers.eml" => %{subject: "😀 Encoded Subject", multipart: false},
    "unix_newlines.eml" => %{subject: "LF-only fixture", multipart: false}
  }

  test "all public eml fixtures parse and preserve core semantics" do
    files = Path.wildcard(Path.join(@fixtures_dir, "*.eml")) |> Enum.sort()
    assert files != []

    Enum.each(files, fn file ->
      base = Path.basename(file)
      expected = Map.fetch!(@expectations, base)

      assert {:ok, mime} = OutlookMsg.open_eml(file)
      assert Mime.get_header(mime, "Subject") == expected.subject
      assert Mime.multipart?(mime) == expected.multipart

      if expected[:parts] do
        assert length(mime.parts) == expected.parts
      else
        assert is_binary(mime.body || "")
      end

      rendered = OutlookMsg.eml_to_string(mime)
      assert {:ok, reparsed} = OutlookMsg.open_eml(rendered)
      assert Mime.get_header(reparsed, "Subject") == expected.subject
      assert Mime.multipart?(reparsed) == expected.multipart
    end)
  end
end
