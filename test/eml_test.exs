defmodule OutlookMsg.EmlTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Mime

  test "open_eml/1 parses raw EML string" do
    raw = [
      "From: Alice <alice@example.com>",
      "To: Bob <bob@example.com>",
      "Subject: Hello",
      "Content-Type: text/plain; charset=utf-8",
      "",
      "Body line 1",
      "Body line 2"
    ] |> Enum.join("\r\n")

    assert {:ok, %Mime{} = mime} = OutlookMsg.open_eml(raw)
    assert Mime.get_header(mime, "Subject") == "Hello"
    assert String.contains?(mime.body || "", "Body line 1")
  end

  test "open_eml/1 parses EML file path" do
    raw = [
      "From: Carol <carol@example.com>",
      "To: Dave <dave@example.com>",
      "Subject: Path Parse",
      "",
      "Test body"
    ] |> Enum.join("\r\n")

    path = Path.join(System.tmp_dir!(), "outlook_msg_test_#{System.unique_integer([:positive])}.eml")
    File.write!(path, raw)

    try do
      assert {:ok, %Mime{} = mime} = OutlookMsg.open_eml(path)
      assert Mime.get_header(mime, "Subject") == "Path Parse"
      assert String.contains?(mime.body || "", "Test body")
    after
      File.rm(path)
    end
  end

  test "eml_to_string/1 serializes parsed MIME" do
    raw = [
      "From: Eve <eve@example.com>",
      "To: Frank <frank@example.com>",
      "Subject: Serialize",
      "",
      "payload"
    ] |> Enum.join("\r\n")

    assert {:ok, %Mime{} = mime} = OutlookMsg.open_eml(raw)
    rendered = OutlookMsg.eml_to_string(mime)
    assert String.contains?(rendered, "Subject: Serialize")
    assert String.contains?(rendered, "payload")
  end

  test "open_eml/1 parses multipart alternative EML" do
    raw = [
      "From: Multipart <multi@example.com>",
      "To: Reader <reader@example.com>",
      "Subject: Multipart Sample",
      "MIME-Version: 1.0",
      "Content-Type: multipart/alternative; boundary=\"b1\"",
      "",
      "--b1",
      "Content-Type: text/plain; charset=utf-8",
      "",
      "plain body",
      "--b1",
      "Content-Type: text/html; charset=utf-8",
      "",
      "<html><body><p>html body</p></body></html>",
      "--b1--",
      ""
    ] |> Enum.join("\r\n")

    assert {:ok, %Mime{} = mime} = OutlookMsg.open_eml(raw)
    assert Mime.multipart?(mime)
    assert length(mime.parts) == 2
    assert Mime.get_header(mime, "Subject") == "Multipart Sample"
  end

  test "open_eml/1 decodes RFC2047 encoded subject value" do
    raw = [
      "From: Encoded <enc@example.com>",
      "To: User <user@example.com>",
      "Subject: =?UTF-8?B?0J/RgNC40LLQtdGC?=",
      "",
      "body"
    ] |> Enum.join("\r\n")

    assert {:ok, %Mime{} = mime} = OutlookMsg.open_eml(raw)
    assert Mime.get_header(mime, "Subject") == "Привет"
  end

  test "open_eml_with_warnings/1 retains body and reports malformed header lines" do
    raw = [
      "From: test@example.com",
      "Malformed Header Without Colon",
      "Subject: Keep Going",
      "",
      "body survives"
    ] |> Enum.join("\r\n")

    assert {:ok, %Mime{} = mime, warnings} = OutlookMsg.open_eml_with_warnings(raw)
    assert Mime.get_header(mime, "Subject") == "Keep Going"
    assert mime.body == "body survives"
    assert Enum.any?(warnings, &String.contains?(&1, "malformed header line ignored"))
  end

  test "open_eml_with_warnings/1 handles multipart without boundary as best effort" do
    raw = [
      "Content-Type: multipart/alternative",
      "",
      "raw multipart-ish body content"
    ] |> Enum.join("\r\n")

    assert {:ok, %Mime{} = mime, warnings} = OutlookMsg.open_eml_with_warnings(raw)
    assert mime.body == "raw multipart-ish body content"
    assert Enum.any?(warnings, &String.contains?(&1, "missing boundary"))
  end
end
