defmodule OutlookMsg.RtfTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Rtf

  describe "decompress/1" do
    test "handles uncompressed RTF (MELA magic)" do
      rtf_content = "{\\rtf1 Hello World}"
      raw_size = byte_size(rtf_content)
      compr_size = raw_size + 12
      magic = 0x414C454D  # "MELA"
      crc = 0

      header = <<compr_size::32-little, raw_size::32-little, magic::32-little, crc::32-little>>
      data = header <> rtf_content

      assert {:ok, result} = Rtf.decompress(data)
      assert result == rtf_content
    end

    test "rejects invalid magic" do
      data = <<20::32-little, 10::32-little, 0::32-little, 0::32-little, "test">>
      assert {:error, :invalid_magic} = Rtf.decompress(data)
    end

    test "handles empty input" do
      assert {:error, :invalid_header} = Rtf.decompress(<<>>)
    end

    test "handles too-short input" do
      assert {:error, :invalid_header} = Rtf.decompress(<<1, 2, 3, 4>>)
    end

    test "uncompressed truncates to raw_size" do
      rtf_content = "{\\rtf1 Hello World} extra trailing data"
      raw_size = 19  # Only "{\\rtf1 Hello World}"
      compr_size = byte_size(rtf_content) + 12
      magic = 0x414C454D
      crc = 0

      header = <<compr_size::32-little, raw_size::32-little, magic::32-little, crc::32-little>>
      data = header <> rtf_content

      assert {:ok, result} = Rtf.decompress(data)
      assert byte_size(result) == raw_size
    end
  end

  describe "rtf_to_html/1" do
    test "returns :none for plain RTF without fromhtml" do
      rtf = "{\\rtf1\\ansi Hello World}"
      assert Rtf.rtf_to_html(rtf) == :none
    end

    test "extracts HTML from RTF with fromhtml marker" do
      rtf = "{\\rtf1\\ansi\\fromhtml1 {\\*\\htmltag <html>}{\\*\\htmltag <body>}Hello{\\*\\htmltag </body>}{\\*\\htmltag </html>}}"
      case Rtf.rtf_to_html(rtf) do
        :none -> flunk("Expected HTML extraction")
        {:ok, html} ->
          assert String.contains?(html, "<html>")
          assert String.contains?(html, "<body>")
      end
    end

    test "returns :none when no fromhtml marker present" do
      rtf = "{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 Courier;}} Hello}"
      assert :none = Rtf.rtf_to_html(rtf)
    end

    test "keeps plain text outside htmltag groups" do
      rtf = "{\\rtf1\\ansi\\fromhtml1 {\\*\\htmltag <p>}Hello world{\\*\\htmltag </p>}}"
      assert {:ok, html} = Rtf.rtf_to_html(rtf)
      assert String.contains?(html, "Hello world")
    end

    test "prefers mhtmltag over matching htmltag" do
      rtf = "{\\rtf1\\ansi\\fromhtml1 {\\*\\mhtmltag84 <img src=\\\"cid:abc\\\"/>}{\\*\\htmltag84 <img src=\\\"abc.jpg\\\"/>}}"
      assert {:ok, html} = Rtf.rtf_to_html(rtf)
      assert String.contains?(html, "cid:abc")
      refute String.contains?(html, "abc.jpg")
    end

    test "keeps htmlrtf0 plain text segments" do
      rtf =
        "{\\rtf1\\ansi\\fromhtml1 {\\*\\htmltag <p>}\\htmlrtf0Hello\\htmlrtf {\\*\\htmltag </p>}}"

      assert {:ok, html} = Rtf.rtf_to_html(rtf)
      assert String.contains?(html, "Hello")
    end

    test "decodes escaped quote bytes from html tags" do
      rtf =
        "{\\rtf1\\ansi\\fromhtml1 {\\*\\htmltag <div id=\\'94x\\'94>}Text{\\*\\htmltag </div>}}"

      assert {:ok, html} = Rtf.rtf_to_html(rtf)
      assert String.contains?(html, "<div id=")
      assert String.contains?(html, "Text")
    end
  end

  describe "rtf_to_text/1" do
    test "extracts plain text from RTF" do
      rtf = "{\\rtf1\\ansi Hello World}"
      result = Rtf.rtf_to_text(rtf)
      assert String.contains?(result, "Hello World")
    end

    test "converts par to newline" do
      rtf = "{\\rtf1\\ansi Line1\\par Line2}"
      result = Rtf.rtf_to_text(rtf)
      assert String.contains?(result, "Line1")
      assert String.contains?(result, "Line2")
    end

    test "handles escaped braces" do
      rtf = "{\\rtf1\\ansi \\{braces\\}}"
      result = Rtf.rtf_to_text(rtf)
      assert String.contains?(result, "{braces}")
    end

    test "handles escaped backslash" do
      rtf = "{\\rtf1\\ansi back\\\\slash}"
      result = Rtf.rtf_to_text(rtf)
      assert String.contains?(result, "back\\slash")
    end

    test "strips special destination groups" do
      rtf = "{\\rtf1\\ansi {\\*\\fonttbl hidden}visible text}"
      result = Rtf.rtf_to_text(rtf)
      refute String.contains?(result, "hidden")
      assert String.contains?(result, "visible text")
    end

    test "handles tab control word" do
      rtf = "{\\rtf1\\ansi col1\\tab col2}"
      result = Rtf.rtf_to_text(rtf)
      assert String.contains?(result, "col1")
      assert String.contains?(result, "col2")
    end
  end
end
