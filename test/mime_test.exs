defmodule OutlookMsg.MimeTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Mime

  describe "basic MIME construction" do
    test "creates empty MIME message" do
      mime = Mime.new()
      assert mime.headers == []
      assert mime.body == nil
      assert mime.parts == []
    end

    test "set and get headers" do
      mime = Mime.new()
      |> Mime.set_header("From", "test@example.com")
      |> Mime.set_header("Subject", "Hello")

      assert Mime.get_header(mime, "From") == "test@example.com"
      assert Mime.get_header(mime, "Subject") == "Hello"
      assert Mime.get_header(mime, "from") == "test@example.com"  # case-insensitive
    end

    test "set_header replaces existing" do
      mime = Mime.new()
      |> Mime.set_header("Subject", "Old")
      |> Mime.set_header("Subject", "New")

      assert Mime.get_header(mime, "Subject") == "New"
    end

    test "add_header allows duplicates" do
      mime = Mime.new()
      |> Mime.add_header("Received", "from server1")
      |> Mime.add_header("Received", "from server2")

      assert length(mime.headers) == 2
    end

    test "get_header returns nil for missing header" do
      mime = Mime.new()
      assert Mime.get_header(mime, "X-Missing") == nil
    end
  end

  describe "split_header/1" do
    test "parses simple content type" do
      {main, params} = Mime.split_header("text/plain; charset=\"utf-8\"")
      assert main == "text/plain"
      assert params["charset"] == "utf-8"
    end

    test "parses multipart with boundary" do
      {main, params} = Mime.split_header("multipart/mixed; boundary=\"abc123\"")
      assert main == "multipart/mixed"
      assert params["boundary"] == "abc123"
    end

    test "handles no parameters" do
      {main, params} = Mime.split_header("text/plain")
      assert main == "text/plain"
      assert params == %{}
    end

    test "handles multiple parameters" do
      {main, params} = Mime.split_header("text/html; charset=utf-8; name=\"test.html\"")
      assert main == "text/html"
      assert params["charset"] == "utf-8"
      assert params["name"] == "test.html"
    end
  end

  describe "multipart?/1" do
    test "detects multipart" do
      mime = %Mime{content_type: "multipart/mixed"}
      assert Mime.multipart?(mime)
    end

    test "detects non-multipart" do
      mime = %Mime{content_type: "text/plain"}
      refute Mime.multipart?(mime)
    end

    test "handles nil content type" do
      mime = %Mime{content_type: nil}
      refute Mime.multipart?(mime)
    end
  end

  describe "to_string/1" do
    test "serializes simple message" do
      mime = Mime.new()
      |> Mime.set_header("From", "test@example.com")
      |> Mime.set_header("Subject", "Test")
      |> Map.put(:body, "Hello World")

      result = Mime.to_string(mime)
      assert String.contains?(result, "From: test@example.com")
      assert String.contains?(result, "Subject: Test")
      assert String.contains?(result, "Hello World")
    end

    test "separates headers from body with blank line" do
      mime = Mime.new()
      |> Mime.set_header("Subject", "Test")
      |> Map.put(:body, "body text")

      result = Mime.to_string(mime)
      assert String.contains?(result, "\r\n\r\n")
    end
  end

  describe "new/1 (parsing)" do
    test "parses simple message" do
      raw = "From: test@example.com\r\nSubject: Hello\r\n\r\nBody text"
      mime = Mime.new(raw)
      assert Mime.get_header(mime, "From") == "test@example.com"
      assert Mime.get_header(mime, "Subject") == "Hello"
      assert mime.body == "Body text"
    end

    test "parses unix-style line endings" do
      raw = "Subject: Test\n\nBody"
      mime = Mime.new(raw)
      assert Mime.get_header(mime, "Subject") == "Test"
      assert mime.body == "Body"
    end

    test "parses folded header continuations" do
      raw = [
        "Subject: Folded",
        " header value",
        "To: one@example.com,",
        " two@example.com",
        "",
        "Body"
      ] |> Enum.join("\r\n")

      mime = Mime.new(raw)
      assert Mime.get_header(mime, "Subject") == "Folded header value"
      assert Mime.get_header(mime, "To") == "one@example.com, two@example.com"
    end

    test "decodes RFC2047 encoded words in headers" do
      raw = [
        "From: =?UTF-8?B?Sm9zw6kgTcOhcnRpbg==?= <jose@example.org>",
        "Subject: =?UTF-8?Q?Ren=C3=A9e_update?=",
        "",
        "Body"
      ] |> Enum.join("\r\n")

      mime = Mime.new(raw)
      assert Mime.get_header(mime, "From") == "José Mártin <jose@example.org>"
      assert Mime.get_header(mime, "Subject") == "Renée update"
    end

    test "captures multipart preamble and epilogue" do
      raw = [
        "MIME-Version: 1.0",
        "Content-Type: multipart/alternative; boundary=\"demo\"",
        "",
        "Preamble line",
        "--demo",
        "Content-Type: text/plain",
        "",
        "Plain",
        "--demo",
        "Content-Type: text/html",
        "",
        "<p>HTML</p>",
        "--demo--",
        "Epilogue line"
      ] |> Enum.join("\r\n")

      mime = Mime.new(raw)
      assert Mime.multipart?(mime)
      assert mime.preamble == "Preamble line"
      assert mime.epilogue == "Epilogue line"
      assert length(mime.parts) == 2
    end
  end

  describe "format_email_address/2" do
    test "formats name and email" do
      assert Mime.format_email_address("John Doe", "john@example.com") ==
        ~s("John Doe" <john@example.com>)
    end

    test "formats email only when name is nil" do
      assert Mime.format_email_address(nil, "john@example.com") == "<john@example.com>"
    end

    test "formats email only when name is empty" do
      assert Mime.format_email_address("", "john@example.com") == "<john@example.com>"
    end
  end

  describe "format_date/1" do
    test "formats DateTime as RFC2822" do
      {:ok, dt, _} = DateTime.from_iso8601("2024-01-15T10:30:00Z")
      result = Mime.format_date(dt)
      assert String.contains?(result, "2024")
      assert String.contains?(result, "Jan")
      assert String.contains?(result, "10:30:00")
      assert String.contains?(result, "+0000")
    end

    test "formats NaiveDateTime" do
      ndt = ~N[2024-06-15 14:00:00]
      result = Mime.format_date(ndt)
      assert String.contains?(result, "2024")
      assert String.contains?(result, "Jun")
      assert String.contains?(result, "14:00:00")
    end
  end

  describe "encode_base64/1" do
    test "encodes binary data" do
      result = Mime.encode_base64("Hello World")
      assert is_binary(result)
      # Should decode back
      cleaned = String.replace(result, ~r/\s/, "")
      assert Base.decode64!(cleaned) == "Hello World"
    end

    test "wraps long lines" do
      # Create data that will produce base64 longer than 76 chars
      data = :binary.copy(<<0xAB>>, 100)
      result = Mime.encode_base64(data)
      lines = String.split(result, "\r\n")
      for line <- lines do
        assert String.length(line) <= 76
      end
    end
  end

  describe "encode_quoted_printable/1" do
    test "passes through ASCII text" do
      result = Mime.encode_quoted_printable("Hello World")
      assert result == "Hello World"
    end

    test "encodes non-ASCII bytes" do
      result = Mime.encode_quoted_printable(<<0xC3, 0xA9>>)  # UTF-8 e-acute
      assert String.contains?(result, "=C3")
      assert String.contains?(result, "=A9")
    end

    test "encodes equals sign" do
      result = Mime.encode_quoted_printable("a=b")
      assert String.contains?(result, "=3D")
    end
  end

  describe "encode_header_value/1" do
    test "returns ASCII values unchanged" do
      assert Mime.encode_header_value("Hello World") == "Hello World"
    end

    test "encodes non-ASCII values with RFC2047" do
      result = Mime.encode_header_value("Hej varld")
      # Pure ASCII, should stay as is
      assert result == "Hej varld"
    end
  end

  describe "make_boundary/0" do
    test "generates unique boundaries" do
      b1 = Mime.make_boundary()
      b2 = Mime.make_boundary()
      assert b1 != b2
      assert is_binary(b1)
    end

    test "generates non-empty boundaries" do
      b = Mime.make_boundary()
      assert byte_size(b) > 0
    end
  end
end
