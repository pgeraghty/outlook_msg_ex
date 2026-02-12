defmodule OutlookMsg.ConvertTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.{Convert, Mime, Msg}
  alias OutlookMsg.Msg.{Recipient, Attachment}
  alias OutlookMsg.Mapi.{PropertySet, Key}

  describe "to_mime/1" do
    test "converts simple message to MIME" do
      msg = build_simple_msg()
      assert {:ok, mime} = Convert.to_mime(msg)
      assert is_struct(mime, Mime)

      # Check Subject header
      assert Mime.get_header(mime, "Subject") == "Test Subject"

      # Check From header
      from = Mime.get_header(mime, "From")
      assert from != nil
      assert String.contains?(from, "sender@test.com")
    end

    test "sets MIME-Version header" do
      msg = build_simple_msg()
      assert {:ok, mime} = Convert.to_mime(msg)
      assert Mime.get_header(mime, "MIME-Version") == "1.0"
    end

    test "includes recipients in headers" do
      msg = build_msg_with_recipients()
      assert {:ok, mime} = Convert.to_mime(msg)

      to = Mime.get_header(mime, "To")
      assert to != nil
      assert String.contains?(to, "recipient@test.com")
    end

    test "includes CC recipients" do
      msg = build_msg_with_cc()
      assert {:ok, mime} = Convert.to_mime(msg)

      cc = Mime.get_header(mime, "Cc")
      assert cc != nil
      assert String.contains?(cc, "cc@test.com")
    end

    test "converts message with text body" do
      msg = build_simple_msg()
      assert {:ok, mime} = Convert.to_mime(msg)

      ct = Mime.get_header(mime, "Content-Type")
      assert ct != nil
      assert String.contains?(ct, "text/plain")
    end

    test "converts message with html body" do
      props = PropertySet.new(%{
        Key.new(0x0037) => "HTML Subject",
        Key.new(0x0C1A) => "Sender",
        Key.new(0x0C1F) => "sender@test.com",
        Key.new(0x0C1E) => "SMTP",
        Key.new(0x1013) => "<html><body>Hello</body></html>"  # PR_BODY_HTML
      })

      msg = %Msg{
        properties: props,
        attachments: [],
        recipients: [],
        storage: nil
      }

      assert {:ok, mime} = Convert.to_mime(msg)
      # Should have HTML content type somewhere
      result = Mime.to_string(mime)
      assert String.contains?(result, "text/html") or String.contains?(result, "Hello")
    end

    test "handles message with no body" do
      props = PropertySet.new(%{
        Key.new(0x0037) => "No Body",
        Key.new(0x0C1A) => "Sender",
        Key.new(0x0C1F) => "sender@test.com",
        Key.new(0x0C1E) => "SMTP"
      })

      msg = %Msg{
        properties: props,
        attachments: [],
        recipients: [],
        storage: nil
      }

      assert {:ok, mime} = Convert.to_mime(msg)
      assert is_struct(mime, Mime)
    end

    test "handles attachments" do
      att_props = PropertySet.new(%{
        Key.new(0x3707) => "file.txt",
        Key.new(0x3701) => "file content",
        Key.new(0x370E) => "text/plain"
      })
      att = Attachment.new(att_props)

      msg = %Msg{
        properties: build_simple_props(),
        attachments: [att],
        recipients: [],
        storage: nil
      }

      assert {:ok, mime} = Convert.to_mime(msg)

      # With attachments, should be multipart/mixed
      ct = Mime.get_header(mime, "Content-Type")
      assert ct != nil
      assert String.contains?(ct, "multipart/mixed")
    end

    test "handles importance header" do
      props = PropertySet.new(%{
        Key.new(0x0037) => "Important",
        Key.new(0x0C1A) => "Sender",
        Key.new(0x0C1F) => "sender@test.com",
        Key.new(0x0C1E) => "SMTP",
        Key.new(0x1000) => "body",
        Key.new(0x0017) => 2  # PR_IMPORTANCE = high
      })

      msg = %Msg{
        properties: props,
        attachments: [],
        recipients: [],
        storage: nil
      }

      assert {:ok, mime} = Convert.to_mime(msg)
      assert Mime.get_header(mime, "Importance") == "high"
      assert Mime.get_header(mime, "X-Priority") == "1"
    end

    test "handles sensitivity header" do
      props = PropertySet.new(%{
        Key.new(0x0037) => "Private",
        Key.new(0x0C1A) => "Sender",
        Key.new(0x0C1F) => "sender@test.com",
        Key.new(0x0C1E) => "SMTP",
        Key.new(0x1000) => "body",
        Key.new(0x0036) => 2  # PR_SENSITIVITY = private
      })

      msg = %Msg{
        properties: props,
        attachments: [],
        recipients: [],
        storage: nil
      }

      assert {:ok, mime} = Convert.to_mime(msg)
      assert Mime.get_header(mime, "Sensitivity") == "Private"
    end

    test "serializes to valid RFC2822 string" do
      msg = build_msg_with_recipients()
      assert {:ok, mime} = Convert.to_mime(msg)

      result = Mime.to_string(mime)
      assert is_binary(result)
      assert String.contains?(result, "Subject: Test Subject")
      assert String.contains?(result, "MIME-Version: 1.0")
    end
  end

  # -------------------------------------------------------------------
  # Test helpers
  # -------------------------------------------------------------------

  defp build_simple_props do
    PropertySet.new(%{
      Key.new(0x0037) => "Test Subject",
      Key.new(0x0C1A) => "Sender Name",
      Key.new(0x0C1F) => "sender@test.com",
      Key.new(0x0C1E) => "SMTP",
      Key.new(0x1000) => "Hello, this is the body."
    })
  end

  defp build_simple_msg do
    %Msg{
      properties: build_simple_props(),
      attachments: [],
      recipients: [],
      storage: nil
    }
  end

  defp build_msg_with_recipients do
    msg = build_simple_msg()
    recipient_props = PropertySet.new(%{
      Key.new(0x3A20) => "Recipient",
      Key.new(0x5909) => "recipient@test.com",
      Key.new(0x0C15) => 1
    })
    recipient = Recipient.new(recipient_props)
    %{msg | recipients: [recipient]}
  end

  defp build_msg_with_cc do
    msg = build_msg_with_recipients()
    cc_props = PropertySet.new(%{
      Key.new(0x3A20) => "CC Person",
      Key.new(0x5909) => "cc@test.com",
      Key.new(0x0C15) => 2
    })
    cc_recipient = Recipient.new(cc_props)
    %{msg | recipients: msg.recipients ++ [cc_recipient]}
  end
end
