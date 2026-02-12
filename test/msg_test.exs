defmodule OutlookMsg.MsgTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Msg.{Recipient, Attachment}
  alias OutlookMsg.Mapi.{PropertySet, Key}

  describe "Recipient" do
    test "creates recipient from property set" do
      props = PropertySet.new(%{
        Key.new(0x3A20) => "John Doe",        # PR_TRANSMITTABLE_DISPLAY_NAME
        Key.new(0x5909) => "john@example.com", # PR_SMTP_ADDRESS
        Key.new(0x0C15) => 1                   # PR_RECIPIENT_TYPE = TO
      })

      recipient = Recipient.new(props)
      assert recipient.name == "John Doe"
      assert recipient.email == "john@example.com"
      assert recipient.type == :to
    end

    test "handles CC recipient type" do
      props = PropertySet.new(%{
        Key.new(0x3A20) => "Jane",
        Key.new(0x5909) => "jane@test.com",
        Key.new(0x0C15) => 2
      })

      recipient = Recipient.new(props)
      assert recipient.type == :cc
    end

    test "handles BCC recipient type" do
      props = PropertySet.new(%{
        Key.new(0x3A20) => "Bob",
        Key.new(0x5909) => "bob@test.com",
        Key.new(0x0C15) => 3
      })

      recipient = Recipient.new(props)
      assert recipient.type == :bcc
    end

    test "defaults to :to when recipient type not set" do
      props = PropertySet.new(%{
        Key.new(0x3A20) => "No Type",
        Key.new(0x5909) => "noone@test.com"
      })

      recipient = Recipient.new(props)
      assert recipient.type == :to
    end

    test "falls back to display_name when transmittable name missing" do
      props = PropertySet.new(%{
        Key.new(0x3001) => "Display Name",     # PR_DISPLAY_NAME
        Key.new(0x5909) => "display@test.com",
        Key.new(0x0C15) => 1
      })

      recipient = Recipient.new(props)
      assert recipient.name == "Display Name"
    end

    test "falls back to email_address when smtp_address missing" do
      props = PropertySet.new(%{
        Key.new(0x3A20) => "Name",
        Key.new(0x3003) => "fallback@test.com",  # PR_EMAIL_ADDRESS
        Key.new(0x0C15) => 1
      })

      recipient = Recipient.new(props)
      assert recipient.email == "fallback@test.com"
    end

    test "to_string formats name and email" do
      props = PropertySet.new(%{
        Key.new(0x3A20) => "Jane",
        Key.new(0x5909) => "jane@test.com",
        Key.new(0x0C15) => 2
      })

      recipient = Recipient.new(props)
      str = Recipient.to_string(recipient)
      assert String.contains?(str, "Jane")
      assert String.contains?(str, "jane@test.com")
    end

    test "to_string with nil name shows just email" do
      props = PropertySet.new(%{
        Key.new(0x5909) => "only@email.com",
        Key.new(0x0C15) => 1
      })

      recipient = Recipient.new(props)
      str = Recipient.to_string(recipient)
      assert str == "<only@email.com>"
    end

    test "to_string with nil email shows just name" do
      props = PropertySet.new(%{
        Key.new(0x3A20) => "Just Name",
        Key.new(0x0C15) => 1
      })

      recipient = Recipient.new(props)
      str = Recipient.to_string(recipient)
      assert str == "Just Name"
    end

    test "String.Chars protocol works" do
      props = PropertySet.new(%{
        Key.new(0x3A20) => "Proto",
        Key.new(0x5909) => "proto@test.com",
        Key.new(0x0C15) => 1
      })

      recipient = Recipient.new(props)
      assert "#{recipient}" == ~s("Proto" <proto@test.com>)
    end
  end

  describe "Attachment" do
    test "creates attachment from property set" do
      props = PropertySet.new(%{
        Key.new(0x3707) => "document.pdf",    # PR_ATTACH_LONG_FILENAME
        Key.new(0x3701) => "PDF content",     # PR_ATTACH_DATA_BIN
        Key.new(0x370E) => "application/pdf"  # PR_ATTACH_MIME_TAG
      })

      att = Attachment.new(props)
      assert att.filename == "document.pdf"
      assert att.data == "PDF content"
      assert att.mime_type == "application/pdf"
    end

    test "falls back to short filename" do
      props = PropertySet.new(%{
        Key.new(0x3704) => "doc.pdf",  # PR_ATTACH_FILENAME (short)
        Key.new(0x3701) => "data"
      })

      att = Attachment.new(props)
      assert att.filename == "doc.pdf"
    end

    test "uses default filename when none specified" do
      props = PropertySet.new(%{
        Key.new(0x3701) => "data"
      })

      att = Attachment.new(props)
      assert att.filename == "attachment"
    end

    test "inline? checks content_id" do
      props = PropertySet.new(%{
        Key.new(0x3707) => "image.png",
        Key.new(0x3712) => "cid:image001",  # PR_ATTACH_CONTENT_ID
        Key.new(0x3701) => "png data"
      })

      att = Attachment.new(props)
      assert Attachment.inline?(att)
    end

    test "inline? checks content_location" do
      props = PropertySet.new(%{
        Key.new(0x3707) => "image.png",
        Key.new(0x3713) => "http://example.com/image.png",  # PR_ATTACH_CONTENT_LOCATION
        Key.new(0x3701) => "png data"
      })

      att = Attachment.new(props)
      assert Attachment.inline?(att)
    end

    test "not inline when no content_id or content_location" do
      props = PropertySet.new(%{
        Key.new(0x3707) => "document.pdf",
        Key.new(0x3701) => "pdf data"
      })

      att = Attachment.new(props)
      refute Attachment.inline?(att)
    end

    test "embedded? returns false by default" do
      props = PropertySet.new(%{
        Key.new(0x3707) => "doc.pdf",
        Key.new(0x3701) => "data"
      })

      att = Attachment.new(props)
      refute Attachment.embedded?(att)
    end

    test "content_id accessor" do
      props = PropertySet.new(%{
        Key.new(0x3712) => "image001@host"
      })

      att = Attachment.new(props)
      assert Attachment.content_id(att) == "image001@host"
    end

    test "extension accessor" do
      props = PropertySet.new(%{
        Key.new(0x3703) => ".pdf"  # PR_ATTACH_EXTENSION
      })

      att = Attachment.new(props)
      assert Attachment.extension(att) == ".pdf"
    end

    test "method accessor" do
      props = PropertySet.new(%{
        Key.new(0x3705) => 1  # PR_ATTACH_METHOD
      })

      att = Attachment.new(props)
      assert Attachment.method(att) == 1
    end
  end
end
