defmodule OutlookMsg.SpecConformanceTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Mime
  alias OutlookMsg.Ole.Header, as: OleHeader
  alias OutlookMsg.Mapi.{Key, Tags}
  alias OutlookMsg.Pst.Header

  @moduletag :spec_conformance

  describe "RFC 5322 / RFC 2047 MIME behavior" do
    test "unfolds folded headers" do
      raw = [
        "Subject: hello",
        " world",
        "",
        "body"
      ] |> Enum.join("\r\n")

      mime = Mime.new(raw)
      assert Mime.get_header(mime, "Subject") == "hello world"
    end

    test "decodes encoded-word headers (B and Q)" do
      raw = [
        "From: =?UTF-8?B?Sm9zw6kgTcOhcnRpbg==?= <jose@example.org>",
        "Subject: =?UTF-8?Q?Ren=C3=A9e_update?=",
        "",
        "body"
      ] |> Enum.join("\r\n")

      mime = Mime.new(raw)
      assert Mime.get_header(mime, "From") == "José Mártin <jose@example.org>"
      assert Mime.get_header(mime, "Subject") == "Renée update"
    end

    test "decodes adjacent encoded words with ignored linear whitespace" do
      raw = [
        "Subject: =?UTF-8?B?8J+YgA==?= =?UTF-8?B?IFNwZWM=?=",
        "",
        "body"
      ] |> Enum.join("\r\n")

      mime = Mime.new(raw)
      assert Mime.get_header(mime, "Subject") == "😀 Spec"
    end

    test "folds long header lines on serialization" do
      long_subject = String.duplicate("SubjectWord ", 12)

      rendered =
        Mime.new()
        |> Mime.set_header("Subject", long_subject)
        |> Map.put(:body, "x")
        |> Mime.to_string()

      [header_block | _] = String.split(rendered, "\r\n\r\n", parts: 2)
      subject_lines = header_block |> String.split("\r\n") |> Enum.filter(&String.starts_with?(&1, "Subject:") or String.starts_with?(&1, " "))
      assert length(subject_lines) > 1
      assert Enum.all?(subject_lines, &(String.length(&1) <= 76))
    end
  end

  describe "MS-OXPROPS critical property tags" do
    @critical_tags [
      {0x001A, :pr_message_class, :pt_tstring},
      {0x0037, :pr_subject, :pt_tstring},
      {0x0C1A, :pr_sender_name, :pt_tstring},
      {0x0C1F, :pr_sender_email_address, :pt_tstring},
      {0x0E04, :pr_display_to, :pt_tstring},
      {0x1000, :pr_body, :pt_tstring},
      {0x1009, :pr_rtf_compressed, :pt_binary},
      {0x1013, :pr_body_html, :pt_binary},
      {0x3705, :pr_attach_method, :pt_long},
      {0x3712, :pr_attach_content_id, :pt_tstring},
      {0x3713, :pr_attach_content_location, :pt_tstring},
      {0x3716, :pr_attach_content_disposition, :pt_tstring},
      {0x39FE, :pr_smtp_address, :pt_tstring},
      {0x403E, :pr_org_email_addr, :pt_tstring}
    ]

    test "lookup/1 returns expected names and types for critical tags" do
      Enum.each(@critical_tags, fn {code, expected_name, expected_type} ->
        assert {expected_name, expected_type} == Tags.lookup(code)
        assert expected_name == Tags.name(code)
      end)
    end

    test "Key.to_sym resolves standard PS_MAPI tags" do
      assert Key.new(0x0037) |> Key.to_sym() == :pr_subject
      assert Key.new(0x1000) |> Key.to_sym() == :pr_body
    end
  end

  describe "MS-PST header parsing strictness" do
    test "returns structured error for unknown index type" do
      header = <<0x21, 0x42, 0x44, 0x4E, 0, 0, 0, 0, 0, 0, 0x99>> <> :binary.copy(<<0>>, 512 - 11)
      assert {:error, {:unknown_index_type, 0x99}} = Header.parse(header)
    end

    test "returns error (not crash) for truncated PST 2003 header" do
      # Valid magic + PST2003 index type, but missing bytes past offset 0x0201.
      truncated = <<0x21, 0x42, 0x44, 0x4E, 0, 0, 0, 0, 0, 0, 0x17>> <> :binary.copy(<<0>>, 512 - 11)
      assert {:error, :data_too_short} = Header.parse(truncated)
    end
  end

  describe "MS-CFB OLE header strictness" do
    defp valid_ole_header(overrides) do
      magic = <<0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1>>
      clsid = <<0::128>>
      minor_version = <<62, 0>>
      major_version = Map.get(overrides, :major_version, <<3, 0>>)
      byte_order = Map.get(overrides, :byte_order, <<0xFE, 0xFF>>)
      sector_shift = Map.get(overrides, :sector_shift, <<9, 0>>)
      mini_shift = Map.get(overrides, :mini_shift, <<6, 0>>)
      reserved = <<0::48>>
      dir_sector_count = <<0::32>>
      fat_count = <<1, 0, 0, 0>>
      dir_start = <<0, 0, 0, 0>>
      txn_sig = <<0::32>>
      mini_cutoff = Map.get(overrides, :mini_cutoff, <<0, 16, 0, 0>>)
      mini_fat_start = <<0xFE, 0xFF, 0xFF, 0xFF>>
      mini_fat_count = <<0::32>>
      difat_start = <<0xFE, 0xFF, 0xFF, 0xFF>>
      difat_count = <<0::32>>
      difat = <<0, 0, 0, 0>> <> :binary.copy(<<0xFF, 0xFF, 0xFF, 0xFF>>, 108)

      magic <> clsid <> minor_version <> major_version <>
        byte_order <> sector_shift <> mini_shift <> reserved <>
        dir_sector_count <> fat_count <> dir_start <> txn_sig <>
        mini_cutoff <> mini_fat_start <> mini_fat_count <>
        difat_start <> difat_count <> difat
    end

    test "rejects invalid byte order" do
      assert {:error, msg} = OleHeader.parse(valid_ole_header(%{byte_order: <<0xFF, 0xFE>>}))
      assert String.contains?(msg, "invalid byte order")
    end

    test "rejects invalid sector shift for v3" do
      assert {:error, msg} = OleHeader.parse(valid_ole_header(%{sector_shift: <<12, 0>>}))
      assert String.contains?(msg, "invalid sector shift for v3")
    end

    test "rejects invalid mini stream cutoff" do
      assert {:error, msg} = OleHeader.parse(valid_ole_header(%{mini_cutoff: <<0, 8, 0, 0>>}))
      assert String.contains?(msg, "invalid mini stream cutoff")
    end
  end
end
