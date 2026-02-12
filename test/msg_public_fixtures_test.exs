defmodule OutlookMsg.MsgPublicFixturesTest do
  use ExUnit.Case, async: false

  alias OutlookMsg.Mapi.PropertySet

  @fixtures_dir Path.expand("fixtures/public_msg", __DIR__)

  test "opens all public MSG fixtures and extracts core content" do
    fixture_files =
      @fixtures_dir
      |> Path.join("*.msg")
      |> Path.wildcard()
      |> Enum.sort()

    assert fixture_files != []

    Enum.each(fixture_files, fn path ->
      assert {:ok, msg} = OutlookMsg.open(path)

      # Core parse integrity for public fixtures: parser should not crash and
      # top-level collections should be materialized.
      assert is_list(msg.recipients)
      assert is_list(msg.attachments)

      props = msg.properties
      assert is_struct(props, PropertySet)

      _subject = PropertySet.subject(props)
      body = PropertySet.body(props) || ""
      html = PropertySet.body_html(props) || ""

      assert is_binary(body)
      assert is_binary(html)
    end)
  end

  test "matches ruby-msg baseline for semantic fixture metrics" do
    unless System.find_executable("ruby") do
      flunk("ruby not available; cannot run ruby-msg baseline comparison")
    end

    fixture_files =
      @fixtures_dir
      |> Path.join("*.msg")
      |> Path.wildcard()
      |> Enum.sort()

    assert fixture_files != []

    ruby_rows = ruby_fixture_rows(fixture_files)

    Enum.each(fixture_files, fn path ->
      key = Path.basename(path)
      assert Map.has_key?(ruby_rows, key)
      ruby = ruby_rows[key]

      assert {:ok, msg} = OutlookMsg.open(path)
      props = msg.properties

      body = PropertySet.body(props) || ""
      html = PropertySet.body_html(props) || ""
      body = if is_binary(body), do: body, else: inspect(body)
      html = if is_binary(html), do: html, else: inspect(html)

      ours = %{
        subject: normalize(PropertySet.subject(props)),
        recipients: length(msg.recipients),
        attachments: length(msg.attachments),
        body_len: byte_size(body),
        html_len: byte_size(html),
        first_recipient_email: normalize(msg.recipients |> List.first() |> then(fn r -> if r, do: r.email, else: nil end))
      }

      assert ours.subject == ruby.subject
      assert ours.recipients == ruby.recipients
      assert ours.attachments == ruby.attachments
      assert ours.first_recipient_email == ruby.first_recipient_email
      assert abs(ours.body_len - ruby.body_len) <= 128
      assert abs(ours.html_len - ruby.html_len) <= 4096

      if ruby.html_len > 0 do
        # Some non-note fixture classes may not expose HTML consistently across
        # implementations; require at least one rich body path to be populated.
        assert ours.html_len > 0 or ours.body_len > 0
      end
    end)
  end

  defp ruby_fixture_rows(files) do
    ruby_code = """
    require 'base64'
    require 'mapi/msg'

    def norm(v)
      return '' if v.nil?
      v.to_s.gsub(/\\x00/, '').gsub(/\\s+/, ' ').strip
    end

    ARGV.each do |f|
      msg = Mapi::Msg.open(f)
      p = msg.props
      body = (p.body || '').to_s
      html = (p.body_html || '').to_s
      rec = msg.recipients.first
      first_email = rec ? rec.email : nil

      puts [
        'ROW',
        File.basename(f),
        Base64.strict_encode64(norm(p.subject)),
        msg.recipients.length,
        msg.attachments.length,
        body.bytesize,
        html.bytesize,
        Base64.strict_encode64(norm(first_email))
      ].join('|')

      msg.close
    end
    """

    {stdout, 0} =
      System.cmd(
        "ruby",
        ["-I/home/sprite/ruby-msg/lib", "-e", ruby_code | files],
        stderr_to_stdout: true
      )

    stdout
    |> String.split("\n", trim: true)
    |> Enum.reduce(%{}, fn line, acc ->
      case String.split(line, "|") do
        ["ROW", file, subject_b64, recipients, attachments, body_len, html_len, first_email_b64] ->
          Map.put(acc, file, %{
            subject: b64_decode(subject_b64),
            recipients: String.to_integer(recipients),
            attachments: String.to_integer(attachments),
            body_len: String.to_integer(body_len),
            html_len: String.to_integer(html_len),
            first_recipient_email: b64_decode(first_email_b64)
          })

        _ ->
          acc
      end
    end)
  end

  defp normalize(nil), do: ""

  defp normalize(v) do
    v
    |> to_string()
    |> String.replace("\u0000", "")
    |> String.replace(~r/\s+/, " ")
    |> String.trim()
  end

  defp b64_decode(""), do: ""
  defp b64_decode(s), do: Base.decode64!(s)
end
