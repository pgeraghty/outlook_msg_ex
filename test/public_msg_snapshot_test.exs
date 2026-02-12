defmodule OutlookMsg.PublicMsgSnapshotTest do
  use ExUnit.Case, async: false

  alias OutlookMsg.Mapi.PropertySet

  @fixtures_dir Path.expand("fixtures/public_msg", __DIR__)
  @snapshot_path Path.expand("../data/snapshots/public_msg_snapshot.exs", __DIR__)

  test "public fixture semantic snapshot stays stable" do
    assert File.exists?(@snapshot_path)

    {expected, _} = Code.eval_file(@snapshot_path)
    assert is_map(expected)

    files =
      @fixtures_dir
      |> Path.join("*.msg")
      |> Path.wildcard()
      |> Enum.sort()

    assert files != []

    actual =
      files
      |> Enum.map(fn file -> {Path.basename(file), metrics_for(file)} end)
      |> Map.new()

    assert Map.keys(actual) |> Enum.sort() == Map.keys(expected) |> Enum.sort()

    Enum.each(Map.keys(expected), fn key ->
      assert Map.fetch!(actual, key) == Map.fetch!(expected, key)
    end)
  end

  defp metrics_for(path) do
    {:ok, msg} = OutlookMsg.open(path)
    p = msg.properties

    body = PropertySet.body(p) || ""
    html = PropertySet.body_html(p) || ""
    body = if is_binary(body), do: body, else: inspect(body)
    html = if is_binary(html), do: html, else: inspect(html)

    first_rec = List.first(msg.recipients)
    first_rec_email = if first_rec, do: first_rec.email, else: nil

    %{
      subject: normalize(PropertySet.subject(p)),
      recipient_count: length(msg.recipients),
      attachment_count: length(msg.attachments),
      first_recipient_email: normalize(first_rec_email),
      body_len: byte_size(body),
      html_len: byte_size(html),
      body_sha256: :crypto.hash(:sha256, body) |> Base.encode16(case: :lower),
      html_sha256: :crypto.hash(:sha256, html) |> Base.encode16(case: :lower)
    }
  end

  defp normalize(nil), do: ""

  defp normalize(v) do
    v
    |> to_string()
    |> String.replace("\u0000", "")
    |> String.replace(~r/\s+/, " ")
    |> String.trim()
  end
end
