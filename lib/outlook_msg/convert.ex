defmodule OutlookMsg.Convert do
  @moduledoc """
  Converts an MSG message to a MIME/RFC2822 email.

  Based on note-mime.rb and convert.rb from ruby-msg. Takes a parsed
  `%OutlookMsg.Msg{}` struct and produces a `%OutlookMsg.Mime{}` struct
  that can be serialized to a standard RFC2822 email string.
  """

  alias OutlookMsg.{Msg, Mime}
  alias OutlookMsg.Msg.{Attachment, Recipient}
  alias OutlookMsg.Mapi.PropertySet

  # -------------------------------------------------------------------
  # Public API
  # -------------------------------------------------------------------

  @doc """
  Convert a `%OutlookMsg.Msg{}` to `%OutlookMsg.Mime{}`.

  Returns `{:ok, mime}` on success.

  ## Example

      {:ok, msg} = OutlookMsg.Msg.open("message.msg")
      {:ok, mime} = OutlookMsg.Convert.to_mime(msg)
      email_string = OutlookMsg.Mime.to_string(mime)
  """
  @spec to_mime(Msg.t()) :: {:ok, Mime.t()}
  def to_mime(%Msg{} = msg) do
    mime = %Mime{}

    mime = populate_headers(mime, msg)
    mime = build_body(mime, msg)
    mime = attach_files(mime, msg)

    {:ok, mime}
  end

  # -------------------------------------------------------------------
  # Header population
  # -------------------------------------------------------------------

  @doc false
  @spec populate_headers(Mime.t(), Msg.t()) :: Mime.t()
  def populate_headers(%Mime{} = mime, %Msg{} = msg) do
    props = msg.properties

    mime
    |> set_from_header(props)
    |> set_recipient_headers(msg.recipients)
    |> set_subject_header(props)
    |> set_date_header(props)
    |> set_message_id_header(props)
    |> set_in_reply_to_header(props)
    |> set_references_header(props)
    |> set_mime_version_header()
    |> set_importance_headers(props)
    |> set_sensitivity_header(props)
  end

  # -------------------------------------------------------------------
  # Body building
  # -------------------------------------------------------------------

  @doc false
  @spec build_body(Mime.t(), Msg.t()) :: Mime.t()
  def build_body(%Mime{} = mime, %Msg{} = msg) do
    text_body = PropertySet.body(msg.properties)
    html_body = PropertySet.body_html(msg.properties)

    cond do
      text_body != nil and html_body != nil ->
        text_part = make_text_part(text_body)
        html_part = make_html_part(html_body)
        alternative = wrap_multipart("alternative", [text_part, html_part], "multipart/alternative")

        %{mime | body: nil, parts: alternative.parts, content_type: "multipart/alternative"}
        |> Mime.set_header("Content-Type",
          "multipart/alternative; boundary=\"#{get_boundary_from_parts(alternative)}\"")

      html_body != nil ->
        encoding = pick_transfer_encoding(html_body)
        encoded_body = encode_body(html_body, encoding)

        %{mime | body: encoded_body}
        |> Mime.set_header("Content-Type", "text/html; charset=utf-8")
        |> Mime.set_header("Content-Transfer-Encoding", encoding)

      text_body != nil ->
        encoding = pick_text_transfer_encoding(text_body)
        encoded_body = encode_body(text_body, encoding)

        %{mime | body: encoded_body}
        |> Mime.set_header("Content-Type", "text/plain; charset=utf-8")
        |> Mime.set_header("Content-Transfer-Encoding", encoding)

      true ->
        %{mime | body: ""}
        |> Mime.set_header("Content-Type", "text/plain; charset=utf-8")
        |> Mime.set_header("Content-Transfer-Encoding", "7bit")
    end
  end

  # -------------------------------------------------------------------
  # Attachment handling
  # -------------------------------------------------------------------

  @doc false
  @spec attach_files(Mime.t(), Msg.t()) :: Mime.t()
  def attach_files(%Mime{} = mime, %Msg{attachments: []}) do
    mime
  end

  def attach_files(%Mime{} = mime, %Msg{attachments: nil}) do
    mime
  end

  def attach_files(%Mime{} = mime, %Msg{attachments: attachments}) do
    {inline_atts, regular_atts} = Enum.split_with(attachments, &Attachment.inline?/1)

    # Build the body part - this is whatever we have so far
    body_part = extract_body_part(mime)

    # If there are inline attachments, wrap body + inline in multipart/related
    body_part =
      if inline_atts != [] do
        inline_parts = Enum.map(inline_atts, &make_attachment_part/1)
        wrap_multipart("related", [body_part | inline_parts], "multipart/related")
      else
        body_part
      end

    # Wrap everything in multipart/mixed with regular attachments
    regular_parts = Enum.map(regular_atts, &make_attachment_part/1)
    all_parts = [body_part | regular_parts]

    mixed = wrap_multipart("mixed", all_parts, "multipart/mixed")
    boundary = get_boundary_from_parts(mixed)

    %{mime |
      body: nil,
      parts: mixed.parts,
      content_type: "multipart/mixed"
    }
    |> Mime.set_header("Content-Type", "multipart/mixed; boundary=\"#{boundary}\"")
  end

  # -------------------------------------------------------------------
  # MIME type guessing
  # -------------------------------------------------------------------

  @doc false
  defp guess_mime_type(filename) do
    ext = filename |> Path.extname() |> String.downcase()

    case ext do
      ".txt" -> "text/plain"
      ".html" -> "text/html"
      ".htm" -> "text/html"
      ".pdf" -> "application/pdf"
      ".doc" -> "application/msword"
      ".docx" -> "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ".xls" -> "application/vnd.ms-excel"
      ".xlsx" -> "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ".ppt" -> "application/vnd.ms-powerpoint"
      ".png" -> "image/png"
      ".jpg" -> "image/jpeg"
      ".jpeg" -> "image/jpeg"
      ".gif" -> "image/gif"
      ".zip" -> "application/zip"
      ".xml" -> "application/xml"
      ".csv" -> "text/csv"
      ".msg" -> "application/vnd.ms-outlook"
      _ -> "application/octet-stream"
    end
  end

  # -------------------------------------------------------------------
  # Private helpers -- Header setters
  # -------------------------------------------------------------------

  defp set_from_header(mime, props) do
    sender_name = PropertySet.get(props, :pr_sender_name)
    sender_email = resolve_sender_email(props)

    if sender_email do
      from = Mime.format_email_address(sender_name, sender_email)
      Mime.set_header(mime, "From", from)
    else
      mime
    end
  end

  defp resolve_sender_email(props) do
    addrtype = PropertySet.get(props, :pr_sender_addrtype)
    sender_email = PropertySet.get(props, :pr_sender_email_address)

    cond do
      addrtype != nil and String.upcase(to_string(addrtype)) == "SMTP" ->
        sender_email

      true ->
        # Try SMTP-specific properties as fallback
        PropertySet.get(props, :pr_sent_representing_smtp_address) ||
          PropertySet.get(props, :pr_sender_smtp_address) ||
          sender_email
    end
  end

  defp set_recipient_headers(mime, recipients) when is_list(recipients) do
    grouped = Enum.group_by(recipients, & &1.type)

    mime
    |> set_recipient_header("To", Map.get(grouped, :to, []))
    |> set_recipient_header("Cc", Map.get(grouped, :cc, []))
    |> set_recipient_header("Bcc", Map.get(grouped, :bcc, []))
  end

  defp set_recipient_headers(mime, _recipients), do: mime

  defp set_recipient_header(mime, _name, []), do: mime

  defp set_recipient_header(mime, name, recipients) do
    value =
      recipients
      |> Enum.map(&Recipient.to_string/1)
      |> Enum.join(", ")

    Mime.set_header(mime, name, value)
  end

  defp set_subject_header(mime, props) do
    case PropertySet.get(props, :pr_subject) do
      nil -> mime
      subject -> Mime.set_header(mime, "Subject", Mime.encode_header_value(subject))
    end
  end

  defp set_date_header(mime, props) do
    date =
      PropertySet.get(props, :pr_message_delivery_time) ||
        PropertySet.get(props, :pr_client_submit_time)

    case date do
      nil -> mime
      dt -> Mime.set_header(mime, "Date", Mime.format_date(dt))
    end
  end

  defp set_message_id_header(mime, props) do
    case PropertySet.get(props, :pr_internet_message_id) do
      nil -> mime
      id -> Mime.set_header(mime, "Message-ID", id)
    end
  end

  defp set_in_reply_to_header(mime, props) do
    case PropertySet.get(props, :pr_in_reply_to_id) do
      nil -> mime
      id -> Mime.set_header(mime, "In-Reply-To", id)
    end
  end

  defp set_references_header(mime, props) do
    case PropertySet.get(props, :pr_internet_references) do
      nil -> mime
      refs -> Mime.set_header(mime, "References", refs)
    end
  end

  defp set_mime_version_header(mime) do
    Mime.set_header(mime, "MIME-Version", "1.0")
  end

  defp set_importance_headers(mime, props) do
    importance_val = PropertySet.get(props, :pr_importance)

    mime
    |> maybe_set_importance(importance_val)
    |> maybe_set_x_priority(importance_val)
  end

  defp maybe_set_importance(mime, 0), do: Mime.set_header(mime, "Importance", "low")
  defp maybe_set_importance(mime, 2), do: Mime.set_header(mime, "Importance", "high")
  defp maybe_set_importance(mime, _), do: mime

  defp maybe_set_x_priority(mime, 0), do: Mime.set_header(mime, "X-Priority", "5")
  defp maybe_set_x_priority(mime, 1), do: Mime.set_header(mime, "X-Priority", "3")
  defp maybe_set_x_priority(mime, 2), do: Mime.set_header(mime, "X-Priority", "1")
  defp maybe_set_x_priority(mime, _), do: mime

  defp set_sensitivity_header(mime, props) do
    case PropertySet.get(props, :pr_sensitivity) do
      1 -> Mime.set_header(mime, "Sensitivity", "Personal")
      2 -> Mime.set_header(mime, "Sensitivity", "Private")
      3 -> Mime.set_header(mime, "Sensitivity", "Company-Confidential")
      _ -> mime
    end
  end

  # -------------------------------------------------------------------
  # Private helpers -- Body parts
  # -------------------------------------------------------------------

  defp make_text_part(text) do
    encoding = pick_text_transfer_encoding(text)
    encoded = encode_body(text, encoding)

    %Mime{
      headers: [
        {"Content-Type", "text/plain; charset=utf-8"},
        {"Content-Transfer-Encoding", encoding}
      ],
      body: encoded,
      content_type: "text/plain"
    }
  end

  defp make_html_part(html) do
    encoding = pick_transfer_encoding(html)
    encoded = encode_body(html, encoding)

    %Mime{
      headers: [
        {"Content-Type", "text/html; charset=utf-8"},
        {"Content-Transfer-Encoding", encoding}
      ],
      body: encoded,
      content_type: "text/html"
    }
  end

  defp make_attachment_part(%Attachment{embedded_msg: %Msg{} = embedded} = att) do
    case to_mime(embedded) do
      {:ok, embedded_mime} ->
        embedded_str = Mime.to_string(embedded_mime)
        filename = att.filename || "message.eml"

        %Mime{
          headers: [
            {"Content-Type", "message/rfc822; name=\"#{filename}\""},
            {"Content-Disposition", "attachment; filename=\"#{filename}\""}
          ],
          body: embedded_str,
          content_type: "message/rfc822"
        }
    end
  end

  defp make_attachment_part(%Attachment{} = att) do
    data = att.data || <<>>
    filename = att.filename || "attachment"
    mime_type = att.mime_type || guess_mime_type(filename)

    encoded_data = Mime.encode_base64(data)
    is_inline = Attachment.inline?(att)

    disposition =
      if is_inline do
        "inline; filename=\"#{filename}\""
      else
        "attachment; filename=\"#{filename}\""
      end

    headers = [
      {"Content-Type", "#{mime_type}; name=\"#{filename}\""},
      {"Content-Disposition", disposition},
      {"Content-Transfer-Encoding", "base64"}
    ]

    # Add Content-ID for inline attachments
    content_id = Attachment.content_id(att)

    headers =
      if content_id != nil do
        cid =
          if String.starts_with?(content_id, "<") do
            content_id
          else
            "<#{content_id}>"
          end

        headers ++ [{"Content-ID", cid}]
      else
        headers
      end

    # Add Content-Location for inline attachments
    content_location = Attachment.content_location(att)

    headers =
      if content_location != nil do
        headers ++ [{"Content-Location", content_location}]
      else
        headers
      end

    %Mime{
      headers: headers,
      body: encoded_data,
      content_type: mime_type
    }
  end

  defp wrap_multipart(_subtype, parts, content_type) do
    boundary = Mime.make_boundary()

    %Mime{
      headers: [
        {"Content-Type", "#{content_type}; boundary=\"#{boundary}\""}
      ],
      body: nil,
      parts: parts,
      content_type: content_type,
      preamble: "This is a multi-part message in MIME format."
    }
  end

  # -------------------------------------------------------------------
  # Private helpers -- Encoding
  # -------------------------------------------------------------------

  # Determine if text is pure 7-bit ASCII.
  defp ascii_only?(<<>>), do: true
  defp ascii_only?(<<byte, rest::binary>>) when byte <= 127, do: ascii_only?(rest)
  defp ascii_only?(_), do: false

  # For text/plain: use 7bit if pure ASCII, otherwise quoted-printable.
  defp pick_text_transfer_encoding(text) do
    if ascii_only?(text), do: "7bit", else: "quoted-printable"
  end

  # For text/html and general content: use base64 if contains high bytes,
  # otherwise quoted-printable for safety; 7bit if pure ASCII.
  defp pick_transfer_encoding(content) do
    cond do
      ascii_only?(content) -> "7bit"
      true -> "quoted-printable"
    end
  end

  defp encode_body(content, "base64"), do: Mime.encode_base64(content)
  defp encode_body(content, "quoted-printable"), do: Mime.encode_quoted_printable(content)
  defp encode_body(content, _), do: content

  # -------------------------------------------------------------------
  # Private helpers -- Multipart utilities
  # -------------------------------------------------------------------

  # Extract the current body as a standalone MIME part so it can be
  # wrapped inside a multipart container.
  defp extract_body_part(%Mime{parts: parts, content_type: ct} = mime)
       when is_list(parts) and parts != [] do
    # Already a multipart -- return it as a sub-part
    boundary = extract_boundary(mime)

    %Mime{
      headers: [{"Content-Type", "#{ct}; boundary=\"#{boundary}\""}],
      body: nil,
      parts: parts,
      content_type: ct,
      preamble: mime.preamble
    }
  end

  defp extract_body_part(%Mime{} = mime) do
    ct_header = Mime.get_header(mime, "Content-Type") || "text/plain; charset=utf-8"
    cte_header = Mime.get_header(mime, "Content-Transfer-Encoding")

    headers = [{"Content-Type", ct_header}]

    headers =
      if cte_header do
        headers ++ [{"Content-Transfer-Encoding", cte_header}]
      else
        headers
      end

    {main_ct, _params} = Mime.split_header(ct_header)

    %Mime{
      headers: headers,
      body: mime.body,
      content_type: main_ct
    }
  end

  # Extract the boundary string from a MIME struct's Content-Type header.
  defp extract_boundary(%Mime{} = mime) do
    case Mime.get_header(mime, "Content-Type") do
      nil ->
        Mime.make_boundary()

      ct_value ->
        {_main, params} = Mime.split_header(ct_value)
        Map.get(params, "boundary", Mime.make_boundary())
    end
  end

  # Get the boundary from a freshly-created multipart wrapper.
  defp get_boundary_from_parts(%Mime{} = mime) do
    extract_boundary(mime)
  end
end
