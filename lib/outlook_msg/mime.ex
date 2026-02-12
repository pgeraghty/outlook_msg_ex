defmodule OutlookMsg.Mime do
  @moduledoc """
  Basic MIME/RFC2822 message builder and parser.

  Provides a struct representing a MIME message with headers, body, and
  multipart children, along with functions for parsing, serializing, and
  manipulating MIME messages.

  Based on mime.rb from ruby-msg.
  """
  alias OutlookMsg.Warning

  # -------------------------------------------------------------------
  # Struct
  # -------------------------------------------------------------------

  defstruct headers: [],
            body: nil,
            parts: [],
            content_type: nil,
            preamble: nil,
            epilogue: nil,
            warnings: []

  @type t :: %__MODULE__{
          headers: [{String.t(), String.t()}],
          body: String.t() | nil,
          parts: [t()],
          content_type: String.t() | nil,
          preamble: String.t() | nil,
          epilogue: String.t() | nil,
          warnings: [Warning.t() | String.t()]
        }

  # Maximum line length before folding (RFC2822 recommends 78, we use 76).
  @max_line_length 76

  # -------------------------------------------------------------------
  # Construction / Parsing
  # -------------------------------------------------------------------

  @doc """
  Creates an empty MIME message.
  """
  @spec new() :: t()
  def new, do: %__MODULE__{}

  @doc """
  Parses a MIME string into a `%Mime{}` struct.

  Splits headers from body at the first blank line, parses headers and
  Content-Type, and recursively parses multipart children when applicable.
  """
  @spec new(String.t()) :: t()
  def new(raw) when is_binary(raw) do
    {header_text, body_text} = split_header_body(raw)
    {headers, header_warnings} = parse_headers(header_text)

    content_type_raw = header_value(headers, "content-type")
    {ct_main, ct_params} = if content_type_raw, do: split_header(content_type_raw), else: {nil, %{}}

    mime = %__MODULE__{
      headers: headers,
      content_type: ct_main,
      warnings: header_warnings
    }

    if ct_main != nil and String.starts_with?(ct_main, "multipart/") do
      boundary = Map.get(ct_params, "boundary", "")
      parse_multipart(mime, body_text, boundary)
    else
      %{mime | body: body_text}
    end
  end

  # -------------------------------------------------------------------
  # Predicates
  # -------------------------------------------------------------------

  @doc """
  Returns `true` if the MIME message has a multipart content type.
  """
  @spec multipart?(t()) :: boolean()
  def multipart?(%__MODULE__{content_type: ct}) when is_binary(ct) do
    String.starts_with?(ct, "multipart/")
  end

  def multipart?(_), do: false

  # -------------------------------------------------------------------
  # Header accessors
  # -------------------------------------------------------------------

  @doc """
  Returns the first header value matching `name` (case-insensitive),
  or `nil` if not found.
  """
  @spec get_header(t(), String.t()) :: String.t() | nil
  def get_header(%__MODULE__{headers: headers}, name) do
    header_value(headers, name)
  end

  @doc """
  Sets (replaces) a header. If a header with the same name (case-insensitive)
  already exists, replaces its first occurrence; otherwise appends the header.
  """
  @spec set_header(t(), String.t(), String.t()) :: t()
  def set_header(%__MODULE__{headers: headers} = mime, name, value) do
    name_down = String.downcase(name)

    case Enum.find_index(headers, fn {n, _} -> String.downcase(n) == name_down end) do
      nil -> %{mime | headers: headers ++ [{name, value}]}
      idx -> %{mime | headers: List.replace_at(headers, idx, {name, value})}
    end
  end

  @doc """
  Adds a header, allowing duplicates.
  """
  @spec add_header(t(), String.t(), String.t()) :: t()
  def add_header(%__MODULE__{headers: headers} = mime, name, value) do
    %{mime | headers: headers ++ [{name, value}]}
  end

  # -------------------------------------------------------------------
  # Header value parsing
  # -------------------------------------------------------------------

  @doc """
  Parses a structured header value into its main value and a map of parameters.

  ## Example

      iex> OutlookMsg.Mime.split_header(~s(multipart/mixed; boundary="abc123"))
      {"multipart/mixed", %{"boundary" => "abc123"}}

      iex> OutlookMsg.Mime.split_header("text/plain; charset=utf-8")
      {"text/plain", %{"charset" => "utf-8"}}
  """
  @spec split_header(String.t()) :: {String.t(), map()}
  def split_header(value) when is_binary(value) do
    [main | params_raw] =
      value
      |> String.split(";")
      |> Enum.map(&String.trim/1)

    params =
      params_raw
      |> Enum.reject(&(&1 == ""))
      |> Enum.reduce(%{}, fn param, acc ->
        case String.split(param, "=", parts: 2) do
          [key, val] ->
            Map.put(acc, String.downcase(String.trim(key)), unquote_value(String.trim(val)))

          _ ->
            acc
        end
      end)

    {String.downcase(String.trim(main)), params}
  end

  # -------------------------------------------------------------------
  # Serialization
  # -------------------------------------------------------------------

  @doc """
  Serializes a `%Mime{}` struct to an RFC2822-formatted string.
  """
  @spec to_string(t()) :: String.t()
  def to_string(%__MODULE__{} = mime) do
    header_str = format_headers(mime.headers)

    if multipart?(mime) do
      {_ct, params} =
        case header_value(mime.headers, "content-type") do
          nil -> {mime.content_type, %{}}
          raw -> split_header(raw)
        end

      boundary = Map.get(params, "boundary", make_boundary())

      parts_str =
        mime.parts
        |> Enum.map(fn part ->
          "--#{boundary}\r\n" <> __MODULE__.to_string(part)
        end)
        |> Enum.join("\r\n")

      preamble = if mime.preamble, do: mime.preamble <> "\r\n", else: ""
      epilogue = if mime.epilogue, do: mime.epilogue, else: ""

      header_str <> "\r\n" <> preamble <> parts_str <> "\r\n--#{boundary}--\r\n" <> epilogue
    else
      body = mime.body || ""
      header_str <> "\r\n" <> body
    end
  end

  # -------------------------------------------------------------------
  # Encoding helpers
  # -------------------------------------------------------------------

  @doc """
  Generates a unique MIME boundary string.
  """
  @spec make_boundary() :: String.t()
  def make_boundary do
    "----=_Part_#{:erlang.unique_integer([:positive])}_#{System.system_time(:millisecond)}"
  end

  @doc """
  Base64-encodes binary data with line wrapping at 76 characters.
  """
  @spec encode_base64(binary()) :: String.t()
  def encode_base64(data) do
    data
    |> Base.encode64()
    |> chunk_string(@max_line_length)
    |> Enum.join("\r\n")
  end

  @doc """
  Quoted-printable encodes a string.

  Non-printable or non-ASCII bytes are encoded as `=XX` where XX is the
  uppercase hex representation. Lines are soft-wrapped at 76 characters
  using `=\\r\\n` continuation.
  """
  @spec encode_quoted_printable(binary()) :: String.t()
  def encode_quoted_printable(data) when is_binary(data) do
    data
    |> :binary.bin_to_list()
    |> Enum.reduce({[], 0}, fn byte, {acc, line_len} ->
      encoded =
        cond do
          byte == ?\t or byte == ?\s ->
            <<byte>>

          byte == ?\r or byte == ?\n ->
            <<byte>>

          byte >= 33 and byte <= 126 and byte != ?= ->
            <<byte>>

          true ->
            "=" <> String.upcase(Base.encode16(<<byte>>))
        end

      enc_len = byte_size(encoded)

      if line_len + enc_len >= 75 and byte != ?\r and byte != ?\n do
        {[acc, "=\r\n", encoded], enc_len}
      else
        new_line_len =
          if byte == ?\n, do: 0, else: line_len + enc_len

        {[acc, encoded], new_line_len}
      end
    end)
    |> elem(0)
    |> IO.iodata_to_binary()
  end

  @doc """
  RFC2047-encodes a header value if it contains non-ASCII characters.

  Returns the original value unchanged if it is pure ASCII.
  """
  @spec encode_header_value(String.t()) :: String.t()
  def encode_header_value(value) when is_binary(value) do
    if ascii?(value) do
      value
    else
      encoded = Base.encode64(value)
      "=?UTF-8?B?#{encoded}?="
    end
  end

  @doc """
  Formats an email address with an optional display name.

  ## Examples

      iex> OutlookMsg.Mime.format_email_address("John Doe", "john@example.com")
      ~s("John Doe" <john@example.com>)

      iex> OutlookMsg.Mime.format_email_address(nil, "john@example.com")
      "<john@example.com>"
  """
  @spec format_email_address(String.t() | nil, String.t()) :: String.t()
  def format_email_address(nil, email), do: "<#{email}>"
  def format_email_address("", email), do: "<#{email}>"
  def format_email_address(name, email), do: ~s("#{name}" <#{email}>)

  @doc """
  Formats a `DateTime` (or `NaiveDateTime`) as an RFC2822 date string.

  ## Example

      iex> OutlookMsg.Mime.format_date(~U[2024-01-01 12:00:00Z])
      "Mon, 01 Jan 2024 12:00:00 +0000"
  """
  @spec format_date(DateTime.t() | NaiveDateTime.t()) :: String.t()
  def format_date(%DateTime{} = dt) do
    day_name = day_of_week_name(Date.day_of_week(dt))
    month_name = month_name(dt.month)

    offset = format_utc_offset(dt.utc_offset + (dt.std_offset || 0))

    "#{day_name}, #{pad2(dt.day)} #{month_name} #{dt.year} " <>
      "#{pad2(dt.hour)}:#{pad2(dt.minute)}:#{pad2(dt.second)} #{offset}"
  end

  def format_date(%NaiveDateTime{} = ndt) do
    day_name = day_of_week_name(Date.day_of_week(ndt))
    month_name = month_name(ndt.month)

    "#{day_name}, #{pad2(ndt.day)} #{month_name} #{ndt.year} " <>
      "#{pad2(ndt.hour)}:#{pad2(ndt.minute)}:#{pad2(ndt.second)} +0000"
  end

  # ===================================================================
  # Private helpers
  # ===================================================================

  # -------------------------------------------------------------------
  # Parsing helpers
  # -------------------------------------------------------------------

  # Splits a raw message into header text and body text at the first blank line.
  defp split_header_body(raw) do
    cond do
      # Try \r\n\r\n first (standard RFC2822)
      String.contains?(raw, "\r\n\r\n") ->
        [header, body] = String.split(raw, "\r\n\r\n", parts: 2)
        {header, body}

      # Fall back to \n\n (common in unix-style messages)
      String.contains?(raw, "\n\n") ->
        [header, body] = String.split(raw, "\n\n", parts: 2)
        {header, body}

      # No blank line found -- treat entire message as headers
      true ->
        {raw, ""}
    end
  end

  # Parses raw header text into a list of {name, value} tuples.
  # Handles continuation lines (lines starting with whitespace).
  defp parse_headers(header_text) do
    lines =
      header_text
      |> String.split(~r/\r?\n/)

    {headers, warnings} =
      lines
      |> Enum.reduce({[], []}, fn line, {acc, warnings} ->
      cond do
        # Continuation line: starts with whitespace
        String.match?(line, ~r/^[ \t]/) and acc != [] ->
          {name, prev_value} = List.last(acc)
          updated = {name, prev_value <> " " <> String.trim(line)}
          {List.replace_at(acc, length(acc) - 1, updated), warnings}

        # New header line
        String.contains?(line, ":") ->
          [name | rest] = String.split(line, ":", parts: 2)
          value = rest |> Enum.join(":") |> String.trim()
          {acc ++ [{String.trim(name), value}], warnings}

        # Skip malformed lines
        String.trim(line) == "" ->
          {acc, warnings}

        true ->
          {acc, warnings ++ [Warning.new(:malformed_header_line, "malformed header line ignored", context: line)]}
      end
    end)

    {
      Enum.map(headers, fn {name, value} -> {name, decode_header_value(value)} end),
      warnings
    }
  end

  # Look up a header value by name (case-insensitive) in a headers list.
  defp header_value(headers, name) do
    name_down = String.downcase(name)

    Enum.find_value(headers, fn {n, v} ->
      if String.downcase(n) == name_down, do: v
    end)
  end

  # -------------------------------------------------------------------
  # Multipart parsing
  # -------------------------------------------------------------------

  # Parses the body of a multipart message, splitting on the boundary and
  # recursively parsing each part.
  defp parse_multipart(mime, body_text, boundary) do
    if boundary == "" do
      warn = Warning.new(:multipart_missing_boundary, "multipart content-type missing boundary; treated as single-part body")
      %{mime | body: body_text, warnings: mime.warnings ++ [warn]}
    else
      delimiter = "--" <> boundary
      terminator = "--" <> boundary <> "--"

    # Split on the terminator first to separate epilogue
    {before_term, epilogue} =
      case String.split(body_text, terminator, parts: 2) do
        [before, rest_after] -> {before, String.trim_leading(rest_after, "\r\n")}
        [only] -> {only, nil}
      end

    # Split the remaining text on the delimiter
    raw_parts = String.split(before_term, delimiter)

    # The first segment before the first delimiter is the preamble
    {preamble, part_texts} =
      case raw_parts do
        [] -> {nil, []}
        [pre | rest] -> {normalize_preamble(pre), rest}
      end

    # Each part text has a leading \r\n (or \n) that separates it from
    # the boundary line. Strip that and parse recursively.
    parts =
      part_texts
      |> Enum.map(&strip_leading_newline/1)
      |> Enum.reject(&(&1 == ""))
      |> Enum.map(&new/1)

    part_warnings =
      parts
      |> Enum.with_index()
      |> Enum.flat_map(fn {part, idx} ->
        Enum.map(part.warnings || [], fn warning ->
          case warning do
            %Warning{} = w -> %{w | context: "part[#{idx}]" <> if(w.context, do: " #{w.context}", else: "")}
            txt -> Warning.new(:nested_part_warning, Kernel.to_string(txt), context: "part[#{idx}]")
          end
        end)
      end)

      %{mime | parts: parts, preamble: preamble, epilogue: epilogue, warnings: mime.warnings ++ part_warnings}
    end
  end

  defp normalize_preamble(""), do: nil
  defp normalize_preamble(text) do
    trimmed = String.trim(text)
    if trimmed == "", do: nil, else: trimmed
  end

  defp strip_leading_newline("\r\n" <> rest), do: rest
  defp strip_leading_newline("\n" <> rest), do: rest
  defp strip_leading_newline(other), do: other

  # -------------------------------------------------------------------
  # Serialization helpers
  # -------------------------------------------------------------------

  # Formats a list of headers into an RFC2822 header block string.
  defp format_headers(headers) do
    headers
    |> Enum.map(fn {name, value} ->
      fold_header(name, value)
    end)
    |> Enum.join("")
  end

  # Folds a single header line if it exceeds the maximum line length.
  # Returns the header with trailing \r\n.
  defp fold_header(name, value) do
    line = "#{name}: #{value}"

    if String.length(line) <= @max_line_length do
      line <> "\r\n"
    else
      fold_long_line(line) <> "\r\n"
    end
  end

  # Folds a long line by inserting \r\n followed by a space at appropriate
  # break points (preferring spaces and commas).
  defp fold_long_line(line) do
    do_fold_line(line, @max_line_length, [])
    |> Enum.reverse()
    |> Enum.join("\r\n ")
  end

  defp do_fold_line("", _max, acc), do: acc

  defp do_fold_line(line, max, acc) do
    if String.length(line) <= max do
      [line | acc]
    else
      # Try to find a break point (space) before the max length
      chunk = String.slice(line, 0, max)
      break_pos = find_break_point(chunk)

      if break_pos > 0 do
        left = String.slice(line, 0, break_pos)
        right = String.slice(line, break_pos..-1//1) |> String.trim_leading()
        do_fold_line(right, max, [left | acc])
      else
        # No good break point found; force break at max
        left = String.slice(line, 0, max)
        right = String.slice(line, max..-1//1)
        do_fold_line(right, max, [left | acc])
      end
    end
  end

  # Finds the last space or comma position suitable for folding.
  defp find_break_point(chunk) do
    chars = String.graphemes(chunk)

    chars
    |> Enum.with_index()
    |> Enum.reverse()
    |> Enum.find_value(0, fn
      {" ", idx} -> idx
      {",", idx} -> idx + 1
      _ -> nil
    end)
  end

  # -------------------------------------------------------------------
  # String helpers
  # -------------------------------------------------------------------

  # Strips surrounding double quotes from a value.
  defp unquote_value(val) do
    val
    |> String.trim_leading("\"")
    |> String.trim_trailing("\"")
  end

  # Decodes RFC2047 encoded words in headers, leaving undecodable chunks unchanged.
  defp decode_header_value(value) when is_binary(value) do
    value = collapse_adjacent_encoded_word_whitespace(value)

    Regex.replace(~r/=\?([^?]+)\?([bBqQ])\?([^?]*)\?=/, value, fn match, charset, encoding, payload ->
      case decode_encoded_word(charset, encoding, payload) do
        {:ok, decoded} -> decoded
        _ -> match
      end
    end)
  end

  # RFC2047 allows linear whitespace between adjacent encoded words; it should be ignored.
  defp collapse_adjacent_encoded_word_whitespace(value) do
    pattern = ~r/(=\?[^?]+\?[bBqQ]\?[^?]*\?=)\s+(?==\?[^?]+\?[bBqQ]\?[^?]*\?=)/
    collapse_adjacent_encoded_word_whitespace(value, pattern)
  end

  defp collapse_adjacent_encoded_word_whitespace(value, pattern) do
    updated = Regex.replace(pattern, value, "\\1")
    if updated == value, do: value, else: collapse_adjacent_encoded_word_whitespace(updated, pattern)
  end

  defp decode_encoded_word(charset, encoding, payload) do
    with {:ok, bytes} <- decode_encoded_payload(encoding, payload),
         {:ok, decoded} <- decode_charset(bytes, charset) do
      {:ok, decoded}
    else
      _ -> :error
    end
  end

  defp decode_encoded_payload(enc, payload) when enc in ["b", "B"] do
    case Base.decode64(payload) do
      {:ok, bin} -> {:ok, bin}
      :error -> :error
    end
  end

  defp decode_encoded_payload(enc, payload) when enc in ["q", "Q"] do
    payload
    |> String.replace("_", " ")
    |> decode_q_encoded()
  end

  defp decode_charset(bytes, charset) do
    case String.downcase(String.trim(charset)) do
      "utf-8" -> {:ok, bytes}
      "us-ascii" -> {:ok, bytes}
      "iso-8859-1" -> {:ok, :unicode.characters_to_binary(bytes, :latin1, :utf8)}
      "latin1" -> {:ok, :unicode.characters_to_binary(bytes, :latin1, :utf8)}
      _ -> {:ok, bytes}
    end
  rescue
    _ -> :error
  end

  defp decode_q_encoded(data) when is_binary(data) do
    do_decode_q_encoded(data, [])
  end

  defp do_decode_q_encoded(<<>>, acc), do: {:ok, IO.iodata_to_binary(Enum.reverse(acc))}

  defp do_decode_q_encoded(<<"=", h1, h2, rest::binary>>, acc) do
    case Integer.parse(<<h1, h2>>, 16) do
      {byte, ""} -> do_decode_q_encoded(rest, [<<byte>> | acc])
      _ -> do_decode_q_encoded(rest, [<<?=, h1, h2>> | acc])
    end
  end

  defp do_decode_q_encoded(<<char, rest::binary>>, acc) do
    do_decode_q_encoded(rest, [<<char>> | acc])
  end

  # Splits a string into chunks of at most `n` characters.
  defp chunk_string(string, n) when is_binary(string) and n > 0 do
    do_chunk_string(string, n, [])
  end

  defp do_chunk_string("", _n, acc), do: Enum.reverse(acc)

  defp do_chunk_string(string, n, acc) do
    case String.split_at(string, n) do
      {chunk, ""} -> Enum.reverse([chunk | acc])
      {chunk, rest} -> do_chunk_string(rest, n, [chunk | acc])
    end
  end

  # Returns `true` if a binary contains only ASCII characters (bytes 0-127).
  defp ascii?(<<>>), do: true
  defp ascii?(<<byte, rest::binary>>) when byte <= 127, do: ascii?(rest)
  defp ascii?(_), do: false

  # -------------------------------------------------------------------
  # Date formatting helpers
  # -------------------------------------------------------------------

  defp day_of_week_name(1), do: "Mon"
  defp day_of_week_name(2), do: "Tue"
  defp day_of_week_name(3), do: "Wed"
  defp day_of_week_name(4), do: "Thu"
  defp day_of_week_name(5), do: "Fri"
  defp day_of_week_name(6), do: "Sat"
  defp day_of_week_name(7), do: "Sun"

  defp month_name(1), do: "Jan"
  defp month_name(2), do: "Feb"
  defp month_name(3), do: "Mar"
  defp month_name(4), do: "Apr"
  defp month_name(5), do: "May"
  defp month_name(6), do: "Jun"
  defp month_name(7), do: "Jul"
  defp month_name(8), do: "Aug"
  defp month_name(9), do: "Sep"
  defp month_name(10), do: "Oct"
  defp month_name(11), do: "Nov"
  defp month_name(12), do: "Dec"

  defp pad2(n) when n < 10, do: "0#{n}"
  defp pad2(n), do: "#{n}"

  # Formats a UTC offset in seconds as "+HHMM" or "-HHMM".
  defp format_utc_offset(offset_seconds) do
    sign = if offset_seconds >= 0, do: "+", else: "-"
    total = abs(offset_seconds)
    hours = div(total, 3600)
    minutes = div(rem(total, 3600), 60)
    "#{sign}#{pad2(hours)}#{pad2(minutes)}"
  end
end

# -------------------------------------------------------------------
# String.Chars protocol implementation
# -------------------------------------------------------------------

defimpl String.Chars, for: OutlookMsg.Mime do
  def to_string(%OutlookMsg.Mime{} = mime) do
    OutlookMsg.Mime.to_string(mime)
  end
end
