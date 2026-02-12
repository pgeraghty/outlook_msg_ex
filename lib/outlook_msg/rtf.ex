defmodule OutlookMsg.Rtf do
  @moduledoc """
  RTF decompression (LZFu algorithm) and RTF-to-HTML/text extraction.

  Outlook MSG files store RTF body content compressed with the LZFu algorithm.
  When the RTF originates from HTML (indicated by the `\\fromhtml` marker),
  the original HTML can be recovered from `\\htmltag` groups embedded in the RTF.

  Based on rtf.rb from ruby-msg.
  """

  import Bitwise

  # -------------------------------------------------------------------
  # Constants
  # -------------------------------------------------------------------

  # The 207-byte pre-buffer used to seed the LZFu circular buffer.
  @rtf_prebuf "{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}" <>
                "{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript " <>
                "\\fdecor MS Sans SerifSymbolArialTimes New Roman" <>
                "Courier{\\colortbl\\red0\\green0\\blue0\r\n\\par " <>
                "\\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx"

  @prebuf_len byte_size(@rtf_prebuf)

  # Magic values for the compression header.
  @magic_compressed 0x75465A4C
  @magic_uncompressed 0x414C454D

  # Circular buffer size.
  @buf_size 4096

  # -------------------------------------------------------------------
  # LZFu Decompression
  # -------------------------------------------------------------------

  @doc """
  Decompresses LZFu-compressed RTF data.

  Takes the raw binary data (starting from the compression header) and returns
  `{:ok, decompressed_rtf}` on success or `{:error, reason}` on failure.

  The 16-byte header layout is:

      bytes 0-3:   compressed size (little-endian uint32)
      bytes 4-7:   raw (uncompressed) size (little-endian uint32)
      bytes 8-11:  magic signature (little-endian uint32)
      bytes 12-15: CRC32 (little-endian uint32, not validated)
  """
  @spec decompress(binary()) :: {:ok, binary()} | {:error, atom()}
  def decompress(
        <<_compr_size::32-little, raw_size::32-little, magic::32-little,
          _crc::32-little, rest::binary>>
      ) do
    case magic do
      @magic_uncompressed ->
        {:ok, binary_part(rest, 0, min(raw_size, byte_size(rest)))}

      @magic_compressed ->
        buf = init_buffer()
        {:ok, lzfu_decompress(rest, buf, @prebuf_len, raw_size, [])}

      _ ->
        {:error, :invalid_magic}
    end
  end

  def decompress(_data), do: {:error, :invalid_header}

  # Initialize circular buffer with the RTF pre-buffer content.
  defp init_buffer do
    prebuf_bytes = :binary.bin_to_list(@rtf_prebuf)

    prebuf_bytes
    |> Enum.with_index()
    |> Enum.reduce(%{}, fn {byte, idx}, buf -> Map.put(buf, idx, byte) end)
  end

  # Core LZFu decompression loop.
  # Reads flag bytes, then processes 8 tokens per flag byte.
  defp lzfu_decompress(<<>>, _buf, _wp, _raw_size, acc) do
    finalize_output(acc)
  end

  defp lzfu_decompress(data, buf, wp, raw_size, acc) do
    case data do
      <<flag_byte, rest::binary>> ->
        process_flag_bits(rest, flag_byte, 0, buf, wp, raw_size, acc)

      <<>> ->
        finalize_output(acc)
    end
  end

  # Process 8 bits of a flag byte (LSB first).
  defp process_flag_bits(data, _flag, 8, buf, wp, raw_size, acc) do
    lzfu_decompress(data, buf, wp, raw_size, acc)
  end

  defp process_flag_bits(<<>>, _flag, _bit, _buf, _wp, raw_size, acc) do
    finalize_output(acc, raw_size)
  end

  defp process_flag_bits(data, flag, bit, buf, wp, raw_size, acc) do
    if (flag >>> bit &&& 1) == 1 do
      # Back-reference: read 2 bytes big-endian
      case data do
        <<val::16-big, rest::binary>> ->
          offset = val >>> 4
          length = (val &&& 0x0F) + 2

          # End marker: offset equals current write position
          if offset == wp do
            finalize_output(acc, raw_size)
          else
            {new_acc, new_buf, new_wp} =
              copy_reference(buf, wp, offset, length, acc)

            process_flag_bits(rest, flag, bit + 1, new_buf, new_wp, raw_size, new_acc)
          end

        _ ->
          # Not enough data for a back-reference
          finalize_output(acc, raw_size)
      end
    else
      # Literal byte
      case data do
        <<byte, rest::binary>> ->
          new_buf = Map.put(buf, wp &&& (@buf_size - 1), byte)
          new_wp = wp + 1
          new_acc = [acc | [byte]]

          process_flag_bits(rest, flag, bit + 1, new_buf, new_wp, raw_size, new_acc)

        <<>> ->
          finalize_output(acc, raw_size)
      end
    end
  end

  # Copy `length` bytes from circular buffer at `offset` to output.
  defp copy_reference(buf, wp, offset, length, acc) do
    Enum.reduce(0..(length - 1), {acc, buf, wp}, fn i, {cur_acc, cur_buf, cur_wp} ->
      pos = (offset + i) &&& (@buf_size - 1)
      byte = Map.get(cur_buf, pos, 0)
      new_buf = Map.put(cur_buf, cur_wp &&& (@buf_size - 1), byte)
      {[cur_acc | [byte]], new_buf, cur_wp + 1}
    end)
  end

  defp finalize_output(acc) do
    IO.iodata_to_binary(acc)
  end

  defp finalize_output(acc, raw_size) do
    result = IO.iodata_to_binary(acc)

    if byte_size(result) > raw_size do
      binary_part(result, 0, raw_size)
    else
      result
    end
  end

  # -------------------------------------------------------------------
  # RTF to HTML extraction
  # -------------------------------------------------------------------

  @doc """
  Extracts HTML from a decompressed RTF string that originated from HTML.

  Returns `{:ok, html_string}` if the RTF contains the `\\fromhtml` marker
  and HTML content could be extracted, or `:none` if the RTF does not
  contain embedded HTML.
  """
  @spec rtf_to_html(binary()) :: {:ok, binary()} | :none
  def rtf_to_html(rtf) when is_binary(rtf) do
    if !String.contains?(rtf, "\\fromhtml") do
      :none
    else
      # Start scanning from the earliest htmltag/mhtmltag marker.
      start =
        case {:binary.match(rtf, "{\\*\\htmltag"), :binary.match(rtf, "{\\*\\mhtmltag")} do
          {:nomatch, :nomatch} -> nil
          {{off, _}, :nomatch} -> off
          {:nomatch, {off, _}} -> off
          {{off1, _}, {off2, _}} -> min(off1, off2)
        end

      case start do
        nil ->
          :none

        offset ->
          stream = binary_part(rtf, offset, byte_size(rtf) - offset)
          html = scan_html_stream(stream, [], nil)
          html = IO.iodata_to_binary(Enum.reverse(html))

          if String.trim(html) == "" do
            :none
          else
            {:ok, html}
          end
      end
    end
  end

  # Scanner equivalent to ruby-msg RTF.rtf2html.
  # Keeps plain text content and selected escapes while skipping RTF control tags.
  defp scan_html_stream(<<>>, acc, _ignore_tag), do: acc

  defp scan_html_stream(data, acc, ignore_tag) do
    cond do
      String.starts_with?(data, "{") ->
        <<_::binary-size(1), rest::binary>> = data
        scan_html_stream(rest, acc, ignore_tag)

      String.starts_with?(data, "}") ->
        <<_::binary-size(1), rest::binary>> = data
        scan_html_stream(rest, acc, ignore_tag)

      true ->
        parse_html_token(data, acc, ignore_tag)
    end
  end

  defp parse_html_token(data, acc, ignore_tag) do
    case Regex.run(~r/\A\\\*\\htmltag(\d+) ?/, data) do
      [full, tag] ->
        rest = binary_part(data, byte_size(full), byte_size(data) - byte_size(full))

        if ignore_tag == tag do
          case :binary.match(rest, "}") do
            :nomatch ->
              scan_html_stream(<<>>, acc, nil)

            {idx, _} ->
              rest2 = binary_part(rest, idx + 1, byte_size(rest) - idx - 1)
              scan_html_stream(rest2, acc, nil)
          end
        else
          scan_html_stream(rest, acc, ignore_tag)
        end

      _ ->
        parse_html_token_fallback(data, acc, ignore_tag)
    end
  end

  defp parse_html_token_fallback(data, acc, ignore_tag) do
    cond do
      (m = Regex.run(~r/\A\\\*\\mhtmltag(\d+) ?/, data)) ->
        [full, tag] = m
        rest = binary_part(data, byte_size(full), byte_size(data) - byte_size(full))
        scan_html_stream(rest, acc, tag)

      (m = Regex.run(~r/\A\\par ?/, data)) ->
        [full] = m
        rest = binary_part(data, byte_size(full), byte_size(data) - byte_size(full))
        scan_html_stream(rest, ["\r\n" | acc], ignore_tag)

      (m = Regex.run(~r/\A\\tab ?/, data)) ->
        [full] = m
        rest = binary_part(data, byte_size(full), byte_size(data) - byte_size(full))
        scan_html_stream(rest, ["\t" | acc], ignore_tag)

      (m = Regex.run(~r/\A\\'([0-9A-Fa-f]{2})/, data)) ->
        [full, hex] = m
        rest = binary_part(data, byte_size(full), byte_size(data) - byte_size(full))
        scan_html_stream(rest, [<<String.to_integer(hex, 16)>> | acc], ignore_tag)

      (m = Regex.run(~r/\A\\pntext/, data)) ->
        [full] = m
        rest = binary_part(data, byte_size(full), byte_size(data) - byte_size(full))

        case :binary.match(rest, "}") do
          :nomatch ->
            scan_html_stream(<<>>, acc, ignore_tag)

          {idx, _} ->
            rest2 = binary_part(rest, idx + 1, byte_size(rest) - idx - 1)
            scan_html_stream(rest2, acc, ignore_tag)
        end

      (m = Regex.run(~r/\A\\htmlrtf0 ?/, data)) ->
        [full] = m
        rest = binary_part(data, byte_size(full), byte_size(data) - byte_size(full))
        scan_html_stream(rest, acc, ignore_tag)

      (m = Regex.run(~r/\A\\htmlrtf/, data)) ->
        [full] = m
        rest = binary_part(data, byte_size(full), byte_size(data) - byte_size(full))

        case Regex.run(~r/\\htmlrtf0 ?/, rest, return: :index) do
          nil ->
            scan_html_stream(<<>>, acc, ignore_tag)

          [{idx, len}] ->
            rest2 = binary_part(rest, idx + len, byte_size(rest) - idx - len)
            scan_html_stream(rest2, acc, ignore_tag)
        end

      (m = Regex.run(~r/\A\\[a-z-]+(\d+)? ?/, data)) ->
        [full | _] = m
        rest = binary_part(data, byte_size(full), byte_size(data) - byte_size(full))
        scan_html_stream(rest, acc, ignore_tag)

      (m = Regex.run(~r/\A[\r\n]/, data)) ->
        [full] = m
        rest = binary_part(data, byte_size(full), byte_size(data) - byte_size(full))
        scan_html_stream(rest, acc, ignore_tag)

      (m = Regex.run(~r/\A\\([{}\\])/, data)) ->
        [full, escaped] = m
        rest = binary_part(data, byte_size(full), byte_size(data) - byte_size(full))
        scan_html_stream(rest, [escaped | acc], ignore_tag)

      true ->
        case data do
          <<byte, rest::binary>> -> scan_html_stream(rest, [<<byte>> | acc], ignore_tag)
          <<>> -> acc
        end
    end
  end

  # Replace \'XX hex escape sequences with the corresponding byte value.
  defp replace_hex_escapes(str) do
    Regex.replace(~r/\\'([0-9a-fA-F]{2})/, str, fn _, hex ->
      <<String.to_integer(hex, 16)>>
    end)
  end

  # -------------------------------------------------------------------
  # RTF to plain text extraction
  # -------------------------------------------------------------------

  @doc """
  Extracts plain text from a decompressed RTF string.

  Strips RTF formatting and control words, returning a plain text
  representation of the content.
  """
  @spec rtf_to_text(binary()) :: binary()
  def rtf_to_text(rtf) when is_binary(rtf) do
    rtf
    # Remove groups that start with {\* (special destinations).
    |> remove_special_groups()
    # Process known control words using regex to avoid partial matches
    # (e.g., \par must not match inside \pard).
    |> replace_control_word("par", "\n")
    |> replace_control_word("line", "\n")
    |> replace_control_word("tab", "\t")
    |> replace_control_word("endash", "-")
    |> replace_control_word("emdash", "--")
    |> replace_control_word("bullet", "*")
    |> replace_control_word("lquote", "'")
    |> replace_control_word("rquote", "'")
    |> replace_control_word("ldblquote", "\"")
    |> replace_control_word("rdblquote", "\"")
    # Replace escaped braces with placeholders before stripping structural braces.
    |> String.replace("\\{", "\x00LBRACE\x00")
    |> String.replace("\\}", "\x00RBRACE\x00")
    |> String.replace("\\\\", "\x00BSLASH\x00")
    # Decode hex escapes.
    |> replace_hex_escapes()
    # Remove all remaining \controlword sequences.
    |> remove_control_words()
    # Remove structural braces.
    |> String.replace("{", "")
    |> String.replace("}", "")
    # Restore escaped characters from placeholders.
    |> String.replace("\x00LBRACE\x00", "{")
    |> String.replace("\x00RBRACE\x00", "}")
    |> String.replace("\x00BSLASH\x00", "\\")
    # Trim leading/trailing whitespace.
    |> String.trim()
  end

  # Replace an RTF control word with its text equivalent, using a regex that
  # ensures the control word is not a prefix of a longer word. RTF control
  # words are terminated by a space (consumed), a non-alpha character, or end
  # of string.
  defp replace_control_word(str, word, replacement) do
    Regex.replace(~r/\\#{word}(?:\r?\n| |(?=[^a-z])|$)/, str, replacement)
  end

  # Remove {\*\destination ...} groups, handling nested braces.
  # Uses a simple iterative approach to handle nesting.
  defp remove_special_groups(str) do
    do_remove_special_groups(str)
  end

  defp do_remove_special_groups(str) do
    # Match the outermost {\* ... } groups.
    # We need to handle nesting, so we use a manual approach.
    case find_special_group(str) do
      nil ->
        str

      {start_pos, end_pos} ->
        prefix = binary_part(str, 0, start_pos)
        suffix_start = end_pos + 1
        suffix = binary_part(str, suffix_start, byte_size(str) - suffix_start)
        do_remove_special_groups(prefix <> suffix)
    end
  end

  # Find the position of the first {\* group and its matching closing brace.
  defp find_special_group(str) do
    case :binary.match(str, "{\\*") do
      {start, _len} ->
        # Find the matching closing brace, accounting for nesting.
        case find_matching_brace(str, start + 1) do
          nil -> nil
          end_pos -> {start, end_pos}
        end

      :nomatch ->
        nil
    end
  end

  # Find the matching closing brace for an opening brace.
  # `pos` should point to the character after the opening brace.
  defp find_matching_brace(str, pos) do
    find_matching_brace(str, pos, 1)
  end

  defp find_matching_brace(_str, pos, 0), do: pos - 1

  defp find_matching_brace(str, pos, _depth) when pos >= byte_size(str), do: nil

  defp find_matching_brace(str, pos, depth) do
    case :binary.at(str, pos) do
      ?{ -> find_matching_brace(str, pos + 1, depth + 1)
      ?} -> find_matching_brace(str, pos + 1, depth - 1)
      ?\\ -> find_matching_brace(str, pos + 2, depth)
      _ -> find_matching_brace(str, pos + 1, depth)
    end
  end

  # Remove remaining RTF control words (e.g., \fonttbl, \fs20, etc.).
  defp remove_control_words(str) do
    Regex.replace(~r/\\[a-z]+[-]?\d*\s?/, str, "")
  end
end
