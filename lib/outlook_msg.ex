defmodule OutlookMsg do
  @moduledoc """
  Elixir library for reading Microsoft Outlook MSG and PST files.

  Provides:
  - OLE/CFB container parsing
  - MSG file reading with full property extraction
  - PST file reading (97 and 2003 formats)
  - MAPI property system with symbolic name resolution
  - RTF decompression (LZFu) and RTF-to-HTML extraction
  - MIME/RFC2822 email conversion
  """

  alias OutlookMsg.{Msg, Pst, Mime, Convert}
  alias OutlookMsg.Warning

  @doc """
  Open an MSG file from a file path or binary data.

  Returns `{:ok, %OutlookMsg.Msg{}}` or `{:error, reason}`.

  ## Examples

      {:ok, msg} = OutlookMsg.open("email.msg")
      msg.properties |> OutlookMsg.Mapi.PropertySet.subject()
  """
  def open(path_or_binary) do
    Msg.open(path_or_binary)
  end

  @doc """
  Open an MSG file and return parsed content plus non-fatal parse warnings.
  """
  def open_with_warnings(path_or_binary) do
    case open_with_report(path_or_binary) do
      {:ok, msg, warnings} -> {:ok, msg, Warning.format_all(warnings)}
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Open an MSG file and return parsed content with structured warnings.
  """
  def open_with_report(path_or_binary) do
    result =
      try do
        Msg.open(path_or_binary)
      rescue
        e -> {:error, {:exception, Exception.message(e)}}
      catch
        kind, reason -> {:error, {:exception, {kind, reason}}}
      end

    case result do
      {:ok, msg} -> {:ok, msg, msg.warnings || []}
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Open/parse an EML (RFC2822) message from a path or raw message string.

  Returns `{:ok, %OutlookMsg.Mime{}}` or `{:error, reason}`.

  If the input contains CR/LF line breaks, it is treated as raw EML text.
  Otherwise, when it points to an existing file, the file is read and parsed.
  """
  def open_eml(path_or_raw) when is_binary(path_or_raw) do
    raw_or_err =
      cond do
        String.contains?(path_or_raw, "\n") or String.contains?(path_or_raw, "\r") ->
          {:ok, path_or_raw}

        File.regular?(path_or_raw) ->
          File.read(path_or_raw)

        true ->
          {:ok, path_or_raw}
      end

    case raw_or_err do
      {:ok, raw} -> {:ok, Mime.new(raw)}
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Open/parse an EML message and return parse warnings alongside the MIME struct.
  """
  def open_eml_with_warnings(path_or_raw) when is_binary(path_or_raw) do
    case open_eml_with_report(path_or_raw) do
      {:ok, mime, warnings} -> {:ok, mime, Warning.format_all(warnings)}
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Open/parse an EML message and return structured warnings alongside the MIME struct.
  """
  def open_eml_with_report(path_or_raw) when is_binary(path_or_raw) do
    result =
      try do
        open_eml(path_or_raw)
      rescue
        e -> {:error, {:exception, Exception.message(e)}}
      catch
        kind, reason -> {:error, {:exception, {kind, reason}}}
      end

    case result do
      {:ok, mime} -> {:ok, mime, mime.warnings || []}
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Open a PST file from a file path or binary data.

  Returns `{:ok, %OutlookMsg.Pst{}}` or `{:error, reason}`.

  ## Examples

      {:ok, pst} = OutlookMsg.open_pst("archive.pst")
      OutlookMsg.Pst.items(pst) |> Enum.each(fn item -> ... end)
  """
  def open_pst(path_or_binary) do
    Pst.open(path_or_binary)
  end

  @doc """
  Open a PST file and return parse warnings alongside the PST struct.
  """
  def open_pst_with_warnings(path_or_binary) do
    case open_pst_with_report(path_or_binary) do
      {:ok, pst, warnings} -> {:ok, pst, Warning.format_all(warnings)}
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Open a PST file and return structured warnings alongside the PST struct.
  """
  def open_pst_with_report(path_or_binary) do
    result =
      try do
        Pst.open(path_or_binary)
      rescue
        e -> {:error, {:exception, Exception.message(e)}}
      catch
        kind, reason -> {:error, {:exception, {kind, reason}}}
      end

    case result do
      {:ok, pst} -> {:ok, pst, pst.warnings || []}
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Convert an MSG message to a MIME struct.

  Returns `{:ok, %OutlookMsg.Mime{}}` or `{:error, reason}`.
  """
  def to_mime(%Msg{} = msg) do
    Convert.to_mime(msg)
  end

  @doc """
  Convert an MSG message to an RFC2822 email string.

  Returns `{:ok, string}` or `{:error, reason}`.
  """
  def to_eml(%Msg{} = msg) do
    with {:ok, mime} <- Convert.to_mime(msg) do
      {:ok, Mime.to_string(mime)}
    end
  end

  @doc """
  Open an MSG file and immediately convert to MIME.

  Returns `{:ok, %OutlookMsg.Mime{}}` or `{:error, reason}`.
  """
  def msg_to_mime(path_or_binary) do
    with {:ok, msg} <- open(path_or_binary) do
      to_mime(msg)
    end
  end

  @doc """
  Open an MSG file and immediately convert to EML string.

  Returns `{:ok, string}` or `{:error, reason}`.
  """
  def msg_to_eml(path_or_binary) do
    with {:ok, msg} <- open(path_or_binary) do
      to_eml(msg)
    end
  end

  @doc """
  Serialize a parsed MIME/EML struct back to RFC2822 string form.

  Returns the EML string.
  """
  def eml_to_string(%Mime{} = mime) do
    Mime.to_string(mime)
  end
end
