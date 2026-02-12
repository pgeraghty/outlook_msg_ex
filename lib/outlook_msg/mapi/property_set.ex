defmodule OutlookMsg.Mapi.PropertySet do
  @moduledoc """
  Wraps a raw MAPI property map (`%Key{} => value`) with friendly access methods.

  Based on `property_set.rb` from ruby-msg. Provides symbolic lookup by
  property name atoms (e.g. `:pr_subject`), integer code lookup, and
  convenience accessors for common message properties such as subject, body,
  sender, recipients, timestamps, importance, and sensitivity.

  Implements the `Access` behaviour so properties can be fetched with
  bracket syntax:

      ps[:pr_subject]

  ## Examples

      iex> alias OutlookMsg.Mapi.{Key, PropertySet}
      iex> props = %{Key.new(0x0037) => "Hello"}
      iex> ps = PropertySet.new(props)
      iex> PropertySet.subject(ps)
      "Hello"
  """

  alias OutlookMsg.Mapi.Key

  defstruct properties: %{}

  @type t :: %__MODULE__{
          properties: %{Key.t() => term()}
        }

  @behaviour Access

  # -------------------------------------------------------------------
  # Constructors
  # -------------------------------------------------------------------

  @doc """
  Creates a new `PropertySet` from a map of `%Key{} => value` pairs.

  ## Examples

      iex> alias OutlookMsg.Mapi.{Key, PropertySet}
      iex> ps = PropertySet.new(%{Key.new(0x0037) => "Test"})
      iex> ps.properties |> map_size()
      1
  """
  @spec new(map()) :: t()
  def new(props) when is_map(props) do
    %__MODULE__{properties: props}
  end

  # -------------------------------------------------------------------
  # Core property access
  # -------------------------------------------------------------------

  @doc """
  Gets a property value by symbolic atom name or integer code.

  When given an atom (e.g. `:pr_subject`), searches all keys for one whose
  `Key.to_sym/1` matches. When given an integer code, searches for a key
  with that code regardless of GUID.

  Returns the value or `nil` if not found.

  ## Examples

      iex> alias OutlookMsg.Mapi.{Key, PropertySet}
      iex> ps = PropertySet.new(%{Key.new(0x0037) => "Hello"})
      iex> PropertySet.get(ps, :pr_subject)
      "Hello"

      iex> alias OutlookMsg.Mapi.{Key, PropertySet}
      iex> ps = PropertySet.new(%{Key.new(0x0037) => "Hello"})
      iex> PropertySet.get(ps, 0x0037)
      "Hello"
  """
  @spec get(t(), atom() | non_neg_integer()) :: term() | nil
  def get(%__MODULE__{properties: props}, name) when is_atom(name) do
    Enum.find_value(props, fn {key, value} ->
      if Key.to_sym(key) == name, do: value
    end)
  end

  def get(%__MODULE__{properties: props}, code) when is_integer(code) do
    Enum.find_value(props, fn {%Key{code: c}, value} ->
      if c == code, do: value
    end)
  end

  @doc """
  Gets a property value by code and GUID.

  Performs an exact `Map.get/2` lookup using a `%Key{}` constructed from
  the given code and GUID.

  Returns the value or `nil` if not found.

  ## Examples

      iex> alias OutlookMsg.Mapi.{Key, Guids, PropertySet}
      iex> key = Key.new(0x0037, Guids.ps_mapi())
      iex> ps = PropertySet.new(%{key => "Hello"})
      iex> PropertySet.get(ps, 0x0037, Guids.ps_mapi())
      "Hello"
  """
  @spec get(t(), non_neg_integer(), binary()) :: term() | nil
  def get(%__MODULE__{properties: props}, code, guid) do
    Map.get(props, %Key{code: code, guid: guid})
  end

  # -------------------------------------------------------------------
  # Access behaviour
  # -------------------------------------------------------------------

  @impl Access
  def fetch(%__MODULE__{} = ps, key) do
    case get(ps, key) do
      nil -> :error
      value -> {:ok, value}
    end
  end

  @impl Access
  def get_and_update(%__MODULE__{properties: props} = ps, key, fun) when is_atom(key) do
    current = get(ps, key)
    {get_value, update_value} = fun.(current)

    new_props =
      case find_key(props, key) do
        nil -> props
        found_key -> Map.put(props, found_key, update_value)
      end

    {get_value, %__MODULE__{ps | properties: new_props}}
  end

  def get_and_update(%__MODULE__{properties: props} = ps, code, fun) when is_integer(code) do
    current = get(ps, code)
    {get_value, update_value} = fun.(current)

    new_props =
      case find_key_by_code(props, code) do
        nil -> props
        found_key -> Map.put(props, found_key, update_value)
      end

    {get_value, %__MODULE__{ps | properties: new_props}}
  end

  @impl Access
  def pop(%__MODULE__{properties: props} = ps, key) when is_atom(key) do
    case find_key(props, key) do
      nil ->
        {nil, ps}

      found_key ->
        {value, new_props} = Map.pop(props, found_key)
        {value, %__MODULE__{ps | properties: new_props}}
    end
  end

  def pop(%__MODULE__{properties: props} = ps, code) when is_integer(code) do
    case find_key_by_code(props, code) do
      nil ->
        {nil, ps}

      found_key ->
        {value, new_props} = Map.pop(props, found_key)
        {value, %__MODULE__{ps | properties: new_props}}
    end
  end

  # -------------------------------------------------------------------
  # Convenience accessors
  # -------------------------------------------------------------------

  @doc "Returns the message subject (`PR_SUBJECT`)."
  @spec subject(t()) :: String.t() | nil
  def subject(%__MODULE__{} = ps), do: get(ps, :pr_subject)

  @doc """
  Returns the plain-text message body (`PR_BODY`).

  If the property is not set directly, attempts to extract body text from
  the compressed RTF body via `body_from_rtf/1`.
  """
  @spec body(t()) :: String.t() | nil
  def body(%__MODULE__{} = ps) do
    case get(ps, :pr_body) do
      nil -> body_from_rtf(ps)
      value -> value
    end
  end

  @doc """
  Returns the decompressed RTF body.

  Decompresses the `PR_RTF_COMPRESSED` property using
  `OutlookMsg.Rtf.decompress/1`. Returns `nil` if the property is not
  present or decompression fails.
  """
  @spec body_rtf(t()) :: String.t() | nil
  def body_rtf(%__MODULE__{} = ps) do
    case get(ps, :pr_rtf_compressed) do
      nil ->
        nil

      compressed ->
        case safe_rtf_call(:decompress, [compressed]) do
          {:ok, rtf} -> rtf
          _ -> nil
        end
    end
  end

  @doc """
  Returns the HTML body of the message.

  First tries the `PR_BODY_HTML` property directly. If not available,
  attempts to extract HTML from the compressed RTF body via RTF-to-HTML
  conversion. Ensures the result is returned as a UTF-8 string.
  """
  @spec body_html(t()) :: String.t() | nil
  def body_html(%__MODULE__{} = ps) do
    case get(ps, :pr_body_html) do
      nil ->
        case body_rtf(ps) do
          nil ->
            nil

          rtf ->
            case safe_rtf_call(:rtf_to_html, [rtf]) do
              {:ok, html} -> ensure_string(html)
              _ -> nil
            end
        end

      value ->
        ensure_string(value)
    end
  end

  @doc "Returns the message class (`PR_MESSAGE_CLASS`), e.g. `\"IPM.Note\"`."
  @spec message_class(t()) :: String.t() | nil
  def message_class(%__MODULE__{} = ps), do: get(ps, :pr_message_class)

  @doc "Returns the sender display name (`PR_SENDER_NAME`)."
  @spec sender_name(t()) :: String.t() | nil
  def sender_name(%__MODULE__{} = ps), do: get(ps, :pr_sender_name)

  @doc "Returns the sender email address (`PR_SENDER_EMAIL_ADDRESS`)."
  @spec sender_email(t()) :: String.t() | nil
  def sender_email(%__MODULE__{} = ps), do: get(ps, :pr_sender_email_address)

  @doc "Returns the display-to recipients (`PR_DISPLAY_TO`)."
  @spec display_to(t()) :: String.t() | nil
  def display_to(%__MODULE__{} = ps), do: get(ps, :pr_display_to)

  @doc "Returns the display-cc recipients (`PR_DISPLAY_CC`)."
  @spec display_cc(t()) :: String.t() | nil
  def display_cc(%__MODULE__{} = ps), do: get(ps, :pr_display_cc)

  @doc "Returns the display-bcc recipients (`PR_DISPLAY_BCC`)."
  @spec display_bcc(t()) :: String.t() | nil
  def display_bcc(%__MODULE__{} = ps), do: get(ps, :pr_display_bcc)

  @doc "Returns the creation time (`PR_CREATION_TIME`) as a `DateTime`."
  @spec creation_time(t()) :: DateTime.t() | nil
  def creation_time(%__MODULE__{} = ps), do: get(ps, :pr_creation_time)

  @doc "Returns the last modification time (`PR_LAST_MODIFICATION_TIME`) as a `DateTime`."
  @spec last_modification_time(t()) :: DateTime.t() | nil
  def last_modification_time(%__MODULE__{} = ps), do: get(ps, :pr_last_modification_time)

  @doc "Returns the message delivery time (`PR_MESSAGE_DELIVERY_TIME`) as a `DateTime`."
  @spec delivery_time(t()) :: DateTime.t() | nil
  def delivery_time(%__MODULE__{} = ps), do: get(ps, :pr_message_delivery_time)

  @doc "Returns the client submit time (`PR_CLIENT_SUBMIT_TIME`) as a `DateTime`."
  @spec client_submit_time(t()) :: DateTime.t() | nil
  def client_submit_time(%__MODULE__{} = ps), do: get(ps, :pr_client_submit_time)

  @doc """
  Returns the message importance as an atom.

  Maps the integer `PR_IMPORTANCE` value:
    - `0` -> `:low`
    - `1` -> `:normal`
    - `2` -> `:high`

  Returns `nil` if not set.
  """
  @spec importance(t()) :: :low | :normal | :high | nil
  def importance(%__MODULE__{} = ps) do
    case get(ps, :pr_importance) do
      0 -> :low
      1 -> :normal
      2 -> :high
      _ -> nil
    end
  end

  @doc """
  Returns the message sensitivity as an atom.

  Maps the integer `PR_SENSITIVITY` value:
    - `0` -> `:none`
    - `1` -> `:personal`
    - `2` -> `:private`
    - `3` -> `:confidential`

  Returns `nil` if not set.
  """
  @spec sensitivity(t()) :: :none | :personal | :private | :confidential | nil
  def sensitivity(%__MODULE__{} = ps) do
    case get(ps, :pr_sensitivity) do
      0 -> :none
      1 -> :personal
      2 -> :private
      3 -> :confidential
      _ -> nil
    end
  end

  @doc "Returns the raw `PR_MESSAGE_FLAGS` integer value."
  @spec message_flags(t()) :: non_neg_integer() | nil
  def message_flags(%__MODULE__{} = ps), do: get(ps, :pr_message_flags)

  # -------------------------------------------------------------------
  # Enumeration helpers
  # -------------------------------------------------------------------

  @doc """
  Returns a list of all `%Key{}` structs in the property set.

  ## Examples

      iex> alias OutlookMsg.Mapi.{Key, PropertySet}
      iex> ps = PropertySet.new(%{Key.new(0x0037) => "Hello"})
      iex> PropertySet.keys(ps) |> length()
      1
  """
  @spec keys(t()) :: [Key.t()]
  def keys(%__MODULE__{properties: props}), do: Map.keys(props)

  @doc """
  Returns a list of all property values in the property set.

  ## Examples

      iex> alias OutlookMsg.Mapi.{Key, PropertySet}
      iex> ps = PropertySet.new(%{Key.new(0x0037) => "Hello"})
      iex> PropertySet.values(ps)
      ["Hello"]
  """
  @spec values(t()) :: [term()]
  def values(%__MODULE__{properties: props}), do: Map.values(props)

  @doc """
  Returns the raw properties map (`%Key{} => value`).

  ## Examples

      iex> alias OutlookMsg.Mapi.{Key, PropertySet}
      iex> props = %{Key.new(0x0037) => "Hello"}
      iex> ps = PropertySet.new(props)
      iex> PropertySet.to_map(ps) == props
      true
  """
  @spec to_map(t()) :: %{Key.t() => term()}
  def to_map(%__MODULE__{properties: props}), do: props

  @doc """
  Returns a map with symbolic atom keys where possible.

  Each `%Key{}` is resolved via `Key.to_sym/1`. If resolution succeeds,
  the atom is used as the map key; otherwise, the original `%Key{}` struct
  is kept.

  ## Examples

      iex> alias OutlookMsg.Mapi.{Key, PropertySet}
      iex> ps = PropertySet.new(%{Key.new(0x0037) => "Hello"})
      iex> sym = PropertySet.to_symbolic_map(ps)
      iex> sym[:pr_subject]
      "Hello"
  """
  @spec to_symbolic_map(t()) :: %{(atom() | Key.t()) => term()}
  def to_symbolic_map(%__MODULE__{properties: props}) do
    Map.new(props, fn {key, value} ->
      case Key.to_sym(key) do
        nil -> {key, value}
        name -> {name, value}
      end
    end)
  end

  # -------------------------------------------------------------------
  # Private helpers
  # -------------------------------------------------------------------

  # Attempts to extract a text body from compressed RTF.
  # First tries RTF-to-HTML conversion, then falls back to RTF-to-text.
  # Returns nil if no RTF is available or conversion fails.
  @spec body_from_rtf(t()) :: String.t() | nil
  defp body_from_rtf(%__MODULE__{} = ps) do
    case body_rtf(ps) do
      nil ->
        nil

      rtf ->
        case safe_rtf_call(:rtf_to_html, [rtf]) do
          {:ok, html} ->
            ensure_string(html)

          _ ->
            case safe_rtf_call(:rtf_to_text, [rtf]) do
              {:ok, text} -> ensure_string(text)
              _ -> nil
            end
        end
    end
  end

  # Safely calls a function on the OutlookMsg.Rtf module.
  # Returns {:ok, result} on success or :error if the module is not
  # available or the function raises.
  @spec safe_rtf_call(atom(), [term()]) :: {:ok, term()} | :error
  defp safe_rtf_call(function, args) do
    case Code.ensure_loaded(OutlookMsg.Rtf) do
      {:module, mod} ->
        if function_exported?(mod, function, length(args)) do
          case apply(mod, function, args) do
            {:ok, value} -> {:ok, value}
            :none -> :error
            {:error, _reason} -> :error
            value -> {:ok, value}
          end
        else
          :error
        end

      {:error, _} ->
        :error
    end
  rescue
    _ -> :error
  end

  # Finds a Key struct in the properties map whose to_sym matches the given atom name.
  @spec find_key(map(), atom()) :: Key.t() | nil
  defp find_key(props, name) when is_atom(name) do
    Enum.find_value(props, fn {key, _value} ->
      if Key.to_sym(key) == name, do: key
    end)
  end

  # Finds a Key struct in the properties map whose code matches the given integer.
  @spec find_key_by_code(map(), non_neg_integer()) :: Key.t() | nil
  defp find_key_by_code(props, code) when is_integer(code) do
    Enum.find_value(props, fn {%Key{code: c} = key, _value} ->
      if c == code, do: key
    end)
  end

  # Ensures a value is returned as a UTF-8 string.
  # If the value is already a string, returns it as-is.
  # If it is a binary, attempts to interpret it as UTF-8.
  @spec ensure_string(term()) :: String.t() | nil
  defp ensure_string(value) when is_binary(value) do
    if String.valid?(value) do
      value
    else
      case :unicode.characters_to_binary(value, :latin1, :utf8) do
        result when is_binary(result) -> result
        _ -> value
      end
    end
  end

  defp ensure_string(_), do: nil
end
