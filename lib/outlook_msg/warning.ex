defmodule OutlookMsg.Warning do
  @moduledoc """
  Structured non-fatal parser warning.
  """

  @enforce_keys [:code, :message]
  defstruct [:code, :message, severity: :warn, context: nil, recoverable: true]

  @type t :: %__MODULE__{
          code: atom(),
          message: String.t(),
          severity: :info | :warn | :error,
          context: String.t() | nil,
          recoverable: boolean()
        }

  @spec new(atom(), String.t(), keyword()) :: t()
  def new(code, message, opts \\ []) when is_atom(code) and is_binary(message) do
    %__MODULE__{
      code: code,
      message: message,
      severity: Keyword.get(opts, :severity, :warn),
      context: Keyword.get(opts, :context),
      recoverable: Keyword.get(opts, :recoverable, true)
    }
  end

  @spec format(t() | String.t()) :: String.t()
  def format(%__MODULE__{} = w) do
    prefix =
      case w.context do
        nil -> "[#{w.severity}:#{w.code}]"
        ctx -> "[#{w.severity}:#{w.code}] #{ctx}:"
      end

    "#{prefix} #{w.message}"
  end

  def format(text) when is_binary(text), do: text

  @spec format_all([t() | String.t()]) :: [String.t()]
  def format_all(warnings) when is_list(warnings), do: Enum.map(warnings, &format/1)
end

defimpl String.Chars, for: OutlookMsg.Warning do
  def to_string(w), do: OutlookMsg.Warning.format(w)
end
