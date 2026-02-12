defmodule OutlookMsg.Msg.Recipient do
  alias OutlookMsg.Mapi.PropertySet

  defstruct [:properties, :name, :email, :type]

  @type_map %{0 => :orig, 1 => :to, 2 => :cc, 3 => :bcc}

  def new(properties) when is_struct(properties, PropertySet) do
    name = PropertySet.get(properties, :pr_transmittable_display_name) ||
           PropertySet.get(properties, :pr_display_name) ||
           PropertySet.get(properties, :pr_recipient_display_name)

    email = PropertySet.get(properties, :pr_smtp_address) ||
            PropertySet.get(properties, :pr_org_email_addr) ||
            PropertySet.get(properties, :pr_email_address)

    type_code = PropertySet.get(properties, :pr_recipient_type) || 1
    type = Map.get(@type_map, type_code, :to)

    %__MODULE__{
      properties: properties,
      name: name,
      email: email,
      type: type
    }
  end

  def to_string(%__MODULE__{name: nil, email: email}), do: "<#{email}>"
  def to_string(%__MODULE__{name: "", email: email}), do: "<#{email}>"
  def to_string(%__MODULE__{name: name, email: nil}), do: name
  def to_string(%__MODULE__{name: name, email: email}), do: ~s("#{name}" <#{email}>)

  defimpl String.Chars do
    def to_string(recipient), do: OutlookMsg.Msg.Recipient.to_string(recipient)
  end
end
