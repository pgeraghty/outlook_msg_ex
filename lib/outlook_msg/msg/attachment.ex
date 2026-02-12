defmodule OutlookMsg.Msg.Attachment do
  alias OutlookMsg.Mapi.PropertySet

  defstruct [:properties, :filename, :data, :mime_type, :embedded_msg, :embedded_ole]

  def new(properties) do
    filename = PropertySet.get(properties, :pr_attach_long_filename) ||
               PropertySet.get(properties, :pr_attach_filename) ||
               "attachment"

    data = PropertySet.get(properties, :pr_attach_data_bin)
    mime_type = PropertySet.get(properties, :pr_attach_mime_tag)

    %__MODULE__{
      properties: properties,
      filename: filename,
      data: data,
      mime_type: mime_type,
      embedded_msg: nil,
      embedded_ole: nil
    }
  end

  def content_id(%__MODULE__{properties: props}), do: PropertySet.get(props, :pr_attach_content_id)
  def content_location(%__MODULE__{properties: props}), do: PropertySet.get(props, :pr_attach_content_location)
  def content_disposition(%__MODULE__{properties: props}), do: PropertySet.get(props, :pr_attach_content_disposition)
  def method(%__MODULE__{properties: props}), do: PropertySet.get(props, :pr_attach_method)
  def rendering_position(%__MODULE__{properties: props}), do: PropertySet.get(props, :pr_rendering_position)
  def extension(%__MODULE__{properties: props}), do: PropertySet.get(props, :pr_attach_extension)

  def inline?(%__MODULE__{} = att) do
    content_id(att) != nil || content_location(att) != nil
  end

  def embedded?(%__MODULE__{embedded_msg: msg}), do: msg != nil
end
