defmodule OutlookMsg.Pst.Item do
  @moduledoc "Represents a PST item (message, folder, etc.)"

  alias OutlookMsg.Mapi.PropertySet

  defstruct [
    :desc,            # Descriptor record
    :properties,      # PropertySet
    :type,            # :message, :folder, :root, :search_folder
    :parent_id,       # Parent descriptor ID
    :children_ids,    # List of child descriptor IDs
    :pst_ref          # Reference back to PST for lazy loading
  ]

  @doc "Create an Item from a descriptor record and parsed properties"
  def new(desc, properties, type) do
    %__MODULE__{
      desc: desc,
      properties: properties,
      type: type,
      parent_id: desc.parent_desc_id,
      children_ids: desc.children || []
    }
  end

  @doc "Determine item type from message class property"
  def detect_type(properties) do
    case PropertySet.get(properties, :pr_message_class) do
      nil ->
        # Check if it has subfolder indicators
        if PropertySet.get(properties, :pr_content_count) != nil or
           PropertySet.get(properties, :pr_subfolders) != nil do
          :folder
        else
          :message
        end
      class when is_binary(class) ->
        class_down = String.downcase(class)
        cond do
          String.starts_with?(class_down, "ipm.note") -> :message
          String.starts_with?(class_down, "ipm.post") -> :message
          String.starts_with?(class_down, "ipm.appointment") -> :appointment
          String.starts_with?(class_down, "ipm.contact") -> :contact
          String.starts_with?(class_down, "ipm.task") -> :task
          String.starts_with?(class_down, "ipm.stickynote") -> :note
          String.starts_with?(class_down, "ipm.activity") -> :journal
          true -> :message
        end
    end
  end

  @doc "Get item display name"
  def display_name(%__MODULE__{properties: props}) do
    PropertySet.get(props, :pr_display_name) || PropertySet.get(props, :pr_subject) || ""
  end

  @doc "Get item subject"
  def subject(%__MODULE__{properties: props}), do: PropertySet.subject(props)

  @doc "Check if item is a folder"
  def folder?(%__MODULE__{type: :folder}), do: true
  def folder?(%__MODULE__{type: :root}), do: true
  def folder?(_), do: false

  @doc "Check if item is a message"
  def message?(%__MODULE__{type: :message}), do: true
  def message?(_), do: false

  @doc "Load attachments for this item (from sub-table)"
  def attachments(%__MODULE__{properties: _props}) do
    # Attachments would be in a sub-table referenced by PR_MESSAGE_ATTACHMENTS
    # This requires the full PST context to resolve
    []
  end

  @doc "Load recipients for this item (from sub-table)"
  def recipients(%__MODULE__{properties: _props}) do
    []
  end

  @doc "Iterate recursively over item and all descendants"
  def each_recursive(%__MODULE__{} = item, fun) do
    fun.(item)
    # Children would need to be resolved from PST descriptor hierarchy
    :ok
  end
end
