defmodule OutlookMsg.Msg do
  alias OutlookMsg.Ole.Storage
  alias OutlookMsg.Msg.{PropertyStore, Attachment, Recipient}
  alias OutlookMsg.Mapi.PropertySet
  alias OutlookMsg.Warning

  defstruct [:properties, :attachments, :recipients, :storage, warnings: []]

  @doc "Open an MSG file from path or binary data"
  def open(path_or_binary) do
    with {:ok, storage} <- Storage.open(path_or_binary) do
      root = Storage.root(storage)

      # Load nameid mapping first
      {nameid, warnings1} = load_nameid(storage, root)

      # Load root properties (32-byte header skip for root message)
      {props, warnings2} = safe_load_properties(storage, root, nameid, 32, "root properties")
      property_set = PropertySet.new(props)

      # Load attachments
      {attachments, warnings3} = load_attachments(storage, root, nameid)

      # Load recipients
      {recipients, warnings4} = load_recipients(storage, root, nameid)

      {:ok, %__MODULE__{
        properties: property_set,
        attachments: attachments,
        recipients: recipients,
        storage: storage,
        warnings: warnings1 ++ warnings2 ++ warnings3 ++ warnings4
      }}
    end
  end

  defp load_nameid(storage, root) do
    case Storage.find(storage, root, "__nameid_version1.0") do
      nil ->
        {%{}, []}

      nameid_dir ->
        try do
          {PropertyStore.parse_nameid(storage, nameid_dir), []}
        rescue
          e ->
            {%{}, [Warning.new(:nameid_parse_failed, "continuing without named-property mapping", context: Exception.message(e))]}
        end
    end
  end

  defp load_attachments(storage, parent, nameid) do
    parent.children
    |> Enum.filter(fn d ->
      d.type in [:storage, :root] and
      String.starts_with?(String.downcase(d.name), "__attach_version1.0_")
    end)
    |> Enum.sort_by(fn d -> d.name end)
    |> Enum.reduce({[], []}, fn d, {atts, warnings} ->
      case safe_load_attachment(storage, d, nameid) do
        {:ok, att, att_warnings} ->
          {atts ++ [att], warnings ++ att_warnings}

        {:error, reason} ->
          {atts, warnings ++ [Warning.new(:attachment_skipped, "attachment skipped", context: "#{d.name}: #{reason}")]}
      end
    end)
  end

  defp load_recipients(storage, parent, nameid) do
    parent.children
    |> Enum.filter(fn d ->
      d.type in [:storage, :root] and
      String.starts_with?(String.downcase(d.name), "__recip_version1.0_")
    end)
    |> Enum.sort_by(fn d -> d.name end)
    |> Enum.reduce({[], []}, fn d, {recips, warnings} ->
      case safe_load_properties(storage, d, nameid, 8, "recipient #{d.name}") do
        {props, rec_warnings} ->
          recip = props |> PropertySet.new() |> Recipient.new()
          {recips ++ [recip], warnings ++ rec_warnings}
      end
    end)
  end

  defp safe_load_attachment(storage, d, nameid) do
    {props, warnings1} = safe_load_properties(storage, d, nameid, 8, "attachment #{d.name}")
    property_set = PropertySet.new(props)
    att = Attachment.new(property_set)

    case PropertySet.get(property_set, :pr_attach_method) do
      5 ->
        case Storage.find(storage, d, "__substg1.0_3701000D") do
          nil ->
            {:ok, att, warnings1}

          embed_dir ->
            {embed_props, warnings2} = safe_load_properties(storage, embed_dir, nameid, 32, "embedded message in #{d.name}")
            embed_ps = PropertySet.new(embed_props)
            {embed_attachments, warnings3} = load_attachments(storage, embed_dir, nameid)
            {embed_recipients, warnings4} = load_recipients(storage, embed_dir, nameid)

            embedded = %OutlookMsg.Msg{
              properties: embed_ps,
              attachments: embed_attachments,
              recipients: embed_recipients,
              storage: storage,
              warnings: warnings2 ++ warnings3 ++ warnings4
            }

            {:ok, %{att | embedded_msg: embedded}, warnings1 ++ warnings2 ++ warnings3 ++ warnings4}
        end

      _ ->
        {:ok, att, warnings1}
    end
    rescue
    e ->
      {:error, Exception.message(e)}
  end

  defp safe_load_properties(storage, dirent, nameid, prefix_size, context) do
    try do
      {PropertyStore.load(storage, dirent, nameid, prefix_size), []}
    rescue
      e ->
        {%{}, [Warning.new(:property_parse_failed, "property parsing failed", context: "#{context}: #{Exception.message(e)}")]}
    end
  end
end
