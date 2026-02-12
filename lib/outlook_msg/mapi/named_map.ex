defmodule OutlookMsg.Mapi.NamedMap do
  @moduledoc """
  Map of `{code, guid_binary}` tuples to symbolic atom names for MAPI named
  properties.

  This replaces the `named_map.yaml` file from ruby-msg. Each entry maps a
  numeric property code together with the 16-byte mixed-endian GUID of the
  property set to a human-readable atom name.

  ## Usage

      iex> guid = OutlookMsg.Mapi.NamedMap.encode_guid("00062008-0000-0000-C000-000000000046")
      iex> OutlookMsg.Mapi.NamedMap.lookup(0x8503, guid)
      :reminder_set

      iex> OutlookMsg.Mapi.NamedMap.all() |> map_size()
      78
  """

  # -------------------------------------------------------------------
  # Compile-time GUID encoding
  # -------------------------------------------------------------------

  # Converts a GUID string like "00062004-0000-0000-C000-000000000046"
  # into a 16-byte mixed-endian binary at compile time.
  #
  # Layout:
  #   - data1 (4 bytes): little-endian
  #   - data2 (2 bytes): little-endian
  #   - data3 (2 bytes): little-endian
  #   - data4 (2 bytes): big-endian
  #   - data5 (6 bytes): big-endian
  @doc false
  @spec encode_guid(String.t()) :: <<_::128>>
  def encode_guid(guid_string) do
    [d1, d2, d3, d4, d5] = String.split(guid_string, "-")

    {data1, ""} = Integer.parse(d1, 16)
    {data2, ""} = Integer.parse(d2, 16)
    {data3, ""} = Integer.parse(d3, 16)
    {data4, ""} = Integer.parse(d4, 16)
    {data5, ""} = Integer.parse(d5, 16)

    <<
      data1::little-unsigned-integer-size(32),
      data2::little-unsigned-integer-size(16),
      data3::little-unsigned-integer-size(16),
      data4::big-unsigned-integer-size(16),
      data5::big-unsigned-integer-size(48)
    >>
  end

  # -------------------------------------------------------------------
  # GUID string constants
  # -------------------------------------------------------------------

  @psetid_address "00062004-0000-0000-C000-000000000046"
  @psetid_appointment "00062002-0000-0000-C000-000000000046"
  @psetid_common "00062008-0000-0000-C000-000000000046"
  @psetid_task "00062003-0000-0000-C000-000000000046"
  @psetid_log "0006200A-0000-0000-C000-000000000046"
  @ps_internet_headers "00020386-0000-0000-C000-000000000046"

  # -------------------------------------------------------------------
  # Named property definitions
  # -------------------------------------------------------------------

  # Each entry is {code, guid_string, atom_name}.
  @named_property_definitions [
    # PSETID_ADDRESS
    {0x8005, @psetid_address, :file_under},
    {0x8006, @psetid_address, :file_under_id},
    {0x800E, @psetid_address, :email1_display_name},
    {0x800F, @psetid_address, :email1_address_type},
    {0x8010, @psetid_address, :email1_email_address},
    {0x8014, @psetid_address, :email1_original_display_name},
    {0x8015, @psetid_address, :email1_original_entryid},
    {0x8023, @psetid_address, :email1_display_name_2},
    {0x8045, @psetid_address, :email2_display_name},
    {0x8046, @psetid_address, :email2_address_type},
    {0x8047, @psetid_address, :email2_email_address},
    {0x8053, @psetid_address, :email2_original_display_name},
    {0x8054, @psetid_address, :email2_original_entryid},
    {0x8080, @psetid_address, :email3_display_name},
    {0x8081, @psetid_address, :email3_address_type},
    {0x8082, @psetid_address, :email3_email_address},
    {0x8084, @psetid_address, :email3_original_display_name},
    {0x8085, @psetid_address, :email3_original_entryid},
    {0x80B2, @psetid_address, :home_address},
    {0x80B8, @psetid_address, :work_address},
    {0x80B9, @psetid_address, :other_address},
    {0x80D8, @psetid_address, :internet_free_busy_address},

    # PSETID_APPOINTMENT
    {0x8201, @psetid_appointment, :appt_sequence},
    {0x8205, @psetid_appointment, :busy_status},
    {0x8208, @psetid_appointment, :location},
    {0x820D, @psetid_appointment, :appt_start_whole},
    {0x820E, @psetid_appointment, :appt_end_whole},
    {0x8213, @psetid_appointment, :appt_duration},
    {0x8214, @psetid_appointment, :appt_color},
    {0x8215, @psetid_appointment, :appt_sub_type},
    {0x8216, @psetid_appointment, :appt_state_flags},
    {0x8217, @psetid_appointment, :response_status},
    {0x8218, @psetid_appointment, :recurring},
    {0x8223, @psetid_appointment, :is_recurring},
    {0x8228, @psetid_appointment, :recurrence_type},
    {0x8229, @psetid_appointment, :recurrence_pattern},
    {0x8231, @psetid_appointment, :time_zone},
    {0x8232, @psetid_appointment, :clip_start},
    {0x8233, @psetid_appointment, :clip_end},
    {0x8234, @psetid_appointment, :original_store_entryid},
    {0x8235, @psetid_appointment, :all_day_event},
    {0x8236, @psetid_appointment, :appt_message_class},

    # PSETID_COMMON
    {0x8501, @psetid_common, :reminder_delta},
    {0x8502, @psetid_common, :reminder_time},
    {0x8503, @psetid_common, :reminder_set},
    {0x8506, @psetid_common, :private},
    {0x8510, @psetid_common, :smart_no_attach},
    {0x8514, @psetid_common, :sidebar_image},
    {0x8516, @psetid_common, :common_start},
    {0x8517, @psetid_common, :common_end},
    {0x8520, @psetid_common, :task_mode},
    {0x8530, @psetid_common, :companies},
    {0x8535, @psetid_common, :billing},
    {0x8539, @psetid_common, :contacts},
    {0x8554, @psetid_common, :current_version},
    {0x8560, @psetid_common, :reminder_signal_time},
    {0x8580, @psetid_common, :internet_account_name},
    {0x8581, @psetid_common, :internet_account_stamp},
    {0x8582, @psetid_common, :use_tnef},

    # PSETID_TASK
    {0x8101, @psetid_task, :task_status},
    {0x8102, @psetid_task, :percent_complete},
    {0x8103, @psetid_task, :team_task},
    {0x8104, @psetid_task, :task_start_date},
    {0x8105, @psetid_task, :task_due_date},
    {0x810F, @psetid_task, :task_date_completed},
    {0x8110, @psetid_task, :task_actual_effort},
    {0x8111, @psetid_task, :task_estimated_effort},
    {0x811C, @psetid_task, :task_complete},
    {0x811F, @psetid_task, :task_owner},
    {0x8121, @psetid_task, :task_delegator},
    {0x8126, @psetid_task, :task_is_recurring},
    {0x8129, @psetid_task, :task_ownership},

    # PSETID_LOG
    {0x8700, @psetid_log, :log_type},
    {0x8706, @psetid_log, :log_start},
    {0x8707, @psetid_log, :log_duration},
    {0x8708, @psetid_log, :log_end},

    # PS_INTERNET_HEADERS
    {0x0002, @ps_internet_headers, :x_sharing_flavor},
    {0x0003, @ps_internet_headers, :x_sharing_local_type}
  ]

  # -------------------------------------------------------------------
  # Build the lookup map at compile time
  # -------------------------------------------------------------------

  @named_map (for {code, guid_string, name} <- @named_property_definitions, into: %{} do
    [d1, d2, d3, d4, d5] = String.split(guid_string, "-")
    {data1, ""} = Integer.parse(d1, 16)
    {data2, ""} = Integer.parse(d2, 16)
    {data3, ""} = Integer.parse(d3, 16)
    {data4, ""} = Integer.parse(d4, 16)
    {data5, ""} = Integer.parse(d5, 16)

    guid_binary =
      <<
        data1::little-unsigned-integer-size(32),
        data2::little-unsigned-integer-size(16),
        data3::little-unsigned-integer-size(16),
        data4::big-unsigned-integer-size(16),
        data5::big-unsigned-integer-size(48)
      >>

    {{code, guid_binary}, name}
  end)

  # -------------------------------------------------------------------
  # Public API
  # -------------------------------------------------------------------

  @doc """
  Looks up a named property by its code and GUID binary.

  Returns the symbolic atom name if the `{code, guid}` pair is known,
  or `nil` if not found.

  ## Parameters

    - `code` - the integer property code (e.g. `0x8503`)
    - `guid` - the 16-byte mixed-endian binary GUID of the property set

  ## Examples

      iex> guid = OutlookMsg.Mapi.NamedMap.encode_guid("00062008-0000-0000-C000-000000000046")
      iex> OutlookMsg.Mapi.NamedMap.lookup(0x8503, guid)
      :reminder_set

      iex> guid = OutlookMsg.Mapi.NamedMap.encode_guid("00062004-0000-0000-C000-000000000046")
      iex> OutlookMsg.Mapi.NamedMap.lookup(0x8010, guid)
      :email1_email_address

      iex> OutlookMsg.Mapi.NamedMap.lookup(0xFFFF, <<0::128>>)
      nil
  """
  @spec lookup(non_neg_integer(), <<_::128>>) :: atom() | nil
  def lookup(code, <<_::binary-size(16)>> = guid) when is_integer(code) do
    Map.get(@named_map, {code, guid})
  end

  def lookup(_code, _guid), do: nil

  @doc """
  Returns the full map of named properties.

  The map keys are `{code, guid_binary}` tuples and the values are atom names.

  ## Examples

      iex> map = OutlookMsg.Mapi.NamedMap.all()
      iex> is_map(map)
      true

      iex> map = OutlookMsg.Mapi.NamedMap.all()
      iex> map_size(map)
      78
  """
  @spec all() :: %{{non_neg_integer(), <<_::128>>} => atom()}
  def all, do: @named_map
end
