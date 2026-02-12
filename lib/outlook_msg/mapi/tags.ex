defmodule OutlookMsg.Mapi.Tags do
  @moduledoc """
  Lookup map from MAPI property tag codes to their symbolic names and types.

  This module replaces the `mapitags.yaml` file from ruby-msg, providing a
  compiled Elixir map for fast lookups of MAPI property tag metadata.

  Each entry maps an integer property tag code to a `{name, type}` tuple where
  `name` is a lowercase atom with the `:pr_` prefix (e.g. `:pr_subject`) and
  `type` is a property type atom (e.g. `:pt_tstring`, `:pt_long`, `:pt_binary`,
  `:pt_boolean`, `:pt_systime`, `:pt_object`).

  The `:pt_tstring` type indicates the property may appear as either an ANSI
  string (PT_STRING8) or a Unicode string (PT_UNICODE) depending on the actual
  type code in the property tag.
  """

  # -------------------------------------------------------------------
  # MAPI property tag -> {name, type} map
  # -------------------------------------------------------------------

  @tags %{
    0x0001 => {:pr_acknowledgement_mode, :pt_long},
    0x0002 => {:pr_alternate_recipient_allowed, :pt_boolean},
    0x0003 => {:pr_authorizing_users, :pt_binary},
    0x0004 => {:pr_auto_forward_comment, :pt_tstring},
    0x0005 => {:pr_auto_forwarded, :pt_boolean},
    0x0006 => {:pr_content_confidentiality_algorithm_id, :pt_binary},
    0x0007 => {:pr_content_correlator, :pt_binary},
    0x0008 => {:pr_content_identifier, :pt_tstring},
    0x0009 => {:pr_content_length, :pt_long},
    0x000A => {:pr_content_return_requested, :pt_boolean},
    0x000B => {:pr_conversation_key, :pt_binary},
    0x000C => {:pr_conversion_eits, :pt_binary},
    0x000D => {:pr_conversion_with_loss_prohibited, :pt_boolean},
    0x000E => {:pr_converted_eits, :pt_binary},
    0x000F => {:pr_deferred_delivery_time, :pt_systime},
    0x0010 => {:pr_deliver_time, :pt_systime},
    0x0011 => {:pr_discard_reason, :pt_long},
    0x0012 => {:pr_disclosure_of_recipients, :pt_boolean},
    0x0013 => {:pr_dl_expansion_history, :pt_binary},
    0x0014 => {:pr_dl_expansion_prohibited, :pt_boolean},
    0x0015 => {:pr_expiry_time, :pt_systime},
    0x0016 => {:pr_implicit_conversion_prohibited, :pt_boolean},
    0x0017 => {:pr_importance, :pt_long},
    0x0018 => {:pr_ipm_id, :pt_binary},
    0x0019 => {:pr_latest_delivery_time, :pt_systime},
    0x001A => {:pr_message_class, :pt_tstring},
    0x001B => {:pr_message_delivery_id, :pt_binary},
    0x001E => {:pr_message_security_label, :pt_binary},
    0x001F => {:pr_obsoleted_ipms, :pt_binary},
    0x0020 => {:pr_originally_intended_recipient_name, :pt_binary},
    0x0021 => {:pr_original_eits, :pt_binary},
    0x0022 => {:pr_originator_certificate, :pt_binary},
    0x0023 => {:pr_originator_delivery_report_requested, :pt_boolean},
    0x0024 => {:pr_originator_return_address, :pt_binary},
    0x0025 => {:pr_parent_key, :pt_binary},
    0x0026 => {:pr_priority, :pt_long},
    0x0027 => {:pr_origin_check, :pt_binary},
    0x0028 => {:pr_proof_of_delivery_requested, :pt_boolean},
    0x0029 => {:pr_read_receipt_requested, :pt_boolean},
    0x002A => {:pr_receipt_time, :pt_systime},
    0x002B => {:pr_recipient_reassignment_prohibited, :pt_boolean},
    0x002C => {:pr_redirection_history, :pt_binary},
    0x002D => {:pr_related_ipms, :pt_binary},
    0x002E => {:pr_original_sensitivity, :pt_long},
    0x002F => {:pr_languages, :pt_tstring},
    0x0030 => {:pr_reply_time, :pt_systime},
    0x0031 => {:pr_report_tag, :pt_binary},
    0x0032 => {:pr_report_disposition, :pt_tstring},
    0x0033 => {:pr_report_disposition_mode, :pt_tstring},
    0x0034 => {:pr_originator_non_delivery_report_requested, :pt_boolean},
    0x0036 => {:pr_sensitivity, :pt_long},
    0x0037 => {:pr_subject, :pt_tstring},
    0x0039 => {:pr_client_submit_time, :pt_systime},
    0x003A => {:pr_report_name, :pt_tstring},
    0x003B => {:pr_sent_representing_search_key, :pt_binary},
    0x003D => {:pr_subject_prefix, :pt_tstring},
    0x003F => {:pr_received_by_entryid, :pt_binary},
    0x0040 => {:pr_received_by_name, :pt_tstring},
    0x0041 => {:pr_sent_representing_entryid, :pt_binary},
    0x0042 => {:pr_sent_representing_name, :pt_tstring},
    0x0043 => {:pr_sent_representing_addrtype, :pt_tstring},
    0x0044 => {:pr_sent_representing_email_address, :pt_tstring},
    0x0045 => {:pr_original_sender_entryid, :pt_binary},
    0x0046 => {:pr_original_sender_name, :pt_tstring},
    0x0047 => {:pr_original_sender_search_key, :pt_binary},
    0x0048 => {:pr_original_sent_representing_entryid, :pt_binary},
    0x0049 => {:pr_original_sent_representing_name, :pt_tstring},
    0x004A => {:pr_original_sent_representing_search_key, :pt_binary},
    0x004B => {:pr_start_date, :pt_systime},
    0x004C => {:pr_end_date, :pt_systime},
    0x004D => {:pr_owner_appt_id, :pt_long},
    0x004E => {:pr_response_requested, :pt_boolean},
    0x0050 => {:pr_reply_recipient_entries, :pt_binary},
    0x0051 => {:pr_reply_recipient_names, :pt_tstring},
    0x0052 => {:pr_received_by_search_key, :pt_binary},
    0x0053 => {:pr_received_by_entryid2, :pt_binary},
    0x0054 => {:pr_received_by_name2, :pt_tstring},
    0x0055 => {:pr_original_sender_addrtype, :pt_tstring},
    0x0056 => {:pr_original_sender_email_address, :pt_tstring},
    0x0057 => {:pr_original_sent_representing_addrtype, :pt_tstring},
    0x0058 => {:pr_original_sent_representing_email_address, :pt_tstring},
    0x0059 => {:pr_conversation_topic, :pt_tstring},
    0x005A => {:pr_conversation_index, :pt_binary},
    0x005B => {:pr_original_display_bcc, :pt_tstring},
    0x005C => {:pr_original_display_cc, :pt_tstring},
    0x005D => {:pr_original_display_to, :pt_tstring},
    0x005E => {:pr_received_by_addrtype, :pt_tstring},
    0x005F => {:pr_received_by_email_address, :pt_tstring},
    0x0060 => {:pr_received_representing_addrtype, :pt_tstring},
    0x0061 => {:pr_received_representing_email_address, :pt_tstring},
    0x0062 => {:pr_original_author_addrtype, :pt_tstring},
    0x0063 => {:pr_original_author_email_address, :pt_tstring},
    0x0064 => {:pr_original_submit_time, :pt_systime},
    0x0065 => {:pr_reply_recipient_entries2, :pt_binary},
    0x0070 => {:pr_conversation_topic2, :pt_tstring},
    0x0071 => {:pr_conversation_index2, :pt_binary},
    0x0072 => {:pr_original_display_bcc2, :pt_tstring},
    0x0073 => {:pr_original_display_cc2, :pt_tstring},
    0x0074 => {:pr_original_display_to2, :pt_tstring},
    0x0075 => {:pr_received_by_addrtype2, :pt_tstring},
    0x0076 => {:pr_received_by_email_address2, :pt_tstring},
    0x0C06 => {:pr_non_receipt_notification_requested, :pt_boolean},
    0x0C15 => {:pr_recipient_type, :pt_long},
    0x0C17 => {:pr_reply_requested, :pt_boolean},
    0x0C19 => {:pr_sender_entryid, :pt_binary},
    0x0C1A => {:pr_sender_name, :pt_tstring},
    0x0C1B => {:pr_supplementary_info, :pt_tstring},
    0x0C1D => {:pr_sender_search_key, :pt_binary},
    0x0C1E => {:pr_sender_addrtype, :pt_tstring},
    0x0C1F => {:pr_sender_email_address, :pt_tstring},
    0x0E01 => {:pr_delete_after_submit, :pt_boolean},
    0x0E02 => {:pr_display_bcc, :pt_tstring},
    0x0E03 => {:pr_display_cc, :pt_tstring},
    0x0E04 => {:pr_display_to, :pt_tstring},
    0x0E06 => {:pr_message_delivery_time, :pt_systime},
    0x0E07 => {:pr_message_flags, :pt_long},
    0x0E08 => {:pr_message_size, :pt_long},
    0x0E09 => {:pr_parent_entryid, :pt_binary},
    0x0E0A => {:pr_sentmail_entryid, :pt_binary},
    0x0E0F => {:pr_responsibility, :pt_boolean},
    0x0E12 => {:pr_message_recipients, :pt_object},
    0x0E13 => {:pr_message_attachments, :pt_object},
    0x0E17 => {:pr_message_status, :pt_long},
    0x0E1B => {:pr_hasattach, :pt_boolean},
    0x0E1D => {:pr_normalized_subject, :pt_tstring},
    0x0E1F => {:pr_rtf_in_sync, :pt_boolean},
    0x0E20 => {:pr_attach_size, :pt_long},
    0x0E21 => {:pr_attach_num, :pt_long},
    0x0E28 => {:pr_primary_send_acct, :pt_tstring},
    0x0E29 => {:pr_next_send_acct, :pt_tstring},
    0x0E62 => {:pr_url_comp_name_set, :pt_boolean},
    0x0E79 => {:pr_trust_sender, :pt_long},
    0x0FF4 => {:pr_access, :pt_long},
    0x0FF7 => {:pr_access_level, :pt_long},
    0x0FF9 => {:pr_record_key, :pt_binary},
    0x0FFE => {:pr_object_type, :pt_long},
    0x0FFF => {:pr_entryid, :pt_binary},
    0x1000 => {:pr_body, :pt_tstring},
    0x1001 => {:pr_report_text, :pt_tstring},
    0x1006 => {:pr_rtf_sync_body_crc, :pt_long},
    0x1007 => {:pr_rtf_sync_body_count, :pt_long},
    0x1008 => {:pr_rtf_sync_body_tag, :pt_tstring},
    0x1009 => {:pr_rtf_compressed, :pt_binary},
    0x1010 => {:pr_rtf_sync_prefix_count, :pt_long},
    0x1011 => {:pr_rtf_sync_trailing_count, :pt_long},
    0x1013 => {:pr_body_html, :pt_binary},
    0x1014 => {:pr_body_html2, :pt_tstring},
    0x1035 => {:pr_internet_message_id, :pt_tstring},
    0x1039 => {:pr_internet_references, :pt_tstring},
    0x1042 => {:pr_in_reply_to_id, :pt_tstring},
    0x1043 => {:pr_list_help, :pt_tstring},
    0x1044 => {:pr_list_subscribe, :pt_tstring},
    0x1045 => {:pr_list_unsubscribe, :pt_tstring},
    0x1046 => {:pr_internet_return_path, :pt_tstring},
    0x3001 => {:pr_display_name, :pt_tstring},
    0x3002 => {:pr_addrtype, :pt_tstring},
    0x3003 => {:pr_email_address, :pt_tstring},
    0x3004 => {:pr_comment, :pt_tstring},
    0x3005 => {:pr_depth, :pt_long},
    0x3007 => {:pr_creation_time, :pt_systime},
    0x3008 => {:pr_last_modification_time, :pt_systime},
    0x300B => {:pr_search_key, :pt_binary},
    0x3010 => {:pr_target_entryid, :pt_binary},
    0x35E0 => {:pr_ipm_subtree_entryid, :pt_binary},
    0x35E2 => {:pr_ipm_outbox_entryid, :pt_binary},
    0x35E3 => {:pr_ipm_wastebasket_entryid, :pt_binary},
    0x35E4 => {:pr_ipm_sentmail_entryid, :pt_binary},
    0x35E5 => {:pr_views_entryid, :pt_binary},
    0x35E6 => {:pr_common_views_entryid, :pt_binary},
    0x35E7 => {:pr_finder_entryid, :pt_binary},
    0x3600 => {:pr_container_flags, :pt_long},
    0x3601 => {:pr_folder_type, :pt_long},
    0x3602 => {:pr_content_count, :pt_long},
    0x3603 => {:pr_content_unread, :pt_long},
    0x3610 => {:pr_subfolders, :pt_boolean},
    0x3613 => {:pr_container_class, :pt_tstring},
    0x3701 => {:pr_attach_data_bin, :pt_binary},
    0x3702 => {:pr_attach_encoding, :pt_binary},
    0x3703 => {:pr_attach_extension, :pt_tstring},
    0x3704 => {:pr_attach_filename, :pt_tstring},
    0x3705 => {:pr_attach_method, :pt_long},
    0x3707 => {:pr_attach_long_filename, :pt_tstring},
    0x3708 => {:pr_attach_pathname, :pt_tstring},
    0x3709 => {:pr_attach_rendering, :pt_binary},
    0x370A => {:pr_attach_tag, :pt_binary},
    0x370B => {:pr_rendering_position, :pt_long},
    0x370E => {:pr_attach_mime_tag, :pt_tstring},
    0x3710 => {:pr_attach_mime_sequence, :pt_long},
    0x3712 => {:pr_attach_content_id, :pt_tstring},
    0x3713 => {:pr_attach_content_location, :pt_tstring},
    0x3716 => {:pr_attach_content_disposition, :pt_tstring},
    0x3714 => {:pr_attach_long_pathname, :pt_tstring},
    0x3900 => {:pr_display_type, :pt_long},
    0x3A00 => {:pr_account, :pt_tstring},
    0x3A02 => {:pr_callback_telephone_number, :pt_tstring},
    0x3A05 => {:pr_generation, :pt_tstring},
    0x3A06 => {:pr_given_name, :pt_tstring},
    0x3A08 => {:pr_business_telephone_number, :pt_tstring},
    0x3A09 => {:pr_home_telephone_number, :pt_tstring},
    0x3A0A => {:pr_initials, :pt_tstring},
    0x3A0B => {:pr_keyword, :pt_tstring},
    0x3A0C => {:pr_language, :pt_tstring},
    0x3A0D => {:pr_location, :pt_tstring},
    0x3A0F => {:pr_mhs_common_name, :pt_tstring},
    0x3A10 => {:pr_organizational_id_number, :pt_tstring},
    0x3A11 => {:pr_surname, :pt_tstring},
    0x3A12 => {:pr_original_entryid, :pt_binary},
    0x3A13 => {:pr_original_display_name, :pt_tstring},
    0x3A15 => {:pr_postal_address, :pt_tstring},
    0x3A16 => {:pr_company_name, :pt_tstring},
    0x3A17 => {:pr_title, :pt_tstring},
    0x3A18 => {:pr_department_name, :pt_tstring},
    0x3A19 => {:pr_office_location, :pt_tstring},
    0x3A1A => {:pr_primary_telephone_number, :pt_tstring},
    0x3A1B => {:pr_business2_telephone_number, :pt_tstring},
    0x3A1C => {:pr_mobile_telephone_number, :pt_tstring},
    0x3A1D => {:pr_radio_telephone_number, :pt_tstring},
    0x3A1E => {:pr_car_telephone_number, :pt_tstring},
    0x3A1F => {:pr_other_telephone_number, :pt_tstring},
    0x3A20 => {:pr_transmittable_display_name, :pt_tstring},
    0x3A21 => {:pr_pager_telephone_number, :pt_tstring},
    0x3A22 => {:pr_user_certificate, :pt_binary},
    0x3A23 => {:pr_primary_fax_number, :pt_tstring},
    0x3A24 => {:pr_business_fax_number, :pt_tstring},
    0x3A25 => {:pr_home_fax_number, :pt_tstring},
    0x3A26 => {:pr_country, :pt_tstring},
    0x3A27 => {:pr_locality, :pt_tstring},
    0x3A28 => {:pr_state_or_province, :pt_tstring},
    0x3A29 => {:pr_street_address, :pt_tstring},
    0x3A2A => {:pr_postal_code, :pt_tstring},
    0x3A2B => {:pr_post_office_box, :pt_tstring},
    0x3A2D => {:pr_telex_number, :pt_tstring},
    0x3A2E => {:pr_isdn_number, :pt_tstring},
    0x3A2F => {:pr_assistant_telephone_number, :pt_tstring},
    0x3A30 => {:pr_home2_telephone_number, :pt_tstring},
    0x3A40 => {:pr_send_rich_info, :pt_boolean},
    0x3A44 => {:pr_middle_name, :pt_tstring},
    0x3A45 => {:pr_display_name_prefix, :pt_tstring},
    0x3A46 => {:pr_profession, :pt_tstring},
    0x3A48 => {:pr_spouse_name, :pt_tstring},
    0x3A4F => {:pr_nickname, :pt_tstring},
    0x3A51 => {:pr_personal_home_page, :pt_tstring},
    0x3A57 => {:pr_company_main_phone_number, :pt_tstring},
    0x3A58 => {:pr_childrens_names, :pt_tstring},
    0x3A70 => {:pr_user_x509_certificate, :pt_binary},
    0x39FE => {:pr_smtp_address, :pt_tstring},
    0x3D01 => {:pr_ab_default_dir, :pt_binary},
    0x3FDE => {:pr_internet_cpid, :pt_long},
    0x3FF1 => {:pr_message_locale_id, :pt_long},
    0x3FFD => {:pr_message_codepage, :pt_long},
    0x403E => {:pr_org_email_addr, :pt_tstring},
    0x5902 => {:pr_internet_content, :pt_binary},
    0x5909 => {:pr_smtp_address, :pt_tstring},
    0x5D01 => {:pr_sender_smtp_address, :pt_tstring},
    0x5D02 => {:pr_sent_representing_smtp_address, :pt_tstring},
    0x5FDE => {:pr_recipient_order, :pt_long},
    0x5FF6 => {:pr_recipient_display_name, :pt_tstring},
    0x5FF7 => {:pr_recipient_entryid, :pt_binary},
    0x5FFF => {:pr_recipient_flags, :pt_long},
    0x6000 => {:pr_recipient_trackstatus, :pt_long},
    0x6001 => {:pr_recipient_trackstatus_time, :pt_systime},
    0x6619 => {:pr_user_entryid, :pt_binary},
    0x661B => {:pr_mailbox_owner_entryid, :pt_binary},
    0x6620 => {:pr_schedule_folder_entryid, :pt_binary},
    0x6622 => {:pr_ipmdrafts_entryid, :pt_binary},
    0x6627 => {:pr_additional_ren_entryids, :pt_binary},
    0x6635 => {:pr_pst_hidden_count, :pt_long},
    0x6636 => {:pr_pst_hidden_unread, :pt_long},
    0x6638 => {:pr_folder_webviewinfo, :pt_binary},
    0x6645 => {:pr_ipm_appointment_entryid, :pt_binary},
    0x6646 => {:pr_ipm_contact_entryid, :pt_binary},
    0x6647 => {:pr_ipm_journal_entryid, :pt_binary},
    0x6648 => {:pr_ipm_note_entryid, :pt_binary},
    0x6649 => {:pr_ipm_task_entryid, :pt_binary},
    0x65E0 => {:pr_source_key, :pt_binary},
    0x65E2 => {:pr_change_key, :pt_binary},
    0x65E3 => {:pr_predecessor_change_list, :pt_binary},
    0x66A1 => {:pr_locale_id, :pt_long},
    0x67F2 => {:pr_ltp_row_id, :pt_long},
    0x67F3 => {:pr_ltp_row_ver, :pt_long},
    0x6800 => {:pr_pst_password, :pt_long}
  }

  # Build the reverse lookup map at compile time: name atom -> integer code
  @names_to_codes @tags
                  |> Enum.sort_by(fn {code, _} -> code end)
                  |> Enum.reduce(%{}, fn {code, {name, _type}}, acc ->
                    # Some aliases share the same property name; keep the lowest code as canonical.
                    Map.put_new(acc, name, code)
                  end)

  # -------------------------------------------------------------------
  # Public API
  # -------------------------------------------------------------------

  @doc """
  Looks up a MAPI property tag by its integer code.

  Returns a `{name, type}` tuple where `name` is the symbolic property name
  atom (e.g. `:pr_subject`) and `type` is the property type atom
  (e.g. `:pt_tstring`), or `nil` if the code is not recognized.

  ## Examples

      iex> OutlookMsg.Mapi.Tags.lookup(0x0037)
      {:pr_subject, :pt_tstring}

      iex> OutlookMsg.Mapi.Tags.lookup(0x0017)
      {:pr_importance, :pt_long}

      iex> OutlookMsg.Mapi.Tags.lookup(0xFFFF)
      nil
  """
  @spec lookup(non_neg_integer()) :: {atom(), atom()} | nil
  def lookup(code) when is_integer(code) do
    Map.get(@tags, code)
  end

  @doc """
  Returns the symbolic name atom for a MAPI property tag code, or `nil`
  if the code is not recognized.

  ## Examples

      iex> OutlookMsg.Mapi.Tags.name(0x0037)
      :pr_subject

      iex> OutlookMsg.Mapi.Tags.name(0x001A)
      :pr_message_class

      iex> OutlookMsg.Mapi.Tags.name(0xFFFF)
      nil
  """
  @spec name(non_neg_integer()) :: atom() | nil
  def name(code) when is_integer(code) do
    case Map.get(@tags, code) do
      {name, _type} -> name
      nil -> nil
    end
  end

  @doc """
  Reverse-looks up a MAPI property tag code by its symbolic name atom.

  Returns the integer code or `nil` if the name is not recognized.

  ## Examples

      iex> OutlookMsg.Mapi.Tags.code(:pr_subject)
      0x0037

      iex> OutlookMsg.Mapi.Tags.code(:pr_message_class)
      0x001A

      iex> OutlookMsg.Mapi.Tags.code(:pr_nonexistent)
      nil
  """
  @spec code(atom()) :: non_neg_integer() | nil
  def code(name) when is_atom(name) do
    Map.get(@names_to_codes, name)
  end
end
