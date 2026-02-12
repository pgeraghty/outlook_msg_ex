defmodule OutlookMsg.Mapi.PropertySetTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Mapi.{PropertySet, Key, Tags, Guids}

  describe "PropertySet" do
    test "get by atom name" do
      key = Key.new(0x0037)  # PR_SUBJECT
      ps = PropertySet.new(%{key => "Hello World"})
      assert PropertySet.get(ps, :pr_subject) == "Hello World"
    end

    test "get by atom name ignores string-named custom properties" do
      key = Key.new(0x0037)  # PR_SUBJECT
      custom = %Key{code: "AntispamSafeLinksWrappedUrls", guid: Guids.ps_public_strings()}
      ps = PropertySet.new(%{custom => "custom", key => "Hello World"})
      assert PropertySet.get(ps, :pr_subject) == "Hello World"
    end

    test "get by integer code" do
      key = Key.new(0x0037)
      ps = PropertySet.new(%{key => "Test Subject"})
      assert PropertySet.get(ps, 0x0037) == "Test Subject"
    end

    test "returns nil for missing property" do
      ps = PropertySet.new(%{})
      assert PropertySet.get(ps, :pr_subject) == nil
      assert PropertySet.get(ps, 0x0037) == nil
    end

    test "subject convenience function" do
      key = Key.new(0x0037)
      ps = PropertySet.new(%{key => "My Subject"})
      assert PropertySet.subject(ps) == "My Subject"
    end

    test "body convenience function" do
      key = Key.new(0x1000)  # PR_BODY
      ps = PropertySet.new(%{key => "Hello body"})
      assert PropertySet.body(ps) == "Hello body"
    end

    test "sender_name convenience function" do
      key = Key.new(0x0C1A)  # PR_SENDER_NAME
      ps = PropertySet.new(%{key => "John Doe"})
      assert PropertySet.sender_name(ps) == "John Doe"
    end

    test "sender_email convenience function" do
      key = Key.new(0x0C1F)  # PR_SENDER_EMAIL_ADDRESS
      ps = PropertySet.new(%{key => "john@example.com"})
      assert PropertySet.sender_email(ps) == "john@example.com"
    end

    test "message_class convenience function" do
      key = Key.new(0x001A)  # PR_MESSAGE_CLASS
      ps = PropertySet.new(%{key => "IPM.Note"})
      assert PropertySet.message_class(ps) == "IPM.Note"
    end

    test "importance returns atom for known values" do
      key = Key.new(0x0017)  # PR_IMPORTANCE
      assert PropertySet.importance(PropertySet.new(%{key => 0})) == :low
      assert PropertySet.importance(PropertySet.new(%{key => 1})) == :normal
      assert PropertySet.importance(PropertySet.new(%{key => 2})) == :high
      assert PropertySet.importance(PropertySet.new(%{})) == nil
    end

    test "sensitivity returns atom for known values" do
      key = Key.new(0x0036)  # PR_SENSITIVITY
      assert PropertySet.sensitivity(PropertySet.new(%{key => 0})) == :none
      assert PropertySet.sensitivity(PropertySet.new(%{key => 1})) == :personal
      assert PropertySet.sensitivity(PropertySet.new(%{key => 2})) == :private
      assert PropertySet.sensitivity(PropertySet.new(%{key => 3})) == :confidential
      assert PropertySet.sensitivity(PropertySet.new(%{})) == nil
    end

    test "keys/1 returns all keys" do
      k1 = Key.new(0x0037)
      k2 = Key.new(0x1000)
      ps = PropertySet.new(%{k1 => "a", k2 => "b"})
      keys = PropertySet.keys(ps)
      assert length(keys) == 2
    end

    test "values/1 returns all values" do
      k1 = Key.new(0x0037)
      k2 = Key.new(0x1000)
      ps = PropertySet.new(%{k1 => "subj", k2 => "body"})
      vals = PropertySet.values(ps)
      assert Enum.sort(vals) == ["body", "subj"]
    end

    test "to_map/1 returns the raw properties map" do
      k1 = Key.new(0x0037)
      props = %{k1 => "test"}
      ps = PropertySet.new(props)
      assert PropertySet.to_map(ps) == props
    end

    test "to_symbolic_map/1 resolves key names" do
      key = Key.new(0x0037)
      ps = PropertySet.new(%{key => "Test"})
      sym_map = PropertySet.to_symbolic_map(ps)
      assert sym_map[:pr_subject] == "Test"
    end

    test "Access behaviour - fetch" do
      key = Key.new(0x0037)
      ps = PropertySet.new(%{key => "Hello"})
      assert ps[:pr_subject] == "Hello"
      assert ps[0x0037] == "Hello"
    end

    test "get by code and guid" do
      guid = Guids.ps_mapi()
      key = Key.new(0x0037, guid)
      ps = PropertySet.new(%{key => "Exact lookup"})
      assert PropertySet.get(ps, 0x0037, guid) == "Exact lookup"
    end
  end

  describe "Key" do
    test "new/1 creates key with PS_MAPI guid" do
      key = Key.new(0x0037)
      assert key.code == 0x0037
      assert key.guid == Guids.ps_mapi()
    end

    test "new/2 creates key with custom guid" do
      guid = Guids.psetid_common()
      key = Key.new(0x8501, guid)
      assert key.code == 0x8501
      assert key.guid == guid
    end

    test "to_sym/1 resolves standard MAPI properties" do
      key = Key.new(0x0037)
      assert Key.to_sym(key) == :pr_subject
    end

    test "to_sym/1 resolves named properties" do
      key = Key.new(0x8503, Guids.psetid_common())
      assert Key.to_sym(key) == :reminder_set
    end

    test "to_sym/1 returns nil for unknown" do
      key = Key.new(0xFFFF)
      assert Key.to_sym(key) == nil
    end

    test "to_sym/1 returns nil for string-named properties" do
      key = Key.new("AntispamSafeLinksWrappedUrls", Guids.ps_public_strings())
      assert Key.to_sym(key) == nil
    end

    test "to_string/1 returns name for known property" do
      key = Key.new(0x0037)
      assert Key.to_string(key) == "pr_subject"
    end

    test "to_string/1 returns hex for unknown property" do
      key = Key.new(0xBEEF)
      assert Key.to_string(key) == "beef"
    end

    test "symbolic?/1 returns true for known key" do
      key = Key.new(0x0037)
      assert Key.symbolic?(key)
    end

    test "symbolic?/1 returns false for unknown key" do
      key = Key.new(0xFFFF)
      refute Key.symbolic?(key)
    end
  end

  describe "Tags" do
    test "lookup/1 finds PR_SUBJECT" do
      assert {:pr_subject, :pt_tstring} = Tags.lookup(0x0037)
    end

    test "lookup/1 returns nil for unknown" do
      assert Tags.lookup(0xDEAD) == nil
    end

    test "name/1 returns atom name" do
      assert Tags.name(0x0037) == :pr_subject
    end

    test "name/1 returns nil for unknown" do
      assert Tags.name(0xDEAD) == nil
    end

    test "code/1 reverse lookup" do
      assert Tags.code(:pr_subject) == 0x0037
    end

    test "code/1 returns nil for unknown name" do
      assert Tags.code(:pr_nonexistent) == nil
    end

    test "lookup for common properties" do
      assert {:pr_body, :pt_tstring} = Tags.lookup(0x1000)
      assert {:pr_sender_name, :pt_tstring} = Tags.lookup(0x0C1A)
      assert {:pr_message_class, :pt_tstring} = Tags.lookup(0x001A)
      assert {:pr_recipient_type, :pt_long} = Tags.lookup(0x0C15)
      assert {:pr_attach_long_filename, :pt_tstring} = Tags.lookup(0x3707)
      assert {:pr_attach_data_bin, :pt_binary} = Tags.lookup(0x3701)
    end
  end

  describe "Guids" do
    test "ps_mapi returns 16-byte binary" do
      guid = Guids.ps_mapi()
      assert byte_size(guid) == 16
    end

    test "name/1 resolves known GUID" do
      assert Guids.name(Guids.ps_mapi()) == :ps_mapi
    end

    test "name/1 returns :unknown for unknown GUID" do
      assert Guids.name(<<0::128>>) == :unknown
    end

    test "all GUIDs are 16 bytes" do
      for {guid, _name} <- Guids.all() do
        assert byte_size(guid) == 16
      end
    end

    test "well-known GUIDs have correct names" do
      assert Guids.name(Guids.ps_mapi()) == :ps_mapi
      assert Guids.name(Guids.ps_public_strings()) == :ps_public_strings
      assert Guids.name(Guids.psetid_common()) == :psetid_common
      assert Guids.name(Guids.psetid_address()) == :psetid_address
      assert Guids.name(Guids.psetid_appointment()) == :psetid_appointment
      assert Guids.name(Guids.psetid_task()) == :psetid_task
    end

    test "all/0 returns a map of expected size" do
      all = Guids.all()
      assert is_map(all)
      # At least the core GUIDs should be present
      assert map_size(all) >= 10
    end
  end
end
