defmodule OutlookMsg.PstItemTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Mapi.{Key, PropertySet}
  alias OutlookMsg.Pst.Item

  defp props(map) do
    PropertySet.new(map)
  end

  describe "detect_type/1" do
    test "classifies well-known message classes" do
      assert Item.detect_type(props(%{Key.new(0x001A) => "IPM.Note"})) == :message
      assert Item.detect_type(props(%{Key.new(0x001A) => "IPM.Contact"})) == :contact
      assert Item.detect_type(props(%{Key.new(0x001A) => "IPM.Task"})) == :task
      assert Item.detect_type(props(%{Key.new(0x001A) => "IPM.StickyNote"})) == :note
      assert Item.detect_type(props(%{Key.new(0x001A) => "IPM.Appointment"})) == :appointment
    end

    test "falls back to folder when folder indicators exist" do
      assert Item.detect_type(props(%{Key.new(0x3602) => 4})) == :folder
      assert Item.detect_type(props(%{Key.new(0x3610) => true})) == :folder
    end

    test "defaults to message when class is missing and no folder indicator exists" do
      assert Item.detect_type(props(%{})) == :message
    end
  end

  describe "predicates and helpers" do
    test "folder?/1 and message?/1 reflect type" do
      assert Item.folder?(%Item{type: :folder})
      assert Item.folder?(%Item{type: :root})
      refute Item.folder?(%Item{type: :message})

      assert Item.message?(%Item{type: :message})
      refute Item.message?(%Item{type: :contact})
    end

    test "display_name/1 prefers display name then subject" do
      i1 = %Item{properties: props(%{Key.new(0x3001) => "Inbox", Key.new(0x0037) => "Fallback"})}
      i2 = %Item{properties: props(%{Key.new(0x0037) => "Subject only"})}
      i3 = %Item{properties: props(%{})}

      assert Item.display_name(i1) == "Inbox"
      assert Item.display_name(i2) == "Subject only"
      assert Item.display_name(i3) == ""
    end
  end
end
