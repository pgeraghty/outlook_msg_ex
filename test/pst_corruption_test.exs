defmodule OutlookMsg.PstCorruptionTest do
  use ExUnit.Case, async: true

  alias OutlookMsg.Pst.{Header, Index, Descriptor}

  defp header97(index1, index2) do
    %Header{
      version: :pst97,
      index1_offset: index1,
      index2_offset: index2
    }
  end

  defp make_page(item_count, level, entries) do
    body = entries <> :binary.copy(<<0>>, max(0, 496 - byte_size(entries)))
    meta = <<item_count::8, 0::8, 0::8, level::8>>
    body <> meta <> :binary.copy(<<0>>, 512 - byte_size(body) - byte_size(meta))
  end

  test "Index.load handles branch loops without recursion blowup" do
    # Branch at offset 512 points back to itself.
    branch_entry = <<1::32-little, 512::32-little, 0::32-little>>
    branch_page = make_page(1, 1, branch_entry)
    data = :binary.copy(<<0>>, 512) <> branch_page

    result =
      try do
        Index.load(data, header97(512, 0))
      rescue
        e -> {:raised, e}
      end

    refute match?({:raised, _}, result)
    assert is_map(result)

    {_, warnings} = Index.load_with_warnings(data, header97(512, 0))
    assert Enum.any?(warnings, &match?(%OutlookMsg.Warning{code: :pst_branch_loop_detected}, &1))
  end

  test "Descriptor.load handles branch loops without recursion blowup" do
    branch_entry = <<1::32-little, 512::32-little, 0::32-little>>
    branch_page = make_page(1, 1, branch_entry)
    data = :binary.copy(<<0>>, 512) <> branch_page

    result =
      try do
        Descriptor.load(data, header97(0, 512))
      rescue
        e -> {:raised, e}
      end

    refute match?({:raised, _}, result)
    assert is_map(result)

    {_, warnings} = Descriptor.load_with_warnings(data, header97(0, 512))
    assert Enum.any?(warnings, &match?(%OutlookMsg.Warning{code: :pst_branch_loop_detected}, &1))
  end

  test "Index.load ignores oversized item_count when records are truncated" do
    # One valid leaf record but item_count claims many.
    leaf_entry = <<5::32-little, 1234::32-little, 99::16-little, 0::16-little>>
    leaf_page = make_page(30, 0, leaf_entry)
    data = :binary.copy(<<0>>, 512) <> leaf_page

    idx = Index.load(data, header97(512, 0))
    assert is_map(idx)
    assert Map.has_key?(idx, 5)
  end

  test "Descriptor.load ignores oversized item_count when records are truncated" do
    leaf_entry = <<33::32-little, 7::32-little, 0::32-little, 0::32-little>>
    leaf_page = make_page(30, 0, leaf_entry)
    data = :binary.copy(<<0>>, 512) <> leaf_page

    desc = Descriptor.load(data, header97(0, 512))
    assert is_map(desc)
    assert Map.has_key?(desc, 33)
  end

  test "fanout branch with mixed valid/invalid child offsets still salvages valid subtree" do
    leaf_entry = <<9::32-little, 2048::32-little, 16::16-little, 0::16-little>>
    leaf_page = make_page(1, 0, leaf_entry)
    branch_entries = <<
      1::32-little, 1024::32-little, 0::32-little,
      2::32-little, 99_999::32-little, 0::32-little
    >>
    branch_page = make_page(2, 1, branch_entries)
    data = :binary.copy(<<0>>, 512) <> branch_page <> leaf_page

    idx = Index.load(data, header97(512, 0))
    assert is_map(idx)
    assert Map.has_key?(idx, 9)
  end
end
