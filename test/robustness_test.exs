defmodule OutlookMsg.RobustnessTest do
  use ExUnit.Case, async: true
  import Bitwise

  @msg_fixtures Path.expand("fixtures/public_msg/*.msg", __DIR__)
  @eml_fixtures Path.expand("fixtures/public_eml/*.eml", __DIR__)

  test "open_with_warnings/1 never raises on corrupted MSG binaries" do
    fixtures =
      Path.wildcard(@msg_fixtures)
      |> Enum.sort()
      |> Enum.take(4)

    assert fixtures != []

    Enum.each(fixtures, fn path ->
      data = File.read!(path)

      for shift <- [13, 257, 1021] do
        mutated = mutate_binary(data, shift)

        result =
          try do
            OutlookMsg.open_with_warnings(mutated)
          rescue
            e -> {:raised, e}
          catch
            kind, reason -> {:thrown, {kind, reason}}
          end

        refute match?({:raised, _}, result), "raised on #{Path.basename(path)} shift=#{shift}: #{inspect(result)}"
        refute match?({:thrown, _}, result), "threw on #{Path.basename(path)} shift=#{shift}: #{inspect(result)}"
        assert match?({:ok, _, _}, result) or match?({:error, _}, result)
      end
    end)
  end

  test "open_eml_with_warnings/1 retains parse path and never raises on malformed EML" do
    fixtures = Path.wildcard(@eml_fixtures) |> Enum.sort()
    assert fixtures != []

    Enum.each(fixtures, fn path ->
      raw = File.read!(path)
      malformed = "BadHeaderWithoutColon\r\n" <> raw <> "\r\n=??broken"

      result =
        try do
          OutlookMsg.open_eml_with_warnings(malformed)
        rescue
          e -> {:raised, e}
        catch
          kind, reason -> {:thrown, {kind, reason}}
        end

      refute match?({:raised, _}, result), "raised on #{Path.basename(path)}: #{inspect(result)}"
      refute match?({:thrown, _}, result), "threw on #{Path.basename(path)}: #{inspect(result)}"
      assert {:ok, _mime, warnings} = result
      assert Enum.any?(warnings, &String.contains?(to_string(&1), "malformed header"))
    end)
  end

  test "open_pst_with_warnings/1 does not raise for random binary inputs" do
    data = :crypto.strong_rand_bytes(600)

    result =
      try do
        OutlookMsg.open_pst_with_warnings(data)
      rescue
        e -> {:raised, e}
      catch
        kind, reason -> {:thrown, {kind, reason}}
      end

    refute match?({:raised, _}, result)
    refute match?({:thrown, _}, result)
    assert match?({:ok, _, _}, result) or match?({:error, _}, result)
  end

  defp mutate_binary(data, shift) when is_binary(data) and is_integer(shift) do
    size = byte_size(data)
    if size == 0 do
      data
    else
      pos = rem(shift, size)
      <<prefix::binary-size(pos), b, rest::binary>> = data
      prefix <> <<bxor(b, 0x5A)>> <> rest
    end
  end
end
