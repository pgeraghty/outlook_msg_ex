import Bitwise

prebuf = "{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx"

# Initialize buffer
buf = prebuf
  |> :binary.bin_to_list()
  |> Enum.with_index()
  |> Enum.reduce(%{}, fn {byte, idx}, buf -> Map.put(buf, idx, byte) end)

wp = 207
output = []

# Payload bytes
payload = [3, 0, 10, 0, 114, 99, 112, 10, 96, 54, 50, 49, 48, 48, 53, 10, 15, 9, 8, 96, 224, 75, 7, 4, 176, 48, 104, 101, 108, 108, 111, 125, 10, 128, 15, 160]

IO.puts("Flag byte: 0x03 = #{Integer.to_string(3, 2)}")
IO.puts("")

# Manual trace: flag = 0x03
# Bit 0 = 1 (back-ref): bytes [0x00, 0x0A] = 0x000A
#   offset = 0x000A >>> 4 = 0, length = (0x000A &&& 0x0F) + 2 = 12
val1 = 0x000A
offset1 = val1 >>> 4
length1 = (val1 &&& 0x0F) + 2
IO.puts("Token 1: back-ref val=0x#{Integer.to_string(val1, 16)}, offset=#{offset1}, length=#{length1}")
chars1 = for i <- 0..(length1-1), do: Map.get(buf, (offset1 + i) &&& 4095, 0)
IO.puts("  Copies: #{inspect(chars1)} = #{inspect(:binary.list_to_bin(chars1))}")

# Bit 1 = 1 (back-ref): bytes [0x00, 0x72] = 0x0072
val2 = 0x0072
offset2 = val2 >>> 4
length2 = (val2 &&& 0x0F) + 2
IO.puts("Token 2: back-ref val=0x#{Integer.to_string(val2, 16)}, offset=#{offset2}, length=#{length2}")
# After token 1, wp should be 207+12=219, and buf[207..218] = same as prebuf[0..11]
# But offset2=7 which is in the prebuf area still
chars2 = for i <- 0..(length2-1), do: Map.get(buf, (offset2 + i) &&& 4095, 0)
IO.puts("  Copies: #{inspect(chars2)} = #{inspect(:binary.list_to_bin(chars2))}")

# Bit 2 = 0 (literal): byte 0x63 = 'c'
IO.puts("Token 3: literal 0x63 = 'c'")

# Bit 3 = 0 (literal): byte 0x70 = 'p'
IO.puts("Token 4: literal 0x70 = 'p'")

# Bit 4 = 0 (literal): byte 0x0a ???
# Expected: 'g' = 0x67
IO.puts("Token 5: literal 0x0A = '\\n' --- BUT expected 'g' = 0x67!")

# WAIT. Let me re-examine the payload hex.
# The hex string from the spec test is:
# 03 00 0a 00 72 63 70 0a 60 36 32 31 30 30 35 ...
# But maybe my hex encoding from the spec is wrong.
# Let me look at the actual MS-OXRTFCP spec example.

# Actually, the hex I built may be wrong. Let me try parsing it differently.
# The MS-OXRTFCP spec example compresses "{\rtf1\ansi\ansicpg1252\pard hello}\r\n"
# which is 43 bytes.

# Maybe the correct compressed payload from the spec is different.
# Let me try the known good test from python/other implementations.

IO.puts("\n\nActual expected result: {\\rtf1\\ansi\\ansicpg1252\\pard hello}\\r\\n")
IO.puts("That's #{byte_size("{\\rtf1\\ansi\\ansicpg1252\\pard hello}\r\n")} bytes")

# Actually, let me check if the flag byte processing order is wrong.
# Some implementations process MSB first instead of LSB first.
# Flag 0x03 = 0b00000011
# LSB first: bit0=1(ref), bit1=1(ref), bit2-7=0(literal)
# MSB first: bit7=0(lit), bit6=0(lit), ..., bit1=1(ref), bit0=1(ref)

# Let me check if MSB first would work:
# MSB first for 0x03: bits are 0,0,0,0,0,0,1,1
# First 6 tokens literal, last 2 back-references
# Tokens: literal 0x00, literal 0x0A, literal 0x00, literal 0x72, literal 0x63, literal 0x70
# That gives us: \x00 \n \x00 r c p
# That doesn't work either.

# Actually wait, some implementations use flag byte where bit 0 is the FIRST token.
# In LSB-first, bit position 0 is checked first. Let me verify my code does this.
# The spec says: "For each bit in the flag byte, starting from the least significant bit"
# So bit 0 first, then bit 1, etc.

# Hmm, but perhaps 0=ref, 1=literal (inverted from what I have)?
# In the MS spec: 0=reference, 1=literal? Let me check.
# Actually I may have the 0 and 1 meanings swapped!
# MS-OXRTFCP says:
# "if bit is 0, the token is a literal"
# "if bit is 1, the token is a reference"
# That's what I have. But let me try swapping to see...

# With SWAPPED: 0=reference, 1=literal
# Flag 0x03: bit0=1(literal), bit1=1(literal), bit2-7=0(reference)
# Token 1 (literal): byte 0x00
# Token 2 (literal): byte 0x0A
# Tokens 3-8 (all references):
#   Token 3: bytes 0x00 0x72 -> 0x0072 -> offset=7, len=4 -> "ansi"
# Hmm that doesn't make sense either.

# Let me look at this from the other direction. If the expected output is
# {\rtf1\ansi\ansicpg1252\pard hello}\r\n
# and the prebuf starts with {\rtf1\ansi\
# Then the first token should give us {\rtf1\ansi\  (11 bytes)
# or {\rtf1\ansi\a (12 bytes, including the \a from \ansi)

# A back-reference with offset=0, length=12 gives prebuf[0..11]:
# { \ r t f 1 \ a n s i \ = "{\rtf1\ansi\"
# That's 12 bytes. Then we need "ansicpg1252\pard hello}\r\n" = 26 more bytes.

# The next back-ref offset=7, length=4 gives buf[7..10] = "ansi"
# Total: "{\rtf1\ansi\ansi" (16 bytes).
# Then literals: c, p -- total "{\rtf1\ansi\ansicpg" (18 bytes).
# Then we need "1252\pard hello}\r\n" (19 more bytes).
# But the next byte in the stream is 0x0A, which is not a literal for 'g'.

# UNLESS... token 5 is not at byte offset 7. Let me recount bytes consumed.
# Flag: byte[0] = 0x03
# Token 1 (ref, 2 bytes): bytes[1,2] = 0x00, 0x0A
# Token 2 (ref, 2 bytes): bytes[3,4] = 0x00, 0x72
# Token 3 (literal, 1 byte): byte[5] = 0x63 = 'c'
# Token 4 (literal, 1 byte): byte[6] = 0x70 = 'p'
# Token 5 (literal, 1 byte): byte[7] = 0x0A -- this is line feed, not 'g'!

# Something is off with the test vector. Let me re-check the hex.
# Actually maybe I transcribed the compressed data wrong from the spec.

# The MS-OXRTFCP Section 4.1 gives this compressed stream (hex):
# Let me use the canonical example that many implementations test against.
# This is from the pyrtfcomp and other reference implementations.
