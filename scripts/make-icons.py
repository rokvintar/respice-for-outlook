"""Generates the add-in icons (blue rounded square with a white 'eye'). Pure stdlib."""
import struct, zlib, math, os, sys

def png(path, size):
    S = 4  # supersampling
    rows = []
    for y in range(size):
        row = bytearray([0])
        for x in range(size):
            acc = [0, 0, 0, 0]
            for sy in range(S):
                for sx in range(S):
                    u = (x + (sx + .5) / S) / size * 2 - 1
                    v = (y + (sy + .5) / S) / size * 2 - 1
                    r = .32
                    qx, qy = max(abs(u) - (1 - r), 0), max(abs(v) - (1 - r), 0)
                    if math.hypot(qx, qy) > r:
                        continue
                    col = (15, 108, 189)
                    # eye: lens = intersection of two circles, pupil in the middle
                    d = .9
                    lens = math.hypot(u, v - d) < 1.25 and math.hypot(u, v + d) < 1.25
                    if lens:
                        col = (255, 255, 255)
                        if math.hypot(u, v) < .24:
                            col = (15, 108, 189)
                    acc[0] += col[0]; acc[1] += col[1]; acc[2] += col[2]; acc[3] += 255
            n = S * S
            a = acc[3] // n
            if acc[3]:
                k = acc[3] // 255
                row += bytes([acc[0] // k, acc[1] // k, acc[2] // k, a])
            else:
                row += bytes([0, 0, 0, 0])
        rows.append(bytes(row))
    raw = zlib.compress(b"".join(rows), 9)
    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xffffffff)
    data = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)) + chunk(b"IDAT", raw) + chunk(b"IEND", b"")
    open(path, "wb").write(data)

out = sys.argv[1] if len(sys.argv) > 1 else "assets"
for s in (16, 32, 64, 80, 128):
    png(os.path.join(out, f"icon-{s}.png"), s)
print("icons written to", out)
