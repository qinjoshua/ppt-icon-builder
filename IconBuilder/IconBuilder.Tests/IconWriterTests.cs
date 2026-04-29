using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Xunit;

namespace IconBuilder.Tests
{
    public class IconWriterTests : IDisposable
    {
        private readonly string _tempDir;

        public IconWriterTests()
        {
            _tempDir = Path.Combine(Path.GetTempPath(), "IconBuilderTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);
        }

        public void Dispose()
        {
            try { Directory.Delete(_tempDir, recursive: true); } catch { }
        }

        private string TempPath(string name) => Path.Combine(_tempDir, name);

        // -- helpers ------------------------------------------------------

        private static Bitmap MakeFilledBitmap(int width, int height, Color color)
        {
            var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(color);
            }
            return bmp;
        }

        private struct DirEntry
        {
            public int Width, Height, BitCount, ByteSize, Offset;
        }

        private static (int count, DirEntry[] entries, byte[] bytes) ReadIco(string path)
        {
            byte[] bytes = File.ReadAllBytes(path);
            using (var ms = new MemoryStream(bytes))
            using (var br = new BinaryReader(ms))
            {
                short reserved = br.ReadInt16();
                short type = br.ReadInt16();
                short count = br.ReadInt16();
                Assert.Equal(0, reserved);
                Assert.Equal(1, type);
                var entries = new DirEntry[count];
                for (int i = 0; i < count; i++)
                {
                    byte w = br.ReadByte();
                    byte h = br.ReadByte();
                    br.ReadByte(); // colorCount
                    br.ReadByte(); // reserved
                    short planes = br.ReadInt16();
                    short bits = br.ReadInt16();
                    int size = br.ReadInt32();
                    int off = br.ReadInt32();
                    Assert.Equal(1, planes);
                    entries[i] = new DirEntry
                    {
                        Width = w == 0 ? 256 : w,
                        Height = h == 0 ? 256 : h,
                        BitCount = bits,
                        ByteSize = size,
                        Offset = off
                    };
                }
                return (count, entries, bytes);
            }
        }

        private static byte[] Slice(byte[] src, int offset, int length)
        {
            var dst = new byte[length];
            Buffer.BlockCopy(src, offset, dst, 0, length);
            return dst;
        }

        // Read a single 32-bit BGRA pixel from a bottom-up ICO BMP payload at logical (x, y from top).
        private static (byte B, byte G, byte R, byte A) ReadBmpPixel(byte[] payload, int width, int height, int x, int yFromTop)
        {
            // BITMAPINFOHEADER is 40 bytes; pixels start at offset 40, bottom-up.
            int yFromBottom = height - 1 - yFromTop;
            int rowOffset = 40 + yFromBottom * width * 4;
            int px = rowOffset + x * 4;
            return (payload[px + 0], payload[px + 1], payload[px + 2], payload[px + 3]);
        }

        private static DirEntry FindEntry(DirEntry[] entries, int size)
        {
            foreach (var e in entries)
            {
                if (e.Width == size && e.Height == size) return e;
            }
            throw new InvalidOperationException($"No ICO entry of size {size}x{size} found.");
        }

        // --- tests --------------------------------------------------------

        [Fact]
        public void WritesFile_WithAllExpectedSizes_AndCorrectBitDepths()
        {
            string outPath = TempPath("a.ico");
            using (var src = MakeFilledBitmap(64, 64, Color.Red))
            {
                IconWriter.WriteIcoFromImage(src, outPath);
            }

            var (count, entries, bytes) = ReadIco(outPath);

            int[] expectedSizes = { 16, 20, 24, 32, 40, 48, 64, 256 };
            Assert.Equal(expectedSizes.Length, count);
            foreach (int s in expectedSizes)
            {
                var e = FindEntry(entries, s);
                Assert.Equal(32, e.BitCount);
            }

            int expectedOffset = 6 + 16 * count;
            for (int i = 0; i < count; i++)
            {
                Assert.Equal(expectedOffset, entries[i].Offset);
                Assert.True(entries[i].ByteSize > 0);
                Assert.True(entries[i].Offset + entries[i].ByteSize <= bytes.Length);
                expectedOffset += entries[i].ByteSize;
            }
            Assert.Equal(bytes.Length, expectedOffset);
        }

        [Fact]
        public void BmpEntries_HaveCorrectBitmapInfoHeader()
        {
            string outPath = TempPath("b.ico");
            using (var src = MakeFilledBitmap(64, 64, Color.Red))
            {
                IconWriter.WriteIcoFromImage(src, outPath);
            }

            var (_, entries, bytes) = ReadIco(outPath);

            int[] bmpSizes = { 16, 20, 24, 32, 40, 48, 64 };
            foreach (int s in bmpSizes)
            {
                var entry = FindEntry(entries, s);
                var payload = Slice(bytes, entry.Offset, entry.ByteSize);
                int biSize = BitConverter.ToInt32(payload, 0);
                int biWidth = BitConverter.ToInt32(payload, 4);
                int biHeight = BitConverter.ToInt32(payload, 8);
                short biPlanes = BitConverter.ToInt16(payload, 12);
                short biBitCount = BitConverter.ToInt16(payload, 14);
                int biCompression = BitConverter.ToInt32(payload, 16);

                Assert.Equal(40, biSize);
                Assert.Equal(entry.Width, biWidth);
                Assert.Equal(entry.Height * 2, biHeight); // doubled for AND mask
                Assert.Equal(1, biPlanes);
                Assert.Equal(32, biBitCount);
                Assert.Equal(0, biCompression);

                int maskRowBytes = ((entry.Width + 31) / 32) * 4;
                int expected = 40 + entry.Width * entry.Height * 4 + maskRowBytes * entry.Height;
                Assert.Equal(expected, entry.ByteSize);
            }
        }

        [Fact]
        public void PngEntry_StartsWithPngMagic_AndIs256x256()
        {
            string outPath = TempPath("c.ico");
            using (var src = MakeFilledBitmap(64, 64, Color.Red))
            {
                IconWriter.WriteIcoFromImage(src, outPath);
            }

            var (_, entries, bytes) = ReadIco(outPath);
            var pngEntry = FindEntry(entries, 256);
            var payload = Slice(bytes, pngEntry.Offset, pngEntry.ByteSize);

            byte[] sig = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
            for (int i = 0; i < sig.Length; i++) Assert.Equal(sig[i], payload[i]);

            int width = (payload[16] << 24) | (payload[17] << 16) | (payload[18] << 8) | payload[19];
            int height = (payload[20] << 24) | (payload[21] << 16) | (payload[22] << 8) | payload[23];
            Assert.Equal(256, width);
            Assert.Equal(256, height);
        }

        [Fact]
        public void IcoCanBeLoaded_ByGdiPlus_AsValidIcon()
        {
            string outPath = TempPath("d.ico");
            using (var src = MakeFilledBitmap(64, 64, Color.Red))
            {
                IconWriter.WriteIcoFromImage(src, outPath);
            }

            using (var icon = new Icon(outPath))
            {
                Assert.NotNull(icon);
            }

            using (var icon48 = new Icon(outPath, 48, 48))
            using (var bmp = icon48.ToBitmap())
            {
                Assert.Equal(48, bmp.Width);
                Assert.Equal(48, bmp.Height);
            }

            using (var icon16 = new Icon(outPath, 16, 16))
            using (var bmp = icon16.ToBitmap())
            {
                Assert.Equal(16, bmp.Width);
                Assert.Equal(16, bmp.Height);
            }

            using (var icon64 = new Icon(outPath, 64, 64))
            using (var bmp = icon64.ToBitmap())
            {
                Assert.Equal(64, bmp.Width);
                Assert.Equal(64, bmp.Height);
            }
        }

        [Fact]
        public void SquareSource_ProducesFullyOpaqueRedIcon()
        {
            string outPath = TempPath("e.ico");
            using (var src = MakeFilledBitmap(80, 80, Color.Red))
            {
                IconWriter.WriteIcoFromImage(src, outPath);
            }

            var (_, entries, bytes) = ReadIco(outPath);
            var entry = FindEntry(entries, 32);
            var payload = Slice(bytes, entry.Offset, entry.ByteSize);

            var (b, g, r, a) = ReadBmpPixel(payload, 32, 32, 16, 16);
            Assert.Equal(255, a);
            Assert.True(r > 240, $"center R should be ~255 but was {r}");
            Assert.True(g < 15, $"center G should be ~0 but was {g}");
            Assert.True(b < 15, $"center B should be ~0 but was {b}");

            // Corner alpha may be slightly < 255 due to bicubic edge bleed, but should still be mostly opaque.
            var corner = ReadBmpPixel(payload, 32, 32, 1, 1);
            Assert.True(corner.A > 200, $"corner alpha should be near opaque but was {corner.A}");
        }

        [Fact]
        public void WideSource_PreservesAspectRatio_WithTransparentTopBottomLetterbox()
        {
            // 100x50 red rectangle => aspect 2:1.
            // In a 32x32 square: image is 32 wide x 16 tall, centered vertically.
            string outPath = TempPath("f.ico");
            using (var src = MakeFilledBitmap(100, 50, Color.Red))
            {
                IconWriter.WriteIcoFromImage(src, outPath);
            }

            var (_, entries, bytes) = ReadIco(outPath);
            var entry = FindEntry(entries, 32);
            var payload = Slice(bytes, entry.Offset, entry.ByteSize);

            for (int y = 0; y < 6; y++)
            {
                for (int x = 0; x < 32; x++)
                {
                    var p = ReadBmpPixel(payload, 32, 32, x, y);
                    Assert.True(p.A == 0, $"Expected transparent at top ({x},{y}) but got A={p.A}");
                }
            }

            for (int y = 26; y < 32; y++)
            {
                for (int x = 0; x < 32; x++)
                {
                    var p = ReadBmpPixel(payload, 32, 32, x, y);
                    Assert.True(p.A == 0, $"Expected transparent at bottom ({x},{y}) but got A={p.A}");
                }
            }

            for (int y = 10; y < 22; y++)
            {
                for (int x = 4; x < 28; x++)
                {
                    var p = ReadBmpPixel(payload, 32, 32, x, y);
                    Assert.True(p.A > 240, $"Expected opaque at middle ({x},{y}) A={p.A}");
                    Assert.True(p.R > 200, $"Expected red at ({x},{y}) R={p.R}");
                    Assert.True(p.G < 30 && p.B < 30, $"Expected pure red ({x},{y}) G={p.G} B={p.B}");
                }
            }
        }

        [Fact]
        public void TallSource_PreservesAspectRatio_WithTransparentLeftRightLetterbox()
        {
            // 50x100 red => 1:2 aspect; in 48x48: 24x48 centered horizontally.
            string outPath = TempPath("g.ico");
            using (var src = MakeFilledBitmap(50, 100, Color.Red))
            {
                IconWriter.WriteIcoFromImage(src, outPath);
            }

            var (_, entries, bytes) = ReadIco(outPath);
            var entry = FindEntry(entries, 48);
            var payload = Slice(bytes, entry.Offset, entry.ByteSize);

            for (int y = 0; y < 48; y++)
            {
                for (int x = 0; x < 9; x++)
                {
                    var p = ReadBmpPixel(payload, 48, 48, x, y);
                    Assert.True(p.A == 0, $"Expected transparent at left ({x},{y}) A={p.A}");
                }
                for (int x = 39; x < 48; x++)
                {
                    var p = ReadBmpPixel(payload, 48, 48, x, y);
                    Assert.True(p.A == 0, $"Expected transparent at right ({x},{y}) A={p.A}");
                }
            }

            for (int y = 4; y < 44; y++)
            {
                for (int x = 16; x < 32; x++)
                {
                    var p = ReadBmpPixel(payload, 48, 48, x, y);
                    Assert.True(p.A > 240, $"Expected opaque at center ({x},{y}) A={p.A}");
                    Assert.True(p.R > 200, $"Expected red at ({x},{y}) R={p.R}");
                }
            }
        }

        [Fact]
        public void TransparentPaddingAroundShape_IsTrimmed_BeforeLetterboxing()
        {
            const int srcSize = 200;
            const int shapeW = 50;
            const int shapeH = 200;
            int xOffset = (srcSize - shapeW) / 2;

            using (var bmp = new Bitmap(srcSize, srcSize, PixelFormat.Format32bppArgb))
            {
                using (var g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.Transparent);
                    using (var brush = new SolidBrush(Color.Red))
                    {
                        g.FillRectangle(brush, xOffset, 0, shapeW, shapeH);
                    }
                }

                string outPath = TempPath("trim.ico");
                IconWriter.WriteIcoFromImage(bmp, outPath);

                var (_, entries, bytes) = ReadIco(outPath);
                var entry = FindEntry(entries, 48);
                var payload = Slice(bytes, entry.Offset, entry.ByteSize);

                for (int y = 0; y < 48; y++)
                {
                    for (int x = 0; x < 9; x++)
                    {
                        var p = ReadBmpPixel(payload, 48, 48, x, y);
                        Assert.True(p.A == 0, $"Expected transparent left margin at ({x},{y}) A={p.A}");
                    }
                    for (int x = 39; x < 48; x++)
                    {
                        var p = ReadBmpPixel(payload, 48, 48, x, y);
                        Assert.True(p.A == 0, $"Expected transparent right margin at ({x},{y}) A={p.A}");
                    }
                }
                for (int y = 4; y < 44; y++)
                {
                    for (int x = 18; x < 30; x++)
                    {
                        var p = ReadBmpPixel(payload, 48, 48, x, y);
                        Assert.True(p.A > 220, $"Expected opaque center at ({x},{y}) A={p.A}");
                        Assert.True(p.R > 200, $"Expected red center at ({x},{y}) R={p.R}");
                    }
                }
            }
        }

        [Fact]
        public void TransparentPaddingTopBottom_IsTrimmed_BeforeLetterboxing()
        {
            const int srcSize = 200;
            const int shapeW = 200;
            const int shapeH = 100;
            int yOffset = (srcSize - shapeH) / 2;

            using (var bmp = new Bitmap(srcSize, srcSize, PixelFormat.Format32bppArgb))
            {
                using (var g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.Transparent);
                    using (var brush = new SolidBrush(Color.Red))
                    {
                        g.FillRectangle(brush, 0, yOffset, shapeW, shapeH);
                    }
                }

                string outPath = TempPath("trim_wide.ico");
                IconWriter.WriteIcoFromImage(bmp, outPath);

                var (_, entries, bytes) = ReadIco(outPath);
                var entry = FindEntry(entries, 32);
                var payload = Slice(bytes, entry.Offset, entry.ByteSize);

                for (int y = 0; y < 6; y++)
                {
                    for (int x = 0; x < 32; x++)
                    {
                        var p = ReadBmpPixel(payload, 32, 32, x, y);
                        Assert.True(p.A == 0, $"Expected transparent top at ({x},{y}) A={p.A}");
                    }
                }
                for (int y = 26; y < 32; y++)
                {
                    for (int x = 0; x < 32; x++)
                    {
                        var p = ReadBmpPixel(payload, 32, 32, x, y);
                        Assert.True(p.A == 0, $"Expected transparent bottom at ({x},{y}) A={p.A}");
                    }
                }
                for (int y = 10; y < 22; y++)
                {
                    for (int x = 4; x < 28; x++)
                    {
                        var p = ReadBmpPixel(payload, 32, 32, x, y);
                        Assert.True(p.A > 240, $"Expected opaque middle at ({x},{y}) A={p.A}");
                    }
                }
            }
        }

        [Fact]
        public void RowOrientation_TopOfImage_IsTopOfRenderedIcon()
        {
            const int srcSize = 64;
            using (var bmp = new Bitmap(srcSize, srcSize, PixelFormat.Format32bppArgb))
            {
                using (var g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.Transparent);
                    using (var topBrush = new SolidBrush(Color.Blue))
                    using (var botBrush = new SolidBrush(Color.Lime))
                    {
                        g.FillRectangle(topBrush, 0, 0, srcSize, srcSize / 2);
                        g.FillRectangle(botBrush, 0, srcSize / 2, srcSize, srcSize / 2);
                    }
                }

                string outPath = TempPath("h.ico");
                IconWriter.WriteIcoFromImage(bmp, outPath);

                var (_, entries, bytes) = ReadIco(outPath);
                var entry = FindEntry(entries, 48);
                var payload = Slice(bytes, entry.Offset, entry.ByteSize);

                var top = ReadBmpPixel(payload, 48, 48, 24, 4);
                Assert.True(top.B > 200 && top.R < 30 && top.G < 30,
                    $"Expected top to be blue but got R={top.R} G={top.G} B={top.B}");

                var bot = ReadBmpPixel(payload, 48, 48, 24, 43);
                Assert.True(bot.G > 200 && bot.R < 30 && bot.B < 30,
                    $"Expected bottom to be green but got R={bot.R} G={bot.G} B={bot.B}");
            }
        }
    }
}
