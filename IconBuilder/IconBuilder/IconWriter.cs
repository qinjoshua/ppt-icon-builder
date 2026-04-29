using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

namespace IconBuilder
{
    internal static class IconWriter
    {
        // ICO entry sizes (ascending). 256 is stored as a PNG; smaller entries as 32-bit BMP.
        public static readonly int[] BmpSizes = new[] { 16, 20, 24, 32, 40, 48, 64 };
        public const int PngSize = 256;

        public static void WriteIcoFromImage(Image source, string outputPath)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));

            // Trim any fully-transparent border so the visible content's aspect ratio drives the
            // letterboxing. Without this, a shape exported by PowerPoint with surrounding padding
            // would be centered relative to the padded bounds rather than its visible content.
            using (var trimmed = TrimTransparentBorder(source))
            {
                var entries = new System.Collections.Generic.List<IconEntry>(BmpSizes.Length + 1);
                foreach (int size in BmpSizes)
                {
                    using (var bmp = ResizeTo32bpp(trimmed, size))
                    {
                        byte[] data = BuildBmpIconData(bmp);
                        entries.Add(new IconEntry(size, size, 32, data, isPng: false));
                    }
                }
                using (var bmpPng = ResizeTo32bpp(trimmed, PngSize))
                {
                    byte[] pngData = BuildPngData(bmpPng);
                    entries.Add(new IconEntry(PngSize, PngSize, 32, pngData, isPng: true));
                }

                WriteIco(outputPath, entries.ToArray());
            }
        }

        // Writes an ICO file using a pre-rendered Bitmap per size slot. For any size that isn't
        // present in `sources`, the largest available source is scaled down to fill the gap.
        // Each source bitmap should ideally already be exactly its size×size 32bpp ARGB content;
        // if not, it is letterbox-resized first.
        public static void WriteIcoFromSizedSources(IDictionary<int, Bitmap> sources, string outputPath)
        {
            if (sources == null || sources.Count == 0)
                throw new ArgumentException("At least one slot must be assigned.", nameof(sources));

            // Use the largest filled slot as the fallback source for missing sizes.
            var keys = new System.Collections.Generic.List<int>(sources.Keys);
            keys.Sort((a, b) => b.CompareTo(a));
            Bitmap fallback = sources[keys[0]];

            var entries = new System.Collections.Generic.List<IconEntry>(BmpSizes.Length + 1);

            foreach (int size in BmpSizes)
            {
                using (var rendered = GetOrRender(sources, fallback, size))
                {
                    byte[] data = BuildBmpIconData(rendered);
                    entries.Add(new IconEntry(size, size, 32, data, isPng: false));
                }
            }

            using (var pngBmp = GetOrRender(sources, fallback, PngSize))
            {
                byte[] pngData = BuildPngData(pngBmp);
                entries.Add(new IconEntry(PngSize, PngSize, 32, pngData, isPng: true));
            }

            WriteIco(outputPath, entries.ToArray());
        }

        // Returns the fully-resolved per-size bitmap dictionary that WriteIcoFromSizedSources
        // would produce: every BmpSize plus PngSize is present, either as a clone of the
        // user-supplied bitmap or as a freshly-rendered fallback from the largest filled slot.
        // Caller owns disposal of the returned bitmaps. Useful for previewing the exact pixels
        // the .ico will contain before deciding whether to save.
        public static IDictionary<int, Bitmap> ResolveSources(IDictionary<int, Bitmap> sources)
        {
            if (sources == null || sources.Count == 0)
                throw new ArgumentException("At least one slot must be assigned.", nameof(sources));

            var keys = new System.Collections.Generic.List<int>(sources.Keys);
            keys.Sort((a, b) => b.CompareTo(a));
            Bitmap fallback = sources[keys[0]];

            var resolved = new System.Collections.Generic.Dictionary<int, Bitmap>();
            foreach (int size in BmpSizes)
                resolved[size] = GetOrRender(sources, fallback, size);
            resolved[PngSize] = GetOrRender(sources, fallback, PngSize);
            return resolved;
        }

        // Returns a freshly-allocated Bitmap of exactly size×size for the given slot. If the
        // caller-provided source is already the right size and 32bpp ARGB, a clone is returned
        // (so the result is always safe for the caller to dispose).
        private static Bitmap GetOrRender(IDictionary<int, Bitmap> sources, Bitmap fallback, int size)
        {
            if (sources.TryGetValue(size, out Bitmap exact)
                && exact != null
                && exact.Width == size
                && exact.Height == size
                && exact.PixelFormat == PixelFormat.Format32bppArgb)
            {
                return (Bitmap)exact.Clone();
            }

            Image src = sources.TryGetValue(size, out Bitmap raw) && raw != null ? (Image)raw : fallback;
            return ResizeTo32bpp(src, size);
        }

        // Returns a Bitmap whose bounds tightly enclose the non-fully-transparent pixels of the source.
        // If the source has no alpha channel or no transparent borders, returns an unchanged 32bpp copy.
        // If the source is entirely transparent, returns a 1x1 transparent bitmap.
        internal static Bitmap TrimTransparentBorder(Image source)
        {
            // Always work in 32bpp ARGB so alpha sampling is well-defined.
            var src = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppArgb);
            try
            {
                using (var g = Graphics.FromImage(src))
                {
                    g.CompositingMode = CompositingMode.SourceCopy;
                    g.Clear(Color.Transparent);
                    g.DrawImage(source, 0, 0, source.Width, source.Height);
                }

                int w = src.Width, h = src.Height;
                var rect = new Rectangle(0, 0, w, h);
                var data = src.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                int stride = data.Stride;
                byte[] buf = new byte[stride * h];
                System.Runtime.InteropServices.Marshal.Copy(data.Scan0, buf, 0, buf.Length);
                src.UnlockBits(data);

                int top = -1, bottom = -1, left = w, right = -1;
                for (int y = 0; y < h; y++)
                {
                    int rowStart = y * stride;
                    int rowLeft = -1, rowRight = -1;
                    for (int x = 0; x < w; x++)
                    {
                        // BGRA: alpha at offset +3
                        if (buf[rowStart + x * 4 + 3] != 0)
                        {
                            if (rowLeft < 0) rowLeft = x;
                            rowRight = x;
                        }
                    }
                    if (rowLeft >= 0)
                    {
                        if (top < 0) top = y;
                        bottom = y;
                        if (rowLeft < left) left = rowLeft;
                        if (rowRight > right) right = rowRight;
                    }
                }

                if (top < 0)
                {
                    // Entirely transparent — return a 1x1 transparent bitmap.
                    return new Bitmap(1, 1, PixelFormat.Format32bppArgb);
                }

                if (top == 0 && left == 0 && bottom == h - 1 && right == w - 1)
                {
                    // Already tight; return the 32bpp copy as-is.
                    return src;
                }

                int tw = right - left + 1;
                int th = bottom - top + 1;
                var cropped = new Bitmap(tw, th, PixelFormat.Format32bppArgb);
                using (var g = Graphics.FromImage(cropped))
                {
                    g.CompositingMode = CompositingMode.SourceCopy;
                    g.Clear(Color.Transparent);
                    g.DrawImage(src, new Rectangle(0, 0, tw, th),
                                new Rectangle(left, top, tw, th), GraphicsUnit.Pixel);
                }
                src.Dispose();
                return cropped;
            }
            catch
            {
                return src;
            }
        }

        // Letterbox-resize `source` to a `size` × `size` 32bpp ARGB bitmap. Internal so other
        // types in this assembly (e.g. the Icon Editor pane) can render slot previews using the
        // same centering semantics as WriteIcoFromImage.
        internal static Bitmap ResizeTo32bpp(Image source, int size)
        {
            var dest = new Bitmap(size, size, PixelFormat.Format32bppArgb);
            dest.SetResolution(96, 96);
            using (var g = Graphics.FromImage(dest))
            {
                g.CompositingMode = CompositingMode.SourceCopy;
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Letterbox/scale preserving aspect ratio.
                float srcW = source.Width;
                float srcH = source.Height;
                if (srcW <= 0 || srcH <= 0)
                {
                    g.DrawImage(source, 0, 0, size, size);
                }
                else
                {
                    float scale = Math.Min(size / srcW, size / srcH);
                    float w = srcW * scale;
                    float h = srcH * scale;
                    float x = (size - w) / 2f;
                    float y = (size - h) / 2f;
                    g.DrawImage(source, x, y, w, h);
                }
            }

            return dest;
        }

        private static byte[] BuildPngData(Bitmap bmp)
        {
            using (var ms = new MemoryStream())
            {
                bmp.Save(ms, ImageFormat.Png);
                return ms.ToArray();
            }
        }

        // Build the ICO image payload for a 32bpp BMP entry:
        // BITMAPINFOHEADER (height doubled) + BGRA pixels (bottom-up) + AND mask (1bpp, all zero).
        private static byte[] BuildBmpIconData(Bitmap bmp)
        {
            int w = bmp.Width;
            int h = bmp.Height;

            // Lock and copy BGRA pixel data.
            var rect = new Rectangle(0, 0, w, h);
            var data = bmp.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            byte[] pixels = new byte[w * h * 4];
            try
            {
                int stride = data.Stride;
                IntPtr scan0 = data.Scan0;
                // Read rows bottom-up so written BMP is correctly oriented.
                for (int y = 0; y < h; y++)
                {
                    int srcRow = h - 1 - y;
                    System.Runtime.InteropServices.Marshal.Copy(
                        IntPtr.Add(scan0, srcRow * stride),
                        pixels, y * w * 4, w * 4);
                }
            }
            finally { bmp.UnlockBits(data); }

            // AND mask: 1bpp, rows padded to 4 bytes. All zeros (alpha channel in pixels is authoritative).
            int maskRowBytes = ((w + 31) / 32) * 4;
            int maskSize = maskRowBytes * h;

            int headerSize = 40;
            byte[] result = new byte[headerSize + pixels.Length + maskSize];

            using (var ms = new MemoryStream(result))
            using (var bw = new BinaryWriter(ms))
            {
                // BITMAPINFOHEADER
                bw.Write(headerSize);          // biSize
                bw.Write(w);                   // biWidth
                bw.Write(h * 2);               // biHeight (XOR + AND mask)
                bw.Write((short)1);            // biPlanes
                bw.Write((short)32);           // biBitCount
                bw.Write(0);                   // biCompression BI_RGB
                bw.Write(pixels.Length);       // biSizeImage
                bw.Write(0);                   // biXPelsPerMeter
                bw.Write(0);                   // biYPelsPerMeter
                bw.Write(0);                   // biClrUsed
                bw.Write(0);                   // biClrImportant

                bw.Write(pixels);
                // AND mask zeros already initialized.
            }
            return result;
        }

        private struct IconEntry
        {
            public int Width;
            public int Height;
            public int BitCount;
            public byte[] Data;
            public bool IsPng;
            public IconEntry(int w, int h, int bits, byte[] data, bool isPng)
            {
                Width = w; Height = h; BitCount = bits; Data = data; IsPng = isPng;
            }
        }

        private static void WriteIco(string path, params IconEntry[] entries)
        {
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            using (var bw = new BinaryWriter(fs))
            {
                // ICONDIR
                bw.Write((short)0);                  // reserved
                bw.Write((short)1);                  // type = icon
                bw.Write((short)entries.Length);     // count

                int dataOffset = 6 + 16 * entries.Length;
                foreach (var e in entries)
                {
                    bw.Write((byte)(e.Width >= 256 ? 0 : e.Width));
                    bw.Write((byte)(e.Height >= 256 ? 0 : e.Height));
                    bw.Write((byte)0);               // color count
                    bw.Write((byte)0);               // reserved
                    bw.Write((short)1);              // planes
                    bw.Write((short)e.BitCount);     // bit count
                    bw.Write(e.Data.Length);         // bytes in res
                    bw.Write(dataOffset);            // image offset
                    dataOffset += e.Data.Length;
                }

                foreach (var e in entries)
                {
                    bw.Write(e.Data);
                }
            }
        }
    }
}
