using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Media.Imaging;
using Clipboard = System.Windows.Forms.Clipboard;

namespace QuickLook.Plugin.ImageViewer.NativeMethods;

internal static class ClipboardEx
{
    [StructLayout(LayoutKind.Sequential)]
    private struct BITMAPINFOHEADER
    {
        public uint biSize;
        public int biWidth;
        public int biHeight;
        public ushort biPlanes;
        public ushort biBitCount;
        public uint biCompression;
        public uint biSizeImage;
        public int biXPelsPerMeter;
        public int biYPelsPerMeter;
        public uint biClrUsed;
        public uint biClrImportant;
    }

    public static void SetClipboardImage(this BitmapSource img)
    {
        if (img == null)
            return;

        var thread = new Thread(state =>
        {
            if (state is not BitmapSource image)
                return;

            try { Clipboard.Clear(); } catch { }

            try
            {
                if (image is BitmapFrame && image.IsFrozen)
                {
                    image = new WriteableBitmap(image);
                }

                using var originalBitmap = image.ToBitmap();

                var data = new DataObject();
                data.SetData(DataFormats.Bitmap, (object)originalBitmap, true);


                var dibData = CreateDibData(originalBitmap);
                if (dibData != null)
                {
                    data.SetData(DataFormats.Dib, dibData, true);
                }

                Clipboard.SetDataObject(data, true);
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
            }
        })
        {
            Name = nameof(ClipboardEx)
        };

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start(img);
    }

    private static byte[] CreateDibData(Bitmap bitmap)
    {
        try
        {
            var bmi = new BITMAPINFOHEADER();
            bmi.biSize = (uint)Marshal.SizeOf(typeof(BITMAPINFOHEADER));
            bmi.biWidth = bitmap.Width;
            bmi.biHeight = -bitmap.Height;
            bmi.biPlanes = 1;
            bmi.biBitCount = (ushort)Image.GetPixelFormatSize(bitmap.PixelFormat);
            bmi.biCompression = 0;
            bmi.biSizeImage = (uint)((bitmap.Width * bmi.biBitCount / 8 + 3) & ~3) * (uint)bitmap.Height;
            bmi.biXPelsPerMeter = 0;
            bmi.biYPelsPerMeter = 0;
            bmi.biClrUsed = 0;
            bmi.biClrImportant = 0;

            var headerSize = Marshal.SizeOf(typeof(BITMAPINFOHEADER));
            var result = new byte[headerSize + (int)bmi.biSizeImage];

            var headerPtr = Marshal.AllocHGlobal(headerSize);
            try
            {
                Marshal.StructureToPtr(bmi, headerPtr, false);
                Marshal.Copy(headerPtr, result, 0, headerSize);
            }
            finally
            {
                Marshal.FreeHGlobal(headerPtr);
            }

            var rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
            var bitmapData = bitmap.LockBits(rect, ImageLockMode.ReadOnly, bitmap.PixelFormat);

            try
            {
                var destRowSize = (bitmap.Width * bmi.biBitCount / 8 + 3) & ~3;
                var srcStride = bitmapData.Stride;

                for (var y = 0; y < bitmap.Height; y++)
                {
                    var srcOffset = y * srcStride;
                    var destOffset = headerSize + y * destRowSize;
                    Marshal.Copy(bitmapData.Scan0 + srcOffset, result, destOffset, destRowSize);
                }
            }
            finally
            {
                bitmap.UnlockBits(bitmapData);
            }

            return result;
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex);
            return null;
        }
    }

    private static Bitmap ToBitmap(this BitmapSource source)
    {
        using var outStream = new MemoryStream();
        BitmapEncoder enc = new PngBitmapEncoder();
        enc.Frames.Add(BitmapFrame.Create(source));
        enc.Save(outStream);
        using var temp = new Bitmap(outStream);
        return new Bitmap(temp);
    }
}