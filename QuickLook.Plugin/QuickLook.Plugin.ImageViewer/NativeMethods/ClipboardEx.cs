using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Windows.Media.Imaging;
using Clipboard = System.Windows.Forms.Clipboard;

namespace QuickLook.Plugin.ImageViewer.NativeMethods;

internal static class ClipboardEx
{
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
                
                // 【关键】必须脱离 using 生命周期，否则微信/QQ粘贴时底层内存已被释放会报错
                var clipboardBitmap = new Bitmap(originalBitmap);

                // 显式使用完整命名空间，彻底避免与 WPF 的 DataObject/DataFormats 冲突
                var data = new System.Windows.Forms.DataObject();

                // 标准位图格式 - 基本支持
                data.SetData(System.Windows.Forms.DataFormats.Bitmap, clipboardBitmap);

                // DIB 格式 - 支持更多应用程序，如微信、QQ、画图等
                using var dibStream = new MemoryStream();
                SaveDib(originalBitmap, dibStream);
                dibStream.Position = 0;
                data.SetData(System.Windows.Forms.DataFormats.Dib, dibStream);

                // PNG 格式 - 给 Word/浏览器用
                using var pngStream = new MemoryStream();
                originalBitmap.Save(pngStream, ImageFormat.Png);
                pngStream.Position = 0;
                // 因为前面用了完整命名空间，这里的重载解析绝对不会错位
                data.SetData("PNG", pngStream, false);

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

    private static Bitmap ToBitmap(this BitmapSource source)
    {
        using var outStream = new MemoryStream();
        BitmapEncoder enc = new PngBitmapEncoder();
        enc.Frames.Add(BitmapFrame.Create(source));
        enc.Save(outStream);
        using var temp = new Bitmap(outStream);
        return new Bitmap(temp);
    }

    private static void SaveDib(Bitmap bitmap, Stream stream)
    {
        // 保存为 DIB 格式
        var bmi = bitmap.GetBitmapHeaderInfo();
        var data = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), System.Drawing.Imaging.ImageLockMode.ReadOnly, bitmap.PixelFormat);
        
        try
        {
            // 写入 BITMAPINFOHEADER
            stream.Write(bmi, 0, bmi.Length);
            
            // 写入像素数据
            var bytes = new byte[data.Stride * data.Height];
            System.Runtime.InteropServices.Marshal.Copy(data.Scan0, bytes, 0, bytes.Length);
            stream.Write(bytes, 0, bytes.Length);
        }
        finally
        {
            bitmap.UnlockBits(data);
        }
    }

    private static byte[] GetBitmapHeaderInfo(this Bitmap bitmap)
    {
        var bmi = new System.Drawing.Imaging.BitmapInfoHeader();
        bmi.Size = (uint)System.Runtime.InteropServices.Marshal.SizeOf(bmi);
        bmi.Width = (int)bitmap.Width;
        bmi.Height = (int)bitmap.Height;
        bmi.Planes = 1;
        bmi.BitCount = (ushort)(Image.GetPixelFormatSize(bitmap.PixelFormat));
        bmi.Compression = 0; // BI_RGB
        bmi.SizeImage = (uint)(bitmap.Width * bitmap.Height * (bmi.BitCount / 8));
        bmi.XPelsPerMeter = 0;
        bmi.YPelsPerMeter = 0;
        bmi.ClrUsed = 0;
        bmi.ClrImportant = 0;
        
        var result = new byte[bmi.Size];
        var ptr = System.Runtime.InteropServices.Marshal.AllocHGlobal((int)bmi.Size);
        try
        {
            System.Runtime.InteropServices.Marshal.StructureToPtr(bmi, ptr, false);
            System.Runtime.InteropServices.Marshal.Copy(ptr, result, 0, (int)bmi.Size);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.FreeHGlobal(ptr);
        }
        return result;
    }
}
