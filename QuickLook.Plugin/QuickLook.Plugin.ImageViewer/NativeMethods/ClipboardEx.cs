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

                // 给微信/QQ/钉钉用
                data.SetData(System.Windows.Forms.DataFormats.Bitmap, clipboardBitmap);

                // 给 Word/浏览器用
                using var pngStream = new MemoryStream();
                originalBitmap.Save(pngStream, ImageFormat.Png);
                
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
}
