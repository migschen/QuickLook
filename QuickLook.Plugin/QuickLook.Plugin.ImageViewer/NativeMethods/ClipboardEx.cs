using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Windows.Media.Imaging;
using Clipboard = System.Windows.Forms.Clipboard;
using WFDataObject = System.Windows.Forms.DataObject;
using WFDataFormats = System.Windows.Forms.DataFormats;

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
                
                // 【关键修复】：必须克隆一份脱离 using 生命周期的 Bitmap 给剪贴板
                // 否则方法结束时 originalBitmap 被释放，微信/QQ去读就会报错（红X或粘贴失败）
                var clipboardBitmap = new Bitmap(originalBitmap);

                // 显式使用 WinForms 的 DataObject，避免与 WPF 冲突
                var data = new WFDataObject();

                // 显式使用 WinForms 的 DataFormats
                // 标准位图，给微信/QQ/钉钉用
                data.SetData(WFDataFormats.Bitmap, clipboardBitmap);

                // PNG 流，给 Word/浏览器用
                using var pngStream = new MemoryStream();
                originalBitmap.Save(pngStream, ImageFormat.Png);
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
        return new Bitmap(temp); // 这里内部的 new Bitmap(temp) 已经处理了流关闭的问题，没问题
    }
}
