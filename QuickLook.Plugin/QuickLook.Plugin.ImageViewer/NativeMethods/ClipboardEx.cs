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

                using var bitmap = image.ToBitmap();
                var data = new System.Windows.Forms.DataObject();

                // 标准位图，给微信/QQ/钉钉用
                data.SetData(DataFormats.Bitmap, bitmap);

                // PNG 流，给 Word/浏览器用
                using var pngStream = new MemoryStream();
                bitmap.Save(pngStream, ImageFormat.Png);
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
