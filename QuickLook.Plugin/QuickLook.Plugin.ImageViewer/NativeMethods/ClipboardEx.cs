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
                
                var clipboardBitmap = new Bitmap(originalBitmap);
                Clipboard.SetImage(clipboardBitmap);
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