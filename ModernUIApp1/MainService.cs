using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media.Imaging;
using static ModernUIApp1.CommonConstants;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;

namespace ModernUIApp1
{
    class MainService
    {

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr FindWindowEx(IntPtr hWndParent, IntPtr hWndChildAfter, string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        static extern bool ClientToScreen(IntPtr hWnd, ref Point lpPoint);

        [DllImport("user32.dll")]
        static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("User32.Dll", EntryPoint = "PostMessageA")]
        static extern bool PostMessage(IntPtr hWnd, uint msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetScrollInfo(IntPtr hwnd, int fnBar, ref SCROLLINFO lpsi);

        [DllImport("gdi32.dll")]
        public static extern bool DeleteObject(IntPtr hObject);

        public static Bitmap ImageCombineV(List<Bitmap> src)
        {
            // 結合後のサイズを計算
            int dstWidth = 0, dstHeight = 0;
            System.Drawing.Imaging.PixelFormat dstPixelFormat = System.Drawing.Imaging.PixelFormat.Format8bppIndexed;

            foreach (Bitmap bmp in src)
            {
                if (dstWidth < bmp.Width) dstWidth = bmp.Width;
                dstHeight += bmp.Height;

                // 最大のビット数を検索
                if (Image.GetPixelFormatSize(dstPixelFormat) < Image.GetPixelFormatSize(bmp.PixelFormat))
                {
                    dstPixelFormat = bmp.PixelFormat;
                }
            }

            var dst = new Bitmap(dstWidth, dstHeight, dstPixelFormat);
            var dstRect = new Rectangle();

            using (var g = Graphics.FromImage(dst))
            {
                foreach (Bitmap bmp in src)
                {
                    dstRect.Width = bmp.Width;
                    dstRect.Height = bmp.Height;

                    // 描画
                    g.DrawImage(bmp, dstRect, 0, 0, bmp.Width, bmp.Height, GraphicsUnit.Pixel);

                    // 次の描画先
                    dstRect.Y = dstRect.Bottom;
                }
            }
            return dst;
        }

        public static void Execute()
        {
            List<Bitmap> bitmaps = new List<Bitmap>();

            IntPtr tmpPtr = FindWindow("Notepad", null);

            if (tmpPtr.Equals(IntPtr.Zero))
            {
                MessageBox.Show("Target application is not open", "Error");
                return;
            }

            IntPtr intPtr = FindWindowEx(tmpPtr, IntPtr.Zero,"Edit", null);

            if (intPtr.Equals(IntPtr.Zero))
            {
                MessageBox.Show("Target application is not open", "Error");
                return;
            }

            SetActiveWindow(tmpPtr);

            Point defaultPnt = default;

            SCROLLINFO VScrInfo = new SCROLLINFO();
            VScrInfo.cbSize = (uint)Marshal.SizeOf(VScrInfo);
            VScrInfo.fMask = (int)ScrollInfoMask.SIF_ALL;

            // GetScrollInfo
            GetScrollInfo(intPtr, (int)ScrollBarDirection.SB_VERT, ref VScrInfo);

            double pageCount = Math.Truncate((double)VScrInfo.nMax / VScrInfo.nPage);

            int finalScrPos = 0;

            for (int i = 0; i < pageCount + 1; i++)
            {
                Point origin = default;

                int operationInt;

                if (i == 0)
                {
                    operationInt = SB_PAGETOP;
                }
                else
                {
                    operationInt = SB_PAGEDOWN;
                }

                PostMessage(intPtr, 0x115, operationInt, 0);

                System.Threading.Thread.Sleep(100);

                ClientToScreen(intPtr, ref origin);

                GetClientRect(intPtr, out RECT rect);

                int height;
                int margin = 0;

                if (i == pageCount - 1)
                {
                    GetScrollInfo(intPtr, (int)ScrollBarDirection.SB_VERT, ref VScrInfo);
                    finalScrPos = VScrInfo.nPos;
                }

                if (i == pageCount)
                {
                    GetScrollInfo(intPtr, (int)ScrollBarDirection.SB_VERT, ref VScrInfo);
                    margin = (int)(rect.Height * (VScrInfo.nPos - finalScrPos) / VScrInfo.nPage);
                    height = margin;
                }
                else
                {
                    height = rect.Height;
                }

                using (var bmp = new Bitmap(rect.Width, height))
                {
                    using (var g = Graphics.FromImage(bmp))
                    {
                        Size size = rect.Size;
                        if (i == pageCount)
                        {
                            size = new System.Drawing.Size(rect.Width, margin);
                            g.CopyFromScreen(new Point(origin.X + rect.Left, origin.Y + rect.Top + rect.Height - margin), defaultPnt, size);
                        }
                        else
                        {
                            g.CopyFromScreen(new Point(origin.X + rect.Left, origin.Y + rect.Top), defaultPnt, size);
                        }
                        bitmaps.Add((Bitmap)bmp.Clone());
                    }
                }
            }

            var hBitmap = ImageCombineV(bitmaps).GetHbitmap();

            try
            {
                var bitmapSource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    hBitmap,
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions()
                );

                Clipboard.SetImage(bitmapSource);
            }
            finally
            {
                DeleteObject(hBitmap);
            }
        }
    }
}
