using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Security;
namespace EVWB
{
    public class GetHtmlToImage
    {
        public static Bitmap GetHtmlImage(Uri UrlString, int Width)
        {
            WebBrowser MyControl = new WebBrowser();
            MyControl.Size = new Size(Width, 10);
            MyControl.Url = UrlString;
            while (MyControl.ReadyState != WebBrowserReadyState.Complete)
            {
                Application.DoEvents();
            }
            MyControl.Height = MyControl.Document.Body.ScrollRectangle.Height + 20;
            MyControl.Url = UrlString;
            WebControlImage.Snapshot snap = new WebControlImage.Snapshot();
            Bitmap MyImage = snap.TakeSnapshot(MyControl.ActiveXInstance, new Rectangle(0, 0, MyControl.Width, MyControl.Height));
            MyControl.Dispose();
            return MyImage;
        }
        /// 
        /// WebBrowser获取图形 
        /// 
        private class WebControlImage
        {
            internal static class NativeMethods
            {
                [StructLayout(LayoutKind.Sequential)]
                public sealed class tagDVTARGETDEVICE
                {
                    [MarshalAs(UnmanagedType.U4)]
                    public int tdSize;
                    [MarshalAs(UnmanagedType.U2)]
                    public short tdDriverNameOffset;
                    [MarshalAs(UnmanagedType.U2)]
                    public short tdDeviceNameOffset;
                    [MarshalAs(UnmanagedType.U2)]
                    public short tdPortNameOffset;
                    [MarshalAs(UnmanagedType.U2)]
                    public short tdExtDevmodeOffset;
                }
                [StructLayout(LayoutKind.Sequential)]
                public class COMRECT
                {
                    public int left;
                    public int top;
                    public int right;
                    public int bottom;
                    public COMRECT()
                    {
                    }
                    public COMRECT(Rectangle r)
                    {
                        this.left = r.X;
                        this.top = r.Y;
                        this.right = r.Right;
                        this.bottom = r.Bottom;
                    }
                    public COMRECT(int left, int top, int right, int bottom)
                    {
                        this.left = left;
                        this.top = top;
                        this.right = right;
                        this.bottom = bottom;
                    }
                    public static NativeMethods.COMRECT FromXYWH(int x, int y, int width, int height)
                    {
                        return new NativeMethods.COMRECT(x, y, x + width, y + height);
                    }
                    public override string ToString()
                    {
                        return string.Concat(new object[] { "Left = ", this.left, " Top ", this.top, " Right = ", this.right, " Bottom = ", this.bottom });
                    }
                }
                [StructLayout(LayoutKind.Sequential)]
                public sealed class tagLOGPALETTE
                {
                    [MarshalAs(UnmanagedType.U2)]
                    public short palVersion;
                    [MarshalAs(UnmanagedType.U2)]
                    public short palNumEntries;
                }
            }
            public class Snapshot
            {
                
                public Bitmap TakeSnapshot(object pUnknown, Rectangle bmpRect)
                {
                    if (pUnknown == null)
                        return null;
                    //必须为com对象 
                    if (!Marshal.IsComObject(pUnknown))
                        return null;
                    //IViewObject 接口 
                    UnsafeNativeMethods.IViewObject ViewObject = null;
                    IntPtr pViewObject = IntPtr.Zero;
                    //内存图 
                    Bitmap pPicture = new Bitmap(bmpRect.Width, bmpRect.Height);
                    Graphics hDrawDC = Graphics.FromImage(pPicture);
                    //获取接口 
                    object hret = Marshal.QueryInterface(Marshal.GetIUnknownForObject(pUnknown),
                    ref UnsafeNativeMethods.IID_IViewObject, out pViewObject);
                    try
                    {
                        ViewObject = Marshal.GetTypedObjectForIUnknown(pViewObject, typeof(UnsafeNativeMethods.IViewObject)) as UnsafeNativeMethods.IViewObject;
                        //调用Draw方法 
                        ViewObject.Draw((int)System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT,
                        -1,
                        IntPtr.Zero,
                        null,
                        IntPtr.Zero,
                        hDrawDC.GetHdc(),
                        new NativeMethods.COMRECT(bmpRect),
                        null,
                        IntPtr.Zero,
                        0);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        throw ex;
                    }
                    //释放 
                    hDrawDC.Dispose();
                    return pPicture;
                }
            }
            [SuppressUnmanagedCodeSecurity]
            internal static class UnsafeNativeMethods
            {
                public static Guid IID_IViewObject = new Guid("{0000010d-0000-0000-C000-000000000046}");
                [ComImport, Guid("0000010d-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
                public interface IViewObject
                {
                    [PreserveSig]
                    int Draw([In, MarshalAs(UnmanagedType.U4)] int dwDrawAspect, int lindex, IntPtr pvAspect, [In] NativeMethods.tagDVTARGETDEVICE ptd, IntPtr hdcTargetDev, IntPtr hdcDraw, [In] NativeMethods.COMRECT lprcBounds, [In] NativeMethods.COMRECT lprcWBounds, IntPtr pfnContinue, [In] int dwContinue);
                    [PreserveSig]
                    int GetColorSet([In, MarshalAs(UnmanagedType.U4)] int dwDrawAspect, int lindex, IntPtr pvAspect, [In] NativeMethods.tagDVTARGETDEVICE ptd, IntPtr hicTargetDev, [Out] NativeMethods.tagLOGPALETTE ppColorSet);
                    [PreserveSig]
                    int Freeze([In, MarshalAs(UnmanagedType.U4)] int dwDrawAspect, int lindex, IntPtr pvAspect, [Out] IntPtr pdwFreeze);
                    [PreserveSig]
                    int Unfreeze([In, MarshalAs(UnmanagedType.U4)] int dwFreeze);
                    void SetAdvise([In, MarshalAs(UnmanagedType.U4)] int aspects, [In, MarshalAs(UnmanagedType.U4)] int advf, [In, MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IAdviseSink pAdvSink);
                    void GetAdvise([In, Out, MarshalAs(UnmanagedType.LPArray)] int[] paspects, [In, Out, MarshalAs(UnmanagedType.LPArray)] int[] advf, [In, Out, MarshalAs(UnmanagedType.LPArray)] System.Runtime.InteropServices.ComTypes.IAdviseSink[] pAdvSink);
                }
            }
        }
    }
}

