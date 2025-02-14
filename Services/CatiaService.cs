using System;
using System.Runtime.InteropServices;
using INFITF;

namespace CATIAAssistant.Services
{
    public class CatiaService
    {
        public INFITF.Application CatiaApplication { get; private set; }

        // CLSIDFromProgID fonksiyonunu tanımlıyoruz.
        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        private static extern int CLSIDFromProgID(
            [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
            out Guid pclsid);

        // GetActiveObject fonksiyonunu P/Invoke ile tanımlıyoruz.
        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetActiveObject(
            [In] ref Guid rclsid,
            IntPtr reserved,
            out IntPtr ppunk);

        public bool Connect()
        {
            try
            {
                // CATIA'nın ProgID'sini kullanarak CLSID'i alıyoruz.
                Guid clsid;
                int hr = CLSIDFromProgID("CATIA.Application", out clsid);
                if (hr != 0)
                    throw new Exception("CLSIDFromProgID failed with HR: " + hr);

                // Aktif nesneyi almaya çalışıyoruz.
                IntPtr pUnk;
                hr = GetActiveObject(ref clsid, IntPtr.Zero, out pUnk);
                if (hr != 0)
                    throw new Exception("GetActiveObject failed with HR: " + hr);

                // COM nesnesini managed nesneye dönüştürüyoruz.
                object obj = Marshal.GetObjectForIUnknown(pUnk);
                Marshal.Release(pUnk);

                CatiaApplication = (INFITF.Application)obj;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
