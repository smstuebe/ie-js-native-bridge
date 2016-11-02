using System.Runtime.InteropServices;

namespace IEExtension
{
    [Guid("DCDC3289-96BB-4227-AF40-74BEC457B2AF"),
     // Tell Interop Marshalling Layer how to interpret v-table of interface.
     // In the .idl file the 'dual' attribute is specified,
     // meaning that the interface must be compatible with Automation 
     // and be derived from IDispatch (early & late binding).
     InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface INativeBridge
    {
        int Test([In, MarshalAs(UnmanagedType.BStr)] string str);
    }

    [ComVisible(true)]
    [TypeLibType(TypeLibTypeFlags.FCanCreate)]
    [Guid("19F9905F-367B-4197-BC35-DD3BDCF1F7E2")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComDefaultInterface(typeof(INativeBridge))]
    [ProgId("IEExtension.NativeBridge")]
    public class NativeBridge : INativeBridge
    {
        public int Test(string str)
        {
            return 0;
        }
    }
}