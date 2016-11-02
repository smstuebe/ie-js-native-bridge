using System;
using System.Runtime.InteropServices;

namespace IEExtension
{
    class Attach
    {
        [DllImport(@"oleaut32.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall)]
        static extern Int32 VariantClear(IntPtr pvarg);

        private const int LOCALE_SYSTEM_DEFAULT = 2048;
        private const ushort DISPATCH_PROPERTYPUT = 4;
        private const int DISPID_PROPERTYPUT = -3;

        public static int AttachBridge(IDispatchEx dispatch, out INativeBridge brigde)
        {
            var nullId = new Guid("00000000-0000-0000-0000-000000000000");

            brigde = new NativeBridge();

            var pNamedArg = Marshal.AllocCoTaskMem(sizeof(Int64));
            Marshal.WriteInt64(pNamedArg, DISPID_PROPERTYPUT);
            var pVariant = Marshal.AllocCoTaskMem(16);
            Marshal.GetNativeVariantForObject(brigde, pVariant);

            try
            {
                var name = "bridge";
                int dispId;
                dispatch.GetDispID(name, 0x00000002U, out dispId);

                var param = new System.Runtime.InteropServices.ComTypes.DISPPARAMS
                {
                    rgvarg = pVariant,
                    rgdispidNamedArgs = pNamedArg,
                    cArgs = 1,
                    cNamedArgs = 0
                };

                var err = new IntPtr[0];
                object varResult;
                var excepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();

                dispatch.Invoke(dispId, ref nullId, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, ref param, out varResult, ref excepInfo, err);
            }
            finally
            {
                VariantClear(pVariant);
                Marshal.FreeCoTaskMem(pVariant);
                VariantClear(pNamedArg);
                Marshal.FreeCoTaskMem(pNamedArg);
            }

            return 0;
        }
    }
}