using System;
using System.Runtime.InteropServices;

namespace IEExtension
{
    [ComImport()]
    [Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDispatch
    {
        [PreserveSig]
        int GetTypeInfoCount(out int Count);

        [PreserveSig]
        int GetTypeInfo
        (
          [MarshalAs(UnmanagedType.U4)] int iTInfo,
          [MarshalAs(UnmanagedType.U4)] int lcid,
          out System.Runtime.InteropServices.ComTypes.ITypeInfo typeInfo
        );

        [PreserveSig]
        int GetIDsOfNames
        (
          ref Guid riid,
          [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)]
          string[] rgsNames,
          int cNames,
          int lcid,
          [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId
        );

        [PreserveSig]
        int Invoke
        (
          int dispIdMember,
          ref Guid riid,
          uint lcid,
          ushort wFlags,
          ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams,
          out object pVarResult,
          ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo,
          out UInt32 pArgErr
        );
    }


    [ComImport()]
    [Guid("A6EF9860-C720-11D0-9337-00A0C90DCAA9")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDispatchEx
    {
        // IDispatch
        int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        System.Runtime.InteropServices.ComTypes.ITypeInfo GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        void GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames, [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        void Invoke(int dispIdMember, ref Guid riid, uint lcid, ushort wFlags, ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, out object pVarResult, ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, IntPtr[] pArgErr);

        // IDispatchEx
        void GetDispID([MarshalAs(UnmanagedType.BStr)] string bstrName, uint grfdex, [Out] out int pid);
        void InvokeEx(int id, uint lcid, ushort wFlags,
            [In] ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pdp,
            [In, Out] IntPtr pvarRes,
            [In, Out] IntPtr pei,
            System.IServiceProvider pspCaller);
        void DeleteMemberByName([MarshalAs(UnmanagedType.BStr)] string bstrName, uint grfdex);
        void DeleteMemberByDispID(int id);
        void GetMemberProperties(int id, uint grfdexFetch, [Out] out uint pgrfdex);
        void GetMemberName(int id, [Out, MarshalAs(UnmanagedType.BStr)] out string pbstrName);
        [PreserveSig]
        [return: MarshalAs(UnmanagedType.I4)]
        int GetNextDispID(uint grfdex, int id, [In, Out] ref int pid);
        void GetNameSpaceParent([Out, MarshalAs(UnmanagedType.IUnknown)] out object ppunk);
    }
}