using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShellPhotoPrint
{
    public partial class Form1 : Form
    {
        public enum HRESULT : int
        {
            DRAGDROP_S_CANCEL = 0x00040101,
            DRAGDROP_S_DROP = 0x00040100,
            DRAGDROP_S_USEDEFAULTCURSORS = 0x00040102,
            DATA_S_SAMEFORMATETC = 0x00040130,
            S_OK = 0,
            S_FALSE = 1,
            E_NOINTERFACE = unchecked((int)0x80004002),
            E_NOTIMPL = unchecked((int)0x80004001),
            OLE_E_ADVISENOTSUPPORTED = unchecked((int)0x80040003),
            E_FAIL = unchecked((int)0x80004005),
            MK_E_NOOBJECT = unchecked((int)0x800401E5)
        }

        public struct POINT
        {
            public long X;
            public long Y;
        }

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern HRESULT SHILCreateFromPath([MarshalAs(UnmanagedType.LPWStr)] string pszPath, out IntPtr ppIdl, ref uint rgflnOut);


        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr ILFindLastID(IntPtr pidl);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr ILClone(IntPtr pidl);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern Boolean ILRemoveLastID(IntPtr pidl);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern void ILFree(IntPtr pidl);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "#740")]
        public static extern HRESULT SHCreateFileDataObject(IntPtr pidlFolder, uint cidl, IntPtr[] apidl, System.Runtime.InteropServices.ComTypes.IDataObject pdtInner, out System.Runtime.InteropServices.ComTypes.IDataObject ppdtobj);

        public const int DROPEFFECT_NONE = (0);
        [ComImport]
        [Guid("00000122-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IDropTarget
        {
            HRESULT DragEnter(
                [In] System.Runtime.InteropServices.ComTypes.IDataObject pDataObj,
                [In] int grfKeyState,
                [In] POINT pt,
                [In, Out] ref int pdwEffect);

            HRESULT DragOver(
                [In] int grfKeyState,
                [In] POINT pt,
                [In, Out] ref int pdwEffect);

            HRESULT DragLeave();

            HRESULT Drop(
                [In] System.Runtime.InteropServices.ComTypes.IDataObject pDataObj,
                [In] int grfKeyState,
                [In] POINT pt,
                [In, Out] ref int pdwEffect);
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            openFileDialog1.Filter = "JPEG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif|All Files (*.*)|*.*|Bitmap Files (*.bmp)|*.bmp|PNG Files (*.png)|*.png|TIFF Files (*.tif, *.tiff)|*.tif;*.tiff|Icon Files (*.ico)|*.ico";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = true;
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                HRESULT hr = HRESULT.E_FAIL;
                IntPtr pidlParent = IntPtr.Zero, pidlFull = IntPtr.Zero, pidlItem = IntPtr.Zero;
                var aPidl = new IntPtr[255];
                uint rgflnOut = 0;
                uint nIndex = 0;
                int nCount = openFileDialog1.FileNames.Length;
                if (nCount == 1)
                {
                    hr = SHILCreateFromPath(openFileDialog1.FileNames[0], out pidlFull, ref rgflnOut);
                    if (hr == HRESULT.S_OK)
                    {
                        pidlItem = ILFindLastID(pidlFull);
                        aPidl[nIndex++] = ILClone(pidlItem);
                        ILRemoveLastID(pidlFull);
                        pidlParent = ILClone(pidlFull);
                        ILFree(pidlFull);
                    }
                }
                else if (nCount > 1)
                {
                    string sPath = Path.GetDirectoryName(openFileDialog1.FileNames[0]);
                    hr = SHILCreateFromPath(sPath, out pidlParent, ref rgflnOut);
                    foreach (String file in openFileDialog1.FileNames)
                    {
                        hr = SHILCreateFromPath(file, out pidlFull, ref rgflnOut);
                        if (hr == HRESULT.S_OK)
                        {
                            pidlItem = ILFindLastID(pidlFull);
                            aPidl[nIndex++] = ILClone(pidlItem);
                            ILFree(pidlFull);
                        }
                    }
                }
                if (nIndex > 0)
                {
                    System.Runtime.InteropServices.ComTypes.IDataObject pdo;
                    hr = SHCreateFileDataObject(pidlParent, nIndex, aPidl, null, out pdo);
                    if (hr == HRESULT.S_OK)
                    {
                        Guid CLSID_PrintPhotosDropTarget = new Guid("60fd46de-f830-4894-a628-6fa81bc0190d");
                        Type DropTargetType = Type.GetTypeFromCLSID(CLSID_PrintPhotosDropTarget, true);
                        object DropTarget = Activator.CreateInstance(DropTargetType);
                        IDropTarget pDropTarget = (IDropTarget)DropTarget;

                        int pdwEffect = DROPEFFECT_NONE;
                        POINT pt;
                        pt.X = 0;
                        pt.Y = 0;
                        hr = pDropTarget.Drop(pdo, 0, pt, pdwEffect);
                    }
                }
                if (pidlParent != IntPtr.Zero)
                {
                    ILFree(pidlParent);
                }

                for (int i = 0; i < nIndex; i++)
                {
                    if (aPidl[i] != IntPtr.Zero)
                    {
                        ILFree(aPidl[i]);
                    }
                }
            }
        }
    }
}
