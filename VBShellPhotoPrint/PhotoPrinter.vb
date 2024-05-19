Imports System
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public NotInheritable Class PhotoPrinter

    Private Shared ReadOnly _instance As New Lazy(Of PhotoPrinter)(Function() New PhotoPrinter(), System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

    Public Shared ReadOnly Property Instance() As PhotoPrinter
        Get
            Return _instance.Value
        End Get
    End Property

    Private Enum HRESULT As UInt32
        DRAGDROP_S_CANCEL = &H40101
        DRAGDROP_S_DROP = &H40100
        DRAGDROP_S_USEDEFAULTCURSORS = &H40102
        DATA_S_SAMEFORMATETC = &H40130
        S_OK = 0
        S_FALSE = 1
        E_NOINTERFACE = &H80004002UI
        E_NOTIMPL = &H80004001UI
        OLE_E_ADVISENOTSUPPORTED = &H80040003UI
        E_FAIL = &H80004005UI
        MK_E_NOOBJECT = &H800401E5UI
    End Enum
    Private Structure POINT
        Public X As Long
        Public Y As Long
    End Structure

    <DllImport("Shell32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function SHILCreateFromPath(
    <MarshalAs(UnmanagedType.LPWStr)> ByVal pszPath As String, <Out> ByRef ppIdl As IntPtr, ByRef rgflnOut As UInteger) As HRESULT
    End Function


    <DllImport("Shell32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function ILFindLastID(ByVal pidl As IntPtr) As IntPtr
    End Function

    <DllImport("Shell32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function ILClone(ByVal pidl As IntPtr) As IntPtr
    End Function

    <DllImport("Shell32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function ILRemoveLastID(ByVal pidl As IntPtr) As Boolean
    End Function

    <DllImport("Shell32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Sub ILFree(ByVal pidl As IntPtr)
    End Sub

    <DllImport("Shell32.dll", CharSet:=CharSet.Unicode, SetLastError:=True, EntryPoint:="#740")>
    Private Shared Function SHCreateFileDataObject(ByVal pidlFolder As IntPtr, ByVal cidl As UInteger, ByVal apidl As IntPtr(), ByVal pdtInner As ComTypes.IDataObject, <Out> ByRef ppdtobj As ComTypes.IDataObject) As HRESULT
    End Function

    Private Const DROPEFFECT_NONE As Integer = 0
    <ComImport>
    <Guid("00000122-0000-0000-C000-000000000046")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IDropTarget
        Function DragEnter(
            <[In]> ByVal pDataObj As ComTypes.IDataObject,
            <[In]> ByVal grfKeyState As Integer,
            <[In]> ByVal pt As POINT,
            <[In], Out> ByRef pdwEffect As Integer) As HRESULT

        Function DragOver(
            <[In]> ByVal grfKeyState As Integer,
            <[In]> ByVal pt As POINT,
            <[In], Out> ByRef pdwEffect As Integer) As HRESULT

        Function DragLeave() As HRESULT

        Function Drop(
            <[In]> ByVal pDataObj As ComTypes.IDataObject,
            <[In]> ByVal grfKeyState As Integer,
            <[In]> ByVal pt As POINT,
            <[In], Out> ByRef pdwEffect As Integer) As HRESULT
    End Interface

    Private Sub New()
    End Sub

    Public Function PrintEm(ByRef FilePaths() As String) As Boolean
        Dim hr = HRESULT.E_FAIL
        Dim pidlParent = IntPtr.Zero, pidlFull = IntPtr.Zero, pidlItem = IntPtr.Zero
        Dim aPidl = New IntPtr(254) {}
        Dim rgflnOut As UInteger = 0
        Dim nIndex As UInteger = 0
        Dim nCount As Integer = FilePaths.Count

        PrintEm = False

        If nCount = 1 Then
            hr = SHILCreateFromPath(FilePaths(0), pidlFull, rgflnOut)

            If hr = HRESULT.S_OK Then
                pidlItem = ILFindLastID(pidlFull)
                aPidl(nIndex) = ILClone(pidlItem)
                ILRemoveLastID(pidlFull)
                pidlParent = ILClone(pidlFull)
                ILFree(pidlFull)

                nIndex += 1
            End If
        ElseIf nCount > 1 Then
            Dim sPath As String = Path.GetDirectoryName(FilePaths(0))

            hr = SHILCreateFromPath(sPath, pidlParent, rgflnOut)
            If hr = HRESULT.S_OK Then
                For Each file As String In FilePaths
                    hr = SHILCreateFromPath(file, pidlFull, rgflnOut)

                    If hr = HRESULT.S_OK Then
                        pidlItem = ILFindLastID(pidlFull)
                        aPidl(nIndex) = ILClone(pidlItem)
                        ILFree(pidlFull)

                        nIndex += 1
                    End If
                Next
            End If
        End If

        If nIndex > 0 Then
            Dim pdo As ComTypes.IDataObject = Nothing
            hr = SHCreateFileDataObject(pidlParent, nIndex, aPidl, Nothing, pdo)

            If hr = HRESULT.S_OK Then
                Dim CLSID_PrintPhotosDropTarget As New Guid("60fd46de-f830-4894-a628-6fa81bc0190d")
                Dim DropTargetType As Type = Type.GetTypeFromCLSID(CLSID_PrintPhotosDropTarget, True)
                Dim DropTarget As IDropTarget = CType(Activator.CreateInstance(DropTargetType), IDropTarget)

                Dim pdwEffect As Integer = DROPEFFECT_NONE
                Dim pt As POINT

                pt.X = 0
                pt.Y = 0

                hr = DropTarget.Drop(pdo, 0, pt, pdwEffect)
                If hr = HRESULT.S_OK Then
                    PrintEm = True
                End If
            End If
        End If

        If pidlParent <> IntPtr.Zero Then
            ILFree(pidlParent)
        End If

        For i As Integer = 0 To nIndex - 1
            If aPidl(i) <> IntPtr.Zero Then
                ILFree(aPidl(i))
            End If
        Next

        Return PrintEm
    End Function
End Class
