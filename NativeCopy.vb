Imports System.Runtime.InteropServices

Public Class NativeCopy

    Private Enum FO_Func As Short
        FO_COPY   = &H2
        FO_DELETE = &H3
        FO_MOVE   = &H1
        FO_RENAME = &H4
    End Enum

    Private Structure SHFILEOPSTRUCT

        Public hwnd  As IntPtr
        Public wFunc As FO_Func

        <MarshalAs(UnmanagedType.LPWStr)>
        Public pFrom As String

        <MarshalAs(UnmanagedType.LPWStr)>
        Public pTo                   As String
        Public fFlags                As UShort
        Public fAnyOperationsAborted As Boolean
        Public hNameMappings         As IntPtr

        <MarshalAs(UnmanagedType.LPWStr)>
        Public lpszProgressTitle As String

    End Structure

    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function SHFileOperation(
       <[In]> ByRef lpFileOp As SHFILEOPSTRUCT) As Integer
    End Function

    Private Shared _ShFile As SHFILEOPSTRUCT

    Public Shared Sub Copy(ByVal sSource As List(Of String), ByVal sTarget As String)
        _ShFile.wFunc = FO_Func.FO_COPY
        _ShFile.pFrom = String.Join(vbNullChar, sSource) + vbNullChar
        _ShFile.pTo   = sTarget
        SHFileOperation(_ShFile)
   End Sub

End Class
