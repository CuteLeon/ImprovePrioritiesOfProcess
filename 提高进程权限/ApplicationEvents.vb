Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic.ApplicationServices

Namespace My
    Partial Friend Class MyApplication
        Private Const TOKEN_QUERY As Integer = 8
        Private Const TOKEN_ADJUST_PRIVILEGES As Integer = 32
        Private Const SE_DEBUG_NAME As String = "SeDebugPrivilege"
        Private Const SE_SHUTDOWN_NAME As String = "SeShutdownPrivilege"
        Private Const ANYSIZE_ARRAY As Integer = 1
        Private Const SE_PRIVILEGE_ENABLED As Integer = 2
        Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As IntPtr
        Private Declare Function OpenProcessToken Lib "advapi32.dll" (ProcessHandle As Integer, DesiredAccess As Integer, ByRef TokenHandle As Integer) As Integer
        Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, ByRef lpLuid As LARGE_INTEGER) As Integer

        Private Declare Function CloseHandle Lib "kernel32" (hObject As Integer) As Integer
        Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" Alias "AdjustTokenPrivileges" (ByVal TokenHandle As Integer, ByVal DisableAllPrivileges As Integer, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Integer, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef ReturnLength As Integer) As Integer
        Public Declare Function GetLastError Lib "kernel32.dll" () As Integer

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
        Private Structure LARGE_INTEGER
            Public LowPart As Integer
            Public HighPart As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
        Private Structure LUID_AND_ATTRIBUTES
            Public pLuid As LARGE_INTEGER
            Public Attributes As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
            Private Structure TOKEN_PRIVILEGES
            Public PrivilegeCount As Integer
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=ANYSIZE_ARRAY)>
            Public Privileges As LUID_AND_ATTRIBUTES()
        End Structure

        Private Sub MyApplication_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup
            Dim hToken As Integer = 0
            Dim CurrentProcess As Integer = GetCurrentProcess()
            If CurrentProcess <> 0 Then
                MsgBox(String.Format("当前进程虚拟句柄：" & CurrentProcess))
            Else
                MsgBox(String.Format("获取当前进程虚拟句柄失败,错误代码：" & GetLastError()))
                Exit Sub
            End If

            Dim htRet As Integer = OpenProcessToken(CurrentProcess, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken)
            If hToken <> 0 Then
                MsgBox(String.Format("打开进程令牌成功，句柄：" & hToken))
            Else
                MsgBox(String.Format("打开进程令牌失败,错误代码：" & GetLastError()))
                Exit Sub
            End If

            Dim sedebugnameValue As LARGE_INTEGER = New LARGE_INTEGER
            Dim LookRet As Integer = LookupPrivilegeValue(Nothing, SE_DEBUG_NAME, sedebugnameValue) '第二个参数是目标权限名称
            If LookRet <> 0 Then
                MsgBox(String.Format("获取系统特权LUID成功！" & sedebugnameValue.LowPart & " - " & sedebugnameValue.HighPart))
            Else
                MsgBox(String.Format("获取系统特权LUID失败,错误代码：" & GetLastError()))
                Exit Sub
            End If
            Dim [TO] As TOKEN_PRIVILEGES = Nothing
            Dim LAA As LUID_AND_ATTRIBUTES = Nothing
            LAA.pLuid = sedebugnameValue
            LAA.Attributes = 2
            [TO].PrivilegeCount = 1
            [TO].Privileges = New LUID_AND_ATTRIBUTES() {LAA}
            Dim SizeOf As Integer = Marshal.SizeOf([TO])
            Dim TO2 As TOKEN_PRIVILEGES = Nothing
            Dim SizeOf2 As Integer = 0

            '函数返回值表示函数执行成功或失败，不表示操作成功或失败
            Dim ATPRes As Integer = AdjustTokenPrivileges(hToken, 0, [TO], SizeOf, TO2, SizeOf2)
            Dim ATPErr As Integer = GetLastError()
            If ATPErr = 0 Then
                MsgBox(String.Format("已提高进程权限"))
            Else
                MsgBox(String.Format("提高进程权限失败,错误代码：" & ATPErr))
                Exit Sub
            End If
        End Sub
    End Class
End Namespace
