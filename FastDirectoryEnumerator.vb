Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Runtime.ConstrainedExecution
Imports System.Runtime.InteropServices
Imports System.Security.Permissions
Imports Microsoft.Win32.SafeHandles



<Serializable>
Public Class FileData
    Public ReadOnly Attributes As FileAttributes

    Public ReadOnly Property CreationTime As DateTime
        Get
            Return Me.CreationTimeUtc.ToLocalTime()
        End Get
    End Property

    Public ReadOnly CreationTimeUtc As DateTime

    Public ReadOnly Property LastAccesTime As DateTime
        Get
            Return Me.LastAccessTimeUtc.ToLocalTime()
        End Get
    End Property

    Public ReadOnly LastAccessTimeUtc As DateTime

    Public ReadOnly Property LastWriteTime As DateTime
        Get
            Return Me.LastWriteTimeUtc.ToLocalTime()
        End Get
    End Property

    Public ReadOnly LastWriteTimeUtc As DateTime
    Public ReadOnly Size As Long
    Public ReadOnly Name As String
    Public ReadOnly Path As String

    Public Overrides Function ToString() As String
        Return Me.Name
    End Function

    Friend Sub New(ByVal dir As String, ByVal findData As WIN32_FIND_DATA)
        Me.Attributes = findData.dwFileAttributes
        Me.CreationTimeUtc = ConvertDateTime(findData.ftCreationTime_dwHighDateTime, findData.ftCreationTime_dwLowDateTime)
        Me.LastAccessTimeUtc = ConvertDateTime(findData.ftLastAccessTime_dwHighDateTime, findData.ftLastAccessTime_dwLowDateTime)
        Me.LastWriteTimeUtc = ConvertDateTime(findData.ftLastWriteTime_dwHighDateTime, findData.ftLastWriteTime_dwLowDateTime)
        Me.Size = CombineHighLowInts(findData.nFileSizeHigh, findData.nFileSizeLow)
        Me.Name = findData.cFileName
        Me.Path = IO.Path.Combine(dir, findData.cFileName)
    End Sub

    Private Shared Function CombineHighLowInts(ByVal high As UInteger, ByVal low As UInteger) As Long
        Return (CLng(high) << &H20) Or low
    End Function

    Private Shared Function ConvertDateTime(ByVal high As UInteger, ByVal low As UInteger) As DateTime
        Dim fileTime As Long = CombineHighLowInts(high, low)
        Return DateTime.FromFileTimeUtc(fileTime)
    End Function
End Class



<Serializable, StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto), BestFitMapping(False)>
Friend Class WIN32_FIND_DATA
    Public dwFileAttributes As FileAttributes
    Public ftCreationTime_dwLowDateTime As UInteger
    Public ftCreationTime_dwHighDateTime As UInteger
    Public ftLastAccessTime_dwLowDateTime As UInteger
    Public ftLastAccessTime_dwHighDateTime As UInteger
    Public ftLastWriteTime_dwLowDateTime As UInteger
    Public ftLastWriteTime_dwHighDateTime As UInteger
    Public nFileSizeHigh As UInteger
    Public nFileSizeLow As UInteger
    Public dwReserved0 As Integer
    Public dwReserved1 As Integer
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
    Public cFileName As String
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=14)>
    Public cAlternateFileName As String

    Public Overrides Function ToString() As String
        Return "File name=" & cFileName
    End Function
End Class



Module FastDirectoryEnumerator

    Public Function EnumerateFiles(ByVal path As String) As IEnumerable(Of FileData)
        Return FastDirectoryEnumerator.EnumerateFiles(path, "*")
    End Function

    Public Function EnumerateFiles(ByVal path As String, ByVal searchPattern As String) As IEnumerable(Of FileData)
        Return FastDirectoryEnumerator.EnumerateFiles(path, searchPattern, SearchOption.TopDirectoryOnly)
    End Function

    Function EnumerateFiles(ByVal thePath As String, ByVal searchPattern As String, ByVal searchOption As SearchOption) As IEnumerable(Of FileData)
        If thePath Is Nothing Then
            Throw New ArgumentNullException(NameOf(thePath))
        End If

        If searchPattern Is Nothing Then
            Throw New ArgumentNullException(NameOf(searchPattern))
        End If

        If (searchOption <> SearchOption.TopDirectoryOnly) AndAlso (searchOption <> SearchOption.AllDirectories) Then
            Throw New ArgumentOutOfRangeException(NameOf(searchOption))
        End If

        Dim fullPath As String = Path.GetFullPath(thePath)
        Return New FileEnumerable(fullPath, searchPattern, searchOption)
    End Function

    Function GetFiles(ByVal path As String, ByVal searchPattern As String, ByVal searchOption As SearchOption) As FileData()
        Dim e As IEnumerable(Of FileData) = FastDirectoryEnumerator.EnumerateFiles(path, searchPattern, searchOption)
        Dim list As List(Of FileData) = New List(Of FileData)(e)
        Dim retval As FileData() = New FileData(list.Count - 1) {}
        list.CopyTo(retval)
        Return retval
    End Function



    Private Class FileEnumerable
        Implements IEnumerable(Of FileData)

        Private ReadOnly m_path As String
        Private ReadOnly m_filter As String
        Private ReadOnly m_searchOption As SearchOption

        Public Sub New(ByVal path As String, ByVal filter As String, ByVal searchOption As SearchOption)
            m_path = path
            m_filter = filter
            m_searchOption = searchOption
        End Sub

        Public Function GetEnumerator() As IEnumerator(Of FileData) Implements IEnumerable(Of FileData).GetEnumerator
            Return New FileEnumerator(m_path, m_filter, m_searchOption)
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return New FileEnumerator(m_path, m_filter, m_searchOption)
        End Function
    End Class



    Private NotInheritable Class SafeFindHandle
        Inherits SafeHandleZeroOrMinusOneIsInvalid

        <ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)>
        <DllImport("kernel32.dll")>
        Private Shared Function FindClose(ByVal handle As IntPtr) As Boolean

        End Function

        <SecurityPermission(SecurityAction.LinkDemand, UnmanagedCode:=True)>
        Friend Sub New()
            MyBase.New(True)
        End Sub

        Protected Overrides Function ReleaseHandle() As Boolean
            Return FindClose(MyBase.handle)
        End Function
    End Class



    <System.Security.SuppressUnmanagedCodeSecurity>
    Private Class FileEnumerator
        Implements IEnumerator(Of FileData)

        <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function FindFirstFile(<MarshalAs(UnmanagedType.LPWStr)> fileName As String, <[In], Out> data As WIN32_FIND_DATA) As SafeFindHandle
        End Function

        <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function FindNextFile(hndFindFile As SafeFindHandle, <[In], Out, MarshalAs(UnmanagedType.LPStruct)> lpFindFileData As WIN32_FIND_DATA) As Boolean
        End Function



        Private Class SearchContext
            Public ReadOnly Path As String
            Public SubdirectoriesToProcess As Stack(Of String)

            Public Sub New(ByVal path As String)
                Me.Path = path
            End Sub
        End Class



        Private m_path As String
        Private ReadOnly m_filter As String
        Private ReadOnly m_searchOption As SearchOption
        Private ReadOnly m_contextStack As Stack(Of SearchContext)
        Private m_currentContext As SearchContext
        Private m_hndFindFile As SafeFindHandle
        Private ReadOnly m_win_find_data As WIN32_FIND_DATA = New WIN32_FIND_DATA()

        Public Sub New(ByVal path As String, ByVal filter As String, ByVal searchOption As SearchOption)
            m_path = path
            m_filter = filter
            m_searchOption = searchOption
            m_currentContext = New SearchContext(path)

            If m_searchOption = SearchOption.AllDirectories Then
                m_contextStack = New Stack(Of SearchContext)()
            End If
        End Sub

        Public ReadOnly Property Current As FileData Implements IEnumerator(Of FileData).Current
            Get
                Return New FileData(m_path, m_win_find_data)
            End Get
        End Property

        Private ReadOnly Property IEnumerator_Current As Object Implements IEnumerator.Current
            Get
                Return Current
            End Get
        End Property

        Public Sub Dispose() Implements IDisposable.Dispose
            If m_hndFindFile IsNot Nothing Then
                m_hndFindFile.Dispose()
            End If
        End Sub

        Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
            Dim retval As Boolean = False

            If m_currentContext.SubdirectoriesToProcess Is Nothing Then

                If m_hndFindFile Is Nothing Then
                    With New FileIOPermission(FileIOPermissionAccess.PathDiscovery, m_path)
                        .Demand()
                    End With

                    Dim searchPath As String = Path.Combine(m_path, m_filter)
                    m_hndFindFile = FindFirstFile(searchPath, m_win_find_data)
                    retval = Not m_hndFindFile.IsInvalid
                Else
                    retval = FindNextFile(m_hndFindFile, m_win_find_data)
                End If
            End If

            If retval Then
                If (CType(m_win_find_data.dwFileAttributes, FileAttributes) And FileAttributes.Directory) = FileAttributes.Directory Then
                    Return MoveNext()
                End If
            ElseIf m_searchOption = SearchOption.AllDirectories Then

                If m_currentContext.SubdirectoriesToProcess Is Nothing Then
                    Dim subDirectories As String() = Directory.GetDirectories(m_path)
                    m_currentContext.SubdirectoriesToProcess = New Stack(Of String)(subDirectories)
                End If

                If m_currentContext.SubdirectoriesToProcess.Count > 0 Then
                    Dim subDir As String = m_currentContext.SubdirectoriesToProcess.Pop()
                    m_contextStack.Push(m_currentContext)
                    m_path = subDir
                    m_hndFindFile = Nothing
                    m_currentContext = New SearchContext(m_path)
                    Return MoveNext()
                End If

                If m_contextStack.Count > 0 Then
                    m_currentContext = m_contextStack.Pop()
                    m_path = m_currentContext.Path

                    If m_hndFindFile IsNot Nothing Then
                        m_hndFindFile.Close()
                        m_hndFindFile = Nothing
                    End If

                    Return MoveNext()
                End If
            End If

            Return retval
        End Function

        Public Sub Reset() Implements IEnumerator.Reset
            m_hndFindFile = Nothing
        End Sub
    End Class
End Module
