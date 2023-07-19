Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Public Class ClassAsyncHelpers

    Public Shared Function GetCheckedListBox(ByRef CB As CheckedListBox) As List(Of String)
        Dim str As New List(Of String)
        For Each item In CB.CheckedItems
            str.Add(item)
        Next
        Return str
    End Function

    Public Shared Async Function GetCheckedListBoxAsync(ByVal CB As CheckedListBox) As Task(Of List(Of String))
        Dim str As List(Of String) = Await Task.Run(Function() GetCheckedListBox(CB))
        Return str
    End Function

    Public Shared Function GetListBox(ByRef CB As ListBox) As List(Of String)
        Dim str As New List(Of String)
        For Each item In CB.Items
            str.Add(item)
        Next
        Return str
    End Function

    Public Shared Async Function GetListBoxAsync(ByVal CB As ListBox) As Task(Of List(Of String))
        Dim str As List(Of String) = Await Task.Run(Function() GetListBox(CB))
        Return str
    End Function

    Public Shared Sub LoadCheckedListBox(ByRef CB As CheckedListBox, ByRef LstStr As List(Of String))
        For Each item In LstStr
            CB.Items.Add(item)
        Next
    End Sub

    Public Shared Async Sub LoadCheckedListBoxAsync(ByVal CB As CheckedListBox, ByVal LstStr As List(Of String))
        Await Task.Run(Sub() LoadCheckedListBox(CB, LstStr))
    End Sub

    Public Shared Sub LoadListBox(ByRef CB As ListBox, ByRef LstStr As List(Of String))
        For Each item In LstStr
            CB.Items.Add(item)
        Next
    End Sub

    Public Shared Async Sub LoadListBoxAsync(ByVal CB As ListBox, ByVal LstStr As List(Of String))
        Await Task.Run(Sub() LoadListBox(CB, LstStr))
    End Sub

    Public Function ExtractLinks(ByVal url As String) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("LinkText")
        dt.Columns.Add("LinkUrl")

        Dim wc As New WebClient
        Dim html As String = wc.DownloadString(url)

        Dim links As MatchCollection = Regex.Matches(html, "<a.*?href=""(.*?)"".*?>(.*?)</a>")

        For Each match As Match In links
            Dim dr As DataRow = dt.NewRow
            Dim matchUrl As String = match.Groups(1).Value
            'Ignore all anchor links
            If matchUrl.StartsWith("#") Then
                Continue For
            End If
            'Ignore all java script calls
            If matchUrl.ToLower.StartsWith("javascript:") Then
                Continue For
            End If
            'Ignore all email links
            If matchUrl.ToLower.StartsWith("mailto:") Then
                Continue For
            End If
            'For internal links, build the url mapped to the base address
            If Not matchUrl.StartsWith("http://") And Not matchUrl.StartsWith("https://") Then
                matchUrl = MapUrl(url, matchUrl)
            End If
            'Add the link data to datatable
            dr("LinkUrl") = matchUrl
            dr("LinkText") = match.Groups(2).Value
            dt.Rows.Add(dr)
        Next

        Return dt
    End Function

    Public Function MapUrl(ByVal baseAddress As String, ByVal relativePath As String) As String

        Dim u As New System.Uri(baseAddress)

        If relativePath = "./" Then
            relativePath = "/"
        End If

        If relativePath.StartsWith("/") Then
            Return u.Scheme + Uri.SchemeDelimiter + u.Authority + relativePath
        Else
            Dim pathAndQuery As String = u.AbsolutePath
            ' If the baseAddress contains a file name, like ..../Something.aspx
            ' Trim off the file name
            pathAndQuery = pathAndQuery.Split("?")(0).TrimEnd("/")
            If pathAndQuery.Split("/")(pathAndQuery.Split("/").Count - 1).Contains(".") Then
                pathAndQuery = pathAndQuery.Substring(0, pathAndQuery.LastIndexOf("/"))
            End If
            baseAddress = u.Scheme + Uri.SchemeDelimiter + u.Authority + pathAndQuery

            'If the relativePath contains ../ then
            ' adjust the baseAddress accordingly

            While relativePath.StartsWith("../")
                relativePath = relativePath.Substring(3)
                If baseAddress.LastIndexOf("/") > baseAddress.IndexOf("//" + 2) Then
                    baseAddress = baseAddress.Substring(0, baseAddress.LastIndexOf("/")).TrimEnd("/")
                End If
            End While

            Return baseAddress + "/" + relativePath
        End If

    End Function

    Private Sub browserImageLoad_DocumentCompleted(ByVal sender As Object, ByVal e As WebBrowserDocumentCompletedEventArgs)
        Dim browser As WebBrowser = CType(sender, WebBrowser)
        Dim collection As HtmlElementCollection
        Dim imgListString As List(Of HtmlElement) = New List(Of HtmlElement)()

        If browser IsNot Nothing Then

            If browser.Document IsNot Nothing Then
                collection = browser.Document.GetElementsByTagName("img")

                If collection IsNot Nothing Then

                    For Each element As HtmlElement In collection
                        Dim wClient As WebClient = New WebClient()
                        Dim urlDownload As String = element.GetAttribute("src")
                        wClient.DownloadFile(urlDownload, urlDownload.Substring(urlDownload.LastIndexOf("/"c)))
                    Next
                End If
            End If
        End If
    End Sub

    Private Sub ExportDatasetToCsv(ByVal MyDataSet As DataSet, ByRef filname As String)
        Dim myWriter As New System.IO.StreamWriter(filname)
        Try
            'Declaration of Variables
            Dim dt As DataTable
            '     Dim dr As DataRow
            Dim myString As String = ""
            Dim bFirstRecord As Boolean = True
            myString = ""
            For Each dt In MyDataSet.Tables
                For Each dr As DataRow In dt.Rows
                    bFirstRecord = True
                    For Each field As Object In dr.ItemArray

                        If Not bFirstRecord Then

                            myString &= ","

                        End If

                        myString &= field.ToString

                        bFirstRecord = False

                    Next
                    'New Line to differentiate next row
                    myString &= Environment.NewLine
                Next
            Next
            'Write the String to the Csv File
            myWriter.WriteLine(myString)
        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
        'Clean up
        myWriter.Close()
    End Sub

End Class