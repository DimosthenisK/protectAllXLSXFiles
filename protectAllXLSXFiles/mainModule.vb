Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

Module XLSXProtect

    Sub Main()
        Dim totalFiles As Integer = 0
        Dim totalSuccesses As Integer = 0
        Dim xlApp As New Excel.Application
        xlApp.DisplayAlerts = False
        Dim xlWorkBooks As Excel.Workbooks = xlApp.Workbooks
        Console.WriteLine("Mass XLSX File Protector")
        Console.WriteLine("Insert password to be used for files: ")
        Dim password As String = Console.ReadLine()
        Console.WriteLine("Insert the folder path that contains the XLSX files (subfolders will be used, too): ")
        Dim path As String = Console.ReadLine()

        Console.WriteLine()
        Console.WriteLine("---------------------------------------------------------------------")
        Console.WriteLine()

        Dim filesToProtect As String() = Directory.GetFiles(path, "*.xlsx", SearchOption.AllDirectories)
        For Each file In filesToProtect
            totalFiles += 1
            Console.WriteLine("File found : " + file)
            Console.WriteLine("Protect it? (Y/y/Yes/yes-N/n/No/no)")
            Dim answer As String = Console.ReadLine()
            If answer = "Y" Or answer = "y" Or answer = "Yes" Or answer = "yes" Then
                Dim cWorkBook As Excel.Workbook = xlWorkBooks.Open(file)
                cWorkBook.SaveAs(Filename:=cWorkBook.FullName, Password:=password)

                Console.WriteLine("File protected successfully")
                totalSuccesses += 1
            ElseIf answer = "N" Or answer = "n" Or answer = "No" Or answer = "no" Then
                Console.WriteLine("File passed.")
            End If
            Console.WriteLine()
        Next
        Console.WriteLine("Finish")
        Console.WriteLine("Found {0} files in total and protected {1} of them.", totalFiles, totalSuccesses)

        Console.ReadLine()
    End Sub

End Module
