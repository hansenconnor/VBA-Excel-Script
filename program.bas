'Rextester.Program.Main is the entry point for your code. Don't change it.
'Compiler version 11.0.50709.17929 for Microsoft (R) .NET Framework 4.5

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text.RegularExpressions

Namespace Rextester
    Public Module Program
        Public Sub Main(args() As string)

            'Prompt for table name to get number of rows
            Sub CountTableRows()
                Dim xTable As ListObject
                Dim xTName As String
                Dim totalRows as Int
                On Error Resume Next
                xTName = Application.InputBox("Please input the table name：", "KuTools For Excel", , , , , , 2)
                Set xTable = ActiveSheet.ListObjects(xTName)
                totalRows = xTable.Range.Rows.Count 'Total Row count

                'Begin loop through all rows
                For i = 0 To totalRows Step 1
                    'Check if column is 1
                     if ActiveSheet.Cells(row, "b") == 1 Then
                     'Success!
                     Rows(totalRows).Insert Shift:=xlDown 'Insert an empty row
                         With Sheets(xTName)
                            .Cells(totalRows + 1, "a").Value = "yes"
                            totalRows = totalRows + 1; 'Increment totalRows as we just added another row
                         End With
                     end if
                Next
                Set xTable = Nothing
            End Sub


            Sub insertRow();
            Worksheets(“Insert row”).Rows(6).Insert Shift:=xlShiftDown;

            Worksheets.Rows(row#).Insert;
        End Sub

        Sub Range_Find_Method()
        'Finds the last non-blank cell on a sheet/range.
        Dim lRow As Long
        Dim lCol As Long

            lRow = Cells.Find(What:="*", _
                            After:=Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row

            MsgBox "Last Row: " & lRow
        End Sub


        Sub LastRow()

        'Find the last row of active worksheet in column 1 (aka col A)
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

        End Sub



        Dim a As Integer
           a = 10

           For i = 0 To a Step 2
              MsgBox "The value is i is : " & i
           Next




    End Module
End Namespace
