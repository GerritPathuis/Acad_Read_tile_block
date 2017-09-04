Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.IO

Public Class Form1
    Dim filepath1 As String = "C:\Repos\Acad_Read_tile_block\Acad_Read_title"
    Dim filepath2 As String = "C:\Repos\Acad_Read_tile_block\Acad_Read_title\test.dwg"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FilterMtextWildcard()
    End Sub

    Public Sub FilterMtextWildcard()
        'See http://help.autodesk.com/view/ACD/2016/ENU/?guid=GUID-3C1A759C-BB88-41A7-B1DE-697C493C92C8
        '' Get the current document editor
        Dim acDocEd As Editor = Application.DocumentManager.MdiActiveDocument.Editor

        '' Create a TypedValue array to define the filter criteria
        Dim acTypValAr(1) As TypedValue
        acTypValAr.SetValue(New TypedValue(DxfCode.Start, "MTEXT"), 0)
        acTypValAr.SetValue(New TypedValue(DxfCode.Text, "*The*"), 1)

        '' Assign the filter criteria to a SelectionFilter object
        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)

        '' Request for objects to be selected in the drawing area
        Dim acSSPrompt As PromptSelectionResult
        acSSPrompt = acDocEd.GetSelection(acSelFtr)

        '' If the prompt status is OK, objects were selected
        If acSSPrompt.Status = PromptStatus.OK Then
            Dim acSSet As SelectionSet = acSSPrompt.Value
            Application.ShowAlertDialog("Number of objects selected: " & acSSet.Count.ToString())
        Else
            Application.ShowAlertDialog("Number of objects selected: 0")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        SaveActiveDrawing()
    End Sub

    Public Sub SaveActiveDrawing()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim strDWGName As String = acDoc.Name

        Dim obj As Object = Application.GetSystemVariable("DWGTITLED")

        '' Check to see if the drawing has been named
        If System.Convert.ToInt16(obj) = 0 Then
            '' If the drawing is using a default name (Drawing1, Drawing2, etc)
            '' then provide a new name
            strDWGName = "c:\MyDrawing.dwg"
        End If

        '' Save the active drawing
        acDoc.Database.SaveAs(strDWGName, True, DwgVersion.Current, acDoc.Database.SecurityParameters)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        GetDrawingInfo()
    End Sub
    Public Sub GetDrawingInfo()
        Dim objArray As ArrayList = New ArrayList
        Dim ed As Editor = Core.Application.DocumentManager.MdiActiveDocument.Editor

        ' Create an OpenFileDialog object.
        ' Initialize the OpenFileDialog to look for DWG files.
        Dim openFile1 As OpenFileDialog = New OpenFileDialog With {
            .InitialDirectory = filepath1,
            .Filter = "Drawing Files|*.dwg",
            .Multiselect = True
        }
        MessageBox.Show("99")
        ' Check if the user selected a file from the OpenFileDialog.
        If openFile1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            If openFile1.FileNames.Length > 0 Then
                Dim strFile As String
                For Each strFile In openFile1.FileNames
                    Using db As Database = New Database(False, True)
                        Try
                            db.ReadDwgFile(strFile, System.IO.FileShare.Read, False, Nothing)
                        Catch UnableToReadExeption As System.Exception
                            ed.WriteMessage(String.Format(vbCrLf & "Can't to read the drawing {0}", strFile))
                            Return
                        End Try
                        objArray.Add("------------------------------")
                        objArray.Add("File processed: " & strFile)
                        objArray.Add("------------------------------")
                        Using trans As Transaction = db.TransactionManager.StartTransaction()
                            ' Open the blocktable, get the modelspace
                            Try
                                Dim bt As BlockTable = DirectCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                                Dim btr As BlockTableRecord = DirectCast(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
                                Dim id As ObjectId

                                For Each id In btr

                                    Dim ent As Entity = CType(trans.GetObject(id, OpenMode.ForRead, False), Entity)
                                    Dim objName As String = ent.GetRXClass().Name.ToString()
                                    objArray.Add(objName)
                                Next
                                btr.Dispose()
                            Catch ex As System.Exception
                                MessageBox.Show("Unexpected Error: " + ex.ToString())
                            End Try
                            trans.Commit()
                        End Using
                        db.Dispose()
                    End Using
                Next
            End If
        End If

        For i As Integer = 0 To objArray.Count - 1
            ed.WriteMessage(vbCrLf & objArray(i).ToString)
        Next
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim strFileName As String = filepath2

        Dim acDocMgr As DocumentCollection = Application.DocumentManager

        If (File.Exists(strFileName)) Then
            DocumentCollectionExtension.Open(acDocMgr, strFileName, False)
        Else
            acDocMgr.MdiActiveDocument.Editor.WriteMessage("File " & strFileName & " does not exist.")
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim strDWGName As String = acDoc.Name

        Dim obj As Object = Application.GetSystemVariable("DWGTITLED")

        '' Check to see if the drawing has been named
        If System.Convert.ToInt16(obj) = 0 Then
            '' If the drawing is using a default name (Drawing1, Drawing2, etc)
            '' then provide a new name
            strDWGName = filepath1 & "\" & RandomString(6) & ".dwg"
        End If

        '' Save the active drawing
        acDoc.Database.SaveAs(strDWGName, True, DwgVersion.Current, acDoc.Database.SecurityParameters)
    End Sub

    Public Function RandomString(ByVal length As Integer) As String
        Dim sb As New System.Text.StringBuilder
        Dim chars() As String = {"a", "b", "c", "d", "e", "f", "g", "h", "i",
        "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w",
        "x", "y", "z"}
        Dim upperBound As Integer = UBound(chars)

        For x As Integer = 1 To length
            sb.Append(chars(Int(Rnd() * upperBound)))
        Next
        Return sb.ToString
    End Function
End Class
