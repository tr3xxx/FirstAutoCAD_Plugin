Imports System.Text
Imports System.Windows.Forms
Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime

Imports Application = Autodesk.AutoCAD.ApplicationServices.Application ' to avoid ambiguity with System.Windows.Forms.Application


<Assembly: CommandClass(GetType(FirstAutoCAD_Plugin.MyCommands))>
Namespace FirstAutoCAD_Plugin

    Public Class MyCommands

        ' define a command method
        <CommandMethod("GetBlockAttributes", CommandFlags.Modal)>
        Public Sub SelectBlock()

            ' get the current doc, editor and db
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim editor As Editor = doc.Editor

            ' create a prompt for the user to select a block
            Dim prompt As New PromptEntityOptions(vbLf & "Select a block reference:")
            prompt.SetRejectMessage(vbLf & "That is not a block reference.")
            prompt.AddAllowedClass(GetType(BlockReference), False) ' only allow block references

            ' get the result of the prompt
            Dim result As PromptEntityResult = editor.GetEntity(prompt)

            If result.Status = PromptStatus.OK Then ' if the user selected a block
                ' open the block reference to get the blocks attributes 
                Using transaction As Transaction = db.TransactionManager.StartTransaction()
                    Dim obj As DBObject = transaction.GetObject(result.ObjectId, OpenMode.ForRead)

                    ' check if the selected object is a block reference (block reference is a type of block)
                    Dim blockRef As BlockReference = TryCast(obj, BlockReference)
                    If blockRef IsNot Nothing Then ' if blockRef exists call dialog method
                        ShowDialog(blockRef)
                    End If

                    transaction.Commit() ' commit and close db connection
                End Using
            Else
                editor.WriteMessage(vbLf & "Block selection cancelled.") ' print to cmd line if user cancels the selection
            End If
        End Sub

        Private Sub ShowDialog(ByVal blockRef As BlockReference)

            ' create a form using .NET Windows Forms
            Dim form As New Form() With {
                .Text = "Block Attributes",
                .Size = New System.Drawing.Size(400, 800),
                .StartPosition = FormStartPosition.CenterScreen
            }

            ' create a label as title for the block information 
            Dim blockInformationTitle As New Label() With {
                .Text = "Block Attributes",
                .Font = New System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Bold), ' set font bold bcs it should work as a subtitle 
                .AutoSize = True,
                .Location = New System.Drawing.Point(10, 10)
            }
            form.Controls.Add(blockInformationTitle) ' add label to the form

            Dim blockInfo As Dictionary(Of String, String) = GetBlockAttributes(blockRef) ' get block attributes

            Dim yPos As Integer = blockInformationTitle.Bottom + 10 ' padding from the title label

            ' itterate through the blockInfo dictionary and create a label for each key value pair
            For Each entry As KeyValuePair(Of String, String) In blockInfo
                Dim label As New Label() With {
                    .Text = $"{entry.Key}: {entry.Value}",
                    .AutoSize = True,
                    .Location = New System.Drawing.Point(10, yPos)
                }
                form.Controls.Add(label)

                yPos += label.Height + 5 ' padding between labels

            Next

            ' export button
            Dim exportButton As New Button() With {
                .Text = "Export to CSV",
                .Location = New System.Drawing.Point(10, yPos),
                .AutoSize = True
            }
            AddHandler exportButton.Click, Sub(sender, e) ExportToCsv(blockInfo) ' click event for export btn
            form.Controls.Add(exportButton)

            Dim closeButtonXPos As Integer = exportButton.Right + 10 ' calc the xPos for the close btn to be next to the export btn

            ' close button
            Dim closeButton As New Button() With {
                .Text = "Close",
                .Location = New System.Drawing.Point(closeButtonXPos, yPos),
                .AutoSize = True
            }
            AddHandler closeButton.Click, Sub(sender, e) form.Close() ' click event for close btn
            form.Controls.Add(closeButton)



            ' show the form as a modal dialog
            Application.ShowModalDialog(form)
        End Sub

        Private Function GetBlockAttributes(ByVal blockRef As BlockReference) As Dictionary(Of String, String)
            Dim blockInfo As New Dictionary(Of String, String)

            ' get basic block information
            blockInfo.Add("Block Name", blockRef.Name)
            blockInfo.Add("Position", $"X={blockRef.Position.X}, Y={blockRef.Position.Y}, Z={blockRef.Position.Z}")
            blockInfo.Add("Scale", $"X: {blockRef.ScaleFactors.X}, Y: {blockRef.ScaleFactors.Y}, Z: {blockRef.ScaleFactors.Z}")
            blockInfo.Add("Rotation", $"{blockRef.Rotation} radians")
            blockInfo.Add("Layer", blockRef.Layer)
            blockInfo.Add("Is Dynamic", If(blockRef.IsDynamicBlock, "Yes", "No"))
            blockInfo.Add("Is Visible", If(blockRef.Visible, "Yes", "No"))
            ' could add more (specific) block info here if needed 

            ' get block attributes by itterating through the block reference attribute collection
            Using transaction As Transaction = blockRef.Database.TransactionManager.StartTransaction()
                For Each attId As ObjectId In blockRef.AttributeCollection
                    Dim attRef As AttributeReference = TryCast(transaction.GetObject(attId, OpenMode.ForRead), AttributeReference)
                    If attRef IsNot Nothing Then ' if attribute exists add it to the dictionary
                        ' add the attribute tag and value to the dictionary
                        Dim attValue As String = If(String.IsNullOrWhiteSpace(attRef.TextString), "No Value Provided", attRef.TextString)

                        If Not blockInfo.ContainsKey(attRef.Tag) Then ' check if the tag already exists in the dictionary
                            blockInfo.Add(attRef.Tag, attValue)
                        End If
                    End If
                Next
                transaction.Commit() ' commit and close db connection
            End Using

            Return blockInfo
        End Function

        Private Sub ExportToCsv(ByVal blockInfo As Dictionary(Of String, String))
            ' create a save file dialog to allow the user to select a file path
            Using saveBlockAttributes As New SaveFileDialog()
                ' set the save file dialog properties to csv and set the initial directory to the users doc folder
                With saveBlockAttributes
                    .Filter = "CSV files (*.csv)|*.csv"
                    .Title = "Save block attributes to CSV"
                    .FileName = $"BlockAttributes.csv"
                    .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                End With

                If saveBlockAttributes.ShowDialog() = DialogResult.OK Then ' if the user selects a file path

                    Try
                        Dim csv As New StringBuilder() ' create a string builder to store the csv content

                        csv.AppendLine("Tag;Value") ' column headers

                        ' iterate through blockInfo and add a csv line for each key value pair
                        For Each entry As KeyValuePair(Of String, String) In blockInfo
                            ' Escape each value to handle commas and newlines
                            Dim line As String = $"""{entry.Key}"";""{entry.Value}""" ' ";" puts the value in the next column
                            csv.AppendLine(line)
                        Next

                        ' write the csv to the file
                        File.WriteAllText(saveBlockAttributes.FileName, csv.ToString())

                        ' show a message box to confirm the export
                        MessageBox.Show("Block attributes sucessfully exported to CSV", "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Catch ex As Exception
                        ' show a message box if an error occurs
                        MessageBox.Show($"An error occured while exporting: {ex.Message}", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End Using
        End Sub
    End Class

End Namespace