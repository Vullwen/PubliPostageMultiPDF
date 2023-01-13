Attribute VB_Name = "NewMacros"
Sub PublipostageMultiPDF()

    ' ouverture folderbrowser
    Dim sFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then
            sFolder = .SelectedItems(1)
        End If
    End With
    If sFolder <> "" Then
        ' var
        Dim docmult As Document, docseul As Document, dernierNumEnr As Long
        Set multidoc = ActiveDocument
        
        multidoc.MailMerge.DataSource.ActiveRecord = wdLastRecord
        dernierNumEnr = multidoc.MailMerge.DataSource.ActiveRecord
        multidoc.MailMerge.DataSource.ActiveRecord = wdFirstRecord
        
        ' cr?ation des docs
        Do While dernierNumEnr > 0
            multidoc.MailMerge.Destination = wdSendToNewDocument
            multidoc.MailMerge.DataSource.FirstRecord = multidoc.MailMerge.DataSource.ActiveRecord
            multidoc.MailMerge.DataSource.LastRecord = multidoc.MailMerge.DataSource.ActiveRecord
            multidoc.MailMerge.Execute False
            Set docseul = ActiveDocument
            
            docseul.ExportAsFixedFormat _
                OutputFileName:=sFolder & Application.PathSeparator & _
                    multidoc.MailMerge.DataSource.DataFields("VIN").Value & ".pdf", _
                ExportFormat:=wdExportFormatPDF
                
            docseul.Close False
            
            If multidoc.MailMerge.DataSource.ActiveRecord >= dernierNumEnr Then
                dernierNumEnr = 0
            Else
                multidoc.MailMerge.DataSource.ActiveRecord = wdNextRecord
            End If

        Loop
    End If
    
    
End Sub

