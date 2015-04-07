'######################################################################
'######################################################################

Public objFileSystem, objFile, strFilePath, DelimitChar


Public Function Txt_file_read()

        RowCount = 1
        
         On Error Resume Next
         Set objFileSystem = CreateObject("Scripting.FileSystemObject")
         
         strFilePath = "d:\test_file.txt"
         
         Set objFile = objFileSystem.OpenTextFile(strFilePath)
         DelimitChar = "|"
        
         
        
        Do Until objFile.AtEndOfStream
        
            strLine = objFile.ReadLine        
            
            
            Txt_value = Split(strLine, DelimitChar)
            
            
                    For I_counter = 0 To UBound(Txt_value)
                    
						ThisWorkbook.Sheets(1).Cells(RowCount, I_counter + 1).Value = Txt_value(I_counter)
                    
                    Next I_counter
            
            
            RowCount = RowCount + 1
        
        Loop
        
        
        ThisWorkbook.Save
        
        objFileSystem.Close
        
        objFileSystem = Nothing
        objFile = Nothing

End Function

