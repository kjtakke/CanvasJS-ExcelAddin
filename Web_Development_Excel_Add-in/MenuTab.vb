'Face ID Icon List
'https://bettersolutions.com/vba/ribbon/face-ids-2003.htm

Private Sub Workbook_Open() 
    Dim cmbBar As CommandBar 
    Dim cmbControl As CommandBarControl 
     
    Set cmbBar = Application.CommandBars("Worksheet Menu Bar") 
    Set cmbControl = cmbBar.Controls.Add(Type:=msoControlPopup, temporary:=True) 'adds a menu item to the Menu Bar
    
    With cmbControl 
        .Caption = "&Site Development" 'names the menu item
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "New" 'adds a description to the menu item
            .OnAction = "NewXLWD" 'runs the specified macro
            .FaceId = 3044 'assigns an icon to the dropdown
        End With 
        With .Controls.Add(Type:=msoControlButton) 
            .Caption = "Open" 
            .OnAction = "OpenXLWD" 
            .FaceId = 9161 
        End With 
        With .Controls.Add(Type:=msoControlButton) 
            .Caption = "Edit" 
            .OnAction = "EditXLWD" 
            .FaceId = 9115 
        End With        
        With .Controls.Add(Type:=msoControlButton) 
            .Caption = "Save" 
            .OnAction = "SaveXLWD" 
            .FaceId = 3104 
        End With 
        With .Controls.Add(Type:=msoControlButton) 
            .Caption = "SAve As" 
            .OnAction = "SaveAsXLWD" 
            .FaceId = 1713 
        End With 
    End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Deploy" 'adds a description to the menu item
            .OnAction = "DeployXLWD" 'runs the specified macro
            .FaceId = 7424 'assigns an icon to the dropdown
        End With 
        With .Controls.Add(Type:=msoControlButton) 
            .Caption = "Developer Notes" 
            .OnAction = "DeveloperNotesXLWD" 
            .FaceId = 3535 
        End With 
        With .Controls.Add(Type:=msoControlButton) 
            .Caption = "Eror Logs" 
            .OnAction = "ErrorLogsXLWD" 
            .FaceId = 9326 
        End With   
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Copy Code" 'adds a description to the menu item
            .OnAction = "CopyCodeXLWD" 'runs the specified macro
            .FaceId = 22 'assigns an icon to the dropdown
        End With 
    End With 
   
    With cmbControl 
        .Caption = "&Single Pages" 'names the menu item   
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Single Table" 'adds a description to the menu item
            .OnAction = "XLSTD" 'runs the specified macro
            .FaceId = 800 'assigns an icon to the dropdown
        End With    
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Singe Chart" 'adds a description to the menu item
            .OnAction = "XLSCD" 'runs the specified macro
            .FaceId = 45 'assigns an icon to the dropdown
        End With     
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Chart and Table" 'adds a description to the menu item
            .OnAction = "XLCTD" 'runs the specified macro
            .FaceId = 641 'assigns an icon to the dropdown
        End With 
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Single Image" 'adds a description to the menu item
            .OnAction = "XLSID" 'runs the specified macro
            .FaceId = 6578 'assigns an icon to the dropdown
        End With         
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Image Slider" 'adds a description to the menu item
            .OnAction = "XLISD" 'runs the specified macro
            .FaceId = 6185 'assigns an icon to the dropdown
        End With         
    End With    
    
    With cmbControl 
        .Caption = "&Other Pages" 'names the menu item  
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Profile" 'adds a description to the menu item
            .OnAction = "XLPD" 'runs the specified macro
            .FaceId = 9035 'assigns an icon to the dropdown
        End With          
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Resume" 'adds a description to the menu item
            .OnAction = "XLRD" 'runs the specified macro
            .FaceId = 9196 'assigns an icon to the dropdown
        End With            
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Business" 'adds a description to the menu item
            .OnAction = "XLBD" 'runs the specified macro
            .FaceId = 4174 'assigns an icon to the dropdown
        End With    
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Article" 'adds a description to the menu item
            .OnAction = "XLAD" 'runs the specified macro
            .FaceId = 6824 'assigns an icon to the dropdown
        End With          
    End With         
        
    With cmbControl 
        .Caption = "&Help" 'names the menu item        
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Resources" 'adds a description to the menu item
            .OnAction = "Resources" 'runs the specified macro
            .FaceId = 9987 'assigns an icon to the dropdown
        End With          
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Help" 'adds a description to the menu item
            .OnAction = "Help" 'runs the specified macro
            .FaceId = 984 'assigns an icon to the dropdown
        End With    
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Documentation" 'adds a description to the menu item
            .OnAction = "Documentation" 'runs the specified macro
            .FaceId = 9325 'assigns an icon to the dropdown
        End With           
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Demo" 'adds a description to the menu item
            .OnAction = "Demo" 'runs the specified macro
            .FaceId = 9456 'assigns an icon to the dropdown
        End With        
    End With              
End Sub 
 
 
Private Sub Workbook_BeforeClose(Cancel As Boolean) 
    On Error Resume Next 'in case the menu item has already been deleted
    Application.CommandBars("Worksheet Menu Bar").Controls("My Macros").Delete 'delete the menu item
End Sub 
 
