Set FileSysObj = CreateObject("Scripting.FileSystemObject")
DesktopPath = FileSysObj.GetAbsolutePathName(".")
Consent = MsgBox("Do you want to run this script?", vbYesNo, "Consent")

If Consent = vbYes Then
    '-- Define a dictionary to map extensions to custom folder names
    Set ExtensionMap = CreateObject("Scripting.Dictionary")
    ExtensionMap.Add "rbxm", "Roblox Studio Files"
    ExtensionMap.Add "rbxl", "Roblox Studio Places"
    ExtensionMap.Add "png", "Images"
    ExtensionMap.Add "jpg", "Images"
    ExtensionMap.Add "jpeg", "Images"
    ExtensionMap.Add "gif", "Gifs"
    ExtensionMap.Add "txt", "Text Files"
    ExtensionMap.Add "exe", "Executable Files and Setups"
    ExtensionMap.Add "msi", "Setups"
    ExtensionMap.Add "ogg", "Sound Files"
    ExtensionMap.Add "mp3", "Sound Files"
    ExtensionMap.Add "wav", "Sound Files"
    ExtensionMap.Add "mp4", "Video Files"
    ExtensionMap.Add "mov", "Video Files"
    ExtensionMap.Add "rar", "Rars and Zip"
    ExtensionMap.Add "zip", "Rars and Zip"
    ExtensionMap.Add "psd", "Photoshop Files"
    ExtensionMap.Add "lua", "Roblox Script Files"
    ExtensionMap.Add "flac", "High Quality Sounds"
    ExtensionMap.Add "lnk", "Shortcuts"
    ExtensionMap.Add "blend", "Blender Projects"
    ExtensionMap.Add "blend1", "Blender Projects"
    ExtensionMap.Add "py", "Python Scripts"
    ExtensionMap.Add "webm", "Video Files"
    ExtensionMap.Add "rbxmx", "Roblox Studio Files"
    ExtensionMap.Add "doc", "Microsoft Word Documents"
    ExtensionMap.Add "docx", "Microsoft Word Documents"
    ExtensionMap.Add "xls", "Microsoft Excel Spreadsheets"
    ExtensionMap.Add "xlsx", "Microsoft Excel Spreadsheets"
    ExtensionMap.Add "ppt", "Microsoft PowerPoint Presentations"
    ExtensionMap.Add "pptx", "Microsoft PowerPoint Presentations"
    ExtensionMap.Add "pdf", "PDF Files"
    ExtensionMap.Add "csv", "Comma Separated Values Files"
    ExtensionMap.Add "svg", "Scalable Vector Graphics Files"
    ExtensionMap.Add "html", "HTML Files"
    ExtensionMap.Add "htm", "HTML Files"
    ExtensionMap.Add "js", "JavaScript Files"
    ExtensionMap.Add "css", "Cascading Style Sheets Files"
    ExtensionMap.Add "json", "JSON Files"
    ExtensionMap.Add "xml", "XML Files"
    ExtensionMap.Add "md", "Markdown Files"
    ExtensionMap.Add "avi", "Video Files"
    ExtensionMap.Add "mkv", "Video Files"
    ExtensionMap.Add "wmv", "Video Files"
    ExtensionMap.Add "bmp", "Image Files"
    ExtensionMap.Add "tiff", "Image Files"
    ExtensionMap.Add "ico", "Image Files"
    ExtensionMap.Add "odt", "OpenDocument Text Files"
    ExtensionMap.Add "ods", "OpenDocument Spreadsheet Files"
    ExtensionMap.Add "odp", "OpenDocument Presentation Files"
    ExtensionMap.Add "epub", "E-book Files"
    ExtensionMap.Add "mobi", "E-book Files for Amazon Kindle"
    ExtensionMap.Add "azw", "E-book Files for Amazon Kindle"
    ExtensionMap.Add "azw3", "E-book Files for Amazon Kindle"
    ExtensionMap.Add "djvu", "Compressed Image Files for Scanned Documents"
    ExtensionMap.Add "cbz", "Comic Book Archive Files"
    ExtensionMap.Add "cbr", "Comic Book Archive Files"
    ExtensionMap.Add "flv", "Adobe Flash Video Files"
    ExtensionMap.Add "swf", "Adobe Flash Animation Files"
    ExtensionMap.Add "ps", "PostScript Files"
    ExtensionMap.Add "eps", "Encapsulated PostScript Files"
    ExtensionMap.Add "ai", "Adobe Illustrator Files"
    ExtensionMap.Add "indd", "Adobe InDesign Files"
    ExtensionMap.Add "dwg", "AutoCAD Drawing Files"
    ExtensionMap.Add "dxf", "AutoCAD Interchange Files"
    ExtensionMap.Add "skp", "SketchUp 3D Modeling Files"
    ExtensionMap.Add "log", "Log Files"
    ExtensionMap.Add "bat", "Windows Batch Files"
    ExtensionMap.Add "cmd", "Windows Batch Files"
    ExtensionMap.Add "sh", "Unix Shell Scripts"
    ExtensionMap.Add "iso", "Optical Disc Image Files"
    ExtensionMap.Add "vmdk", "Virtual Machine Disk Image Files"
    ExtensionMap.Add "vdi", "Virtual Machine Disk Image Files"
    ExtensionMap.Add "vhd", "Virtual Machine Disk Image Files"
    ExtensionMap.Add "ova", "Open Virtualization Format Files"
    ExtensionMap.Add "ovf", "Open Virtualization Format Files"
    ExtensionMap.Add "qcow2", "QEMU Copy-On-Write Disk Image Files"
    ExtensionMap.Add "raw", "Raw Image Files"
    ExtensionMap.Add "cr2", "Digital Camera Raw Image Files"
    ExtensionMap.Add "nef", "Digital Camera Raw Image Files"
    ExtensionMap.Add "arw", "Digital Camera Raw Image Files"
    ExtensionMap.Add "dng", "Digital Camera Raw Image Files"
    ExtensionMap.Add "3gp", "3GPP Multimedia Files"
    ExtensionMap.Add "3g2", "3GPP2 Multimedia Files"
    ExtensionMap.Add "asf", "Advanced Systems Format Media Files"
    ExtensionMap.Add "wma", "Advanced Systems Format Media Files"
    ExtensionMap.Add "wmv", "Advanced Systems Format Media Files"
    ExtensionMap.Add "svelte", "Svelte Files"
    ExtensionMap.Add "jsx", "JSX Files"
    ExtensionMap.Add "tsx", "TSX Files"
    ExtensionMap.Add "go", "GoLang Files"
    ExtensionMap.Add "rs", "Rust Files"

    ' Add more custom mappings here

    FileCount = 0
    
    ' Add an array variable to store the list of files being moved
    Dim FilesMoved()
    ReDim FilesMoved(0)
    
    ' Add an array variable to store the list of folders being created
    Dim FoldersCreated()
    ReDim FoldersCreated(0)

For Each File In FileSysObj.GetFolder(DesktopPath).Files
    Extension = FileSysObj.GetExtensionName(File.Path)
    If Extension <> "" And LCase(Extension) <> "vbs" Then
        If ExtensionMap.Exists(Extension) Then
            TargetFolder = FileSysObj.BuildPath(DesktopPath, ExtensionMap(Extension))
        Else
            TargetFolder = FileSysObj.BuildPath(DesktopPath, Extension)
        End If

        If Not FileSysObj.FolderExists(TargetFolder) Then
            ReDim Preserve FoldersCreated(UBound(FoldersCreated) + 1)
            FoldersCreated(UBound(FoldersCreated)) = TargetFolder
            FileSysObj.CreateFolder(TargetFolder)
        End If

        TargetFile = FileSysObj.BuildPath(TargetFolder, File.Name)
        If FileSysObj.FileExists(TargetFile) Then
            RemoveDuplicate = MsgBox(File.Name &" already exists in "& TargetFolder &". Do you want to remove the duplicate?", vbYesNo, "Remove Duplicate")
            If RemoveDuplicate = vbYes Then
                File.Delete
            End If
        Else
            If FileCount >= 100 Then
                WScript.Echo "Files moved so far: (" & UBound(FilesMoved) & "): " & Join(FilesMoved, ", ")
                ContinueMoving = MsgBox("Do you want to continue moving files?", vbYesNo, "Continue Moving")
                If ContinueMoving = vbNo Then
                    Exit For
                Else
                    ReDim FilesMoved(0)
                    FileCount = 0
                End If
            End If

            ReDim Preserve FilesMoved(UBound(FilesMoved) + 1)
            FilesMoved(UBound(FilesMoved)) = File.Name
            FileSysObj.MoveFile File.Path, TargetFile
            FileCount = FileCount + 1
        End If
    End If
Next
    
    ' Print out the final list of files that have been moved and folders that have been created.
    If UBound(FilesMoved) > 0 Then
        WScript.Echo "Files Moved (" & UBound(FilesMoved) & "): " & Join(FilesMoved, ", ")
    Else
        WScript.Echo "Files Moved: None"
    End If
    
    If UBound(FoldersCreated) > 0 Then
        WScript.Echo "Folders created (" & UBound(FoldersCreated) & "): " & Join(FoldersCreated, ", ")
    Else
        WScript.Echo "Folders created: None"
    End If
    
End If