Option Explicit
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim from, dest, jpegExt, jpgExt, currFile
Dim formattedDate, yearPath, datePath
Dim fileAmount, folderAmount, folderCount
Dim folder, subFolder
Dim list

from = WScript.Arguments(0) & "\"
dest = WScript.Arguments(1) & "\"
jpegExt = "jpeg"
jpgExt = "jpg"
fileAmount = 0
folderAmount = 0

function pd(n, totalDigits) 
  if totalDigits > len(n) then 
    pd = String(totalDigits - len(n), "0") & n 
  else 
    pd = n 
  end if 
end function 

function search(folder)
  for each subFolder in fso.GetFolder(folder).SubFolders
    for each currFile in fso.GetFolder(subFolder.Path).Files
      formattedDate = Year(currFile.DateLastModified) & "-" & Pd(Month(currFile.DateLastModified),2) & "-" & Pd(Day(currFile.DateLastModified),2)
      yearPath = dest & Year(currFile.DateLastModified) & "\"
      datePath = dest & Year(currFile.DateLastModified) & "\" & formattedDate & "\"
      if fso.GetExtensionName(currFile.Path) = jpegExt or fso.GetExtensionName(currFile.Path) = jpgExt then
        if fso.FolderExists(datePath) then
          currFile.Copy datePath
          fileAmount = fileAmount + 1
        elseif fso.FolderExists(yearPath) then
          fso.CreateFolder datePath
          currFile.Copy datePath
          fileAmount = fileAmount + 1
          folderAmount = folderAmount + 1
        else
          fso.CreateFolder yearPath
          fso.CreateFolder datePath
          currFile.Copy datePath
          fileAmount = fileAmount + 1
          folderAmount = folderAmount + 1
          folderAmount = folderAmount + 1
        end if
      end if
    next
    call search(subFolder)
  next
end function

function output(folder)
  for each subFolder in fso.GetFolder(folder).SubFolders
  folderCount = fso.GetFolder(subFolder).Files.count
    if folderCount > 1 then
      wscript.echo "--------"
      wscript.echo folderCount & " files"
      set list = CreateObject("System.Collections.ArrayList")
      for each currFile in fso.GetFolder(subFolder.Path).Files
        list.Add(currFile.Name)
      next
      wscript.echo join(list.ToArray(), ", ")
      wscript.echo "were moved to folder"
      wscript.echo SubFolder.Path
    elseif folderCount = 1 then
      wscript.echo "--------"
      wscript.echo folderCount & " file"
      for each currFile in fso.GetFolder(subFolder.Path).Files
        wscript.echo currFile.name
      next
      wscript.echo "was moved to folder"
      wscript.echo subFolder.Path
    end if
    call output(subFolder)
  next
end function

if not fso.FolderExists(from) then
  Wscript.Quit
elseif not fso.FolderExists(dest) then
  Wscript.Quit
end if

for each currFile in fso.GetFolder(from).Files
  formattedDate = Year(currFile.DateLastModified) & "-" & Pd(Month(currFile.DateLastModified),2) & "-" & Pd(Day(currFile.DateLastModified),2)
  yearPath = dest & Year(currFile.DateLastModified) & "\"
  datePath = dest & Year(currFile.DateLastModified) & "\" & formattedDate & "\"
  if fso.GetExtensionName(currFile.Path) = jpegExt or fso.GetExtensionName(currFile.Path) = jpgExt then
    if fso.FolderExists(datePath) then
      currFile.Copy datePath
      fileAmount = fileAmount + 1
    elseif fso.FolderExists(yearPath) then
      fso.CreateFolder datePath
      currFile.Copy datePath
      fileAmount = fileAmount + 1
      folderAmount = folderAmount + 1
    else
      fso.CreateFolder yearPath
      fso.CreateFolder datePath
      currFile.Copy datePath
      fileAmount = fileAmount + 1
      folderAmount = folderAmount + 1
      folderAmount = folderAmount + 1
    end if
  end if
next
call search(from)

if fileAmount = 1 and folderAmount = 1 then
  wscript.echo fileAmount & " picture was sorted into " &  folderAmount & " folder."
elseif fileAmount = 1 then
  wscript.echo fileAmount & " picture was sorted into " &  folderAmount & " folders."
elseif folderAmount = 1 then
  wscript.echo fileAmount & " pictures were sorted into " &  folderAmount & " folder."
else 
  wscript.echo fileAmount & " pictures were sorted into " &  folderAmount & " folders."
end if

call output(dest)