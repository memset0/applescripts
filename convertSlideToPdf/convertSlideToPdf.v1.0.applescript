# 批量转换幻灯片到 pdf
# v1.0 by memset0 @ 2020.6

on listFile(rootFolder)
	tell application "Finder"
		return every file of rootFolder
	end tell
end listFile

on filterFileByExtension(sourceFileList, neededExtension)
	set distFileList to {}
	repeat with nowFile in sourceFileList
		set fileExtension to the name extension of nowFile
		if fileExtension = neededExtension then
			set end of distFileList to nowFile
		end if
	end repeat
	return distFileList
end filterFileByExtension

on listFileName(fileList)
	set fileNameSet to {}
	repeat with nowFile in fileList
		set fileName to the name of nowFile
		set end of fileNameSet to fileName
	end repeat
	return fileNameSet
end listFileName

on convertSlideNametoPdfName(slideName)
	if slideName ends with ".ppt" then
		return (text 1 thru -4 of slideName) & "pdf"
	end if
	if slideName ends with ".pptx" then
		return (text 1 thru -5 of slideName) & "pdf"
	end if
	return slideName & ".pdf"
end convertSlideNametoPdfName

on convertSlideToPdf(slideList)
	set pdfList to {}
	tell application "Microsoft PowerPoint"
		launch
		repeat with thisSlide in slideList
			set slideName to the name of thisSlide
			set slidePath to thisSlide as string
			set pdfPath to my convertSlideNametoPdfName(slidePath) # quirk?
			open thisSlide
			display dialog "是否转换幻灯片\"" & slideName & "\"?"
			save active presentation in pdfPath as save as PDF
			set end of pdfList to pdfPath as alias
		end repeat
	end tell
	tell application "Microsoft PowerPoint"
		display dialog "是否关闭 Microsoft PowerPoint"
		quit
	end tell
	return pdfList
end convertSlideToPdf

set allFileList to listFile(choose folder "Select a folder")
set fileList to filterFileByExtension(allFileList, "ppt") 
set fileList to fileList & filterFileByExtension(allFileList, "pptx")

set console to convertSlideToPdf(fileList)