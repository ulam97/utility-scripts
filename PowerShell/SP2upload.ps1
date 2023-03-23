#Senran Princess
#animate image rolls to a set of mp4 animation file.

$dirsource = "C:\Source\Folder"
$dirtemp = "C:\Temporary\Folder"
$dirtarget = "C:\Output\Folder"
$dim = "854x640"                               #Animation dimension
$HSpeed = 1                                    #Set speed

#Get frame files
$filelists = get-childitem $dirsource

#Process each frame sets
foreach ($hentai in $filelists) {
	del $dirtemp\*
	$filename = $hentai -replace ".{4}$"        #remove file extension .jpg or .png
    
    #Chop frame files
	magick convert $dirsource\$filename.jpg -crop $dim $dirtemp\$filename.jpg
	
    #Calculating speed based on number of frames in the set
	$FrameCount = Get-ChildItem $dirtemp -File | Measure-Object | %{$_.Count}
	$FramePS = [Math]::Ceiling($FrameCount*$HSpeed)

	#Animate the frame set
	ffmpeg -framerate $FramePS -i $dirtemp\$filename-%d.jpg -c:v libx264 -vf "fps=$FramePS,format=yuv420p" $dirtarget\$filename.mp4
}