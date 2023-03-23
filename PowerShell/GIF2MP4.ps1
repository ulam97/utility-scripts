# powershell D:\Clutter\testedit\GIF2MP4.ps1
#$HSrc = "E:\mega\src"       #Input folder
#$HSrc = "D:\Clutter\testedit\base"
$HSrc = "D:\Clutter\src"
#$HTemp = "E:\mega\temp"           #Temporary folder
$HTemp = "D:\Clutter\testedit\mem1"
$HOut = "D:\Clutter\testedit\out1"        #Output folder

#Set speed by frame per seconds (FPS)
$HSpeed = 1

#padsize
#$iw = 854
#$ih = 641
#$iw = 600
#$ih = 339
#$iw = 560
#$ih = 315
#$iw = 1136
#$ih = 639
#$iw = 1046
#$ih = 589
#SC
#$iw,$ih = 1280,741


#function to convert
function GIF_2_MP4{
  param ($HAni, $HSrc, $HTem, $HSpeed)

  del $HTem\*
	$filename = $HAni -replace ".{4}$"        #chop .gif extension
    
  #Chop frame files
  magick convert -coalesce $HSrc\$HAni $HTem\$filename-%05d.jpg

  $FrameCount = Get-ChildItem $HTem -File | Measure-Object | %{$_.Count}
	#$FramePS = [Math]::Ceiling($FrameCount*$HSpeed) #SP
	$FramePS = 10 #mgcm 30, AL 20, neto 24, shg 10, urajuta 10, skbelf 15

	#Animate the frame set
	#echo 'Y' | ffmpeg -framerate $FramePS -i $HTem\$filename-%05d.jpg -c:v libx264 -vf "fps=$FramePS,format=yuv420p" $HOut\$filename.mp4

	#with pad
	echo 'Y' | ffmpeg -framerate $FramePS -i $HTem\$filename-%05d.jpg -c:v libx264 -vf "pad=ceil(iw/2)*2:ceil(ih/2)*2,fps=$FramePS,format=yuv420p" $HOut\$filename.mp4

}

#Get frame files
$filelists = get-childitem $HSrc

#Process each frame sets
foreach ($hentai in $filelists) {
	GIF_2_MP4 $hentai $HSrc $HTemp $HSpeed	
}