#!/bin/bash
# Convert all doc files to docx using libreoffice 
outdir=DOC_TO_DOCX_CONVERSIONS
mkdir $outdir
ls *.doc |while read line; do 
echo libreoffice --headless  --convert-to docx --outdir $outdir $line>> libreoffice-conversion-record.log; 
fname=$(echo $line | sed -re 's/\.doc/\.docx/'); 
echo $fname;
if [ ! -e $outdir/$fname ]; then 
     	# running Libreoffice with a timeout because occasionally files fail to transfer
	timeout 30s libreoffice --headless  --convert-to docx --outdir $outdir $line; 
	if [ $? = 143 ]; then
		killall -9 soffice.bin 
		echo "Couldn't convert $line ">> libreoffice-failed-record.log
	fi
fi
done
