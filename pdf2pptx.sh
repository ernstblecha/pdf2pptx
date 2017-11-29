#!/bin/bash
# Alireza Shafaei - shafaei@cs.ubc.ca - Jan 2016
# Forked by Jasper Chan - jasper.chan@alumni.ubc.ca - Nov 2017

resolution=1024
density=300
#colorspace="-depth 8"
colorspace="-colorspace sRGB -background white -alpha remove"
makewide=false
nobar=true
istex=false
filename=

if [ $# -eq 0 ]; then
    echo "No arguments supplied!"
    echo "Usage: ./pdf2pptx.sh [options] file"
    echo "Generates file.pdf.pptx from a beamer tex/pdf file"
    echo ""
    echo "Options:"
    echo "-w, --wide        Render pptx for widescreen format"
    echo "-b, --bar         Include beamer's navigation bar (only works on .tex files)"
    exit 1
fi

while [ "$1" != "" ]; do
    case $1 in
        -w | --wide )   makewide=true
                        ;;
        -b | --bar )    nobar=false
                        ;;
        * )             filename=$1
    esac
    shift
done

if [ ${filename: -4} == ".tex" ]; then
    istex=true
    cp $filename "$filename.orig"
    if [ $nobar = true ]; then
        echo "Removing the navigation bar from your tex file..."
        sed -i 's/\documentclass{beamer}/\documentclass{beamer}\n\\beamertemplatenavigationsymbolsempty/' $filename
    fi
    echo "Rendering .tex file..."
    latexmk -pdf $1
    latexmk -c $1
    mv "$filename.orig" $filename
    filename="${filename::-4}.pdf"
    echo "Checking for $filename"
    if [ ! -f $filename ]; then
        echo "File not found, check if it was rendered successfully?"
        exit
    fi
    echo "File found!"
elif [ $nobar = true ]; then
    echo "Warning: Can't remove the navigation bar from a pdf."
fi

echo "Converting $filename..."
tempname="$filename.temp"
if [ -d $tempname ]; then
	echo "Removing ${tempname}"
	rm -rf $tempname
fi

mkdir $tempname
convert -density $density $colorspace -resize "x${resolution}" $filename ./$tempname/slide.png

if [ $? -eq 0 ]; then
	echo "Extraction success!"
else
	echo "Error with extraction"
	exit
fi

pptname="$filename.pptx.base"
fout="$filename.pptx"
rm -rf $pptname
cp -r template $pptname

mkdir $pptname/ppt/media

cp ./$tempname/*.png "$pptname/ppt/media/"

function add_slide {
	pat='slide1\.xml\"\/>'
	id=$1
	id=$((id+8))
	entry='<Relationship Id=\"rId'$id'\" Type=\"http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/slide\" Target=\"slides\/slide-'$1'\.xml"\/>'
	rep="${pat}${entry}"
	sed -i "s/${pat}/${rep}/g" ../_rels/presentation.xml.rels 

	pat='slide1\.xml\" ContentType=\"application\/vnd\.openxmlformats-officedocument\.presentationml\.slide+xml\"\/>'
	entry='<Override PartName=\"\/ppt\/slides\/slide-'$1'\.xml\" ContentType=\"application\/vnd\.openxmlformats-officedocument\.presentationml\.slide+xml\"\/>'
	rep="${pat}${entry}"
	sed -i "s/${pat}/${rep}/g" ../../\[Content_Types\].xml

	sid=$1
	sid=$((sid+256))
	pat='<p:sldIdLst>'
	entry='<p:sldId id=\"'$sid'\" r:id=\"rId'$id'\"\/>'
	rep="${pat}${entry}"
	sed -i "s/${pat}/${rep}/g" ../presentation.xml
}

function make_slide {
	cp ../slides/slide1.xml ../slides/slide-$1.xml
	cat ../slides/_rels/slide1.xml.rels | sed "s/image1\.JPG/slide-${slide}.png/g" > ../slides/_rels/slide-$1.xml.rels
	add_slide $1
}

pushd $pptname/ppt/media/
count=`ls -ltr | wc -l`
for (( slide=$count-2; slide>=0; slide-- ))
do
	echo "Processing "$slide
	make_slide $slide
done

if [ "$makewide" = true ]; then
	pat='<p:sldSz cx=\"9144000\" cy=\"6858000\" type=\"screen4x3\"\/>'
	wscreen='<p:sldSz cy=\"6858000\" cx=\"12192000\"\/>'
	sed -i "s/${pat}/${wscreen}/g" ../presentation.xml
fi
popd

pushd $pptname
rm -rf ../$fout
zip -q -r ../$fout .
popd

rm -rf $pptname
rm -rf $tempname
