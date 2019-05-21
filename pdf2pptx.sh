#!/bin/bash
# Alireza Shafaei - shafaei@cs.ubc.ca - Jan 2016

resolution=1024
density=300
#colorspace="-depth 8"
colorspace="-colorspace sRGB -background white -alpha remove"
makeWide=true

if [ $# -eq 0 ]; then
	echo "No arguments supplied!"
	echo "Usage: ./pdf2pptx.sh file.pdf"
	echo "       .Generates file.pdf.pptx in widescreen format (by default)"
	echo "       ./pdf2pptx.sh file.pdf notwide"
	echo "       .Generates file.pdf.pptx in 4:3 format"
	exit 1
fi

if [ $# -eq 2 ]; then
	if [ "$2" == "notwide" ]; then
		makeWide=false
	fi
fi

if [ $# -eq 2 ]; then
	if [ "$2" == "notwide" ]; then
		makeWide=false
	fi
fi

tempname=$(mktemp -d 2>/dev/null || mktemp -d -t 'tmp')
fout="$(pwd)/$(basename $1).pptx"

# deletes the temp directory on exit
# based on https://stackoverflow.com/questions/4632028/how-to-create-a-temporary-directory#34676160
function cleanup {
	rm -rf "$tempname"
}

# check if tmp dir was created
if [[ ! "$tempname" || ! -d "$tempname" ]]; then
	echo "Could not create tempdir"
	exit 1
else
	trap cleanup EXIT
fi

# $colorspace may contain multiple parameters passed to convert
# shellcheck disable=SC2086
if convert -density "$density" $colorspace -resize "x${resolution}" "$1" "$tempname/ppt/media/slide.png"; then
	echo "Extraction succ!"
else
	echo "Error with extraction"
	exit 1
fi

fout=$1.pptx
cp -r template "$tempname"

mkdir "$tempname/ppt/media"
cp "$tempname/"*".png" "$tempname/ppt/media/"
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
	cp ../slides/slide1.xml "../slides/slide-$1.xml"
	sed "s/image1\.JPG/slide-${slide}.png/g" ../slides/_rels/slide1.xml.rels > "../slides/_rels/slide-$1.xml.rels"
	add_slide "$1"
}

pushd "$tempname/ppt/media/" || exit
	count=$(find . -maxdepth 1 -name "*.png" -printf '%i\n' | wc -l)
	for (( slide=count-1; slide>=0; slide-- )); do
		echo "Processing $slide"
		make_slide "$slide"
	done

	if [ "$makeWide" = true ]; then
		pat='<p:sldSz cx=\"9144000\" cy=\"6858000\" type=\"screen4x3\"\/>'
		wscreen='<p:sldSz cy=\"6858000\" cx=\"12192000\"\/>'
		sed -i "s/${pat}/${wscreen}/g" ../presentation.xml
	fi
popd || exit 1

pushd "$tempname" || exit
	rm -rf "../$fout"
	zip -q -r "../$fout" .
popd || exit 1

rm -rf "$tempname"
