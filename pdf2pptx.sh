#!/bin/bash
# Alireza Shafaei - shafaei@cs.ubc.ca - Jan 2016

resolution=1024
density=300
#colorspace="-depth 8"
colorspace="-colorspace sRGB -background white -alpha remove"
makeWide=true
usepdftoppm=true

if ! hash pdftoppm 2>/dev/null; then
	echo "pdftoppm not available!"
	usepdftoppm=false
	if ! hash convert 2>/dev/null; then
		echo "convert not available"
		exit 1
	fi
fi

if ! hash seq 2>/dev/null; then
	echo "seq not available!"; exit 1
fi

if ! hash zip 2>/dev/null; then
	echo "zip not available!"; exit 1
fi

if ! hash unzip 2>/dev/null; then
	echo "unzip not available!"; exit 1
fi

if ! hash basename 2>/dev/null; then
	echo "basename not available!"; exit 1
fi

if ! hash basename 2>/dev/null; then
	echo "base64 not available!"; exit 1
fi

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

tempname=$(mktemp -d 2>/dev/null || mktemp -d -t 'tmp')
fout="$(pwd)/$(basename "$1").pptx"

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

function templatedata {
	# files from "template.zip"  base64 encoded
	cat <<TEMPLATE
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
TEMPLATE
}

function slideXmlTemplate {
	cat <<TEMPLATE
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId2"><a:lum/></a:blip><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree><p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="4032986112"/></p:ext></p:extLst></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>
TEMPLATE
}

function slideXmlRelTemplate {
	local num=$1
	cat <<TEMPLATE
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/slide-${num}.png"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>
TEMPLATE
}

function presentationXmlRelsTemplate {
	local list=$1
	cat <<TEMPLATE
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>${list}<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/></Relationships>
TEMPLATE
}

function contentTypesXmlTemplate {
	local list=$1
	cat <<TEMPLATE
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="JPG" ContentType="image/jpeg"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>${list}<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>
TEMPLATE
}

function presentationXmlTemplate {
	local list=$1
	local screen=$2
	cat <<TEMPLATE
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483656" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${list}</p:sldIdLst>${screen}<p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr></p:defaultTextStyle><p:extLst><p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"/></p:ext></p:extLst></p:presentation>
TEMPLATE
}

function coreXmlTemplate {
	local author="pdf2pptx"
	local datetime
	datetime=$(date -Iseconds)
	cat <<TEMPLATE
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>PowerPoint Presentation</dc:title><dc:creator>${author}</dc:creator><cp:lastModifiedBy>${author}</cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">${datetime}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${datetime}</dcterms:modified></cp:coreProperties>
TEMPLATE
}

function make_slide {
	local num=$1
	local id=$((${1#0}+8))
	local sid=$((${1#0}+256))

	slideXmlTemplate "$num" > "$tempname/ppt/slides/slide-$num.xml"
	slideXmlRelTemplate "$num" > "$tempname/ppt/slides/_rels/slide-$num.xml.rels"

	relationList+='<Relationship Id="rId'$id'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide-'$1'.xml"/>'
	contentList+='<Override PartName="/ppt/slides/slide-'$1'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
	slideList+='<p:sldId id="'$sid'" r:id="rId'$id'"/>'
}

# temporary file is needed as explained here https://unix.stackexchange.com/questions/2690/how-to-redirect-output-of-wget-as-input-to-unzip
# in short: unzip <(...) is only supported in busybox implementation...
# busybox unzip -q <(templatedata | base64 -d) -d "$tempname"
# other options: use tar,... instead of unzip
templatedata | base64 -d > "$tempname/template.zip"
unzip -q "$tempname/template.zip" -d "$tempname"
rm "$tempname/template.zip"

if $usepdftoppm; then
	if pdftoppm -r "$density" -scale-to "$resolution" -png "$1" "$tempname/ppt/media/slide"; then
		echo "Extraction successful!"
	else
		echo "Error during extraction"; exit 1
	fi
	pdftoppm -r "$density" -scale-to-x 256 -scale-to-y -1 -singlefile -jpeg "$1" "$tempname/docProps/thumbnail"
	mv "$tempname/docProps/thumbnail.jpg" "$tempname/docProps/thumbnail.jpeg"
else
	# $colorspace may contain multiple parameters passed to convert
	# shellcheck disable=SC2086
	if convert -density "$density" $colorspace -resize "x${resolution}" "$1" "$tempname/ppt/media/slide.png"; then
		echo "Extraction successful!"
	else
		echo "Error during extraction"; exit 1
	fi
	# $colorspace may contain multiple parameters passed to convert
	# shellcheck disable=SC2086
	convert -density "$density" $colorspace -resize 256x "$1" "$tempname/docProps/thumbnail.jpeg"
fi

relationList=""
contentList=""
slideList=""
count=$(find "$tempname/ppt/media/" -maxdepth 1 -name "*.png" -printf '%i\n' | wc -l)
if $usepdftoppm; then
	for slide in $(seq -w 1 "$count"); do
		echo "Processing slide $slide"
		make_slide "$slide"
	done
else
	for slide in $(seq 1 "$count"); do
		echo "Processing slide $slide"
		make_slide "$slide"
	done
fi

if [ "$makeWide" = true ]; then
	screen='<p:sldSz cy="6858000" cx="12192000"/>'
else
	screen='<p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>'
fi

presentationXmlRelsTemplate "$relationList" > "$tempname/ppt/_rels/presentation.xml.rels"
contentTypesXmlTemplate "$contentList" > "$tempname/[Content_Types].xml"
presentationXmlTemplate "$slideList" "$screen" > "$tempname/ppt/presentation.xml"
coreXmlTemplate > "$tempname/docProps/core.xml"

rm -f "$fout"

cd "$tempname" || exit 1
zip -q -r "$fout" .
