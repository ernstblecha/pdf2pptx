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

function templatedata {
	# files from "template/" zipped and base64 encoded
	cat <<TEMPLATE
UEsDBBQAAAAIADm1s06QrFwRhwEAAOAGAAATABwAW0NvbnRlbnRfVHlwZXNdLnhtbFVUCQADjb/h
XI2/4Vx1eAsAAQToAwAABOgDAAC1VclOwzAQvfMVka8occsBIdS0BzYJsVSifIBJJqnBsS3bLc3f
M0laFKosRW0ukcYzbxmPNZnMNpnw1mAsVzIk42BEPJCRirlMQ/K+uPeviGcdkzETSkJIcrBkNj2b
LHIN1kOwtCFZOqevKbXREjJmA6VBYiZRJmMOQ5NSzaIvlgK9GI0uaaSkA+l8V3CQ6eQWErYSzrvb
4HFlRMuUeDdVXSEVEp4V+OKcNiI+NTRDykQzxoCweximteARc5inaxnv9eJv+wgQWdbYJdf2HAta
FIpMu0A77nH+0NnMK47M8Bi8OTPuhWVYQLV2VBuwCCnpg27xhu5UkvAIYhWtMoQEdbJM/AmDjHG5
67vNjBV4+Mysw+dVD8andlbjPsjT1s0wPvocFJi5UdoOMZ+SuM/BmsP3IA5+ifscOFwUUH2PH0JJ
06vIPgS8uVzAybuuUR/0+p5YrlbO1oNhXmLF3eUJ4eW8cCUb+L+H3TYs0L5GIjCOd9/CryJSH900
FAszhrhBm5Y/qOkPUEsDBAoAAAAAAOtotU4AAAAAAAAAAAAAAAAEABwAcHB0L1VUCQAD6tvjXI2/
4Vx1eAsAAQToAwAABOgDAABQSwMECgAAAAAAObWzTgAAAAAAAAAAAAAAABEAHABwcHQvc2xpZGVM
YXlvdXRzL1VUCQADjb/hXI2/4Vx1eAsAAQToAwAABOgDAABQSwMEFAAAAAgAObWzTnCINDPiAgAA
gQcAACEAHABwcHQvc2xpZGVMYXlvdXRzL3NsaWRlTGF5b3V0MS54bWxVVAkAA42/4VyNv+FcdXgL
AAEE6AMAAAToAwAAtVVdcpswEH7vKRj6TMRfMPbEzhgbdzqTJp46OYACwmYCkirJrt1OZnKt9jg5
SVcCYjdJO3lwXpBYdlf7fd+yOjvf1pW1IUKWjA5t78S1LUIzlpd0ObRvrmdObFtSYZrjilEytHdE
2uejD2d8IKv8Au/YWlmQgsoBHtorpfgAIZmtSI3lCeOEwreCiRoreBVLlAv8HVLXFfJdN0I1Lqnd
xou3xLOiKDMyZdm6JlQ1SQSpsILy5arkssvG35KNCyIhjYn+uyS14wD2tsL0zraMm9iAwbNHgDxb
VLlFcQ2GxHhoo+TXghC9o5tPgi/4XBjfy81cWGWuY9sYG7UfWjfUBJkNeha+7LZ4sC1ErVegwNoO
bRBqp59I28hWWVljzPbWbHX1im+2Sl/xRt0B6OBQjaop7iUcv4MzxYpY8wpnZMWqnAjLewLYlS75
BcvupEUZQNNMNEifPBr4euWrlvpcQd/9ABFxVdhwIJTruXbHkHZGh3XJjke1TVi+04fewmqMeFBJ
tVC7ipgXrh8FKKhR/Jz0onCa9vqOH8UzJ5x4YyfuBwG8Bp7fn/rjWd+97/ohB6iqrMmsXK4FuVor
W+cSwAi0AfwvhDo3C9vKS6H2fKuRh/weNJcXaZaV4RrON7rRfI4F/vq/DMiUjPbQUCfLv8UJOnFm
jCmQ5FAe/xjyFEo0+nxbYwEndBJ5x5PovbgJO24WVZkT63Jd3z5jKDgGQzAeIfWrJPnv0MdhPIuC
cZA4vXjSc8Jekjrj9DR1Us/1Qjd1wygMnvpYauQUqntT+z4+/Pr4+PD7qM2LDgcmTK8LqdqdtRYl
4EmSfuRP4sRJvBD+y2m/54xn0akzOw3CcJLE40mQ3uvB64WDTBAzwj/n3fD3whfjvy4zwSQr1EnG
6vYeQZx9J4Kz0lwlntsO/w2u4Bc6jeN+6EZR0KoFtXWrqRY1F4HplEp8wfxqY3oFDgOtJ8bE4a5r
W2Xvgg7uztEfUEsDBAoAAAAAADm1s04AAAAAAAAAAAAAAAAXABwAcHB0L3NsaWRlTGF5b3V0cy9f
cmVscy9VVAkAA42/4VyNv+FcdXgLAAEE6AMAAAToAwAAUEsDBBQAAAAIADm1s05A+3BatAAAADYB
AAAsABwAcHB0L3NsaWRlTGF5b3V0cy9fcmVscy9zbGlkZUxheW91dDEueG1sLnJlbHNVVAkAA42/
4VyNv+FcdXgLAAEE6AMAAAToAwAAjc+9CsIwEAfw3acIt5u0DiLStIsIDi6iD3Ak1zbYJiEXRd/e
jBYcHO/r9+ea7jVP4kmJXfAaalmBIG+CdX7QcLse1zsQnNFbnIInDW9i6NpVc6EJc7nh0UUWBfGs
Ycw57pViM9KMLEMkXyZ9SDPmUqZBRTR3HEhtqmqr0rcB7cIUJ6shnWwN4vqO9I8d+t4ZOgTzmMnn
HxGKJ2fpjJwpFRbTQFmDlN/9xVItSwSotlGLd9sPUEsDBBQAAAAIADm1s06Y0TLyggIAACQNAAAU
ABwAcHB0L3ByZXNlbnRhdGlvbi54bWxVVAkAA42/4VyNv+FcdXgLAAEE6AMAAAToAwAA7Zfdbpsw
FIDv9xTItxMF8xeKQqpmHdOkToqa9gFccBpUYyPbSZNOe/ceExNoq0l9AK5y7PPLdyzneH51aJiz
p1LVgucIX/jIobwUVc2fcvRwX7gpcpQmvCJMcJqjI1XoavFt3matpIpyTTR4OhCFq4zkaKt1m3me
Kre0IepCtJSDbiNkQzQs5ZNXSfIC0RvmBb6feA2pObL+8iv+YrOpS3ojyl0D6U9BJGVdHWpbt6qP
1n4l2vgr3pekyJ6ud4+K6kJwrQAOWsBnK1b9IUpT+bu6VfrDjlNXOQpwNIvSMIkT5MjM7IAGI28x
9/7jbmVvvOjk9atTHnJ0iaPI96E15TFHSRqn3UIfW2iIKiWlPDqEJkGbcaGpsm5nS+PWx+isKroh
O6bv6UGv9ZHRxZyYvdVKWuluJR1GzBmg3H1Yd8WPTdie4RZsGiJvcwQpCHuC88OQAzb35HH92mcE
Bpp1JpTc8qV8NiAd0y5ul6DaQio4E6sdL/UJ9LkKBZFwauI8U2mOKLSo0yvB6qqoGesWpsP0B5PO
nkA2fcC25HdWXdaO24aUwO57w12mjSXJKPmgoOSkKNUHRakGHHcGh3fmYdEEA5oonpmCJz4dFMsn
HPj0ECY+4cAnGvjgcIaTCVBPxQKKR4DSIE0nQD0VCygZAAVBmvgToJ6KBTQbAZpF4XRHn6lYQOkA
yNCZLukzFQvocgQoiWfTJX2m0k2yn0fMNgPZzrYgOTtZ5+jvz+K6WAZh6PpJWLhRsIzdFP703Mub
IixivLzG/vU/M3nj2EzEv3Z1RSFIP+Pj+NOU39SlFEps9EUpGvtc8FrxQmUr6u7FgIPTjH8ayaGW
/rcfw8evgsUbUEsDBAoAAAAAABV7tU4AAAAAAAAAAAAAAAAKABwAcHB0L21lZGlhL1VUCQADGfzj
XOrb41x1eAsAAQToAwAABOgDAABQSwMECgAAAAAAFXu1TgAAAAAAAAAAAAAAAA8AHABwcHQvbWVk
aWEvLmtlZXBVVAkAAxn841wZ/ONcdXgLAAEE6AMAAAToAwAAUEsDBAoAAAAAADm1s04AAAAAAAAA
AAAAAAAKABwAcHB0L3RoZW1lL1VUCQADjb/hXI2/4Vx1eAsAAQToAwAABOgDAABQSwMEFAAAAAgA
ObWzTl71rkD0BQAAqBoAABQAHABwcHQvdGhlbWUvdGhlbWUxLnhtbFVUCQADjb/hXI2/4Vx1eAsA
AQToAwAABOgDAADtWUuP2zYQvvdXELo7esv2It7Alu2kzW4SZDcpcqQlWmKWEg2S3l0jCFAkxwIF
iqZFLwV666FoGyABekl/TdoUbQrkL5SS/KBsOo9mAwRobMAWh98MP84Mh5R0/sJpRsAxYhzTvGPY
5ywDoDyiMc6TjnHjcNhoGYALmMeQ0Bx1jBnixoXdj87DHZGiDAGpnvMd2DFSISY7pskjKYb8HJ2g
XPaNKcugkE2WmDGDJ9JsRkzHsgIzgzg3QA4zafXqeIwjBA4Lk8buwviAyJ9c8EIQEXYQlSNu0YiP
7OKPz3hIGDiGpGPI0WJ6cohOhQEI5EJ2dAyr/Bjm7nlzqUTEFl1Fb1h+5npzhfjIKfVYMloqep7v
Bd2lfaeyv4kbNAfBIFjaKwEwiuR87Q2s32v3+v4cq4CqS43tfrPv2jW8Yt/dwHf94lvDuyu8t4Ef
DsOVDxVQdelrfNJ0Qq+G91f4YAPftLp9r1nDl6CU4PxoA235gRsuZruEjCm5pIW3fW/YdObwFcpU
cqzSz8XLMy6DtykbSlgZYihwDsRsgsYwkugQEjxiGOzhJJXpN4E55VJsOdbQcuVv8fXKq9IvcAdB
RbsSRXxDVLACPGJ4IjrGJ9KqoUBePPnpxZNH4MWTh0/vPX5679en9+8/vfeLRvESzBNV8fkPX/7z
3Wfg70ffP3/wtR7PVfwfP3/++29f6YFCBT775uGfjx8++/aLv358oIF3GRyp8EOcIQ6uoBNwnWZy
bpoB0Ii9mcZhCrGq0c0TDnNY6GjQA5HW0FdmkEANrofqHrzJZL3QAS9Ob9cIH6RsKrAGeDnNasB9
SkmPMu2cLhdjqV6Y5ol+cDZVcdchPNaNHa7FdzCdyCzHOpNhimo0rxEZcpigHAlQ9NEjhDRqtzCu
+XUfR4xyOhbgFgY9iLUuOcQjoVe6hDMZl5mOoIx3zTf7N0GPEp35PjquI+WqgERnEpGaGy/CqYCZ
ljHMiIrcgyLVkTyYsajmcC5kpBNEKBjEiHOdzlU2q9G9LCuMPuz7ZJbVkUzgIx1yD1KqIvv0KExh
NtFyxnmqYj/mRzJFIbhGhZYEra+Qoi3jAPOt4b6JkXiztX1DFld9ghQ9U6ZbEojW1+OMjCHK59tB
raRnOH9lfV+r7P6Hyq6v7F2GtUtrvZ5vw61X8ZCyGL//RbwPp/k1JNfNhxr+oYb/H2v4tvV89pV7
VaxN9fBemslecZIfY0IOxIygPV4Wey4nGQ+lsGyUqsvbh0kqL+eD1nAJg+U1YFR8ikV6kMKJHMwu
R0j43HTCwYRyuV0YW22X280026dxJbXtxR2rVIBiJZfbzUIuNydRSYPm6tZsab5sJVwl4JdGX5+E
MlidhKsh0XRfj4RtnRWLtoZFy34ZC1OJilyFABZPPHyvYiSzDhIUF3Gq9BfRPfNIb3NmfdqOZnpt
78wiXSOhpFudhJKGKYzRuviMY91u60PtaGk0W+8i1uZmbSB5vQVO5JpzfWkmgpOOMZYHRXmZTaQ9
XlRPSJK8Y0Ri7uj/UlkmjIs+5GkFK7uq+WdYIAYIzmSuq2Eg+Yqb7TSt95dc23r/PGeuBxmNxygS
WySrpuyrjGh73xJcNOhUkj5I4xMwIlN2HUpH+U27cGCMuVh6M8ZMSe6VF9fK1Xwp1p6krZYoJJMU
zncUtZhX8PJ6SUeZR8l0fVamzoWjZHgWu+6rldaK5pYNpLm1ir27TV5h5epZ+dpa125ZL98l3n5D
UKi19NRcPbVte8cZHgiU4YItfnO2RvMtd4P1rDWV02XZ2nhxQUe3Zeb35aF1SgSvHgicyjuFcPGw
uaoEpXRRXU4FmDLcMe5YftcLHT9sWC1/0PBcz2q0/K7b6Pq+aw982+r3nLvSKSLNbL8aeyjvashs
/l6mlG+8m8kWh+1zEc1MWp6GzVK5fDdjO9vfzQAsPXMncIZtt90LGm23O2x4/V6r0Q6DXqMfhM3+
sB/6rfbwrgGOS7DXdUMvGLQagR2GDS+wCvqtdqPpOU7Xa3ZbA697d+5rOfPF/8K9Ja/dfwFQSwME
FAAAAAgAObWzTsu/Zf2BAQAALwMAABEAHABwcHQvcHJlc1Byb3BzLnhtbFVUCQADjb/hXI2/4Vx1
eAsAAQToAwAABOgDAACt0t1u2yAUB/D7PYXFPeHD2ImtOBWOiVRpF9PUPgDCOEE1xgLSdpr67mNp
2qabJlVTrwCh89fvwFlfPdoxu9c+GDc1gCwwyPSkXG+mfQNub3ZwBbIQ5dTL0U26AT90AFebL+u5
nr0OeooypspvPks5U6hlAw4xzjVCQR20lWHhZj2lu8F5K2M6+j3qvXxI+XZEFOMSWWkmcK73H6l3
w2CU7pw62gR4DvF6PEnCwczhJW3+SNplH+9Im9SkfoxfQzzvsqM3DfgpluVWVIzDEudbyAijsK1E
C8uO5EuMCeZ0+fS7mrC6N0FJ319budeiN7GTUb7gCPuLZ43yLrghLpSz5z7R7B60n505tUrw+b3u
5dgADNBmjU6498YuJxyXlMNlteKQ5bSCvO062LZ8VZQlxQXBr0Y9yOMYT8ZuNp/Io/SfwF1XiB3n
HcRiKyArcgGrVU4gK1uatyItOXsGFrU6SB9vvFR3aWq+66GVQfevzOJ/mPSSSS6R6O3T0Z9DvvkF
UEsDBAoAAAAAADm1s04AAAAAAAAAAAAAAAARABwAcHB0L3NsaWRlTWFzdGVycy9VVAkAA42/4VyN
v+FcdXgLAAEE6AMAAAToAwAAUEsDBBQAAAAIADm1s06OiGvQyQYAANMwAAAhABwAcHB0L3NsaWRl
TWFzdGVycy9zbGlkZU1hc3RlcjEueG1sVVQJAAONv+Fcjb/hXHV4CwABBOgDAAAE6AMAAO1b727b
NhD/vqcQtI+DK1H/bdQpbCfeAqRtUKcPQEu0rYWiNIp2nQ4D+ix7i+1x+iQ7UqIlJ3YQrwmQBEYB
izqSx+P9fnciL+jbd+uMGivCyzRnfRO9sU2DsDhPUjbvm5+vxp3INEqBWYJpzkjfvCGl+e7kp7dF
r6TJe1wKwg1Qwcoe7psLIYqeZZXxgmS4fJMXhEHfLOcZFvDK51bC8RdQnVHLse3AynDKzHo+f8j8
fDZLY3Kax8uMMFEp4YRiAeaXi7QotbbiIdoKTkpQo2ZvmXQC+4snNJHP6bz6/URmRpqswUm2jWAE
7inNZES5scK0b07nyLRO3lr14LolJ5fFFSdEttjqV15MikuuVviwuuSgE1SaBsMZuFcqUB31MKua
pBrWrelz3cS99Yxn8gnuMcBCAPFG/lpSRtbCiCth3EjjxccdY+PF2Y7Rll7Aai0qd1UZd3c7jt7O
VSooMS4pjskipwlwBW12qG0vi4s8vi4NlsPepCuqrW5GVPuXz2JhiJsC1Aqp1tQukZ1W25Byt1cC
Jwr8artu4CMn2PZPGEVBaNf7Rq7j+4G7tXvcK3gpfiV5ZshG3+QkFooIeHVRimqoHqJMKmuDxHqY
Jzdy5BSe4CQIOJi/yPlX06DnrOybXeR5sLZQL54fOvDC2z3TrR5BRzlVKGEWg56+GQuubGHA78FS
5LO0tqhaUnbRUkzEDSVq24X8UWIOBlEs452wzucJxHsmRpRgtqGFOBnRNL42RG6QJBVGHfcKBsgO
oFIuJNRySiVhySXm+NNtzUnKRYtVhfKS9o6lKbWfWO6GWBK1Nq+cx+CVdJVZB/mP0AtFjh84/j38
8lwfuW70/Pl1MKUKifmKbrjzgxST3lMMK7coZunVtpZEBy45IXHOEoOSFaEPUO8cqP5qkfKHa3cP
1D7Ol1wsHqzeO1R9Otup/emC29PBfYrF9kfDfYzgTgTs8ytEBaazOsidHwnywIUPhI+2g9yx/dDT
Qa6+Mv7zj/Gtb4jVDmvVXlEkWYTpHPhBlbEJmUn4pTsRnJqq01BO02ScUrrjaCTW1YlJpExUktC3
bc2UzeDqrdFj6ZVUszakarcMVDyf0USR6E83chAKQrfT9T2n49nuKbTss04wCEanzuh0PHbtv0zN
CWCaSDMyTudLTj4uKygeEh7IckI4MKKgCY6ZPC/uDY//GRS+DopxnsuE2A4L7zHCYgaYKyD/WGIO
K9Sh4R4cGq7tRN37YsO1owC95tjQR7DnFx2Py8lAc3ICthDjwzKb3mKm/xjMhAsmqN5FTu/wvA3A
3kvOV5+4nys1N4nbPjsbhC4669iDMSTuYTjoDCMb8rjb9cNB1x0Px2ebxF1K5jFgx0Pz9fdv//z8
/du/j5CtrfZ9HugD6NctY8lT2Mhw2A2cUTTsDJE37nin3bAzGAd+Z+y7njcaRoORCxuBOcjrxZyo
6sN5ousWyLtTucjSmOdlPhNv4jyrSyBWkX8hvMhTVQVBdl1KURChLorCKPSdbh0nYJt+KmutproR
U/4eF8Z0juDbLhD4dw2t5Bpa07kjZY6UOVIGLRzHhAkYUTe0xNGSzRhXS1wt8bTE0xJfS3wtCbQE
csyCpuwanCEfpjHL6W+VQLeqHANZ4gLf5EtxntRItCRVNQJ5oRe5AVznDd6TEn6e6AR0d7pYK4KW
qi1vuHtPQgZw/ApPJ1/rOK1iUwUmwRdsyK9VZUdWp1j9Cl0LoFnK5pdLFgvZrzSzSRFXaTK+jOtI
69pNpLUHDGVtaXvoJiA3vdPlh5xV17JWzFdGXhPODoh/63Z0gzlySyoUZ5D0++Yv2e8dKuqMim91
EFwXl8pbHXFZ696ZK7a9X6jseQeKDPMLQBhO5XJjKYOkAE7taMHzQUqUdWy2smcLrHEO+bXxzoCn
GKwuMMtLeLUde2gHtgdP/Q9iqEhFvBjjLKXykwWCeIF5ScQm602XI5Aocd/8/u1v8zYdnOip6MD2
0YHtowO7nw6q6TSQB5EfvRDI/eeE+JMlgEdE3GkQdxvEEfJc+wj54ZDbLwByt4Hca0EO8DpHyA+G
HL2EvO41kPutT7m+hx0hf32Q+w3kQQtyH3kv5fh2hPxAyIMG8rAFeTdEx+PbK4U8bCCPGshdz+ke
j2+vFPKogbzbgjyKguPx7ZVC3tVVmlZdpujlYkH4pkoDMy4rYtS7u1tibYZsl3SehCQvzce7Sx/q
zwBH/+wtFGgnHP2z51bthuiJsvBLc9DuOyiKnCg6OuieG5v6jB8dtP9+E3ruMUffdxsAc49J+r6z
c+CHxyS9fdJsHy6t9h9qrdb/Rjj5D1BLAwQKAAAAAAA5tbNOAAAAAAAAAAAAAAAAFwAcAHBwdC9z
bGlkZU1hc3RlcnMvX3JlbHMvVVQJAAONv+Fcjb/hXHV4CwABBOgDAAAE6AMAAFBLAwQUAAAACAA5
tbNONHbfJMoAAAC9AQAALAAcAHBwdC9zbGlkZU1hc3RlcnMvX3JlbHMvc2xpZGVNYXN0ZXIxLnht
bC5yZWxzVVQJAAONv+Fcjb/hXHV4CwABBOgDAAAE6AMAAK2QwWrDMAyG73sKo/vspIdSRp1exqCw
U+keQNhKYprYxlLH8vYzFEoDPeywi0C/9H/60f7wM0/qmwqHFC20ugFF0SUf4mDh6/zxugPFgtHj
lCJZWIjh0L3sTzShVA+PIbOqkMgWRpH8Zgy7kWZknTLFOulTmVFqWwaT0V1wILNpmq0pjwzoVkx1
9BbK0W9AnZdMf2Gnvg+O3pO7zhTlyQkj1UsViGUgsaD1TbnVVlcemOcx2v+MwVPw9IlLusoqzIO+
WronM6uvd79QSwMECgAAAAAAObWzTgAAAAAAAAAAAAAAAAsAHABwcHQvc2xpZGVzL1VUCQADjb/h
XI2/4Vx1eAsAAQToAwAABOgDAABQSwMEFAAAAAgAObWzTn8ikLPfAQAAqgMAABUAHABwcHQvc2xp
ZGVzL3NsaWRlMS54bWxVVAkAA42/4VyNv+FcdXgLAAEE6AMAAAToAwAAjVNLb9swDL7vVwi6J4oT
N0uM2EWdNkOBbS2WFjsrshwL0AuUkqYY9t8ryzaSrT30IlHkR/LjQ6vrk5LoyMEJo3OcjCcYcc1M
JfQ+x89Pm9ECI+eprqg0muf4lTt8XXxZ2czJCgVn7TKa48Z7mxHiWMMVdWNjuQ622oCiPjxhTyqg
LyGokmQ6mcyJokLj3h8+42/qWjB+a9hBce27IMAl9YG4a4R1QzT7mWgWuAthovc/lIpQGdvKqr13
++58hGJFs50UdiOkRJUVOQ59AuN/C99sG2pDYxI8gBBkXO14lWO4r6ZRLQ+KFCvS2VuFA/aLM0+i
7IF71rRiHeL3enJhIOfsLYrXdcB8dxE2ECQDX2efgPNW0sdvYLe2tYaifh4fAYmqZYo0VYEyJr2h
h5HOKQrkP/f9INLsVINq7zARdIqdeG3PWAs/ecQ6JTtrWfPwAZY1dx+gyZCAXCQll2WFHKH2XkIH
CNP4U5bL+XS9KEdlkm5G6e3y6+hmM78aba5mabouFzfr2d3fdrhJmjHgce73w/4G5budUYKBcab2
Y2ZUv3zEmhcO1oi4f8mkX+IjlTlOJ7PpcjFPkinuuhe4DXdkS857xST8oPbhGLsZknkO66iy4YN0
3hcQEr9a8QZQSwMECgAAAAAAObWzTgAAAAAAAAAAAAAAABEAHABwcHQvc2xpZGVzL19yZWxzL1VU
CQADjb/hXI2/4Vx1eAsAAQToAwAABOgDAABQSwMEFAAAAAgAObWzTvI0Zf3QAAAAvQEAACAAHABw
cHQvc2xpZGVzL19yZWxzL3NsaWRlMS54bWwucmVsc1VUCQADjb/hXI2/4Vx1eAsAAQToAwAABOgD
AACtkD1rwzAQhvf+CnF7JDtDKSVylpCQ0iGU9Acc0tkWsT7QKSX+91XpEkOGDh3v63kfbrO9+Ul8
UWYXg4ZWNiAomGhdGDR8nverFxBcMFicYiANMzFsu6fNB01Y6g2PLrGokMAaxlLSq1JsRvLIMiYK
ddLH7LHUMg8qobngQGrdNM8q3zOgWzDF0WrIR7sGcZ4T/YUd+94Z2kVz9RTKgwjlfM2uQMwDFQ1S
Kk/W4W+/lW+nA6jHGu1/avDkLL3jHK9lIXPXXyy1skb8mKnF17tvUEsDBAoAAAAAADm1s04AAAAA
AAAAAAAAAAAKABwAcHB0L19yZWxzL1VUCQADjb/hXI2/4Vx1eAsAAQToAwAABOgDAABQSwMEFAAA
AAgAObWzTm9Tl34CAQAAzwMAAB8AHABwcHQvX3JlbHMvcHJlc2VudGF0aW9uLnhtbC5yZWxzVVQJ
AAONv+Fcjb/hXHV4CwABBOgDAAAE6AMAAK2TTU7DMBBG95zCmj1xUpWCUJ1uUKUukBCEA5hkklg4
tuUxhdweq4WSVFXEIst59nx+/ltvvjrN9uhJWSMgS1JgaEpbKdMIeC2213fAKEhTSW0NCuiRYJNf
rZ9RyxB7qFWOWAwxJKANwd1zTmWLnaTEOjRxpLa+kyGWvuFOlu+yQb5I0xX3wwzIR5lsVwnwu+oW
WNE7/E+2rWtV4oMtPzo04cISPMg3jS+h13ETrJC+wSBgAJOYCPyyyGJOEdKqwj+FQ/lDsymJbHaJ
R0kB/ZnKEY5mTGqtZr2k2Ds4m0N5hJMON3M67BV+PnnrBs/khKYklnNKOI90JnFCvxJ89A/zb1BL
AwQUAAAACAA5tbNOytho5pABAABJAwAAEQAcAHBwdC92aWV3UHJvcHMueG1sVVQJAAONv+Fcjb/h
XHV4CwABBOgDAAAE6AMAAI2TTW/bMAyG7/0Vgu6rnCFNUyNO0WHoLj0MiNe7Kim2Cn1BlFMnv360
7DYfyKE3k3z58qEsrR57a8hORdDeVXR2W1CinPBSu6ai/+rnH0tKIHEnufFOVXSvgD6ub1ah3Gn1
8TcS7HdQ8oq2KYWSMRCtshxufVAOa1sfLU8YxobJyD/Q1xr2sygWzHLt6NQfv9Pvt1st1G8vOqtc
Gk2iMjwhO7Q6wKdb+I5biArQJnefIxkO6RW3qygYWbedfXNcmyFD17i4G0xyiOu3Ph5+8bhBHzwd
y3tt9UHJLMQByUclX9Q2ETjg8d49LO8p4V3yT/K9g1TRgrJTae1DVj7MF4tcYufzBi0YLdUxFBsj
JxhwPNT+T9RyMM7FqbJDRMENIs5yHoZgveIl9GT477MFJdg0K/JQTO+vpNlXXyh91I12pMfifI6q
/aBaTipxpGs6hH2BNBW+WEe3802cTwpq1aeT5U7WvkAeyc5wj6nrqEXmLC4p2dXRDR7jJnCBN5YI
bL7HG4KPQew/P0eX8Rms/wNQSwMEFAAAAAgAObWzTuqwe0qkAAAAtQAAABMAHABwcHQvdGFibGVT
dHlsZXMueG1sVVQJAAONv+Fcjb/hXHV4CwABBOgDAAAE6AMAAA3MyQ6CMBRA0b1f0bx9LdaiSCgE
UFfu1A94ShmSDqZtHGL8d1ne3OQU1dto8lQ+TM5KWC0TIMreXTfZQcL1cqQZkBDRdqidVRI+KkBV
LgrM402f40erU4hkRmzIUcIY4yNnLNxHZTAs3UPZ+fXOG4xz+oF1Hl8zbjTjSbJhBicLpFO9hG/a
cp4KUdPt4bChYi04bRKR0Sxt9u3uuF+16/oHrPwDUEsDBAoAAAAAADm1s04AAAAAAAAAAAAAAAAJ
ABwAZG9jUHJvcHMvVVQJAAONv+Fcjb/hXHV4CwABBOgDAAAE6AMAAFBLAwQUAAAACAA5tbNO4Wsl
KoMZAABbHAAAFwAcAGRvY1Byb3BzL3RodW1ibmFpbC5qcGVnVVQJAAONv+Fcjb/hXHV4CwABBOgD
AAAE6AMAAO14Z1RUS5TuaaIKKDkjICAoUZIgoUUkJ0HJWVBoELDJSfpKzggISJYcFFpyzllylhya
jCBNaBtoul9775uZtWbmrTczv96Pt0/ts06t2nt/+6tTVafq4H7gVgFKdWU1ZQAEAgFW+AvALQCK
ACEBwZ+CFyJ8Ib5GTExERHyDlJTkGvkNcnKyG2RkFDepKSluUt0kI6Okp6SioaWjoyO/xcBIT8tI
TUtH+ycIiBDvQ0R8nZj4Oi0FGQXtf1tw7QDVNaAdBBCC7gAEVCBCKhCuG2AHABAx6G8B/reACPA5
kpBeu36DDG9QQwkQgAgJCYgI/2SNb32HbweIqIipOR8okNDoWpPegdKKvo/Pucb1pKKDTm/8iFvs
pWvg9Rv0DIxMzDx3efnu3ReXkHwoJf1I8amSsoqqmvrzF/oGhkbGJja2r17b2UMc3Nw9PL28fXyD
gkNCw8IjIhMSPyYlp3xKTcvNyy8oLCouKa2sqq6pratvaOzs6u7p7esf+D4xOTU9M/tjbn5tHbGx
ubW9s7uHPD45PUP9Rp9f/OH1h+e/yH/KiwrPi4CIiJCI9A8vEIHXHwMqImLOByTUCrqk1lCaO6Lv
r9E+ic+p6LjOJaZ3RPfSdfwGPbf4Gg/yD7W/mf3XiAX+j5j9K7F/4zUPkBOC8C+PkAoAA0jII+bd
AyzTHoOJGYPgu6IFc3OTloPadOgD04g1re7GmCblJ5J0qcYVx0HxWS+1Q5saMPprcq8XNNtmyjzu
nZmDtU0JizojUKaFMTeeHPDFW4iQDH+nqny2FnoIJunqzYCPYlVOhKjFN2GQ84XytrCz8aMTb0eD
2OUJr0lhepahdAjzUucIqV4Y4xu/F7MWICRdazbMTz7Gz8XFRlzy9117JuaHnx5E03HJhsgi6s+Y
zujQs/k/zOSuvk1nHe5pQPLrPJQie3KqkokpRT4kb+kG6kuWIrb27qHtEBbpubsX1OamljrQgtZs
ToBfoHDZsdfl07f7nuYlz4NOlKJB/ac5ihR+j7Jnz7CmWS2y8jRpI0UwNHsA6tzZRw4HhJzFyB/N
L0EMZ35zToDvOtqpLJ7zNamOmJIStkrlz8qylHl7WDupaNXXqJXlNit7j8imyiNoftzO+L3e4isc
ayKZ4lHJdyazYPkTUUAqouRgrcaToJLFffPyVfS34+PtLTtpsxC4NBIWKX7g2zs60DCV4C3eZEof
gtCPWOaq3gjjLh9jhP6OYETwqunlMS6LNmFbTyC6XZ2CnV6w2WU7raNZUxwQbPl49mWIkEDNYIOA
Ez3DEm9GsNCMxUKEzDvPugJwj+V1f4U1GOWifYyzhJYyS4RgmqCgBUIgnUSB7Hvinb9U4JNysOAA
MQiW5SgzNglq7QGVzHYRR7kqf6cXVFU9cATzJikIpLKp3clej4s1eNQaUT5vbGppsuhAX7WQ35Tk
eZ4IOXRVtGN5s0F6dfa7vL6LYiDA6U2bIIymVvwo6ztm+Sxls2SLbDXTYaeLrTzTw8eHO6H7TdnB
hg/C/vsXkmfNrqvvfNe/DOANJDwbJg+/qgUdUJZNN4c7Mwy9uFoZPPEfkms56fFNiZSXMXa4EA4X
dHI6I/MrU78X/1ogj0QjnamyW9LXREr+Z+vpEgxRmLTOgxRnGHPI0oaX7nlF8fEdK1JMoYRKRVd8
DMZ7H71KqYU2Fiia6uSMM0s57oL3LzCj/PVHN5veJFzBNZsqUglo6xpqk7jT3YJBIAsfLhww+yL/
SCc8YAkl/FNQ56DN7tJwDtvHm7qBuHq91DCHnVVHpVvGYtnH/K0di+ckq3zy54eL0mWYymzCfr41
tVCR77ljtzxEN+BWXVXB7dYDC8OKVyK9gl8eLaGqoHkWJgJitYW/xxQEolMShytMIDZ6YQeg9O1F
ZgmruKDYT7oF9/hKHDyUQ+jPbfjb9k9kwzqXCVuQBjlfTNK/mLNCPOFePpJFW3c6o+4oK9yIXqjr
z2N558Cm2M5QUa2o3hDUnQtJWxbL65FNnzOQlKJgzqosPR1RQPE2t6JrSN4jNdaTW/IGxSje/cqb
PMzVPY5w0XvlSfEjl2oSLFLq69K7HdbRGBOCbE43bsZOrPjQT+kJWVfFKHSwEVcFvR8Ps3DIYzao
OcMBgS6Mzyz2ffdLeuYxwsi7OTbvoC6F2sdGgoPLJp8yX9eqEikre3BY3ZL2HvGWJow+ma8LO9Oi
Z5GuoXh02MvSj1mACFnZz7hjj+kJ324Xr7dUvUD/znO4eJosZurwrDr5oeQvkaeJ7nHWhe35MYj7
tt9c+6TJOKR471S0SmFeT7RyGw3Qp9Itgt9MF0xv19XXRDIepE7axzDZeiYyccEOO65l87K6N7p/
NI2OhdupJNUPvoBerHRcmapMWhi2TjpLSFnVNtaPWyXNoxgPJX24Sx6X66TTG0i+NkdwWCHiJzH8
ZY6JKiSZ+ZZ7NKpG3J61o1yS10hBCiNkfBffzuBukwFy3m4qlZMbHx19T4xrs1AmymVORBEFsjRd
XDUtJ6qTTvy04RU/lQjMWH9ukxqh29aWhDsetYU1P5qfivGz8yJYMe9iq3JOw37fTlKzvEE9/QHT
x/N+lZpUh4d/SXUP5e6lJ1vTyo7OhAuUMXSlVTS1GvgZEUe32u61efc2ubbbJIhFbjuXbVbXJzAh
7F/HRMq63jMwICWi9KX0Mxw/CZgwkYeXOZ7ssS9QmTr4jlQVvyrLaSEh67K1CVXNeroIVHP8jL+R
yGQgo55NgY6CG6Hz+kwO4YWKbi7Wfj6C3Q+WNSc16T5lZ86lRkiPCvRlGTemQYMDGNwMb0ANg8td
h5c/J5ZpL+1dfzT9tNjZMLFGgUeH1yFQfTdijrrpBbWHhMyf2yRNKY9XMfev5y0KvL47EdPU1x8x
w3l9K3gV9Td+mT94tKrAg/fDq/vTvdjtU04cQM5zhDmL1ZnFQMBrQVh22Og9sCUSnr/zVb4oFAdk
617GD19lB+GfHuCAdsNzdgkc0EXiL4sDsvQC2hBaOCD0OtoLB8iy4wco84VJJOx5F2y0/qofXmwN
3h6CHdFhVYZwwFE7DripY+p1BeuNQ1tZhrGfHOIAOHp2kvb5RSE79l3focIZKy/mIg5L0NZ/noJl
JF2rw1w+wOAAdKkLRuZyrQhzUbePlTjdpH1neAsHyM8uhp0JCOMdFvH24TigQx7LlH3u7jKLAyLz
wUc62aEqZekBbRj9d3JK4C/P8Pk7HipeZLH8gVIZ3MZe3/+T0mD8FQxdqoP2KurEAedIMBwH0Fk+
G9fZO4XFv+uHl1rjgO3FgA3/Bnz2SLSJDsYbBsbeT8fXDtDGOOAxDrgiwwPCS+yLznzB/a0b4NKx
v6FUBsFHh20hsNObOGCs7W8Ilc7EWS32eByw/mXPlw9F24ZHkC9TPGdYjEvCAWuLbddxwA4jGPnr
HwAEWmtFw3IA73tRGhnQm3IMr4d/Gccq/ebaxAe71WlxoYfF936PHqyrLoAKj8ID++Pa+jU3u6II
8+XwlrUGPrR/Lmw1HXYtDtNJc+yCvZ+IZas4+utsfPJPpRWem10Z9ieujc654aJ/Jh51Gg3G07LF
pP/TjVrCQUXlkdmJMKT+WRKtn5Snws8SHHD/Mi4RvDaHA0hGMR1c/9jigLDXGGYc8JkV9rd1Ae05
1w+43c3+c14sW+KRTWsIFYYFb1iIHz4VWLLs8796zh/+6W0zXZ1kuaAEubj4trUfbcRtO4QMGysY
ffzIDLTC0IJPgLC/c23MzYw3+1JxzI55hANWU3FAagD7wN8d6/Grg2a2aywbTfYH94C2l1ZXCaDb
XBZEI23WwNFZzGg/oXg/N79ypa0+r2vfOa62pfqGvmyy9XE5n0pJS50xO49pQ4ft9EQTzGJ73aLz
U7gqVsGMMPsymzXLaAktkV3J3zMm73L3JkT5z0NHlC5l74u8feW4/mV38xDBJVGzBp4XWjAa9zf1
eFdg9niZLmp4yu7aePdOog75McdbqeeOYqv0554xC3X3LzC1BicVEM5YZY57Ap8TQP3s2RK0j+Ri
msw/fT1APHRqXM1Xl93qFcX497G6zX8dO3ISRvWPnam+ea1FqPeet+WSaHzRSo28KVeTEsS4ImJW
xSfDDnkpmTlMirjCAeGHZz+sj0497TTGbacizSxI7FIG6rKksPq9k2xDfulVJIAIKO7LkreLWmsC
k5D7oEwE50uk7Dcu2pcKPGqsrKVZcqQFOleJtrw0UapEC+QBVjFeTjIKZoicHxmwSjliwEclrM0+
yslQbfqXkHtY+O68/nLxAPF3dTOT2oGx2BuvcMB08uX+CTSQn8vgqohnfPtp89EQaqTKrCiRrduN
qZvKXKy09gXnknM9L0coF2ExKl34FsTX/pPJ7kxtp63R5rGV/gmdocMXK1eCI9COVMmKZtWHQNS3
io3NPNJyt4VWnS1IOFIHDVHDAbXs2nFdrmbLZKkCy3QNT43yYy9YGK7WRo5eZ16TMdwgVx2hxS8D
oaU1dVUtLemqkNXa2gSsn9FrXQ8nl9uZxAtsGRHykGCrwPxkRj26tJbRS8lJqEtlJ2tmT3NFxiFF
VvYycKXsQHqESuHeldmtjRMlF4ghSNaNfgvd7eHMklkfnWd5Vol2KUCnq4wlbTU3Tgksz1yoROcn
z8y6bireTyafqT+vpWT5cI0D4m2HiqhGknaTCV9rDfNxLNYSQFVpXvd+mcbbpEd4cKM/oEfopqL/
wPG7q6zPreyT1CPz4+77+jW10oHEIU+evlX5Xt95NCT1U6dMtVEucIVnp7KOIfxrFHRT+BNd7Bao
IoNmgG0Y8LmdcqljXzcfPOMGOTHRrpK1h4ti9oScrJ6IAjLzB6p8xtq9XRKWaDckAkszrGCV0jXL
bCb8cnb8+yz/g1f0KC8W9eLVc6+jC8x3xAfCh7nwkFPbeHP3dDqPzPIgMvf+1FBmTcSv+OnGtCaK
mJYGEj6u6Pd/6QLt7YyYJBUMke8z0g9ofsTP1odInmooZOYiIDrNfCbHSyCFXyywwmiYfub5t1v3
f2+T5gmB5gklXl3lCMkqHKF0KHfSXi7oOH3pKzJXDGAN1uHh9dTnZEof9ADY8qC8Ap/pUSxs9F1a
cYNVlXU1mk1OoTwKydI27Yc+8oV5344OUwoWJFqSMyEQ5+HmHqIb4aaAHWAjBbwjdlLWLdJVh25p
fM7qi3aAVRdMeT4YtMshTi3njhSu0STQLnbRHeBPZBi13Xh7bqDbwV2+nHKpaLSHZUDerCKJ86VR
F3GKfPxs5O6Hvtsg84W6e99zG3Zo/Nhqdn8zy2dDyRQlMps/fHtlv8xzX2FEgXHHSxK2LF+9glY5
Ha1cWoDZzV+lIVMPrY8NbXtakdDoB3InM037s50vChHqcQKN8q9xQBm4fOJj4nQWp82loal3kF5d
xfjpvSw13kkE5DAzL/z6VkWspuQQde6VCGiAbyGQoqm2WQ25t7XSS4WfvhNrDwVLe2ydhbaJfByh
jnSDjClbz6Oqj4/f/hDwBsAlt8h3bwx1aVSRD7DaXmWKGoh9mi1S4qCUsYnMdPcjbklDvinrWVyQ
HejWaEFdDdWKQhcmmivjWJE6BzKQtZbnFt/MCbaOERrJr+brbmS5v10w1+Ev4YK0V3uANPyHclBs
HcbFAVMGvuJ1Dyx3aLp4crJYgmOP7wyKw960P4M2igq8GgoDeQkpE35ESYy5T9c31NPT5/olkB+f
8zfq0RYLQE/882P9qO+9GjvdLQ8ZSqaHQzebLRyfFKCkB8uhGQ/2ZCaYVi/CDhRKTPZm5S3DBLcm
BTizHFeE0LJNVpfeGb2eIRwS7+euDlYpMw72mrWQHRgGBJwa0liTz6X+bZIaNaK6QSYwUmrywqmb
PHM8mPDyq9oGaNcTyzZRRmdAITlCZ6TUGL6qRMCZwdF+mecGndySr75cotGZSpNsqqhrtF9UjRFX
3zQiICSPvXmd6bc4aCMZVh3sNV+WlSZOUtvYVKv1RLhmj7KHqQZiwLGQHifSRK6y5EMKFnQzMTzw
oA1gy9abdhe1mB+ZON7jl56iZaDf5YZQ2qQebQaJkFV4JtyCfeYdO91Pfl/wrTa+bMk2JIvurqTZ
wStEoOhbFc20EFnhvN4xuomvHxdCjaDOQYPkOUbCrAQAKdvjHsDHI/uAgXZZ4DZb/Vh+1dRT119i
FrwfuY22kp4PLYR8f9ydsjC7eZP6SVZ/EQdznmRQoZk5WYn7pbHQE/GXdiEaIc5Mrspvof2/DHf/
0lWiBf5PmppPfY+bhCg2T45/VF4QbWiEfPKjMHWddAwRKJ2leDtHME9mSzRa07+ymNr2trx1B5y/
bswjk1DXb2ElvDZwohdZA1TsnVSpvbk7ifKpZ8qI8na2OR/cGGlRW/wcV2CxzveDNJ4W2SXCpk26
r31CMzTX4IWQcRuRcxyQ0AH2kmHU+P0Qv5EBVqPMQCSS98NXg4rdyr5r0EBKstRC3pcfZKfdvmfI
dMiTI7WDir21Bd30Eo2nnbQuVxxabZpd3nTo8ZGsimgveHI1vP+RYJQQ6rvkWmfBN+WE5YVk3Zas
bZszTFNaQjeeMI5udUlUvpmr5jMzs5PmzhqMS7s7htGb7/m6C//kfmorOvcxUqud/ZPLA5LtvUeo
Qofb4hWTWM45e7G9mUE/18hxY4pHdh/4vxa4vWUZ1Opd1iMEZfO9TxB7zKoSKTEMXr/bUOjQeCvC
w16wGfxOMPOqNTt/eSiFPLEPFeXz5O5DEP34Iv0ox+pFa9A6yy0qCP1FUjDWKf/82gSX8hxRiYEh
KCmI4Ab+KxLIP+D+zrHTRKVXh8Y7tqvHxCK+UMcFJec847eTwD/85K6+2hQ9iwXf9JoXFhyvIzIz
G84ZtS4xG9isv5ae7uSoNHlYUF+XNrN67qp+WIv1RWzE0zM0OS1u/OTP2GADTn9VK22myk79uEnm
zNgvxQW2uHe29sgl0F9PFlw4WOwz4e072Zk9NxvAGguR3W/aDC/xNG/S1hjKkGT96Ko93fcNbUCh
PduFor9o6hPKIafqExiXtuYnBmVskFO7Qi8mCU0i+T4T/Bf1pl0P+y33bIaVBYs9MjGz+Q6Jlub4
fsPvxi+uX2MTfdv8AVrtmzFdgzbM6fkCg5jlp2XOvXdagtf2s9GWMLFAM0proVOG1krA35vAIOhq
zO397geOrvO/ptaT7qoNxX0YoB0XZuolMNN8EP1qYH/YRfcxK+HOfHdawTobocalwLG6TaP3ul2w
azEUYzzPnv0LUeORxRh6uNKVZuYbqHL7c20F2cdwPl2hjmag2cbTF2ri/gykRPv/9f9JhYmDj8bZ
GWE7hd4IHEA6m41xjUPbx0WDT5/LQnFA0P4oVvox+8QZHQZ/bOt66o8/7+WmIuIw7BYw5HfwWq68
CA4Y10T2YW/BwehNAKYxZXjScjSKKAqX0JAzAbP4v0Kn5pmfJ+8XDRftM7gpVVedb+mvJP5F4zw5
lUXh/g1jYurXPwObp1l3WmpONiQfJr46fc2fRHmucULTYFR9FSp5V3VYBAzjg7cv4A9H9LpfWFzm
n68Hh7zN8V9YNzu0Mk1lACjCWE9lMT8M2CMN0vPDsMwq62FLQikvHdellgkYMe0K4CZpcpVn0Ixg
0l9ldGiaIn+z2suAFpIcJP6o+Vedfs0jpaXyEt+OgZievqxoEOMHLqGVVSlp59jE9bIo60lUHPoL
80/9B6kuTFEhLOny9iAN2/cDwOo2Ewd760fDExzghvQ3XjQWgFWT9rApM3E2nWxo7PebSHsq1+ZB
3JJOmJb6uT8++UzCka7nb1CZ5nncZv8rsrYiwdgA4sLdzm2tcT1UlWizYoCNxNsq4FLX8r2vcM+g
TLgH2sNynnQ9KWbCgbowyc2pZHmUbC6e5ulxyfyy1W0H0WiKHtV2y7SPlozRaQUb4KNWLB9P3FyO
5G/XPcPIZdgnA31bvTnz92asgaLKrO8/3NohicSQhK0vfa2jJr3Kwd5XuK2qNWsYLX8I4lbLcD5v
/aUY2zyb7f8EfY4YZGdYUVjFAVFgyhr2m/AUlGPUgeeR5rmScL2zRVz9zUcg1Vs723eiZzfCQgOE
vAOdtHs8sm+jEyjEOBblakvENKrJP3xufXg7SJZ8Y9ZghUrQYEkHI7hCz5TMWlnWvWnxWAVIwS+k
21L3fWQNw6ukBixTkCqhVaqDPwP4rp4OM8EHTeHhEXQfJxSQfj00i8VDb1MPaaXPpXhhbJbH7NHJ
bdRlto0tsbDIZeFIihIUJPnwpVv94WLTOSMha1zLgc6luowrlg82D21pgNEG5L2Ont0ihneXiWpx
+uMHUatolN7KnQrTORPZ66xcoCawqI8d0+rhfgZJXbcGvRXN+nbUM3omUbPR43KfXsaN09MdCA6A
qJzDaMEHfDhAeD1uinU8DVM/3cq1oGQnkMxO728q0dlzrMQztSr76WxZJB9rcQWJxzwLvNfc0HrX
u9P9uFeIc4pHt7ly+76pFFlbm5nNFPyfUCFXReybGN2LpDQxy5TnGNVqDVPDsPW6qILs0WgdmmXN
z91mzpx1bAWGXr8rVRR0XMI9sx3vywt7V1hBAgwYaCpUTzCWUgN9n08vSK8K8JGe4YBqBKYtEvys
StnxmN4lalnYmJU31mP2saoBdvqTrafUiiYOCOlAsF+wna1gWP78IJkBz49jlW2u4OwbdvuH9sMB
CrDOVWTc1c2LX1iKQvwc2fn3HuNXZewb1v83B+H/4FBxa8MON/e/AFBLAwQUAAAACAA5tbNOJ6yZ
SVsBAAC2AgAAEQAcAGRvY1Byb3BzL2NvcmUueG1sVVQJAAONv+Fcjb/hXHV4CwABBOgDAAAE6AMA
AI2SyW7CMBRF9/2KyPvEmUqrKAnqIFZFigRVq+4s5wFW40G2S6BfXyeQAIJFl9Y97+T5Ovl0xxtv
C9owKQoUBSHyQFBZM7Eu0Pty5j8iz1giatJIAQXag0HT8i6nKqNSQ6WlAm0ZGM+JhMmoKtDGWpVh
bOgGODGBI4QLV1JzYt1Rr7Ei9JusAcdhOMEcLKmJJbgT+mo0oqOypqNS/eimF9QUQwMchDU4CiJ8
Yi1obm4O9MkZyZndK7iJDuFI7wwbwbZtgzbpUbd/hD/nb4v+qj4TXVUUUJnXNLPMNlBWsgVdSSas
V2kwbmNiXdc5HomOpRqIlbp8apiGX+ItNmRFgPXUkHWdN8TYuXudFYP6eX+NXyPdlIYt6x64THpi
PObHug6fgNpz18wOpQzJR/LyupyhMg6jiR9GfvywjOMsTrM0/eq2u5g/CflxgX8b0yi7T86Mg6Ds
N7781co/UEsDBBQAAAAIADm1s064rXyZCwIAAFYFAAAQABwAZG9jUHJvcHMvYXBwLnhtbFVUCQAD
jb/hXI2/4Vx1eAsAAQToAwAABOgDAACdVMFuGjEQvfcrLJ/aAxgIjSJkHEVEiEMoSCzJ2VnPsla9
tmW7JOnX1+uFzdKipIlPb2aenkfPM6bXz5VCe3BeGj3Fw/4AI9C5EVLvpnibzXtXGPnAteDKaJji
F/D4mn2ha2csuCDBo6ig/RSXIdgJIT4voeK+H8s6VgrjKh5i6HbEFIXM4dbkvyrQgYwGg0sCzwG0
ANGzrSBuFCf78FlRYfK6P3+fvdiox2gGlVU8AFslNsqiHFDSpmlmAleZrICNY7oN6INxwrMBJQ2g
N9YqmfMQ3WJLmTvjTRHQQXVtnsCtjdSBki4xmgU+NpeieeqdrXTP5w5Ao01pntDX8eTiGyVniHTN
Hd85bsvURyeiGyUFeDai5IDoDxMg0RpAF1II0IdqTJ/EdLmcKWlT4QjpJucKZtE8VnDlo0evCboA
Xs/FmksXmfsw2UMejENe/o6TcYnRI/dQWz7Fe+4k1wE3tCZIWFkfHJsbHTzaehCUtMkEu9wulmN2
kQgRvElstA4P/N/aww9oJ/tQJoMC/4ErRuevIK2PEZ863FyxKuKbh/cMTz3gTpc3UV9122vRjCv5
6ORbNXQnd2U4yzjdoDOE1y1A3XH+LPfEn78cuZP6p9/azNzWS3wY2NMk3ZTcgYj/QzvQbYIuonVO
1fxZyfUOxJHzb6He/Pvmm2TD7/1BPGnJj7l6d48fGPsDUEsDBAoAAAAAADm1s04AAAAAAAAAAAAA
AAAGABwAX3JlbHMvVVQJAAONv+Fcjb/hXHV4CwABBOgDAAAE6AMAAFBLAwQUAAAACAA5tbNOnvYV
6PsAAADhAgAACwAcAF9yZWxzLy5yZWxzVVQJAAONv+Fcjb/hXHV4CwABBOgDAAAE6AMAAK2S20oD
MRCG732KMPfdbKuISLO9EaF3IusDjMnsburmQDKV9u0NBQ8LaxHsZWb++fgmyXpzcKN4p5Rt8AqW
VQ2CvA7G+l7BS/u4uAORGb3BMXhScKQMm+Zq/UwjcpnJg41ZFIjPCgbmeC9l1gM5zFWI5EunC8kh
l2PqZUT9hj3JVV3fyvSTAc2EKbZGQdqaaxDtMdL/2NIRo0FGqUOiRUxlOrEtq4gWU0+swAT9VMr5
lKgKGeS80OqyQjzs3atHO86ofPWqXaT+N6Hl34VC11lND0HvHXme85omvp1iZBkT5VI8pc/d0M0l
hejA5A2Z84+GMX4aycnPbD4AUEsBAh4DFAAAAAgAObWzTpCsXBGHAQAA4AYAABMAGAAAAAAAAQAA
AKSBAAAAAFtDb250ZW50X1R5cGVzXS54bWxVVAUAA42/4Vx1eAsAAQToAwAABOgDAABQSwECHgMK
AAAAAADraLVOAAAAAAAAAAAAAAAABAAYAAAAAAAAABAA7UHUAQAAcHB0L1VUBQAD6tvjXHV4CwAB
BOgDAAAE6AMAAFBLAQIeAwoAAAAAADm1s04AAAAAAAAAAAAAAAARABgAAAAAAAAAEADtQRICAABw
cHQvc2xpZGVMYXlvdXRzL1VUBQADjb/hXHV4CwABBOgDAAAE6AMAAFBLAQIeAxQAAAAIADm1s05w
iDQz4gIAAIEHAAAhABgAAAAAAAEAAACkgV0CAABwcHQvc2xpZGVMYXlvdXRzL3NsaWRlTGF5b3V0
MS54bWxVVAUAA42/4Vx1eAsAAQToAwAABOgDAABQSwECHgMKAAAAAAA5tbNOAAAAAAAAAAAAAAAA
FwAYAAAAAAAAABAA7UGaBQAAcHB0L3NsaWRlTGF5b3V0cy9fcmVscy9VVAUAA42/4Vx1eAsAAQTo
AwAABOgDAABQSwECHgMUAAAACAA5tbNOQPtwWrQAAAA2AQAALAAYAAAAAAABAAAApIHrBQAAcHB0
L3NsaWRlTGF5b3V0cy9fcmVscy9zbGlkZUxheW91dDEueG1sLnJlbHNVVAUAA42/4Vx1eAsAAQTo
AwAABOgDAABQSwECHgMUAAAACAA5tbNOmNEy8oICAAAkDQAAFAAYAAAAAAABAAAApIEFBwAAcHB0
L3ByZXNlbnRhdGlvbi54bWxVVAUAA42/4Vx1eAsAAQToAwAABOgDAABQSwECHgMKAAAAAAAVe7VO
AAAAAAAAAAAAAAAACgAYAAAAAAAAABAA7UHVCQAAcHB0L21lZGlhL1VUBQADGfzjXHV4CwABBOgD
AAAE6AMAAFBLAQIeAwoAAAAAABV7tU4AAAAAAAAAAAAAAAAPABgAAAAAAAAAAACkgRkKAABwcHQv
bWVkaWEvLmtlZXBVVAUAAxn841x1eAsAAQToAwAABOgDAABQSwECHgMKAAAAAAA5tbNOAAAAAAAA
AAAAAAAACgAYAAAAAAAAABAA7UFiCgAAcHB0L3RoZW1lL1VUBQADjb/hXHV4CwABBOgDAAAE6AMA
AFBLAQIeAxQAAAAIADm1s05e9a5A9AUAAKgaAAAUABgAAAAAAAEAAACkgaYKAABwcHQvdGhlbWUv
dGhlbWUxLnhtbFVUBQADjb/hXHV4CwABBOgDAAAE6AMAAFBLAQIeAxQAAAAIADm1s07Lv2X9gQEA
AC8DAAARABgAAAAAAAEAAACkgegQAABwcHQvcHJlc1Byb3BzLnhtbFVUBQADjb/hXHV4CwABBOgD
AAAE6AMAAFBLAQIeAwoAAAAAADm1s04AAAAAAAAAAAAAAAARABgAAAAAAAAAEADtQbQSAABwcHQv
c2xpZGVNYXN0ZXJzL1VUBQADjb/hXHV4CwABBOgDAAAE6AMAAFBLAQIeAxQAAAAIADm1s06OiGvQ
yQYAANMwAAAhABgAAAAAAAEAAACkgf8SAABwcHQvc2xpZGVNYXN0ZXJzL3NsaWRlTWFzdGVyMS54
bWxVVAUAA42/4Vx1eAsAAQToAwAABOgDAABQSwECHgMKAAAAAAA5tbNOAAAAAAAAAAAAAAAAFwAY
AAAAAAAAABAA7UEjGgAAcHB0L3NsaWRlTWFzdGVycy9fcmVscy9VVAUAA42/4Vx1eAsAAQToAwAA
BOgDAABQSwECHgMUAAAACAA5tbNONHbfJMoAAAC9AQAALAAYAAAAAAABAAAApIF0GgAAcHB0L3Ns
aWRlTWFzdGVycy9fcmVscy9zbGlkZU1hc3RlcjEueG1sLnJlbHNVVAUAA42/4Vx1eAsAAQToAwAA
BOgDAABQSwECHgMKAAAAAAA5tbNOAAAAAAAAAAAAAAAACwAYAAAAAAAAABAA7UGkGwAAcHB0L3Ns
aWRlcy9VVAUAA42/4Vx1eAsAAQToAwAABOgDAABQSwECHgMUAAAACAA5tbNOfyKQs98BAACqAwAA
FQAYAAAAAAABAAAApIHpGwAAcHB0L3NsaWRlcy9zbGlkZTEueG1sVVQFAAONv+FcdXgLAAEE6AMA
AAToAwAAUEsBAh4DCgAAAAAAObWzTgAAAAAAAAAAAAAAABEAGAAAAAAAAAAQAO1BFx4AAHBwdC9z
bGlkZXMvX3JlbHMvVVQFAAONv+FcdXgLAAEE6AMAAAToAwAAUEsBAh4DFAAAAAgAObWzTvI0Zf3Q
AAAAvQEAACAAGAAAAAAAAQAAAKSBYh4AAHBwdC9zbGlkZXMvX3JlbHMvc2xpZGUxLnhtbC5yZWxz
VVQFAAONv+FcdXgLAAEE6AMAAAToAwAAUEsBAh4DCgAAAAAAObWzTgAAAAAAAAAAAAAAAAoAGAAA
AAAAAAAQAO1BjB8AAHBwdC9fcmVscy9VVAUAA42/4Vx1eAsAAQToAwAABOgDAABQSwECHgMUAAAA
CAA5tbNOb1OXfgIBAADPAwAAHwAYAAAAAAABAAAApIHQHwAAcHB0L19yZWxzL3ByZXNlbnRhdGlv
bi54bWwucmVsc1VUBQADjb/hXHV4CwABBOgDAAAE6AMAAFBLAQIeAxQAAAAIADm1s07K2GjmkAEA
AEkDAAARABgAAAAAAAEAAACkgSshAABwcHQvdmlld1Byb3BzLnhtbFVUBQADjb/hXHV4CwABBOgD
AAAE6AMAAFBLAQIeAxQAAAAIADm1s07qsHtKpAAAALUAAAATABgAAAAAAAEAAACkgQYjAABwcHQv
dGFibGVTdHlsZXMueG1sVVQFAAONv+FcdXgLAAEE6AMAAAToAwAAUEsBAh4DCgAAAAAAObWzTgAA
AAAAAAAAAAAAAAkAGAAAAAAAAAAQAO1B9yMAAGRvY1Byb3BzL1VUBQADjb/hXHV4CwABBOgDAAAE
6AMAAFBLAQIeAxQAAAAIADm1s07hayUqgxkAAFscAAAXABgAAAAAAAAAAACkgTokAABkb2NQcm9w
cy90aHVtYm5haWwuanBlZ1VUBQADjb/hXHV4CwABBOgDAAAE6AMAAFBLAQIeAxQAAAAIADm1s04n
rJlJWwEAALYCAAARABgAAAAAAAEAAACkgQ4+AABkb2NQcm9wcy9jb3JlLnhtbFVUBQADjb/hXHV4
CwABBOgDAAAE6AMAAFBLAQIeAxQAAAAIADm1s064rXyZCwIAAFYFAAAQABgAAAAAAAEAAACkgbQ/
AABkb2NQcm9wcy9hcHAueG1sVVQFAAONv+FcdXgLAAEE6AMAAAToAwAAUEsBAh4DCgAAAAAAObWz
TgAAAAAAAAAAAAAAAAYAGAAAAAAAAAAQAO1BCUIAAF9yZWxzL1VUBQADjb/hXHV4CwABBOgDAAAE
6AMAAFBLAQIeAxQAAAAIADm1s06e9hXo+wAAAOECAAALABgAAAAAAAEAAACkgUlCAABfcmVscy8u
cmVsc1VUBQADjb/hXHV4CwABBOgDAAAE6AMAAFBLBQYAAAAAHgAeAIEKAACJQwAAAAA=
TEMPLATE
}

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

# temporary file is needed as explained here https://unix.stackexchange.com/questions/2690/how-to-redirect-output-of-wget-as-input-to-unzip
# in short: unzip <(...) is only supported in busybox implementation...
# busybox unzip -q <(templatedata | base64 -d) -d "$tempname"
# other options: use tar,... instead of unzip
templatedata | base64 -d > "$tempname/template.zip"
unzip -q "$tempname/template.zip" -d "$tempname"
rm "$tempname/template.zip"

# $colorspace may contain multiple parameters passed to convert
# shellcheck disable=SC2086
if convert -density "$density" $colorspace -resize "x${resolution}" "$1" "$tempname/ppt/media/slide.png"; then
	echo "Extraction successful!"
else
	echo "Error during extraction"
	exit 1
fi

pushd "$tempname/ppt/media/" || exit 1
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

	convert -resize 256x slide-0.png ../../docProps/thumbnail.jpeg

popd || exit 1

pushd "$tempname" || exit 1
	rm $fout
	zip -q -r "$fout" .
popd || exit 1
