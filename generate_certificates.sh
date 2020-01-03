#! /bin/sh

template_name=$1
new_name=$2
rendered_doc="tmp.docx"
rendered_pdf="tmp.pdf"
template_string="AttendeeName"

mkdir tmp/
unzip "$template_name" -d tmp/
cp tmp/word/document.xml document.xml
sed -i "s/$template_string/$new_name/g" tmp/word/document.xml
cd tmp/ && zip -r ../$rendered_doc * && cd ..
cp document.xml tmp/word/document.xml
libreoffice --headless --convert-to pdf $rendered_doc
rm -rf tmp/
rm document.xml
rm $rendered_doc
echo "Here is your certificate." | mail -s "Test Certificate" -a $rendered_pdf bhig93@gmail.com
rm $rendered_pdf
