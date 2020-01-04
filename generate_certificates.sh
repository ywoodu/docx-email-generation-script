#! /bin/sh

template_name=$1
csv_name=$2
rendered_doc="tmp.docx"
rendered_pdf="tmp.pdf"
template_string="AttendeeName"

mkdir tmp/
mkdir generated_certificates/
unzip "$template_name" -d tmp/
cp tmp/word/document.xml document.xml
while IFS=, read -r last_name first_name email
do
  name="$first_name $last_name"
  qualified_file_name="${last_name}_$first_name.pdf"
  echo $name
  sed -i "s/$template_string/$name/g" tmp/word/document.xml
  cd tmp/ && zip -r ../$rendered_doc * && cd ..
  cp document.xml tmp/word/document.xml
  libreoffice --headless --convert-to pdf $rendered_doc
  mv $rendered_pdf $qualified_file_name
  echo "Here is your certificate." | mail -s "Test Certificate" -a $qualified_file_name $email
  mv $qualified_file_name generated_certificates/
  rm $rendered_doc
done < $csv_name
rm -rf tmp/
rm document.xml
