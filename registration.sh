#!/bin/bash

year=$(date +%Y)
month=$(date +%b)

name= "" #Insert Name
vacationPath= "" #Vacation fullpath
signaturePath="" #Signature fullpath
outputPath="" #"Documents Full path concat with $name-$month-$year.xlsx"

java -jar "/Users/#YOUR USER IN OSX/Documents/registration.jar" cli "$month" "$year" "$name" "$vacationPath" "$signaturePath" "$outputPath"

echo "$(date) - Created registration file for $month/$year - $outputPath"