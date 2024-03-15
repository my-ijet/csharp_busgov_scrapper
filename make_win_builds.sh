#!/bin/bash

dotnet publish -r win-x64 -c release -p:PublishSingleFile=true -p:PublishTrimmed=true ;
dotnet publish -r win-x86 -c release -p:PublishSingleFile=true -p:PublishTrimmed=true ;

cp -r ./bin/Release/net8.0/win-x64/publish ./win-x64
cp -r ./bin/Release/net8.0/win-x86/publish ./win-x86

cp ./список.xlsx ./win-x64/
cp ./список.xlsx ./win-x86/

rm -rf ./win-x64/selenium-manager/linux
rm -rf ./win-x64/selenium-manager/macos
rm -rf ./win-x86/selenium-manager/linux
rm -rf ./win-x86/selenium-manager/macos

rm ./win-x64/*.pdb
rm ./win-x86/*.pdb

cd ./win-x64
zip -r ../busgov_scrapper-win-x64.zip *
cd ../win-x86
zip -r ../busgov_scrapper-win-x86.zip *

cd ../
rm -rf win-x64
rm -rf win-x86
