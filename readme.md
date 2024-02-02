for make a release
```bash
dotnet publish -r win-x64 -c release -p:PublishSingleFile=true -p:PublishTrimmed=true
```
or
```bash
dotnet publish -r win-x64 -c release -p:PublishAOT=true
```