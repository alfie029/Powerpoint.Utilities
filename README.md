# Powerpoint Utilities

## Powerpoint.DefaultFontSize

[![Build-CSharp](https://github.com/alfie029/Powerpoint.Utilities/actions/workflows/build-csharp.yml/badge.svg)](https://github.com/alfie029/Powerpoint.Utilities/actions/workflows/build-csharp.yml)

At some particular case, user wanna change
the default font-size to value rather than its default (18pt), 
this tool was in front of Office Open XML and give helps.

### Installation
download nuget package from 
[https://github.com/alfie029/Powerpoint.Utilities/releases](https://github.com/alfie029/Powerpoint.Utilities/releases)

and install via `dotnet tool`
```
dotnet tool install --global --add-source ./ Powerpoint.DefaultFontSize
```

##### remove the tool
```
dotnet tool uninstall -g Powerpoint.DefaultFontSize
```

### Basic Usage

- #### read default font-size from given PPTX file
```
Powerpoint.DefaultFontSize get --file YourSlides.pptx
```

- #### change default font-size for given PPTX file
```
Powerpoint.DefaultFontSize set --size 16 --file YourSlides.pptx
```

### Build

- clone the code
```
git clone https://github.com/alfie029/Powerpoint.Utilities.git
cd Powerpoint.Utilities
```
- restore package
```
dotnet restore  # full solution
dotnet restore Powerpoint.DefaultFontSize/Powerpoint.DefaultFontSize.csproj
```
