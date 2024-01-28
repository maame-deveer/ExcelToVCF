# ExcelToVCF Converter with QR Code Generation



## Description

The ExcelToVCF Converter with QR Code Generation is a simple class library designed to streamline the extraction of contact data from Excel files, the conversion of this data into virtual contact files (.vcf), and the automatic generation of scannable QR codes for each contact. This library simplifies the process of handling contact information, making it easy to integrate into your projects.

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
- [License](#license)

## Installation

1. Download the library directly from this repo as it is currently not available as an npm package.

2. Extract the downloaded zip file to a location of your choice.

3. In your project, include the library by referencing the ExtractAndConvert.csproj file in your project.

```bash
// Example for including the library in your project, in your .csproj file
<ItemGroup>
    <PackageReference Include="./path/to/ExtractAndConvert.csproj"/>
</ItemGroup>

```
## Usage

// Example usage for converting details in worksheet to virtual contact files (.vcf)
```bash
using ExtractAndConvert.SheetToVCF;

string worksheet = "./path/to/worksheet";
string fileSaveDirectory = "./path/to/foldertosavefilesto";

//generate contact vcf files
WorkSheetToVCF toVCF = new WorkSheetToVCF();
toVCF.ToVCF(worksheet, fileSaveDirectory);
```

//Example usage for converting .vcf file to qrcode
```bash
using ExtractAndConvert.GenQrCode;

string vcfFile = "./path/to/vcffile";
string imageSaveDirectory = "./path/to/foldertosaveqrcodesto";

//generate qrcode
VCFToQRCode qrCode = new VCFToQRCode();
qrCode.ConvertToQRCode(vcfFile,imageSaveDirectory);
```

## License

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)