# Binary(doc,xls,ppt) to OpenXMLTranslator
Port and update of http://b2xtranslator.sourceforge.net/

Original code: https://sourceforge.net/projects/b2xtranslator/

The main goal of the Office Binary (doc, xls, ppt) Translator to Open XML Project is to create software tools, plus guidance, showing how a document written using the Binary Formats (doc, xls, ppt) can be translated to Office Open XML.

## Overview

The main goal of the Office Binary (doc, xls, ppt) Translator to Open XML project is to create software tools, plus guidance, showing how a document written using the Binary Formats (doc, xls, ppt) can be translated into the Office Open XML format (aka OpenXML). As a result customers can use these tools to migrate from the binary formats to OpenXML; thus, enabling them to more easily access their existing content in the new world of XML. The Translator will be available under the open source Berkeley Software Distribution (BSD) license, which allows that anyone can use the mapping and code, submit bugs and feedback, or contribute to the project.

On February 15th 2008, Microsoft has made it even easier to get access to the binary formats documentation. This documentation has been completely revamped by June 30, 2008 and the latest version can be downloaded [from here](https://msdn.microsoft.com/en-us/library/dd208104.aspx). All these specifications are available under the [Open Specification Promise](http://www.microsoft.com/interop/osp).

The Office Open XML file formats have been approved and published as ISO/IEC 29500. The pre-ISO version or ECMA-376 1st Edition, which is implemented in Office 2007 SP2, and ECMA-376 2nd edition, which is technically aligned with ISO/IEC 29500, are available free of charge from Ecma-International. ISO/IEC 29500 can be purchased from ISO/IEC.

Microsoft also published a set of document-format implementation notes for ECMA-376 1st Edition. The goal of publishing these notes is to help other implementers improve interoperability with Office, by transparently documenting the details of Microsoft's OpenXML implementation. To get to the ECMA-376 implementer notes, go to the DII home page and click on Reference and then select ECMA-376 1st Edition from the dropdown list.

While Microsoft provides with Office 2007 and the File Format Compatibiliy pack for earlier Office versions a migration path from binary Office formats to OpenXML, the Office Binary (doc, xls, ppt) Translator to Open XML project is still necessary due to the following reasons

* Enables the back-office / batch scenario due to its a command-line-based architecture
* Provides a cross-platform story via .Net/Mono, i.e. it the translators run, for example, on SUSE Linux
* Proves the usability and completeness of the file format specifications
* Allows that anyone uses the mapping, code snippets, etc. due to the open source development approach based on the liberate BSD license

We have chosen to use an Open Source development model that allows developers from all around the world to participate and contribute to the project.

## Main Contributors

### [Evolution](https://www.evolutionjobs.com/) (.NET Core Port)
Fork and port of .NET 2.0 Mono original to .NET Standard 2.0

### [DIaLOGIKa](http://www.dialogika.de/) (Analysis and Development)

DIaLOGIKa - a German systems and software house founded in 1982 - conducts projects on behalf of industry, finance, and governmental and supranational clients such as the institutions of the European Union (EU).

From the beginning DIaLOGIKa has focused – among others – on technically demanding projects in the field of multilingual text and data processing such as document format conversion. DIaLOGIKa has also contributed to the OpenXML/ODF Translator project.

### [Microsoft](http://www.microsoft.com/interop) (Architectural Guidance, Technical Support & Project Management)
