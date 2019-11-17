# ConvertToOFX
* Version: 3
* Website: http://www.norcalico.com/ConvertToOFX/
* Source Code: https://github.com/jtg-github/ConvertToOFX

ConvertToOFX is a Windows program that attempts to modify files in a way that allows them to be imported into Microsoft Money. Currently, it works with many variations of QFX files, which are known under several names: "Quicken file" or "Quicken Web Connect" or "Quicken Direct Connect".

It does this by editing the file in several ways. QFX files are a variation of OFX files. However, Microsoft Money stopped development around 2010 and does not support the latest formats. So the QFX files are modified to remove extra elements that would otherwise cause trouble for Microsoft Money.

After the file is automatically converted, you can manually inspect and edit the file if you want to. Then, you can either save the file or call the Microsoft Money Import Handler to import the file. The Import Handler can be called from this program, or by manually opening an OFX file with the Import Handler program.

# License and Dependencies
This program is free and open source. It is distributed under the MIT license. See the LICENSE file. 

It uses the TinyXML-2 library for XML parsing. TinyXML-2 can be found here: https://github.com/leethomason/tinyxml2


# Requirements
This program only runs on Windows. It was tested to work with Windows 7 and Windows 10.


# How to Use
1) Open a QFX file using this program. You can use your web browser's "Open with" feature to select this program when opening a QFX file from the internet. It will automatically display in the left window pane.

Alternatively, if you have a QFX file saved locally, click "File" and then click "Open..." and select the file you wish to open.

2) Convert the QFX to OFX by selecting "OFX Actions" from the menu and then "Convert". If it encountered any issues, it will display messages. A lot of issues can be ignored but are displayed just in case.

3) Inspect the output in the right window pane. This will be in XML. Most of the time you won't need to change anything, but if you see something wrong, go ahead and modify it.

4) Send to the Microsoft Money Import Handler by clicking "OFX Actions" from the menu and then "Send to Import Handler".


# Bugs
If you encounter any issues, you can create an issue on the GitHub project. You can also try contacting me on the website for this project.

I may not be able to fix all bugs quickly: I don't have a computer that supports Visual Studio 2019. I created this program by using free Azure credits. Also, the Microsoft Win32 API is a royal pain to deal with. But I would still like to hear if you encounter bugs.
