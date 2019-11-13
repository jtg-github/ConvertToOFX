// compile with: /D_UNICODE /DUNICODE /DWIN32 /D_WINDOWS /c

/******************************************************************************
* This program attempts to convert data into an OFX file that
* Microsoft Money can properly read. Currently, it only supports QFX files
* (also known as Quicken files or Quicken Web Connect files), which are already
* somewhat closely related to the OFX files that Money can read.
*
* At a high level, this program works by removing extra elements from the
* Transaction sections of the Quicken file. MS Money would otherwise choke on
* these elements.
*
* The Win32 API is ugly. This is also my first Win32 program. Hello World!
*
* This program depends on the TinyXML2 project for XML processing.
*
******************************************************************************/

#include "tinyxml2.h"

#include <cassert>
#include <ctype.h>
#include <map>
#include <regex>
#include <set>
#include <shobjidl.h> 
#include <sstream>
#include <stdlib.h>
#include <string>
#include <tchar.h>
#include <vector>
#include <windows.h>
#include <winhttp.h>
#pragma comment(lib, "winhttp.lib")

#define ID_FILE_OPEN 0
#define ID_FILE_EXIT 1
#define ID_ACTIONS_CONVERT_TO_OFX 2
#define ID_ACTIONS_SAVE_OFX 3
#define ID_ACTIONS_SEND_TO_MONEY 4
#define ID_HELP_ONLINE 5
#define ID_HELP_ABOUT 6
#define ID_HELP_PRIVACY_NOTICE 7
#define ID_CONFIG_CHANGE_IMPORT_HANDLER_LOCATION 8
#define ID_CONFIG_DEDUPE_MEMO 9

#define IDC_MAIN_EDIT 101
#define IDC_OFX_EDIT 102
#define IDC_BUTTON_OPEN 103
#define IDC_BUTTON_CONVERT_AND_IMPORT 104


// Global variables
std::wstring VERSION_ID = L"1";
static TCHAR szWindowClass[] = _T("ConvertToOFX");  // main window class name
static TCHAR szTitle[] = _T("ConvertToOFX");
const wchar_t INPUT_DEFAULT_TEXT[] =
L"Open a file to display text here.\r\n\r\n"
"Currently, only QFX files are supported.";
const wchar_t OFX_DEFAULT_TEXT[] =
L"OFX Output will appear here after input is converted.";
static std::wstring IMPORT_HANDLER_EXE =
L"C:\\Program Files (x86)\\Microsoft Money Plus\\MNYCoreFiles\\mnyimprt.exe";
const wchar_t privacyMessage[] =
L"GDPR Privacy Notice: This program does not collect any financial data. "
"The only data it collects is related to usage: We want to identify how "
"many people use this program. In order to do that, this program 'pings' a "
"webserver upon starting. This program does not send any financial data!";
// For the headers, spacing matters! I've observed some files not getting 
// accepted because of spacing before and after an '='! Very strange, but
// this is Windows so it shouldn't be surprising.
const std::string XML_HEADER =
"<?xml version=\"1.0\" encoding=\"utf-8\" ?>";
const std::string XML_OFX_HEADER =
"<?OFX OFXHEADER=\"200\" VERSION=\"202\" SECURITY=\"NONE\" "
"OLDFILEUID=\"NONE\" NEWFILEUID=\"NONE\" ?>";
// Allowed child elements under <STMTTRN> fields. Everything else gets deleted.
const std::set<std::string> STMTTRN_WHITELIST{ "TRNTYPE", "DTPOSTED", "TRNAMT",
    "FITID", "CHECKNUM", "NAME", "MEMO", "CCACCTTO", "DTUSER" };
// Where to find the <BANKTRANLIST> for a Message Set Type 
std::map<std::string, std::vector<std::string>> TYPE_TO_BANKTRANLIST_MAP = {
    {"CREDITCARDMSGSRSV1",
        {"OFX", "CREDITCARDMSGSRSV1", "CCSTMTTRNRS", "CCSTMTRS", "BANKTRANLIST"}},
    {"BANKMSGSRSV1",
        {"OFX", "BANKMSGSRSV1",       "STMTTRNRS",   "STMTRS",   "BANKTRANLIST"}},
};
bool dedupeMemoField = true;

// Attempt to fix the imbalanced input into well-formatted XML.
std::string FixXML(const std::string& input) {
    // A lot of banks really mangle their XML and this is the #1 problem
    // with preparing a file for MS Money. By now, we've detected XML mis-match
    // issues. Let's try to fix it by guessing where matching tags could go...
    // * Most commonly, it occurs when an element does not have a closing tag
    //   after a value. E.g. '<status><value>1</status>' is missing a 
    //   '</value>'.
    //
    // MS Money doesn't require perfect XML to work, but getting it perfect 
    // will help us. MS Money trips up on extra fields. Once we have valid XML,
    // we can use standard XML APIs to remove problematic fields. Determining 
    // which fields are problematic is an educated guess.

    std::string fixedXML;  // Will contain our final result.
    std::string tag;  // *Anything* in brackets: <.*> 
    std::string value;  // Anything not inside brackets
    bool processTag = false;
    std::vector<std::string> tagStack;

    // To determine where to add tags, we tokenize XML tags and values.
    // Upon encountering a closing bracket ('>'), our fixing logic kicks in.
    // We use a stack to determine if there is an imbalance, and if to fix.
    for (const char& c : input) {
        if (c == '>') {
            tag += c;

            if (tag.find("</") == 0) {
                // Closing tag (e.g. </item>). Make sure we match with the top
                // of the tag stack
                if (tagStack.size() == 0) {
                    // This XML is so bad we can't fix it.
                    // Give up. The user will get a message
                    // later when this XML fails to parse.
                    fixedXML = "No XML tag to match closing tag with. "
                        "Giving up. XML:\n" + fixedXML;
                    return fixedXML;
                }
                std::string tagRawValue = tag.substr(2, tag.length() - 3);
                std::string topOfStack = tagStack.back();
                tagStack.pop_back();
                std::string topOfStackRawValue = topOfStack.substr(
                    1,
                    topOfStack.length() - 2);
                while (topOfStackRawValue != tagRawValue) {
                    if (tagStack.size() == 0) {
                        // This XML is so bad we can't fix it.
                        // Give up. The user will get a message
                        // later when this XML fails to parse.
                        fixedXML = "Ran into issues trying to fix this XML:\n"
                            + fixedXML;
                        return fixedXML;
                    }
                    fixedXML += value + "</" + topOfStackRawValue + ">";
                    value = "";
                    topOfStack = tagStack.back();
                    topOfStackRawValue = topOfStack.substr(1,
                        topOfStack.length() - 2);
                    tagStack.pop_back();
                }
                fixedXML += value + tag;
                value = tag = "";
            }
            else if (tag.find("/>") == tag.length() - 2 || tag.find("<?")==0) {
                // Self contained tag. Write it out directly. Nothing to 
                // balance here.
                fixedXML += value + tag;
                value = tag = "";
            }
            else {
                // This is an opening tag, e.g. <tag>
                if (value.length() > 0) {
                    // We have a value. This value is associated with the item 
                    // at the top of the stack. Since we didn't encounter a 
                    // closing element, this must be imbalanced, and we 
                    // manually add the closing element after the value.
                    if (tagStack.size() == 0) {
                        // This XML is so bad we can't fix it.
                        // Give up. The user will get a message
                        // later when this XML fails to parse.
                        fixedXML =
                            "No XML tag to match value with. Giving up. XML:\n"
                            + fixedXML;
                        return fixedXML;
                    }
                    std::string topOfStack = tagStack.back();
                    tagStack.pop_back();
                    std::string prevTagRawValue = topOfStack.substr(1,
                        topOfStack.length() - 2);
                    fixedXML += value + "</" + prevTagRawValue + ">" + tag;
                }
                else {
                    // No value already present, so we just write out the 
                    // element.
                    fixedXML += tag;
                }
                tagStack.push_back(tag);
                value = tag = "";
            }

            // We've processed a tag. If we encounter a character later, it 
            // belongs to a value and not an XML tag.
            processTag = false;
        }
        else if (c == '<') {
            // We are going to process an XML element
            assert(tag.length() == 0);
            tag += c;
            processTag = true;
            // If we have a value waiting, let's trim extra whitespace off
            if (value.length() > 1) {
                value.erase(value.find_last_not_of(" ") + 1);
            }
        }
        else if (processTag) {
            // Continue processing this character as an XML element
            tag += c;
        }
        else {
            // We are processing a value, not an XML element
            if (value.length() == 0 && (isspace(c))) {
                // Ignore leading (left-side) whitespace
                continue;
            }
            if (c == '\r' || c == '\n') {
                // Ignore new lines in values
                continue;
            }
            value += c;
        }
    }

    // If there is anything in the tagStack, close it out
    for (int i = tagStack.size() - 1; i >= 0; --i) {
        std::string topOfStack = tagStack.back();
        tagStack.pop_back();
        std::string tagRawValue = topOfStack.substr(1,
            topOfStack.length() - 2);
        fixedXML += "</" + tagRawValue + ">";
    }
    return fixedXML;
}

// Remove any STMTTRN child elements that are not whitelisted. 
void PruneSTMTTRN(tinyxml2::XMLElement* banktranlist) {
    // We need to prune extra elements because they can cause MS Money to 
    // reject the file. This increases our chances of success.

    // The <STMTTRN> is where all the magic and trouble happens.
    // Extra elements under it will choke MS Money.
    // Let's remove any extra elements that don't match a whitelist.
    /*
        Here's an example of a sanitized STMTTRN:
        <STMTTRN>
          <TRNTYPE>CHECK</TRNTYPE>
          <DTPOSTED>20190101120000.000[0:GMT]</DTPOSTED>
          <TRNAMT>-1.00</TRNAMT>
          <FITID>0123456789ABCDEF</FITID>
          <CHECKNUM>100</CHECKNUM>
          <NAME>CHECK# 100 CHECK WITHDRAWAL</NAME>
          <MEMO>CHECK# 100 CHECK WITHDRAWAL</MEMO>
        </STMTTRN>
    */
    // * Other possibly accepted child elements: <DTUSER>, <CCACCTTO>
    // * TRNTYPE has been observed to be CHECK, DEBIT, CREDIT. Any others?
    // * I've observed STMTRN elements as children under the following:
    //   * <OFX><CREDITCARDMSGSRSV1><CCSTMTTRNRS><CCSTMTRS><BANKTRANLIST>
    //   * <OFX><BANKMSGSRSV1><STMTTRNRS><STMTRS><BANKTRANLIST>
    //   * Are there others that I should care about?
    tinyxml2::XMLElement* stmttrn = banktranlist->FirstChildElement("STMTTRN");
    while (stmttrn) {  // For every STMTTRN
        tinyxml2::XMLNode* child = stmttrn->FirstChild();
        while (child) {  // For every child element under STMTTRN
            tinyxml2::XMLNode* currentChild = child;
            child = child->NextSibling();
            if (STMTTRN_WHITELIST.find(currentChild->Value()) ==
                STMTTRN_WHITELIST.end()) {
                stmttrn->DeleteChild(currentChild);
            }
        }

        if (dedupeMemoField) {
            // If <NAME> == <MEMO>, then delete MEMO field.
            // Personally, I hate when this gets duplicated. Waste of space!
            tinyxml2::XMLElement* name = stmttrn->FirstChildElement("NAME");
            tinyxml2::XMLElement* memo = stmttrn->FirstChildElement("MEMO");
            if (name && memo && name->GetText() && memo->GetText() && 
                (strcmp(name->GetText(), memo->GetText()) == 0)) {
                stmttrn->DeleteChild(memo);
            }
        }

        stmttrn = stmttrn->NextSiblingElement("STMTTRN");
    }
}

// Is the XML balanced correctly with proper opening and closing tags?
bool isXMLBalanced(const std::string& xml) {
    // TinyXML2 does not always appear to be correct when determining if
    // the XML is valid and balanced. Thus, this is used to make that 
    // determination. 
    // We use a stack to determine if the XML is balanced. It's very
    // similar to the previous FixXML() method.
    // Since our XML input files don't have attribute values, the code is a 
    // little simpler. If the XML ever gets attributes, then 
    // expect this code to break.
    bool isBalanced = true;
    std::string tag;  // *Anything* in brackets: <.*> 
    std::string value;  // Anything not inside brackets
    bool processTag = false;
    std::vector<std::string> tagStack;

    for (const char& c : xml) {
        if (c == '>') {
            tag += c;
            if (tag.find("</") == 0) {
                // Closing tag (e.g. </item>)
                if (tagStack.size() == 0) {
                    // No opening tag to match this closing tag. Unbalanced.
                    return false;
                }
                std::string tagRawValue = tag.substr(2, tag.length() - 3);
                std::string topOfStack = tagStack.back();
                tagStack.pop_back();
                std::string topOfStackRawValue = topOfStack.substr(
                    1,
                    topOfStack.length() - 2);
                if (topOfStackRawValue != tagRawValue) {
                    // The top of the stack does not match this closing tag. 
                    // Therefore, XML is unbalanced.
                    return false;
                }
                value = tag = "";  // tags match. close out tag and value vars.
            }
            else if (tag.find("/>") == tag.length() - 2 || tag.find("<?")==0) {
                // Self contained tag. Nothing to balance here.
                value = tag = "";
                processTag = false;
                continue;
            }
            else {
                // This is the end of an opening tag, e.g. <tag>
                if (value.length() > 0) {
                    // We have a previous value that wasn't closed out with a 
                    // closing tag. Therefore, the xml is unbalanced.
                    return false;
                }
                tagStack.push_back(tag);
                value = tag = "";
            }

            // We've processed a tag. If we encounter a character later, it 
            // belongs to a value and not an XML tag.
            processTag = false;
        }
        else if (c == '<') {
            // We are going to process an XML element
            assert(tag.length() == 0);
            tag += c;
            processTag = true;
        }
        else if (processTag) {
            // Continue processing this character as an XML element
            tag += c;
        }
        else {
            if (value.length() == 0 && (isspace(c))) {
                // Ignore leading (left-side) whitespace
                continue;
            }
            if (c == '\r' || c == '\n') {
                // Ignore new lines in values
                continue;
            }
            value += c;
        }
    }

    if (tagStack.size() > 0) {
        return false;
    }
    return true;
}

void SetOfxWindowDebugText(HWND hWnd, const std::string& xml) {
    HWND hOfxEdit = GetDlgItem(hWnd, IDC_OFX_EDIT);
    std::string text =
        "This text is invalid and only for debugging purposes!"
        "\r\n\r\n" + xml;
    SetWindowTextA(hOfxEdit, text.c_str());
}

// Convert whatever is in the Input window (should be QFX XML) to 
// a MS Money-acceptable OFX format.
bool ConvertInputToOFX(HWND hWnd) {
    // Polish the input text before converting into OFX-happy XML.
    std::string polishedText;

    // First, get the input from the main window and shove into a ANSI string
    HWND hEdit = GetDlgItem(hWnd, IDC_MAIN_EDIT);
    std::string inputText;
    int len = GetWindowTextLength(hEdit) + 1;
    std::vector<wchar_t> buf(len);
    GetWindowText(hEdit, &buf[0], len);
    std::wstring wide = &buf[0];
    std::string s(wide.begin(), wide.end());

    // Remove anything before <OFX> - it's junk to Money or headers that can
    // be replaced. Some banks, i.e. Wells Fargo, jam everything into one line.
    // std::regex's search and match have memory issues with long lines, so
    // that's why this logic is very simple - to avoid using that library.
    std::istringstream ss(s);
    std::string line;
    // First, look for the start of <OFX>
    while (std::getline(ss, line)) {
        std::size_t start = line.find("<OFX>");
        if (start != std::string::npos) {
            polishedText += line.substr(start, std::string::npos);
            break;
        }
    }
    // Now append the rest
    while (std::getline(ss, line)) {
        polishedText += line + "\n";
    }

    // Add OFX XML-style headers. Version may be wrong, but Money doesn't care.
    polishedText = XML_HEADER + "\n" + XML_OFX_HEADER + "\n" + polishedText;

    // Convert text to an XML object
    tinyxml2::XMLDocument doc;
    const char* inxml = polishedText.c_str();
    doc.Parse(inxml);

    // Check if the elements are balanced. TinyXML is not always 100% correct
    // (sometimes give success / code 0 for broken XML. So we also check again.
    if (doc.ErrorID() == 14 ||
        (doc.ErrorID() == 0 && !isXMLBalanced(polishedText))) {
        // Mis-matched brackets. Let's attempt to fix it, 
        // but inform the user first.
        std::string msg;
        if (doc.ErrorID() == 14) {
            msg = "Input XML does not have matching brackets according to the "
                "XML Parser. Will attempt to fix it!\n\n"
                "XML Parser Error Message: ";
            msg += doc.ErrorStr();
        }
        else {
            msg = "Input XML appears to be unbalanced. Will try to fix it!";
        }
        MessageBoxA(NULL,
            msg.c_str(),
            "FYI: XML is unbalanced",
            MB_OK | MB_ICONWARNING);
        std::string fixedXML = FixXML(polishedText);
        inxml = fixedXML.c_str();
        doc.Parse(inxml);
        if (doc.ErrorID() != 0) {
            std::string msg = "Could not fix the XML. This XML either needs "
                "to be fixed at the source, or this program needs extra "
                "modifications to handle the XML.\n\nParser error message: ";
            msg += std::to_string(doc.ErrorID()) + ":\n";
            msg += doc.ErrorStr();
            MessageBoxA(NULL,
                msg.c_str(),
                "Error Parsing XML Again",
                MB_OK | MB_ICONSTOP);
            SetOfxWindowDebugText(hWnd, inxml);
            return false;
        }
    }
    else if (doc.ErrorID() != 0) {
        std::string msg =
            "Parser Error when processing XML. Cannot continue: ";
        msg += std::to_string(doc.ErrorID()) + ":\n";
        msg += doc.ErrorStr();
        MessageBoxA(NULL,
            msg.c_str(),
            "Error Parsing XML",
            MB_OK | MB_ICONERROR);
        SetOfxWindowDebugText(hWnd, inxml);
        return false;
    }

    // Determine if this is a CreditCard statement or a Bank statement.
    tinyxml2::XMLElement* ofxRoot = doc.FirstChildElement("OFX");
    if (!ofxRoot) {
        MessageBoxA(NULL,
            "OFX is missing <OFX> element at the root. Cannot parse.",
            "Error Parsing XML",
            MB_OK);
        SetOfxWindowDebugText(hWnd, inxml);
        return false;
    }
    tinyxml2::XMLHandle docHandle(&doc);

    // Determine what "types" this statement contains. Some include both
    // Credit Card and Bank statements, or have both with 1 empty. It's crazy.
    // Currently, only CreditCard and Bank statements are supported,
    // since they are the only 2 types which I see have a BANKTRANLIST child.
    // There are other types, but I couldn't see if they have BANKTRANLIST
    // children. For a list of other types, see: 
    // https://schemas.liquid-technologies.com/OFX/2.1.1/?page=ofxresponse.html
    std::vector<std::string> statementTypes = { };
    if (docHandle.FirstChildElement("OFX").
        FirstChildElement("CREDITCARDMSGSRSV1").ToElement()) {
        // We have a Credit Card Statement
        statementTypes.push_back("CREDITCARDMSGSRSV1");
    }
    if (docHandle.FirstChildElement("OFX").
        FirstChildElement("BANKMSGSRSV1").ToElement()) {
        // We have a Bank Statement
        statementTypes.push_back("BANKMSGSRSV1");
    }

    if (statementTypes.size() == 0) {
        // Error: Found zero statements
        MessageBoxA(NULL,
            "OFX is missing valid elements under the <OFX> root (like "
            "<CREDITCARDMSGSRSV1> or <BANKMSGSRSV1>). Cannot parse.",
            "Error Parsing XML",
            MB_OK);
        SetOfxWindowDebugText(hWnd, inxml);
        return false;
    }

    // Using the Statement Type as a key name, we lookup its expected 
    // <BANKTRANLIST> path from a map data structure.
    for (std::string& type : statementTypes) {
        std::vector<std::string> pathToBanktranlist =
            TYPE_TO_BANKTRANLIST_MAP[type];
        if (pathToBanktranlist.size() == 0) {
            MessageBoxA(NULL,
                "Cannot find TYPE_TO_BANKTRANLIST_MAP mapping; Code error! "
                "Stopping!",
                "Fatal Error",
                MB_OK | MB_ICONSTOP);
            SetOfxWindowDebugText(hWnd, inxml);
            continue;
        }

        // Now that we have an expected <BANKTRANLIST> path, see if the item
        // exists at that path.
        tinyxml2::XMLHandle banktranlistHandle(&doc);
        bool hasError = false;
        for (size_t i = 0; i < pathToBanktranlist.size(); ++i) {
            banktranlistHandle = banktranlistHandle.FirstChildElement(
                pathToBanktranlist[i].c_str());
            if (!banktranlistHandle.ToElement()) {
                // Possible Problem: Could not find element in expected path...
                // Let user know right now and stop processing for this type.
                hasError = true;
                std::string fullPath;
                for (size_t j = 0; j < pathToBanktranlist.size(); ++j) {
                    fullPath += "<" + pathToBanktranlist[j] + ">";
                }
                std::string errorMsg = "Not modifiying " + type + " because "
                    "we encountered problems locating this element: " +
                    pathToBanktranlist[i] + " in the path " + fullPath +
                    ". We were expecting it to be present. This might be a "
                    "problem (or not, if it was purposely left out): inspect "
                    "the output to make sure you are okay with results.";
                MessageBoxA(NULL,
                    errorMsg.c_str(),
                    "FYI: Possible Error",
                    MB_OK | MB_ICONINFORMATION);
                break;
            }
        }

        // We got the <BANKTRANLIST> element. Now, prune unnecessary elements.
        if (!hasError) {
            tinyxml2::XMLElement* banktranlist =
                banktranlistHandle.ToElement();
            PruneSTMTTRN(banktranlist);
        }
    }

    // Pretty Print XML
    tinyxml2::XMLPrinter printer;
    doc.Print(&printer);
    const char* prettyXml = printer.CStr();

    // We want the editor to display text nicely. In Windows land,
    // we need to add carriage returns before the newlines...
    std::string normalizedString;
    const char* p = &prettyXml[0];
    while (*p != '\0') {
        if (*p == '\n') {
            normalizedString += "\r\n";
        }
        else {
            normalizedString += *p;
        }
        ++p;
    }
    HWND hOfxEdit = GetDlgItem(hWnd, IDC_OFX_EDIT);

    if (s == normalizedString) {
        MessageBoxA(NULL,
            "FYI: Nothing changed after attempting to convert!",
            "FYI",
            MB_OK | MB_ICONINFORMATION);
    }
    SetWindowTextA(hOfxEdit, normalizedString.c_str());
    return true;
}

// Create the Menu Bar
void CreateMainMenu(HWND hWnd) {
    HMENU hMenu = CreateMenu();
    HMENU hFileSubMenu = CreatePopupMenu();
    HMENU hActionsSubMenu = CreatePopupMenu();
    HMENU hConfigSubMenu = CreatePopupMenu();
    HMENU hHelpSubMenu = CreatePopupMenu();
    
    AppendMenu(hMenu, MF_STRING | MF_POPUP,
        (UINT)hFileSubMenu, _T("&File"));
    AppendMenu(hMenu, MF_STRING | MF_POPUP,
        (UINT)hActionsSubMenu, _T("OFX &Actions"));
    AppendMenu(hMenu, MF_STRING | MF_POPUP,
        (UINT)hConfigSubMenu, _T("Confi&g"));
    AppendMenu(hMenu, MF_STRING | MF_POPUP,
        (UINT)hHelpSubMenu, _T("&Help"));

    AppendMenu(hFileSubMenu, MF_STRING, ID_FILE_OPEN,
        _T("&Open File...\tALT+O"));
    AppendMenu(hFileSubMenu, MF_STRING, ID_FILE_EXIT,
        _T("E&xit"));

    AppendMenu(hActionsSubMenu, MF_STRING, ID_ACTIONS_CONVERT_TO_OFX,
        _T("&Convert To OFX\tALT+C"));
    AppendMenu(hActionsSubMenu, MF_STRING, ID_ACTIONS_SAVE_OFX,
        _T("&Save OFX As...\tALT+S"));
    AppendMenu(hActionsSubMenu, MF_STRING, ID_ACTIONS_SEND_TO_MONEY,
        _T("Send OFX To Money &Import Handler\tALT+I"));

    AppendMenu(hConfigSubMenu,
        MF_STRING,
        ID_CONFIG_CHANGE_IMPORT_HANDLER_LOCATION,
        _T("&Change Money Import Handler Location"));
    AppendMenu(hConfigSubMenu,
        MF_STRING,
        ID_CONFIG_DEDUPE_MEMO,
        _T("&Delete the MEMO field if identical to NAME field"));

    AppendMenu(hHelpSubMenu, MF_STRING, ID_HELP_ONLINE,
        _T("On-Line &Documentation"));
    AppendMenu(hHelpSubMenu, MF_STRING, ID_HELP_PRIVACY_NOTICE,
        _T("&Privacy Notice (GDPR)"));
    AppendMenu(hHelpSubMenu, MF_STRING, ID_HELP_ABOUT,
        _T("&About"));

    if (dedupeMemoField) {
        CheckMenuItem(hConfigSubMenu, ID_CONFIG_DEDUPE_MEMO, MF_CHECKED);
    }
    SetMenu(hWnd, hMenu);
}

// After a user selects a file, load the contents into the input Text Box
void LoadFile(const PWSTR filename, HWND hWnd) {
    HANDLE hFile = CreateFile(filename, GENERIC_READ, 0, NULL, OPEN_ALWAYS,
        FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile != INVALID_HANDLE_VALUE) {
        DWORD dwFileSize = GetFileSize(hFile, NULL);
        if (dwFileSize != 0xFFFFFFFF) {
            DWORD dwRead;
            LPBYTE fileByteArray = new BYTE[dwFileSize];

            if (ReadFile(hFile, fileByteArray, dwFileSize, &dwRead, NULL)) {
                // If the file is UTF-8 or Unicode, we need to convert it to
                // Wide Characters. I can't get it working with Unicode yet.
                std::string fileTxt(reinterpret_cast<const char*>(
                    fileByteArray),
                    dwRead);
                int num_chars = MultiByteToWideChar(CP_UTF8,
                    MB_ERR_INVALID_CHARS, fileTxt.c_str(), fileTxt.length(),
                    NULL, 0);
                std::wstring fileTxtWideChars;
                HWND hEdit = GetDlgItem(hWnd, IDC_MAIN_EDIT);
                if (num_chars) {
                    fileTxtWideChars.resize(num_chars);
                    if (MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS,
                        fileTxt.c_str(), fileTxt.length(),
                        &fileTxtWideChars[0], num_chars)) {
                        SetWindowText(hEdit, fileTxtWideChars.c_str());
                    }
                    else {
                        MessageBox(NULL,
                            L"Error reading characters in file. "
                            "Maybe there is an invalid character?",
                            L"Error",
                            MB_OK | MB_ICONERROR);
                    }
                }
                else {
                    if (!SetWindowTextA(hEdit, fileTxt.c_str())) {
                        MessageBox(NULL, L"Error displaying text as ANSI.",
                            L"Error", MB_OK | MB_ICONERROR);
                    }
                }
            }
            delete[] fileByteArray;
        }
        CloseHandle(hFile);
    }
    else {
        // User probably selected Cancel
        MessageBox(NULL,
            L"No File Selected.",
            L"Warning: Nothing Selected",
            MB_OK | MB_ICONWARNING);
    }
}

// Write out the contents of the OFX window to disk
void WriteOutFile(const PWSTR filename, HWND hWnd) {
    HANDLE hFile = CreateFile(filename, GENERIC_WRITE, FILE_SHARE_READ, NULL,
        OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    HWND hOfx = GetDlgItem(hWnd, IDC_OFX_EDIT);

    std::string inputText;
    int len = GetWindowTextLength(hOfx) + 1;

    std::vector<char> buf(len);
    GetWindowTextA(hOfx, &buf[0], len);
    std::string s = &buf[0];
    DWORD bytesWritten;
    WriteFile(hFile, s.c_str(), len, &bytesWritten, NULL);
    CloseHandle(hFile);
}

// Display the Save File Dialog and return chosen new file name.
PWSTR SaveFileWindow() {
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED |
        COINIT_DISABLE_OLE1DDE);
    wchar_t tmp[1] = L"";
    PWSTR pszFilePath = tmp;
    if (SUCCEEDED(hr)) {
        IFileSaveDialog* pFileSave;

        // Create the FileOpenDialog object.
        hr = CoCreateInstance(CLSID_FileSaveDialog, NULL, CLSCTX_ALL,
            IID_IFileSaveDialog, reinterpret_cast<void**>(&pFileSave));
        if (SUCCEEDED(hr)) {

            const COMDLG_FILTERSPEC c_rgSaveTypes[] =
            { {L"OFX Files (*.ofx)",       L"*.ofx"}, };
            hr = pFileSave->SetFileTypes(ARRAYSIZE(c_rgSaveTypes),
                c_rgSaveTypes);
            hr = pFileSave->SetFileTypeIndex(1);  // Uses 1-based index. groan.
            hr = pFileSave->SetDefaultExtension(L"ofx");
            // Show the Open dialog box.
            hr = pFileSave->Show(NULL);
            // Get the file name from the dialog box.
            if (SUCCEEDED(hr)) {
                IShellItem* pItem;
                hr = pFileSave->GetResult(&pItem);
                if (SUCCEEDED(hr)) {
                    hr = pItem->GetDisplayName(SIGDN_FILESYSPATH,
                        &pszFilePath);

                    // Display the file name to the user.
                    if (SUCCEEDED(hr)) {
                        pFileSave->Release();
                        return pszFilePath;
                        //CoTaskMemFree(pszFilePath);
                    }
                    pItem->Release();
                }
            }
            pFileSave->Release();
        }
        CoUninitialize();
    }
    return pszFilePath;
}

// Prompt the user to select a file to open and return that file name.
PWSTR OpenFileWindow() {
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED |
        COINIT_DISABLE_OLE1DDE);
    wchar_t tmp[1] = L"";
    PWSTR pszFilePath = tmp;
    if (SUCCEEDED(hr)) {
        IFileOpenDialog* pFileOpen;

        // Create the FileOpenDialog object.
        hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL,
            IID_IFileOpenDialog, reinterpret_cast<void**>(&pFileOpen));

        if (SUCCEEDED(hr)) {
            // Show the Open dialog box.
            hr = pFileOpen->Show(NULL);

            // Get the file name from the dialog box.
            if (SUCCEEDED(hr)) {
                IShellItem* pItem;
                hr = pFileOpen->GetResult(&pItem);
                if (SUCCEEDED(hr)) {
                    hr = pItem->GetDisplayName(SIGDN_FILESYSPATH,
                        &pszFilePath);
                    // Display the file name to the user.
                    if (SUCCEEDED(hr)) {
                        return pszFilePath;
                    }
                    pItem->Release();
                }
            }
            pFileOpen->Release();
        }
        CoUninitialize();
    }
    return pszFilePath;
}

// Send the OFX output to the MS Money Import Handler
void SendToMoneyImportHandler(HWND hWnd) {
    // Save the OFX as a temporary file and then call msmnyimprt.exe with
    // that file as a parameter.
    // The MS Money Import Handler appears to be very simple, and we could
    // probably replicate the code here. It appears to do the following:
    // 1) Create a temporary file with the OFX
    // 2) Update Registry keys that point to that file
    // 3) Prompt the user to start Money. Then Money processes the file and
    //    deletes it and clears the registry key.
    // Since mnyimprt.exe just works, we'll avoid doing that for now.

    // First get the OFX text. Make sure it is not empty.
    HWND hOfx = GetDlgItem(hWnd, IDC_OFX_EDIT);
    std::string inputText;
    int len = GetWindowTextLength(hOfx) + 1;
    if (len == 1) {
        MessageBox(hWnd,
            _T("OFX Text is empty. Nothing to Import!"),
            _T("Error"),
            MB_OK | MB_ICONERROR);
        return;
    }
    std::vector<wchar_t> buf(len);
    GetWindowText(hOfx, &buf[0], len);
    std::wstring wide = &buf[0];
    if (std::equal(buf.begin(), buf.end(), std::begin(OFX_DEFAULT_TEXT))) {
        MessageBox(hWnd,
            _T("OFX Text is not valid. "
                "The right text pane needs to be updated!"),
            _T("Error"),
            MB_OK | MB_ICONERROR);
        return;
    }
    std::string ofxText(wide.begin(), wide.end());

    // The OFX needs to be in a temporary file for the MS Money Import Handler
    // to process it.
    DWORD dwRetVal = 0;
    TCHAR tmpFileName[MAX_PATH];
    TCHAR tmpFilePath[MAX_PATH];
    // Get the Temorary File Directory
    dwRetVal = GetTempPath(MAX_PATH, tmpFilePath);
    if (dwRetVal > MAX_PATH || (dwRetVal == 0)) {
        MessageBox(hWnd,
            _T("Error Getting Temporary File Path. Alternatively, you should "
                "save the OFX data and open that file with the Money "
                "Import Handler."),
            _T("Error"),
            MB_OK | MB_ICONERROR);
        return;
    }
    // Get a Temporary File Name inside the temp directory
    dwRetVal = GetTempFileName(tmpFilePath,  // directory for tmp files
        _T("ofx"),  // temp file name prefix, only first 3 letters used! 
        0,
        tmpFileName);
    if (dwRetVal == 0) {
        MessageBox(hWnd,
            _T("Unable to get Temporary File Name. Alternatively, you should "
                "save the OFX data and open that file with the Money Import "
                "Handler"),
            _T("Error"),
            MB_OK | MB_ICONERROR);
        return;
    }
    // Write out temporary file
    WriteOutFile(tmpFileName, hWnd);

    // Call the Import Handler with the file. The Import Handler, under the 
    // hood, appears to create another copy of the file and update two
    // Registry keys. When Money starts, it processes these keys and 
    // subsequently the file(s) before deleting them.
    SHELLEXECUTEINFO ShExecInfo;
    ShExecInfo.cbSize = sizeof(SHELLEXECUTEINFO);
    ShExecInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
    ShExecInfo.hwnd = hWnd;
    ShExecInfo.lpVerb = L"open";
    ShExecInfo.lpFile = IMPORT_HANDLER_EXE.c_str();
    ShExecInfo.lpParameters = tmpFileName;
    ShExecInfo.lpDirectory = NULL;
    ShExecInfo.nShow = SW_SHOWNORMAL;
    ShExecInfo.hInstApp = NULL;
    ShellExecuteEx(&ShExecInfo);
    // We want to wait before proceeding
    WaitForSingleObject(ShExecInfo.hProcess, INFINITE);
    CloseHandle(ShExecInfo.hProcess);

    int retVal = (int)ShExecInfo.hInstApp;
    if (retVal == 2) {
        // Import Handler EXE was not found
        std::wstring errorMsg = L"Error: Do you need to change the location "
            "of the Money Import Handler? Could not locate the "
            "MS Money Import Handler at: ";
        errorMsg += IMPORT_HANDLER_EXE;
        MessageBox(hWnd, errorMsg.c_str(), _T("Error"), MB_OK | MB_ICONERROR);
    }

    // Finally, delete the temporary file if one was created
    if (!DeleteFile(tmpFileName)) {
        std::wstring warnMsg =
            L"Warning: Could not delete the temporary file. "
            "You may want to delete the file manually. File was created at: ";
        warnMsg += tmpFileName;
        MessageBox(hWnd,
            warnMsg.c_str(),
            _T("Warning: Did Not Remove Temp File"),
            MB_OK | MB_ICONWARNING);
    }

    return;
}

// For those special people out there that have mnyimprt.exe elsewhere
// This change does not persist, so one improvement is to persist the value.
void ChangeImportHandlerLocation(HWND hWnd) {
    std::string msg = "Select a different location for mnyimprt.exe. You will "
        "need to do this every time you use this program, as the new location "
        "is NOT saved. I recommend you manually create the folder structure "
        "and copy mnyimprt.exe to: "
        "C:\\Program Files(x86)\\Microsoft Money Plus\\MNYCoreFiles\\mnyimprt.exe";
    MessageBoxA(hWnd,
        msg.c_str(),
        "FYI",
        MB_OK | MB_ICONINFORMATION);
    PWSTR NEW_IMPORT_HANDLER_EXE = OpenFileWindow();
    if (NEW_IMPORT_HANDLER_EXE[0] != L'\0') {  // Ensure user didn't hit Cancel
        IMPORT_HANDLER_EXE = NEW_IMPORT_HANDLER_EXE;
    }
}

// Main Window callback
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    PAINTSTRUCT ps;
    HDC hdc;
    HWND hEdit;
    HWND hwndButton;
    RECT rcClient;
    PWSTR filename;
    int buttonWidth = 75;
    int buttonHeight = 28;

    switch (message) {

    case WM_CREATE: {
        GetClientRect(hWnd, &rcClient);

        // Input / Source File window
        hEdit = CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", INPUT_DEFAULT_TEXT,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL | ES_MULTILINE |
            ES_AUTOVSCROLL | ES_AUTOHSCROLL,
            0, 0, rcClient.right * 49 / 100,
            rcClient.bottom - (5 + 28 + 5),
            hWnd,
            (HMENU)IDC_MAIN_EDIT,
            GetModuleHandle(NULL),
            NULL);
        if (hEdit == NULL)
            MessageBox(hWnd,
                _T("Could not create source edit box."),
                _T("Error"),
                MB_OK | MB_ICONERROR);
        SendMessage(hEdit,
            WM_SETFONT,
            (WPARAM)GetStockObject(DEFAULT_GUI_FONT),
            MAKELPARAM(FALSE, 0));
        // Set the TextBox limit to 1 MB of TCHARs
        SendMessage(hEdit, EM_SETLIMITTEXT, (WPARAM)1000000, 0);
        SetFocus(hEdit);

        // Converted OFX Text window
        hEdit = CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", OFX_DEFAULT_TEXT,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL | ES_MULTILINE |
            ES_AUTOVSCROLL | ES_AUTOHSCROLL,
            (int)rcClient.right * 51 / 100,
            0,
            (int)rcClient.right * 49 / 100, rcClient.bottom - (5 + 28 + 5),
            hWnd,
            (HMENU)IDC_OFX_EDIT,
            GetModuleHandle(NULL),
            NULL);
        if (hEdit == NULL) {
            MessageBox(hWnd,
                _T("Could not create ofx edit box."),
                _T("Error"),
                MB_OK | MB_ICONERROR);
        }
        SendMessage(hEdit,
            WM_SETFONT,
            (WPARAM)GetStockObject(DEFAULT_GUI_FONT),
            MAKELPARAM(FALSE, 0));
        // Set the TextBox limit to 1 MB of TCHARs
        SendMessage(hEdit, EM_SETLIMITTEXT, (WPARAM)1000000, 0);

        hwndButton = CreateWindow(
            L"BUTTON",
            L"Convert and Import!",  // Button text 
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON | BS_CENTER,
            rcClient.right - 20 - buttonWidth * 2,  // x position 
            rcClient.bottom - (buttonHeight + 5),  // y position 
            buttonWidth * 2,  // Button width
            buttonHeight,  // Button height
            hWnd,  // Parent window
            (HMENU)IDC_BUTTON_CONVERT_AND_IMPORT,
            GetModuleHandle(NULL),
            NULL);

        hwndButton = CreateWindow(
            L"BUTTON",
            L"Open...",  // Button text 
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON | BS_CENTER,
            20,         // x position 
            rcClient.bottom - (buttonHeight + 5),         // y position 
            buttonWidth,        // Button width
            buttonHeight,        // Button height
            hWnd,     // Parent window
            (HMENU)IDC_BUTTON_OPEN,
            GetModuleHandle(NULL),
            NULL);
        break;
    }
    case WM_SIZE: {
        // Resize the UI Elements when the window gets resized
        GetClientRect(hWnd, &rcClient);

        // The Input / Source window is the left 40% of the screen.
        hEdit = GetDlgItem(hWnd, IDC_MAIN_EDIT);
        SetWindowPos(hEdit, NULL,
            0,  // x starting position
            0,  // y starting position
            rcClient.right * 49 / 100,  // width 
            rcClient.bottom - (5 + buttonHeight + 5),  // height
            SWP_NOZORDER);

        // The OFX window gets the right 40% of the screen.
        hEdit = GetDlgItem(hWnd, IDC_OFX_EDIT);
        SetWindowPos(hEdit, NULL,
            rcClient.right * 51 / 100,  // x start
            0,  // y start
            rcClient.right * 49 / 100,  // width
            rcClient.bottom - (5 + buttonHeight + 5),  // height
            SWP_NOZORDER);

        // Bottom buttons
        hEdit = GetDlgItem(hWnd, IDC_BUTTON_CONVERT_AND_IMPORT);
        SetWindowPos(hEdit, NULL,
            rcClient.right - 20 - buttonWidth * 2,   // x start
            rcClient.bottom - (buttonHeight + 5),  // y start
            buttonWidth * 2,   // width
            buttonHeight,  // height
            SWP_NOZORDER);

        hEdit = GetDlgItem(hWnd, IDC_BUTTON_OPEN);
        SetWindowPos(hEdit, NULL,
            20,  // x start
            rcClient.bottom - (buttonHeight + 5),  // y start
            buttonWidth,   // width
            buttonHeight,  // height
            SWP_NOZORDER);
        break;
    }
    case WM_PAINT: {
        hdc = BeginPaint(hWnd, &ps);
        EndPaint(hWnd, &ps);
        break;
    }
    case WM_DESTROY: {
        PostQuitMessage(0);
        break;
    }
    case WM_COMMAND: {
        // A Menu Item was selected
        switch (LOWORD(wParam)) {
        case ID_FILE_OPEN:
        case IDC_BUTTON_OPEN: {
            filename = OpenFileWindow();
            if (filename != L"") {
                LoadFile(filename, hWnd);
            }
            break;
        }
        case IDC_BUTTON_CONVERT_AND_IMPORT: {
            // Combine two steps into one button
            bool result = ConvertInputToOFX(hWnd);
            if (result) {
                // Only send to Money if we parse successfully!
                SendToMoneyImportHandler(hWnd);
            }
            break;
        }
        case ID_FILE_EXIT: {
            PostQuitMessage(0);
            break;
        }
        case ID_ACTIONS_CONVERT_TO_OFX: {
            ConvertInputToOFX(hWnd);
            break;
        }
        case ID_ACTIONS_SAVE_OFX: {
            PWSTR filename = SaveFileWindow();
            WriteOutFile(filename, hWnd);
            break;
        }
        case ID_ACTIONS_SEND_TO_MONEY: {
            SendToMoneyImportHandler(hWnd);
            break;
        }
        case ID_HELP_ABOUT: {
            std::wstring aboutUrl =
                L"http://www.norcalico.com/ConvertToOFX/about/" +
                VERSION_ID + L".html";
            std::wstring msg =
                L"ConvertToOFX Version: " + VERSION_ID + L"\n\n"
                L"This program uses the TinyXML-2 project (zlib License) \n\n"
                L"A web browser will now open to show more information: " +
                aboutUrl;
            MessageBox(hWnd, msg.c_str(), _T("About"), MB_OK);
            ShellExecute(NULL,
                L"open",
                aboutUrl.c_str(),
                NULL,
                NULL,
                SW_SHOWNORMAL);
            break;
        }
        case ID_HELP_ONLINE: {
            ShellExecute(NULL,
                L"open",
                L"http://www.norcalico.com/ConvertToOFX/",
                NULL,
                NULL,
                SW_SHOWNORMAL);
            break;
        }
        case ID_HELP_PRIVACY_NOTICE: {
            MessageBox(hWnd, privacyMessage, _T("Privacy Notice"), MB_OK);
            break;
        }
        case ID_CONFIG_CHANGE_IMPORT_HANDLER_LOCATION: {
            ChangeImportHandlerLocation(hWnd);
            break;
        }
        case ID_CONFIG_DEDUPE_MEMO: {
            HMENU mainMenu = GetMenu(hWnd);
            HMENU configSubMenu = GetSubMenu(mainMenu, 2);
            if (dedupeMemoField) {
                // Option was previously selected, so disable it
                CheckMenuItem(configSubMenu, 
                              ID_CONFIG_DEDUPE_MEMO,
                              MF_UNCHECKED);
                dedupeMemoField = false;
            }
            else {
                // Option was previously unselected, so enable it
                CheckMenuItem(configSubMenu,
                              ID_CONFIG_DEDUPE_MEMO,
                              MF_CHECKED);
                dedupeMemoField = true;
            }
            break;
        }
        default:
            break;
        }
    }
    default: {
        // From the docs: Calls the default window procedure to provide default
        // processing for any window messages that an application does not 
        // process. This function ensures that every message is processed. 
        return DefWindowProc(hWnd, message, wParam, lParam);
        break;
    }  // end default
    }  // end switch

    return 0;
}

int CALLBACK WinMain(
    _In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPSTR     lpCmdLine,
    _In_ int       nCmdShow
) {
    WNDCLASSEX wcex;
    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.cbClsExtra = 0;
    wcex.cbWndExtra = 0;
    wcex.hInstance = hInstance;
    wcex.hIcon = LoadIcon(hInstance, IDI_APPLICATION);
    wcex.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wcex.lpszMenuName = NULL;
    wcex.lpszClassName = szWindowClass;
    wcex.hIconSm = LoadIcon(wcex.hInstance, IDI_APPLICATION);

    if (!RegisterClassEx(&wcex)) {
        MessageBox(NULL,
            _T("Call to RegisterClassEx failed!"),
            _T("Cannot create window. Exiting."),
            NULL);
        return 1;
    }
    HWND hWnd = CreateWindow(
        szWindowClass,
        szTitle,
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT,
        700, 600,
        NULL,
        NULL,
        hInstance,
        NULL
    );
    if (!hWnd) {
        MessageBox(NULL,
            _T("Call to CreateWindow failed!"),
            _T("Error"),
            MB_ICONERROR);
        return 1;
    }

    // Set the Application Icon to an Exclamation Icon.
    SetClassLongPtr(hWnd,          // window handle 
        GCLP_HICON,              // changes icon 
        (LONG)LoadIcon(NULL, IDI_EXCLAMATION)
    );

    CreateMainMenu(hWnd);

    // The parameters to ShowWindow explained:
    // hWnd: the value returned from CreateWindow
    // nCmdShow: the fourth parameter from WinMain
    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    // If given a filename param (e.g. "Open with..."), open the file here.
    int argCount;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLine(), &argCount);
    if (argCount == 2) {
        LoadFile(argv[1], hWnd);
    }

    // Create Keyboard Accelerators that are invoked using ALT + KEY
    const int ACCELERATOR_TABLE_SIZE = 4;
    ACCEL accelTable[ACCELERATOR_TABLE_SIZE];
    accelTable[0].cmd = ID_ACTIONS_CONVERT_TO_OFX;
    accelTable[0].fVirt = FALT | FVIRTKEY;
    accelTable[0].key = 0x43; //'C' key
    accelTable[1].cmd = ID_ACTIONS_SAVE_OFX;
    accelTable[1].fVirt = FALT | FVIRTKEY;
    accelTable[1].key = 0x53; //'S' key
    accelTable[2].cmd = ID_ACTIONS_SEND_TO_MONEY;
    accelTable[2].fVirt = FALT | FVIRTKEY;
    accelTable[2].key = 0x49; //'I' key
    accelTable[3].cmd = ID_FILE_OPEN;
    accelTable[3].fVirt = FALT | FVIRTKEY;
    accelTable[3].key = 0x4F; //'O' key
    HACCEL accels = CreateAcceleratorTable(accelTable, ACCELERATOR_TABLE_SIZE);

    // To get an idea of usage metrics, ping my web server on startup.
    // The user agent is set as the MD5 hash of the Computer Name, which
    // allows us to roughly determine a unique instance.
    char buffer[MAX_COMPUTERNAME_LENGTH + 1] = { "UNKNOWN" };
    DWORD len = 7;
    GetComputerNameA(buffer, &len);
    // Now MD5 hash/obfuscate the computer name
    BYTE hashedComputerName[256] = {};
    DWORD hashedLength = 256;
    HCRYPTPROV hCryptProv;
    HCRYPTHASH hHash;
    if (CryptAcquireContext(&hCryptProv, NULL, NULL, PROV_RSA_FULL,
        CRYPT_VERIFYCONTEXT | CRYPT_MACHINE_KEYSET)) {
        if (CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, &hHash)) {
            if (CryptHashData(hHash, (const BYTE*)buffer, len, 0)) {
                CryptGetHashParam(hHash, HP_HASHVAL, hashedComputerName,
                    &hashedLength, 0);
            }
        }
    }
    // The hashed value is in a byte array. Convert it to string.
    std::ostringstream oss;
    for (DWORD i = 0; i < hashedLength; ++i) {
        oss << std::hex << static_cast<const int>(hashedComputerName[i]);
    }
    std::string compNameMD5 = oss.str();
    std::wstring md5 = std::wstring(compNameMD5.begin(), compNameMD5.end());
    md5 += L" v" + VERSION_ID;
    CryptDestroyHash(hHash);
    CryptReleaseContext(hCryptProv, 0);

    // Send a HTTP ping to help determine application usage
    // User-Agent is a md5 of the computer name + version id.
    HINTERNET hSession = WinHttpOpen(md5.c_str(),
        WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
        WINHTTP_NO_PROXY_NAME,
        WINHTTP_NO_PROXY_BYPASS,
        WINHTTP_FLAG_ASYNC);
    if (hSession) {
        // Use HTTP because no sensitive data is sent and HTTPS can
        // have issues owing to skewed clocks, faulty certs, and more.
        HINTERNET hConnect = WinHttpConnect(hSession, L"www.norcalico.com",
            INTERNET_DEFAULT_HTTP_PORT, 0);
        // Create an HTTP request handle.
        if (hConnect) {
            HINTERNET hRequest = WinHttpOpenRequest(hConnect,
                L"GET", L"/ConvertToOFX/usage/", NULL, WINHTTP_NO_REFERER,
                WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_REFRESH);
            // Send a request.
            if (hRequest) {
                WinHttpSendRequest(hRequest,
                    WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                    WINHTTP_NO_REQUEST_DATA, 0,
                    0, 0);
            }
        }
    }

    // Main message loop:
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        if (!TranslateAccelerator(
            hWnd,
            accels,
            &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int)msg.wParam;
}
