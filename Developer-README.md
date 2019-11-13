# Developer README for ConvertToOFX

To make changes, you will need Visual Studio with the Windows 10 SDK. Despite using the Windows 10 SDK, strive to ensure that any APIs used are backward compatible with older versions of Windows (e.g. Windows 7) as much as reasonably possible.


# Dependencies

This project requires the following dependencies:
* TinyXML-2: https://github.com/leethomason/tinyxml2


# How to Build (Details Step by Step Instructions)

1) First, install Visual Studio Community 2019, and make sure you install the Windows 10 SDK.

2) Place the contents of TinyXML-2 (https://github.com/leethomason/tinyxml2) into a directory in your repos folder.

3) Place the contents of ConvertToOFX into a directory in your repos folder.

4) In the ConvertToOFX directory, open src/ConvertToOFX.sln with Visual Studio.

5) In the ConvertToOFX project's properties, under "Configuration Properties" -> "VC++ Directories", for All Configurations and All Platforms, append the "Include Directories" to also include the tinyxml2 folder.

6) Change your Solution Configuration to "Release", select the appropriate Platform as needed, and select "Debug" from the menu and "Start without Debugging"

7) The ConvertToOFX.exe file will now be under your ConvertToOFX directory, in the platform directory (e.g. x86), in the Release folder.


# Notes on Signing the EXE

To sign the EXE, one must perform the following:

0) Install PublishOnce tools to get the SignTool.exe. It's installed via the Visual Studio Installer, modify your existing installation to get this package.

1) Generate a certificate or import the certificate:
  * To Generate, run these 3 commands:
    * `$cert = New-SelfSignedCertificate -DnsName www.norcalico.com -Type CodeSigning -CertStoreLocation Cert:\CurrentUser\My$($cert.Thumbprint)`
    * `$CertPassword = ConvertTo-SecureString -String "PASSWORD" -Force –AsPlainText`
    * `Export-PfxCertificate -Cert "cert:\CurrentUser\My\$($cert.Thumbprint)" -FilePath "c:\selfsigncert.pfx" -Password $CertPassword`

2) Sign: C:\Program Files (x86)\Microsoft SDKs\ClickOnce\SignTool>signtool.exe sign "C:\Users\username\Source\Repos\ConvertToOFX\Release\ConvertToOFX.exe"

3) Make a backup copy of c:\selfsigncert.pfx and the PASSWORD you used. In the future, you will only need to run step 2 to sign new binaries.


# Suggestions for Future
These are features I want to add but did not get a chance to. Maybe one day:
* Find a way to persist changes to the location of the Money Import Handler (maybe use an INI file?). Right now, the user must change the location manually every time.
* Check if there is a new version available and let the user know to update.
** The problem I ran into here was that I want to do it asynchronously. The Async HTTP code is horrible. I worried that I would introduce crash conditions with such code. I scrapped it because this is not vital functionality and the risks were worse than the benefits.
* Add an option to convert XML tags to uppercase (and persist that option)
* Add logic to run this from the command prompt and disable all MessageBox prompts (like a "batch" mode)
* Add parsing for other statement types. Need to investigate what types Money supports.
* Add the ability to encrypt and submit un-parseable files (with explicit user permission in each case) so that I can inspect them and fix bugs.

# Statistics
* I estimate this took me 90 hours to write. I have a passing familiarity with C++ but this was my very first Win32 API program, so it took me a while to understand how to use the API.

## Azure Pricing
I used Azure Credits to develop this since I don't have a Windows machine that can run Visual Studio. I used a B2S VM (Currently: $0.0496/hr) and a 128 GB Standard HDD disk (Type S10, $5.888/month). The B1S VM is too slow. Total Cost reported by Azure: $3.74.
* Storage breakdown: $2.96: Implies 1/2 month of usage, which is accurate.
  * I could have used a smaller 32 GB disk for this project.
* VM breakdown: $0.78: Implies 15 hours of usage. This does not look accurate. If I had 90 hours of usage, I'd expect cost to be 90*$0.0496: $4.47
