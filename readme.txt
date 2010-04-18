About adomxmldom.pas and gtcAdomxmldom.pas:

gtcAdomxmldom.pas is intended to be compiled with the v4 or v5 version of
Dieter Köhler's ADOM components, as downloaded directly from his site.
His source files are prefixed by dk, in order to allow us to compile and use
his latest version of ADOM, while coexisting with the version supplied with
Delphi 2010 onwards.

adomxmldom.pas is intended to be equal to the source supplied in the RTL of
Delphi 2010 onwards.
It is identical to gtcAdomxmldom.pas, except for the unit name and the use of a
conditional complation constant that identifies it as being the RTL version.


In order to install and use gtcAdomxmldom in Delphi 2010, install the
packages\gtcAdomxmldomD14.dpk.
Then, when you want to use a TXMLDocument, set its DOMVendor to 'ADOM XML v4 (custom)'.
The RTL version of the ADOM wrapper is selected by setting DOMVendor to
'ADOM XML v4'.

