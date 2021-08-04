# XL_VBanywhere
## VB-Code running in VBC as well as in VBA7-x86 and x64  

[![GitHub](https://img.shields.io/github/license/OlimilO1402/XL_VBanywhere?style=plastic)](https://github.com/OlimilO1402/XL_VBanywhere/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/XL_VBanywhere?style=plastic)](https://github.com/OlimilO1402/XL_VBanywhere/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/XL_VBanywhere/total.svg)](https://github.com/OlimilO1402/XL_VBanywhere/releases/download/v1.0.0/XL_VBanywhere.zip)
[![Follow](https://img.shields.io/github/followers/OlimilO1402.svg?style=social&label=Follow&maxAge=2592000)](https://github.com/OlimilO1402/XL_VBanywhere/watchers)

Tutorial published on 04. aug. 2021. at ActiveVB.de/VBA-Forum

Tutorial in german, deutsch

#VB-Code, lauffähig in VBC als auch VBA7-x86 und -x64  
## VB bedingte Kompilierung  

In der VBC-IDE gibt es in den Projekteigenschaften unter "Erstellen" "Argumente für bedingte Kompilierung" die Möglichkeit Konstanten zu definieren mit denen das Verhalten des Compilers nach Belieben gesteuert werden kann.
Ebenso in der VBA-IDE zu finden unter "Extras" "Eigenschaften von VBAProjekt..."
Mehrere Konstanten werden mit einem Doppelpunkt ":" voneinander getrennt z.B.:

Mode_Beta = 0 : Mode_Debug = 1 : VBC = 1

![XL_VBanywhere Image](ProjekteigArgFBedKomp.png "ProjekteigArgFBedKomp Image")

Zusätzlich kann man mit dem #If-Statement-für-bedingte-Kompilierung diese Konstanten abfragen um dem Compiler zu sagen was er Kompilieren soll und was er beim Kompilieren weglassen soll. z.B.:

<PRE><FONT SIZE=2 FACE=Consolas><FONT COLOR=#0000FF>Sub</FONT> Main()
<FONT COLOR=#0000FF>End</FONT> <FONT COLOR=#0000FF>Sub</FONT>

##VBA7 x86 und x64  
Seit ca 2007 gibt es VBA7 auch für die 64-Bit Plattform. Mit VBA7 wurden 2 neue vorbelegte Kompiler-Konstanten und neue Schlüsselwörter eingeführt.  
