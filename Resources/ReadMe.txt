Dies ist ein kleines Beispiel f�r einen XmlParser
urspr�nglich war der Code in C# und wurde mit ICSharpCode.com
nach VB.NET �bersetzt, die Klassen:
SmallXmlParser, IContentHandler, DefaultHandler
IAttrList und AttrListImpl stammen aus der Feder von 
Atsushi Enomoto, vom Mono.NET-Framework. <atsushi@ximian.com>
(siehe "\\mono-1.1.16.1\mcs\class\corlib\Mono.Xml\")

Tutorial zum Namespace *corlib.System.IO
>>Datei Einlese- und Schreibvorg�nge

Q: Wie spielen die einzelnen Klassen im Namespace corlib.System.IO
des .net-Frameworks zusammen und wie gebraucht man die einzelnen Klassen
f�rs Dateihandling in einem Projekt?

A: Um unter .net eine Datei zu �ffnen, zu schreiben oder lesen, gibt es 
verschiedene M�glichkeiten.
Die grundlegende Klasse f�r Datenlese- und Schreibvorg�nge ist die 
abstrakte Klasse:
 Stream
von ihr abgeleitet ist zun�chst die Klasse: 
 FileStream
(Sie implementiert also das Interface Stream, und �berall wo an eine 
Funktion ein �bergabeparameter vom Typ Stream gebraucht wird, kann man 
ebensogut ein Objekt vom Typ FileStream �bergeben.) 

um Textdaten zu Lesen oder zu Schreiben gibt es die zwei verschiedenen 
Klassentypen, die das Suffix "Reader" bzw. "Writer" haben.

1. Ganz allgemein um "Textdateien" zu lesen gibt es die grundlegende 
abstrakte Klasse:
 TextReader 
davon abgeleitet sind die Klassen
 StreamReader und
 StringReader
(Sie implementieren also das Interface TextReader, und �berall wo an eine 
Funktion ein �bergabeparameter vom diesem Typ gebraucht wird, kann man 
ebensogut ein Objekt von den beiden abgeleiteten Klassen �bergeben.)

2. Ganz allgemein um "Textdateien" zu schreiben gibt es die grundlegende 
abstrakte Klasse:
 TextWriter
davon abgeleitet sind die Klassen
 StreamWriter und
 StringWriter 
(Sie implementieren also das Interface TextWriter, und �berall wo an eine 
Funktion ein �bergabeparameter vom diesem Typ gebraucht wird, kann man 
ebensogut ein Objekt von den beiden abgeleiteten Klassen �bergeben.)
 
F�r einfachere Programme um nur mal schnell eben eine Datei zu speichern 
ist die Klasse FileStream allerdings ausreichend, da sie bereits grundlegende
Funktionen hat um Dateien zu �ffnen, zu schreiben oder zu lesen, zu speichern
bzw. zu schlie�en.

Die Klassen StreamWriter und StreamReader k�nnen im Zusammenhang mit Encoding
ben�tzt werden.
Die Klassen StringWriter und StringReader schreiben nur Strings.
Die Klsse StringWriter implementiert dazu �brigens ein Objekt vom Typ
StringBuilder, aus dem Namespace *corlib.System.Text. 