Dies ist ein kleines Beispiel für einen XmlParser
ursprünglich war der Code in C# und wurde mit ICSharpCode.com
nach VB.NET übersetzt, die Klassen:
SmallXmlParser, IContentHandler, DefaultHandler
IAttrList und AttrListImpl stammen aus der Feder von 
Atsushi Enomoto, vom Mono.NET-Framework. <atsushi@ximian.com>
(siehe "\\mono-1.1.16.1\mcs\class\corlib\Mono.Xml\")

Tutorial zum Namespace *corlib.System.IO
>>Datei Einlese- und Schreibvorgänge

Q: Wie spielen die einzelnen Klassen im Namespace corlib.System.IO
des .net-Frameworks zusammen und wie gebraucht man die einzelnen Klassen
fürs Dateihandling in einem Projekt?

A: Um unter .net eine Datei zu öffnen, zu schreiben oder lesen, gibt es 
verschiedene Möglichkeiten.
Die grundlegende Klasse für Datenlese- und Schreibvorgänge ist die 
abstrakte Klasse:
 Stream
von ihr abgeleitet ist zunächst die Klasse: 
 FileStream
(Sie implementiert also das Interface Stream, und überall wo an eine 
Funktion ein Übergabeparameter vom Typ Stream gebraucht wird, kann man 
ebensogut ein Objekt vom Typ FileStream übergeben.) 

um Textdaten zu Lesen oder zu Schreiben gibt es die zwei verschiedenen 
Klassentypen, die das Suffix "Reader" bzw. "Writer" haben.

1. Ganz allgemein um "Textdateien" zu lesen gibt es die grundlegende 
abstrakte Klasse:
 TextReader 
davon abgeleitet sind die Klassen
 StreamReader und
 StringReader
(Sie implementieren also das Interface TextReader, und überall wo an eine 
Funktion ein Übergabeparameter vom diesem Typ gebraucht wird, kann man 
ebensogut ein Objekt von den beiden abgeleiteten Klassen übergeben.)

2. Ganz allgemein um "Textdateien" zu schreiben gibt es die grundlegende 
abstrakte Klasse:
 TextWriter
davon abgeleitet sind die Klassen
 StreamWriter und
 StringWriter 
(Sie implementieren also das Interface TextWriter, und überall wo an eine 
Funktion ein Übergabeparameter vom diesem Typ gebraucht wird, kann man 
ebensogut ein Objekt von den beiden abgeleiteten Klassen übergeben.)
 
Für einfachere Programme um nur mal schnell eben eine Datei zu speichern 
ist die Klasse FileStream allerdings ausreichend, da sie bereits grundlegende
Funktionen hat um Dateien zu öffnen, zu schreiben oder zu lesen, zu speichern
bzw. zu schließen.

Die Klassen StreamWriter und StreamReader können im Zusammenhang mit Encoding
benützt werden.
Die Klassen StringWriter und StringReader schreiben nur Strings.
Die Klsse StringWriter implementiert dazu übrigens ein Objekt vom Typ
StringBuilder, aus dem Namespace *corlib.System.Text. 