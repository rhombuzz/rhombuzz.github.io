---
layout: post
title: "Crashkurs VBA"
description: "Ein zügiger Einstieg in die Abgründe von VBA für Excel."
tags: [vba, excel]
comments: false
published: false
---

Möge der folgende Einstieg den lieben Kolleginnen und Kollegen von nachhaltigem Nutzen sein ...

## Intro

Visual Basic (VB) ist eine objektorientierte Programmiersprache. VBA (for Applications) ist ein Ableger, der zur Ablaufsteuerung bzw. Automatisierung von MS Office-Produkten verwendet wird. Die Objektorientierung bietet den Vorteil, dass zahlreiche applikations- also excel-interne Eigenschaften (Objekte und Attribute) und Methoden (Prozeduren) angesprochen werden können. 

Die Objektorientierung führt zu Code mit der unverkennbaren Punkt-Semantik, wie bspw. `Worksheets("Auswertung").Range("A1:D5").Value = 55`. VBA wirkt zu Beginn oft schwierig, da man optisch nicht zwischen Eigenschaften und Methoden unterscheiden kann bzw. auch für Fortgeschrittene häufig unklar bleibt, wie bestimmte Eigenschaften oder Objekte -- von denen man weiß, dass es sie geben muss -- eigentlich heißen.

Excel-Dateien mit VBA-Code werden in Daten mit der Endung `*.xlsm` gespeichert. Die einzelnen Makros und Funktionen werden in sog. Modulen abgelegt; diese sind vergleichbar mit Arbeitsblättern nur eben für VBA-Code. Prozeduren sind eigenständige "Programme" bzw. Makros und werden mit dem Schlüsselwort `Sub` definiert. Funktionen sind eigenständige Zellfunktionen vergleichbar mit excel-internen wie `Summe` oder `Mittelwert`, die bei Eingabe von Argumenten einen Rückgabewert in die Zelle schreiben; sie werden mit dem Schlüsselwort `Function` definiert.

Es gibt arithmetische Operatoren: `+ - * / ^`, Vergleichsoperatoren `= < > <= >= <> is like` und logische Operatoren `not and or xor eqv`.

Weiterführendes unter <https://docs.microsoft.com/de-de/office/vba/api/overview/excel>

## Einstieg in Funktionen 

Funktionen sind eigenständige Programme, die entweder durch ihre Verwendung in einer Zelle oder auch durch anderen VBA-Code aufgerufen werden können. Funktionen haben idR mehrere Eingabewerte und -parameter und liefern **einen** Rückgabewert. Die Eingabeparameter werden idR positional (d.h. in einer feststehenden Reihenfolge) übergeben.

Im Gegensatz zu den "deutschsprachigen" Arbeitsblättern, ist VBA "englisch". D.h. das Semikolon ist hier ein Komma (bspw. `round(23, 0)`), das Komma ist ein Punkt (bspw. bei der Zahl `2.3`) In VBA integrierte Excel-Functionen müssen mit ihrer englischen Bezeichnung angegeben werden.

### erste Schritte

Eine Funktion wird mit ihrem Namen (hier `sum`) definiert, in der folgenden Klammer werden die Variablennamen der Eingabeparameter (hier `a` und `b`) spezifiziert. Der Rückgabewert wird im weiteren Ablauf dem Funktionsnamen zugewiesen.

```visualbasic
Function sum(a, b)
    sum = a + b
End Function
```

Variablen (für Zwischenergebnisse o.ä.) lassen sich innerhalb der Funktion in beliebiger Anzahl generieren, indem man ihrem Namen einfach einen Wert zuweist. Man kann leere Variablen auch mit dem Schlüsselwort `Dim` erstellen.

```visualbasic
Function mult(a, b)
    f1 = a
    f2 = b
    f2 = f1 * f2    'f2 wird mit einem neuen Wert überschrieben
    mult = f2
End Function
```

Konstanten sind unveränderliche Variablen, die nicht nur innerhalb einer Funktion sondern im gesamten Modul gelten. Sie werden mittels des Schlüsselwortes `Const` definiert und stehen **vor** der ersten Funktion des Moduls (warum auch immer).

```visualbasic
Const cnst = 1000
```

Funktionen können nicht nur Konstanten sondern auch andere Funktionen aufrufen. Dies können vba-interne Funktionen ([Liste der Namen](https://www.excelfunctions.net/vba-functions.html)) als auch jede vorhandene Excel-Funktion sein ([Liste der Namen](https://docs.microsoft.com/de-de/office/vba/api/excel.worksheetfunction)).

```visualbasic
Function sum_mult(x, y)
    tmp = mult(x, cnst)    
    sum_mult = sum(tmp, y)    
    ' vba-interne Funktion Runden
    sum_mult = Round(sum_mult, 1)    
    ' excel-interne Funktion Runden 
    sum_mult = Application.WorksheetFunction.Round(sum_mult, 0)
End Function
```

## Kontrollstrukturen

Mittels `If-Then-Else` werden Fallunterscheidungen getroffen. D.h. Teile des Codes werden nur ausgeführt, wenn vorgenannte Bedingungen erfüllt sind. Der Bedingungsblock wird mit den Schlüsselwort `End If` abgeschlossen.

```visualbasic
Function if_vgl(a)
    If a < 0 Then
        if_vgl = "kleiner Null"        
    ElseIf a = 1 Then
        if_vgl = "gleich Eins"        
    ElseIf a >= 2 And a <= 4 Then
        if_vgl = "zwischen Zwei und Vier" 
    ElseIf a = 23 Or a = 42 Or a = 666 Then
        if_vgl = "vorzügliche Wahl"       
    Else
        if_vgl = "irgendwas anderes"        
    End If
End Function
```

Geht es darum, bestimmte Zahlenwerte oder -bereiche zu unterscheiden, bietet sich die Verwendung von `Select Case` an.

```visualbasic
Function case_vgl(nmbr)
    Select Case nmbr
        Case Is < 0
            case_vgl = "kleiner Null"
        Case 1
            case_vgl = "gleich Eins"
        Case 2 To 4
            case_vgl = "zwischen Zwei und Vier"
        Case 23, 42, 666
            case_vgl = "vorzügliche Wahl"
        Case Else
            case_vgl = "irgendwas anderes"
    End Select
End Function
```

Für das gleiche Ergebnis wie in den obigen zwei Beispielen würde man die Zellenformel `=WENN(C37<0;"kleiner Null";WENN(C37=1;"gleich Eins";WENN(UND(C37>=2;C37<=4);"zwischen Zwei und Vier";WENN(ODER(C37=23;C37=42;C37=666);"vorzügliche Wahl";"irgendwas anderes"))))` benötigen.[^1] 
Das ist nicht sonderlich übersichtlich.

[^1]: Der Wert für `nmbr` steht im Beispiel in Zelle `C37`.

## Schleifen (loops)

Mit Schleifen werden Anweisungen wiederholt. Die Schleife bricht ab, wenn entweder eine Bedingung erfüllt ist oder eine (feststehende) Anzahl von Durchläufen erreicht.

Bei der `while`-Bedingungsschleife steht die Anzahl der Durchläufe nicht vorab fest; sie stoppt erst, wenn die Abbruchbedingung erfüllt ist.[^2]

[^2]: MS bricht hier mit dem bisherigen Schema von `End NAME` sondern erfindet für das Ende der `while`-Schleife ein neues Schlüsselwort: `Wend`. 

```visualbasic
Function sqr_till_max(nmbr, Optional limit = 888)
    While nmbr <> 1 And (nmbr ^ 2) < limit
        nmbr = nmbr ^ 2
    Wend
    sqr_till_max = nmbr
End Function
```

Bei der `for-next`-Zählschleife ist die Anzahl der Durchläufe im weitesten Sinne determiniert. Im simpelsten Fall durch einen Eingangsparameter.

```visualbasic
Function mad_adder(nmbr, runs)
    For i = 1 To runs
        nmbr = nmbr + nmbr
    Next i
    mad_adder = nmbr
End Function
```

Bei `for-each-next` entspricht die Anzahl der Durchläufe der Anzahl der Elemente einer Gruppe. So ist das excel-interne Objekt `Worksheets` eine Gruppe einzelner Arbeitsblätter.

```visualbasic
Function wrksht_nr()
    wrksht_nr = 0
    For Each blatt In Worksheets
        wrksht_nr = wrksht_nr + 1
    Next
End Function
```

Der Datentyp `Array` besteht aus `for-each`-iterierbaren Elementen. Das folgende Beispiel zeigt die Funktionsweise der Schleife Schritt für Schritt durch die Ausgabe in der `MsgBox`[^3] und gibt am Ende den Wert des Elementes mit dem Index Nummer 1 zurück. Da die Indizes eines `Array` in VBA [mit 0 beginnen](https://en.wikipedia.org/wiki/Zero-based_numbering), wird deshalb der zweite Wert angezeigt.

[^3]: Das Verwenden einer `MsgBox` in einer Funktion ist keine gute Idee, da sie jedes mal aufgerufen wird, wenn die Zelle mit der Formel neu berechnet wird. `MsgBox` finden daher eher in Makros Anwendung.

```visualbasic
Function array_loop()    
    a = Array(1, 2, 3)
    For Each i In a
        MsgBox (i)
    Next
    array_loop = a(1)
End Function
```

## Zellbereiche mit dem Range-Objekt verarbeiten

Bis jetzt wurden den Funktionen nur einzelne Zellen als Eingabeparameter übergeben. Häufig will man aber mit Zellbereichen (vorab unbekannter Größe) arbeiten. Damit das funktioniert, muss der Eingabeparameter von vornherein `As Range` definiert werden. Weiterführendes zum `Range`-Objekt gibt es [hier](https://docs.microsoft.com/de-de/office/vba/api/excel.range(object)). 

Im folgenden Beispiel werden die Eigenschaften `.Columns`, `.Rows` und `.Cells` des `Range` angesprochen und jeweils mittels der Methode `.Count` gezählt. `.Columns`, `.Rows` und `.Cells` sind -- das suggeriert die Mehrzahl bereits -- wieder Gruppen von Spalten, Zeilen oder Zellen. Eigenschaften -- oder besser Unterobjekte -- wie `.Cells` haben selbst wieder Attribute (wie `.Value` oder `.Font`) und unterstützen bestimmte Methoden (wie `.Select`, `.Copy`, `.Insert` oder `.Count`).

```visualbasic
Function xrange(rng As Range)
    spaltz = rng.Columns.Count
    zeilez = rng.Rows.Count
    zellez = rng.Cells.Count
    xrange = ("Spalten: " & spaltz & ", Zeilen: " & zeilez & ", Zellen: " & zellez)  
End Function
```

Ein `Range` kann als eine Gruppe von Zellen (oder Spalten oder Zeilen) verstanden werden. Die einzelnen Elemente können über ihren Index angesprochen werden, der -- im Gegensatz zum `Array` -- in diesem Fall mit 1 beginnt (man merkt, Konsistenz ist nicht so MS' Sache). 

Der Rückgabewert des nächsten Beispiels entspricht der Eigenschaft `.Value` der Zelle mit dem entsprechenden Index. Die entsprechende Zelle wird durch die Eigenschaft `.Cells` gefunden. Die einzelnen `.Cells` eines Range-Objektes werden mit dem Index zunächst spalten- dann zeilenweise durchgezählt.

```visualbasic
Function getvaluefromindex(rng As Range, idx)
    getvaluefromindex = rng.Cells(idx).Value  
End Function
```

kurzer Exkurs: Das Worksheet-Objekt besitzt ebenfalls eine Eigenschaft namens `.Cells`. Mit dieser kann man eine Zelle über ihre absolute Position (Zeile, Spalte) im Blatt angesprochen werden.

```visualbasic
Function getvaluefromposition(zeile, spalte)
    getvaluefromposition = ActiveSheet.Cells(zeile, spalte)
End Function
```

Zurück zum `Range` mit dem nächsten Beispiel: Die einzelnen Zellen des Bereiches werden in der Schleife über ihre Position / ihren Index ausgelesen. Die Zahl der loop-Durchläufe wird über die Gesamtzahl der Zellen des eingegebenen Zellbereiches festgelegt.

```visualbasic
Function mean(rng As Range)    
    summe = 0      
    For i = 1 To rng.Cells.Count
        summe = summe + rng.Cells(i).Value
    Next i    
    mean = summe / rng.Cells.Count        
End Function
```

Alternativ zum Durchlaufen der einzelnen Indizes  lassen sich Zellbereiche auch mit der `for-each`-Schleife verarbeiten. Die einzelnen Zellen werden der Schleife als iterierbare Teilelemente des Range-Objektes übergeben. Die Zahl der loop-Durchläufe muss nicht vorab vom user bestimmt werden, da der loop stoppt, wenn es keine Zellen mehr zu übergeben gibt.

```visualbasic
Function cell_adder(rng As Range, Optional min = 0, Optional max = 9999)
    cell_adder = 0    
    For Each zelle In rng
        If zelle >= min And zelle <= max Then
            cell_adder = cell_adder + zelle
        End If
    Next zelle    
End Function
```

VBA ist bei der Angabe der erfragten Attribute sehr 'nachsichtig'. In der eben gezeigten Funktion müssten eigentlich die Attribute der in der Schleife verwendeten Objekte weiter spezifizert werden (nämlich `rng.Cells` und `zelle.Value`). VBA verwendetet aber ohne genaue Angaben Standardwerte. Das erscheint -- gerade für Anfänger -- praktisch, da häufige Fehler gar nicht erst auftreten. Führt aber bei komplexeren Abläufen zu schwer nachvollziehbarem Verhalten.

Ein genaue Angabe des Attributes ist spätestens dann vonnöten, wenn man nicht nur über Zellen loopen möchte.

```visualbasic
Function col_loop(bereich As Range)
    For Each spalte In bereich.Columns
        For Each zelle In spalte.Cells
            MsgBox ("Zelle: " & zelle.Value & vbNewLine & "Spalte: " & spalte.Column)
        Next zelle
    Next spalte
End Function
```

----
