---
layout: post
title: "Style Test"
tags: [test]
comments: false
published: false
---

Dieser Text demonstriert die Style-Elemente  von Jekyll bzw. Markdown. Schau in den Quelltext um die Anweisungen für die eingebetten Elemente zu sehen.  

---

## 1. Header 

# Header H1 (reserved for post titles)

## Header H2

### Header H3

#### Header H4

##### Header H5

###### Header H6

##### H4 Center
{: .center}

##### H4 Right
{: .right}

## 2. Text

Links können im [Text](https://rhombuzz.github.io) stehen oder vollständig gezeigt werden <https://rhombuzz.github.io>. 

Inline markup funktioniert wie folgt:

**Das ist fett.**

*Das ist kursiv.*

_Das ist auch kursiv._

`Das ist Code.`

<u>Das ist unterstrichen</u>.

<del>Das ist durchgestrichen</del>.

<mark>Igitt, das ist gelb markiert.</mark> 

Formeln: 5<sup>3</sup> = 125. Wasser ist H<sub>2</sub>O. 

Es gibt sogar Fussnoten. [^1]

[^1]: <http://en.wikipedia.org/wiki/Syntax_highlighting>

Benutze zwei Leerzeichen   
rechts am Ende der Zeile  
um Zeilenumbrüche zu erzeugen.


## 3. Images

![example image]({{ site.baseurl }}/images/logo.jpg "Small-sized example image")

![Center example image](/images/logo.jpg "Center"){: .center-image}


## 4. Blockquotes

> Blockquote. Wer liebt nicht Zitate?
>> nested Blockquote.


## 5. Listen und Aufzählungen

### ungeordnete Liste

* Lorem ipsum dolor sit amet, consectetur adipiscing elit.
* Nam ultrices nunc in nisi pellentesque ultricies. Cras scelerisque ipsum in ante laoreet viverra. Pellentesque eget quam et augue molestie tincidunt ac ut ex. Sed quis velit vulputate, rutrum nisl sit amet, molestie neque. Vivamus sed augue at turpis suscipit fringilla.
* Integer pretium nisl vitae justo aliquam, at varius nisi blandit.
  1. Nunc vehicula nulla ac odio gravida vestibulum sed nec mauris.
  2. Duis at diam eget arcu dapibus consequat.
* Etiam vel elit in purus iaculis pretium.

### nummerierte Liste

1. Quisque ullamcorper leo non ex pretium, in fermentum libero imperdiet.
2. Donec eu nulla euismod, rhoncus ipsum nec, faucibus elit.
3. Nam blandit purus gravida, accumsan sem in, lacinia orci.
  * Duis congue dui nec nisi posuere, at luctus velit semper.
  * Suspendisse in lorem id lacus elementum pretium nec vel nibh.
4. Aliquam eget ipsum laoreet, maximus risus vitae, iaculis leo.

### Definition (Liste)

kramdown
: A Markdown-superset converter

Maruku
: Another Markdown-superset converter


## 6. Tabellen

| Header1 | Header2 | Header3 |
|:--------|:-------:|--------:|
| cell1   | cell2   | cell3   |
| cell4   | cell5   | cell6   |
|----
| cell1   | cell2   | cell3   |
| cell4   | cell5   | cell6   |
| cell7   | cell8   | cell9   |
|=====
| Foot1   | Foot2   | Foot3


## 7. Code Snippets

### Syntax Highlighting


Python Code wird dem Schlüsselwort {% raw %}```python{% endraw %} eingeleitet.

```python
# python

def test(name):
    print('Hello' + name)

print test(john)
```

VB Code wird dem Schlüsselwort {% raw %}```visualbasic{% endraw %} eingeleitet.

```visualbasic
' visual basic

Function test(a, b)
    test = a + b + 1000
End Function
```

Der Code-Block wird jeweils mit {% raw %}```{% endraw %} abgeschlossen.
Das Styling und die Farben werden in `/_sass/_highlighter.scss` modifiziert.

Ein Standard-Code-Block ohne highlighting

    <div id="awesome">
      <p>This is great isn't it?</p>
    </div>



### GitHub Gist Embed

An example of a Gist embed below.

<script src="https://gist.github.com/mmistakes/43a355923921d22cd993.js"></script>

Below is a partial code showing main steps of merge function.

<code data-gist-id="0fe211678316cc53370c" data-gist-file="merge_tables_datatable.R" data-gist-line="50-52,57,65-69,80,88-90,100-106"></code>


Und Schlussendlich: Horizontale Linien 
----
****
