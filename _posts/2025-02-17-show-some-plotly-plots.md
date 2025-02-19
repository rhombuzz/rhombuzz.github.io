---
layout: post
title: "Plotly trifft Jekyll"
description: "Hier werden verschiedene Wege getestet, um Plotly-Grafiken einzubinden."
tags: [plotly]
comments: false
published: True
---

1) Mittels Jinja wird eine eigenständige html-Datei erstellt, die einfach verlinkt werden kann.

[SGB II-Quote Choropleth]({{ site.baseurl }}{% link pages/sgb2Q_choropleth_bar_v12.html %})  
[SGB II-KdU Choropleth]({{ site.baseurl }}{% link pages/KdU_choropleth_bar_v1.html %})

2) Eine nicht-eigenständige html-Datei wird mittels iframe eingebunden.

... das ist ja nicht so gut gelaufen.

