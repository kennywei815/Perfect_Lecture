Perfect Lecture
======
Extend PowerPoint with technologies such as a LaTeX subsystem and a TTS system to allow much easier recording and revising of online course videos.

- Support MS Office 2010 and above.
- Integrated with [TeX4Office](https://github.com/kennywei815/TeX4Office) & [InsertImage++](https://github.com/kennywei815/InsertImagePlus)
- Output both slide format _(.pptx)_ & video format _(.wmv & .mp4)_
- Implemented a TTS subsystem with full support for [Speech Synthesis Markup Language(SSML)](https://www.w3.org/TR/speech-synthesis11/)
- Implemented a LaTeX subsystem and equation transpositon commands for math derivation
- Implemented a pointer subsystem which allows for using laser pointers in the slideshow

# Prerequisite
You need to install these packages first:
- [Ghostscript](https://www.ghostscript.com/download/gsdnld.html) for PDF file processing
- A LaTeX distribution: [MikTeX](https://miktex.org/download) (easier to install on Windows, recommended) or [TeX Live](https://www.tug.org/texlive/) (cross-platform)
- [Python3 interpreter](https://www.python.org/downloads/)

You can find archived installers for these programs with installation manuals in the [Prerequisite folder](https://github.com/kennywei815/Perfect_Lecture/blob/master/Prerequisite)!

# Quick Start

### Install
Just download [Setup_Perfect_Lecture.exe](https://github.com/kennywei815/Perfect_Lecture/raw/master/Setup_Perfect_Lecture.exe) and install it on your machine!

You can find the user manuals in the start menu. <br />

![start_menu.png](https://github.com/kennywei815/Perfect_Lecture/blob/master/www/start_menu.png)

### Usage

Just 1 step: Open the slides and type Perfect Lecture Scripts(enclosed by `<script>...</script>`) in the note page below. Then click "Compile with Perfect Lecture" in the "Add-ins" tab!
![step1_compile_with_perfect_lecture.png](https://github.com/kennywei815/Perfect_Lecture/blob/master/www/step1_compile_with_perfect_lecture.png)

The result slides and videos with machine-synthesized narrative will be put in the same directory.
![step4_results.PNG](https://github.com/kennywei815/Perfect_Lecture/blob/master/www/step4_results.PNG)

# License
Copyright (c) 2017 魏誠寬(Cheng-Kuan Wei) and 汪治平教授(Prof. Jyhpyng Wang) Licensed under the Apache License.

This software contains code derived from Jonathan Le Roux and Zvika Ben-Haim's IguanaTeX project which is originally released under the Creative Commons Attribution 3.0 License and combined into TeX4Office as a whole under Apache License 2.0 .