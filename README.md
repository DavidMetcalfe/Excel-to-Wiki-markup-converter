# Excel to Wiki markup generator

This VBA module takes selected columns/rows in Excel and converts them into Wiki-compatible tables.

This repository contains a modified version of Othmar Lippuner's code, which can be located on the [Dutch language Wikipedia](https://de.wikipedia.org/wiki/Wikipedia:Technik/Text/Basic/EXCEL-Tabellenumwandlung), or there's an [English translation]( https://de.wikipedia.org/wiki/Wikipedia:Technik/Text/Basic/EXCEL-Tabellenumwandlung) available as well.

## Instructions

### Installation:
1. Download the .BAS module from this respository.
2. Import the .BAS module file into a VBA-project in an Excel file of your choosing.

### Usage:
1. Select the range you wan't to publish in Excel
2. Execute the macro 'Format_as_wikitable'
3. Copy the complete wiki text in the 'wikioutput' Sheet.
4. Paste the clipboard text into the Wiki article of your choosing.
