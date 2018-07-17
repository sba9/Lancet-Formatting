The LancetFormatting macro is meant to speed up the formatting and copyediting process for papers being submitted to the Lancet for publication. In this repo, there are four files necessary for supporting and running the macro.

# Installation/Use
The only file needed to run this macro is LancetFormatting.bas. After downloading the file, open MS Word. Press __alt + F11__ to open the macro editor screen. Go to __File > Import File__, and then import the macro file. 

To run the macro, go back to the document. Switch to the __View__ tab, and select __Macros__. In the dialog box, find __LancetFormatting__, and hit __Run__. 

# Files
## LancetFormatting
LancetFormatting.bas is the macro itself. To install and run the macro in Word, press alt+F11 to bring up the macro editor. Go to File > Import File to load the macro. To run, go to the document, and then to the View tab > Macros and select "LancetFormatting". The macro will then search through the document and:

- Replaces American English words with British English equivalents
- Swaps i.e./e.g. with ie/eg
- Replaces other words with Lancet and IHME standards
- Replaces decimal points with floating decimal points
- Uses en-dashes instead of hyphens in year ranges
- Uses em-dashes instead of hyphens for all other hyphens except in negative numbers
- Highlights in red any uncertainty intervals that seem to straddle zero

All changes will be highlighted in "dark yellow" for visibility.

## VBA_macro_generator
VBA_macro_generator.py is a script to update the LancetFormatting code. Because VBA doesn't allow file IO for macros, the dictionary of word replacements needs to be inserted manually; this script saves time and energy by generating it programmatically whenever the user updates the dictionary.

## British English
British English.csv is the dictionary for all word replacements. When LancetFormatting finds a word or phrase from the second column in the document, it replaces it with the corresponding entry in the first column.

## test bed
test bed.docx contains examples of text that should or should not be adjusted by the macro. It is meant to test any updates made to the code.
