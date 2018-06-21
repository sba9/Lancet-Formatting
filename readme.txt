The LancetFormatting macro is meant to speed up the formatting and copyediting process for papers being submitted to the Lancet for publication. In this repo, there are four files necessary for supporting and running the macro.

LancetFormatting.bas is the macro itself. To install and run the macro in Word, press alt+F11 to bring up the macro editor. Go to File > Import File to load the macro. To run, go to the document, and then to the View tab > Macros and select "LancetFormatting". The macro will then search through the document and:
- Replaces American English words with British English equivalents
- Swaps i.e./e.g. with ie/eg
- Replaces other words with Lancet and IHME standards
- Replaces decimal points with floating decimal points
- Uses en-dashes instead of hyphens in year ranges
- Uses em-dashes instead of hyphens for all other hyphens except in negative numbers
- Highlights uncertainty intervals that seem to straddle zero

VBA_macro_generator.py is a script to update the LancetFormatting code. Because VBA doesn't allow file IO for macros, the dictionary of word replacements needs to be inserted manually; this script saves time and energy by generating it programmatically whenever the user updates the dictionary.

British English.csv is the dictionary for all word replacements. When LancetFormatting finds a word or phrase from the second column in the document, it replaces it with the corresponding entry in the first column.

test bed.docx contains examples of text that should or should not be adjusted by the macro. It is meant to test any updates made to the code.
