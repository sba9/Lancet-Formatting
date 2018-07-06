# -*- coding: cp1252 -*-
import csv,datetime
filepath = "British English.csv"
now = datetime.datetime.now()
f = csv.reader(open(filepath,"rb"))
raw_output = []
for row in f:
    if row[0] == "UK":
        continue
    raw_output.append(row)
#####################################
le = len(raw_output)
output = []
macro_count = 0
step_size = 1000
start_indices = range(0,le,step_size)
stop_indices = []
for i in range(0,len(start_indices)):
    if start_indices[i] == start_indices[-1]:
        stop_indices.append(le)
    else:
        stop_indices.append(start_indices[i]+999)

output.append("""Option Explicit
'Sam Albertson, 2018-05-25, updated %s\n""" % str(now.date()))
output.append("""Sub LancetFormatting()
Call Macro0
Call Macro1
Call DecimalMacro
Call EnDashMacro
Call EmDashMacro
Call UncertaintyFlagMacro
End Sub
'Swaps American English words with their British English Counterparts.'
'Covers the first 1000 words in the dictionary.'
""")

for i in range(0,len(start_indices)):
    start_index = start_indices[i]
    stop_index = stop_indices[i]
    output.append("Sub Macro"+str(macro_count)+"()\n")
    output.append("    Dim swapWords("+str(stop_index-start_index+1)+", 2) As String\n")
    row_count = 1
    for j in range(start_index,stop_index):
        row = raw_output[j]
        output.append("    swapWords("+str(row_count)+",1) = \"" + row[0] + "\"\n")
        output.append("    swapWords("+str(row_count)+",2) = \"" + row[1] + "\"\n")
        row_count += 1
    output.append("    Dim i As Integer\n")
    output.append("    For i = 1 To "+str(stop_index-start_index+1)+"\n")
    output.append("          Selection.Find.ClearFormatting\n")
    output.append("          Selection.Find.Replacement.ClearFormatting\n")
    output.append("            With Selection.Find\n")
    output.append("            .Text = swapWords(i, 2)\n")
    output.append("            .Replacement.Text = swapWords(i, 1)\n")
    output.append("            .Forward = True\n")
    output.append("            .Wrap = wdFindContinue\n")
    output.append("            .Format = False\n")
    output.append("            .MatchCase = False\n")
    output.append("            .MatchWholeWord = True\n")
    output.append("            .MatchWildcards = False\n")
    output.append("            .MatchSoundsLike = False\n")
    output.append("            .MatchAllWordForms = False\n")
    output.append("       End With\n")
    output.append("       Options.DefaultHighlightColorIndex = wdDarkYellow\n")
    output.append("       Selection.Find.replacement.Highlight = True\n")
    output.append("       Selection.Find.Execute Replace:=wdReplaceAll\n")
    output.append("   Next i\n")
    output.append("End Sub\n")
    output.append("\n")
    macro_count += 1

output.append("""'Swaps out decimal points for floating decimals.'
Sub DecimalMacro()

Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "(<[0-9]@)(\.)([0-9]@)"
.Replacement.Text = "\\1·\\3"
.Forward = True
.Wrap = wdFindContinue
.Format = False
.MatchCase = False
.MatchWholeWord = True
.MatchWildcards = True
.MatchSoundsLike = False
.MatchAllWordForms = False
End With
Options.DefaultHighlightColorIndex = wdDarkYellow
Selection.Find.Replacement.Highlight = True
Selection.Find.Execute Replace:=wdReplaceAll

End Sub

'Swaps out hyphens for en-dashes in year ranges.'
Sub EnDashMacro()

Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "<([12][019][0-9]{2})-([12][019][0-9]{2})>"
.Replacement.Text = "\\1–\\2"
.Forward = True
.Wrap = wdFindContinue
.Format = False
.MatchCase = False
.MatchWholeWord = True
.MatchWildcards = True
.MatchSoundsLike = False
.MatchAllWordForms = False
End With
Options.DefaultHighlightColorIndex = wdDarkYellow
Selection.Find.Replacement.Highlight = True
Selection.Find.Execute Replace:=wdReplaceAll

End Sub

'Swaps out hyphens for em-dashes when not attached to negative numbers or year ranges.'
Sub EmDashMacro()

Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "-([!0-9]@)"
.Replacement.Text = "—\\1"
.Forward = True
.Wrap = wdFindContinue
.Format = False
.MatchCase = False
.MatchWholeWord = True
.MatchWildcards = True
.MatchSoundsLike = False
.MatchAllWordForms = False
End With
Options.DefaultHighlightColorIndex = wdDarkYellow
Selection.Find.Replacement.Highlight = True
Selection.Find.Execute Replace:=wdReplaceAll

End Sub

'Highlights uncertainty intervals that straddle zero'
Sub UncertaintyFlagMacro()

Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "(-[0-9\,\.]@ to [0-9\,\.]@)"
.Replacement.Text = "\\1"
.Forward = True
.Wrap = wdFindContinue
.Format = False
.MatchCase = False
.MatchWholeWord = True
.MatchWildcards = True
.MatchSoundsLike = False
.MatchAllWordForms = False
End With
Options.DefaultHighlightColorIndex = wdRed
Selection.Find.Replacement.Highlight = True
Selection.Find.Execute Replace:=wdReplaceAll

End Sub""")

#####################################
f_out = open("VBA_formatted_words.txt","w")
f_out.writelines(output)
f_out.close()
