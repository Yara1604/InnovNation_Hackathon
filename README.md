# InnovNation_Hackathon

This code is an attempt at integrating multiple libraries in order to be able to read a whole handwritten page and turn it into an editable word document.

So far, the only things we've been able to accomplish in this prototype are:
1. Text extraction (w/ Spell checking to reduce chance of error)
2. Highlighted text extraction
3. Placing the highlighted & regular text in a word document

In order to run this prototype you will need these packages:
easyocr
cv2
numpy
docx  
docx.shared
spellchecker
