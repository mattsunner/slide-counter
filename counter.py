"""
Slide Counter Tool Script

Author: Matthew Sunner, 2020
"""

# Imports
import os
import re
import zipfile
import tkinter
from tkinter import filedialog, messagebox

# Results Functions - Errors & Results


def show_error(text):
    tkinter.messagebox.showerror('Error', text)


def show_results(text):
    tkinter.messagebox.showinfo("Results", text)

# Counting Function


def count(files):
    decks = {}
    results = ""

    # Iterate through files selected
    for file in files:
        # Check for PowerPoint extension (.pptx)
        if os.path.abspath(file).endswith('.pptx'):
            decks[(os.path.abspath(file))] = 0
        else:
            show_error(
                "The file %s is not a .pptx file and will be ignored." % (file))

    # Iterate through the decks dict
    for deck, count in decks.items():
        try:
            archive = zipfile.ZipFile(deck, 'r')
            contents = archive.namelist()
        except Exception as e:
            show_error("Error reading %s (%s). Count will be 0." %
                       (os.path.basename(deck), e))
        else:
            for fileentry in contents:
                if(re.findall("ppt/slides/slide", fileentry)):
                    decks[deck] += 1

    results += ("Slides\tDeck\n")

    # Iterate through a sorted dictionary and print out the slide count and name of each deck.
    for deck, count in sorted(decks.items()):
        results += ("%s\t%s\n" % (count, os.path.basename(deck)))

    # Summary Checkbox
    if (show_summary.get()):
        results += ("- - - - -\n")
        total = 0
        for count in decks.values():
            total += count
        results += ("%s total slides in %s decks." % (total, len(decks)))

    show_results(results)

# GUI Interaction


# Tkinter interface Creation
app = tkinter.Tk()
app.geometry('275x175')
app.title("SlideCountr")


def clicked():
    t = tkinter.filedialog.askopenfilenames()
    count(t)


# Window Header
header = tkinter.Label(app, text="SlideCountr",
                       fg="black", font=("Arial Bold", 16))
header.pack(side="top", ipady=10)

# Descriptive Text
text = tkinter.Label(
    app, text="Select some .pptx files, and the\n tool will count how many slides\n are contained within them.")
text.pack()

# Summary Checkbox
show_summary = tkinter.BooleanVar()
show_summary.set(True)
summary = tkinter.Checkbutton(app, text="Show summary", var=show_summary)
summary.pack(ipady=10)

# File Picker Button
open_files = tkinter.Button(app, text="Choose decks...", command=clicked)
open_files.pack(fill="x")

# Initialize Tkinter Window
app.mainloop()
