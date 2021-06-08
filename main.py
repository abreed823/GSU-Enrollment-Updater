import random
from tkinter import *
import pygame
from tkinter import filedialog
import nav_pages

root = Tk()
root.title('Breedlove')

# Initialize Pygame Mixer
pygame.mixer.init()


# Called when exit button is pressed
def exit():
    quit()


# Called when program first runs
# Allows file upload and edits given Excel file
def upload():
    # Only accepts Excel files
    file_path = filedialog.askopenfilename(filetypes=(("Excel Files", "*.xlsx"), ))

    # Ensures that user actually selected a file rather than just pressing cancel
    if len(file_path) > 0:
        # Changes label text after file is selected
        main_label.config(text="Please wait while your file is being updated. Do not close this window.\nIn the mean "
                               "time, enjoy this calming music :)")
        # Hides upload button
        main_button.pack_forget()
        root.update()

        # Randomly plays one of five songs
        song_list = ["elevator.mp3", "sweet.mp3", "hipjazz.mp3", "smile.mp3", "sunny.mp3"]
        pygame.mixer.music.load(random.choice(song_list))
        pygame.mixer.music.play(loops=0)

        # Calls the rest of the program that updates the spreadsheet
        page_1 = nav_pages.PageOne()
        page_1.select_term()

        page_2 = nav_pages.PageTwo()
        page_2.filter_classes(file_path)

        # Displayed when update process is finished
        main_label.config(text='File update complete. Open the file on your machine to view the changes.')
        # Changes button from 'upload' to 'exit'
        main_button.config(text='Exit', command=exit)
        main_button.pack()
        root.update()


# Creating and packing label
main_label = Label(root, text="Please upload your Excel file below.")
main_label.pack()

# Creates and packing button button
main_button = Button(root, text="Upload", padx=30, command=upload)
main_button.pack()

root.mainloop()
