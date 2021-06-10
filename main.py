import nav_pages
import pygame
import random
from tkinter import *
from tkinter import filedialog
import tkinter.font as font
import urllib.request

root = Tk()
root.title('Breedlove')

# Sets Geometry/centers window on screen
window_width = 900
window_height = 200

position_right = int(root.winfo_screenwidth()/2 - window_width/2)
position_top = int(root.winfo_screenheight()/2 - window_height/2)

root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

# Initialize Pygame Mixer
pygame.mixer.init()


def disable_event():
    pass


# Checks wifi connection before program runs
def wifi_is_connected(url='http://google.com', timeout=3):
    connection = False
    try:
        urllib.request.urlopen(url, timeout=timeout)
        connection = True
        return connection
    except Exception as e:
        print(e)
        return connection


# Called when program first runs
# Allows file upload and edits given Excel file
def upload():
    # Program will not continue unless WiFi connection is established
    if not wifi_is_connected():
        main_label.config(text='Please check your WiFi connection and try again.')
        root.update()
    else:
        # Only accepts Excel files
        file_path = filedialog.askopenfilename(filetypes=(('Excel Files', '*.xlsx'), ))

        # Ensures that user actually selected a file rather than just pressing cancel
        if len(file_path) > 0:
            # Changes label text after file is selected
            main_label.config(text='Please wait while your file is being updated.'
                                   '\n'
                                   '\nIn the mean time, enjoy this snazzy music :)')
            # Hides upload button
            main_button.pack_forget()

            # Adds music credit
            music_credit_label = Label(root, text='Music courtesy of www.bensound.com', bg='#F0FBFD', fg='#581845')
            music_credit_label.pack(side=RIGHT)

            root.update()

            # Randomly plays one of five songs
            song_list = ['./songs/elevator.mp3', './songs/sweet.mp3', './songs/hipjazz.mp3', './songs/smile.mp3',
                         './songs/sunny.mp3']
            pygame.mixer.music.load(random.choice(song_list))
            pygame.mixer.music.play(loops=0)

            root.protocol("WM_DELETE_WINDOW", disable_event)

            # Calls the rest of the program that updates the spreadsheet
            page_1 = nav_pages.PageOne()
            page_1.select_term()

            page_2 = nav_pages.PageTwo()
            page_2.filter_classes(file_path)

            # Displayed when update process is finished
            main_label.config(text='File update complete.'
                                   '\nOpen the file on your machine to view the changes.')
            # Changes button from 'upload' to 'exit'
            main_button.config(text='Exit', command=root.quit)
            music_credit_label.pack_forget()
            main_button.pack()
            music_credit_label.pack(side=RIGHT)
            root.update()


# Creating and packing label
main_label = Label(root, text='Please upload your Excel file below.', bg='#F0FBFD', fg='#581845')
main_label['font'] = font.Font(size=40)
main_label.pack()

# Creates and packing button button
main_button = Button(root, text='Upload', padx=30, fg='#581845', height=1, width=10, command=upload)
main_button['font'] = font.Font(size=50)
main_button.pack(pady=(75, 0))

root.configure(bg='#F0FBFD')
root.mainloop()
