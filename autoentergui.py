# CONVERT GOOGLE SHEET TO CSV WITH GOOGLE SCRIPT
# https://script.google.com/u/0/home/my?pli=1
# ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

from tkinter import *
import pandas as pd
import pyautogui as pg
import time, datetime

# Update the tk day variable with the new day
def next_day():
    day.set(day.get() + 1)
    print("Day ", day.get())
    l2.config(text=str(day.get()))

def back_day():
    day.set(day.get() - 1)
    print("Day ", day.get())
    l2.config(text=str(day.get()))

# Uses the coordinates from get_pos() to set the date
# The coordinates are used to account for potentially different window size and scale
def set_date():
    pg.click(m_positions[0][0], m_positions[0][1])
    pg.press("left")
    pg.press("right", presses=3)
    pg.press("delete", presses=7)
    pg.write(str(day.get()) + "/" + str(datetime.date.today().year))
    pg.click(m_positions[1][0], m_positions[1][1])

# Type out all the numbers for the bank deposits
def type_bank_deposit():
    # Make sure mouse positions for the date and first entry boxes exist
    if m_positions[0] == 0 or m_positions[0] == 0:
        return
    set_date()

    # Update the path for the csv of the current day
    file.set("qbcsv/0(" + str(day.get()) + ").csv")

    # Load the csv for the current day
    c = pd.read_csv(file.get(), usecols = [5, 9])

    # Store all the regular numbers needed
    regulars = [c.iloc[1, 1], c.iloc[2, 1], c.iloc[3, 1], c.iloc[12, 1], c.iloc[13, 1], c.iloc[14, 1], c.iloc[17, 0], (c.iloc[18, 0] - c.iloc[17, 0])]

    # Check for and store any extras (there will only ever be up to 7)
    extras = []
    for i in range(7):
        if not pd.isna(c.iloc[i+4, 1]):
            if c.iloc[i+4, 1] != 0:
                extras.append(-abs(float(c.iloc[i+4, 1])))
    
    # Type out all the numbers
    time.sleep(0.1)
    for r in regulars:
        pg.write(str(round(r, 2)))
        time.sleep(0.1)
        pg.press('down')
    if len(extras) > 0:
        for e in extras:
            pg.write(str(round(e, 2)))
            time.sleep(0.1)
            pg.press('down')

def type_credit_deposit():
    # Get the csv again
    file.set("qbcsv/0(" + str(day.get()) + ").csv")
    c = pd.read_csv(file.get(), usecols = [5, 9])

    # Store the required numbers
    values = [c.iloc[28, 1], c.iloc[29, 1], c.iloc[30, 1]]

    # Type the values
    for n in range(0, 2):
        time.sleep(1)
        print(n)
    print("starting")
    for v in values:
        pg.write(str(round(v, 2)))
        time.sleep(.1)
        pg.press('down')

def get_mpos():
    # Send a popup box to show it's working
    pg.alert("Move mouse to the date and hold")

    # Short delay before setting the value. May be a bit short
    # but has worked for me so far
    time.sleep(1)

    # Store the date box position for use later
    x, y = pg.position()
    m_positions[0] = (x, y)

    # Same thing for the first entry box
    pg.alert("Move mouse to first box and hold")
    time.sleep(1)
    x, y = pg.position()
    m_positions[1] = (x, y)

    # print(m_positions)
    # Visual of where the stored positions are to confirm the locations
    # Need to change the timing to make it more useful
    pg.moveTo(m_positions[0][0], m_positions[0][1], duration=0.1)
    pg.moveTo(m_positions[1][0], m_positions[1][1], duration=0.1)
    pg.moveTo(m_positions[0][0], m_positions[0][1], duration=0.1)
    pg.moveTo(m_positions[1][0], m_positions[1][1], duration=0.1)
    pg.alert("Positions collected, if they're off you can try again")

# Tkinter setup
root = Tk()
root.geometry("200x60")
root.title("Autoenter")
root.configure(bg='#22252b')
root.call('wm', 'attributes', '.', '-topmost', '1')

# Some tkinter variable setup. Needed to make things work with tk buttons
day = IntVar()
day.set(1)
file = StringVar()
file.set("qbcsv/0(" + str(day.get()) + ").csv")
m_positions = [0, 0]

# Set up buttons
b_next = Button(root, text="Next", bg='#2d3138', fg='#ccd5e6', 
                command=next_day, borderwidth = 0, highlightthickness = 0, padx=6, pady=6)
b_back = Button(root, text="Back", bg='#2d3138', fg='#ccd5e6', 
                command=back_day, borderwidth = 0, highlightthickness = 0, padx=6, pady=6)
b_type1 = Button(root, text="One", bg='#2d3138', fg='#ccd5e6', 
                command=type_bank_deposit, borderwidth = 0, highlightthickness = 0, padx=6, pady=6)
b_type2 = Button(root, text="Two", bg='#2d3138', fg='#ccd5e6', 
                command=type_credit_deposit, borderwidth = 0, highlightthickness = 0, padx=6, pady=6)

# This button breaks the layout, but everything still works
b_mpos = Button(root, text="Get mouse positions", bg='#2d3138', fg='#ccd5e6', command=get_mpos, borderwidth = 0, highlightthickness = 0, padx=6, pady=6)

l1 = Label(root, text="Day: ", bg='#2d3138', fg='#ccd5e6')
# Want to change this to something other than a label so I can tpye to set the day
# but as long as I know what day I'm on it's fine for now
l2 = Label(root, text=str(day.get()))

# Add the buttons and labels to a grid layout
b_next.grid(row=0, column=3, padx=6)
b_back.grid(row=0, column=0, padx=6)
l1.grid(row=0, column=1, padx=6)
l2.grid(row=0, column=2, padx=6)
b_type1.grid(row=1, column=1, padx=6)
b_type2.grid(row=1, column=2, padx=6)

# This is what breaks the layout
# Can possibly use frames to separate this and fix it
b_mpos.grid(row=2, column=1, padx=6)

# Start the tkinter window
root.mainloop()