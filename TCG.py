import tkinter as tk
from tkinter import filedialog, simpledialog
from PIL import Image, ImageTk
import TCG_format
import TCG_generate

# Window generation and base variable declaration
window = tk.Tk()
window.geometry('700x265')
window.title('TCG')
label = tk.Label(window, text = 'Test Case Generator', font = ('Arial Bold', 20), bg = 'linen').pack()
label = tk.Label(window, text = '(TCG)', font = ('Arial Bold', 15), bg = 'linen').pack()
window.config(bg = 'linen')
system_list, base_name, sim_loc = '','',''
runs = 0
clicked = tk.StringVar(window)
options = ['Template 1', 'Template 2']

# Point to location of system list file (must be .xlsx)
def get_list_location():
  global system_list
  system_list = filedialog.askopenfilename()

  if clicked.get() != 'Template 1':
    TCG_generate.get_elnots(system_list, out_list = True)
  if system_list != "":
    global startLab, loc, locLab
    loc.config(bg = 'spring green')
    locLab.config(text = system_list, bg = 'light gray')
    startLab.destroy()

def set_sim_location():
  global sim_loc
  while sim_loc == "":
    sim_loc = filedialog.askdirectory()
  global SIM_location
  SIM_location.config(bg = 'spring green')

# Allows user to select a name for later file output naming
def get_base_name():
  global base_name
  base_name = simpledialog.askstring('Base Name Input', 'Input a base name for output files, enclose in quotes if name has spaces',
                                     parent = window)

  if base_name != "" and base_name != None:
    global startLab2, nameButton, nameLab
    nameButton.config(bg = 'spring green')
    nameLab.config(text = base_name, bg = 'light gray')
    startLab2.destroy()

# Easter egg function, on press it opens a window with "I have no idea what I'm doing" dog meme
def dog():
  dog_window = tk.Toplevel()
  img = Image.open('easter_egg.jpg')
  image1 = ImageTk.PhotoImage(img)
  ref = tk.Canvas(dog_window)
  ref.image = image1
  dog_label = tk.Label(dog_window, image = image1, borderwidth = 0, highlightthickness = 0)
  dog_label.pack()

def set_template(template):
  template = clicked.get()
  if template == 'Template 2':
    pop = tk.Toplevel()
    pop.lift()
    pop.geometry('525x400')
    pop.wm_title('Preliminary Steps for Other Simulator-based Test Case')
    
    intro = tk.Label(pop, text = 'Currently unable to access restricted simulator library via this program.\n  The following steps need to be taken to generate test case:', font = ('Arial Bold', 12))
    intro.pack()

    step_one = tk.Label(pop, text = '1. Input system list below.  This operation will output a text file\n in the same location as your system list containing all identifiers present.', font = ('Arial', 10), fg = 'Red')
    step_one.pack(pady = 10)

    identifier_list_button = tk.Button(pop, text = 'System List Location', font = ('Arial', 10), command = lambda: [get_list_location(), pop.lift()], background = 'light slate gray', activebackground = 'light slate blue')
    identifier_list_button.pack()

    step_two = tk.Label(pop, text = '2. Instructions to download local copy of simulator library for specified identifier', font = ('Arial', 10), fg = 'Red')
    step_two.pack(pady = 10)

    step_three = tk.Label(pop, text = '3. Use the button below to provide the location of the downloaded simulation files\n Then feel free to click Exit', font = ('Arial', 10), fg = 'Red')
    step_three.pack(pady = 10)

    SIM_location = tk.Button(pop, text = 'Downloaded simulation library location', font = ('Arial', 10), command = lambda: [set_sim_location(), pop.lift()], background = 'light slate gray', activebackground = 'light slate blue')
    SIM_location.pack()

    exit_button = tk.Button(pop, text = 'Exit', font = ('Arial', 10), command = pop.destroy, background = 'light slate gray', activebackground = 'light slate blue')
    exit_button.place(x = 250, y = 350)

    dog_button = tk.Button(pop, text = r'¯\_(ツ)_/¯', font = ('Arial', 6), command = dog, background = 'light slate gray', activebackground = 'light slate blue')
    dog_button.place(x = 480, y = 380)

# Run everything, requires inputs to be completed before running
def run():
  if system_list != "" and base_name != "" and clicked.get() != "":
    global runLab, runs
    list_location = system_list
    name = base_name
    sim_lib = sim_loc
    template = clicked.get()
    test_case, summaries, available = TCG_generate.generate_test_case(list_location, name, sim_lib, template)
    TCG_format.format_test_case(test_case, summaries, available, template)
    if runs > 0:
      runLab = tk.Label(window, text = 'Run complete... Again! (x' + str(runs + 1) + ')', bg = 'spring green', font = ('Arial', 10))
      runLab.place(x = 265, y = 235)
      runs += 1
    else:
      runLab.place(x = 298, y = 235)
      runs += 1

    global loc, locLab, nameButton, nameLab
    loc.config(bg = 'light slate gray')
    nameButton.config(bg = 'light slate gray')
    locLab.config(text = '', width = 55, font = ('Arial', 10), bg = 'dim gray')
    nameLab.config(text = '', width = 50, font = ('Arial', 10), bg = 'dim gray')

  if system_list == '':
    global startLab
    startLab = tk.Label(window, text = 'No System List Location Set', bg = 'red', fg = 'white', font = ('Arial', 10))
    startLab.place(x = 420, y = 204)

  if base_name == '':
    global startLab2
    startLab2 = tk.Label(window, text = 'No Base Name Set', bg = 'red', fg = 'white', font = ('Arial', 10))
    startLab2.place(x = 150, y = 204)

  if clicked.get() == '':
    error_window = tk.Toplevel()
    img = Image.open('system_error.jpg')
    image1 = ImageTk.PhotoImage(img)
    ref = tk.Canvas(error_window)
    ref.image = image1
    error_label = tk.Label(error_window, image = image1, borderwidth = 0, highlightthickness = 0)
    error_label.pack()

drop = tk.OptionMenu(window, clicked, *options, command = set_template)
drop.place(x = 65, y = 6)
dropLab = tk.Label(window, text = 'Template: ', font = ('Arial', 10), bg = 'linen')
dropLab.place(x = 10, y = 10)

loc = tk.Button(window, text = 'System List Location', font = ('Arial', 10), command = get_list_location, background = 'light slate gray', activebackground = 'light slate blue')
loc.place(x = 25, y = 100)
locLab = tk.Label(window, text = '', width = 50, font = ('Arial', 10), bg = 'dim gray')
locLab.place(x = 180, y = 104)

nameButton = tk.Button(window, text = 'Set Output File Base Name', font = ('Arial', 10), command = get_base_name, background = 'light slate gray', activebackground = 'light slate blue')
nameButton.place(x = 10, y = 150)
nameLab = tk.Label(window, text = '', width = 50, font = ('Arial', 10), bg = 'dim gray')
nameLab.place(x = 195, y = 154)

start_button = tk.Button(window, text = 'Generate Test Case', font = ('Arial', 10), command = run, background = 'light slate gray', activebackground = 'light slate blue')
start_button.place(x = 280, y = 200)
startLab = tk.Label(window, text = 'No System List Location Set', bg = 'red', fg = 'white', font = ('Arial', 10))
startLab2 = tk.Label(window, text = 'No Base Name Set', bg = 'red', fg = 'white', font = ('Arial', 10))
runLab = tk.Label(window, text = 'Run complete!', bg = 'spring green', font = ('Arial', 10))

window.mainloop()
