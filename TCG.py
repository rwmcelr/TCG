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

def set_system(sys):
  sys = clicked.get()
  if sys == 'Template 2':
    pop = tk.Toplevel()
    pop.lift()
    pop.geometry('525x400')
    pop.wm_title('Preliminary Steps for Other Simulator-based Test Case')
    
