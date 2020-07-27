from tkinter import *
import GUI_Second_Window
import User_Interface

def call_GUI1():
    win1 = Toplevel(root)
    User_Interface.main()

def call_GUI2():
    win2 = Toplevel(root)
    GUI_Second_Window.GUI2(win2)
    return

# the first gui owns the root window
if __name__ == "__main__":
    root = Tk()
    root.title('Caller GUI')
    root.minsize(720, 600)
    button_1 = Button(root, text='Call GUI1', width='20', height='20', command=call_GUI1)
    button_1.pack()
    button_2 = Button(root, text='Call GUI2', width='20', height='20', command=call_GUI2)
    button_2.pack()
    root.mainloop()