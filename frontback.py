import os, sys
from tkinter import *
from tkinter import messagebox as mBox

root = Tk()
root.title('cebbank robot')


def _quit():   
	root.quit()   
	root.destroy()
	exit() 

#read GUI
def _check():
	os.system("python RobotLocXpath.py")
		
Button(root,text='Start',width=10,command=_check).grid(row=8,column=0,sticky=W,padx=20,pady=5)
Button(root,text='Exit',width=10,command=_quit).grid(row=8,column=4,sticky=E,padx=20,pady=5)

mainloop()

