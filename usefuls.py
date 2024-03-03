# import tkinter as tk
# from tkinter import ttk
# from tkinter import messagebox

# top = tk.Tk()

# def helloCallBack():
#    messagebox.showinfo( "Hello Python", "Hello World")

# B = ttk.Button(top, text ="Hello", command = helloCallBack)
# B.pack()

# top.mainloop()
# ### useful for informational pop-up boxes


import tkinter as tk
# Object creation for tkinter
parent = tk.Tk()
button = tk.Button(text="QUIT",
                   bd=10,
                   bg="grey",
                   fg="red",
                   command=quit,
                   activeforeground="Orange",
                   activebackground="blue",
                   font="Andalus",
                   height=2,
                   highlightcolor="purple",
                   justify="right",
                   padx=10,
                   pady=10,
                   relief="raised",
                   )
# pack geometry manager for organizing a widget before palcing them into the parent widget.
# possible options "Fill" [X=HORIZONTAL,Y=VERTICAL,BOTH]
#                  "side" [LEFT,RIGHT,TOP,UP]
#                  "expand" [YES,NO]
button.pack(fill=tk.BOTH,side=tk.LEFT,expand=tk.YES)
# kick the program
parent.mainloop()