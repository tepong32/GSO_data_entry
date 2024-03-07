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
                   fg="green",
                   command=quit,
                   activeforeground="Orange",
                   activebackground="Purple",
                   font="Roboto",
                   height=2,
                   highlightcolor="Red",
                   justify="right",
                   padx=10,
                   pady=10,
                   relief="groove",
                   )
# pack geometry manager for organizing a widget before placing them into the parent widget.
# possible options "Fill" [X=HORIZONTAL,Y=VERTICAL,BOTH]
#                  "side" [LEFT,RIGHT,TOP,UP]
#                  "expand" [YES,NO]
button.pack(fill=tk.BOTH,side=tk.LEFT,expand=tk.YES)
# kick the program
parent.mainloop()



#### running functions in sequence using lambda
# define the functions
# def fun1():
#     print("Function 1")

# def fun2():
#     print("Function 2")

# bind the functions to the "command" parameter thru a list
# button = ttk.Button(root, text="Save", command=lambda: [fun1(), fun2()])