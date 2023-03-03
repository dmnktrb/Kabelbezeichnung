import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import showinfo

import pandas as pd
import re


#https://www.icelscpa.it/flex/cm/pages/ServeAttachment.php/L/IT/D/1%252Ff%252F4%252FD.fe13caf2d5f7dd45ea47/P/BLOB%3AID%3D4/E/pdf?mode=inline
#https://www.pythontutorial.net/tkinter/tkinter-stringvar/
#https://www.pythontutorial.net/tkinter/tkinter-object-oriented-window/

class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()

        # start of line(^) and end of line($) necessary for exact match
        self.pattern_unel = "^(A|F|FF|R|U)(E|E4|G|4|G7|G9|G10|G16|G17|G18|M|R|R2|S17|S18|T)(D|O|W|X)(AC|C|H|H1|H2|H3|H4|H5|Q)(A|F|N|Q|Z)(E4|G|K|M1|M2|M16|M21|R|R16|R18|T)\s*(100/100V|300V/300V|300/500V|450/700V|0,6/1kV)$"
        self.pattern_cenelec = "^(H|N|A)(01|03|05|07|1)(B|N|N2|Q|G|R|S|V|V2|V3|V4|V5|Z|Z1|Z2)(C|C4)(B|G|N|N2|Q|R|V|V2|V3|V4|V5|Z|Z1|Z2|S|G)(H|H2|H3|H6|H7|H8)(D|E|F|H|K|R|U|Y)(G|X)$"

        self.pattern = f"({self.pattern_unel})|({self.pattern_cenelec})"

        self.df_unel = pd.read_excel("cei35011.xlsx")
        self.df_cenelec = pd.read_excel("cei2027.xlsx")

        self.title = "Leitungsdimensionierung"
        self.cable_name = tk.StringVar()

        # calculate the size and position of the window
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = int(screen_width * 3/4)
        window_height = int(screen_height * 3/4)
        window_x = int((screen_width - window_width) / 2)
        window_y = int((screen_height - window_height) / 2)

        # set the size and position of the window
        self.geometry(f"{window_width}x{window_height}+{window_x}+{window_y}")

        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.columnconfigure(2, weight=1)

        
        self.initialize_widgets()
    
    def initialize_widgets(self) -> None:
        
        padding_header = {'padx': 10, 'pady': 10}
        padding = {'padx': 5, 'pady': 5}
        ttk.Label(self, text="Tool für Kabelbezeichnung", font=("Arial", 30)).grid(column=1, row=0, **padding_header)
        ttk.Label(self, text='Kabelbezeichnung:', font=("Arial", 30)).grid(column=0, row=1, sticky='e', **padding)
    
        style = ttk.Style()
        style.configure('TEntry', font=("Arial", 18), padding=(10, 10, 10, 10))
        name_entry = ttk.Entry(self, textvariable=self.cable_name, width=30, style='TEntry', font=("Arial", 15))
        name_entry.grid(column=1, row=1, **padding)
        name_entry.focus()
    
        style.configure('TButton', font=("Arial", 14))
        submit_button = ttk.Button(self, text='Submit', command=self.show_output, width=10, style='TButton')
        submit_button.grid(column=2, row=1, **padding)
    
        self.output_label = ttk.Label(self, font=("Arial", 14), padding=(10, 10))
        self.output_label.grid(column=0, row=2, columnspan=3, **padding)


    def show_output(self) -> None:
        """
        Trigger function for submit button, shows all the output
        """
        

        if self.isValid():
            self.output_label.config(text=f"{self.cable_name.get().upper()}\n\n{self.write_output()}\n{self.get_substrings()}", font=("Comic Sans MS", 14))
        else:
            self.output_label.config(text="Incorrect input")

    def isValid(self) -> bool:
        """
        Returns bool value for validity
        """

        cable = self.cable_name.get().upper()

        #wrap in () so the whole pattern is concatenated with the OR

        #stops program if wrong input
        if re.match(self.pattern, cable):
            return True
        return False
    

    #def get_pattern()    
    def get_substrings(self):
        """
        Returns list of substrings
        """

        word = self.cable_name.get().upper()
        
        match_unel = re.match(self.pattern_unel, word)
        match_cenelec = re.match(self.pattern_cenelec, word)
       
        #
        if match_unel:
            substrings = match_unel.groups()
        
        elif match_cenelec:
            substrings = match_cenelec.groups()
        else:
            #in case it is invalid
            return

        return substrings

    def get_norm(self):
        """
        returns string of norm
        """
        word = self.cable_name.get().upper()

        #INSERT HERE
        if self.isValid():
            if re.match(self.pattern_cenelec, word):
                return "CEI2027"
            elif re.match(self.pattern_unel, word):
                return "CEI35011"
            
    #get the header values of all columns according to the norm
    def get_headers(self):
        if self.get_norm() == "CEI35011":
            df = pd.read_excel("cei35011.xlsx")
        elif self.get_norm() == "CEI2027":
            df = pd.read_excel("cei2027.xlsx")
        
        
        return list(df.columns)
    
    def get_pairings(self):
        l = self.get_substrings()
        
        #Use this like a dictionary bc dictionary doesnt work with duplicate keys like "A"
        keys = []
        values = []
        #header values of column with the same inices as dict entrys - can access both category and values with same index
        #INSERT HERE

        #ERROR: WHEN PROVIDING WORNG INPUT IT DOESNT CRASH, BUT IT DOES WHEN PROVIDING CORRECT ONE
        if self.get_norm() == "CEI35011":
            df = pd.read_excel("cei35011.xlsx")
        elif self.get_norm() == "CEI2027":
            df = pd.read_excel("cei2027.xlsx")
        else:
            #do nothing bc is invalid
            return

        dic = {}
        for i in range(len(l)):
            substr = l[i]
    
            for value in df.iloc[:, i]:
                if pd.isna(value):
                    continue
                parts = value.split(';')
                if substr == parts[0]:
                    keys.append(parts[0])
                    values.append(parts[1])
                if substr is None:
                    keys.append("None")
                    values.append("None")

        #has to be done in order tu be returned, [0] is keys and [1] is values
        list_of_keys_values = [keys, values]
        return list_of_keys_values
    
    def write_output(self):
        """
        Returns f string with newlines for all bindings + headers
        """
        string = f"Norm: {self.get_norm()}\n\n"
        #pairings and according headers have the same index because missing values get inserted as "None" into dictionary and are not being left out
        headers = self.get_headers()

        processed_headers = [s.strip().capitalize().replace('_', ' ') for s in headers]
        
        keys = self.get_pairings()[0]
        values = self.get_pairings()[1]

        for i in range(len(headers)):
            string += f"{processed_headers[i]}: {keys[i]} - {values[i]} \n"
            continue
        return string       
        
if __name__ == "__main__":
    app = App()
    app.mainloop()
    
