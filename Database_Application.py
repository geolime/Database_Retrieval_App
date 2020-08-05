# Widget GUI for integration with internal database infrastructure
# Created by Leif Holmquist
# Date: 2020-06-03

# Check for tkinter wrappers to import
try:
    # Python2
    from Tkinter import *
    from Tkinter import tkFileDialog
except ImportError:
    # Python3
    from tkinter import *
    from tkinter import filedialog

# Import packages
from tkinter import messagebox
from psycopg2 import *
import pandas as pd
import csv
import webbrowser
import os

# Path check for objects to be compliled in an executable
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path) 

# Application definition
class MyApp(Frame):
    def __init__(self, parent=None):
        Frame.__init__(self, parent)
        self.parent = parent
        self.init_window()
        
    def init_window(self):
    
		# Title of Database Selection Window
        self.parent.title("DATABASE QUERY")
        
        # Allowing the widget to take the full space of the root window
        self.pack(fill=BOTH, expand=1)
        
        # Constants for controlling layout of buttons ------
        button_width = 6            
        button_padx = "2m"     
        button_pady = "1m"     
        buttons_frame_padx =  "3m"    
        buttons_frame_pady =  "2m"          
        buttons_frame_ipadx = "3m"    
        buttons_frame_ipady = "1m"   
        
        # Top frame
        self.myContainer = Frame(self.parent) ###
        self.myContainer.pack(expand=YES, fill=BOTH)
        
        # Control frame
        self.control_frame = Frame(self.myContainer) ###
        self.control_frame.pack(side=LEFT, expand=NO,  padx=75, pady=20, ipadx=5, ipady=5) 
        
        # SQL CONNECTION
        global conn
        global cur
        conn = connect("dbname=main_database user=postgres password=****")
        cur = conn.cursor()
        
        # Databases
        databases = [
            ("Apartments"),
            ("Small", "Houses")
            ]
        
        self.database_frame = Frame(self.control_frame, borderwidth=5)
        self.database_frame.pack(side=LEFT, expand=YES, fill=Y, anchor=N)
        Label(self.database_frame, text="\n Select \n DATABASE:").pack()
        
        def databasewindow():
            if btn1.get() == 'Small Houses':
            
                #Main Frames
                window = Toplevel(self.database_frame)
                window.title("Sweden Small House Transactions")
                window.iconbitmap(resource_path("RGB.ico"))       
                
                Container = Frame(window)
                Container.pack(expand=YES, fill=BOTH)
                
				# Definition for opening a local filepath
                def callback(filepath):
                    webbrowser.open_new(filepath)
                
                buttons_frame = Frame(Container)
                buttons_frame.pack(side=TOP, expand=NO, fill=Y, ipadx=5, ipady=5)
                
                send_frame = Frame(buttons_frame)
                send_frame.pack(side=RIGHT, expand=YES, fill=BOTH, ipadx=5, ipady=5)
                
                #Sub-Frames
                municipality_frame = Frame(buttons_frame, borderwidth=5)
                year_frame = Frame(buttons_frame, borderwidth=5)
                select_frame = Frame(buttons_frame, borderwidth=5)
                variable_frame = Frame(buttons_frame, borderwidth=5)
                stats_frame = Frame(buttons_frame, borderwidth=5)
        
                municipality_frame.pack(side=LEFT, expand=YES, anchor=N)        
                year_frame.pack(side=LEFT, expand=YES, anchor=N)
                select_frame.pack(side=LEFT, expand=YES, anchor=N)
                variable_frame.pack(side=LEFT, expand=YES, anchor=N)
                stats_frame.pack(side=LEFT, expand=YES, anchor=N)
        
                Label(municipality_frame, text="Select \n Municipalities(s):").pack()
                Label(year_frame, text="Select transaction \n Year(s):").pack()
                Label(select_frame, text="Select the data \n to be exported:").pack()
                Label(variable_frame, text="Select the variables \n you wish to export:").pack()
                Label(stats_frame, text="Select the statistics \n you wish to export:").pack()
                
                #Database parameters
                table = 'small_houses."Houses"'
                table_var = """'Houses'"""
                column_muni = 'kommun'
                column_year = 'RIGHT(datum, 4)'
                
                #Small House Municipalities
                cur.execute('SELECT DISTINCT (%s) FROM %s ORDER BY %s' %(column_muni, table, column_muni))
                municipality = [item[0] for item in cur.fetchall()]
                print(municipality)
                
                #Small House Transaction Years
                cur.execute('SELECT DISTINCT %s FROM %s ORDER BY %s' %(column_year, table, column_year)) 
                year = [item[0] for item in cur.fetchall()]
                print(year)
                
                #Small House Variables
                cur.execute("""SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = %s""" %(table_var)) 
                variables = [item[0] for item in cur.fetchall()]
                print(variables)
                
                # Checkbar Class for variables
                class Checkbar_unchecked(Frame):
                    def __init__(self, parent=None, picks=[], side=LEFT, anchor=W):
                        Frame.__init__(self, parent)
                        self.vars = []
                        self.texts = []
                        for pick in picks:
                            var = IntVar()
                            chk = Checkbutton(self, text=pick, variable=var)
                            chk.pack(side=TOP, anchor=anchor, expand=YES)
                            self.vars.append(var)
                            self.texts.append(pick)
                            var.set(0)
                            #print(var)
                            print(pick)
                            print(self.texts)
                            #print(self.vars)
                    def state(self):
                        return map((lambda var: var.get()), self.vars)
                    def text(self):
                        return map((lambda var: var.get()), self.texts)
                #####
                
                #Statistic Options
                stats = [
                    ("Avg. Price"),
                    ("Avg. Price per Sq. Meter"),
                    ("Avg. K/T"),
                    ("Total Transactions")
                    ]
                
                #Municipal check
                btn2=StringVar()
                btn2.set("ALL")
                def muni_selected():
                    selected = btn2.get()
                    print(selected)
                    list_municipality.config(state=DISABLED if selected == "ALL" else NORMAL)
                
                comp_muni=[
                    ("ALL"), 
                    ("Specific Municipalitie(s)")
                    ]
                    
                for option, componentm in enumerate(comp_muni):
                    button = Radiobutton(municipality_frame, text=componentm, 
                    indicatoron=1, padx = 20, value=componentm, variable=btn2, command=muni_selected)
                    button.pack(anchor=W)
                #####
                
                #Region List
                def CurSelect(event):
                    muni = btn2.get()
                    yr = btn3.get()
                    
                    if muni == "Specific Municipalitie(s)" and yr == "Specific Year(s)":
                        values = [list_municipality.get(idx) for idx in list_municipality.curselection()]
                        values2 = [list_year.get(idx) for idx in list_year.curselection()]
                        print (', '.join(values))
                        print (', '.join(values2))
                    elif muni == "ALL" and yr == "Specific Year(s)":
                        values2 = [list_year.get(idx) for idx in list_year.curselection()]
                        print (', '.join(values2))
                    elif muni == "Specific Municipalitie(s)" and yr == "ALL":
                        values = [list_municipality.get(idx) for idx in list_municipality.curselection()]
                        print (', '.join(values))
        
                scroll_municipality = Scrollbar(municipality_frame)
                scroll_municipality.pack(side=RIGHT, fill=Y)
        
                list_municipality= Listbox(municipality_frame, selectmode=EXTENDED, exportselection=0)
                list_municipality.bind('<<ListboxSelect>>',CurSelect)
                list_municipality["width"] = 15
                list_municipality.pack()

                for item in municipality:
                    list_municipality.insert(END, item)
        
                list_municipality.config(yscrollcommand=scroll_municipality.set)
                list_municipality.config(state=DISABLED)
                scroll_municipality.config(command=list_municipality.yview)
                #####   
                
                #Year check
                btn3=StringVar()
                btn3.set("ALL")
                def year_selected():
                    selected = btn3.get()
                    print(selected)
                    list_year.config(state=DISABLED if selected == "ALL" else NORMAL)
                
                comp_year=[
                    ("ALL"), 
                    ("Specific Year(s)")
                    ]
                    
                for option, componenty in enumerate(comp_year):
                    button = Radiobutton(year_frame, text=componenty, 
                    indicatoron=1, padx = 20, value=componenty, variable=btn3, command=year_selected)
                    button.pack(anchor=W)
                #####
                
                #Year List              
                scroll_year = Scrollbar(year_frame)
                scroll_year.pack(side=RIGHT, fill=Y)
        
                list_year = Listbox(year_frame, selectmode=EXTENDED, exportselection=0)
                list_year.bind('<<ListboxSelect>>',CurSelect)
                list_year["width"] = button_width
                list_year.pack()
                
                for item in year:
                    list_year.insert(END, item)
        
                list_year.config(yscrollcommand=scroll_year.set)
                list_year.config(state=DISABLED)
                scroll_year.config(command=list_year.yview)             
                ######
                
                #Variables and Stats check
                def allselected():
                    selected = btn4.get()
                    print(selected)

                btn4=StringVar()
                btn4.set("ALL")
                
                components=[
                    ("ALL"), 
                    ("Specific Variables"),
                    ("Statistics")
                    ]
                    
                for option, component in enumerate(components):
                    button = Radiobutton(select_frame, text=component, 
                    indicatoron=1, padx = 20, value=component, variable=btn4, command=allselected)
                    button.pack(anchor=W)
                #####
                    
                # Variable List
                variable_list = Checkbar_unchecked(variable_frame, variables)
                variable_list.pack(side=TOP, fill=X)
                variable_list.config(relief=GROOVE, bd=2)
                
                def variablestates():
                    print(list(variable_list.state()))
                    print(variable_list)
                #Button(variable_frame, text='Peek', command=variablestates).pack(side=RIGHT)   
                #####
                
                #Stats List
                stats_list = Checkbar_unchecked(stats_frame, stats)
                stats_list.pack(side=TOP, fill=X)
                stats_list.config(relief=GROOVE, bd=2)
                #stats_list.config(state=DISABLED)
                
                def statstates():
                        print(list(stats_list.state()))
                #Button(stats_frame, text='Peek', command=statstates).pack(side=RIGHT)  
                #####
                
                # Statistics (ALL or Specific)
                def statcheck():
                    selected = btn5.get()
                    print(selected)

                btn5=StringVar()
                btn5.set("ALL")
                
                stattype=[
                    ("ALL"), 
                    ("Groups")
                    ]
                    
                for option, componentst in enumerate(stattype):
                    button = Radiobutton(stats_frame, text=componentst, 
                    indicatoron=1, padx = 20, value=componentst, variable=btn5, command=statcheck)
                    button.pack(anchor=W)
                #####
                
                #Export to file
                def print_to_file(self, colnames,filename=None):
                    count = 0
                    while count < 1:
                        if filename is None:
                            ftypes = [('Excel Workbook', '.xlsx'), 
                                        ('Comma Seperated Delimited', '.csv'), 
                                        ('All Files', '*')]
                            title='Save Small House Transactions:'
                            initialdir='C:/Users/Desktop'
                            initialfile='Small_House_Transactions_'
                            filename = filedialog.asksaveasfilename(filetypes=ftypes, title=title, initialdir=initialdir, initialfile=initialfile, defaultextension='.xlsx')
                            if filename.endswith('.xlsx'):
                                self.to_excel(filename, sheet_name='Sheet1', header=colnames)
                                count += 1
                            elif filename.endswith('.csv'):
                                self.to_csv(filename, header=colnames, encoding='utf-8-sig')
                                count += 1
                            elif filename is (''):
                                return
                            else:
                                MsgBox = messagebox.askquestion("Invalid Filetype", "This is an invalid filetype, continue anyway?")
                                if MsgBox == 'yes':
                                    #messagebox.showinfo('Return','You will now return to the application screen')
                                    return
                                else:
                                    filename=None
                            print(filename)
                    if not filename: 
                        return
                #####
                
                #SQL Query Text box
                def select():

                    global queryresult
                    queryresult = ''
                    muni = btn2.get()
                    yr = btn3.get()
                    dataex = btn4.get()
                    statbtn = btn5.get()
                    
                    if muni == "ALL" and yr == "ALL" and dataex == "ALL":
                        queryresult = 'SELECT * \nFROM %s' %(table)
                        
                    elif muni == "Specific Municipalitie(s)" and yr == "ALL" and dataex == "ALL":
                        sqlmunicipality = list()
                        countx = 0
                        for i in list_municipality.curselection():
                            selected = list_municipality.get(i)
                            sqlmunicipality.append(selected)
                            countx += 1
                        if countx > 1:
                            tuplemunicipality = tuple(sqlmunicipality)
                            queryresult = 'SELECT * \nFROM %s \nWHERE %s in {}'.format(tuplemunicipality) %(table, column_muni)
                        else:
                            queryresult = """SELECT * \nFROM %s \nWHERE %s in ('""" + """,""".join((str(n) for n in sqlmunicipality)) + """')""" %(table, column_muni)
                        
                    elif muni == "ALL" and yr == "Specific Year(s)" and dataex == "ALL":
                        sqlyear = list()
                        county = 0
                        for i in list_year.curselection():
                            selected = list_year.get(i)
                            sqlyear.append(selected)
                            county += 1
                        if county > 1:
                            tupleyear = tuple(sqlyear)
                            queryresult = 'SELECT * \nFROM %s \nWHERE CAST(%s AS INT) in {}'.format(tupleyear) %(table, column_year)
                        else:
                            queryresult = """SELECT * \nFROM %s \nWHERE CAST(%s AS INT) in ('""" + """,""".join((str(n) for n in sqlyear)) + """')""" %(table, column_year)
                        
                    elif muni == "Specific Municipalitie(s)" and yr == "Specific Year(s)" and dataex == "ALL":
                        sqlmunicipality = list()
                        countx = 0
                        for i in list_municipality.curselection():
                            selected = list_municipality.get(i)
                            sqlmunicipality.append(selected)
                            countx += 1
                        sqlyear = list()
                        county = 0
                        for i in list_year.curselection():
                            selected = list_year.get(i)
                            sqlyear.append(selected)
                            county += 1
                        tuplemunicipality = tuple(sqlmunicipality)
                        tupleyear = tuple(sqlyear)
                        if countx > 1 and county > 1:
                            queryresult = 'SELECT * \nFROM %s \nWHERE CAST(%s AS INT) in {} \nAND %s in {}'.format(tupleyear, tuplemunicipality) %(table, column_year, column_muni)
                        elif countx > 1 and county == 1:
                            queryresult = """SELECT * \nFROM %s \nWHERE CAST(%s AS INT) in ('""" + """,""".join((str(n) for n in sqlyear)) + """') \nAND %s in {}""".format(tuplemunicipality) %(table, column_year, column_muni)
                        elif countx == 1 and county > 1:
                            queryresult = """SELECT * \nFROM %s \nWHERE %s in ('""" + """,""".join((str(n) for n in sqlmunicipality)) + """') \nAND CAST(%s AS INT) in {}""".format(tupleyear) %(table, column_muni, column_year)
                        else:
                            queryresult = """SELECT * \nFROM %s \nWHERE CAST(%s) AS INT) in ('""" + """,""".join((str(n) for n in sqlyear)) + """') \nAND %s in ('""" + """,""".join((str(n) for n in sqlmunicipality)) + """')""" %(table, column_year, column_muni)
                        
                    elif muni == "ALL" and yr == "ALL" and dataex == "Specific Variables":
                        sqlvariables = []
                        count = 0
                        for variable in list(variable_list.state()):
                            if variable == 1:
                                if variable_list.texts[count] == '\ufeffgatuadress':
                                    sqlvariables.append('gatuadress')
                                else:
                                    sqlvariables.append(variable_list.texts[count])
                                    print (variable, variable_list.texts[count])
                                count += 1
                            else:
                                print (variable, variable_list.texts[count])
                                count += 1
                        queryresult = 'SELECT %s \nFROM small_houses."Houses"' %(str(sqlvariables).replace("'", "").strip('[]'))
                        
                    elif muni == "ALL" and yr == "ALL" and dataex == "Statistics":
                        sqlstats = []
                        count = 0
                        nozero = 0
                        for variable in list(stats_list.state()):
                            if variable == 1:
                                if stats_list.texts[count] == 'Avg. Price':
                                    sqlstats.append('ROUND(AVG(CAST(köpesumma AS INT)),2) AS avg_price')
                                elif stats_list.texts[count] == 'Avg. Price per Sq. Meter':
                                    sqlstats.append('ROUND(AVG(CAST(köpesumma AS INT)/CAST(boarea AS INT)),2) AS avg_price_m2')
                                    nozero = 1
                                elif stats_list.texts[count] == 'Avg. K/T':
                                    sqlstats.append('ROUND(AVG(CAST(k_t AS DECIMAL)),2) AS avg_kt')
                                else:
                                    sqlstats.append('COUNT(kommun) AS total_sales')
                                count += 1
                            else:
                                print (variable, stats_list.texts[count])
                                count += 1
                        if statbtn == "ALL":
                            if nozero == 1:
                                queryresult = 'SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 AND CAST(boarea AS INT) != 0' %(str(sqlstats).replace("'", "").strip('[]'))
                            else:
                                queryresult = 'SELECT %s \nFROM small_houses."Houses"' %(str(sqlstats).replace("'", "").strip('[]'))
                        else:
                            if nozero == 1:
                                queryresult = 'SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 AND CAST(boarea AS INT) != 0 \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)' %(str(sqlstats).replace("'", "").strip('[]'))
                            else:
                                queryresult = 'SELECT kommun, CAST(RIGHT(datum,4) AS INT), %s \nFROM small_houses."Houses" \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)' %(str(sqlstats).replace("'", "").strip('[]'))
                    
                    elif muni == "Specific Municipalitie(s)" and yr == "ALL" and dataex == "Specific Variables":
                        sqlmunicipality = list()
                        countx = 0
                        sqlvariables = []
                        count = 0
                        for variable in list(variable_list.state()):
                            if variable == 1:
                                if variable_list.texts[count] == '\ufeffgatuadress':
                                    sqlvariables.append('gatuadress')
                                else:
                                    sqlvariables.append(variable_list.texts[count])
                                    print (variable, variable_list.texts[count])
                                count += 1
                            else:
                                print (variable, variable_list.texts[count])
                                count += 1
                        for i in list_municipality.curselection():
                            selected = list_municipality.get(i)
                            sqlmunicipality.append(selected)
                            countx += 1
                        if countx > 1:
                            tuplemunicipality = tuple(sqlmunicipality)
                            queryresult = 'SELECT %s \nFROM small_houses."Houses" \nWHERE kommun in {}'.format(tuplemunicipality) %(str(sqlvariables).replace("'", "").strip('[]'))
                        else:
                            queryresult = """SELECT %s \nFROM small_houses."Houses" \nWHERE kommun in ('""" %(str(sqlvariables).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """')"""
                    
                    elif muni == "ALL" and yr == "Specific Year(s)" and dataex == "Specific Variables":
                        sqlyear = list()
                        county = 0
                        sqlvariables = []
                        count = 0
                        for variable in list(variable_list.state()):
                            if variable == 1:
                                if variable_list.texts[count] == '\ufeffgatuadress':
                                    sqlvariables.append('gatuadress')
                                else:
                                    sqlvariables.append(variable_list.texts[count])
                                    print (variable, variable_list.texts[count])
                                count += 1
                            else:
                                print (variable, variable_list.texts[count])
                                count += 1
                        for i in list_year.curselection():
                            selected = list_year.get(i)
                            sqlyear.append(selected)
                            county += 1
                        if county > 1:
                            tupleyear = tuple(sqlyear)
                            queryresult = 'SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(RIGHT(datum,4) AS INT) in {}'.format(tupleyear) %(str(sqlvariables).replace("'", "").strip('[]'))
                        else:
                            queryresult = """SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(RIGHT(datum,4) AS INT) in ('""" %(str(sqlvariables).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """')"""
                        
                    elif muni == "Specific Municipalitie(s)" and yr == "Specific Year(s)" and dataex == "Specific Variables":
                        sqlmunicipality = list()
                        countx = 0
                        sqlyear = list()
                        county = 0
                        sqlvariables = []
                        count = 0
                        for variable in list(variable_list.state()):
                            if variable == 1:
                                if variable_list.texts[count] == '\ufeffgatuadress':
                                    sqlvariables.append('gatuadress')
                                else:
                                    sqlvariables.append(variable_list.texts[count])
                                    print (variable, variable_list.texts[count])
                                count += 1
                            else:
                                print (variable, variable_list.texts[count])
                                count += 1
                        for i in list_municipality.curselection():
                            selected = list_municipality.get(i)
                            sqlmunicipality.append(selected)
                            countx += 1
                        for i in list_year.curselection():
                            selected = list_year.get(i)
                            sqlyear.append(selected)
                            county += 1
                        tuplemunicipality = tuple(sqlmunicipality)
                        tupleyear = tuple(sqlyear)
                        if countx > 1 and county > 1:
                            queryresult = 'SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(RIGHT(datum,4) AS INT) in {} \nAND kommun in {}'.format(tupleyear, tuplemunicipality) %(str(sqlvariables).replace("'", "").strip('[]'))
                        elif countx > 1 and county == 1:
                            queryresult = """SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(RIGHT(datum,4) AS INT) in ('""" %(str(sqlvariables).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in {}""".format(tuplemunicipality)
                        elif countx == 1 and county > 1:
                            queryresult = """SELECT %s \nFROM small_houses."Houses" \nWHERE kommun in ('""" %(str(sqlvariables).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """') \nAND CAST(RIGHT(datum,4) AS INT) in {}""".format(tupleyear)
                        else:
                            queryresult = """SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(RIGHT(datum,4) AS INT) in ('""" %(str(sqlvariables).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in ('""" + """,""".join((str(n) for n in sqlmunicipality)) + """')"""
                    
                    elif muni == "Specific Municipalitie(s)" and yr == "ALL" and dataex == "Statistics":
                        sqlmunicipality = list()
                        countx = 0
                        sqlstats = []
                        count = 0
                        nozero = 0
                        for variable in list(stats_list.state()):
                            if variable == 1:
                                if stats_list.texts[count] == 'Avg. Price':
                                    sqlstats.append('ROUND(AVG(CAST(köpesumma AS INT)),2) AS avg_price')
                                elif stats_list.texts[count] == 'Avg. Price per Sq. Meter':
                                    sqlstats.append('ROUND(AVG(CAST(köpesumma AS INT)/CAST(boarea AS INT)),2) AS avg_price_m2')
                                    nozero = 1
                                elif stats_list.texts[count] == 'Avg. K/T':
                                    sqlstats.append('ROUND(AVG(CAST(k_t AS DECIMAL)),2) AS avg_kt')
                                else:
                                    sqlstats.append('COUNT(kommun) AS total_sales')
                                count += 1
                            else:
                                print (variable, stats_list.texts[count])
                                count += 1
                        for i in list_municipality.curselection():
                            selected = list_municipality.get(i)
                            sqlmunicipality.append(selected)
                            countx += 1
                        if statbtn == "ALL":
                            if countx > 1:
                                tuplemunicipality = tuple(sqlmunicipality)
                                if nozero == 1:
                                    queryresult = 'SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 \nAND CAST(boarea AS INT) != 0 \nAND kommun in {}'.format(tuplemunicipality) %(str(sqlstats).replace("'", "").strip('[]'))
                                else:
                                    queryresult = 'SELECT %s \nFROM small_houses."Houses" \nWHERE kommun in {}'.format(tuplemunicipality) %(str(sqlstats).replace("'", "").strip('[]'))
                            else:
                                if nozero == 1:
                                    queryresult = """SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 \nAND CAST(boarea AS INT) != 0 \nAND kommun in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """')"""
                                else:
                                    queryresult = """SELECT %s \nFROM small_houses."Houses" \nWHERE kommun in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """')"""
                        else:
                            if countx > 1:
                                tuplemunicipality = tuple(sqlmunicipality)
                                if nozero == 1:
                                    queryresult = 'SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 \nAND CAST(boarea AS INT) != 0 \nAND kommun in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)'.format(tuplemunicipality) %(str(sqlstats).replace("'", "").strip('[]'))
                                else:
                                    queryresult = 'SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE kommun in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)'.format(tuplemunicipality) %(str(sqlstats).replace("'", "").strip('[]'))
                            else:
                                if nozero == 1:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 \nAND CAST(boarea AS INT) != 0 \nAND kommun in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """') \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)"""
                                else:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE kommun in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """') \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)"""

                    elif muni == "ALL" and yr == "Specific Year(s)" and dataex == "Statistics":
                        sqlyear = list()
                        county = 0
                        sqlstats = []
                        count = 0
                        nozero = 0
                        for variable in list(stats_list.state()):
                            if variable == 1:
                                if stats_list.texts[count] == 'Avg. Price':
                                    sqlstats.append('ROUND(AVG(CAST(köpesumma AS INT)),2) AS avg_price')
                                elif stats_list.texts[count] == 'Avg. Price per Sq. Meter':
                                    sqlstats.append('ROUND(AVG(CAST(köpesumma AS INT)/CAST(boarea AS INT)),2) AS avg_price_m2')
                                    nozero = 1
                                elif stats_list.texts[count] == 'Avg. K/T':
                                    sqlstats.append('ROUND(AVG(CAST(k_t AS DECIMAL)),2) AS avg_kt')
                                else:
                                    sqlstats.append('COUNT(kommun) AS total_sales')
                                count += 1
                            else:
                                print (variable, stats_list.texts[count])
                                count += 1
                        for i in list_year.curselection():
                            selected = list_year.get(i)
                            sqlyear.append(selected)
                            county += 1
                        if statbtn == "ALL":
                            if county > 1:
                                tupleyear = tuple(sqlyear)
                                if nozero == 1:
                                    queryresult = 'SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 \nAND CAST(boarea AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) in {}'.format(tupleyear) %(str(sqlstats).replace("'", "").strip('[]'))
                                else:
                                    queryresult = 'SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(RIGHT(datum,4) AS INT) in {}'.format(tupleyear) %(str(sqlstats).replace("'", "").strip('[]'))
                            else:
                                if nozero == 1:
                                    queryresult = """SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 \nAND CAST(boarea AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """')"""
                                else:
                                    queryresult = """SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(RIGHT(datum,4) AS INT) in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """')"""
                        else:
                            if county > 1:
                                tupleyear = tuple(sqlyear)
                                if nozero == 1:
                                    queryresult = 'SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 \nAND CAST(boarea AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)'.format(tupleyear) %(str(sqlstats).replace("'", "").strip('[]'))
                                else:
                                    queryresult = 'SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE CAST(RIGHT(datum,4) AS INT) in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)'.format(tupleyear) %(str(sqlstats).replace("'", "").strip('[]'))
                            else:
                                if nozero == 1:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 \nAND CAST(boarea AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)"""
                                else:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE CAST(RIGHT(datum,4) AS INT) in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)"""
                        if sqlyear == """''""":
                            queryresult = """''"""
                        
                    elif muni == "Specific Municipalitie(s)" and yr == "Specific Year(s)" and dataex == "Statistics":
                        sqlmunicipality = list()
                        countx = 0
                        sqlyear = list()
                        county = 0
                        sqlstats = []
                        count = 0
                        nozero = 0
                        for variable in list(stats_list.state()):
                            if variable == 1:
                                if stats_list.texts[count] == 'Avg. Price':
                                    sqlstats.append('ROUND(AVG(CAST(köpesumma AS INT)),2) AS avg_price')
                                elif stats_list.texts[count] == 'Avg. Price per Sq. Meter':
                                    sqlstats.append('ROUND(AVG(CAST(köpesumma AS INT)/CAST(boarea AS INT)),2) AS avg_price_m2')
                                    nozero = 1
                                elif stats_list.texts[count] == 'Avg. K/T':
                                    sqlstats.append('ROUND(AVG(CAST(k_t AS DECIMAL)),2) AS avg_kt')
                                else:
                                    sqlstats.append('COUNT(kommun) AS total_sales')
                                count += 1
                            else:
                                print (variable, stats_list.texts[count])
                                count += 1
                        for i in list_municipality.curselection():
                            selected = list_municipality.get(i)
                            sqlmunicipality.append(selected)
                            countx += 1
                        for i in list_year.curselection():
                            selected = list_year.get(i)
                            sqlyear.append(selected)
                            county += 1
                        tuplemunicipality = tuple(sqlmunicipality)
                        tupleyear = tuple(sqlyear)
                        if statbtn == "ALL":
                            if countx > 1 and county > 1:
                                if nozero == 1:
                                    queryresult = 'SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 \nAND CAST(boarea AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) in {} \nAND kommun in {}'.format(tupleyear, tuplemunicipality) %(str(sqlstats).replace("'", "").strip('[]'))
                                else:
                                    queryresult = 'SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(RIGHT(datum,4) AS INT) in {} \nAND kommun in {}'.format(tupleyear, tuplemunicipality) %(str(sqlstats).replace("'", "").strip('[]'))
                            elif countx > 1 and county == 1:
                                if nozero == 1:
                                    queryresult = """SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 \nAND CAST(boarea AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in {}""".format(tuplemunicipality)
                                else:
                                    queryresult = """SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(RIGHT(datum,4) AS INT) in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in {}""".format(tuplemunicipality)
                            elif countx == 1 and county > 1:
                                if nozero == 1:
                                    queryresult = """SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 \nAND CAST(boarea AS INT) != 0 \nAND kommun in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """') \nAND CAST(RIGHT(datum,4) AS INT) in {}""".format(tupleyear)
                                else:
                                    queryresult = """SELECT %s \nFROM small_houses."Houses" \nWHERE kommun in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """') \nAND CAST(RIGHT(datum,4) AS INT) in {}""".format(tupleyear)
                            else:
                                if nozero == 1:
                                    queryresult = """SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 \nAND CAST(boarea AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in ('""" + """,""".join((str(n) for n in sqlmunicipality)) + """')"""
                                else:
                                    queryresult = """SELECT %s \nFROM small_houses."Houses" \nWHERE CAST(RIGHT(datum,4) AS INT) in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in ('""" + """,""".join((str(n) for n in sqlmunicipality)) + """')"""
                        else:
                            if countx > 1 and county > 1:
                                if nozero == 1:
                                    queryresult = 'SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 \nAND CAST(boarea AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) in {} \nAND kommun in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)'.format(tupleyear, tuplemunicipality) %(str(sqlstats).replace("'", "").strip('[]'))
                                else:
                                    queryresult = 'SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE CAST(RIGHT(datum,4) AS INT) in {} \nAND kommun in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)'.format(tupleyear, tuplemunicipality) %(str(sqlstats).replace("'", "").strip('[]'))
                            elif countx > 1 and county == 1:
                                if nozero == 1:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 \nAND CAST(boarea AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)""".format(tuplemunicipality)
                                else:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE CAST(RIGHT(datum,4) AS INT) in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)""".format(tuplemunicipality)
                            elif countx == 1 and county > 1:
                                if nozero == 1:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 \nAND CAST(boarea AS INT) != 0 \nAND kommun in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """') \nAND CAST(RIGHT(datum,4) AS INT) in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)""".format(tupleyear)
                                else:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE kommun in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """') \nAND CAST(RIGHT(datum,4) AS INT) in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)""".format(tupleyear)
                            else:
                                if nozero == 1:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE CAST(köpesumma AS INT) != 0 \nAND CAST(boarea AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in ('""" + """,""".join((str(n) for n in sqlmunicipality)) + """') \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)"""
                                else:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM small_houses."Houses" \nWHERE CAST(RIGHT(datum,4) AS INT) in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in ('""" + """,""".join((str(n) for n in sqlmunicipality)) + """') \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT)"""
                    
                    conn.rollback()
                    print(queryresult)
                    cur.execute(queryresult)
                    colnames = [desc[0] for desc in cur.description]
                    print(colnames)
                    #Get Data batch
                    while True:
                        df = pd.DataFrame(cur.fetchall())
                        if len(df) == 0:
                            break
                        else:
                            print_to_file(df,colnames)
                ######              
                
                # Send inquiry request button
                sendButtonFrame = Frame(send_frame)
                sendButtonFrame.pack(side=RIGHT, expand=YES, anchor=W)
        
                def question():
                    MsgBox = messagebox.askquestion("Confirm Data Inquiry", "Send this data request for Small House transactions?")
                    if MsgBox == 'yes':
                        select()
                
                sendButton = Button(sendButtonFrame, text="Send", 
                    background="green", command=question,
                    width=button_width,   
                    padx=button_padx,     
                    pady=button_pady      
                    )               
                sendButton.pack(side=RIGHT, anchor=E)
                
                #Close Top Level Window
                def close_top():
                    if messagebox.askokcancel("Quit", "Do you want to quit?"):
                        window.destroy()
                window.protocol("WM_DELETE_WINDOW", close_top)
                #####
                
            elif btn1.get() == 'Apartments':
                
                #Main Frames
                window = Toplevel(self.database_frame)
                window.title("Apartment Transactions")
                window.iconbitmap(resource_path("RGB.ico"))
                
                Container = Frame(window)
                Container.pack(expand=YES, fill=BOTH)
                
                def callback(filepath):
                    webbrowser.open_new(filepath)
                
                buttons_frame = Frame(Container)
                buttons_frame.pack(side=TOP, expand=NO, fill=Y, ipadx=5, ipady=5)
                
                send_frame = Frame(buttons_frame)
                send_frame.pack(side=RIGHT, expand=YES, fill=BOTH, ipadx=5, ipady=5)
                
                #Sub-Frames
                municipality_frame = Frame(buttons_frame, borderwidth=5)
                year_frame = Frame(buttons_frame, borderwidth=5)
                select_frame = Frame(buttons_frame, borderwidth=5)
                variable_frame = Frame(buttons_frame, borderwidth=5)
                stats_frame = Frame(buttons_frame, borderwidth=5)
        
                municipality_frame.pack(side=LEFT, expand=YES, anchor=N)        
                year_frame.pack(side=LEFT, expand=YES, anchor=N)
                select_frame.pack(side=LEFT, expand=YES, anchor=N)
                variable_frame.pack(side=LEFT, expand=YES, anchor=N)
                stats_frame.pack(side=LEFT, expand=YES, anchor=N)
        
                Label(municipality_frame, text="Select \n Municipalities(s):").pack()
                Label(year_frame, text="Select transaction \n Year(s):").pack()
                Label(select_frame, text="Select the data \n to be exported:").pack()
                Label(variable_frame, text="Select the variables \n you wish to export:").pack()
                Label(stats_frame, text="Select the statistics \n you wish to export:").pack()
                
                #Database parameters
                table = 'apartments."Apartments"'
                table_var = """'Apartments'"""
                column_muni = 'kommun'
                column_year = 'RIGHT(datum, 4)'
                
                
                #Apartment Municipalities
                cur.execute('SELECT DISTINCT (%s) FROM %s ORDER BY %s' %(column_muni, table, column_muni))
                municipality = [item[0] for item in cur.fetchall()]
                print(municipality)
                
                #Apartment Transaction Years
                cur.execute('SELECT DISTINCT %s FROM %s ORDER BY %s' %(column_year, table, column_year)) 
                year = [item[0] for item in cur.fetchall()]
                print(year)
                
                #Apartment Variables
                cur.execute("""SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = %s""" %(table_var)) 
                variables = [item[0] for item in cur.fetchall()]
                print(variables)
                
                #Checkbar Class
                class Checkbar_unchecked(Frame):
                    def __init__(self, parent=None, picks=[], side=LEFT, anchor=W):
                        Frame.__init__(self, parent)
                        self.vars = []
                        self.texts = []
                        for pick in picks:
                            var = IntVar()
                            chk = Checkbutton(self, text=pick, variable=var)
                            chk.pack(side=TOP, anchor=anchor, expand=YES)
                            self.vars.append(var)
                            self.texts.append(pick)
                            var.set(0)
                            #print(var)
                            print(pick)
                            print(self.texts)
                            #print(self.vars)
                    def state(self):
                        return map((lambda var: var.get()), self.vars)
                    def text(self):
                        return map((lambda var: var.get()), self.texts)
                #####
                
                #Statistic Options
                stats = [
                    ("Avg. Price"),
                    ("Avg. Price per Sq. Meter"),
                    ("Avg. Rent"),
                    ("Total Transactions")
                    ]
                
                #Municipal check
                btn2=StringVar()
                btn2.set("ALL")
                def muni_selected():
                    selected = btn2.get()
                    print(selected)
                    list_municipality.config(state=DISABLED if selected == "ALL" else NORMAL)
                
                comp_muni=[
                    ("ALL"), 
                    ("Specific Municipalitie(s)")
                    ]
                    
                for option, componentm in enumerate(comp_muni):
                    button = Radiobutton(municipality_frame, text=componentm, 
                    indicatoron=1, padx = 20, value=componentm, variable=btn2, command=muni_selected)
                    button.pack(anchor=W)
                #####
                
                #Region List
                def CurSelect(event):
                    muni = btn2.get()
                    yr = btn3.get()
                    
                    if muni == "Specific Municipalitie(s)" and yr == "Specific Year(s)":
                        values = [list_municipality.get(idx) for idx in list_municipality.curselection()]
                        values2 = [list_year.get(idx) for idx in list_year.curselection()]
                        print (', '.join(values))
                        print (', '.join(values2))
                    elif muni == "ALL" and yr == "Specific Year(s)":
                        values2 = [list_year.get(idx) for idx in list_year.curselection()]
                        print (', '.join(values2))
                    elif muni == "Specific Municipalitie(s)" and yr == "ALL":
                        values = [list_municipality.get(idx) for idx in list_municipality.curselection()]
                        print (', '.join(values))
        
                scroll_municipality = Scrollbar(municipality_frame)
                scroll_municipality.pack(side=RIGHT, fill=Y)
        
                list_municipality= Listbox(municipality_frame, selectmode=EXTENDED, exportselection=0)
                list_municipality.bind('<<ListboxSelect>>',CurSelect)
                list_municipality["width"] = 15
                list_municipality.pack()

                for item in municipality:
                    list_municipality.insert(END, item)
        
                list_municipality.config(yscrollcommand=scroll_municipality.set)
                list_municipality.config(state=DISABLED)
                scroll_municipality.config(command=list_municipality.yview)
                #####   
                
                #Year check
                btn3=StringVar()
                btn3.set("ALL")
                def year_selected():
                    selected = btn3.get()
                    print(selected)
                    list_year.config(state=DISABLED if selected == "ALL" else NORMAL)
                
                comp_year=[
                    ("ALL"), 
                    ("Specific Year(s)")
                    ]
                    
                for option, componenty in enumerate(comp_year):
                    button = Radiobutton(year_frame, text=componenty, 
                    indicatoron=1, padx = 20, value=componenty, variable=btn3, command=year_selected)
                    button.pack(anchor=W)
                #####
                
                #Year List              
                scroll_year = Scrollbar(year_frame)
                scroll_year.pack(side=RIGHT, fill=Y)
        
                list_year = Listbox(year_frame, selectmode=EXTENDED, exportselection=0)
                list_year.bind('<<ListboxSelect>>',CurSelect)
                list_year["width"] = button_width
                list_year.pack()
                
                for item in year:
                    list_year.insert(END, item)
        
                list_year.config(yscrollcommand=scroll_year.set)
                list_year.config(state=DISABLED)
                scroll_year.config(command=list_year.yview)             
                ######
                
                #Variables and Stats check
                def allselected():
                    selected = btn4.get()
                    print(selected)
                    #variable_list.config(state=DISABLED if selected == "ALL" or selected == "Statistics" else NORMAL)
                    #stats_list.config(state=DISABLED if selected == "ALL" or selected == "Specific Variables" else NORMAL)

                btn4=StringVar()
                btn4.set("ALL")
                
                components=[
                    ("ALL"), 
                    ("Specific Variables"),
                    ("Statistics")
                    ]
                    
                for option, component in enumerate(components):
                    button = Radiobutton(select_frame, text=component, 
                    indicatoron=1, padx = 20, value=component, variable=btn4, command=allselected)
                    button.pack(anchor=W)
                #####
                    
                # Variable List
                variable_list = Checkbar_unchecked(variable_frame, variables)
                variable_list.pack(side=TOP, fill=X)
                variable_list.config(relief=GROOVE, bd=2)
                #variable_list.config(state=DISABLED)
                
                def variablestates():
                    print(list(variable_list.state()))
                    print(variable_list)
                #Button(variable_frame, text='Peek', command=variablestates).pack(side=RIGHT)   
                #####
                
                #Stats List
                stats_list = Checkbar_unchecked(stats_frame, stats)
                stats_list.pack(side=TOP, fill=X)
                stats_list.config(relief=GROOVE, bd=2)
                #stats_list.config(state=DISABLED)
                
                def statstates():
                        print(list(stats_list.state()))
                #Button(stats_frame, text='Peek', command=statstates).pack(side=RIGHT)  
                #####
                
                # Statistics (ALL or Specific)
                def statcheck():
                    selected = btn5.get()
                    print(selected)

                btn5=StringVar()
                btn5.set("ALL")
                
                stattype=[
                    ("ALL"), 
                    ("Groups")
                    ]
                    
                for option, componentst in enumerate(stattype):
                    button = Radiobutton(stats_frame, text=componentst, 
                    indicatoron=1, padx = 20, value=componentst, variable=btn5, command=statcheck)
                    button.pack(anchor=W)
                #####
                
                #Export to file
                def print_to_file(self, colnames,filename=None):
                    count = 0
                    while count < 1:
                        if filename is None:
                            ftypes = [('Excel Workbook', '.xlsx'), 
                                        ('Comma Seperated Delimited', '.csv'), 
                                        ('All Files', '*')]
                            title='Save Apartment Transactions:'
                            initialdir='C:/Users/Desktop'
                            initialfile='Apartment_Transactions_'
                            filename = filedialog.asksaveasfilename(filetypes=ftypes, title=title, initialdir=initialdir, initialfile=initialfile, defaultextension='.xlsx')
                            if filename.endswith('.xlsx'):
                                self.to_excel(filename, sheet_name='Sheet1', header=colnames)
                                count += 1                          
                            elif filename.endswith('.csv'):
                                self.to_csv(filename, header=colnames, encoding='utf-8-sig')
                                count += 1
                            elif filename is (''):
                                return
                            else:
                                MsgBox = messagebox.askquestion("Invalid Filetype", "This is an invalid filetype, continue anyway?")
                                if MsgBox == 'yes':
                                    #messagebox.showinfo('Return','You will now return to the application screen')
                                    return
                                else:
                                    filename=None
                            print(filename)
                    if not filename: 
                        return
                #####
                
                #SQL Query Text box
                def select():
                    global queryresult
                    queryresult = ''
                    muni = btn2.get()
                    yr = btn3.get()
                    dataex = btn4.get()
                    statbtn = btn5.get()
                    
                    if muni == "ALL" and yr == "ALL" and dataex == "ALL":
                        queryresult = 'SELECT * \nFROM apartments."Apartments"'
                        
                    elif muni == "Specific Municipalitie(s)" and yr == "ALL" and dataex == "ALL":
                        sqlmunicipality = list()
                        countx = 0
                        for i in list_municipality.curselection():
                            selected = list_municipality.get(i)
                            sqlmunicipality.append(selected)
                            countx += 1
                        if countx > 1:
                            tuplemunicipality = tuple(sqlmunicipality)
                            queryresult = 'SELECT * \nFROM apartments."Apartments" \nWHERE kommun in {}'.format(tuplemunicipality)
                        else:
                            queryresult = """SELECT * \nFROM apartments."Apartments" \nWHERE kommun in ('""" + """,""".join((str(n) for n in sqlmunicipality)) + """')"""
                        
                    elif muni == "ALL" and yr == "Specific Year(s)" and dataex == "ALL":
                        sqlyear = list()
                        county = 0
                        for i in list_year.curselection():
                            selected = list_year.get(i)
                            sqlyear.append(selected)
                            county += 1
                        if county > 1:
                            tupleyear = tuple(sqlyear)
                            queryresult = 'SELECT * \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) in {}'.format(tupleyear)
                        else:
                            queryresult = """SELECT * \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) in ('""" + """,""".join((str(n) for n in sqlyear)) + """')"""
                        
                    elif muni == "Specific Municipalitie(s)" and yr == "Specific Year(s)" and dataex == "ALL":
                        sqlmunicipality = list()
                        countx = 0
                        for i in list_municipality.curselection():
                            selected = list_municipality.get(i)
                            sqlmunicipality.append(selected)
                            countx += 1
                        sqlyear = list()
                        county = 0
                        for i in list_year.curselection():
                            selected = list_year.get(i)
                            sqlyear.append(selected)
                            county += 1
                        tuplemunicipality = tuple(sqlmunicipality)
                        tupleyear = tuple(sqlyear)
                        if countx > 1 and county > 1:
                            queryresult = 'SELECT * \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in {} \nAND kommun in {}'.format(tupleyear, tuplemunicipality)
                        elif countx > 1 and county == 1:
                            queryresult = """SELECT * \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in ('""" + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in {}""".format(tuplemunicipality)
                        elif countx == 1 and county > 1:
                            queryresult = """SELECT * \nFROM apartments."Apartments" \nWHERE kommun in ('""" + """,""".join((str(n) for n in sqlmunicipality)) + """') \nAND CAST(RIGHT(datum,4) AS INT) as year in {}""".format(tupleyear)
                        else:
                            queryresult = """SELECT * \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in ('""" + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in ('""" + """,""".join((str(n) for n in sqlmunicipality)) + """')"""
                        
                    elif muni == "ALL" and yr == "ALL" and dataex == "Specific Variables":
                        sqlvariables = []
                        count = 0
                        for variable in list(variable_list.state()):
                            if variable == 1:
                                if variable_list.texts[count] == '\ufeffgatuadress':
                                    sqlvariables.append('gatuadress')
                                else:
                                    sqlvariables.append(variable_list.texts[count])
                                    print (variable, variable_list.texts[count])
                                count += 1
                            else:
                                print (variable, variable_list.texts[count])
                                count += 1
                        queryresult = 'SELECT %s \nFROM apartments."Apartments"' %(str(sqlvariables).replace("'", "").strip('[]'))
                        
                    elif muni == "ALL" and yr == "ALL" and dataex == "Statistics":
                        sqlstats = []
                        count = 0
                        nozero = 0
                        for variable in list(stats_list.state()):
                            if variable == 1:
                                if stats_list.texts[count] == 'Avg. Price':
                                    sqlstats.append('ROUND(AVG(CAST(köpesumma AS INT)),2) AS avg_price')
                                elif stats_list.texts[count] == 'Avg. Price per Sq. Meter':
                                    sqlstats.append('ROUND(AVG(CAST(kvmpris AS INT)),2) AS avg_price_m2')
                                    nozero = 1
                                elif stats_list.texts[count] == 'Avg. Rent':
                                    sqlstats.append('ROUND(AVG(CAST(avgift AS INT)),2) AS avg_rent')
                                else:
                                    sqlstats.append('COUNT(kommun) AS total_sales')
                                count += 1
                            else:
                                print (variable, stats_list.texts[count])
                                count += 1
                        if statbtn == "ALL":
                            if nozero == 1:
                                queryresult = 'SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0' %(str(sqlstats).replace("'", "").strip('[]'))
                            else:
                                queryresult = 'SELECT %s \nFROM apartments."Apartments"' %(str(sqlstats).replace("'", "").strip('[]'))
                        else:
                            if nozero == 1:
                                queryresult = 'SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year' %(str(sqlstats).replace("'", "").strip('[]'))
                            else:
                                queryresult = 'SELECT kommun, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year' %(str(sqlstats).replace("'", "").strip('[]'))
                    
                    elif muni == "Specific Municipalitie(s)" and yr == "ALL" and dataex == "Specific Variables":
                        sqlmunicipality = list()
                        countx = 0
                        sqlvariables = []
                        count = 0
                        for variable in list(variable_list.state()):
                            if variable == 1:
                                if variable_list.texts[count] == '\ufeffgatuadress':
                                    sqlvariables.append('gatuadress')
                                else:
                                    sqlvariables.append(variable_list.texts[count])
                                    print (variable, variable_list.texts[count])
                                count += 1
                            else:
                                print (variable, variable_list.texts[count])
                                count += 1
                        for i in list_municipality.curselection():
                            selected = list_municipality.get(i)
                            sqlmunicipality.append(selected)
                            countx += 1
                        if countx > 1:
                            tuplemunicipality = tuple(sqlmunicipality)
                            queryresult = 'SELECT %s \nFROM apartments."Apartments" \nWHERE kommun in {}'.format(tuplemunicipality) %(str(sqlvariables).replace("'", "").strip('[]'))
                        else:
                            queryresult = """SELECT %s \nFROM apartments."Apartments" \nWHERE kommun in ('""" %(str(sqlvariables).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """')"""
                    
                    elif muni == "ALL" and yr == "Specific Year(s)" and dataex == "Specific Variables":
                        sqlyear = list()
                        county = 0
                        sqlvariables = []
                        count = 0
                        for variable in list(variable_list.state()):
                            if variable == 1:
                                if variable_list.texts[count] == '\ufeffgatuadress':
                                    sqlvariables.append('gatuadress')
                                else:
                                    sqlvariables.append(variable_list.texts[count])
                                    print (variable, variable_list.texts[count])
                                count += 1
                            else:
                                print (variable, variable_list.texts[count])
                                count += 1
                        for i in list_year.curselection():
                            selected = list_year.get(i)
                            sqlyear.append(selected)
                            county += 1
                        if county > 1:
                            tupleyear = tuple(sqlyear)
                            queryresult = 'SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in {}'.format(tupleyear) %(str(sqlvariables).replace("'", "").strip('[]'))
                        else:
                            queryresult = """SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in ('""" %(str(sqlvariables).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """')"""
                        
                    elif muni == "Specific Municipalitie(s)" and yr == "Specific Year(s)" and dataex == "Specific Variables":
                        sqlmunicipality = list()
                        countx = 0
                        sqlyear = list()
                        county = 0
                        sqlvariables = []
                        count = 0
                        for variable in list(variable_list.state()):
                            if variable == 1:
                                if variable_list.texts[count] == '\ufeffgatuadress':
                                    sqlvariables.append('gatuadress')
                                else:
                                    sqlvariables.append(variable_list.texts[count])
                                    print (variable, variable_list.texts[count])
                                count += 1
                            else:
                                print (variable, variable_list.texts[count])
                                count += 1
                        for i in list_municipality.curselection():
                            selected = list_municipality.get(i)
                            sqlmunicipality.append(selected)
                            countx += 1
                        for i in list_year.curselection():
                            selected = list_year.get(i)
                            sqlyear.append(selected)
                            county += 1
                        tuplemunicipality = tuple(sqlmunicipality)
                        tupleyear = tuple(sqlyear)
                        if countx > 1 and county > 1:
                            queryresult = 'SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in {} \nAND kommun in {}'.format(tupleyear, tuplemunicipality) %(str(sqlvariables).replace("'", "").strip('[]'))
                        elif countx > 1 and county == 1:
                            queryresult = """SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year AS INT) in ('""" %(str(sqlvariables).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in {}""".format(tuplemunicipality)
                        elif countx == 1 and county > 1:
                            queryresult = """SELECT %s \nFROM apartments."Apartments" \nWHERE kommun in ('""" %(str(sqlvariables).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """') \nAND CAST(RIGHT(datum,4) AS INT) as year in {}""".format(tupleyear)
                        else:
                            queryresult = """SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in ('""" %(str(sqlvariables).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in ('""" + """,""".join((str(n) for n in sqlmunicipality)) + """')"""
                    
                    elif muni == "Specific Municipalitie(s)" and yr == "ALL" and dataex == "Statistics":
                        sqlmunicipality = list()
                        countx = 0
                        sqlstats = []
                        count = 0
                        nozero = 0
                        for variable in list(stats_list.state()):
                            if variable == 1:
                                if stats_list.texts[count] == 'Avg. Price':
                                    sqlstats.append('ROUND(AVG(CAST(köpesumma AS INT)),2) AS avg_price')
                                elif stats_list.texts[count] == 'Avg. Price per Sq. Meter':
                                    sqlstats.append('ROUND(AVG(CAST(kvmpris AS INT)),2) AS avg_price_m2')
                                    nozero = 1
                                elif stats_list.texts[count] == 'Avg. Rent':
                                    sqlstats.append('ROUND(AVG(CAST(avgift AS INT)),2) AS avg_rent')
                                else:
                                    sqlstats.append('COUNT(kommun) AS total_sales')
                                count += 1
                            else:
                                print (variable, stats_list.texts[count])
                                count += 1
                        for i in list_municipality.curselection():
                            selected = list_municipality.get(i)
                            sqlmunicipality.append(selected)
                            countx += 1
                        if statbtn == "ALL":
                            if countx > 1:
                                tuplemunicipality = tuple(sqlmunicipality)
                                if nozero == 1:
                                    queryresult = 'SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0 \nAND kommun in {}'.format(tuplemunicipality) %(str(sqlstats).replace("'", "").strip('[]'))
                                else:
                                    queryresult = 'SELECT %s \nFROM apartments."Apartments" \nWHERE kommun in {}'.format(tuplemunicipality) %(str(sqlstats).replace("'", "").strip('[]'))
                            else:
                                if nozero == 1:
                                    queryresult = """SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0 \nAND kommun in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """')"""
                                else:
                                    queryresult = """SELECT %s \nFROM apartments."Apartments" \nWHERE kommun in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """')"""
                        else:
                            if countx > 1:
                                tuplemunicipality = tuple(sqlmunicipality)
                                if nozero == 1:
                                    queryresult = 'SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0 \nAND kommun in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year'.format(tuplemunicipality) %(str(sqlstats).replace("'", "").strip('[]'))
                                else:
                                    queryresult = 'SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE kommun in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year'.format(tuplemunicipality) %(str(sqlstats).replace("'", "").strip('[]'))
                            else:
                                if nozero == 1:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0 \nAND kommun in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """') \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year"""
                                else:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE kommun in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """') \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year"""

                    elif muni == "ALL" and yr == "Specific Year(s)" and dataex == "Statistics":
                        sqlyear = list()
                        county = 0
                        sqlstats = []
                        count = 0
                        nozero = 0
                        for variable in list(stats_list.state()):
                            if variable == 1:
                                if stats_list.texts[count] == 'Avg. Price':
                                    sqlstats.append('ROUND(AVG(CAST(köpesumma AS INT)),2) AS avg_price')
                                elif stats_list.texts[count] == 'Avg. Price per Sq. Meter':
                                    sqlstats.append('ROUND(AVG(CAST(kvmpris AS INT)),2) AS avg_price_m2')
                                    nozero = 1
                                elif stats_list.texts[count] == 'Avg. Rent':
                                    sqlstats.append('ROUND(AVG(CAST(avgift AS INT)),2) AS avg_rent')
                                else:
                                    sqlstats.append('COUNT(kommun) AS total_sales')
                                count += 1
                            else:
                                print (variable, stats_list.texts[count])
                                count += 1
                        for i in list_year.curselection():
                            selected = list_year.get(i)
                            sqlyear.append(selected)
                            county += 1
                        if statbtn == "ALL":
                            if county > 1:
                                tupleyear = tuple(sqlyear)
                                if nozero == 1:
                                    queryresult = 'SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) as year in {}'.format(tupleyear) %(str(sqlstats).replace("'", "").strip('[]'))
                                else:
                                    queryresult = 'SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in {}'.format(tupleyear) %(str(sqlstats).replace("'", "").strip('[]'))
                            else:
                                if nozero == 1:
                                    queryresult = """SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) as year in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """')"""
                                else:
                                    queryresult = """SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """')"""
                        else:
                            if county > 1:
                                tupleyear = tuple(sqlyear)
                                if nozero == 1:
                                    queryresult = 'SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) as year in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year'.format(tupleyear) %(str(sqlstats).replace("'", "").strip('[]'))
                                else:
                                    queryresult = 'SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year'.format(tupleyear) %(str(sqlstats).replace("'", "").strip('[]'))
                            else:
                                if nozero == 1:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) as year in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year"""
                                else:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year"""
                        
                    elif muni == "Specific Municipalitie(s)" and yr == "Specific Year(s)" and dataex == "Statistics":
                        sqlmunicipality = list()
                        countx = 0
                        sqlyear = list()
                        county = 0
                        sqlstats = []
                        count = 0
                        nozero = 0
                        for variable in list(stats_list.state()):
                            if variable == 1:
                                if stats_list.texts[count] == 'Avg. Price':
                                    sqlstats.append('ROUND(AVG(CAST(köpesumma AS INT)),2) AS avg_price')
                                elif stats_list.texts[count] == 'Avg. Price per Sq. Meter':
                                    sqlstats.append('ROUND(AVG(CAST(kvmpris AS INT)),2) AS avg_price_m2')
                                    nozero = 1
                                elif stats_list.texts[count] == 'Avg. Rent':
                                    sqlstats.append('ROUND(AVG(CAST(avgift AS INT)),2) AS avg_rent')
                                else:
                                    sqlstats.append('COUNT(kommun) AS total_sales')
                                count += 1
                            else:
                                print (variable, stats_list.texts[count])
                                count += 1
                        for i in list_municipality.curselection():
                            selected = list_municipality.get(i)
                            sqlmunicipality.append(selected)
                            countx += 1
                        for i in list_year.curselection():
                            selected = list_year.get(i)
                            sqlyear.append(selected)
                            county += 1
                        tuplemunicipality = tuple(sqlmunicipality)
                        tupleyear = tuple(sqlyear)
                        if statbtn == "ALL":
                            if countx > 1 and county > 1:
                                if nozero == 1:
                                    queryresult = 'SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) as year in {} \nAND kommun in {}'.format(tupleyear, tuplemunicipality) %(str(sqlstats).replace("'", "").strip('[]'))
                                else:
                                    queryresult = 'SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in {} \nAND kommun in {}'.format(tupleyear, tuplemunicipality) %(str(sqlstats).replace("'", "").strip('[]'))
                            elif countx > 1 and county == 1:
                                if nozero == 1:
                                    queryresult = """SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) as year in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in {}""".format(tuplemunicipality)
                                else:
                                    queryresult = """SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in {}""".format(tuplemunicipality)
                            elif countx == 1 and county > 1:
                                if nozero == 1:
                                    queryresult = """SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0 \nAND kommun in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """') \nAND CAST(RIGHT(datum,4) AS INT) as year in {}""".format(tupleyear)
                                else:
                                    queryresult = """SELECT %s \nFROM apartments."Apartments" \nWHERE kommun in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """') \nAND CAST(RIGHT(datum,4) AS INT) as year in {}""".format(tupleyear)
                            else:
                                if nozero == 1:
                                    queryresult = """SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) as year in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in ('""" + """,""".join((str(n) for n in sqlmunicipality)) + """')"""
                                else:
                                    queryresult = """SELECT %s \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in ('""" + """,""".join((str(n) for n in sqlmunicipality)) + """')"""
                        else:
                            if countx > 1 and county > 1:
                                if nozero == 1:
                                    queryresult = 'SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) as year in {} \nAND kommun in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year'.format(tupleyear, tuplemunicipality) %(str(sqlstats).replace("'", "").strip('[]'))
                                else:
                                    queryresult = 'SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in {} \nAND kommun in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year'.format(tupleyear, tuplemunicipality) %(str(sqlstats).replace("'", "").strip('[]'))
                            elif countx > 1 and county == 1:
                                if nozero == 1:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) as year in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year""".format(tuplemunicipality)
                                else:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year""".format(tuplemunicipality)
                            elif countx == 1 and county > 1:
                                if nozero == 1:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0 \nAND kommun in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """') \nAND CAST(RIGHT(datum,4) AS INT) as year in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year""".format(tupleyear)
                                else:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE kommun in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlmunicipality)) + """') \nAND CAST(RIGHT(datum,4) AS INT) as year in {} \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year""".format(tupleyear)
                            else:
                                if nozero == 1:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE CAST(kvmpris AS INT) != 0 \nAND CAST(RIGHT(datum,4) AS INT) as year in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in ('""" + """,""".join((str(n) for n in sqlmunicipality)) + """') \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year"""
                                else:
                                    queryresult = """SELECT kommun as municipality, CAST(RIGHT(datum,4) AS INT) as year, %s \nFROM apartments."Apartments" \nWHERE CAST(RIGHT(datum,4) AS INT) as year in ('""" %(str(sqlstats).replace("'", "").strip('[]')) + """,""".join((str(n) for n in sqlyear)) + """') \nAND kommun in ('""" + """,""".join((str(n) for n in sqlmunicipality)) + """') \nGROUP BY kommun, CAST(RIGHT(datum,4) AS INT) as year"""
                    
                    conn.rollback()
                    print(queryresult)
                    cur.execute(queryresult)
                    colnames = [desc[0] for desc in cur.description]
                    #Get Data batch
                    while True:
                        df = pd.DataFrame(cur.fetchall())
                        if len(df) == 0:
                            break
                        else:
                            print_to_file(df,colnames)
                ######              
                
                # Send inquiry request button
                sendButtonFrame = Frame(send_frame)
                sendButtonFrame.pack(side=RIGHT, expand=YES, anchor=W)
        
                def question():
                    MsgBox = messagebox.askquestion("Confirm Data Inquiry", "Send this data request for Apartment transactions?")
                    if MsgBox == 'yes':
                        select()
                
                sendButton = Button(sendButtonFrame, text="Send", 
                    background="green", command=question,
                    width=button_width,   
                    padx=button_padx,     
                    pady=button_pady      
                    )               
                sendButton.pack(side=RIGHT, anchor=E)
                
                #Close Top Level Window
                def close_top():
                    if messagebox.askokcancel("Quit", "Do you want to quit?"):
                        window.destroy()
                window.protocol("WM_DELETE_WINDOW", close_top)
                #####
            
        # Database buttons
        btn1=StringVar()
        for option, databases in enumerate(databases):
            button = Radiobutton(self.control_frame, text=databases, 
            indicatoron=0, width = 20, padx = 20, value=databases, variable=btn1, command=databasewindow)
            button["width"] = button_width
            button.pack(side=TOP)    
        #####   
root = Tk()
root.iconbitmap(resource_path("RGB.ico"))
myapp = MyApp(root)
def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        root.destroy()
        cur.close()
        conn.close()
root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop() 
