from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import os
import time
import win32com.client as win32
from win32com.client import constants as c
import pythoncom
import pandas as pd
import threading

window_1 = Tk()
window_1.title("Generar archivos PMI separados por AREA")
window_1.resizable(True,True)
window_1.config(background="white")

frame_1=Frame()
frame_1.pack(fill="both",expand="True")
frame_1.config(bg="white")

#---Variables
script_thread = ""
v_script_terminate = False
v_script_completed = False
v_source_dir = ""
v_source_filenameext = ""
v_source_filename = ""
v_source_fileext = ""
v_sourcedir = StringVar()
v_sourcefile = StringVar()
v_targetdir = StringVar()
v_sourcefilenameext = StringVar()
v_sourcedir.set("")
v_sourcefile.set("")
v_targetdir.set("")
v_considerations = "SCRIPT CONSIDERATIONS: For now, this script works for .xlsx files extension and the key cell which will be used to filter its values and split the file must be A9, and the name of worksheet must be \"Resumen por Distrito\""

label_considerations=Message(frame_1, text=v_considerations, bg="white")
label_considerations.config(anchor="w", width=550, font=("TkDefaultFont", 10, 'bold'))
label_considerations.grid(row=0,column=0, padx=2, pady=2, columnspan = 2 )
label_sourcefile=Label(frame_1, text="Archivo fuente:", bg="white")
label_sourcefile.grid(row=5,column=0, sticky="e", padx=2, pady=2)
entry_sourcefile=Entry(frame_1, bg="white smoke", textvariable=v_sourcefile)
entry_sourcefile.config(width=80)
entry_sourcefile.grid(row=5,column=1, padx=2, pady=2)
Button_selsourcefile=Button(frame_1,text="Seleccionar", command=lambda: cmd_selsourcefile())
Button_selsourcefile.grid(row=5,column=2, padx=2, pady=2)
label_sourcedir=Label(frame_1, text="Carpeta fuente:", bg="white")
label_sourcedir.grid(row=6,column=0, sticky="e", padx=2, pady=2)
entry_sourcedir=Label(frame_1, bg="white", textvariable=v_sourcedir)
entry_sourcedir.grid(row=6,column=1, sticky="w", padx=2, pady=2)
label_sourcefilename=Label(frame_1, text="Archivo fuente:", bg="white")
label_sourcefilename.grid(row=7,column=0, sticky="e", padx=2, pady=2)
entry_sourcefilename=Label(frame_1, bg="white", textvariable=v_sourcefilenameext)
entry_sourcefilename.grid(row=7,column=1, sticky="w", padx=2, pady=2)
label_targetdir=Label(frame_1, text="Carpeta destino:", bg="white")
label_targetdir.grid(row=8,column=0, sticky="e", padx=2, pady=2)
entry_targetdir=Entry(frame_1, bg="white smoke", textvariable=v_targetdir)
entry_targetdir.grid(row=8,column=1, padx=2, pady=2)
entry_targetdir.config(width=80)
Button_seltargetdir=Button(frame_1,text="Seleccionar", command=lambda: cmd_seltargetdir())
Button_seltargetdir.grid(row=8,column=2, padx=2, pady=2)
Button_execproc=Button(frame_1,text="Ejecutar proceso", command=lambda: cmd_execproc())
Button_execproc.grid(row=9,column=1, padx=2, pady=2)
label_progressbar=Label(frame_1, text="", bg="white")
label_progressbar.grid(row=10,column=1, sticky="w", padx=2, pady=2, columnspan = 1)
label_progressbar.grid_remove()
progressbar_i = ttk.Progressbar(frame_1, orient="horizontal", length=40, mode="indeterminate")
progressbar_i.grid(row=10,column=0, padx=2, pady=2, sticky='e', columnspan = 1)
progressbar_i.start(5)
progressbar_i.grid_remove()
progressbar = ttk.Progressbar(frame_1, orient="horizontal", length=550, mode="determinate", maximum=23)
progressbar.grid(row=11,column=0, padx=2, pady=2, sticky='nsew', columnspan = 3)
progressbar.grid_remove()
Button_cancel=Button(frame_1,text="Cancelar Script", command=lambda: cmd_cancelproc())
Button_cancel.grid(row=12,column=1, padx=2, pady=2, columnspan = 1)
Button_cancel.grid_remove()

def cmd_selsourcefile():
    v_source_file = filedialog.askopenfilename(initialdir = str(os.getcwd()),filetypes=[("Excel files", "*.xlsx")])
    if os.path.isfile(v_source_file):
        v_source_file = os.path.normpath(v_source_file)
        v_sourcefile.set(v_source_file)
        v_source_dir = os.path.dirname(os.path.abspath(v_source_file))
        v_source_filenameext = os.path.basename(os.path.abspath(v_source_file))
        v_sourcefilenameext.set(v_source_filenameext)
        v_sourcedir.set(v_source_dir)
        v_targetdir.set(v_source_dir)
        
def cmd_seltargetdir():
    v_target_dir = filedialog.askdirectory(initialdir=os.getcwd())
    if os.path.isdir(v_target_dir):
        v_target_dir = os.path.normpath(v_target_dir)
        v_targetdir.set(v_target_dir)

def cmd_cancelproc():
    global v_script_terminate, v_script_completed
    if v_script_completed:  #---Verify flag to only quit
        window_1.destroy()  #---Quit Window program
    else:
        if not v_script_terminate:
            canceling_process_setcontrols() #---Set controls gui while process is CANELING..
            v_script_terminate = True   #---Set the Flag that script thread will verify to calcel
            script_thread.join()    #---Wait here until script thread verify the Flag and has been calceled 
            v_script_completed = True   #---Set this Flag so only quit and does not try cancel script thread again
            canceled_process_setcontrols()  #---Set controls gui while process is already CANCELED

def canceled_process_setcontrols():
    #---Disable, enable and set controls while process is already CANCELED
    window_1.protocol('WM_DELETE_WINDOW')
    Button_cancel["state"] = "normal"
    Button_cancel["text"] = "Terminar"
    Button_cancel.update()
    label_progressbar["text"] = "Script Cancelado"
    label_progressbar.update()
    progressbar_i.grid_remove()
    progressbar_i.update()

def canceling_process_setcontrols():
    #---Disable, enable and set controls while process is CANCELING
    window_1.protocol('WM_DELETE_WINDOW')
    Button_cancel["state"] = "disabled"
    Button_cancel.update()
    label_progressbar["text"] = "Cancelando Script, Espere..."
    label_progressbar.update()
    progressbar_i.grid_remove()
    progressbar_i.update()

def runningprocess_setcontrols():
    #---Disable, enable and set controls while process is running
    window_1.protocol('WM_DELETE_WINDOW', cmd_cancelproc)
    entry_sourcefile["state"] = "disabled"
    Button_selsourcefile["state"] = "disabled"
    entry_targetdir["state"] = "disabled"
    Button_seltargetdir["state"] = "disabled"
    Button_execproc["state"] = "disabled"
    entry_sourcefile.update()
    Button_selsourcefile.update()
    entry_targetdir.update()
    Button_seltargetdir.update()
    Button_execproc.update()
    label_progressbar.grid()
    progressbar_i.grid()
    progressbar.grid()
    Button_cancel.grid()
    label_progressbar.update()
    progressbar_i.update()
    progressbar.update()
    Button_cancel.update()

def finishedprocess_setcontrols():
    window_1.protocol('WM_DELETE_WINDOW') 
    Button_cancel["state"] = "normal"
    Button_cancel["text"] = "Terminar"
    Button_cancel.update()
    label_progressbar["text"] = "Script Completado"
    label_progressbar.update()
    progressbar_i.grid_remove()
    progressbar_i.update()

def runningprocess_inputvalidation():
    #---Imput validation
    #---Elimitate last slash or backslash character if exists in target directory field
    v_target_dir = ""
    v_target_dir = v_targetdir.get()
    v_target_dir = v_target_dir.strip()
    v_target_dir = v_target_dir.rstrip("\\")
    v_target_dir = v_target_dir.rstrip("/")
    v_target_dir = v_target_dir.strip()
    v_targetdir.set(v_target_dir)
    entry_targetdir.update()
    if not os.path.isfile(v_sourcefile.get()):
        messagebox.showerror("Error", "Archivo fuente inválido")
        return False
    if not os.path.isdir(v_targetdir.get()) or v_targetdir.get() == ".":
        messagebox.showerror("Error", "Carpeta destino inválida")
        return False
    return True

def cmd_execproc():
    global script_thread
    if runningprocess_inputvalidation():    #---Validate imput fields
        runningprocess_setcontrols()    #---Set the controls gui while process is runing
        script_thread = threading.Thread(target=script) #---Assign in a separate thread the script process
        script_thread.start()   #---Start the script process

def script():
    global v_script_completed
    p_sourcefilename = v_sourcefilenameext.get()
    p_sourcedir = v_sourcedir.get()
    p_targetdir = v_targetdir.get()   

    label_progressbar["text"] = "Analizando archivo fuente..."
    if not v_script_terminate:    
        label_progressbar.update()

    progressbar["value"] = 1
    if not v_script_terminate:            
        progressbar.update()

    p_sourcedir = p_sourcedir+"\\"
    p_targetdir = p_targetdir+"\\"
    source_nameonly, source_extension = os.path.splitext(p_sourcefilename)
    sourcefile = p_sourcedir+p_sourcefilename

    #---Update progressbar in the gui
    label_progressbar["text"] = "Analizando archivo fuente (Abriendo archivo fuente)..."
    progressbar["value"] = 2
    if not v_script_terminate:
        label_progressbar.update()
    if not v_script_terminate:
        progressbar.update()

    #--- Open source file using pandas library
    df_s = pd.read_excel(sourcefile, sheet_name="Resumen por Distrito", skiprows = range(1,8), header=[1])
    #---create list of values from especified column which will be used to separate source file
    list_of_areas = df_s['Area Central'].unique().tolist()
    #---Eliminate pandas dataframe 
    del df_s

    #--Verifica terminacion del proceso
    if v_script_terminate:
        return
    #---Update progressbar in the gui
    label_progressbar["text"] = "Analizando archivo fuente (El archivo fuente de dividirá en " + str(len(list_of_areas)) + " partes)"
    progressbar["value"] = 4
    if not v_script_terminate:
        label_progressbar.update()
    if not v_script_terminate:
        progressbar.update()
    
    #---OPEN excel instance aplication using win32com library (NOTE:Excel needs to be installed)
    pythoncom.CoInitialize()
    xlApp = win32.gencache.EnsureDispatch("Excel.Application")  #---Generate an instance of Excel App
    xlApp.Visible = 0   #---1 Excel is visible, 0 Excel is not visible
    xlApp.DisplayAlerts = False     #---Sets no display alerts

    #--Verifica terminacion del proceso
    if v_script_terminate:
        xlApp.Application.Quit()
        return
    
    #---For cicle from list of values
    for v_value in list_of_areas:
        targetfile = p_targetdir+source_nameonly+" "+v_value+source_extension

        #---Update progressbar in the gui
        label_progressbar["text"] = "Generando archivo de " + v_value
        progressbar["value"] = progressbar["value"] + 1
        if not v_script_terminate:
            label_progressbar.update()
        if not v_script_terminate:
            progressbar.update()

        #---Open source file and rename it with target file name
        wb1 = xlApp.Workbooks.Open(Filename=sourcefile)
        wb1.SaveAs(Filename=targetfile)
        #---Add an aditional temp worksheet "Sheet3" to work with the data
        wb1.Sheets.Add(After=wb1.Sheets(2))
        #---ws1 will be the original worksheet
        ws1 = wb1.Sheets("Resumen por Distrito")
        #---wsp will be the temp worksheet
        wsp = wb1.Sheets(3)
        #---In original worksheet filter the value with AutoFilter on specified column with "area" value
        ws1.Range('a9:if9').AutoFilter(Field=1, Criteria1=v_value)
        #---With used_range.Row and used_range.Col 
        #---you get the number of row and col in your range
        #---Copy the Range from the cell A1 to the last row/col
        #---11 = xlCellTypeLastCell from VBA Range.SpecialCells Method
        ws1f_used_range = ws1.UsedRange.SpecialCells(11)
        ws1.Range("A1", ws1.Cells(ws1f_used_range.Row, ws1f_used_range.Column)).Copy()
        #---Paste cells copied to temp worksheet
        wsp.Activate()
        wsp.Range('a1').Select()
        wsp.PasteSpecial()
        wsp.Range('a1').Select()
        #---Delete all rows from original worksheet except the header
        ws1.Range('a9:if9').AutoFilter(Field=1)
        ws1_used_range = ws1.UsedRange.SpecialCells(11)
        ws1.Range("A10", ws1.Cells(ws1_used_range.Row, ws1_used_range.Column)).EntireRow.Delete()
        #---Copy data from temp worksheet to original worksheet and delete temp worksheet 
        wsp_used_range = wsp.UsedRange.SpecialCells(11)
        wsp.Range("A10", wsp.Cells(wsp_used_range.Row, wsp_used_range.Column)).Copy()
        ws1.Activate()
        ws1.Range('a10').Select()
        ws1.PasteSpecial()
        ws1.Range('a1').Select()
        wb1.Sheets(3).Delete()
        #---Finaly save changes and close excel file
        wb1.Close(SaveChanges=True)        

        #--Verifica terminacion del proceso
        if v_script_terminate:
            break

    #--Verifica terminacion del proceso
    if v_script_terminate:
        #---Close the excel instance
        xlApp.Application.Quit()
        return

    #---Disable button CANCEL since now on
    window_1.protocol('WM_DELETE_WINDOW') 
    Button_cancel["state"] = "disabled"
    if not v_script_terminate:
        Button_cancel.update()

    #---Generar archivo "Monterrey Total"
    targetfile = p_targetdir+source_nameonly+" "+"Monterrey Total"+source_extension

    #---Update progressbar in the gui
    label_progressbar["text"] = "Generando archivo de Monterrey Total"
    progressbar["value"] = progressbar["value"] + 1
    if not v_script_terminate:
        label_progressbar.update()
    if not v_script_terminate:
        progressbar.update()

    #---Open source file and rename it with target file name
    wb1 = xlApp.Workbooks.Open(Filename=sourcefile)
    wb1.SaveAs(Filename=targetfile)
    #---Add an aditional temp worksheet "Sheet3" to work with the data
    wb1.Sheets.Add(After=wb1.Sheets(2))
    #---ws1 will be the original worksheet
    ws1 = wb1.Sheets("Resumen por Distrito")
    #---wsp will be the temp worksheet
    wsp = wb1.Sheets(3)
    #---Create list needed values
    list_of_values = ["MONTERREY 1","MONTERREY 2","MONTERREY FORANEAS"]
    #---For cicle from list of values
    for v_value in list_of_values:
        #---In original worksheet filter the value with AutoFilter on specified column with "area" value
        ws1.Range('A9:IF9').AutoFilter(Field=1, Criteria1=v_value)
        #---With used_range.Row and used_range.Col 
        #---you get the number of row and col in your range
        #---Copy the Range from the cell A1 to the last row/col
        #---11 = xlCellTypeLastCell from VBA Range.SpecialCells Method
        ws1f_used_range = ws1.UsedRange.SpecialCells(11)
        ws1.Range("A10", ws1.Cells(ws1f_used_range.Row, ws1f_used_range.Column)).Copy()
        #---Paste cells copied to temp worksheet
        wsp.Activate()
        if list_of_values.index(v_value) == 0:
            wsp.Range('A1').Select()
        else:
            wsp_used_range = wsp.UsedRange.SpecialCells(11)
            wsp.Range('A'+str(wsp_used_range.Row+1)).Select()
        wsp.PasteSpecial()
        wsp.Range('A1').Select()
    #---Delete all rows from original worksheet except the header
    ws1.Range('A9:IF9').AutoFilter(Field=1)
    ws1_used_range = ws1.UsedRange.SpecialCells(11)
    ws1.Range("A10", ws1.Cells(ws1_used_range.Row, ws1_used_range.Column)).EntireRow.Delete()
    #---Copy data from temp worksheet to original worksheet and delete temp worksheet 
    wsp_used_range = wsp.UsedRange.SpecialCells(11)
    wsp.Range("A1", wsp.Cells(wsp_used_range.Row, wsp_used_range.Column)).Copy()
    ws1.Activate()
    ws1.Range('A10').Select()
    ws1.PasteSpecial()
    ws1.Range('A1').Select()
    wb1.Sheets(3).Delete()    
    #---Finaly save changes and close excel file
    wb1.Close(SaveChanges=True)        

    #---Close the excel instance    
    xlApp.Application.Quit()

    #---Update progressbar in the gui
    progressbar["value"] = progressbar["maximum"]
    if not v_script_terminate:
        progressbar.update()

    #---Enable Button at proccess finished
    v_script_completed = True
    finishedprocess_setcontrols()

if __name__ == "__main__":
    window_1.mainloop()



