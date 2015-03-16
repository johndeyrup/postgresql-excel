'''
Created on Feb 26, 2015
Utility to work across Navicat PostgresSQL database
and excel spreadsheets for data entry. The program checks
the database for information and allows the user to modify
the data in the database, and add the data to a spreadsheet.
@author: John Deyrup
'''

#Win32com let's us share objects across applications
import win32com.client as win32
#Used to query PostgreSQL Database. If user is open close it
import psycopg2 
from tkinter import *

#Connects to the arilabdb using my user name and password
conn = psycopg2.connect(database="database_name",user="username",password='password',host='server address',port='5432')
 
class Frm(Frame):
    '''
    Creates a frame to store objects; i.e. buttons, text boxes etc
    '''
    def __init__(self, parent):
        '''
        Initiate the frame object
        '''
        Frame.__init__(self, parent, background="light blue")            
        self.parent = parent
        self.parent.title("Casent Checker")
        self.pack(fill=BOTH, expand=1)
        self.centerWindow()
    
    def centerWindow(self):
        '''
        Center a 800 x 600 frame
        '''
        frame_width = 800
        frame_height = 600    
        screen_width = self.parent.winfo_screenwidth()
        screen_height = self.parent.winfo_screenheight()
        frame_xpos = (screen_width - frame_width)/2
        frame_ypos = (screen_height - frame_height)/2
        self.parent.geometry('%dx%d+%d+%d' % (frame_width, frame_height, frame_xpos, frame_ypos))    

class Btn(Button):
    '''
    Create a button
    '''
    def __init__(self, parent, txt, cmd, wid, hei, xpos, ypos):
        '''
        Create a button with the following input parameters;
        text inside the button, command to call method, width, height,
        x-position, y-position
        '''
        Button.__init__(self, parent, text= txt, command=cmd)
        self.pack()
        self.place(width= wid, height=hei, x=xpos, y=ypos)
        
class Ent(Entry):
    '''
    Creates a text box with the input parameters;
    width, height, x-position, and y-position
    '''
    def __init__(self, parent, wid, hei, xpos, ypos):
        Entry.__init__(self, parent)
        self.pack()
        self.place(width = wid, height = hei, x=xpos, y=ypos)
        self.focus_set()
        
class Lbl(Label):
    '''
    Creates a label with specified text, size, and location
    '''
    def __init__(self, parent, txt, wid, hei, xpos, ypos, ft):
        Label.__init__(self, parent, text=txt, font=ft, padx=20)
        self.pack  
        self.place(width = wid, height = hei, x=xpos, y=ypos)     
        

def main():
    '''
    Main method
    '''
    
    root = Tk()
    ex = Frm(root)
    
    def createExcel():
        '''
        Dispatch commands to excel
        '''
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        return excel
    
    def restart():
        '''
        On call destroy the window and recreate it
        '''
        root.destroy()
        start()
        
    def quit_prog():
        '''
        Exit program
        '''
        quit()
        
    #Create quit and restart button and put it into the frame with text, command, width, height, x offset, and yoffset
    quit_btn = Btn(ex, "Quit", quit_prog, 30, 30, 760, 540)
    restart_btn = Btn(ex, "Restart", restart, 45, 30, 700, 540)
    #Binds the key press Enter to get the value in the textbox
    cas_ent = Ent(ex, 100, 20, 130, 50)
    cas_ent.bind("<Return>", lambda e: in_db(cas_ent.get()))
    inst_lbl = Lbl(ex, "Please enter \n a casent #", 100, 50, 20, 50, 16)
    inst_lbl.focus_force()
    
    #Get a worksheet from a given file name 
    def getSheetExcel(file_name):
        #Open the file passed in as arg: file_name
        table = createExcel().Workbooks.Open('C:\\Users\\arilab\\workspace\\ChangeCharacters\\%s' % (file_name))
        ws = table.Worksheets("Sheet1")
        return ws
    
    #Quits and saves excel worksheet
    def quitExcel():
        excel = createExcel()
        excel.DisplayAlerts = False
        excel.ActiveWorkbook.Save()
        excel.Quit()
        
    #Returns the first empty row in a worksheet   
    def getEmptyRows(ws):
        return ws.UsedRange.Rows.Count+1
    
    #Returns the first empty column in a worksheet
    def getEmptyColumns(ws):
        return ws.UsedRange.Columns.Count+1
    
    #Create a list of collection codes that we have recorded with their associated information
    #This will let us compare an entered specimen to previously determined mismatches     
    def create_Mismatch_List():
        ws = getSheetExcel('ngs_mismatch.xlsx')
        rowEnd = getEmptyRows(ws)-1
        columnEnd = getEmptyColumns(ws) - 1
        alist = []
        #Creates a 2d array to store cell values
        for i in range(rowEnd):
            new = []
            for j in range(columnEnd):
                new.append(ws.Cells(i+1,j+1).Value)
            alist.append(new)        
        return alist
    
    
    #Checks mismatch list for collection code; If the collection code is in the mismatch list
    #the program returns the error type row[2] and the error description row[3] 
    def find_mismatch_info(collection_code):
        for row in create_Mismatch_List():
            if(collection_code in row):
                return row[2], row[3]
                break
            else:
                continue
            break
        quitExcel()
        return None
    
    #Update specimen with known specimen code that is missing other data      
    def update_specimen(in_list, casent, collection_code, life_stage, taxon_code):
        destroy_objects(in_list)
        cur = get_cursor()
        cur.execute("UPDATE specimen SET collection_event_code=(%s), basis_of_record=%s, located_at=%s, lifestagesex=%s, medium=%s, taxon_code=%s, owned_by=%s \
        WHERE specimen_code= (%s)", (collection_code, "Preserved specimen", "OIST", life_stage, "Pin", taxon_code, "OIST", casent,));
        conn.commit()
        check_lbl(casent)
                
    #Get a collection code from a FBA code                
    def check_fba(fba_code):
        cur = get_cursor()
        cur.execute("SELECT fba_code, fj_collection_code FROM fba_table")
        fba_list = cur.fetchall()
        for row in fba_list:
            if(row[0] == fba_code):
                return row[1]
            
    #Returns the fields taxon code, life stage, specimen code, collection code, locality name, latitude and longitude for a given specimen  
    def print_fields(casent):
        cur = get_cursor()
        cur.execute("SELECT taxon_code, lifestagesex, specimen_code, collection_event_code FROM specimen WHERE specimen_code= (%s)", (casent,));
        results_list = cur.fetchall()[0]
        results_list = list(results_list)
        cur.execute("SELECT locality_code FROM collection_event WHERE collection_event_code=(%s)", (results_list[3],));
        locality_code = cur.fetchall()[0][0]
        cur.execute("SELECT locality_name, latitude, longitude FROM locality WHERE locality_code=(%s)", (locality_code,));
        for elem in cur.fetchall()[0]:
            results_list.append(elem)
        cur.close()
        return (results_list)
    
    #Destroy/Clear objects from frame
    def destroy_objects(object_list):
        for obj in object_list:
            obj.destroy()
    
    #Create labels of given text and width 
    def create_prompt(text, wid):
        return Lbl(ex, text, wid, 20, 110, 40, 12)
    
    #Inserts input values + constant values e.g. owned_by, located_at etc. into the database
    def insert_into_db(in_list, casent, collection_code, lifestage, taxon_code):
        cur = get_cursor()
        destroy_objects(in_list)
        cur.execute("INSERT INTO specimen (specimen_code, collection_event_code, basis_of_record, located_at, owned_by, lifestagesex, medium, taxon_code) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)", (casent, collection_code, "Preserved specimen", "OIST", "OIST", lifestage, "pin", taxon_code))
        #Commits the change to the database
        conn.commit()
        in_db(casent[6:])
    
    '''    
    Workflow for specimen: If the specimen is not in the database, prompt the user to add it.
    If the specimen code is in the database, but there is no collection code prompt the user to
    add a collection code. If specimen code and collection code are in the database attempt to 
    get locality information.
    '''
    def in_db(casent):
        casent_num = "CASENT" + casent
        if (casent_num not in create_specimen_list()):
            #Clear screen
            cas_ent.destroy()
            inst_lbl.destroy()
            prompt_lbl_list = [Lbl(ex, "Could not find %s in the database please add it" % casent_num, 490, 20, 110, 40, 12), Lbl(ex, "Is there a FBA Code?", 170, 20, 110, 90, 12),
                               Lbl(ex,"Type (Y/N)", 170, 20, 300, 90, 12)]
            ex.focus_set()
            ex.bind("y", lambda e: create_FBA(casent_num, prompt_lbl_list, False))
            ex.bind("n", lambda e: add_all(casent_num, prompt_lbl_list, False))
            
        if (casent_num in create_specimen_list()):
            #Clear screen
            cas_ent.destroy()
            inst_lbl.destroy()
            if (get_collection_code(casent_num) == None):
                no_cc_lbl = Lbl(ex, "No collection code associated with %s:" % casent_num, 380, 20, 110, 40, 12)
                fba_lbl = Lbl(ex, "Is there an FBA Label?", 170, 20, 110, 80, 12)
                fba_lbl_y = Lbl(ex, "Type: (Y/N)", 100, 20, 300, 80, 12)
                out_list = []
                out_list.append(no_cc_lbl)
                out_list.append(fba_lbl)
                out_list.append(fba_lbl_y)
                ex.focus_set()
                ex.bind("y", lambda e: create_FBA(casent_num, out_list, True))
                ex.bind("n", lambda e: add_all(casent_num, out_list, True))
            if (get_collection_code(casent_num) != None):
                check_lbl(casent_num)
    
    #Attempts to add all fields into the database
    def add_all(casent_num, in_list, isUpdate):
        destroy_objects(in_list)
        prompt_ent_list = []
        prompts_list = ["Collection Code:", "Life Stage:", "Taxon Code:"]
        prompt_lbl_list = [Lbl(ex, "Could not find %s in the database please add it" % casent_num, 490, 20, 110, 40, 12), Lbl(ex, "Is there a FBA Code?", 170, 20, 110, 90, 12),
                               Lbl(ex,"N", 170, 20, 300, 90, 12)]
        yoff = 130
        entoff = 130
        for prompt in prompts_list:                
            prompt_lbl_list.append(Lbl(ex, prompt, 170, 20, 110, yoff, 12))
            yoff += 50
        for i in range(3):
            prompt_ent_list.append(Ent(ex, 170, 20, 300, entoff))
            entoff += 50
        out_list = prompt_lbl_list + prompt_ent_list
        if(not isUpdate):
            prompt_ent_list[0].bind("<Return>", lambda e: insert_into_db(out_list, casent_num, prompt_ent_list[0].get(), prompt_ent_list[1].get(), prompt_ent_list[2].get()))
            prompt_ent_list[1].bind("<Return>", lambda e: insert_into_db(out_list, casent_num, prompt_ent_list[0].get(), prompt_ent_list[1].get(), prompt_ent_list[2].get()))
            prompt_ent_list[2].bind("<Return>", lambda e: insert_into_db(out_list, casent_num, prompt_ent_list[0].get(), prompt_ent_list[1].get(), prompt_ent_list[2].get()))    
        if(isUpdate):
            prompt_ent_list[0].bind("<Return>", lambda e: update_specimen(out_list, casent_num, prompt_ent_list[0].get(), prompt_ent_list[1].get(), prompt_ent_list[2].get()))
            prompt_ent_list[1].bind("<Return>", lambda e: update_specimen(out_list, casent_num, prompt_ent_list[0].get(), prompt_ent_list[1].get(), prompt_ent_list[2].get()))
            prompt_ent_list[2].bind("<Return>", lambda e: update_specimen(out_list, casent_num, prompt_ent_list[0].get(), prompt_ent_list[1].get(), prompt_ent_list[2].get()))
    
    #If the collection code is an FBA code enter the FBA code, the program will then figure out the corresponding collection code and add it to the database        
    def create_FBA(casent_num, prompt_lbl_list, isinDbase):
        ex.unbind("y")
        destroy_objects(prompt_lbl_list)
        prompt = create_prompt("Please enter an FBA Code, e.g. FBA123 enter 123 then hit enter once all fields are entered", 670)
        fba_text_list = ["FBA Code:", "Life Stage", "Taxon Code"]
        lbl_obj_list = []
        ent_obj_list = []
        yoff = 80 
        entoff = 80
        for word in fba_text_list:
            lbl_obj_list.append(Lbl(ex, word, 170, 20, 110, yoff, 12))
            yoff += 40
        for i in range(len(fba_text_list)):
            ent_obj_list.append(Ent(ex, 170, 20, 300, entoff))
            entoff += 40
        out_list = lbl_obj_list + ent_obj_list
        out_list.append(prompt)
        if(not isinDbase):
            ent_obj_list[0].bind("<Return>", lambda e: insert_into_db(out_list, casent_num, check_fba("FBA" + str(ent_obj_list[0].get())), ent_obj_list[1].get(), ent_obj_list[2].get()))
            ent_obj_list[1].bind("<Return>", lambda e: insert_into_db(out_list, casent_num, check_fba("FBA" + str(ent_obj_list[0].get())), ent_obj_list[1].get(), ent_obj_list[2].get()))
            ent_obj_list[2].bind("<Return>", lambda e: insert_into_db(out_list, casent_num, check_fba("FBA" + str(ent_obj_list[0].get())), ent_obj_list[1].get(), ent_obj_list[2].get()))
        if(isinDbase):
            ent_obj_list[0].bind("<Return>", lambda e: update_specimen(out_list, casent_num, check_fba("FBA" + str(ent_obj_list[0].get())), ent_obj_list[1].get(), ent_obj_list[2].get()))
            ent_obj_list[1].bind("<Return>", lambda e: update_specimen(out_list, casent_num, check_fba("FBA" + str(ent_obj_list[0].get())), ent_obj_list[1].get(), ent_obj_list[2].get()))
            ent_obj_list[2].bind("<Return>", lambda e: update_specimen(out_list, casent_num, check_fba("FBA" + str(ent_obj_list[0].get())), ent_obj_list[1].get(), ent_obj_list[2].get()))
    
    #Displays the fields to check on the label
    def check_lbl(casent_num):
        yoff_label_t = 40
        yoff_field = 40
        label_types = ["Taxon:", "Life stage/ Sex:", "CASENT#:", "Collection Code:", "Locality Name:", "Latitude", "Longitude", "Information Accurate?"]
        field_results = print_fields(casent_num)
        label_list = []
        field_list = []
        for label_t in label_types:
            label_list.append(Lbl(ex, label_t, 170, 20, 110, yoff_label_t, 12))
            yoff_label_t += 50
        for field in field_results:
            field_list.append(Lbl(ex, field, 450, 20, 300, yoff_field, 12))
            yoff_field += 50
        yes_lbl = Lbl(ex, "Type: (Y/N)", 100, 20, 300, yoff_field, 12)
        ex.focus_set()
        ex.bind("y", lambda e: add_to_extraction(casent_num))
        ex.bind("n", lambda e: check_mismatch(casent_num, field_results[3], label_list, field_list, yes_lbl))
    
    #Looks in the known errors excel spreadsheet for the collection code, if it is there add another copy of it with the corresponding collection code
    #If the collection code is not in the spreadsheet prompt the user to add it                 
    def check_mismatch(casent_num, ccode, label_list, entry_list, yes_lbl):
        result = find_mismatch_info(ccode)
        if(result != None):
            #If there was an error in the database enter the error type and error description, collection code and casent will be automatically added.
            error_type = result[0]
            error_descrp = result[1] 
            ws = getSheetExcel('ngs_mismatch.xlsx')
            emptyRow = getEmptyRows(ws)
            ws.Cells(emptyRow,1).Value = casent_num
            ws.Cells(emptyRow,2).Value = ccode
            #This returns error type
            ws.Cells(emptyRow,3).Value = error_type
            #This returns error description 
            ws.Cells(emptyRow,4).Value = error_descrp
            #Save and quit spreadsheet
            quitExcel()
            restart()
        else:
            for s_labels in label_list:
                s_labels.destroy()
            for ents in entry_list:
                ents.destroy()
            yes_lbl.destroy()
            error_list = ["CASENT#:", "Collection Code:", "Error Type:", "Error Description:"]
            error_lbls = []
            txts_lists = [casent_num, ccode]
            error_txts = []
            y_off = 40
            for errors in error_list:
                error_lbls.append(Lbl(ex, errors, 150, 20, 110, y_off, 12))
                y_off += 50
            ent_off = 40
            for txts in txts_lists:
                error_lbls.append(Lbl(ex, txts, 450, 20, 300, ent_off, 12))
                ent_off += 50
            for i in range(2):
                error_txts.append(Ent(ex, 450, 20, 300, ent_off))
                ent_off += 50
            error_txts[0].bind("<Return>", lambda e: input_errors(casent_num, ccode, error_txts[0].get(), error_txts[1].get()))
            error_txts[1].bind("<Return>", lambda e: input_errors(casent_num, ccode, error_txts[0].get(), error_txts[1].get()))
            def input_errors(casents, ccodes, errortypes, errordescp):
                ws = getSheetExcel('ngs_mismatch.xlsx')
                emptyRow = getEmptyRows(ws)
                ws.Cells(emptyRow,1).Value = casents
                ws.Cells(emptyRow,2).Value = ccodes
                #This returns error type
                ws.Cells(emptyRow,3).Value = errortypes
                #This returns error description 
                ws.Cells(emptyRow,4).Value = errordescp
                #Save and quit spreadsheet
                quitExcel()
                restart()
                
    #Add Casent number to Excel Spreadsheet
    def add_to_extraction(casent):
        ws = getSheetExcel("dna_extraction_list.xlsx")
        emptyRow = getEmptyRows(ws)
        ws.Cells(emptyRow, 2).Value = casent
        quitExcel()
        restart()
    
    #Determine collection code from a given specimen code    
    def get_collection_code(casent):
        cur = get_cursor()
        cur.execute("SELECT specimen_code, collection_event_code FROM specimen")
        new_list = cur.fetchall()
        cur.close
        for row in new_list:
            if(row[0] == casent):
                return row[1]
           
    #Creates a cursor which allows python to execute sql queries
    def get_cursor():
        cur = conn.cursor()
        return cur
    
    #Returns a list of all specimens in the database    
    def create_specimen_list():
        cur = get_cursor()
        #Execute a sql query to get all specimen codes
        cur.execute("SELECT specimen_code FROM specimen")
        #Returns all the results from the query
        check_list = cur.fetchall()
        cur.close()
        new_list = []
        for rows in check_list:
            new_list.append(rows[0])
        return new_list
    
    root.mainloop()
      
def start():
    if __name__ == '__main__':
        main()

start()