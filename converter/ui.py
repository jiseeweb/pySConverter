import os
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from conversion import Converter

c = Converter()

window = Tk()

window.title('Converter')
window.geometry('500x550')

def main_program():

    program_name = Label(window, text='pyConversion')
    program_name.config(font=('MS Sans', 25))
    program_name.pack()
    # SAVE FOLDER
    set_save_location_label = Label(window, text='Save Folder')
    save_location_var = StringVar()
    save_location_var.set(c.default_save_path)
    save_location = Entry(window, justify=LEFT, width=50, textvariable=save_location_var)
    set_save_location_label.pack()
    save_location.pack()

    save_location_browse = Button(window, text='Browse...', command=lambda : browse_directory(save_location, filedialog.askdirectory))
    save_location_browse.pack()

    # DB FILE PATH
    set_db_location_label = Label(window, text='DB filepath')
    db_fp = Entry(window, justify=LEFT, width=50)
    set_db_location_label.pack()
    db_fp.pack()

    def browse_directory(entry_field, methodname):
        str_var = StringVar()
        file_dialog = methodname()
        abs_path = os.path.abspath(file_dialog)
        str_var.set(abs_path)
        entry_field.configure(textvariable=str_var)

    db_path_btn = Button(window, text='Browse...', command=lambda : browse_directory(db_fp, filedialog.askopenfilename))
    db_path_btn.pack()


    # DRIVER NAME
    set_driver_name_label = Label(window, text='Driver name')
    default_driver = StringVar()
    default_driver.set(c.driver_name)
    set_driver_name = Entry(window, justify=LEFT, width=50, textvariable=default_driver)
    set_driver_name_label.pack()
    set_driver_name.pack()


    def check_state(entry_state=NORMAL):
        save_location_browse.config(state=entry_state)
        db_path_btn.config(state=entry_state)
        save_location.config(state=entry_state)
        db_fp.config(state=entry_state)
        set_driver_name.config(state=entry_state)


    def save_all_settings():
        #Save all settings that was given in the entry fields,
        #uses default values if leave blank.
        allowed_ext = ['.mdb', '.accdb']
        is_exists = os.path.exists(db_fp.get())
        c.save_dest(save_location.get())
        if is_exists:
            filepath, ext = os.path.splitext(db_fp.get())
            if ext in allowed_ext:
                c.set_dbq_path(db_fp.get())
                check_state(DISABLED)
            else:
                messagebox.showinfo(title='File not supported',\
                message="The database file that you're trying to update is not supported.\nOperation Aborted.\nSupported files(.mdb, .accdb)")
                return None
        else:
            messagebox.showinfo(title='File not supported',\
            message="Database filepath doesn\'t exist.")
            return None
        try:
            c.set_driver_name(set_driver_name.get())
        except FileNotFoundError as err:
            messagebox.showinfo(title='File not Found', message='File or directory not found.\nPlease check your specified paths. again.')
        messagebox.showinfo(title='Saved', message='Successfully saved.')


    user_info_table_name_label = Label(window, text='User Info Table name (Case sensitive)')
    user_info_table_var = StringVar()
    user_info_table_var.set('USERINFO')
    user_info_table_name = Entry(window, justify=LEFT, width=50, textvariable=user_info_table_var)
    user_info_table_name_label.pack()
    user_info_table_name.pack()


    #Note that it should be separated by comma or space
    targets_label = Label(window, text='Column and Table name (Case sensitive)')
    target_var = StringVar()
    target_var.set(f'{c.user_id},{c.default_table_target}')
    targets = Entry(window, justify=LEFT, width=50, textvariable=target_var)
    targets_label.pack()
    targets.pack()

    status_check = StringVar()

    def perform_update():

        status_check.set('Starting update...')
        c.autosave = True
        try:
            c.build_connection()
        except FileNotFoundError as err:
            messagebox.showinfo(title='File not Found', message='File or directory not found.\nPlease check your specified paths. again.')
        c.start_connection()
        user_info = user_info_table_name.get().strip()
        target_row_col = tuple(targets.get().strip().split(','))
        c.get_userinfo(user_info)
        target_table = c.select_targets(target_row_col)
        c.update_database(target_table)
        status_check.set('Done.')
        messagebox.showinfo(title='Success', message='Finished.')

    status_label = Label(window, text='Status')
    status = Label(window, textvariable=status_check)
    status_label.pack()
    status.pack()

    note_label = Label(window, \
    text='''
    Notes:
    1.) USERINFO and TABLE names are case sensitive.
    2.) DB path must be a file ( eg; path/to/mdb/att.mdb )
    ''', justify=LEFT)
    note_label.pack()


    edit_settings = Button(window, text='Edit Settings', command= lambda : check_state())
    edit_settings.pack(fill=X)

    save_settings = Button(window, text='Save', command=save_all_settings)
    save_settings.pack(fill=X)

    update_data = Button(window, text='Update', command=perform_update)
    update_data.pack(fill=X)

    quit = Button(window, text='Exit', command=window.quit)
    quit.pack(fill=X)

main_program()

window.mainloop()
