import os
from tkinter import *
from tkinter import messagebox
from conversion import Converter

c = Converter()

window = Tk()

window.title('Converter')
window.geometry('500x500')

def main_program():

    program_name = Label(window, text='MConversion')
    program_name.config(font=('MS Sans', 25))
    program_name.pack()
    # SAVE FOLDER
    set_save_location_label = Label(window, text='Save Folder')
    save_location_var = StringVar()
    save_location_var.set(c.default_save_path)
    save_location = Entry(window, justify=LEFT, width=50, textvariable=save_location_var)
    set_save_location_label.pack()
    save_location.pack()

    # DB FILE PATH
    set_db_location_label = Label(window, text='DB filepath')
    db_fp = Entry(window, justify=LEFT, width=50)
    set_db_location_label.pack()
    db_fp.pack()

    # DRIVER NAME
    set_driver_name_label = Label(window, text='Driver name')
    default_driver = StringVar()
    default_driver.set(c.driver_name)
    set_driver_name = Entry(window, justify=LEFT, width=50, textvariable=default_driver)
    set_driver_name_label.pack()
    set_driver_name.pack()


    def save_all_settings():
        #Save all settings that was given in the entry fields,
        #uses default values if leave blank.
        allowed_ext = ['.mdb', '.accdb']
        c.save_dest(save_location.get())
        filepath, ext = os.path.splitext(db_fp.get())

        if ext in allowed_ext:
            c.set_dbq_path(db_fp.get())
        else:
            messagebox.showinfo(title='File not supported',\
            message="The database file that you're trying to update is not supported.\nOperation Aborted.\nSupported files(.mdb, .accdb)")
            return None
        c.set_driver_name(set_driver_name.get())
        print(c.default_save_path , '\n', c.dbq_path, '\n', c.driver_name)

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
        c.build_connection()
        c.start_connection()
        user_info = user_info_table_name.get().strip()
        target_row_col = tuple(targets.get().strip().split(','))
        c.get_userinfo(user_info)
        target_table = c.select_targets(target_row_col)
        c.update_database(target_table)
        status_check.set('Done.')

    status_label = Label(window, text='Status')
    status = Label(window, textvariable=status_check)
    status_label.pack()
    status.pack()


    save_settings = Button(window, text='Save', command=save_all_settings)
    save_settings.pack(fill=X)

    update_data = Button(window, text='Update', command=perform_update)
    update_data.pack(fill=X)

    quit = Button(window, text='Exit', command=window.quit)
    quit.pack(fill=X)

main_program()

window.mainloop()
