import os
import streamlit as st
import string
import subprocess

def set_env_variable(name, value):
    # Set environment variable for current session
    os.environ[name] = value
    
    # Set environment variable permanently
    subprocess.run(['setx', name, value])

def list_drives():
            drives = []
            for drive in string.ascii_uppercase:
                if os.path.exists(f'{drive}:\\'):
                    drives.append(drive)
            return drives

def list_folders_and_files(path):
    try:
        items = os.listdir(path)
        folders = [f for f in items if os.path.isdir(os.path.join(path, f))]
        files = [f for f in items if os.path.isfile(os.path.join(path, f))]
        return folders, files
    except PermissionError:
        return [], []
       
def list_folders(path):
    try:
        items = os.listdir(path)
        folders = [f for f in items if os.path.isdir(os.path.join(path, f))]
        return folders
    except PermissionError:
        return []

def set_env_variable(name, value):
    # Set environment variable for current session
    os.environ[name] = value
    
    # Set environment variable permanently
    subprocess.run(['setx', name, value])
    
def main():
    # Streamlit app
    st.subheader("Setup App za Doc Koder za Windows")
    st.caption("Ver. 12.06.2024.")
        
    #### default values
    #
    # DOCK_TEMPLATE = './zasisku/mustre/tpl.docx' 
    # DOCK_ZAPOSLENI = './zasisku/zaposleni'
    # DOCK_FIXED = './zasisku/dokumenti'
    # DOCK_PODELA = './zasisku/podela/podela.csv'
    #
    ####

    folder_list = st.selectbox("Odaberi putanju za:", ["Template", "Opste dokumente", "Podelu", "Zaposlene"])
    if folder_list == "Template": # file odakle cita kljuceve, izvorni docx obicno
        folder = "DOCK_TEMPLATE"
        selection_mode = "Folders and Files"
    elif folder_list == "Opste dokumente" : # folder odakle cita opste dokumente, obicno zakone, sistematizaciju i pravilnike
         folder = "DOCK_FIXED"
         selection_mode = "Folders only"
    elif folder_list == "Podelu": # file odakle cita podelu tj listu promena
        folder = "DOCK_PODELA"
        selection_mode = "Folders and Files"
    elif folder_list == "Zaposlene" : # folder odakle cita listu zaposlenih, obicno folderi sa njihovim imenima
         folder = "DOCK_ZAPOSLENI"
         selection_mode = "Folders only"
   
    drives = list_drives()
    drive = st.selectbox("Select Drive", drives)
    current_path = f"{drive}:\\"

    folder_levels = []
    selected_file = None

    while True:
        folders, files = list_folders_and_files(current_path)
        if not folders and (selection_mode == "Folders only" or not files):
            break

        # Select folder or end selection
        selected_folder = st.selectbox(f"Select folder in {current_path} or 'End Selection' to stop:", ["End Selection"] + folders)
        if selected_folder == "End Selection":
            break
        elif selected_folder:
            current_path = os.path.join(current_path, selected_folder)
            folder_levels.append(selected_folder)

    # Show file selection only if "Folders and Files" mode is selected and user finished folder selection
    if selection_mode == "Folders and Files" and selected_folder == "End Selection":
        folders, files = list_folders_and_files(current_path)
        if files:
            selected_file = st.selectbox(f"Select file in {current_path}", [""] + files)
            if selected_file:
                current_path = os.path.join(current_path, selected_file)

    st.write("Selected path:", current_path)
    if selected_file:
        st.write("Selected file:", selected_file)
        st.write("Full file path:", current_path)

    # Set the environment variable
    if st.button("Set Environment Variable"):
        set_env_variable(folder, current_path)
        st.write(f"Environment variable {folder} set to {current_path}")
    
if __name__ == "__main__":
    main()
    
    