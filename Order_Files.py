from tkinter import *
from tkinter import filedialog
from openpyxl import load_workbook
import os
from difflib import SequenceMatcher
import shutil
from tkinter import messagebox

# ******************* FILE CONFIG *****************
# El file de configuracion debe estar en excel 
# una sola columna exam:

#   *Titulo de clase
#    clase_1
#    clase_2 
#   *Titulo de clase2
#    clase_1
#    clase_2 

# los titulos de clase "sub-folders" deben tener un * al inicio del name

# el folder del curso tendra el nombre que tenga el file de config,no importan las extensiones ni mayusculas, 
# el sistema busca coincidencias y reorganiza conforme a ellas; el folder sera creado en el directorio donde se 
# encuentra el file de configuracion


# ***************** FOLDER DIRECTORY *************************
# la carpeta directorio es donde estaran todos los archivos para ser renombrados, el sistema crea una copia de ellos
#no deben estar en sub-carpetas, solo en el folder raiz



#variables globales
directory = ""
file_config = ""
StatusRun = True

#funcion inicial
def run():

    #parametros para interfaz grafica
    Color_Fondo = "gray12"
    Color_Boton="gray8"
    Color_Texto = "white"

    #Inicio de interfaz grafica
    v = Tk()
    v.title("Rename cursos-platzi By JC")
    v.geometry("480x160")
    v.configure(background=Color_Fondo)
    v.resizable(False,False)

    # ********** Directorio de archivos ********

    # TextBox
    directory_files = StringVar()
    txt_directory = Entry(v,font=("arial",10,"bold"),width=65,bg=Color_Fondo,bd=0,fg=Color_Texto,
                textvariable=directory_files, justify=CENTER,state="disabled",disabledbackground="white",
                disabledforeground="black")
    txt_directory.place(x=10,y=10)
    txt_directory.focus()

    #botton
    btn_Directory = Button(v,text="Open directory",font=("arial",12,"bold"),width=13,height=1,bg=Color_Boton,bd=0,fg="white",command = lambda: select_directory(directory_files))
    btn_Directory.place(x=80,y=80)


    # ********** Archivo de configuracion ********

    # Textbox
    config_file = StringVar()
    txt_configFile = Entry(v,font=("arial",10,"bold"),width=65,bg="white",bd=0,fg="black",textvariable=config_file, justify=CENTER,state="disabled",disabledbackground="white",
                            disabledforeground="black")
    txt_configFile.place(x=10,y=40)
    txt_configFile.focus()

    #boton 
    btn_FileConf = Button(v,text="Open FileConfig",font=("arial",12,"bold"),width=13,height=1,bg=Color_Boton,bd=0,fg="white",command = lambda: select_file_conf(config_file))
    btn_FileConf.place(x=240,y=80)

    #********** Boton de rename ****************
    btn_rename = Button(v,text="Rename",font=("arial",12,"bold"),width=13,height=1,bg=Color_Boton,bd=0,fg="white",command = lambda: Rename_All(directory_files,config_file))
    btn_rename.place(x=160,y=120)

    while StatusRun:
        v.mainloop()


#funcion para captar el folder de archivos
def select_directory(txt):
    global directory
    directory = filedialog.askdirectory(initialdir="/",title="select your directory files")
    txt.set(directory)


#funcion para captar el documento de configuracion
def select_file_conf(txt):
    global file_config
    file_config = filedialog.askopenfilename(initialdir="/",title="Select your config file",defaultextension=".xlsx")
    txt.set(file_config)


# funcion que copia y ordena los archivos
def Rename_All(TxtCourseFolder,TxtconfigFile):
    global StatusRun
    directory_list = {}

    # *********************** LISTAR ARCHIVOS EN EL DIRECTORIO ***************************
    for file in os.listdir(directory):

        # si la descarga contiene este nombre
        if "en curso de" in os.path.basename(file).lower(): 
            className = file[:file.lower().find("en curso de")].strip()

        else: 
            className = os.path.splitext(file)[0].strip()

        #direccion original del archivo sin cortar
        pathFile = os.path.join(directory, file)

        #tipo del archivo 
        typefile = os.path.splitext(file)[1]

        #la llave es el nombre ya modificado y los valores es el filepath original y el tipo del file 
        directory_list[className] = (pathFile,typefile)
            


    # ************************* CARGAR EL EXCEL *******************
    workbook = load_workbook(filename=file_config)

    courses_sheet = workbook.active


    # ******************** CREACION FOLDER EN DONDE ESTE EL ARCHIVO DE CONFIGURACION ******************************

    courseName = (os.path.splitext(os.path.basename(file_config))[0])

    dir_name = os.path.dirname(file_config)

    #Crear un folder para el curso
    directoryCourse = os.path.join(dir_name,courseName)
    os.makedirs(directoryCourse,exist_ok=True)


    #************************ ORGANIZACION DE ARCHIVOS ***************************

    #nombre del folder de subclase en lista
    actuallyFolderFile= ""
    counterFolder = 1
    counterFile = 1

    for cell in courses_sheet["A"]:

        #para cada folder principal creo una carpeta "solo para los valores de celda que contengan el char * "
        if "*" in cell.value or "-" in cell.value:
            actuallyFolderName = "{}) ".format(counterFolder)+ cell.value[1:].strip()
            actuallyFolderFile = os.path.join(directoryCourse,actuallyFolderName)

            os.makedirs(actuallyFolderFile,exist_ok=True)

            counterFolder +=1
            

        #empezar a organizar los archivos "solo para valores que no contengan el char * "
        elif "*" not in cell.value.strip() and "-" not in cell.value.strip():

            for name in directory_list.keys():

                coincidence = SequenceMatcher(None,cell.value.strip(),name)
                
                #si el nivel de coincidencia con el nombre es mayor al 85 % 
                if coincidence.ratio() >= 0.85:

                    #pathfile del archivo original "pepe sin cortar curso pepe.ts"
                    sourceFile = directory_list[name][0]

                    #nuevo nombre del archivo "1) pepe_ya_coratdo.html"
                    NameFile = "{}) ".format(counterFile) + name + directory_list[name][1]

                    #pathfile donde movere y renombrare el archivo
                    destinationFile = os.path.join(actuallyFolderFile,NameFile)

                    #muevo el archivo
                    shutil.copy(sourceFile,destinationFile)

                    counterFile +=1

    #cerramos el workbook
    workbook.close()

    
    opcion = messagebox.askretrycancel(message="Ejecucion realizada con exito ¿Desea renombrar otro curso prro?", title="Título")

    if opcion == False:
        StatusRun = False
    else:
        TxtconfigFile.set("")
        TxtCourseFolder.set("")
        
if __name__ == "__main__":
    run()
