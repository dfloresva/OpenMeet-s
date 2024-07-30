import pandas as pd
import webbrowser


class EnlaceExcel:
    def __init__(self, archivo_excel, nombre_columna):
        self.archivo_excel = archivo_excel
        self.nombre_columna = nombre_columna
        self.enlaces = []

        self.leer_excel()
        self.abrir_enlaces()

    def leer_excel(self):
        try:
            df = pd.read_excel(self.archivo_excel, sheet_name=0)
            if self.nombre_columna in df.columns:
                self.enlaces = df[self.nombre_columna].dropna().tolist()
                #print(self.enlaces)
            else:
                print(
                    f"La columna '{self.nombre_columna}' no se encuentra en el archivo Excel."
                )
        except Exception as e:
            print(f"Error de lectura del archivo: {e}")

    def abrir_enlaces(self):
        for enlace in self.enlaces:
            webbrowser.open(enlace)


class MenuSeleccion:
    def __init__(self, archivo_excel):
        self.archivo_excel = archivo_excel
        enlace_excel = self.ingresar_salon(self.archivo_excel)

    def ingresar_salon(self, archivo_excel):
        salonTemp = input("Num salon: ")
        # AQUI ENTRARIA UNA CONDICIONAL DE CONTROL DE INGRESO DE TEXTO

        return EnlaceExcel(archivo_excel, self.conc_String(salonTemp))

    def conc_String(self, salon):
        return "SP_" + salon


#  ASIGNACION DE VARIABLES DE DIRECCION Y ELECCION
archivo_excel1 = "enlaces.xlsx"
archivo_excel2 = "InformeTemasDesarrollados.xlsx"
nombre_columna = "SP_01"
nombre_columna = "INFORMES"

menuSeleccion = MenuSeleccion(archivo_excel1)

#enlace_excel = EnlaceExcel(archivo_excel2, nombre_columna)


# salonTemp = "01"
# print("SP_" + salonTemp)
