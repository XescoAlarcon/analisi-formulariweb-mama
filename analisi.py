import os
import win32com.client
from collections import Counter
from datetime import datetime
import pandas as pd
import re
import xlsxwriter

def pedir_anyo():
    anyo_actual = datetime.now().year
    while True:
        entrada = input(f"Introduce el año (formato AAAA, entre 2018 y {anyo_actual}): ")
        if entrada.isdigit() and len(entrada) == 4:
            anyo = int(entrada)
            if 2018 <= anyo <= anyo_actual:
                return anyo
        print("Año no válido. Inténtalo de nuevo.")

def mostrar_asuntos_por_ano(anyo):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    pst_path = os.path.join(os.path.dirname(__file__), "mama.pst")
    # Añadir el PST si no está ya cargado
    try:
        outlook.AddStore(pst_path)
    except Exception:
        pass  # Ya está cargado

    # Buscar la carpeta raíz del PST
    for store in outlook.Stores:
        if store.FilePath.lower() == pst_path.lower():
            root_folder = store.GetRootFolder()
            break
    else:
        print("No se encontró el archivo PST.")
        return

    # Navegar a la carpeta correspondiente
    if anyo == datetime.now().year:
        ruta = ["Bandeja de entrada", "Web", str(anyo)]
        carpeta_base = root_folder
    else:
        ruta = ["Bandeja de entrada", "Web", "Antiguos", str(anyo)]
        carpeta_base = root_folder

    try:
        for nombre in ruta:
            carpeta_base = carpeta_base.Folders[nombre]
    except Exception:
        print(f"No se encontró la carpeta para el año {anyo}.")
        return

    asuntos_validos = {
        "Formulario mamografia, canvi de visita",
        "Formulario mamografía, cambio de visita",
        "Formulario mamografia, anul·lar visita",
        "Formulario mamografía, anular visita"
    }
    print(f"Asunto y CIP filtrados en el año {anyo}:")

    total_cambio = 0
    total_anular = 0
    total_cip_incorrectos = 0
    distribucion_cambio_centro = {}
    distribucion_anular_motivo = {}  # Nueva estructura para anulación
    filas_cambios = []
    filas_anulaciones = []


    patron_cip = re.compile(r"^CIP:[A-Za-z]{4}1\d{4,}$")

    for item in carpeta_base.Items:
        if hasattr(item, "Subject") and item.Subject in asuntos_validos:
            cip = ""
            centro = ""
            edad = None
            motivo = ""
                   
            if hasattr(item, "Body"):
                lineas = item.Body.splitlines()
                if len(lineas) >= 3:
                    cip = lineas[2].replace(" ","")
                # Buscar centro solo si es "canvi/cambio de visita"
                if item.Subject.endswith("canvi de visita") or item.Subject.endswith("cambio de visita"):
                    for linea in lineas:
                        if linea.startswith("Centro Sanitario: "):
                            centro = linea[len("Centro Sanitario: "):].strip()
                            break
                        elif linea.startswith("Centre Sanitari: "):
                            centro = linea[len("Centre Sanitari: "):].strip()
                            break
                # Buscar motivo anulacion solo si es "anul·lar/anular visita"
                elif item.Subject.endswith("anul·lar visita") or item.Subject.endswith("anular visita"):
                    for linea in lineas:
                        if linea.startswith("Motiu de l'anul·lació: "):
                            motivo = linea[len("Motiu de l'anul·lació: "):].strip()
                            break
                        elif linea.startswith("Motivo de la anulación: "):
                            motivo = linea[len("Motivo de la anulación: "):].strip()
                            break
                                     # Sintetizar motivo
                    if motivo.startswith("Em faig regularment") or motivo.startswith("Me hago regularmente"):
                        motivo = "MX EXT"
                    elif motivo.startswith("Solament vull anul·lar") or motivo.startswith("Solo quiero anular"):
                        motivo = "MX < 6"
                    elif motivo.startswith("Ja he tingut") or motivo.startswith("Ya he tenido"):
                        motivo = "CA MAMA"
                    elif motivo.startswith("Tinc una altra") or motivo.startswith("Tengo otra"):
                        motivo = "MALALTIA BENIGNA"
                    elif motivo.startswith("Vaig ser estudiada") or motivo.startswith("Fui estudiada"):
                        motivo = "UCG"
                    elif motivo.startswith("De moment no") or motivo.startswith("De momento no"):
                        motivo = "NO INTERÈS"
                    elif motivo.startswith("Altres motius") or motivo.startswith("Otros motivos"):
                        motivo = "ALTRES"
                    else: 
                        motivo = "** " + motivo + " **"
            # Validar el CIP y calcular edad si es correcto
            if patron_cip.match(cip):
                try:
                    anyo_nacimiento = int("19" + cip[9:11])
                    edad = datetime.now().year - anyo_nacimiento
                except Exception:
                    edad = None
            else:
                total_cip_incorrectos += 1
                
            # Recopilar datos individuales
            if item.Subject.endswith("canvi de visita") or item.Subject.endswith("cambio de visita"):
                filas_cambios.append({
                    "Asunto": item.Subject,
                    "CIP": cip,
                    "Centro": centro,
                    "Edad": edad
                })
            elif item.Subject.endswith("anul·lar visita") or item.Subject.endswith("anular visita"):
                filas_anulaciones.append({
                    "Asunto": item.Subject,
                    "CIP": cip,
                    "Motivo": motivo,
                    "Edad": edad
                })  
                
            # Agrupar y contar
            if item.Subject.endswith("canvi de visita") or item.Subject.endswith("cambio de visita"):
                total_cambio += 1
                if centro and edad is not None:
                    # Clasificar por rango de edad
                    if edad < 50 or edad >= 80:
                        rango = "missing"
                    elif 50 <= edad <= 59:
                        rango = "[50-59]"
                    elif 60 <= edad <= 69:
                        rango = "[60-69]"
                    elif 70 <= edad <= 79:
                        rango = "[70-79]"
                    else:
                        rango = "missing"
                    if centro not in distribucion_cambio_centro:
                        distribucion_cambio_centro[centro] = {"[50-59]": 0, "[60-69]": 0, "[70-79]": 0, "missing": 0}
                    distribucion_cambio_centro[centro][rango] += 1
                elif centro:
                    if centro not in distribucion_cambio_centro:
                        distribucion_cambio_centro[centro] = {"[50-59]": 0, "[60-69]": 0, "[70-79]": 0, "missing": 0}
                    distribucion_cambio_centro[centro]["missing"] += 1
            elif item.Subject.endswith("anul·lar visita") or item.Subject.endswith("anular visita"):
                total_anular += 1
                # Distribución por motivo y rango de edad
                if motivo:
                    if edad is not None:
                        if edad < 50 or edad >= 80:
                            rango = "missing"
                        elif 50 <= edad <= 59:
                            rango = "[50-59]"
                        elif 60 <= edad <= 69:
                            rango = "[60-69]"
                        elif 70 <= edad <= 79:
                            rango = "[70-79]"
                        else:
                            rango = "missing"
                    else:
                        rango = "missing"
                    if motivo not in distribucion_anular_motivo:
                        distribucion_anular_motivo[motivo] = {"[50-59]": 0, "[60-69]": 0, "[70-79]": 0, "missing": 0}
                    distribucion_anular_motivo[motivo][rango] += 1
    print(f"\n--------------------------------------------------")
    print(f"Total 'canvi de visita' o 'cambio de visita': {total_cambio}")
    print(f"--------------------------------------------------\n")
    
    print("Distribución por centro sanitario y rango de edad para 'canvi/cambio de visita':")
    for centro, rangos in distribucion_cambio_centro.items():
        total_centro = sum(rangos.values())
        print(f"  {centro}:")
        for rango, cantidad in rangos.items():
            print(f"    {rango}: {cantidad}")
        print(f"    Total: {total_centro}")
    print(f"\n-----------------------------------------------")            
    print(f"Total 'anul·lar visita' o 'anular visita': {total_anular}")
    print(f"-----------------------------------------------\n")

    print("Distribución por motivo y rango de edad para 'anul·lar/anular visita':")
    for motivo, rangos in distribucion_anular_motivo.items():
        total_motivo = sum(rangos.values())
        print(f"  {motivo}:")
        for rango, cantidad in rangos.items():
            print(f"    {rango}: {cantidad}")
        print(f"    Total: {total_motivo}")

    print(f"\n============================")            
    print(f"Total CIPs incorrectos: {total_cip_incorrectos}")
    print(f"============================\n")            

    # Preguntar si se quiere exportar a Excel
    exportar = input("¿Quieres exportar los datos a Excel? (s/n): ").strip().lower()
    if exportar == "s":

        # Crear DataFrames de las distribuciones
        df_cambios_dist = pd.DataFrame([
            {"Centro": centro, "Rango edad": rango, "Cantidad": cantidad}
            for centro, rangos in distribucion_cambio_centro.items()
            for rango, cantidad in rangos.items()
        ])
        # Añadir totales por centro
        for centro, rangos in distribucion_cambio_centro.items():
            total = sum(rangos.values())
            df_cambios_dist = pd.concat([
                df_cambios_dist,
                pd.DataFrame([{"Centro": centro, "Rango edad": "Total", "Cantidad": total}])
            ], ignore_index=True)

        df_anulaciones_dist = pd.DataFrame([
            {"Motivo": motivo, "Rango edad": rango, "Cantidad": cantidad}
            for motivo, rangos in distribucion_anular_motivo.items()
            for rango, cantidad in rangos.items()
        ])
        # Añadir totales por motivo
        for motivo, rangos in distribucion_anular_motivo.items():
            total = sum(rangos.values())
            df_anulaciones_dist = pd.concat([
                df_anulaciones_dist,
                pd.DataFrame([{"Motivo": motivo, "Rango edad": "Total", "Cantidad": total}])
            ], ignore_index=True)

        nombre_fichero = f"datos_{anyo}.xlsx"
        with pd.ExcelWriter(nombre_fichero, engine="xlsxwriter") as writer:
            pd.DataFrame(filas_cambios).to_excel(writer, sheet_name="Cambios de visita", index=False)
            pd.DataFrame(filas_anulaciones).to_excel(writer, sheet_name="Motivos cancelación", index=False)
            df_cambios_dist.to_excel(writer, sheet_name="Distribución cambios", index=False)
            df_anulaciones_dist.to_excel(writer, sheet_name="Distribución anulaciones", index=False)
        print(f"Datos exportados a {nombre_fichero}")

# Ejemplo de uso:
anyo = pedir_anyo()
mostrar_asuntos_por_ano(anyo)