import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd
import win32com.client
import subprocess
import time
import pyperclip
import os

# ==========================================
# CONFIGURACIÓN
# ==========================================

SAP_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
CONEXION_SAP = "SAP_CONNECTION"

archivo_seleccionado = ""

# ==========================================
# FUNCIÓN PRINCIPAL
# ==========================================

def seleccionar_archivo():

    global archivo_seleccionado

    archivo_seleccionado = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )

    if archivo_seleccionado:

        label_archivo.config(
            text=os.path.basename(archivo_seleccionado)
        )

def ejecutar_proceso():

    usuario = entry_user.get().strip()
    password = entry_pass.get().strip()

    if not usuario or not password:

        messagebox.showerror(
            "Error",
            "Ingrese usuario y contraseña SAP"
        )

        return

    try:

        # ==========================================
        # LEER ARCHIVO
        # ==========================================

        if not archivo_seleccionado:

            messagebox.showerror(
                "Error",
                "Seleccione un archivo"
            )

            return

        df = pd.read_excel(archivo_seleccionado)

        pedidos = [
            f"*{p}*"
            for p in df["Pedido"].astype(str).tolist()
        ]

        if len(pedidos) == 0:

            messagebox.showerror(
                "Error",
                "No se encontraron pedidos"
            )

            return

        # ==========================================
        # CONEXIÓN SAP
        # ==========================================

        subprocess.Popen(SAP_PATH)

        time.sleep(3)

        SapGuiAuto = win32com.client.GetObject("SAPGUI")

        application = SapGuiAuto.GetScriptingEngine

        connection = application.OpenConnection(
            CONEXION_SAP,
            True
        )

        time.sleep(2)

        session = connection.Children(0)

        # ==========================================
        # LOGIN
        # ==========================================

        session.findById(
            "wnd[0]/usr/txtRSYST-BNAME"
        ).text = usuario

        session.findById(
            "wnd[0]/usr/pwdRSYST-BCODE"
        ).text = password

        session.findById("wnd[0]").sendVKey(0)

        time.sleep(5)

        # ======================================
        # SESIÓN YA ABIERTA
        # ======================================

        try:

            print("Sesión SAP ya abierta detectada")

            session.findById(
                "wnd[1]/usr/radMULTI_LOGON_OPT1"
            ).select()

            session.findById(
                "wnd[1]/tbar[0]/btn[0]"
            ).press()

            print("Continuando con el inicio de sesión")

        except Exception:

            print("No apareció el popup de sesión abierta")

        # ==========================================
        # ENTRAR A BANDEJA
        # ==========================================

        session.findById("wnd[0]").maximize()

        session.findById(
            "wnd[0]/tbar[1]/btn[36]"
        ).press()

        time.sleep(2)

        # ==========================================
        # SELECCIONAR NODO
        # ==========================================

        session.findById(
            "wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[0]/shell"
        ).selectedNode = "          2"

        time.sleep(1)

        shell = session.findById(
            "wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell"
        )

        # ==========================================
        # FILTRAR OBJDES
        # ==========================================

        shell.setCurrentCell(-1, "OBJDES")

        shell.selectColumn("OBJDES")

        shell.selectedRows = ""

        shell.contextMenu()

        shell.selectContextMenuItem("&FILTER")

        time.sleep(1)

        # ==========================================
        # ABRIR SELECCIÓN MÚLTIPLE
        # ==========================================

        session.findById(
            "wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH"
        ).press()

        time.sleep(1)

        # ==========================================
        # COPIAR AL CLIPBOARD
        # ==========================================

        texto_pedidos = "\r\n".join(pedidos)

        pyperclip.copy(texto_pedidos)

        time.sleep(1)

        # ==========================================
        # UPLOAD DESDE PORTAPAPELES
        # ==========================================

        session.findById(
            "wnd[2]/tbar[0]/btn[24]"
        ).press()

        time.sleep(2)

        session.findById(
            "wnd[2]/tbar[0]/btn[8]"
        ).press()

        time.sleep(1)

        session.findById(
            "wnd[1]/tbar[0]/btn[0]"
        ).press()

        time.sleep(2)

        # ==========================================
        # SELECCIONAR RESULTADOS
        # ==========================================

        filas = shell.RowCount

        cantidad = min(len(pedidos), filas)

        if cantidad == 0:

            messagebox.showwarning(
                "Sin resultados",
                "No se encontraron pedidos"
            )

            return

        shell.setCurrentCell(0, "OBJDES")

        time.sleep(1)

        shell.selectedRows = f"0-{cantidad-1}"

        shell.selectionChanged()

        # ==========================================
        # ABRIR PEDIDOS
        # ==========================================

        shell.doubleClickCurrentCell()

        time.sleep(2)

        # ==========================================
        # VOLVER
        # ==========================================

        session.findById(
            "wnd[0]/tbar[0]/btn[3]"
        ).press()

        time.sleep(2)

        # ==========================================
        # LIBERACIÓN AUTOMÁTICA
        # ==========================================

        for i in range(cantidad):

            try:

                session.findById(
                    "wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell"
                ).sapEvent(
                    "",
                    "",
                    "sapevent:DECI:0001"
                )

            except:

                messagebox.showerror(
                    "Error",
                    "No se encontró botón liberar"
                )

                break

            time.sleep(2)

            # ======================================
            # POPUP
            # ======================================

            try:

                session.findById(
                    "wnd[1]/usr/btnSPOP-OPTION1"
                ).press()

                time.sleep(1)

            except:
                pass

            # ======================================
            # NAVEGAR
            # ======================================

            if i < cantidad - 1:

                try:
                    session.findById("wnd[0]").sendVKey(3)
                    time.sleep(1)
                except:
                    pass

                try:
                    session.findById("wnd[0]").sendVKey(3)
                    time.sleep(1)
                except:
                    pass
        # ==========================================
        # LIBERACION FINAL
        # ==========================================
        
        print("Liberación final del último pedido...")

        try:

            session.findById("wnd[0]").sendVKey(3)
            time.sleep(2)

            session.findById("wnd[0]").sendVKey(3)
            time.sleep(2)

            # 2. Intentar liberar otra vez
            session.findById(
                "wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell"
            ).sapEvent(
                "",
                "",
                "sapevent:DECI:0001"
            )

            time.sleep(2)

            # 3. Popup si aparece
            try:
                session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            except:
                pass

        except Exception as e:
            print("No fue necesaria liberación final:", e)
        
        # ==========================================
        # FINALIZADO
        # ==========================================

        messagebox.showinfo(
            "Proceso finalizado",
            f"Se liberaron {cantidad} pedidos correctamente"
        )

    except Exception as e:

        messagebox.showerror(
            "Error",
            str(e)
        )

# ==========================================
# INTERFAZ
# ==========================================

ventana = tk.Tk()

ventana.title("SAP Order Release Automation")

ventana.geometry("400x300")

ventana.resizable(False, False)

tk.Label(
    ventana,
    text="Usuario SAP"
).pack(pady=10)

entry_user = tk.Entry(
    ventana,
    width=35
)

entry_user.pack()

tk.Label(
    ventana,
    text="Contraseña SAP"
).pack(pady=10)

entry_pass = tk.Entry(
    ventana,
    show="*",
    width=35
)

entry_pass.pack()

tk.Button(
    ventana,
    text="Seleccionar Archivo",
    command=seleccionar_archivo,
    width=20
).pack(pady=10)

label_archivo = tk.Label(
    ventana,
    text="Ningún archivo seleccionado"
)

label_archivo.pack()

tk.Button(
    ventana,
    text="Iniciar Liberación",
    command=ejecutar_proceso,
    width=20,
    height=2
).pack(pady=25)

ventana.mainloop()
