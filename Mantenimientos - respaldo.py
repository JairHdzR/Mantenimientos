# ============================================================
# 1) IMPORTACIONES Y CONFIGURACIÓN BÁSICA
# ============================================================
import os
import calendar
import hashlib
import sqlite3
import secrets
import uuid
from datetime import datetime, date, timedelta
from typing import Union

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import logging

# tkcalendar para DateEntry (calendario desplegable)
try:
    from tkcalendar import DateEntry
    TKCAL_OK = True
except Exception:
    TKCAL_OK = False

# --- Importación opcional de pandas para Excel ---
try:
    import pandas as pd
    PANDAS_OK = True
except Exception:
    PANDAS_OK = False

# Archivo de base de datos (SQLite)
ARCHIVO_BD = "mantenimiento_es.db"
ICONO_APP_ICO = "escudouvm.ico"   # recomendado para Windows

# Logging básico (puedes comentar si no lo usas)
logging.basicConfig(level=logging.WARNING, format="%(levelname)s:%(message)s")


# ============================================================
# 2) UTILIDADES DE UI (ICONO, ESTILO, CURSOR Y PANTALLA COMPLETA)
# ============================================================
def aplicar_icono_aplicacion(ventana: Union[tk.Tk, tk.Toplevel]):
    """Establece el icono de la aplicación en la ventana indicada."""
    if getattr(ventana, "_icono_aplicado", False):
        return
    try:
        if os.path.exists(ICONO_APP_ICO):
            ventana.iconbitmap(ICONO_APP_ICO)
            ventana._icono_aplicado = True
        else:
            if not getattr(aplicar_icono_aplicacion, "_warned", False):
                logging.warning(f"Archivo de icono '{ICONO_APP_ICO}' no encontrado.")
                aplicar_icono_aplicacion._warned = True
    except Exception as e:
        logging.warning(f"No se pudo aplicar el icono: {e}")


def aplicar_estilo_global():
    """Estilo minimalista oscuro con acentos elegantes."""
    estilo = ttk.Style()
    try:
        estilo.theme_use("clam")
    except Exception:
        pass

    # Paleta
    fondo        = "#0B1220"
    tarjeta      = "#101826"
    superficie   = "#0E1624"
    texto        = "#E6E8ED"
    texto_muted  = "#B8C1D1"
    primario     = "#7C9BFF"
    primario_hov = "#A6B8FF"
    exito        = "#22C55E"
    peligro      = "#EF4444"
    borde        = "#1D2A3E"
    cabecera     = "#0B1220"

    # Fondos base
    estilo.configure(".", background=fondo, foreground=texto)
    estilo.configure("App.TFrame", background=fondo)
    estilo.configure("Card.TFrame", background=tarjeta, borderwidth=1, relief="flat")

    # Labels
    estilo.configure("Titulo.TLabel",
                     font=("Segoe UI", 26, "bold"),
                     foreground=texto,
                     background=tarjeta)
    estilo.configure("Cuerpo.TLabel",
                     font=("Segoe UI", 12),
                     foreground=texto_muted,
                     background=tarjeta)

    # Entradas / Combos
    estilo.configure("Entrada.TEntry",
                     fieldbackground=superficie,
                     foreground=texto,
                     background=tarjeta,
                     bordercolor=borde,
                     lightcolor=borde,
                     darkcolor=borde,
                     padding=8)
    estilo.map("Entrada.TEntry",
               fieldbackground=[("focus", superficie)],
               bordercolor=[("focus", primario)])

    estilo.configure("Combo.TCombobox",
                     fieldbackground=superficie,
                     foreground=texto,
                     background=tarjeta,
                     bordercolor=borde,
                     lightcolor=borde,
                     darkcolor=borde)
    estilo.map("Combo.TCombobox",
               fieldbackground=[("readonly", superficie)],
               bordercolor=[("focus", primario)])

    # Botones
    estilo.configure("Primario.TButton",
                     font=("Segoe UI", 12, "bold"),
                     padding=10,
                     background=primario,
                     foreground="#0B1220",
                     borderwidth=0)
    estilo.map("Primario.TButton",
               background=[("active", primario_hov), ("pressed", primario_hov)],
               foreground=[("active", "#0B1220"), ("pressed", "#0B1220")])

    estilo.configure("Fantasma.TButton",
                     font=("Segoe UI", 12, "bold"),
                     padding=10,
                     background=tarjeta,
                     foreground=texto,
                     borderwidth=1)
    estilo.map("Fantasma.TButton",
               background=[("active", superficie)],
               foreground=[("active", texto)],
               bordercolor=[("!disabled", borde)])

    estilo.configure("Success.TButton",
                     font=("Segoe UI", 12, "bold"),
                     padding=10,
                     background=exito,
                     foreground="#06250F",
                     borderwidth=0)
    estilo.map("Success.TButton", background=[("active", "#4ADE80")])

    estilo.configure("Danger.TButton",
                     font=("Segoe UI", 12, "bold"),
                     padding=10,
                     background=peligro,
                     foreground="#2A0B0B",
                     borderwidth=0)
    estilo.map("Danger.TButton", background=[("active", "#F87171")])

    # Treeview
    estilo.configure("Treeview",
                     background=superficie,
                     fieldbackground=superficie,
                     foreground=texto,
                     bordercolor=borde,
                     rowheight=26)
    estilo.configure("Treeview.Heading",
                     font=("Segoe UI", 11, "bold"),
                     background=cabecera,
                     foreground=texto,
                     bordercolor=borde)
    estilo.map("Treeview.Heading", background=[("active", cabecera)])

    # Notebook (pestañas)
    estilo.configure("TNotebook", background=fondo, borderwidth=0)
    estilo.configure("TNotebook.Tab",
                     background=tarjeta,
                     foreground=texto_muted,
                     font=("Segoe UI", 12, "bold"),
                     padding=[16, 8],
                     borderwidth=0)
    estilo.map("TNotebook.Tab",
        background=[
            ("selected", primario_hov),
            ("active", primario)
        ],
        foreground=[
            ("selected", "#0B1220"),
            ("active", "#0B1220"),
            ("!selected", texto_muted)
        ]
    )


def habilitar_pantalla_completa(ventana: tk.Tk) -> None:
    """Arranca en pantalla completa y permite alternar con F11/ESC."""
    try:
        ventana.resizable(True, True)
        ventana.state('normal')
    except Exception:
        pass
    ventana._es_fullscreen = True
    try:
        ventana.attributes("-fullscreen", True)
    except Exception:
        pass

    def _alternar_fullscreen(_event=None):
        ventana._es_fullscreen = not getattr(ventana, "_es_fullscreen", False)
        try:
            ventana.attributes("-fullscreen", ventana._es_fullscreen)
        except Exception:
            pass

    ventana.bind("<F11>", _alternar_fullscreen)
    ventana.bind("<Escape>", _alternar_fullscreen)


def poner_caret_blanco(*widgets):
    """Fuerza el cursor de inserción (caret) en blanco en Entry/Text."""
    for w in widgets:
        try:
            w.configure(insertbackground="white")
        except Exception:
            try:
                w.tk.call(w._w, "configure", "-insertbackground", "white")
            except Exception:
                pass


# ============================================================
# 2.1) UTILIDADES DE FECHA (CONVERSIÓN)
# ============================================================
def _a_iso(fecha_ddmmaaaa: str): 
    """Convierte 'DD-MM-AAAA' -> 'YYYY-MM-DD'."""
    return datetime.strptime(fecha_ddmmaaaa, "%d-%m-%Y").strftime("%Y-%m-%d")

def _a_ddmmaaaa(fecha_iso: str):
    """Convierte 'YYYY-MM-DD' -> 'DD-MM-AAAA'."""
    try:
        return datetime.strptime(fecha_iso, "%Y-%m-%d").strftime("%d-%m-%Y")
    except Exception:
        return fecha_iso or ""


# ============================================================
# 3) SEGURIDAD (HASH DE CONTRASEÑAS) Y USUARIOS
# ============================================================
def crear_hash(contrasena: str, sal: str):
    return hashlib.sha256((sal + contrasena).encode("utf-8")).hexdigest()


def crear_usuario(con, usuario: str, contrasena: str, rol: str) -> None:
    sal = secrets.token_hex(16)
    hash_pwd = crear_hash(contrasena, sal)
    with con:
        con.execute(
            "INSERT INTO usuarios(usuario, hash_contrasena, sal, rol) VALUES(?,?,?,?)",
            (usuario, hash_pwd, sal, rol)
        )


def verificar_usuario(con, usuario: str, contrasena: str):
    cur = con.execute(
        "SELECT id, usuario, hash_contrasena, sal, rol FROM usuarios WHERE usuario=?",
        (usuario,)
    )
    fila = cur.fetchone()
    if not fila:
        return None
    uid, uname, hash_guardado, sal, rol = fila
    if crear_hash(contrasena, sal) == hash_guardado:
        return {"id": uid, "usuario": uname, "rol": rol}
    return None


# ============================================================
# 4) INICIALIZACIÓN DE BD (TABLAS, SEED, MIGRACIONES SUAVES)
# ============================================================
def iniciar_bd(archivo_bd=ARCHIVO_BD):
    """Inicializa la base de datos y retorna la conexión."""
    primera_vez = not os.path.exists(archivo_bd)
    con = sqlite3.connect(archivo_bd)
    con.execute("PRAGMA foreign_keys = ON;")

    # usuarios
    con.execute("""
        CREATE TABLE IF NOT EXISTS usuarios(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT UNIQUE NOT NULL,
            hash_contrasena TEXT NOT NULL,
            sal TEXT NOT NULL,
            rol TEXT CHECK(rol IN ('administrador','laboratorista')) NOT NULL
        );
    """)

    # equipos (+ fecha_registro, + creado_por)
    con.execute("""
        CREATE TABLE IF NOT EXISTS equipos(
            id_equipo TEXT PRIMARY KEY,
            nombre TEXT NOT NULL,
            marca TEXT,
            modelo TEXT,
            serie TEXT,
            ubicacion TEXT,
            descripcion TEXT
        );
    """)
    # columnas nuevas si no existen
    try:
        con.execute("ALTER TABLE equipos ADD COLUMN fecha_registro TEXT;")
    except sqlite3.OperationalError:
        pass
    try:
        con.execute("ALTER TABLE equipos ADD COLUMN creado_por INTEGER;")
    except sqlite3.OperationalError:
        pass

    # mantenimientos (+ registrado_en)
    con.execute("""
        CREATE TABLE IF NOT EXISTS mantenimientos(
            id_mantenimiento TEXT PRIMARY KEY,
            equipo_id TEXT NOT NULL,
            fecha TEXT NOT NULL,
            tipo TEXT CHECK(tipo IN ('Preventivo','Correctivo')) NOT NULL,
            notas TEXT,
            estado TEXT CHECK(estado IN ('Pendiente','Completado')) NOT NULL DEFAULT 'Pendiente',
            proveedor TEXT,
            costo REAL DEFAULT 0,
            creado_por INTEGER,
            FOREIGN KEY(equipo_id) REFERENCES equipos(id_equipo) ON DELETE CASCADE ON UPDATE CASCADE,
            FOREIGN KEY(creado_por) REFERENCES usuarios(id) ON DELETE SET NULL
        );
    """)
    try:
        con.execute("ALTER TABLE mantenimientos ADD COLUMN registrado_en TEXT;")
    except sqlite3.OperationalError:
        pass

    # historicos
    con.execute("""
        CREATE TABLE IF NOT EXISTS historicos(
            id_historico TEXT PRIMARY KEY,
            id_mantenimiento TEXT,
            equipo_id TEXT,
            fecha TEXT,
            tipo TEXT,
            notas TEXT,
            estado TEXT,
            proveedor TEXT,
            costo REAL,
            creado_por INTEGER,
            registrado_en TEXT
        );
    """)

    # ajustes
    con.execute("""
        CREATE TABLE IF NOT EXISTS ajustes(
            clave TEXT PRIMARY KEY,
            valor TEXT
        );
    """)

    # Seed
    if primera_vez:
        usuarios_semilla = [
            ("Adriana Resendis", "laboratorista"),
            ("Norma Angelica", "laboratorista"),
            ("Alfredo Bernal", "laboratorista"),
            ("Mayreli Sánchez", "laboratorista"),
            ("Ian Aguilar", "laboratorista"),
            ("Carlos Lira", "laboratorista"),
            ("Antonio Hernández", "laboratorista"),
            ("Anabell Nochebuena", "laboratorista"),
            ("Edgar Martell", "laboratorista"),
            ("Marte Gómez", "laboratorista"),
            ("Maria del Carmen Rodriguez Chan", "administrador"),
            ("Admin", "administrador"),
        ]
        try:
            for nombre, rol in usuarios_semilla:
                crear_usuario(con, nombre, "1234", rol)
        except sqlite3.IntegrityError:
            pass

        with con:
            con.execute("INSERT OR IGNORE INTO ajustes(clave, valor) VALUES('dia_mantenimiento','1');")
            con.execute("INSERT OR IGNORE INTO ajustes(clave, valor) VALUES('preaviso_dias_fin_mes','5');")
            con.execute("INSERT OR IGNORE INTO ajustes(clave, valor) VALUES('ultima_revision_alerta','');")

    return con


# ============================================================
# 5) AJUSTES
# ============================================================
def obtener_ajuste(con, clave: str, por_defecto=None):
    cur = con.execute("SELECT valor FROM ajustes WHERE clave=?", (clave,))
    fila = cur.fetchone()
    return fila[0] if fila and fila[0] is not None else por_defecto


def establecer_ajuste(con, clave: str, valor: str) -> None:
    with con:
        con.execute(
            "INSERT INTO ajustes(clave, valor) VALUES(?, ?) "
            "ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor;",
            (clave, valor)
        )


# ============================================================
# 6) ALERTAS (LÓGICA)
# ============================================================
def fecha_fin_de_mes(d: date) -> date:
    ultimo = calendar.monthrange(d.year, d.month)[1]
    return date(d.year, d.month, ultimo)


def contar_equipos_sin_mantenimiento_en_mes(con) -> int:
    hoy = date.today()
    inicio_mes = date(hoy.year, hoy.month, 1).strftime("%Y-%m-%d")
    proximo_mes = date(hoy.year + (hoy.month // 12), (hoy.month % 12) + 1, 1).strftime("%Y-%m-%d")

    con_mant = con.execute("""
        SELECT COUNT(DISTINCT e.id_equipo)
        FROM equipos e
        JOIN mantenimientos m ON m.equipo_id = e.id_equipo
        WHERE m.fecha >= ? AND m.fecha < ?;
    """, (inicio_mes, proximo_mes)).fetchone()[0]
    total = con.execute("SELECT COUNT(*) FROM equipos;").fetchone()[0]
    return max(total - con_mant, 0)


def revisar_alertas(con, forzar=False) -> None:
    hoy = date.today()
    hoy_str = hoy.strftime("%Y-%m-%d")

    ultima = obtener_ajuste(con, "ultima_revision_alerta", "")
    if (ultima == hoy_str) and (not forzar):
        return

    dia_alerta = int(obtener_ajuste(con, "dia_mantenimiento", "1") or "1")
    dias_pre = int(obtener_ajuste(con, "preaviso_dias_fin_mes", "5") or "5")

    fecha_alerta = date(hoy.year, hoy.month, min(dia_alerta, 28))
    fin_mes = fecha_fin_de_mes(hoy)
    fecha_preaviso = fin_mes - timedelta(days=dias_pre)

    mensajes = []
    if hoy == fecha_alerta:
        mensajes.append("📅 ¡Mes de mantenimientos! Recuerda ejecutar y registrar los servicios programados.")
    if hoy == fecha_preaviso:
        faltantes = contar_equipos_sin_mantenimiento_en_mes(con)
        mensajes.append(f"⏳ Quedan {dias_pre} días para cerrar el mes. "
                        f"Equipos sin mantenimiento registrado este mes: {faltantes}.")

    if mensajes:
        messagebox.showinfo("Alertas de mantenimiento", "\n\n".join(mensajes))
        establecer_ajuste(con, "ultima_revision_alerta", hoy_str)
        return

    if forzar:
        faltantes = contar_equipos_sin_mantenimiento_en_mes(con)
        resumen = [
            f"🔔 Próxima alerta principal (día del mes): {fecha_alerta.strftime('%Y-%m-%d')}",
            f"🔔 Alerta previa ({dias_pre} días antes del fin de mes): {fecha_preaviso.strftime('%Y-%m-%d')}",
            f"📊 Equipos SIN mantenimiento registrado este mes: {faltantes}",
        ]
        messagebox.showinfo("Estado de alertas", "\n".join(resumen))

    establecer_ajuste(con, "ultima_revision_alerta", hoy_str)


# ============================================================
# 7) SELECTOR DE FECHA (FALLBACK SIN LIBRERÍAS)
# ============================================================
class DatePicker(tk.Toplevel):
    """Calendario manual para elegir fecha (DD-MM-AAAA) si no hay tkcalendar."""
    def __init__(self, master, entry_obj: tk.Entry, fecha_inicial: str = ""):
        super().__init__(master)
        self.title("Seleccionar fecha")
        aplicar_icono_aplicacion(self)
        self.resizable(False, False)
        self.entry_obj = entry_obj

        # que no se pierda en fullscreen
        self.transient(master)
        self.grab_set()
        self.focus_set()
        self.attributes("-topmost", True)
        self.bind("<Escape>", lambda e: self.destroy())

        hoy = date.today()
        if fecha_inicial:
            try:
                dt = datetime.strptime(fecha_inicial, "%d-%m-%Y").date()
                self.year, self.month = dt.year, dt.month
            except Exception:
                self.year, self.month = hoy.year, hoy.month
        else:
            self.year, self.month = hoy.year, hoy.month

        body = ttk.Frame(self)
        body.pack(padx=8, pady=8)

        nav = ttk.Frame(body)
        nav.pack(fill="x")

        self.lb_mes = ttk.Label(nav, text="", style="Cuerpo.TLabel")
        btn_prev = ttk.Button(nav, text="◀", width=3, style="Fantasma.TButton",
                              command=self._prev_month)
        btn_next = ttk.Button(nav, text="▶", width=3, style="Fantasma.TButton",
                              command=self._next_month)
        btn_prev.pack(side="left")
        self.lb_mes.pack(side="left", expand=True)
        btn_next.pack(side="right")

        self.grid_dias = ttk.Frame(body)
        self.grid_dias.pack(pady=6)

        self._render()
        self.update_idletasks()
        try:
            x = self.entry_obj.winfo_rootx()
            y = self.entry_obj.winfo_rooty() + self.entry_obj.winfo_height()
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

    def _prev_month(self):
        if self.month == 1:
            self.month = 12
            self.year -= 1
        else:
            self.month -= 1
        self._render()

    def _next_month(self):
        if self.month == 12:
            self.month = 1
            self.year += 1
        else:
            self.month += 1
        self._render()

    def _render(self):
        for w in self.grid_dias.winfo_children():
            w.destroy()
        self.lb_mes.config(text=f"{calendar.month_name[self.month]} {self.year}")

        for i, dname in enumerate(["Lu","Ma","Mi","Ju","Vi","Sa","Do"]):
            ttk.Label(self.grid_dias, text=dname, style="Cuerpo.TLabel").grid(row=0, column=i, padx=4, pady=4)

        monthcal = calendar.Calendar(firstweekday=0).monthdayscalendar(self.year, self.month)
        for r, week in enumerate(monthcal, start=1):
            for c, day in enumerate(week):
                if day == 0:
                    ttk.Label(self.grid_dias, text=" ").grid(row=r, column=c, padx=2, pady=2)
                    continue
                def _mkcmd(yy=self.year, mm=self.month, dd=day):
                    return lambda: self._elegir(yy, mm, dd)
                ttk.Button(self.grid_dias, text=str(day), width=3, style="Fantasma.TButton",
                           command=_mkcmd()).grid(row=r, column=c, padx=2, pady=2)

    def _elegir(self, y, m, d):
        self.entry_obj.delete(0, "end")
        self.entry_obj.insert(0, f"{d:02d}-{m:02d}-{y:04d}")  # DD-MM-AAAA
        self.destroy()


# ============================================================
# 8) DIÁLOGOS: EQUIPO / MANTENIMIENTO / USUARIO / ALERTAS
# ============================================================
class DialogoEquipo(tk.Toplevel):
    def __init__(self, master, titulo="Equipo", datos=None, usuario_actual=None):
        super().__init__(master)
        self.title(titulo)
        aplicar_icono_aplicacion(self)
        self.resizable(False, False)
        self.resultado = None
        self.usuario_actual = usuario_actual

        marco = ttk.Frame(self, style="Card.TFrame", padding=12)
        marco.grid(row=0, column=0, sticky="nsew")

        etiquetas = ["ID de equipo *", "Nombre *", "Marca", "Modelo", "Serie", "Ubicación", "Descripción"]
        for i, txt in enumerate(etiquetas):
            ttk.Label(marco, text=txt, style="Cuerpo.TLabel").grid(row=i, column=0, sticky="e", padx=6, pady=4)

        self.e_id = ttk.Entry(marco, width=38, style="Entrada.TEntry")
        self.e_nombre = ttk.Entry(marco, width=38, style="Entrada.TEntry")
        self.e_marca = ttk.Entry(marco, width=38, style="Entrada.TEntry")
        self.e_modelo = ttk.Entry(marco, width=38, style="Entrada.TEntry")
        self.e_serie = ttk.Entry(marco, width=38, style="Entrada.TEntry")
        self.e_ubic = ttk.Entry(marco, width=38, style="Entrada.TEntry")
        self.t_desc = tk.Text(marco, width=38, height=4)

        for w in (self.e_id, self.e_nombre, self.e_marca, self.e_modelo, self.e_serie, self.e_ubic, self.t_desc):
            poner_caret_blanco(w)

        self.e_id.grid(row=0, column=1, padx=6, pady=4)
        self.e_nombre.grid(row=1, column=1, padx=6, pady=4)
        self.e_marca.grid(row=2, column=1, padx=6, pady=4)
        self.e_modelo.grid(row=3, column=1, padx=6, pady=4)
        self.e_serie.grid(row=4, column=1, padx=6, pady=4)
        self.e_ubic.grid(row=5, column=1, padx=6, pady=4)
        self.t_desc.grid(row=6, column=1, padx=6, pady=4)

        if datos:
            self.e_id.insert(0, datos.get("id_equipo", ""))
            self.e_nombre.insert(0, datos.get("nombre", ""))
            self.e_marca.insert(0, datos.get("marca", ""))
            self.e_modelo.insert(0, datos.get("modelo", ""))
            self.e_serie.insert(0, datos.get("serie", ""))
            self.e_ubic.insert(0, datos.get("ubicacion", ""))
            self.t_desc.insert("1.0", datos.get("descripcion", "") or "")

        zona_botones = ttk.Frame(marco, style="Card.TFrame")
        zona_botones.grid(row=7, column=0, columnspan=2, pady=8)
        ttk.Button(zona_botones, text="Guardar", command=self._on_guardar, style="Success.TButton").pack(side="left", padx=6)
        ttk.Button(zona_botones, text="Cancelar", command=self.destroy, style="Fantasma.TButton").pack(side="left", padx=6)

        self.bind("<Return>", lambda e: self._on_guardar())
        self.grab_set()
        self.e_id.focus_set()

    def _on_guardar(self):
        id_equipo = self.e_id.get().strip()
        nombre = self.e_nombre.get().strip()
        if not id_equipo or not nombre:
            messagebox.showwarning("Validación", "El ID de equipo y el Nombre son obligatorios.")
            return
        self.resultado = {
            "id_equipo": id_equipo,
            "nombre": nombre,
            "marca": self.e_marca.get().strip(),
            "modelo": self.e_modelo.get().strip(),
            "serie": self.e_serie.get().strip(),
            "ubicacion": self.e_ubic.get().strip(),
            "descripcion": self.t_desc.get("1.0", "end").strip(),
            "fecha_registro": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        self.destroy()


class DialogoMantenimiento(tk.Toplevel):
    """ID oculto (auto). Fecha con DateEntry (tkcalendar) o fallback. registrado_en/creado_por automáticos."""
    def __init__(self, master, titulo="Mantenimiento", ids_equipos=(), datos=None):
        super().__init__(master)
        self.title(titulo)
        aplicar_icono_aplicacion(self)
        self.resizable(False, False)
        self.resultado = None

        marco = ttk.Frame(self, style="Card.TFrame", padding=12)
        marco.grid(row=0, column=0, sticky="nsew")

        labels = ["Equipo (ID) *", "Fecha (DD-MM-AAAA) *", "Tipo *", "Estado", "Proveedor", "Costo (num)", "Notas"]
        for i, txt in enumerate(labels):
            ttk.Label(marco, text=txt, style="Cuerpo.TLabel").grid(row=i, column=0, sticky="e", padx=6, pady=4)

        self.cmb_equipo = ttk.Combobox(marco, values=[str(e) for e in ids_equipos], state="readonly", width=36, style="Combo.TCombobox")

        # Fecha: preferir DateEntry de tkcalendar
        if TKCAL_OK:
            self.e_fecha = DateEntry(marco, date_pattern='dd-mm-yyyy', locale='es_MX', width=18)
            self.e_fecha.grid(row=1, column=1, padx=6, pady=4, sticky="w")
        else:
            self.e_fecha = ttk.Entry(marco, width=27, style="Entrada.TEntry")
            poner_caret_blanco(self.e_fecha)
            btn_cal = ttk.Button(marco, text="📅", width=3, style="Fantasma.TButton",
                                 command=lambda: DatePicker(self, self.e_fecha, self.e_fecha.get().strip()))
            f_fecha = ttk.Frame(marco, style="Card.TFrame")
            f_fecha.grid(row=1, column=1, padx=6, pady=4, sticky="w")
            self.e_fecha.pack(in_=f_fecha, side="left")
            btn_cal.pack(in_=f_fecha, side="left", padx=6)

        self.cmb_tipo = ttk.Combobox(marco, values=["Preventivo", "Correctivo"], state="readonly", width=36, style="Combo.TCombobox")
        self.cmb_tipo.current(0)
        self.cmb_estado = ttk.Combobox(marco, values=["Pendiente", "Completado"], state="readonly", width=36, style="Combo.TCombobox")
        self.cmb_estado.current(0)
        self.e_prov = ttk.Entry(marco, width=38, style="Entrada.TEntry")
        self.e_costo = ttk.Entry(marco, width=38, style="Entrada.TEntry")
        self.t_notas = tk.Text(marco, width=38, height=4)
        poner_caret_blanco(self.e_prov, self.e_costo, self.t_notas)

        self.cmb_equipo.grid(row=0, column=1, padx=6, pady=4, sticky="w")
        if not TKCAL_OK:
            # ya colocado arriba en grid
            pass

        self.cmb_tipo.grid(row=2, column=1, padx=6, pady=4, sticky="w")
        self.cmb_estado.grid(row=3, column=1, padx=6, pady=4, sticky="w")
        self.e_prov.grid(row=4, column=1, padx=6, pady=4, sticky="w")
        self.e_costo.grid(row=5, column=1, padx=6, pady=4, sticky="w")
        self.t_notas.grid(row=6, column=1, padx=6, pady=4, sticky="w")

        if datos:
            if datos.get("equipo_id"):
                vals = self.cmb_equipo["values"]
                if str(datos["equipo_id"]) in vals:
                    self.cmb_equipo.set(str(datos["equipo_id"]))
            # fecha ISO -> DD-MM-AAAA
            if TKCAL_OK:
                try:
                    self.e_fecha.set_date(datetime.strptime(datos.get("fecha",""), "%Y-%m-%d").date())
                except Exception:
                    pass
            else:
                self.e_fecha.insert(0, _a_ddmmaaaa(datos.get("fecha", "")))
            if datos.get("tipo") in ["Preventivo", "Correctivo"]:
                self.cmb_tipo.set(datos["tipo"])
            if datos.get("estado") in ["Pendiente", "Completado"]:
                self.cmb_estado.set(datos["estado"])
            self.e_prov.insert(0, datos.get("proveedor", "") or "")
            self.e_costo.insert(0, str(datos.get("costo") if datos.get("costo") is not None else ""))
            self.t_notas.insert("1.0", datos.get("notas", "") or "")

        zona_botones = ttk.Frame(marco, style="Card.TFrame")
        zona_botones.grid(row=7, column=0, columnspan=2, pady=8)
        ttk.Button(zona_botones, text="Guardar", command=self._on_guardar, style="Success.TButton").pack(side="left", padx=6)
        ttk.Button(zona_botones, text="Cancelar", command=self.destroy, style="Fantasma.TButton").pack(side="left", padx=6)

        self.bind("<Return>", lambda e: self._on_guardar())
        self.grab_set()
        self.cmb_equipo.focus_set()

    def _on_guardar(self):
        eqid = self.cmb_equipo.get().strip()
        if not eqid:
            messagebox.showwarning("Validación", "Equipo (ID) es obligatorio.")
            return

        fecha_txt = self.e_fecha.get().strip()  # con DateEntry ya viene 'dd-mm-yyyy'
        if not fecha_txt:
            messagebox.showwarning("Validación", "La fecha es obligatoria.")
            return
        try:
            fecha_iso = _a_iso(fecha_txt)  # convierte a YYYY-MM-DD para la BD
        except ValueError:
            messagebox.showwarning("Validación", "La fecha debe ser DD-MM-AAAA.")
            return

        tipo = self.cmb_tipo.get()
        estado = self.cmb_estado.get()
        proveedor = self.e_prov.get().strip()
        costo_txt = self.e_costo.get().strip()
        try:
            costo_valor = float(costo_txt) if costo_txt else 0.0
        except ValueError:
            messagebox.showwarning("Validación", "Costo debe ser numérico.")
            return

        notas = self.t_notas.get("1.0", "end").strip()
        auto_id = f"M-{datetime.now().strftime('%Y%m%d%H%M%S')}-{uuid.uuid4().hex[:6]}"
        self.resultado = {
            "id_mantenimiento": auto_id,
            "equipo_id": eqid,
            "fecha": fecha_iso,  # ISO para DB
            "tipo": tipo,
            "estado": estado,
            "proveedor": proveedor,
            "costo": costo_valor,
            "notas": notas,
            "registrado_en": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        self.destroy()


class DialogoUsuario(tk.Toplevel):
    def __init__(self, master, titulo="Usuario", datos=None):
        super().__init__(master)
        self.title(titulo)
        aplicar_icono_aplicacion(self)
        self.resizable(False, False)
        self.resultado = None

        marco = ttk.Frame(self, style="Card.TFrame", padding=12)
        marco.grid(row=0, column=0, sticky="nsew")

        ttk.Label(marco, text="Usuario *", style="Cuerpo.TLabel").grid(row=0, column=0, sticky="e", padx=6, pady=4)
        ttk.Label(marco, text="Contraseña *", style="Cuerpo.TLabel").grid(row=1, column=0, sticky="e", padx=6, pady=4)
        ttk.Label(marco, text="Rol *", style="Cuerpo.TLabel").grid(row=2, column=0, sticky="e", padx=6, pady=4)

        self.e_user = ttk.Entry(marco, width=36, style="Entrada.TEntry")
        self.e_pass = ttk.Entry(marco, width=36, style="Entrada.TEntry", show="*")
        self.cmb_rol = ttk.Combobox(marco, values=["administrador", "laboratorista"], state="readonly", width=34, style="Combo.TCombobox")
        self.cmb_rol.current(1)
        poner_caret_blanco(self.e_user, self.e_pass)

        self.e_user.grid(row=0, column=1, padx=6, pady=4, sticky="w")
        self.e_pass.grid(row=1, column=1, padx=6, pady=4, sticky="w")
        self.cmb_rol.grid(row=2, column=1, padx=6, pady=4, sticky="w")

        if datos:
            self.e_user.insert(0, datos.get("usuario", ""))
            if datos.get("rol") in ["administrador", "laboratorista"]:
                self.cmb_rol.set(datos["rol"])

        zona_botones = ttk.Frame(marco, style="Card.TFrame")
        zona_botones.grid(row=3, column=0, columnspan=2, pady=8)
        ttk.Button(zona_botones, text="Guardar", command=self._on_guardar, style="Success.TButton").pack(side="left", padx=6)
        ttk.Button(zona_botones, text="Cancelar", command=self.destroy, style="Fantasma.TButton").pack(side="left", padx=6)

        self.bind("<Return>", lambda e: self._on_guardar())
        self.grab_set()
        self.e_user.focus_set()

    def _on_guardar(self):
        usuario = self.e_user.get().strip()
        contrasena = self.e_pass.get().strip()
        rol = self.cmb_rol.get().strip()
        if not usuario or not rol:
            messagebox.showwarning("Validación", "Usuario y rol son obligatorios.")
            return
        self.resultado = {"usuario": usuario, "contrasena": contrasena if contrasena else None, "rol": rol}
        self.destroy()


class DialogoConfigAlertas(tk.Toplevel):
    def __init__(self, master, dia_actual: int, preaviso_actual: int):
        super().__init__(master)
        self.title("Programar alertas mensuales")
        aplicar_icono_aplicacion(self)
        self.resizable(False, False)
        self.resultado = None

        marco = ttk.Frame(self, style="Card.TFrame", padding=12)
        marco.grid(row=0, column=0, sticky="nsew")

        ttk.Label(marco, text="Día del mes para alerta principal (1–28)", style="Cuerpo.TLabel").grid(
            row=0, column=0, padx=8, pady=6, sticky="e")
        self.sp_dia = tk.Spinbox(marco, from_=1, to=28, width=6)
        self.sp_dia.grid(row=0, column=1, padx=8, pady=6, sticky="w")

        ttk.Label(marco, text="Días ANTES del fin de mes para alerta previa", style="Cuerpo.TLabel").grid(
            row=1, column=0, padx=8, pady=6, sticky="e")
        self.sp_pre = tk.Spinbox(marco, from_=1, to=27, width=6)
        self.sp_pre.grid(row=1, column=1, padx=8, pady=6, sticky="w")

        info = ("Ejemplo recomendado:\n"
                "- Día del mes = 1 (aviso principal el 1° de cada mes)\n"
                "- Días antes fin de mes = 5 (aviso el 26–31 según el mes)")
        ttk.Label(marco, text=info, style="Cuerpo.TLabel", justify="left").grid(
            row=2, column=0, columnspan=2, padx=8, pady=6)

        zona_botones = ttk.Frame(marco, style="Card.TFrame")
        zona_botones.grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(zona_botones, text="Guardar", command=self._on_guardar, style="Success.TButton").pack(side="left", padx=6)
        ttk.Button(zona_botones, text="Cancelar", command=self.destroy, style="Fantasma.TButton").pack(side="left", padx=6)

        self.sp_dia.delete(0, "end"); self.sp_dia.insert(0, dia_actual)
        self.sp_pre.delete(0, "end"); self.sp_pre.insert(0, preaviso_actual)
        self.grab_set()

    def _on_guardar(self):
        try:
            dia = int(self.sp_dia.get())
            pre = int(self.sp_pre.get())
            if not (1 <= dia <= 28): raise ValueError
            if not (1 <= pre <= 27): raise ValueError
        except Exception:
            messagebox.showwarning("Validación", "Rangos válidos: día (1–28), preaviso (1–27).")
            return
        self.resultado = (dia, pre)
        self.destroy()


# ============================================================
# 9) PESTAÑAS
# ============================================================
class PestanaEquipos(ttk.Frame):
    """Pestaña de gestión de equipos."""
    def __init__(self, padre, con, usuario_actual):
        super().__init__(padre)
        self.con = con
        self.usuario_actual = usuario_actual

        superior = ttk.Frame(self, style="App.TFrame")
        superior.pack(fill="x")
        self.lbl_total = ttk.Label(superior, text="Total de equipos: 0", style="Cuerpo.TLabel")
        self.lbl_total.pack(side="left", padx=8, pady=6)

        self.arbol = ttk.Treeview(
            self,
            columns=("id_equipo", "nombre", "marca", "modelo", "serie", "ubicacion", "descripcion", "fecha_registro", "creado_por"),
            show="headings", height=14
        )
        cabeceras = {
            "id_equipo": "ID equipo", "nombre": "Nombre", "marca": "Marca",
            "modelo": "Modelo", "serie": "Serie", "ubicacion": "Ubicación",
            "descripcion": "Descripción", "fecha_registro": "Fecha registro", "creado_por": "Registrado por (ID)"
        }
        anchos = {"id_equipo": 120, "nombre": 160, "marca": 120, "modelo": 120, "serie": 120,
                  "ubicacion": 140, "descripcion": 220, "fecha_registro": 140, "creado_por": 130}
        for c in self.arbol["columns"]:
            self.arbol.heading(c, text=cabeceras[c])
            self.arbol.column(c, width=anchos[c], anchor="w")
        self.arbol.pack(fill="both", expand=True, padx=8, pady=6)

        zona_botones = ttk.Frame(self, style="App.TFrame")
        zona_botones.pack(pady=4)
        ttk.Button(zona_botones, text="Agregar", command=self._agregar, style="Primario.TButton").pack(side="left", padx=4)
        ttk.Button(zona_botones, text="Editar", command=self._editar, style="Fantasma.TButton").pack(side="left", padx=4)

        self.btn_borrar = ttk.Button(zona_botones, text="Eliminar", command=self._eliminar, style="Danger.TButton")
        self.btn_borrar.pack(side="left", padx=4)
        if self.usuario_actual["rol"] != "administrador":
            self.btn_borrar.state(["disabled"])

        ttk.Button(zona_botones, text="Importar Excel", command=self._importar_excel, style="Fantasma.TButton").pack(side="left", padx=12)
        ttk.Button(zona_botones, text="Exportar Excel", command=self._exportar_excel, style="Fantasma.TButton").pack(side="left", padx=4)

        self._refrescar()

    def _refrescar(self):
        for i in self.arbol.get_children():
            self.arbol.delete(i)
        cur = self.con.execute("""
            SELECT id_equipo, nombre, marca, modelo, serie, ubicacion,
                   COALESCE(descripcion,''), COALESCE(fecha_registro,''), COALESCE(creado_por,'')
            FROM equipos ORDER BY nombre ASC;
        """)
        filas = cur.fetchall()
        for f in filas:
            self.arbol.insert("", "end", values=f)
        self.lbl_total.config(text=f"Total de equipos: {len(filas)}")

    def _seleccionado(self):
        sel = self.arbol.selection()
        if not sel:
            messagebox.showinfo("Info", "Seleccione un equipo.")
            return None
        valores = self.arbol.item(sel[0], "values")
        claves = ["id_equipo", "nombre", "marca", "modelo", "serie", "ubicacion", "descripcion", "fecha_registro", "creado_por"]
        return dict(zip(claves, valores))

    def _agregar(self):
        dlg = DialogoEquipo(self, "Agregar equipo", usuario_actual=self.usuario_actual)
        self.wait_window(dlg)
        if dlg.resultado:
            try:
                with self.con:
                    self.con.execute("""
                        INSERT INTO equipos(id_equipo, nombre, marca, modelo, serie, ubicacion, descripcion, fecha_registro, creado_por)
                        VALUES(?,?,?,?,?,?,?,?,?)
                    """, (
                        dlg.resultado["id_equipo"], dlg.resultado["nombre"], dlg.resultado["marca"],
                        dlg.resultado["modelo"], dlg.resultado["serie"], dlg.resultado["ubicacion"],
                        dlg.resultado["descripcion"], dlg.resultado["fecha_registro"], self.usuario_actual["id"]
                    ))
                self._refrescar()
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", "El ID de equipo ya existe. Usa otro.")

    def _editar(self):
        datos = self._seleccionado()
        if not datos:
            return
        dlg = DialogoEquipo(self, "Editar equipo", datos=datos, usuario_actual=self.usuario_actual)
        self.wait_window(dlg)
        if dlg.resultado:
            try:
                with self.con:
                    self.con.execute("""
                        UPDATE equipos
                        SET id_equipo=?, nombre=?, marca=?, modelo=?, serie=?, ubicacion=?, descripcion=?
                        WHERE id_equipo=?
                    """, (
                        dlg.resultado["id_equipo"], dlg.resultado["nombre"], dlg.resultado["marca"],
                        dlg.resultado["modelo"], dlg.resultado["serie"], dlg.resultado["ubicacion"],
                        dlg.resultado["descripcion"], datos["id_equipo"]
                    ))
                self._refrescar()
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", "El nuevo ID de equipo ya existe. Usa otro.")

    def _eliminar(self):
        if self.usuario_actual["rol"] != "administrador":
            return
        datos = self._seleccionado()
        if not datos:
            return
        if not messagebox.askyesno("Confirmar", f"¿Eliminar equipo '{datos['nombre']}' (ID {datos['id_equipo']}) y sus mantenimientos?"):
            return
        with self.con:
            self.con.execute("DELETE FROM equipos WHERE id_equipo=?", (datos["id_equipo"],))
        self._refrescar()

    # Excel
    def _importar_excel(self):
        if not PANDAS_OK:
            messagebox.showwarning("Excel", "Instala pandas y openpyxl:  pip install pandas openpyxl")
            return
        archivo = filedialog.askopenfilename(title="Selecciona Excel de equipos",
                                             filetypes=[("Excel", "*.xlsx *.xls")])
        if not archivo:
            return
        try:
            df = pd.read_excel(archivo)
            df.columns = [str(c).strip().lower() for c in df.columns]
            df = df.dropna(how="all")  # Elimina filas completamente vacías
        except Exception as e:
            messagebox.showerror("Excel", f"No se pudo leer el archivo:\n{e}")
            return

        for col in ["id_equipo", "nombre", "marca", "modelo", "serie", "ubicacion", "descripcion"]:
            if col not in df.columns:
                df[col] = ""

        nuevos = 0
        with self.con:
            for _, r in df.iterrows():
                id_eq = str(r["id_equipo"]).strip()
                nom = str(r["nombre"]).strip()
                if not id_eq or not nom:
                    continue
                fecha_reg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                try:
                    self.con.execute("""
                        INSERT INTO equipos(id_equipo, nombre, marca, modelo, serie, ubicacion, descripcion, fecha_registro, creado_por)
                        VALUES(?,?,?,?,?,?,?,?,?)
                    """, (
                        id_eq, nom, str(r.get("marca", "") or ""), str(r.get("modelo", "") or ""),
                        str(r.get("serie", "") or ""), str(r.get("ubicacion", "") or ""),
                        str(r.get("descripcion", "") or ""), fecha_reg, self.usuario_actual["id"]
                    ))
                    nuevos += 1
                except sqlite3.IntegrityError:
                    pass
        self._refrescar()
        messagebox.showinfo("Excel", f"Equipos importados. Nuevos: {nuevos}")

    def _exportar_excel(self):
        if not PANDAS_OK:
            messagebox.showwarning("Excel", "Instala pandas y openpyxl:  pip install pandas openpyxl")
            return
        archivo = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=[("Excel", "*.xlsx")],
                                               title="Guardar inventario")
        if not archivo:
            return
        try:
            df = pd.read_sql_query("""
                SELECT id_equipo, nombre, marca, modelo, serie, ubicacion, descripcion, fecha_registro, creado_por
                FROM equipos ORDER BY nombre;
            """, self.con)
            df.to_excel(archivo, index=False)
            messagebox.showinfo("Excel", "Inventario exportado correctamente.")
        except Exception as e:
            messagebox.showerror("Excel", f"No se pudo exportar:\n{e}")


class PestanaMantenimientos(ttk.Frame):
    """Pestaña de gestión de mantenimientos."""
    def __init__(self, padre, con, usuario_actual):
        super().__init__(padre)
        self.con = con
        self.usuario_actual = usuario_actual

        self.arbol = ttk.Treeview(
            self,
            columns=("id_mantenimiento", "equipo_id", "fecha", "tipo", "estado", "proveedor", "costo", "notas", "registrado_en", "creado_por"),
            show="headings", height=14
        )
        cabeceras = {
            "id_mantenimiento": "ID (oculto)", "equipo_id": "Equipo (ID)", "fecha": "Fecha",
            "tipo": "Tipo", "estado": "Estado", "proveedor": "Proveedor", "costo": "Costo",
            "notas": "Notas", "registrado_en": "Registrado en", "creado_por": "Registrado por (ID)"
        }
        anchos = {"id_mantenimiento": 0, "equipo_id": 140, "fecha": 110, "tipo": 110,
                  "estado": 110, "proveedor": 160, "costo": 100, "notas": 260,
                  "registrado_en": 150, "creado_por": 140}
        for c in self.arbol["columns"]:
            self.arbol.heading(c, text=cabeceras[c])
            self.arbol.column(c, width=anchos[c], anchor="w")
        self.arbol.column("id_mantenimiento", width=0, stretch=False, anchor="w")  # oculto
        self.arbol.pack(fill="both", expand=True, padx=8, pady=6)

        zona_botones = ttk.Frame(self, style="App.TFrame")
        zona_botones.pack(pady=4)
        ttk.Button(zona_botones, text="Agregar", command=self._agregar, style="Primario.TButton").pack(side="left", padx=4)
        ttk.Button(zona_botones, text="Editar", command=self._editar, style="Fantasma.TButton").pack(side="left", padx=4)
        ttk.Button(zona_botones, text="Completado", command=lambda: self._cambiar_estado("Completado"), style="Success.TButton").pack(side="left", padx=4)
        ttk.Button(zona_botones, text="Pendiente", command=lambda: self._cambiar_estado("Pendiente"), style="Fantasma.TButton").pack(side="left", padx=4)
        ttk.Button(zona_botones, text="Exportar Excel", command=self._exportar_excel, style="Fantasma.TButton").pack(side="left", padx=12)
        ttk.Button(zona_botones, text="Importar Excel", command=self._importar_excel, style="Fantasma.TButton").pack(side="left", padx=4)

        self._refrescar()

    def _ids_equipos(self):
        return [row[0] for row in self.con.execute("SELECT id_equipo FROM equipos ORDER BY id_equipo;").fetchall()]

    def _refrescar(self):
        for i in self.arbol.get_children():
            self.arbol.delete(i)
        q = """
        SELECT m.id_mantenimiento, m.equipo_id, m.fecha, m.tipo, m.estado,
               COALESCE(m.proveedor,''), COALESCE(m.costo,0), COALESCE(m.notas,''),
               COALESCE(m.registrado_en,''), COALESCE(m.creado_por,'')
        FROM mantenimientos m
        ORDER BY m.fecha DESC, m.id_mantenimiento DESC;
        """
        for fila in self.con.execute(q).fetchall():
            # fila: (id_mant, equipo_id, fecha_iso, tipo, estado, proveedor, costo, notas, registrado_en, creado_por)
            fila = list(fila)
            fila[2] = _a_ddmmaaaa(fila[2])  # mostrar DD-MM-AAAA
            self.arbol.insert("", "end", values=tuple(fila))

    def _id_seleccionado(self):
        sel = self.arbol.selection()
        if not sel:
            messagebox.showinfo("Info", "Seleccione un mantenimiento.")
            return None
        valores = self.arbol.item(sel[0], "values")
        return str(valores[0])

    def _agregar(self):
        ids = self._ids_equipos()
        if not ids:
            messagebox.showinfo("Info", "Agregue equipos primero.")
            return
        dlg = DialogoMantenimiento(self, "Agregar mantenimiento", ids_equipos=ids)
        self.wait_window(dlg)
        if dlg.resultado:
            try:
                # --- NUEVO: Mover registro anterior a historicos si existe para el mismo equipo_id ---
                cur = self.con.execute("""
                    SELECT * FROM mantenimientos WHERE equipo_id=? ORDER BY fecha DESC LIMIT 1
                """, (dlg.resultado["equipo_id"],))
                anterior = cur.fetchone()
                if anterior:
                    # Mapear columnas
                    columnas = [desc[0] for desc in cur.description]
                    datos_ant = dict(zip(columnas, anterior))
                    id_historico = f"H-{datetime.now().strftime('%Y%m%d%H%M%S')}-{uuid.uuid4().hex[:6]}"
                    self.con.execute("""
                        INSERT INTO historicos(
                            id_historico, id_mantenimiento, equipo_id, fecha, tipo, notas, estado, proveedor, costo, creado_por, registrado_en
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        id_historico,
                        datos_ant["id_mantenimiento"],
                        datos_ant["equipo_id"],
                        datos_ant["fecha"],
                        datos_ant["tipo"],
                        datos_ant["notas"],
                        datos_ant["estado"],
                        datos_ant["proveedor"],
                        datos_ant["costo"],
                        datos_ant["creado_por"],
                        datos_ant["registrado_en"]
                    ))
                    # Eliminar el anterior de mantenimientos
                    self.con.execute("DELETE FROM mantenimientos WHERE id_mantenimiento=?", (datos_ant["id_mantenimiento"],))
                # --- FIN NUEVO ---

                with self.con:
                    self.con.execute("""
                        INSERT INTO mantenimientos(id_mantenimiento, equipo_id, fecha, tipo, notas, estado, proveedor, costo, creado_por, registrado_en)
                        VALUES(?,?,?,?,?,?,?,?,?,?)
                    """, (
                        dlg.resultado["id_mantenimiento"], dlg.resultado["equipo_id"], dlg.resultado["fecha"],
                        dlg.resultado["tipo"], dlg.resultado["notas"], dlg.resultado["estado"],
                        dlg.resultado["proveedor"], dlg.resultado["costo"], self.usuario_actual["id"],
                        dlg.resultado["registrado_en"]
                    ))
                self._refrescar()
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", "ID interno duplicado o equipo inexistente.")

    def _editar(self):
        mid = self._id_seleccionado()
        if not mid:
            return
        fila = self.con.execute("""
            SELECT id_mantenimiento, equipo_id, fecha, tipo, estado,
                   COALESCE(proveedor,''), COALESCE(costo,0), COALESCE(notas,'')
            FROM mantenimientos WHERE id_mantenimiento=?
        """, (mid,)).fetchone()
        if not fila:
            return
        datos = {
            "id_mantenimiento": fila[0], "equipo_id": fila[1], "fecha": fila[2],
            "tipo": fila[3], "estado": fila[4], "proveedor": fila[5], "costo": fila[6], "notas": fila[7]
        }
        ids = self._ids_equipos()
        dlg = DialogoMantenimiento(self, "Editar mantenimiento", ids_equipos=ids, datos=datos)
        self.wait_window(dlg)
        if dlg.resultado:
            try:
                with self.con:
                    self.con.execute("""
                        UPDATE mantenimientos
                        SET equipo_id=?, fecha=?, tipo=?, notas=?, estado=?, proveedor=?, costo=?
                        WHERE id_mantenimiento=?
                    """, (
                        dlg.resultado["equipo_id"], dlg.resultado["fecha"], dlg.resultado["tipo"],
                        dlg.resultado["notas"], dlg.resultado["estado"], dlg.resultado["proveedor"],
                        dlg.resultado["costo"], mid
                    ))
                self._refrescar()
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", "Datos inválidos o equipo inexistente.")

    def _cambiar_estado(self, estado: str):
        mid = self._id_seleccionado()
        if not mid:
            return
        with self.con:
            self.con.execute("UPDATE mantenimientos SET estado=? WHERE id_mantenimiento=?", (estado, mid))
        self._refrescar()

    # Excel
    def _exportar_excel(self):
        if not PANDAS_OK:
            messagebox.showwarning("Excel", "Instala pandas y openpyxl:  pip install pandas openpyxl")
            return
        archivo = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=[("Excel", "*.xlsx")],
                                               title="Guardar mantenimientos")
        if not archivo:
            return
        consulta = """
        SELECT id_mantenimiento, equipo_id, fecha, tipo, estado, proveedor, costo, notas, registrado_en, creado_por
        FROM mantenimientos
        ORDER BY fecha DESC;
        """
        try:
            df = pd.read_sql_query(consulta, self.con)
            df.to_excel(archivo, index=False)
            messagebox.showinfo("Excel", "Mantenimientos exportados correctamente.")
        except Exception as e:
            messagebox.showerror("Excel", f"No se pudo exportar:\n{e}")

    def _importar_excel(self):
        if not PANDAS_OK:
            messagebox.showwarning("Excel", "Instala pandas y openpyxl:  pip install pandas openpyxl")
            return
        archivo = filedialog.askopenfilename(title="Selecciona Excel de mantenimientos",
                                             filetypes=[("Excel", "*.xlsx *.xls")])
        if not archivo:
            return
        try:
            df = pd.read_excel(archivo)
            df.columns = [str(c).strip().lower() for c in df.columns]
            df = df.dropna(how="all")  # Elimina filas completamente vacías
        except Exception as e:
            messagebox.showerror("Excel", f"No se pudo leer el archivo:\n{e}")
            return

        for col in ["id_mantenimiento", "equipo_id", "fecha", "tipo", "estado", "proveedor", "costo", "notas"]:
            if col not in df.columns:
                df[col] = "" if col != "costo" else 0

        nuevos = 0
        with self.con:
            for _, r in df.iterrows():
                idm = str(r["id_mantenimiento"]).strip() or f"IMP-{uuid.uuid4().hex[:10]}"
                eqid = str(r["equipo_id"]).strip()
                if not eqid:
                    continue

                # Fecha normalizada: acepta DD-MM-AAAA o ISO o Timestamp
                fecha_val = r["fecha"]
                fecha_str = ""
                if PANDAS_OK and isinstance(fecha_val, (pd.Timestamp,)):
                    fecha_str = fecha_val.strftime("%Y-%m-%d")
                else:
                    txt = str(fecha_val).strip()
                    if not txt:
                        continue
                    try:
                        fecha_str = _a_iso(txt)  # si viene DD-MM-AAAA
                    except Exception:
                        try:
                            datetime.strptime(txt, "%Y-%m-%d")
                            fecha_str = txt
                        except Exception:
                            continue

                tipo = str(r["tipo"])
                if tipo not in ["Preventivo", "Correctivo"]:
                    tipo = "Preventivo"
                estado = str(r["estado"])
                if estado not in ["Pendiente", "Completado"]:
                    estado = "Pendiente"
                proveedor = str(r.get("proveedor", "") or "")
                try:
                    costo = float(r.get("costo", 0) or 0)
                except Exception:
                    costo = 0.0
                notas = str(r.get("notas", "") or "")
                registrado_en = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                try:
                    self.con.execute("""
                        INSERT INTO mantenimientos(id_mantenimiento, equipo_id, fecha, tipo, notas, estado, proveedor, costo, creado_por, registrado_en)
                        VALUES(?,?,?,?,?,?,?,?,?,?)
                    """, (idm, eqid, fecha_str, tipo, notas, estado, proveedor, costo, self.usuario_actual["id"], registrado_en))
                    nuevos += 1
                except sqlite3.IntegrityError:
                    pass
        self._refrescar()
        messagebox.showinfo("Excel", f"Mantenimientos importados. Nuevos: {nuevos}")


class PestanaHistoricos(ttk.Frame):
    """Pestaña de visualización de históricos de mantenimientos."""
    def __init__(self, padre, con):
        super().__init__(padre)
        self.con = con

        self.arbol = ttk.Treeview(
            self,
            columns=("id_historico", "id_mantenimiento", "equipo_id", "fecha", "tipo", "estado", "proveedor", "costo", "notas", "registrado_en", "creado_por"),
            show="headings", height=14
        )
        cabeceras = {
            "id_historico": "ID Hist.",
            "id_mantenimiento": "ID Mant.",
            "equipo_id": "Equipo (ID)",
            "fecha": "Fecha",
            "tipo": "Tipo",
            "estado": "Estado",
            "proveedor": "Proveedor",
            "costo": "Costo",
            "notas": "Notas",
            "registrado_en": "Registrado en",
            "creado_por": "Registrado por (ID)"
        }
        anchos = {
            "id_historico": 0, "id_mantenimiento": 0, "equipo_id": 140, "fecha": 110, "tipo": 110,
            "estado": 110, "proveedor": 160, "costo": 100, "notas": 260, "registrado_en": 150, "creado_por": 140
        }
        for c in self.arbol["columns"]:
            self.arbol.heading(c, text=cabeceras[c])
            self.arbol.column(c, width=anchos.get(c, 100), anchor="w")
        self.arbol.column("id_historico", width=0, stretch=False, anchor="w")  # oculto
        self.arbol.column("id_mantenimiento", width=0, stretch=False, anchor="w")  # oculto
        self.arbol.pack(fill="both", expand=True, padx=8, pady=6)

        self._refrescar()

    def _refrescar(self):
        for i in self.arbol.get_children():
            self.arbol.delete(i)
        q = """
        SELECT id_historico, id_mantenimiento, equipo_id, fecha, tipo, estado, proveedor, costo, notas, registrado_en, creado_por
        FROM historicos
        ORDER BY fecha DESC, id_historico DESC;
        """
        for fila in self.con.execute(q).fetchall():
            fila = list(fila)
            # fecha ISO -> DD-MM-AAAA
            fila[3] = _a_ddmmaaaa(fila[3])
            self.arbol.insert("", "end", values=tuple(fila))


class PestanaUsuarios(ttk.Frame):
    """Pestaña de gestión de usuarios."""
    def __init__(self, padre, con, usuario_actual):
        super().__init__(padre)
        self.con = con
        self.usuario_actual = usuario_actual

        self.arbol = ttk.Treeview(self, columns=("id", "usuario", "rol"), show="headings", height=12)
        for col, txt, ancho in [("id", "ID", 60), ("usuario", "Usuario", 260), ("rol", "Rol", 160)]:
            self.arbol.heading(col, text=txt)
            self.arbol.column(col, width=ancho, anchor="w")
        self.arbol.pack(fill="both", expand=True, padx=8, pady=8)

        zona_botones = ttk.Frame(self, style="App.TFrame")
        zona_botones.pack(pady=4)
        ttk.Button(zona_botones, text="Agregar", command=self._agregar, style="Primario.TButton").pack(side="left", padx=4)
        ttk.Button(zona_botones, text="Editar", command=self._editar, style="Fantasma.TButton").pack(side="left", padx=4)
        ttk.Button(zona_botones, text="Eliminar", command=self._eliminar, style="Danger.TButton").pack(side="left", padx=4)

        self._refrescar()

    def _refrescar(self):
        for i in self.arbol.get_children():
            self.arbol.delete(i)
        cur = self.con.execute("SELECT id, usuario, rol FROM usuarios ORDER BY id;")
        for fila in cur.fetchall():
            self.arbol.insert("", "end", values=fila)

    def _seleccionado(self):
        sel = self.arbol.selection()
        if not sel:
            messagebox.showinfo("Info", "Seleccione un usuario.")
            return None
        vals = self.arbol.item(sel[0], "values")
        return {"id": int(vals[0]), "usuario": vals[1], "rol": vals[2]}

    def _agregar(self):
        dlg = DialogoUsuario(self, "Agregar usuario")
        self.wait_window(dlg)
        if dlg.resultado:
            usuario = dlg.resultado["usuario"]
            contrasena = dlg.resultado["contrasena"]
            rol = dlg.resultado["rol"]
            if not contrasena:
                messagebox.showwarning("Validación", "La contraseña es obligatoria.")
                return
            try:
                crear_usuario(self.con, usuario, contrasena, rol)
                self._refrescar()
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", "El usuario ya existe.")

    def _editar(self):
        datos = self._seleccionado()
        if not datos:
            return
        dlg = DialogoUsuario(self, "Editar usuario", datos=datos)
        self.wait_window(dlg)
        if dlg.resultado:
            usuario = dlg.resultado["usuario"]
            rol = dlg.resultado["rol"]
            nueva_pwd = dlg.resultado["contrasena"]
            with self.con:
                try:
                    self.con.execute("UPDATE usuarios SET usuario=?, rol=? WHERE id=?", (usuario, rol, datos["id"]))
                except sqlite3.IntegrityError:
                    messagebox.showerror("Error", "Nombre de usuario duplicado.")
                    return
                if nueva_pwd:
                    sal = secrets.token_hex(16)
                    hash_pwd = crear_hash(nueva_pwd, sal)
                    self.con.execute("UPDATE usuarios SET hash_contrasena=?, sal=? WHERE id=?",
                                     (hash_pwd, sal, datos["id"]))
            self._refrescar()

    def _eliminar(self):
        datos = self._seleccionado()
        if not datos:
            return
        if datos["id"] == self.usuario_actual["id"]:
            messagebox.showwarning("Restricción", "No puedes eliminar tu propio usuario conectado.")
            return
        if not messagebox.askyesno("Confirmar", f"¿Eliminar usuario '{datos['usuario']}'?"):
            return
        with self.con:
            self.con.execute("DELETE FROM usuarios WHERE id=?", (datos["id"],))
        self._refrescar()


# ============================================================
# 10) VENTANAS PRINCIPALES (LOGIN con Frame centrado + APP)
# ============================================================
class VentanaInicioSesion(tk.Tk):
    """Ventana de inicio de sesión (fullscreen, tarjeta centrada)."""
    def __init__(self, con):
        super().__init__()
        self.con = con
        self.title("Login - Control de Mantenimientos")
        aplicar_estilo_global()
        aplicar_icono_aplicacion(self)
        habilitar_pantalla_completa(self)

        # Contenedor raíz
        cont = ttk.Frame(self, style="App.TFrame")
        cont.pack(fill="both", expand=True)

        # Tarjeta centrada
        tarjeta = ttk.Frame(cont, style="Card.TFrame", padding=40)
        tarjeta.place(relx=0.5, rely=0.5, anchor="center")
        tarjeta.configure(width=560)

        ttk.Label(tarjeta, text="Control de Mantenimientos", style="Titulo.TLabel").grid(
            row=0, column=0, columnspan=2, pady=(0, 20))

        ttk.Label(tarjeta, text="Usuario", style="Cuerpo.TLabel").grid(
            row=1, column=0, sticky="e", padx=(0, 12), pady=10)
        self.e_usuario = ttk.Entry(tarjeta, width=30, style="Entrada.TEntry")
        self.e_usuario.grid(row=1, column=1, sticky="w", pady=10)

        ttk.Label(tarjeta, text="Contraseña", style="Cuerpo.TLabel").grid(
            row=2, column=0, sticky="e", padx=(0, 12), pady=10)
        self.e_contra = ttk.Entry(tarjeta, width=30, style="Entrada.TEntry", show="*")
        self.e_contra.grid(row=2, column=1, sticky="w", pady=10)

        poner_caret_blanco(self.e_usuario, self.e_contra)

        zona_botones = ttk.Frame(tarjeta, style="Card.TFrame")
        zona_botones.grid(row=3, column=0, columnspan=2, pady=(18, 0))
        ttk.Button(zona_botones, text="Ingresar", command=self._iniciar_sesion, style="Primario.TButton").grid(row=0, column=0, padx=8)
        ttk.Button(zona_botones, text="Salir", command=self.destroy, style="Fantasma.TButton").grid(row=0, column=1, padx=8)

        self.bind("<Return>", lambda e: self._iniciar_sesion())
        self.e_usuario.focus_set()

        for i in range(2):
            tarjeta.grid_columnconfigure(i, weight=1)

    def _iniciar_sesion(self):
        usuario = self.e_usuario.get().strip()
        contrasena = self.e_contra.get().strip()
        user = verificar_usuario(self.con, usuario, contrasena)
        if not user:
            messagebox.showerror("Acceso denegado", "Usuario o contraseña incorrectos.")
            return
        self.destroy()
        app = AplicacionPrincipal(self.con, user)
        app.mainloop()


class AplicacionPrincipal(tk.Tk):
    """Ventana principal (inicia en fullscreen)."""
    def __init__(self, con, usuario_actual):
        super().__init__()
        self.con = con
        self.usuario_actual = usuario_actual

        self.title("Control de Mantenimientos")
        aplicar_estilo_global()
        aplicar_icono_aplicacion(self)
        habilitar_pantalla_completa(self)

        # Barra superior
        barra = ttk.Frame(self, style="App.TFrame")
        barra.pack(fill="x")
        ttk.Label(
            barra,
            text=f"Usuario: {usuario_actual['usuario']}  |  Rol: {usuario_actual['rol']}",
            style="Cuerpo.TLabel"
        ).pack(side="left", padx=10, pady=8)

        ttk.Button(barra, text="Comprobar alertas ahora",
                   command=lambda: revisar_alertas(self.con, forzar=True),
                   style="Fantasma.TButton").pack(side="right", padx=6)
        ttk.Button(barra, text="Configurar alertas",
                   command=self._configurar_alertas,
                   style="Fantasma.TButton").pack(side="right", padx=6)
        ttk.Button(barra, text="Pantalla completa (F11/ESC)",
                   command=lambda: self.event_generate("<F11>"),
                   style="Fantasma.TButton").pack(side="right", padx=6)
        ttk.Button(barra, text="Cerrar sesión",
                   command=self._cerrar_sesion,
                   style="Fantasma.TButton").pack(side="right", padx=6)

        # Notebook
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=8)

        self.tab_equipos = PestanaEquipos(self.nb, self.con, self.usuario_actual)
        self.nb.add(self.tab_equipos, text="Equipos")

        self.tab_mants = PestanaMantenimientos(self.nb, self.con, self.usuario_actual)
        self.nb.add(self.tab_mants, text="Mantenimientos")

        # --- NUEVO: Pestaña de Históricos ---
        self.tab_historicos = PestanaHistoricos(self.nb, self.con)
        self.nb.add(self.tab_historicos, text="Históricos")
        # --- FIN NUEVO ---

        if self.usuario_actual["rol"] == "administrador":
            self.tab_usuarios = PestanaUsuarios(self.nb, self.con, self.usuario_actual)
            self.nb.add(self.tab_usuarios, text="Usuarios")

        # Revisar alertas al iniciar
        self.after(300, lambda: revisar_alertas(self.con, forzar=False))

        # Menú
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        menu_cfg = tk.Menu(menubar, tearoff=0)
        menu_cfg.add_command(label="Configurar alertas", command=self._configurar_alertas)
        menu_cfg.add_command(label="Comprobar alertas ahora", command=lambda: revisar_alertas(self.con, forzar=True))
        menubar.add_cascade(label="Configuración", menu=menu_cfg)

    def _configurar_alertas(self):
        dia = int(obtener_ajuste(self.con, "dia_mantenimiento", "1") or "1")
        pre = int(obtener_ajuste(self.con, "preaviso_dias_fin_mes", "5") or "5")
        dlg = DialogoConfigAlertas(self, dia_actual=dia, preaviso_actual=pre)
        self.wait_window(dlg)
        if dlg.resultado:
            dia_nuevo, pre_nuevo = dlg.resultado
           
            establecer_ajuste(self.con, "dia_mantenimiento", str(dia_nuevo))
            establecer_ajuste(self.con, "preaviso_dias_fin_mes", str(pre_nuevo))
            messagebox.showinfo("Alertas", f"Configurado: día={dia_nuevo}, aviso previo={pre_nuevo} días antes de fin de mes.")

    def _cerrar_sesion(self):
        self.destroy()
        ejecutar_login(self.con)


# ============================================================
# 11) ARRANQUE
# ============================================================
def ejecutar_login(con):
    login = VentanaInicioSesion(con)
    login.mainloop()


if __name__ == "__main__":
    conexion = iniciar_bd()
    ejecutar_login(conexion)
