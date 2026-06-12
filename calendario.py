import os
import sys
import tkinter as tk
from tkinter import ttk
from datetime import datetime
from dateutil.relativedelta import relativedelta
from tkcalendar import Calendar

class CalendarioEventosApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Calendario de Eventos")
        
        # Configuración de ventana sin bordes (Estilo Widget Windows 11)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(bg="#f3f3f3")

        # Banderas de control
        self.bloqueo_sincronizacion = False
        self.esta_fijado = False  # Estado inicial: No fijado (Cierre inteligente activo)

        # 1. Cargar Configuración de Colores y Eventos desde el mismo Excel
        self.colores_por_tipo = self.cargar_configuracion_colores()
        self.eventos = self.cargar_eventos()

        # Contenedor principal
        self.main_frame = ttk.Frame(self, padding=10)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # --- BOTÓN DE PIN / CHINCHETA (Esquina Superior Derecha) ---
        top_bar = tk.Frame(self.main_frame, bg="#f3f3f3")
        top_bar.grid(row=0, column=0, columnspan=2, sticky="ne", padx=5)
        
        self.btn_pin = tk.Button(
            top_bar, text="📍 Fijar", font=("Segoe UI", 8, "bold"),
            bg="#e0e0e0", fg="#333333", bd=0, padx=8, pady=2,
            cursor="hand2", command=self.conmutar_fijacion
        )
        self.btn_pin.pack(anchor="ne")

        hoy = datetime.today()
        siguiente_mes = hoy + relativedelta(months=1)

        # --- CALENDARIO 1: MES ACTUAL (IZQUIERDA) ---
        self.cal_actual = Calendar(
            self.main_frame, selectmode='day', 
            locale='es_ES', firstweekday='sunday',
            year=hoy.year, month=hoy.month, day=hoy.day,
            background='#2c3e50', foreground='white', 
            headersbackground='#34495e', headersforeground='white',
            normalbackground='white', normalforeground='black',
            weekendbackground='#f5f5f5', weekendforeground='black'
        )
        self.cal_actual.grid(row=1, column=0, padx=10, pady=10)

        # --- CALENDARIO 2: MES SIGUIENTE (DERECHA) ---
        self.cal_siguiente = Calendar(
            self.main_frame, selectmode='day', 
            locale='es_ES', firstweekday='sunday',
            year=siguiente_mes.year, month=siguiente_mes.month, day=1,
            background='#2c3e50', foreground='white', 
            headersbackground='#34495e', headersforeground='white',
            normalbackground='white', normalforeground='black',
            weekendbackground='#f5f5f5', weekendforeground='black'
        )
        self.cal_siguiente.grid(row=1, column=1, padx=10, pady=10)

        # 2. LEYENDA DE COLORES (Parte Inferior)
        self.crear_leyenda_inferior()

        # Inyección de sincronización de navegación bidireccional corregida
        self.inyectar_navegacion_sincronizada()

        # Dibujar marcas y eventos iniciales
        self.refrescar_toda_la_interfaz()

        # Crear el Tooltip flotante para el Hover
        self.tooltip = tk.Toplevel(self)
        self.tooltip.withdraw()
        self.tooltip.overrideredirect(True)
        self.tooltip_label = tk.Label(self.tooltip, bg="#2d2d2d", fg="white", 
                                      font=("Segoe UI", 9), padx=5, pady=3, 
                                      relief=tk.SOLID, borderwidth=1)
        self.tooltip_label.pack()

        # Vincular eventos de Mouse (Hover y Clic)
        self.bind_eventos_mouse(self.cal_actual)
        self.bind_eventos_mouse(self.cal_siguiente)
        
        # 3. Habilitar arrastre de la ventana usando el fondo del main_frame
        self.main_frame.bind("<Button-1>", self.iniciar_arrastre)
        self.main_frame.bind("<B1-Motion>", self.ejecutar_arrastre)

        # Ubicar la ventana cerca de la esquina superior izquierda al arrancar
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        self.geometry(f'{width}x{height}+50+50')

        # Cierre inteligente al perder el foco (Solo si no está fijado)
        self.bind("<FocusOut>", self.evaluar_cierre)
        self.after(100, self.focus_force)

    def conmutar_fijacion(self):
        """Alterna el estado de fijación del widget en primer plano"""
        self.esta_fijado = not self.esta_fijado
        if self.esta_fijado:
            self.btn_pin.config(text="📌 Fijado", bg="#3498db", fg="white")
            self.attributes("-topmost", True)
        else:
            self.btn_pin.config(text="📍 Fijar", bg="#e0e0e0", fg="#333333")
            self.focus_force()

    def cargar_configuracion_colores(self):
        """Lee la pestaña 'Configuracion' del Excel. Si no existe, la crea con valores por defecto."""
        excel_path = 'eventos.xlsx'
        colores_defecto = {
            'Fin sprint': '#3498db',
            'Inicio sprint': '#2ecc71',
            'Festivo': '#e74c3c',
            'Predeterminado': '#95a5a6'
        }
        
        from openpyxl import load_workbook, Workbook
        if not os.path.exists(excel_path):
            wb = Workbook()
            ws_ev = wb.active
            ws_ev.title = "Eventos"
            ws_ev.append(['fecha', 'evento', 'tipo'])
            ws_ev.append([datetime.today().strftime('%Y-%m-%d'), 'Evento de Prueba', 'Inicio sprint'])
            ws_cfg = wb.create_sheet(title="Configuracion")
            ws_cfg.append(['tipo_evento', 'codigo_hex'])
            for k, v in colores_defecto.items():
                ws_cfg.append([k, v])
            wb.save(excel_path)
            return colores_defecto

        try:
            wb = load_workbook(excel_path)
            if "Configuracion" not in wb.sheetnames:
                ws_cfg = wb.create_sheet(title="Configuracion")
                ws_cfg.append(['tipo_evento', 'codigo_hex'])
                for k, v in colores_defecto.items():
                    ws_cfg.append([k, v])
                wb.save(excel_path)
                return colores_defecto
            
            ws = wb["Configuracion"]
            colores_dict = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] and row[1]:
                    colores_dict[str(row[0]).strip()] = str(row[1]).strip()
            
            if 'Predeterminado' not in colores_dict:
                colores_dict['Predeterminado'] = '#95a5a6'
                
            wb.close()
            return colores_dict
        except Exception as e:
            print(f"Error al cargar configuración de colores: {e}")
            return colores_defecto

    def cargar_eventos(self):
        """Lee la pestaña 'Eventos' del Excel utilizando openpyxl (Sin rastro de Pandas)"""
        excel_path = 'eventos.xlsx'
        try:
            from openpyxl import load_workbook
            wb = load_workbook(excel_path, data_only=True)
            nombre_hoja = "Eventos" if "Eventos" in wb.sheetnames else wb.sheetnames[0]
            ws = wb[nombre_hoja]
            
            eventos_dict = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0]: 
                    continue
                    
                fecha_val = row[0]
                if isinstance(fecha_val, datetime):
                    fecha_str = fecha_val.strftime('%Y-%m-%d')
                else:
                    try:
                        fecha_str = str(fecha_val).split(" ")[0]
                    except:
                        fecha_str = str(fecha_val).split(" ")[0]
                
                evento = str(row[1]) if row[1] else ""
                tipo = str(row[2]).strip() if row[2] else "Predeterminado"
                
                if fecha_str not in eventos_dict:
                    eventos_dict[fecha_str] = []
                eventos_dict[fecha_str].append({'texto': evento, 'tipo': tipo})
                
            wb.close()
            return eventos_dict
        except Exception as e:
            print(f"Error al cargar el archivo Excel: {e}")
            return {}

    def crear_leyenda_inferior(self):
        """Crea de forma dinámica la sección de leyenda en la parte baja"""
        leyenda_frame = tk.Frame(self.main_frame, bg="#f3f3f3", pady=5)
        leyenda_frame.grid(row=2, column=0, columnspan=2, sticky="ew")
        
        leyenda_frame.columnconfigure(0, weight=1)
        container = tk.Frame(leyenda_frame, bg="#f3f3f3")
        container.pack(anchor="center")

        lbl_color_hoy = tk.Label(container, bg="#f1c40f", width=2, height=1, relief=tk.SOLID, bd=1)
        lbl_color_hoy.pack(side=tk.LEFT, padx=(5, 2))
        lbl_texto_hoy = tk.Label(container, text="Hoy", bg="#f3f3f3", font=("Segoe UI", 8, "bold"), fg="black")
        lbl_texto_hoy.pack(side=tk.LEFT, padx=(0, 15))

        for tipo, hex_color in self.colores_por_tipo.items():
            if tipo == 'Predeterminado': continue
            
            lbl_color = tk.Label(container, bg=hex_color, width=2, height=1, relief=tk.SOLID, bd=1)
            lbl_color.pack(side=tk.LEFT, padx=(5, 2))
            
            lbl_texto = tk.Label(container, text=tipo, bg="#f3f3f3", font=("Segoe UI", 8), fg="black")
            lbl_texto.pack(side=tk.LEFT, padx=(0, 15))

    def inyectar_navegacion_sincronizada(self):
        """Asegura sincronización bidireccional exacta sin importar qué calendario se mueva"""
        
        # Guardamos las funciones de navegación nativas de tkcalendar
        orig_next_izq = self.cal_actual._next_month
        orig_prev_izq = self.cal_actual._prev_month
        orig_next_der = self.cal_siguiente._next_month
        orig_prev_der = self.cal_siguiente._prev_month

        # --- CASO A: El usuario interactúa con el calendario IZQUIERDO ---
        def nueva_next_izq():
            if self.bloqueo_sincronizacion: return
            self.bloqueo_sincronizacion = True
            
            orig_next_izq()  # Avanza mes izquierdo
            # Sincroniza el derecho basándose en la nueva fecha del izquierdo
            fecha_ref = datetime(self.cal_actual.__year, self.cal_actual.__month, 1) + relativedelta(months=1)
            self.cal_siguiente.display_month(fecha_ref)
            
            self.refrescar_toda_la_interfaz()
            self.bloqueo_sincronizacion = False

        def nueva_prev_izq():
            if self.bloqueo_sincronizacion: return
            self.bloqueo_sincronizacion = True
            
            orig_prev_izq()  # Retrocede mes izquierdo
            fecha_ref = datetime(self.cal_actual.__year, self.cal_actual.__month, 1) + relativedelta(months=1)
            self.cal_siguiente.display_month(fecha_ref)
            
            self.refrescar_toda_la_interfaz()
            self.bloqueo_sincronizacion = False

        # --- CASO B: El usuario interactúa con el calendario DERECHO ---
        def nueva_next_der():
            if self.bloqueo_sincronizacion: return
            self.bloqueo_sincronizacion = True
            
            orig_next_der()  # Avanza mes derecho
            # Sincroniza el izquierdo restándole un mes exacto al derecho
            fecha_ref = datetime(self.cal_siguiente.__year, self.cal_siguiente.__month, 1) - relativedelta(months=1)
            self.cal_actual.display_month(fecha_ref)
            
            self.refrescar_toda_la_interfaz()
            self.bloqueo_sincronizacion = False

        def nueva_prev_der():
            if self.bloqueo_sincronizacion: return
            self.bloqueo_sincronizacion = True
            
            orig_prev_der()  # Retrocede mes derecho
            fecha_ref = datetime(self.cal_siguiente.__year, self.cal_siguiente.__month, 1) - relativedelta(months=1)
            self.cal_actual.display_month(fecha_ref)
            
            self.refrescar_toda_la_interfaz()
            self.bloqueo_sincronizacion = False

        # Asignación de los nuevos métodos acoplados simétricamente
        self.cal_actual._next_month = nueva_next_izq
        self.cal_actual._prev_month = nueva_prev_izq
        self.cal_siguiente._next_month = nueva_next_der
        self.cal_siguiente._prev_month = nueva_prev_der

    def refrescar_toda_la_interfaz(self):
        self.cal_actual.calevent_remove('all')
        self.cal_siguiente.calevent_remove('all')
        
        hoy = datetime.today()
        for cal in [self.cal_actual, self.cal_siguiente]:
            cal.calevent_create(hoy, 'Hoy', 'tag_hoy')
            cal.tag_config('tag_hoy', background='#f1c40f', foreground='black')

        hoy_str = hoy.strftime('%Y-%m-%d')
        for fecha_str, lista_eventos in self.eventos.items():
            if fecha_str == hoy_str:
                continue
            try:
                fecha_dt = datetime.strptime(fecha_str, '%Y-%m-%d')
                tipo_evento = lista_eventos[0]['tipo']
                color = self.colores_por_tipo.get(tipo_evento, self.colores_por_tipo['Predeterminado'])
                
                tag_name = f"tag_{fecha_str}"
                self.cal_actual.calevent_create(fecha_dt, lista_eventos[0]['texto'], tag_name)
                self.cal_actual.tag_config(tag_name, background=color, foreground='white')
                
                self.cal_siguiente.calevent_create(fecha_dt, lista_eventos[0]['texto'], tag_name)
                self.cal_siguiente.tag_config(tag_name, background=color, foreground='white')
            except ValueError:
                continue

    def bind_eventos_mouse(self, calendario):
        calendario.bind("<Motion>", lambda event: self.verificar_hover(event, calendario))
        calendario.bind("<Leave>", lambda event: self.ocultar_tooltip())
        calendario.bind("<<CalendarSelected>>", lambda event: self.mostrar_evento_seleccionado(calendario))

    def verificar_hover(self, event, calendario):
        x, y = event.x, event.y
        region = calendario.identify(x, y)
        
        if region == "day":
            date = calendario.get_date_at(x, y)
            if date:
                fecha_str = date.strftime('%Y-%m-%d')
                hoy_str = datetime.today().strftime('%Y-%m-%d')
                
                textos = []
                if fecha_str == hoy_str:
                    textos.append("★ Hoy")
                if fecha_str in self.eventos:
                    textos.extend([f"• {ev['texto']} ({ev['tipo']})" for ev in self.eventos[fecha_str]])
                
                if textos:
                    tooltip_text = "\n".join(textos)
                    gx = self.winfo_pointerx() + 15
                    gy = self.winfo_pointery() + 15
                    self.tooltip_label.config(text=tooltip_text)
                    self.tooltip.geometry(f"+{gx}+{gy}")
                    self.tooltip.deiconify()
                    return
        self.ocultar_tooltip()

    def ocultar_tooltip(self):
        self.tooltip.withdraw()

    def mostrar_evento_seleccionado(self, calendario):
        fecha_dt = calendario.selection_get()
        if not fecha_dt: return
            
        fecha_str = fecha_dt.strftime('%Y-%m-%d')
        if fecha_str not in self.eventos: return

        textos = [f"• {ev['texto']} ({ev['tipo']})" for ev in self.eventos[fecha_str]]
        detalle_eventos = "\n".join(textos)

        popup = tk.Toplevel(self)
        popup.title("Detalle de Eventos")
        popup.configure(bg="#2c3e50")
        popup.geometry("320x200")
        popup.resizable(False, False)
        
        px = self.winfo_x() + (self.winfo_width() // 2) - 160
        py = self.winfo_y() + (self.winfo_height() // 2) - 100
        popup.geometry(f"+{px}+{py}")
        popup.attributes("-topmost", True)

        lbl_fecha = tk.Label(popup, text=f"Eventos: {fecha_dt.strftime('%d/%m/%Y')}", 
                             font=("Segoe UI", 11, "bold"), bg="#2c3e50", fg="#f1c40f", pady=10)
        lbl_fecha.pack()

        txt_box = tk.Text(popup, bg="#34495e", fg="white", font=("Segoe UI", 10), wrap=tk.WORD, bd=0, padx=10, pady=5)
        txt_box.insert(tk.END, detalle_eventos)
        txt_box.config(state=tk.DISABLED)
        txt_box.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        btn_cerrar = tk.Button(popup, text="Cerrar", bg="#e74c3c", fg="white", font=("Segoe UI", 9, "bold"),
                               activebackground="#c0392b", activeforeground="white", bd=0, 
                               command=popup.destroy, cursor="hand2", padding=5)
        btn_cerrar.pack(pady=10)

        popup.focus_set()
        popup.grab_set()

    def iniciar_arrastre(self, event):
        self.x_offset = event.x
        self.y_offset = event.y

    def ejecutar_arrastre(self, event):
        x = self.winfo_x() + (event.x - self.x_offset)
        y = self.winfo_y() + (event.y - self.y_offset)
        self.geometry(f"+{x}+{y}")

    def evaluar_cierre(self, event):
        if self.esta_fijado:
            return
        self.after(120, self._verificar_foco_real)

    def _verificar_foco_real(self):
        focus_widget = self.focus_get()
        if focus_widget is None:
            ventanas_secundarias = [w for w in self.winfo_children() if isinstance(w, tk.Toplevel) and w.winfo_viewable()]
            ventanas_activas = [v for v in ventanas_secundarias if v != self.tooltip]
            if not ventanas_activas:
                self.destroy()

if __name__ == "__main__":
    app = CalendarioEventosApp()
    app.mainloop()