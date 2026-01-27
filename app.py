import os
import shutil
config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.py')
default_config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.py.default')
if not os.path.exists(config_path) and os.path.exists(default_config_path):
    shutil.copy(default_config_path, config_path)
import json
import tkinter as tk
from tkinter import filedialog, colorchooser, messagebox, ttk
from tkinter import font as tkfont
from PIL import Image, ImageDraw, ImageFont, ImageTk
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import pandas as pd
import threading
import queue
from concurrent.futures import ThreadPoolExecutor
import base64
import io
import re
import logging
import argparse
import sys
from datetime import datetime
import subprocess
import time
import config
import socket

# Configuração do sistema de logging
def setup_logging(verbose=False):
    """Configura o sistema de logging da aplicação"""
    log_level = logging.DEBUG if verbose else logging.WARNING
    log_format = '%(levelname)s: %(message)s'
    
    # Configuração básica do logging
    logging.basicConfig(level=log_level, format=log_format)
    
    # Personalizando handlers
    console = logging.StreamHandler(sys.stdout)
    console.setLevel(log_level)
    
    # Formatador personalizado para categorias [Info], [Erro], etc.
    class CustomFormatter(logging.Formatter):
        FORMATS = {
            logging.DEBUG: '[Debug] %(message)s',
            logging.INFO: '[Info] %(message)s',
            logging.WARNING: '[Warning] %(message)s',
            logging.ERROR: '[Error] %(message)s',
            logging.CRITICAL: '[Critical] %(message)s'
        }

        def format(self, record):
            log_fmt = self.FORMATS.get(record.levelno)
            formatter = logging.Formatter(log_fmt)
            return formatter.format(record)
    
    console.setFormatter(CustomFormatter())
    
    # Remove handlers anteriores
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    
    # Adiciona o handler personalizado
    logging.root.addHandler(console)
    
    # Adiciona um handler de arquivo apenas se estiver em modo verbose
    if verbose:
        log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        os.makedirs(log_dir, exist_ok=True)
        
        log_file = os.path.join(log_dir, f"visumaker_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        file_formatter = logging.Formatter('%(asctime)s - %(levelname)s: %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        file_handler.setFormatter(file_formatter)
        logging.root.addHandler(file_handler)
        
        logging.info(f"Logs will be saved at: {log_file}")

# Parse de argumentos de linha de comando
parser = argparse.ArgumentParser(description='VisuMaker - Gerador de Documentos Visuais')
parser.add_argument('-v', '--verbose', action='store_true', help='Ativa modo verbose com logs detalhados')
parser.add_argument('-csv', '--csv', type=str, help='Caminho para um arquivo CSV para carregar automaticamente', default=None)
parser.add_argument('-proj', '--proj', type=str, help='Caminho para um template HTML para carregar automaticamente', default=None)
args = parser.parse_args()

# Configura o sistema de logging baseado nos argumentos
setup_logging(verbose=args.verbose)

# Otimizações para a biblioteca Pillow/PIL
Image.MAX_IMAGE_PIXELS = None  # Desativa limite de tamanho de imagem
os.environ['PILLOW_CHUNK_SIZE'] = '1024'  # Aumenta tamanho do chunk para operações

# Import para mapear fontes
import ctypes
from ctypes import wintypes

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("VisuMaker")
        self.geometry("1200x700")
        self.configure(bg='#303030')  # Tema escuro
        
        # Define o ícone da aplicação
        try:
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.ico")
            if os.path.exists(icon_path):
                self.iconbitmap(default=icon_path)
        except Exception as e:
            logging.warning(f"Não foi possível carregar ícone: {e}")
        
        # Configura o protocolo para capturar o evento de fecho
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Força o diretório de trabalho para o diretório do script
        os.chdir(os.path.dirname(os.path.abspath(__file__)))
        
        # Inicialização das configurações de email
        self.email_config = {
            "authenticated": False,  # Por predefinição, não está autenticado
            "use_template": False,
            "template_path": "",
            "subject": "Documento Visual",
            "attachment_name": "Documento_{nome}.png",
            "cc": [],
            "bcc": []
        }
        
        # Carrega configurações do ficheiro config.py
        if hasattr(config, 'DEFAULT_EMAIL_CONFIG'):
            self.email_config.update(config.DEFAULT_EMAIL_CONFIG)
        
        # Widget oculto para armazenar o texto do email (usado ao configurar email)
        self.email_text = tk.Text(self)
        self.email_text.insert("1.0", config.EMAIL_BODY if hasattr(config, 'EMAIL_BODY') else "")
        self.email_text.pack_forget()  # Não mostra na interface
        
        # Configuração de estilos globais para ttk
        self._setup_styles()
        
        # Carrega ou cria mapeamento de fontes
        self.fonts_map_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts_map.json")
        self.fonts_map = self.load_fonts_map()
        
        # Força a atualização do mapeamento de fontes
        #self.update_fonts_map()
        
        # Iniciando variáveis de controlo de camadas e itens
        self.items = {}          # Armazena todos os itens por ID
        self.item_order = []     # Ordem de desenho das camadas (do fundo para a frente)
        self.layer_names = {}    # Nomes amigáveis das camadas
        self.layer_widgets = {}  # Widgets utilizados para representar as camadas na UI
        self.visible_items = {}  # Controlo de visibilidade por ID
        self.current_item = None # Item selecionado atualmente
        self.selected_items = set()  # Conjunto de itens selecionados (para seleção múltipla)
        self.df = None           # DataFrame para armazenar dados do CSV
        self.forget_csv_btn = None # Referência ao botão de esquecer CSV
        
        # Estado inicial para o canvas
        self.model_image = None     # A imagem de fundo
        self.pil_model = None       # A imagem PIL original
        self.model_scale = 1.0      # Escala inicial da imagem modelo
        self.zoom_factor = 1.0      # Zoom aplicado à visualização
        
        # Rastreamento de interação com o rato
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.drag_item = None
        self.drag_mode = None
        self.dragging_handle = None
        self.drag_data = {}     # Dados extras para operações específicas
        
        # Estado de pan/scroll
        self.pan_start_x = 0
        self.pan_start_y = 0
        self.is_panning = False

        # Canvas de edição (área principal)
        self.main_frame = tk.Frame(self, bg='#303030')
        self.main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Frame para controlos de zoom
        zoom_frame = tk.Frame(self.main_frame, bg='#3a3a3a', height=30)
        zoom_frame.pack(side=tk.TOP, fill=tk.X)
        
        # Novo slider para zoom em vez dos botões
        tk.Label(zoom_frame, text="Zoom:", bg='#3a3a3a', fg='white').pack(side=tk.LEFT, padx=5, pady=2)
        
        self.zoom_var = tk.DoubleVar(value=1.0)
        zoom_slider = ttk.Scale(zoom_frame, from_=0.1, to=5.0, 
                               orient=tk.HORIZONTAL, variable=self.zoom_var, 
                               command=self.on_zoom_change)
        zoom_slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=2)
        
        # Botão para repor zoom para 100%
        tk.Button(zoom_frame, text="100%", bg='#444444', fg='white',
                 command=self.zoom_reset).pack(side=tk.LEFT, padx=5, pady=2)
        
        tk.Button(zoom_frame, text="◯", bg='#444444', fg='white',
                 command=self.center_view).pack(side=tk.LEFT, padx=5, pady=2)
        
        # Label mostrando o zoom atual
        self.zoom_label = tk.Label(zoom_frame, text="100%", bg='#3a3a3a', fg='white')
        self.zoom_label.pack(side=tk.LEFT, padx=5, pady=2)
        
        # Canvas com borda como Photoshop
        self.canvas_frame = tk.Frame(self.main_frame, bd=1, relief=tk.SUNKEN, bg='#1e1e1e')
        self.canvas_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        # Contentor do canvas com scrollbars
        self.canvas_container = tk.Frame(self.canvas_frame, bg='#1e1e1e')
        self.canvas_container.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbars
        self.h_scrollbar = tk.Scrollbar(self.canvas_container, orient=tk.HORIZONTAL)
        self.h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.v_scrollbar = tk.Scrollbar(self.canvas_container)
        self.v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Canvas principal
        self.canvas = tk.Canvas(self.canvas_container, bg='#202020', 
                              width=800, height=600,
                              xscrollcommand=self.h_scrollbar.set,
                              yscrollcommand=self.v_scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Configura scrollbars
        self.h_scrollbar.config(command=self.canvas.xview)
        self.v_scrollbar.config(command=self.canvas.yview)
        
        # Bind eventos do rato
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.canvas.bind("<Shift-ButtonPress-1>", self.on_shift_press)
        
        # Para pan/movimentação do canvas (apenas com botão direito)
        self.canvas.bind("<ButtonPress-3>", self.start_pan)  # Botão direito
        self.canvas.bind("<B3-Motion>", self.pan_canvas)    # Arrastar com botão direito
        self.canvas.bind("<ButtonRelease-3>", self.end_pan) # Soltar botão direito
        
        # Zoom com roda do rato
        self.canvas.bind("<MouseWheel>", self.mouse_zoom)  # Windows
        self.canvas.bind("<Button-4>", self.mouse_zoom)    # Linux scroll up
        self.canvas.bind("<Button-5>", self.mouse_zoom)    # Linux scroll down

        # Barra de ferramentas superior (estilo Photoshop)
        tool_frame = tk.Frame(self, bg='#3a3a3a', height=40)
        tool_frame.pack(side=tk.TOP, fill=tk.X)
        
        # Botões de ferramentas com ícones
        tools = [
            ("Novo", self.new_project),
            ("Modelo", self.load_model),
            ("Texto", self.add_text),
            ("Imagem", self.add_image),
        ]
        
        for i, (txt, cmd) in enumerate(tools):
            btn = tk.Button(tool_frame, text=txt, command=cmd, 
                           bg='#444444', fg='white', relief=tk.FLAT,
                           activebackground='#555555', activeforeground='white')
            btn.pack(side=tk.LEFT, padx=5, pady=5)

        # Painel lateral direito (painéis de camadas e propriedades)
        self.side_panel = tk.Frame(self, bg='#252525', width=320)
        self.side_panel.pack(side=tk.RIGHT, fill=tk.Y)
        self.side_panel.pack_propagate(False)  # Impede que o painel encolha
        
        # Adiciona scrollbar à barra lateral
        self.sidebar_canvas = tk.Canvas(self.side_panel, bg='#252525', highlightthickness=0)
        self.sidebar_scrollbar = tk.Scrollbar(self.side_panel, orient=tk.VERTICAL, command=self.sidebar_canvas.yview)
        self.sidebar_canvas.configure(yscrollcommand=self.sidebar_scrollbar.set)
        
        # Empacota a scrollbar e o canvas
        self.sidebar_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.sidebar_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Frame interior para conter todos os widgets da barra lateral
        self.sidebar_interior = tk.Frame(self.sidebar_canvas, bg='#252525', padx=0, pady=0)
        # Criação do item com posição explícita 0,0
        self.sidebar_interior_id = self.sidebar_canvas.create_window(0, 0, window=self.sidebar_interior, anchor=tk.NW, width=300)
        
        # Configura eventos para rolagem
        self.sidebar_interior.bind('<Configure>', self._configure_sidebar_interior)
        self.sidebar_canvas.bind('<Configure>', self._configure_sidebar_canvas)
        
        # Permite rolagem com a roda do rato
        self.sidebar_canvas.bind_all('<MouseWheel>', self._on_sidebar_mousewheel)
        self.sidebar_canvas.bind_all('<Button-4>', self._on_sidebar_mousewheel)
        self.sidebar_canvas.bind_all('<Button-5>', self._on_sidebar_mousewheel)
        
        # Painel de camadas (estilo Photoshop)
        layers_frame = tk.LabelFrame(self.sidebar_interior, text="Camadas", 
                                    bg='#252525', fg='white', 
                                    font=('Arial', 10, 'bold'))
        layers_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Contentor para a lista de camadas
        self.layers_list = tk.Frame(layers_frame, bg='#333333')
        self.layers_list.pack(fill=tk.X, padx=5, pady=5)

        # Botões de camada
        layer_btns = tk.Frame(layers_frame, bg='#252525')
        layer_btns.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Button(layer_btns, text="▲", bg='#444444', fg='white',
                 command=self.move_layer_up).pack(side=tk.LEFT, padx=2)
        tk.Button(layer_btns, text="▼", bg='#444444', fg='white',
                 command=self.move_layer_down).pack(side=tk.LEFT, padx=2)
        tk.Button(layer_btns, text="❌", bg='#444444', fg='white',
                 command=self.delete_layer).pack(side=tk.LEFT, padx=2)
        
        # Ferramentas de alinhamento
        align_frame = tk.LabelFrame(self.sidebar_interior, text="Alinhamento", 
                                  bg='#252525', fg='white', 
                                  font=('Arial', 10, 'bold'))
        align_frame.pack(fill=tk.X, padx=5, pady=5)
        
        align_row1 = tk.Frame(align_frame, bg='#252525')
        align_row1.pack(fill=tk.X, padx=5, pady=2)
        
        tk.Button(align_row1, text="◀ Esq", bg='#444444', fg='white',
                command=lambda: self.align_selected("left")).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        tk.Button(align_row1, text="Centro ↔", bg='#444444', fg='white',
                command=lambda: self.align_selected("centerx")).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        tk.Button(align_row1, text="Dir ▶", bg='#444444', fg='white',
                command=lambda: self.align_selected("right")).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        
        align_row2 = tk.Frame(align_frame, bg='#252525')
        align_row2.pack(fill=tk.X, padx=5, pady=2)
        
        tk.Button(align_row2, text="▲ Topo", bg='#444444', fg='white',
                command=lambda: self.align_selected("top")).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        tk.Button(align_row2, text="Centro ↕", bg='#444444', fg='white',
                command=lambda: self.align_selected("centery")).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        tk.Button(align_row2, text="Base ▼", bg='#444444', fg='white',
                command=lambda: self.align_selected("bottom")).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        
        # Botões para centralizar
        center_row = tk.Frame(align_frame, bg='#252525')
        center_row.pack(fill=tk.X, padx=5, pady=2)
        
        tk.Button(center_row, text="Centralizar Tudo", bg='#444444', fg='white',
                command=lambda: self.align_selected("center")).pack(fill=tk.X, padx=2)
        
        # Painel de propriedades (estilo Photoshop)
        self.props_frame = tk.LabelFrame(self.sidebar_interior, text="Propriedades", 
                                       bg='#252525', fg='white', 
                                       font=('Arial', 10, 'bold'))
        self.props_frame.pack(fill=tk.X, padx=5, pady=5)

        # Painel de ações
        actions_frame = tk.LabelFrame(self.sidebar_interior, text="Ações", 
                                     bg='#252525', fg='white', 
                                     font=('Arial', 10, 'bold'))
        actions_frame.pack(fill=tk.X, padx=5, pady=5)
        
        act_btns = [
            ("Carregar CSV", self.load_csv),
            ("Esquecer CSV", self.forget_csv),
            ("Gerar & Enviar", self.generate_all),
            ("Apenas Gerar", self.generate_images_only),
            ("Testar Documento", self.generate_test_certificate),
            ("Guardar Layout", self.save_layout),
            ("Carregar Layout", self.load_layout),
            ("Atualizar Fontes", self.update_fonts_map),
            ("Editar Config", self.open_config_editor),
        ]
        self.forget_csv_btn = None  # Referência ao botão para atualizar estado
        for (txt, cmd) in act_btns:
            btn = tk.Button(actions_frame, text=txt, command=cmd, 
                     bg='#444444', fg='white')
            btn.pack(fill=tk.X, pady=3, padx=5)
            if txt == "Esquecer CSV":
                self.forget_csv_btn = btn
                btn.config(state=tk.NORMAL if self.df is not None else tk.DISABLED)

        # Estado inicial para o canvas
        self.canvas_state = {
            'move_mode': False,
            'drag_data': {'item': None, 'x': 0, 'y': 0},
            'select_box': None,
            'multi_select_active': False
        }
        
        # Variáveis para controle de redraw e zoom
        self.redraw_job = None
        self.zoom_level = 100
        
        # Inicializa a imagem do modelo
        self.model_img = None

    def forget_csv(self):
        """Esquece o CSV carregado e desativa o botão."""
        self.df = None
        if self.forget_csv_btn:
            self.forget_csv_btn.config(state=tk.DISABLED)
        messagebox.showinfo("CSV Esquecido", "O CSV carregado foi esquecido.")

    def open_config_editor(self):
        """Abre um editor de configurações (glorified text editor)."""
        ConfigEditor(self)

    def add_layer_to_list(self, item_id, name):
        # Adiciona à lista de camadas e à ordem
        self.layer_names[item_id] = name
        if item_id not in self.item_order:
            self.item_order.append(item_id)
        
        # Atualiza a listbox
        self.refresh_layers_list()
        
        # Seleciona a nova camada
        self.select_layer(item_id)
        
    def select_layer(self, item_id):
        # Verifica se o item existe antes de prosseguir
        if item_id not in self.items:
            #print(f"Item não encontrado: {item_id}")
            return
            
        # Atualiza a seleção atual
        old_selection = self.current_item
        self.current_item = item_id
        self.selected_items = {item_id}  # Apenas este item na seleção
        
        # Atualiza visual na lista
        for layer_id in self.layer_widgets:
            # Reset todas as camadas
            self.layer_widgets[layer_id]['frame'].config(bg='#333333')
            self.layer_widgets[layer_id]['label'].config(bg='#333333')
            
        # Destaca a selecionada
        if item_id in self.layer_widgets:
            self.layer_widgets[item_id]['frame'].config(bg='#4a90d9')
            self.layer_widgets[item_id]['label'].config(bg='#4a90d9')
        
        # Mostra propriedades do item
        try:
            self.show_properties(item_id)
        except Exception as e:
            logging.error(f"Error showing properties: {str(e)}")
            messagebox.showinfo("Erro", f"Erro ao mostrar propriedades: {str(e)}")
        
        # Atualiza seleção no canvas
        self.schedule_redraw()
        
    def move_layer_up(self):
        if not self.current_item or self.current_item not in self.item_order:
            return
            
        index = self.item_order.index(self.current_item)
        if index == 0:
            return  # Já está no topo
            
        # Move na lista
        self.item_order.remove(self.current_item)
        self.item_order.insert(index - 1, self.current_item)
        
        # Atualiza lista e canvas
        self.refresh_layers_list()
        self.schedule_redraw()
        
    def move_layer_down(self):
        if not self.current_item or self.current_item not in self.item_order:
            return
            
        index = self.item_order.index(self.current_item)
        if index >= len(self.item_order) - 1:
            return  # Já está na base
            
        # Move na lista
        self.item_order.remove(self.current_item)
        self.item_order.insert(index + 1, self.current_item)
        
        # Atualiza lista e canvas
        self.refresh_layers_list()
        self.schedule_redraw()
        
    def delete_layer(self):
        if not self.current_item:
            return
            
        item_id = self.current_item
        
        # Remove das estruturas de dados
        if item_id in self.item_order:
            self.item_order.remove(item_id)
        if item_id in self.layer_names:
            self.layer_names.pop(item_id)
        if item_id in self.visible_items:
            self.visible_items.pop(item_id)
        if item_id in self.items:
            self.items.pop(item_id)
        if item_id in self.selected_items:
            self.selected_items.remove(item_id)
            
        # Limpa a seleção atual
        self.current_item = None
            
        # Atualiza interface
        self.refresh_layers_list()
        self.schedule_redraw()
        
        # Limpa o painel de propriedades
        for w in self.props_frame.winfo_children():
            w.destroy()

    def load_model(self):
        path = filedialog.askopenfilename(filetypes=[("Imagem", ".png .jpg")])
        if not path: return
        
        # Cria uma janela de progresso
        progress = tk.Toplevel(self)
        progress.title("A Carregar Modelo")
        progress.geometry("300x80")
        progress.transient(self)
        
        tk.Label(progress, text="A carregar imagem...").pack(pady=10)
        progress_bar = ttk.Progressbar(progress, mode="indeterminate")
        progress_bar.pack(fill=tk.X, padx=20, pady=10)
        progress_bar.start()
        progress.update()
        progress.update_idletasks()  # Força a atualização da interface
        
        # Captura self e progress para usar na função aninhada
        _self = self
        _progress = progress
        _path = path
        
        # Função que executa em thread separada
        def load_image_thread():
            try:
                img = Image.open(_path).convert("RGBA")
                # Coloca a imagem carregada na fila para a thread principal
                _self.after(0, lambda: _self._finish_loading_model(img, _progress))
            except Exception as e:
                # Em caso de erro, notifica a thread principal
                _self.after(0, lambda: _self._handle_load_error(str(e), _progress))
                
        # Inicia o carregamento em thread separada
        threading.Thread(target=load_image_thread, daemon=True).start()
    
    def _finish_loading_model(self, img, progress_window):
        """Finaliza o carregamento do modelo na thread principal"""
        # Limpa o cache do modelo anterior
        if hasattr(self, 'model_cache'):
            self.model_cache = None
            
        # Define o novo modelo
        self.model_img = img
        
        # Atualiza tamanho do canvas para corresponder à imagem
        w, h = self.model_img.size
        self.canvas.config(width=w, height=h)

        # Redefine o zoom para 100% ao carregar um novo modelo
        self.zoom_factor = 1.0
        self.zoom_label.config(text="100%")
        
        # Limpa a região de scroll para centralizar o novo modelo
        self.canvas.config(scrollregion=(0, 0, w, h))
        
        # Centraliza a visualização
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)
        
        # Se já existiam camadas, pergunta se deseja limpar
        if self.items:
            if messagebox.askyesno("Limpar Camadas", "Deseja limpar todas as camadas existentes?\n\nO novo modelo pode ter um tamanho diferente do anterior."):
                self.clear_all_layers()
        
        # Redraw o canvas
        self.schedule_redraw(immediate=True)
        
        # Fecha janela de progresso
        progress_window.destroy()
        
        # Feedback para o utilizador
        messagebox.showinfo("Modelo Carregado", f"Novo modelo carregado com sucesso.\nDimensões: {w}x{h} pixels")
    
    def clear_all_layers(self):
        """Limpa todas as camadas e reinicia o estado"""
        # Reset das estruturas de dados
        self.items = {}
        self.visible_items = {}
        self.layer_names = {}
        self.item_order = []
        self.current_item = None
        self.selected_items = set()
        self.selection_rectangles = {}
        
        # Limpa todos os widgets de camadas
        for widget in self.layers_list.winfo_children():
            widget.destroy()
        
        self.layer_widgets = {}
        
        # Limpa o painel de propriedades
        for w in self.props_frame.winfo_children():
            w.destroy()
    
    def _handle_load_error(self, error_msg, progress_window):
        """Trata erro de carregamento na thread principal"""
        progress_window.destroy()
        messagebox.showerror("Erro ao carregar imagem", error_msg)

    def schedule_redraw(self, immediate=False):
        """Agenda um redraw do canvas com debounce para melhorar performance"""
        if self.redraw_job is not None:
            self.after_cancel(self.redraw_job)
            self.redraw_job = None
            
        if immediate:
            self._redraw_canvas_now()
        else:
            # Agenda o redraw para 30ms depois para evitar múltiplas chamadas
            # e manter a interface responsiva
            self.redraw_job = self.after(30, self._redraw_canvas_now)
    
    def _redraw_canvas_now(self):
        """Executa o redraw imediatamente"""
        self.redraw_job = None
        self.redraw_canvas()

    def redraw_canvas(self):
        """Redesenha o canvas inteiro - esta função é potencialmente pesada"""
        # Reset do canvas apenas se necessário
        if not self.model_img: return
        
        # Preserva os dados de arrasto para restaurar após redesenho
        preserve_drag_data = self.drag_data.copy() if hasattr(self, 'drag_data') else None
        
        # Limpa o canvas completamente
        self.canvas.delete("all")
        
        # Configuração do canvas para o zoom
        scaled_width = int(self.model_img.width * self.zoom_factor)
        scaled_height = int(self.model_img.height * self.zoom_factor)
        
        # Configura a região de scroll
        self.canvas.config(width=min(scaled_width, 800), height=min(scaled_height, 600),
                          scrollregion=(0, 0, scaled_width, scaled_height))
        
        # Desenha o fundo (modelo) com zoom
        self._render_model_image()
        
        # Limpa as referências de retângulos de seleção
        self.selection_rectangles = {}

        # Desenha apenas as camadas visíveis para melhorar performance
        for item_id in self.item_order:
            # Pula se não estiver visível
            if item_id not in self.visible_items:
                continue
                
            info = self.items[item_id]
            
            if info['tipo'] == 'texto':
                self._draw_text_item(item_id, info)
            else:  # Imagem
                self._draw_image_item(item_id, info)
        
        # Garante que os controles de seleção estão acima das camadas
        for item_id in self.selection_rectangles:
            self.canvas.tag_raise(f"sel_{item_id}")
            
        # Restaura os dados de arrasto
        if preserve_drag_data:
            self.drag_data = preserve_drag_data
    
    def _draw_text_item(self, item_id, info):
        """Desenha um item de texto no canvas - função auxiliar para melhorar performance"""
        # Calcula a posição com zoom (sempre usando o canto superior esquerdo como referência)
        x_pos = int(info['xy'][0] * self.zoom_factor)
        y_pos = int(info['xy'][1] * self.zoom_factor)
        
        # Ajusta o tamanho da fonte para o zoom
        font_size = int(info['size'] * self.zoom_factor)
        
        # Obtém a largura definida pelo usuário
        width = int(info.get('width', 200) * self.zoom_factor)
        
        # Obtém propriedades de formatação
        # Força bg_color para vazio para garantir fundo transparente em todas as caixas
        bg_color = ""  # Sempre transparente
        text_align = info.get('text_align', 'left')
        
        # Definir a cor do contorno - sempre visível mesmo para caixas transparentes
        outline_color = "#dddddd"
        if item_id in self.selected_items:
            outline_color = "#4a90d9"  # Contorno mais visível quando selecionado
        
        # Calcular a altura do texto com base no conteúdo
        # Cria um objeto de fonte para medir o texto
        try:
            font_obj = tkfont.Font(family=info['font_family'], size=font_size)
            if font_obj:
                # Calcula o número de linhas após quebrar o texto para a largura
                text_lines = []
                current_line = ""
                
                # Divide o texto por espaços
                words = info['texto'].split()
                if not words:
                    text_lines = [""]  # Texto vazio
                else:
                    for word in words:
                        # Testa adicionar a próxima palavra
                        test_line = current_line + " " + word if current_line else word
                        # Verifica se cabe na largura
                        if font_obj.measure(test_line) <= width - 10:  # 10 de padding
                            current_line = test_line
                        else:
                            # Não cabe, adiciona a linha atual e começa nova
                            if current_line:
                                text_lines.append(current_line)
                            current_line = word
                    
                    # Adiciona a última linha
                    if current_line:
                        text_lines.append(current_line)
                
                # Quebras de linha explícitas no texto
                explicit_lines = info['texto'].split('\n')
                if len(explicit_lines) > 1:
                    # Recalcula considerando as quebras de linha explícitas
                    text_lines = []
                    for line in explicit_lines:
                        if not line.strip():
                            text_lines.append("")
                            continue
                            
                        words = line.split()
                        current_line = ""
                        for word in words:
                            test_line = current_line + " " + word if current_line else word
                            if font_obj.measure(test_line) <= width - 10:
                                current_line = test_line
                            else:
                                if current_line:
                                    text_lines.append(current_line)
                                current_line = word
                        
                        if current_line:
                            text_lines.append(current_line)
                
                # Calcula a altura com base no número de linhas e tamanho da fonte
                # Adiciona um pequeno espaço extra para melhor visualização
                line_count = max(1, len(text_lines))
                line_height = font_size * 1.2  # Altura aproximada da linha
                height = int(line_count * line_height) + 10  # Adiciona padding
            else:
                # Fallback se não conseguir criar a fonte
                height = int(font_size * 3)  # Altura padrão
        except Exception as e:
            logging.error(f"Error calculating text height: {e}")
            height = int(font_size * 3)  # Altura padrão
        
        # Atualiza a altura no dicionário de informações
        info['height'] = int(height / self.zoom_factor)  # Volta para escala real
        
        # Cria um retângulo para representar a caixa de texto (sempre com fundo transparente)
        rect_id = self.canvas.create_rectangle(
            x_pos, y_pos, x_pos + width, y_pos + height,
            fill="", outline=outline_color, width=1,  # fill="" garante que o fundo é sempre transparente
            tags=(item_id, "text_box", f"item_{item_id}"))
        
        # Cria string de estilo de fonte (negrito, itálico, sublinhado)
        font_style = ""
        if info.get('bold', False):
            font_style += " bold"
        if info.get('italic', False):
            font_style += " italic"
        
        # Ajusta o alinhamento de texto - apenas para renderização, não afeta a bounding box
        anchor = tk.NW
        text_x = x_pos + 5  # Padding inicial
        
        if text_align == 'center':
            # Para texto centralizado, usamos o centro da caixa
            anchor = tk.N
            text_x = x_pos + width // 2
        elif text_align == 'right':
            # Para texto à direita, usamos o canto superior direito
            anchor = tk.NE
            text_x = x_pos + width - 5  # Padding
        
        # Mapeia os valores de alinhamento para os valores aceitáveis pelo Tkinter
        justify_map = {
            'left': 'left',
            'center': 'center',
            'right': 'right',
            'justify': 'left'  # Fallback para justify
        }
        
        # Garante que o valor seja válido
        justify_value = justify_map.get(text_align, 'left')
        
        # Cria o texto no canvas
        text_id = self.canvas.create_text(
            text_x, y_pos + 5,  # Pequeno padding interno
            text=info['texto'], 
            font=(info['font_family'], font_size, font_style),
            fill=info['color'],
            anchor=anchor,  # Alinhamento baseado na configuração
            width=width - 10,  # Largura do texto com padding
            justify=justify_value,  # Justificação interna do texto
            tags=(item_id, "text", f"item_{item_id}"))
        
        # Se estiver selecionado, desenha bounding box
        if item_id in self.selected_items:
            # A bounding box sempre usa as coordenadas originais da caixa de texto
            bbox = (x_pos, y_pos, x_pos + width, y_pos + height)
            self._draw_selection_handles(item_id, bbox)
    
    def _draw_image_item(self, item_id, info):
        """Desenha um item de imagem no canvas - função auxiliar para melhorar performance"""
        try:
            # Debug para verificar o ID e posição
            logging.debug(f"Drawing image '{item_id}', position: {info['xy']}")
            
            # Carrega a imagem apenas se necessário
            # Verifica se já temos uma versão em cache da imagem
            if not hasattr(info, 'img_cache') or info.get('img_cache_zoom', 0) != self.zoom_factor:
                # Carrega a imagem
                im = Image.open(info['path']).convert("RGBA")
                
                # Redimensiona com o zoom
                scaled_size = (int(info['size'][0] * self.zoom_factor), 
                            int(info['size'][1] * self.zoom_factor))
                im = im.resize(scaled_size, Image.LANCZOS)
                
                # Aplica opacidade corretamente
                if info['opacity'] < 1.0:
                    # Criar uma cópia da imagem com o canal alpha modificado
                    alpha = int(255 * info['opacity'])
                    
                    # Criar um novo array de pixels para alpha
                    data = im.getdata()
                    new_data = []
                    for item in data:
                        # item é (r, g, b, a)
                        new_a = int(item[3] * (alpha / 255))
                        new_data.append((item[0], item[1], item[2], new_a))
                    
                    # Atualiza a imagem com os novos valores alpha
                    im.putdata(new_data)
            
            # Salva no cache para uso futuro
            info['tkimg'] = ImageTk.PhotoImage(im)
            info['img_cache_zoom'] = self.zoom_factor
            
            # Posição com zoom
            x = int(info['xy'][0] * self.zoom_factor)
            y = int(info['xy'][1] * self.zoom_factor)
            
            # Criamos a imagem com tags adequadas para poder identificá-la
            img_id = self.canvas.create_image(
                x, y,
                image=info['tkimg'], anchor='nw',
                tags=(item_id, "image", f"item_{item_id}"))
            
            # Se estiver selecionado, desenha bounding box
            if item_id in self.selected_items:
                scaled_size = (int(info['size'][0] * self.zoom_factor), 
                            int(info['size'][1] * self.zoom_factor))
                bbox = (x, y, x + scaled_size[0], y + scaled_size[1])
                self._draw_selection_handles(item_id, bbox)
                
        except Exception as e:
            # Registra erro ao processar imagem no canvas
            logging.error(f"Error processing image {info.get('path')}: {str(e)}")

    def _draw_selection_handles(self, item_id, bbox):
        """Desenha apenas a bounding box de seleção para um item"""
        # Verifica se o ID é válido e existe no dicionário de itens
        if not item_id or item_id not in self.items:
            logging.debug(f"Warning: Attempt to draw selection for invalid item_id: {item_id}")
            return
            
        # Evita IDs genéricos que podem causar conflitos
        if item_id == "text" or item_id == "handle" or item_id == "selection":
            logging.debug(f"Warning: Attempt to draw selection for generic tag: {item_id}")
            return
        
        # Debug para verificar
        logging.debug(f"Drawing selection for item: {item_id}, bbox: {bbox}")
        
        # Desenha apenas o retângulo de seleção, sem handles
        sel_rect_id = self.canvas.create_rectangle(
                            bbox[0], bbox[1], bbox[2], bbox[3],
                            outline='#4a90d9', width=2, 
                            dash=(5,5),  # Linha tracejada
                            tags=(f"sel_{item_id}", "selection", item_id))
        
        # Salva referência do retângulo
        self.selection_rectangles[item_id] = sel_rect_id

    def _render_model_image(self):
        """Renderiza a imagem de fundo do modelo"""
        # Verifica se temos um modelo para renderizar
        if not self.model_img:
            return
            
        # Implementa um sistema de cache para evitar redimensionamento repetido
        if not hasattr(self, 'model_cache') or not self.model_cache or self.model_cache.get('zoom') != self.zoom_factor:
            try:
                # Redimensiona com o zoom
                scaled_width = int(self.model_img.width * self.zoom_factor)
                scaled_height = int(self.model_img.height * self.zoom_factor)
                
                # Usa um algoritmo eficiente para redimensionar
                # Se o zoom for maior que 1, usa LANCZOS para qualidade
                # Se for menor, usa BILINEAR que é mais rápido para redução
                scale_method = Image.LANCZOS if self.zoom_factor >= 1 else Image.BILINEAR
                
                # Redimensiona apenas quando necessário
                scaled_img = self.model_img.resize((scaled_width, scaled_height), scale_method)
                self.model_tk = ImageTk.PhotoImage(scaled_img)
                
                # Armazena no cache
                self.model_cache = {
                    'zoom': self.zoom_factor,
                    'img': self.model_tk
                }
            except Exception as e:
                logging.error(f"Error resizing model: {e}")
                return
        else:
            # Usa a versão em cache
            self.model_tk = self.model_cache['img']
            
        # Desenha a imagem do modelo
        self.canvas.create_image(0, 0, image=self.model_tk, anchor='nw', tags="model")

    def add_text(self):
        if not self.model_img:
            messagebox.showwarning("Aviso", "Carregue primeiro um modelo.")
            return
            
        # Nomeação automática da camada
        layer_name = f"Texto {len([k for k,v in self.items.items() if v['tipo']=='texto'])+1}"
        
        families = list(tkfont.families(self))
        families.sort()
        family = "Arial" if "Arial" in families else families[0]
        size = config.FONT_SIZE
        color = "#000000"  # Preto por defeito
        
        # Posiciona no centro da imagem
        x = self.model_img.width//2
        y = self.model_img.height//2
        
        # Define largura inicial da caixa de texto como 1/3 da largura do modelo
        width = min(self.model_img.width // 3, 400)
        height = min(self.model_img.height // 5, 200)
        
        # Cria o item com texto vazio
        item_id = f"text_{len(self.items)}"
        self.items[item_id] = {
            'tipo': 'texto',
            'texto': 'Texto Editável',
            'font_family': family,
            'size': size,
            'color': color,
            'xy': [x, y],
            'width': width,  # Largura inicial proporcional ao modelo
            'height': height,  # Altura inicial proporcional ao modelo
            'bg_color': ""  # Força fundo transparente para todas as caixas de texto
        }
        
        # Marca como visível
        self.visible_items[item_id] = True
        
        # Adiciona à lista de camadas
        self.add_layer_to_list(item_id, layer_name)
        
        # Atualiza o canvas
        self.schedule_redraw(immediate=True)
        
        # Mostra propriedades
        self.show_properties(item_id)

    def add_image(self):
        if not self.model_img:
            messagebox.showwarning("Aviso", "Carregue primeiro um modelo.")
            return
            
        path = filedialog.askopenfilename(filetypes=[("PNG/JPG", ".png .jpg")])
        if not path: return
        
        # Nomeação automática da camada
        base_name = os.path.basename(path)
        layer_name = f"Imagem: {base_name}"
        
        # Cria uma janela de progresso
        progress = tk.Toplevel(self)
        progress.title("Carregando Imagem")
        progress.geometry("300x80")
        progress.transient(self)
        
        tk.Label(progress, text="A carregar imagem...").pack(pady=10)
        progress_bar = ttk.Progressbar(progress, mode="indeterminate")
        progress_bar.pack(fill=tk.X, padx=20, pady=10)
        progress_bar.start()
        progress.update()
        
        # Captura as variáveis para usar na função aninhada
        _self = self
        _progress = progress
        _path = path
        _layer_name = layer_name
        
        # Função que executa em thread separada
        def load_image_thread():
            try:
                im = Image.open(_path).convert("RGBA")
                # Coloca a imagem carregada na fila para a thread principal
                _self.after(0, lambda: _self._finish_adding_image(im, _path, _layer_name, _progress))
            except Exception as e:
                # Em caso de erro, notifica a thread principal
                _self.after(0, lambda: _self._handle_load_error(str(e), _progress))
                
        # Inicia o carregamento em thread separada
        threading.Thread(target=load_image_thread, daemon=True).start()
    
    def _finish_adding_image(self, img, path, layer_name, progress_window):
        """Finaliza o processo de adicionar imagem na thread principal"""
        try:
            size = (img.width//2, img.height//2)  # Reduz para metade por defeito
            
            # Posiciona no centro da imagem
            x = (self.model_img.width - size[0]) // 2
            y = (self.model_img.height - size[1]) // 2
            opacity = 1.0
            
            # Cria o item
            item_id = f"img_{len(self.items)}"
            self.items[item_id] = {
                'tipo': 'imagem',
                'path': path,
                'size': size,
                'xy': [x, y],
                'opacity': opacity,
                'preserve_ratio': True  # Adiciona opção para preservar proporção
            }
            
            # Marca como visível
            self.visible_items[item_id] = True
            
            # Adiciona à lista de camadas
            self.add_layer_to_list(item_id, layer_name)
            
            # Atualiza o canvas
            self.schedule_redraw(immediate=True)
            
            # Mostra propriedades
            self.show_properties(item_id)
            
            # Fecha janela de progresso
            progress_window.destroy()
        except Exception as e:
            progress_window.destroy()
            messagebox.showerror("Erro", f"Não foi possível processar a imagem: {str(e)}")

    def on_press(self, event):
        """Handler para clique no canvas"""
        # Converte coordenadas do evento para coordenadas reais do canvas
        canvas_x = self.canvas.canvasx(event.x)
        canvas_y = self.canvas.canvasy(event.y)
        
        # Inicializa dados de arrasto com coordenadas de canvas
        self.drag_data = {"x": canvas_x, "y": canvas_y, "item": None}
        
        #print(f"DEBUG - on_press: Inicializando drag_data com canvas_x={canvas_x}, canvas_y={canvas_y}")
        
        # Encontrar todos os itens na posição do clique (usando coordenadas de canvas)
        items_at_position = self.canvas.find_overlapping(canvas_x-2, canvas_y-2, canvas_x+2, canvas_y+2)
        
        # Se não encontrou nada, tenta com uma área maior
        if not items_at_position:
            items_at_position = self.canvas.find_overlapping(canvas_x-10, canvas_y-10, canvas_x+10, canvas_y+10)
            #print(f"Ampliando busca para área maior: {canvas_x-10}, {canvas_y-10}, {canvas_x+10}, {canvas_y+10}")
            
        # Debug: Mostrar os itens encontrados e suas tags
        #if items_at_position:
        #    print(f"Itens encontrados na posição canvas ({canvas_x}, {canvas_y}):")
        #    for item in items_at_position:
        #        tags = self.canvas.gettags(item)
        #        print(f"  Item: {item}, Tags: {tags}")
        #else:
        #    print(f"Nenhum item encontrado na posição canvas ({canvas_x}, {canvas_y})")
            
        # 1. Primeiro tenta encontrar um item diretamente
        item_id = None
        
        # Verificar cada item encontrado e suas tags
        for canvas_item in items_at_position:
            tags = self.canvas.gettags(canvas_item)
            
            # Ignorar o item 'model' que é o fundo
            if 'model' in tags:
                continue
                
            # Verificar se alguma das tags está no dicionário de itens
            for tag in tags:
                if tag in self.items:
                    item_id = tag
                    #print(f"Item encontrado: {item_id}")
                    break
                # Verificar se é uma tag de item específico (item_XXX)
                elif tag.startswith('item_'):
                    item_tag = tag.replace('item_', '')
                    if item_tag in self.items:
                        item_id = item_tag
                        #print(f"Item encontrado via tag item_: {item_id}")
                        break
            if item_id:
                break
                
        # 2. Se não encontrou, procura por qualquer tag 'text_box' que é o retângulo do texto
        if not item_id:
            for canvas_item in items_at_position:
                tags = self.canvas.gettags(canvas_item)
                if "text_box" in tags or "text" in tags or "image" in tags:
                    # Extrair o item_id associado ao elemento
                    for tag in tags:
                        if tag in self.items:
                            item_id = tag
                            #print(f"Item encontrado via tipo de elemento: {item_id}")
                            break
                        # Verificar se é uma tag de item específico (item_XXX)
                        elif tag.startswith('item_'):
                            item_tag = tag.replace('item_', '')
                            if item_tag in self.items:
                                item_id = item_tag
                                #print(f"Item encontrado via tag item_: {item_id}")
                                break
                if item_id:
                    break
                    
        # Debug final
        if item_id:
            #print(f"Item selecionado: {item_id}, Tipo: {self.items[item_id]['tipo']}")
            
            if item_id:
                # Se não estiver pressionando Shift, limpa a seleção anterior
                if not (event.state & 0x0001):  # Shift não está pressionado
                    self.selected_items.clear()
                    
                # Seleciona o item atual
                self.current_item = item_id
                self.selected_items.add(item_id)
                
                # Guarda informações para arrastar
                self.drag_data["item"] = item_id
            #print(f"Preparando para arrastar: {self.drag_data}")
            
            # Atualiza a seleção visual na lista de camadas
            self.select_layer_without_properties(item_id)
            
            # Mostra as propriedades do item selecionado
            self.show_properties(item_id)
            
            # Depuração após show_properties para confirmar que drag_data ainda está intacto
            #print(f"DEBUG - Após show_properties: drag_data = {self.drag_data}")
                
            # Atualiza o canvas para mostrar a seleção
            self.schedule_redraw()
            
            # Depuração após schedule_redraw para confirmar que drag_data ainda está intacto
            #print(f"DEBUG - Após schedule_redraw: drag_data = {self.drag_data}")
            return
        
        # Clicou no fundo, limpa a seleção
        self.selected_items.clear()
        self.current_item = None
        self.schedule_redraw()
        
        # Limpa a seleção na lista de camadas
        for item_id in self.layer_widgets:
            self.layer_widgets[item_id]['frame'].config(bg='#333333')
            self.layer_widgets[item_id]['label'].config(bg='#333333')
        
        # Limpa painel de propriedades
        for w in self.props_frame.winfo_children():
            w.destroy()

    def select_layer_without_properties(self, item_id):
        """Seleciona uma camada sem atualizar o painel de propriedades"""
        if item_id not in self.items:
            #print(f"Item não encontrado: {item_id}")
            return
            
        # Atualiza a seleção atual
        old_selection = self.current_item
        self.current_item = item_id
        self.selected_items = {item_id}  # Apenas este item na seleção
        
        # Atualiza visual na lista
        for layer_id in self.layer_widgets:
            # Reset todas as camadas
            self.layer_widgets[layer_id]['frame'].config(bg='#333333')
            self.layer_widgets[layer_id]['label'].config(bg='#333333')
            
        # Destaca a selecionada
        if item_id in self.layer_widgets:
            self.layer_widgets[item_id]['frame'].config(bg='#4a90d9')
            self.layer_widgets[item_id]['label'].config(bg='#4a90d9')

    def on_drag(self, event):
        """Handler para arrastar elementos"""
        # Verifica se temos um item para arrastar
        if not self.drag_data.get("item"):
            #print("Nenhum item para arrastar")
            return
            
        # Pega o ID do item
        item_id = self.drag_data["item"]
            
        # Verifica se o ID é válido
        if item_id not in self.items:
            #print(f"Item inválido: {item_id}")
            return
        
        # Converte coordenadas do evento para coordenadas reais do canvas
        canvas_x = self.canvas.canvasx(event.x)
        canvas_y = self.canvas.canvasy(event.y)
        
        #print(f"Arrastando item: {item_id}, de ({self.drag_data['x']}, {self.drag_data['y']}) para canvas ({canvas_x}, {canvas_y})")
        
        # Calcula o deslocamento usando coordenadas de canvas
        delta_x = canvas_x - self.drag_data["x"]
        delta_y = canvas_y - self.drag_data["y"]
        
        # Conversão do delta para coordenadas reais (sem zoom)
        real_delta_x = delta_x / self.zoom_factor
        real_delta_y = delta_y / self.zoom_factor
        
        #print(f"DEBUG - on_drag: Calculando deslocamento: delta=({delta_x}, {delta_y}), real_delta=({real_delta_x}, {real_delta_y})")
        
        # Atualiza as coordenadas do item no dicionário de dados
        self.items[item_id]['xy'][0] += real_delta_x
        self.items[item_id]['xy'][1] += real_delta_y
        
        #print(f"DEBUG - on_drag: Nova posição do item {item_id}: xy=({self.items[item_id]['xy'][0]}, {self.items[item_id]['xy'][1]})")
        
        # Atualiza as coordenadas de todos os objetos selecionados juntos (para seleção múltipla)
        for selected_id in self.selected_items:
            if selected_id != item_id and selected_id in self.items:
                # Atualiza os dados
                self.items[selected_id]['xy'][0] += real_delta_x
                self.items[selected_id]['xy'][1] += real_delta_y
        
        # Atualiza dados de arrasto com novas coordenadas de canvas
        self.drag_data["x"] = canvas_x
        self.drag_data["y"] = canvas_y
        
        #print(f"DEBUG - on_drag: Atualizado drag_data para: {self.drag_data}")
        
        # Redesenha o canvas diretamente sem usar schedule_redraw
        try:
            #print("DEBUG - on_drag: Iniciando redraw_canvas")
            self.redraw_canvas()
            #print("DEBUG - on_drag: Concluído redraw_canvas")
        except Exception as e:
            #print(f"Erro ao redesenhar canvas durante arrasto: {e}")
            pass

    def on_release(self, event):
        """Finaliza a operação de arrasto"""
        try:
            # Restaura o cursor normal
            self.canvas.config(cursor="")
            
            # Força uma atualização final quando solta o mouse
            self.redraw_canvas()
            
            # Obtém o item atual antes de limpar drag_data
            current_drag_item = self.drag_data.get("item")
            
            # Converte coordenadas do evento para coordenadas reais do canvas
            canvas_x = self.canvas.canvasx(event.x)
            canvas_y = self.canvas.canvasy(event.y)
            
            # Reset das variáveis de arrasto, mas mantendo o item selecionado
            # Isso permite que o próximo arrasto funcione corretamente
            if current_drag_item:
                self.drag_data = {"x": canvas_x, "y": canvas_y, "item": current_drag_item}
            else:
                self.drag_data = {"x": 0, "y": 0, "item": None}
                
        except Exception as e:
            logging.error(f"Erro ao finalizar arrasto: {e}")
            
            # Garante o reset do estado
            self.drag_data = {"x": 0, "y": 0, "item": None}
            self.canvas.config(cursor="")

    def show_properties(self, item_id):
        """Exibe o painel de propriedades para o item selecionado"""
        # Guarda os dados de arrasto atuais antes de qualquer modificação
        preserve_drag_data = self.drag_data.copy() if hasattr(self, 'drag_data') else None
        
        # Limpa frame de props
        for w in self.props_frame.winfo_children():
            w.destroy()
        
        if item_id not in self.items:
            return
        
        info = self.items[item_id]
        
        # Título do painel
        name_frame = tk.Frame(self.props_frame, bg='#252525')
        name_frame.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Label(name_frame, text="Nome:", bg='#252525', fg='white').pack(side=tk.LEFT)
        
        self.layer_name_var = tk.StringVar(value=self.layer_names.get(item_id, f"Camada {len(self.item_order)}"))
        name_entry = tk.Entry(name_frame, textvariable=self.layer_name_var, 
                             bg='#333333', fg='white', insertbackground='white')
        name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Atualiza o nome ao digitar
        def update_name(*args):
            new_name = self.layer_name_var.get()
            self.layer_names[item_id] = new_name
                
            # Atualiza apenas o texto da camada específica sem redesenhar toda a sidebar
            if item_id in self.layer_widgets and 'name_var' in self.layer_widgets[item_id]:
                    self.layer_widgets[item_id]['name_var'].set(new_name)
        
        self.layer_name_var.trace_add("write", update_name)
        
        def update_prop(key, val):
            info[key] = val
            # Preserva os dados de arrasto antes de redesenhar
            temp_drag_data = self.drag_data.copy() if hasattr(self, 'drag_data') else None
            self.redraw_canvas()
            # Restaura os dados de arrasto após redesenhar
            if temp_drag_data:
                self.drag_data = temp_drag_data

        # Cria um notebook para separar as propriedades em abas
        notebook = ttk.Notebook(self.props_frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
        if info['tipo'] == 'texto':
            # Aba 1: Conteúdo do texto
            texto_frame = tk.Frame(notebook, bg='#252525')
            notebook.add(texto_frame, text="Texto")
            
            # Editar texto
            text_frame = tk.Frame(texto_frame, bg='#252525')
            text_frame.pack(fill=tk.X, padx=5, pady=5)
            
            tk.Label(text_frame, text="Texto:", bg='#252525', fg='white').pack(anchor=tk.W)
            
            # Área de texto multilinha para permitir texto mais longo
            text_entry = tk.Text(text_frame, height=4, width=30,
                                bg='#333333', fg='white', insertbackground='white')
            text_entry.insert("1.0", info['texto'])
            text_entry.pack(fill=tk.X, pady=2)
            
            # Função para atualizar o texto quando modificado
            def update_text(*args):
                new_text = text_entry.get("1.0", "end-1c")
                info['texto'] = new_text
                # Preserva os dados de arrasto durante o redesenho
                temp_drag_data = self.drag_data.copy() if hasattr(self, 'drag_data') else None
                self.redraw_canvas()
                # Restaura após o redesenho
                if temp_drag_data:
                    self.drag_data = temp_drag_data
            
            # Atualiza ao perder foco e ao soltar qualquer tecla
            text_entry.bind("<FocusOut>", lambda e: update_text())
            text_entry.bind("<KeyRelease>", lambda e: update_text())
            
            # Barra de formatação estilo Word (negrito, itálico, sublinhado)
            format_frame = tk.Frame(texto_frame, bg='#252525')
            format_frame.pack(fill=tk.X, padx=5, pady=5)
            
            # Inicializa valores de formatação ou usa padrões
            info.setdefault('bold', False)
            info.setdefault('italic', False)
            info.setdefault('text_align', 'left')
            
            # Sistema para verificar fontes disponíveis
            def check_font_styles_available(font_family):
                """Verifica quais estilos estão disponíveis para a fonte selecionada"""
                styles_available = {
                    'bold': False,
                    'italic': False
                }
                
                font_name_lower = font_family.lower()
                #print(f"Verificando estilos para fonte: '{font_name_lower}'")
                
                # Verifica se a fonte está no mapeamento
                if font_name_lower in self.fonts_map:
                    font_data = self.fonts_map[font_name_lower]
                    #print(f"Dados da fonte encontrados: {type(font_data)}")
                    
                    # Verifica o formato do JSON - pode ser o formato antigo ou o novo
                    if isinstance(font_data, dict) and 'styles' in font_data:
                        # Novo formato com informações de estilos
                        styles = font_data['styles']
                        #print(f"Formato novo! Estilos encontrados: {styles}")
                        styles_available['bold'] = styles.get('bold', False) or styles.get('bold_italic', False)
                        styles_available['italic'] = styles.get('italic', False) or styles.get('bold_italic', False)
                    else:
                        # Formato antigo - acessa diretamente
                        #print(f"Formato antigo! Dados: {font_data}")
                        styles_available['bold'] = font_data.get('bold', False) or font_data.get('bold_italic', False)
                        styles_available['italic'] = font_data.get('italic', False) or font_data.get('bold_italic', False)
                    
                    #print(f"Resultado final: bold={styles_available['bold']}, italic={styles_available['italic']}")
                else:
                    # Fonte não encontrada no mapeamento, usar valores padrão
                    # Vamos ser conservadores e assumir que existe apenas a versão normal
                    styles_available['bold'] = False
                    styles_available['italic'] = False
                    #print(f"Aviso: Fonte '{font_family}' não encontrada no mapeamento.")
                
                return styles_available
            
            # Botões de formatação
            bold_btn = tk.Button(format_frame, text="B", width=2, font=('Arial', 9, 'bold'),
                                bg='#444444' if not info.get('bold', False) else '#666666', 
                                fg='white', relief=tk.FLAT)
            bold_btn.pack(side=tk.LEFT, padx=2)
            
            italic_btn = tk.Button(format_frame, text="I", width=2, font=('Arial', 9, 'italic'),
                                 bg='#444444' if not info.get('italic', False) else '#666666', 
                                 fg='white', relief=tk.FLAT)
            italic_btn.pack(side=tk.LEFT, padx=2)

            # Botão Placeholders
            def show_placeholders():
                # Placeholders globais
                globais = []
                for k in getattr(config, 'GLOBAL_PLACEHOLDERS', {}).keys():
                    globais.append(f'{{GLOBAL_{k}}}')
                # Placeholders do CSV
                csv_cols = []
                if hasattr(self, 'df') and self.df is not None:
                    csv_cols = [f'{{{col}}}' for col in self.df.columns]
                # Mensagem
                msg = 'Placeholders globais disponíveis:\n' + '\n'.join(globais)
                if csv_cols:
                    msg += '\n\nPlaceholders do CSV disponíveis:\n' + '\n'.join(csv_cols)
                messagebox.showinfo('Placeholders disponíveis', msg)
            placeholders_btn = tk.Button(format_frame, text='Placeholders', command=show_placeholders, bg='#444444', fg='white')
            placeholders_btn.pack(side=tk.LEFT, padx=8)
            
            # Funções para atualizar formatação
            def toggle_bold():
                if not bold_btn['state'] == 'disabled':
                    info['bold'] = not info.get('bold', False)
                    bold_btn.config(bg='#666666' if info['bold'] else '#444444')
                    update_font_style()
                
            def toggle_italic():
                if not italic_btn['state'] == 'disabled':
                    info['italic'] = not info.get('italic', False)
                    italic_btn.config(bg='#666666' if info['italic'] else '#444444')
                    update_font_style()
            
            def update_font_style():
                # Preserva os dados de arrasto durante o redesenho
                temp_drag_data = self.drag_data.copy() if hasattr(self, 'drag_data') else None
                self.redraw_canvas()
                # Restaura após o redesenho
                if temp_drag_data:
                    self.drag_data = temp_drag_data
            
            bold_btn.config(command=toggle_bold)
            italic_btn.config(command=toggle_italic)
            
            # Inicialmente verifica as opções disponíveis para a fonte atual
            styles_available = check_font_styles_available(info['font_family'])
            
            # Configura o estado inicial dos botões
            bold_btn.config(state='normal' if styles_available['bold'] else 'disabled')
            italic_btn.config(state='normal' if styles_available['italic'] else 'disabled')
            
            # Se a fonte não suporta o estilo, mas está ativado, desativa-o
            if not styles_available['bold'] and info.get('bold', False):
                info['bold'] = False
                bold_btn.config(bg='#444444')
                
            if not styles_available['italic'] and info.get('italic', False):
                info['italic'] = False
                italic_btn.config(bg='#444444')
            
            # Botões de alinhamento
            align_frame = tk.Frame(texto_frame, bg='#252525')
            align_frame.pack(fill=tk.X, padx=5, pady=5)
            
            tk.Label(align_frame, text="Alinhamento:", bg='#252525', fg='white').pack(side=tk.LEFT)
            
            align_left_btn = tk.Button(align_frame, text="⫷", width=2,
                                     bg='#666666' if info.get('text_align', 'left') == 'left' else '#444444', 
                                     fg='white', relief=tk.FLAT)
            align_left_btn.pack(side=tk.LEFT, padx=2)
            
            align_center_btn = tk.Button(align_frame, text="⫶", width=2,
                                       bg='#666666' if info.get('text_align', 'left') == 'center' else '#444444', 
                                       fg='white', relief=tk.FLAT)
            align_center_btn.pack(side=tk.LEFT, padx=2)
            
            align_right_btn = tk.Button(align_frame, text="⫸", width=2,
                                      bg='#666666' if info.get('text_align', 'left') == 'right' else '#444444', 
                                      fg='white', relief=tk.FLAT)
            align_right_btn.pack(side=tk.LEFT, padx=2)
            
            align_justify_btn = tk.Button(align_frame, text="≡", width=2,
                                        bg='#666666' if info.get('text_align', 'left') == 'justify' else '#444444', 
                                        fg='white', relief=tk.FLAT)
            align_justify_btn.pack(side=tk.LEFT, padx=2)
            
            # Funções para atualizar alinhamento
            def set_align(align):
                info['text_align'] = align
                align_left_btn.config(bg='#666666' if align == 'left' else '#444444')
                align_center_btn.config(bg='#666666' if align == 'center' else '#444444')
                align_right_btn.config(bg='#666666' if align == 'right' else '#444444')
                align_justify_btn.config(bg='#666666' if align == 'justify' else '#444444')
                # Preserva os dados de arrasto durante o redesenho
                temp_drag_data = self.drag_data.copy() if hasattr(self, 'drag_data') else None
                self.redraw_canvas()
                # Restaura após o redesenho
                if temp_drag_data:
                    self.drag_data = temp_drag_data
            
            align_left_btn.config(command=lambda: set_align('left'))
            align_center_btn.config(command=lambda: set_align('center'))
            align_right_btn.config(command=lambda: set_align('right'))
            align_justify_btn.config(command=lambda: set_align('justify'))
            
            # Aba 2: Fonte
            fonte_frame = tk.Frame(notebook, bg='#252525')
            notebook.add(fonte_frame, text="Fonte")
            
            # Fonte (com dropdown)
            font_frame = tk.Frame(fonte_frame, bg='#252525')
            font_frame.pack(fill=tk.X, padx=5, pady=5)
            
            tk.Label(font_frame, text="Fonte:", bg='#252525', fg='white').pack(anchor=tk.W)
            
            # Lista de fontes do sistema
            families = list(tkfont.families(self))
            families.sort()
            
            font_var = tk.StringVar(value=info['font_family'])
            font_combo = ttk.Combobox(font_frame, textvariable=font_var, values=families,
                                    width=20, state="readonly")
            font_combo.pack(fill=tk.X, pady=2)
            
            def on_font_change(event):
                new_font = font_var.get()
                info['font_family'] = new_font
                
                # Adiciona logs de depuração
                #print(f"\n--- Verificando estilos para fonte '{new_font}' ---")
                
                # Verifica quais estilos estão disponíveis para a nova fonte
                styles_available = check_font_styles_available(new_font)
                #print(f"Estilos disponíveis: {styles_available}")
                
                # Atualiza o estado dos botões
                bold_btn.config(state='normal' if styles_available['bold'] else 'disabled')
                italic_btn.config(state='normal' if styles_available['italic'] else 'disabled')
                
                #print(f"Estado do botão negrito: {bold_btn['state']}")
                #print(f"Estado do botão itálico: {italic_btn['state']}")
                
                # Verifica o estado dos estilos no item
                #print(f"Estado atual do item - Negrito: {info.get('bold', False)}, Itálico: {info.get('italic', False)}")
                
                # Se a fonte não suporta o estilo atual, desativa-o
                if not styles_available['bold'] and info.get('bold', False):
                    info['bold'] = False
                    bold_btn.config(bg='#444444')
                    #print("Desativando estilo negrito que não é suportado")
                    
                if not styles_available['italic'] and info.get('italic', False):
                    info['italic'] = False
                    italic_btn.config(bg='#444444')
                    #print("Desativando estilo itálico que não é suportado")
                
                #print("------------------------------------------\n")
                
                # Preserva os dados de arrasto durante o redesenho
                temp_drag_data = self.drag_data.copy() if hasattr(self, 'drag_data') else None
                self.redraw_canvas()
                # Restaura após o redesenho
                if temp_drag_data:
                    self.drag_data = temp_drag_data
                
            font_combo.bind("<<ComboboxSelected>>", on_font_change)
            
            # Tamanho da fonte
            size_frame = tk.Frame(fonte_frame, bg='#252525')
            size_frame.pack(fill=tk.X, padx=5, pady=5)
            
            tk.Label(size_frame, text="Tamanho:", bg='#252525', fg='white').pack(anchor=tk.W)
            
            # Frame para conter o slider e a caixa de entrada
            size_control_frame = tk.Frame(size_frame, bg='#252525')
            size_control_frame.pack(fill=tk.X, pady=2)
            
            # Calcula o tamanho máximo da fonte baseado na altura do modelo (aproximadamente 10% da altura)
            max_size = 72  # Valor padrão
            if self.model_img:
                max_size = int(self.model_img.height * 0.1)  # 10% da altura do modelo
                max_size = max(72, max_size)  # Não deixa ser menor que 72
            
            size_var = tk.IntVar(value=info['size'])
            size_slider = ttk.Scale(size_control_frame, from_=8, to=max_size, 
                                  orient=tk.HORIZONTAL, variable=size_var)
            size_slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            
            # Caixa de entrada para inserir o valor diretamente
            size_entry = tk.Entry(size_control_frame, width=5, bg='#333333', fg='white', 
                                 insertbackground='white', justify='center')
            size_entry.pack(side=tk.LEFT, padx=5)
            size_entry.insert(0, str(info['size']))
            
            # Label para mostrar o valor atual (podemos remover já que temos o entry)
            # size_label = tk.Label(size_frame, text=str(info['size']), bg='#252525', fg='white')
            # size_label.pack(anchor=tk.E, padx=5)
            
            def on_size_change(event=None):
                try:
                    val = size_var.get()
                    if val > 0:
                        info['size'] = val
                        size_entry.delete(0, tk.END)
                        size_entry.insert(0, str(val))
                        self.redraw_canvas()
                except ValueError:
                    pass
            
            def on_size_entry_change(event=None):
                try:
                    val = int(size_entry.get())
                    if val > 0:
                        if val < 8:
                            val = 8
                        elif val > max_size:
                            val = max_size
                        info['size'] = val
                        size_var.set(val)
                        self.redraw_canvas()
                except ValueError:
                    # Se o valor não for válido, restaura o valor anterior
                    size_entry.delete(0, tk.END)
                    size_entry.insert(0, str(info['size']))
                
            # Atualiza quando arrastar o slider
            size_slider.bind("<B1-Motion>", on_size_change)
            size_slider.bind("<ButtonRelease-1>", on_size_change)
            
            # Atualiza quando mudar o valor na caixa de entrada
            size_entry.bind("<Return>", on_size_entry_change)
            size_entry.bind("<FocusOut>", on_size_entry_change)
            
            # Cor com botão de seleção
            color_frame = tk.Frame(fonte_frame, bg='#252525')
            color_frame.pack(fill=tk.X, padx=5, pady=5)
            
            tk.Label(color_frame, text="Cor do texto:", bg='#252525', fg='white').pack(side=tk.LEFT)
            
            # Amostra de cor
            color_sample = tk.Canvas(color_frame, width=20, height=20, bg=info['color'])
            color_sample.pack(side=tk.LEFT, padx=5)
            
            def choose_color():
                color = colorchooser.askcolor(initialcolor=info['color'])[1]
                if color:
                    info['color'] = color
                    color_sample.config(bg=color)
                    self.redraw_canvas()
                    
            color_btn = tk.Button(color_frame, text="Escolher Cor", command=choose_color,
                                bg='#444444', fg='white')
            color_btn.pack(side=tk.LEFT, padx=5)
            
            # Aba 3: Tamanho e cor de fundo
            bg_frame = tk.Frame(notebook, bg='#252525')
            notebook.add(bg_frame, text="Caixa")
            
            # Cor de fundo - Comentado por não funcionar com o texto
            """
            bg_color_frame = tk.Frame(bg_frame, bg='#252525')
            bg_color_frame.pack(fill=tk.X, padx=5, pady=5)
            
            # Inicializa cor de fundo se não existir
            if 'bg_color' not in info:
                info['bg_color'] = ""  # Transparente por padrão
                
            tk.Label(bg_color_frame, text="Cor de fundo:", bg='#252525', fg='white').pack(side=tk.LEFT)
            
            # Amostra de cor
            bg_color_sample = tk.Canvas(bg_color_frame, width=20, height=20, 
                                      bg=info.get('bg_color', '') or '#252525')
            bg_color_sample.pack(side=tk.LEFT, padx=5)
            
            def choose_bg_color():
                color = colorchooser.askcolor(initialcolor=info.get('bg_color', '#ffffff'))[1]
                if color:
                    info['bg_color'] = color
                    bg_color_sample.config(bg=color)
                    self.redraw_canvas()
                    
            bg_color_btn = tk.Button(bg_color_frame, text="Escolher Cor", command=choose_bg_color,
                                   bg='#444444', fg='white',
                                   state=tk.DISABLED if not info.get('bg_color', '') else tk.NORMAL)
            bg_color_btn.pack(side=tk.LEFT, padx=5)
            
            # Caixa de transparência
            bg_transparent_frame = tk.Frame(bg_frame, bg='#252525')
            bg_transparent_frame.pack(fill=tk.X, padx=5, pady=5)
            
            bg_transparent_var = tk.BooleanVar(value=not info.get('bg_color', ''))
            
            def toggle_transparent():
                if bg_transparent_var.get():
                    # Guarda a cor atual antes de tornar transparente
                    if info.get('bg_color'):
                        info['_last_bg_color'] = info['bg_color']
                    info['bg_color'] = ""  # Transparente
                    bg_color_sample.config(bg='#252525')  # Cor do fundo da UI para indicar transparência
                    bg_color_btn.config(state=tk.DISABLED)  # Desativa botão de escolha de cor
                else:
                    # Restaura cor anterior ou usa branco
                    last_color = info.get('_last_bg_color', '#ffffff')
                    info['bg_color'] = last_color
                    bg_color_sample.config(bg=last_color)
                    bg_color_btn.config(state=tk.NORMAL)  # Ativa botão de escolha de cor
                self.redraw_canvas()
                
            bg_transparent_check = tk.Checkbutton(bg_transparent_frame, text="Fundo transparente",
                                               variable=bg_transparent_var, command=toggle_transparent,
                                               bg='#252525', fg='white',
                                               selectcolor='#333333', activebackground='#252525')
            bg_transparent_check.pack(anchor=tk.W)
            """
            
            # Sempre configura fundo transparente para todas as caixas de texto
            info['bg_color'] = ""  # Força fundo transparente
            
            # Largura da caixa
            width_frame = tk.Frame(bg_frame, bg='#252525')
            width_frame.pack(fill=tk.X, padx=5, pady=5)
            
            tk.Label(width_frame, text="Largura:", bg='#252525', fg='white').pack(anchor=tk.W)
            
            # Frame para conter o slider e a caixa de entrada
            width_control_frame = tk.Frame(width_frame, bg='#252525')
            width_control_frame.pack(fill=tk.X, pady=2)
            
            # Inicializa largura se não existir
            if 'width' not in info:
                info['width'] = 200  # Valor padrão
                
            # Obtém a largura do modelo para definir o intervalo máximo do slider
            max_width = 800
            if self.model_img:
                max_width = self.model_img.width  # Permite até a largura do modelo
                
            width_var = tk.IntVar(value=info.get('width', 200))
            width_slider = ttk.Scale(width_control_frame, from_=50, to=max_width, 
                                   orient=tk.HORIZONTAL, variable=width_var)
            width_slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            
            # Caixa de entrada para inserir o valor diretamente
            width_entry = tk.Entry(width_control_frame, width=5, bg='#333333', fg='white', 
                                  insertbackground='white', justify='center')
            width_entry.pack(side=tk.LEFT, padx=5)
            width_entry.insert(0, str(info.get('width', 200)))
            
            def on_width_change(event=None):
                try:
                    w = width_var.get()
                    if w > 0:
                        info['width'] = w
                        width_entry.delete(0, tk.END)
                        width_entry.insert(0, str(w))
                        self.redraw_canvas()
                except:
                    pass
                    
            def on_width_entry_change(event=None):
                try:
                    w = int(width_entry.get())
                    if w <= 0:
                        w = 50
                    if w > max_width:
                        w = max_width
                    
                    info['width'] = w
                    width_var.set(w)
                    self.redraw_canvas()
                except ValueError:
                    # Se o valor não for válido, restaura o valor anterior
                    width_entry.delete(0, tk.END)
                    width_entry.insert(0, str(info['width']))
                
            # Atualiza quando arrastar os sliders
            width_slider.bind("<B1-Motion>", on_width_change)
            width_slider.bind("<ButtonRelease-1>", on_width_change)
            
            # Atualiza quando mudar o valor na caixa de entrada
            width_entry.bind("<Return>", on_width_entry_change)
            width_entry.bind("<FocusOut>", on_width_entry_change)
        
        elif info['tipo'] == 'imagem':
            # Primeira aba: Aparência
            aparencia_frame = tk.Frame(notebook, bg='#252525')
            notebook.add(aparencia_frame, text="Aparência")
            
            # Opacidade
            opacity_frame = tk.Frame(aparencia_frame, bg='#252525')
            opacity_frame.pack(fill=tk.X, padx=5, pady=5)
            
            tk.Label(opacity_frame, text="Opacidade:", bg='#252525', fg='white').pack(anchor=tk.W)
            
            opacity_var = tk.DoubleVar(value=info['opacity'])
            # Cria um frame para conter o slider e a entrada
            opacity_control_frame = tk.Frame(opacity_frame, bg='#252525')
            opacity_control_frame.pack(fill=tk.X, pady=2)
            
            # Insere o slider no frame de controle
            opacity_slider = ttk.Scale(opacity_control_frame, from_=0, to=1, 
                                     orient=tk.HORIZONTAL, variable=opacity_var)
            opacity_slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            
            # Adiciona uma caixa de entrada para valores precisos
            opacity_entry = tk.Entry(opacity_control_frame, width=5, bg='#333333', fg='white', 
                                   insertbackground='white', justify='center')
            opacity_entry.pack(side=tk.LEFT, padx=2)
            opacity_entry.insert(0, f"{int(info['opacity']*100)}")
            
            # Adiciona o símbolo de percentagem
            tk.Label(opacity_control_frame, text="%", bg='#252525', fg='white').pack(side=tk.LEFT)
            
            def on_opacity_change(event=None):
                try:
                    val = opacity_var.get()
                    info['opacity'] = val
                    opacity_entry.delete(0, tk.END)
                    opacity_entry.insert(0, f"{int(val*100)}")
                    self.redraw_canvas()
                except:
                    pass
                
            def on_opacity_entry_change(event=None):
                try:
                    val = float(opacity_entry.get()) / 100
                    if 0 <= val <= 1:
                        info['opacity'] = val
                        opacity_var.set(val)
                        self.redraw_canvas()
                    else:
                        # Se estiver fora do intervalo, ajusta
                        if val < 0:
                            val = 0
                        elif val > 1:
                            val = 1
                        info['opacity'] = val
                        opacity_entry.delete(0, tk.END)
                        opacity_entry.insert(0, str(int(val*100)))
                        opacity_var.set(val)
                except ValueError:
                    # Se o valor não for válido, restaura o valor anterior
                    opacity_entry.delete(0, tk.END)
                    opacity_entry.insert(0, str(int(info['opacity']*100)))
                
            # Atualiza quando arrastar o slider
            opacity_slider.bind("<B1-Motion>", on_opacity_change)
            opacity_slider.bind("<ButtonRelease-1>", on_opacity_change)
            
            # Atualiza quando mudar o valor na caixa de entrada
            opacity_entry.bind("<Return>", on_opacity_entry_change)
            opacity_entry.bind("<FocusOut>", on_opacity_entry_change)
            
            # Segunda aba: Tamanho
            tamanho_frame = tk.Frame(notebook, bg='#252525')
            notebook.add(tamanho_frame, text="Tamanho")
            
            # Checkbox para manter proporção
            preserve_frame = tk.Frame(tamanho_frame, bg='#252525')
            preserve_frame.pack(fill=tk.X, padx=5, pady=5)
            
            preserve_var = tk.BooleanVar(value=info.get('preserve_ratio', True))
            preserve_check = tk.Checkbutton(preserve_frame, text="Manter proporção original",
                                         variable=preserve_var, 
                                         bg='#252525', fg='white',
                                         selectcolor='#333333', activebackground='#252525')
            preserve_check.pack(anchor=tk.W, pady=2)
            
            # Original aspect ratio
            orig_width, orig_height = info['size']
            aspect_ratio = orig_width / orig_height if orig_height > 0 else 1
            
            def toggle_preserve():
                info['preserve_ratio'] = preserve_var.get()
                
            preserve_check.config(command=toggle_preserve)
            
            # Dimensões com sliders
            size_frame = tk.Frame(tamanho_frame, bg='#252525')
            size_frame.pack(fill=tk.X, padx=5, pady=5)
            
            tk.Label(size_frame, text="Largura:", bg='#252525', fg='white').pack(anchor=tk.W)
            
            # Frame para conter o slider e a caixa de entrada da largura
            width_control_frame = tk.Frame(size_frame, bg='#252525')
            width_control_frame.pack(fill=tk.X, pady=2)
            
            # Obtém a largura do modelo para definir o intervalo máximo do slider
            max_width = 1000
            if self.model_img:
                max_width = self.model_img.width * 2  # Permite imagens até 2x maiores que o modelo
                
            width_var = tk.IntVar(value=info['size'][0])
            width_slider = ttk.Scale(width_control_frame, from_=10, to=max_width, 
                                   orient=tk.HORIZONTAL, variable=width_var)
            width_slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            
            # Caixa de entrada para inserir o valor diretamente
            width_entry = tk.Entry(width_control_frame, width=5, bg='#333333', fg='white', 
                                  insertbackground='white', justify='center')
            width_entry.pack(side=tk.LEFT, padx=5)
            width_entry.insert(0, str(info['size'][0]))
            
            # Frame para conter o slider e a caixa de entrada da altura
            tk.Label(size_frame, text="Altura:", bg='#252525', fg='white').pack(anchor=tk.W)
            
            height_control_frame = tk.Frame(size_frame, bg='#252525')
            height_control_frame.pack(fill=tk.X, pady=2)
            
            # Obtém a altura do modelo para definir o intervalo máximo do slider
            max_height = 1000
            if self.model_img:
                max_height = self.model_img.height * 2  # Permite imagens até 2x maiores que o modelo
                
            height_var = tk.IntVar(value=info['size'][1])
            height_slider = ttk.Scale(height_control_frame, from_=10, to=max_height, 
                                    orient=tk.HORIZONTAL, variable=height_var)
            height_slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            
            # Caixa de entrada para inserir o valor diretamente
            height_entry = tk.Entry(height_control_frame, width=5, bg='#333333', fg='white', 
                                   insertbackground='white', justify='center')
            height_entry.pack(side=tk.LEFT, padx=5)
            height_entry.insert(0, str(info['size'][1]))
            
            def on_width_change(event=None):
                try:
                    w = width_var.get()
                    if w <= 0:
                        return
                    
                    # Manter proporção se necessário
                    if preserve_var.get():
                        h = int(w / aspect_ratio)
                        height_var.set(h)
                        height_entry.delete(0, tk.END)
                        height_entry.insert(0, str(h))
                        info['size'] = (w, h)
                    else:
                        h = height_var.get()
                        info['size'] = (w, h)
                    
                    width_entry.delete(0, tk.END)
                    width_entry.insert(0, str(w))
                    self.redraw_canvas()
                except:
                    pass
            
            def on_width_entry_change(event=None):
                try:
                    w = int(width_entry.get())
                    if w <= 0:
                        w = 10
                    if w > max_width:
                        w = max_width
                        
                    # Manter proporção se necessário
                    if preserve_var.get():
                        h = int(w / aspect_ratio)
                        height_var.set(h)
                        height_entry.delete(0, tk.END)
                        height_entry.insert(0, str(h))
                        info['size'] = (w, h)
                    else:
                        h = height_var.get()
                        info['size'] = (w, h)
                    
                    width_var.set(w)
                    self.redraw_canvas()
                except ValueError:
                    # Se o valor não for válido, restaura o valor anterior
                    width_entry.delete(0, tk.END)
                    width_entry.insert(0, str(info['size'][0]))
                
            def on_height_change(event=None):
                try:
                    h = height_var.get()
                    if h <= 0:
                        return
                        
                    # Manter proporção se necessário
                    if preserve_var.get():
                        w = int(h * aspect_ratio)
                        width_var.set(w)
                        width_entry.delete(0, tk.END)
                        width_entry.insert(0, str(w))
                        info['size'] = (w, h)
                    else:
                        w = width_var.get()
                        info['size'] = (w, h)
                    
                    height_entry.delete(0, tk.END)
                    height_entry.insert(0, str(h))
                    self.redraw_canvas()
                except:
                    pass
                
            def on_height_entry_change(event=None):
                try:
                    h = int(height_entry.get())
                    if h <= 0:
                        h = 10
                    if h > max_height:
                        h = max_height
                        
                    # Manter proporção se necessário
                    if preserve_var.get():
                        w = int(h * aspect_ratio)
                        width_var.set(w)
                        width_entry.delete(0, tk.END)
                        width_entry.insert(0, str(w))
                        info['size'] = (w, h)
                    else:
                        w = width_var.get()
                        info['size'] = (w, h)
                    
                    height_var.set(h)
                    self.redraw_canvas()
                except ValueError:
                    # Se o valor não for válido, restaura o valor anterior
                    height_entry.delete(0, tk.END)
                    height_entry.insert(0, str(info['size'][1]))
                
            # Atualiza quando arrastar os sliders
            width_slider.bind("<B1-Motion>", on_width_change)
            width_slider.bind("<ButtonRelease-1>", on_width_change)
            height_slider.bind("<B1-Motion>", on_height_change)
            height_slider.bind("<ButtonRelease-1>", on_height_change)
            
            # Atualiza quando mudar o valor nas caixas de entrada
            width_entry.bind("<Return>", on_width_entry_change)
            width_entry.bind("<FocusOut>", on_width_entry_change)
            height_entry.bind("<Return>", on_height_entry_change)
            height_entry.bind("<FocusOut>", on_height_entry_change)
        
        # POSIÇÃO - APENAS UMA ABA COMUM PARA AMBOS OS TIPOS
        # Aba de posição (comum a todos os elementos)
        posicao_frame = tk.Frame(notebook, bg='#252525')
        notebook.add(posicao_frame, text="Posição")
        
        # Posição com sliders
        pos_frame = tk.Frame(posicao_frame, bg='#252525')
        pos_frame.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Label(pos_frame, text="Posição X:", bg='#252525', fg='white').pack(anchor=tk.W)
        
        # Frame para conter o slider e a caixa de entrada X
        x_control_frame = tk.Frame(pos_frame, bg='#252525')
        x_control_frame.pack(fill=tk.X, pady=2)
        
        # Obtém a largura do modelo para definir o intervalo do slider
        max_x = 800
        if self.model_img:
            max_x = self.model_img.width
            
        x_var = tk.IntVar(value=int(info['xy'][0]))
        x_slider = ttk.Scale(x_control_frame, from_=0, to=max_x, 
                           orient=tk.HORIZONTAL, variable=x_var)
        x_slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Caixa de entrada para inserir o valor diretamente
        x_entry = tk.Entry(x_control_frame, width=5, bg='#333333', fg='white', 
                          insertbackground='white', justify='center')
        x_entry.pack(side=tk.LEFT, padx=5)
        x_entry.insert(0, str(int(info['xy'][0])))
        
        tk.Label(pos_frame, text="Posição Y:", bg='#252525', fg='white').pack(anchor=tk.W)
        
        # Frame para conter o slider e a caixa de entrada Y
        y_control_frame = tk.Frame(pos_frame, bg='#252525')
        y_control_frame.pack(fill=tk.X, pady=2)
        
        # Obtém a altura do modelo para definir o intervalo do slider
        max_y = 600
        if self.model_img:
            max_y = self.model_img.height
            
        y_var = tk.IntVar(value=int(info['xy'][1]))
        y_slider = ttk.Scale(y_control_frame, from_=0, to=max_y, 
                           orient=tk.HORIZONTAL, variable=y_var)
        y_slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Caixa de entrada para inserir o valor diretamente
        y_entry = tk.Entry(y_control_frame, width=5, bg='#333333', fg='white', 
                          insertbackground='white', justify='center')
        y_entry.pack(side=tk.LEFT, padx=5)
        y_entry.insert(0, str(int(info['xy'][1])))
        
        def on_x_change(event=None):
            try:
                x = x_var.get()
                info['xy'][0] = x
                x_entry.delete(0, tk.END)
                x_entry.insert(0, str(x))
                self.redraw_canvas()
            except:
                pass
        
        def on_x_entry_change(event=None):
            try:
                x = int(x_entry.get())
                if x < 0:
                    x = 0
                if x > max_x:
                    x = max_x
                
                info['xy'][0] = x
                x_var.set(x)
                self.redraw_canvas()
            except ValueError:
                # Se o valor não for válido, restaura o valor anterior
                x_entry.delete(0, tk.END)
                x_entry.insert(0, str(int(info['xy'][0])))
            
        def on_y_change(event=None):
            try:
                y = y_var.get()
                info['xy'][1] = y
                y_entry.delete(0, tk.END)
                y_entry.insert(0, str(y))
                self.redraw_canvas()
            except:
                pass
        
        def on_y_entry_change(event=None):
            try:
                y = int(y_entry.get())
                if y < 0:
                    y = 0
                if y > max_y:
                    y = max_y
                
                info['xy'][1] = y
                y_var.set(y)
                self.redraw_canvas()
            except ValueError:
                # Se o valor não for válido, restaura o valor anterior
                y_entry.delete(0, tk.END)
                y_entry.insert(0, str(int(info['xy'][1])))
        
        # Atualiza quando arrastar os sliders
        x_slider.bind("<B1-Motion>", on_x_change)
        x_slider.bind("<ButtonRelease-1>", on_x_change)
        y_slider.bind("<B1-Motion>", on_y_change)
        y_slider.bind("<ButtonRelease-1>", on_y_change)
        
        # Atualiza quando mudar o valor nas caixas de entrada
        x_entry.bind("<Return>", on_x_entry_change)
        x_entry.bind("<FocusOut>", on_x_entry_change)
        y_entry.bind("<Return>", on_y_entry_change)
        y_entry.bind("<FocusOut>", on_y_entry_change)
        
        # Restaura dados de arrasto
        if preserve_drag_data:
            self.drag_data = preserve_drag_data

    def load_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV", ".csv")])
        if not path: 
            return
        
        try:
            # Carrega o CSV e dá feedback para o usuário
            self.df = pd.read_csv(path, encoding='utf-8')
            rows, cols = self.df.shape
            
            # Ativa o botão 'Esquecer CSV' se existir
            if hasattr(self, 'forget_csv_btn') and self.forget_csv_btn:
                self.forget_csv_btn.config(state=tk.NORMAL)
            
            # Mostra informações sobre os dados carregados
            info_message = f"CSV carregado com sucesso:\n\n"
            info_message += f"- {rows} registos\n"
            info_message += f"- Colunas disponíveis: {', '.join(self.df.columns)}\n\n"
            info_message += "Use esses nomes de colunas como placeholders no texto."
            
            messagebox.showinfo("CSV Carregado", info_message)
        except Exception as e:
            # Dá feedback mais detalhado em caso de erro
            error_message = f"Erro ao carregar o CSV:\n{str(e)}\n\n"
            error_message += "Verifique se o arquivo está no formato correto.\n"
            error_message += "O formato esperado é CSV com cabeçalho na primeira linha."
            messagebox.showerror("Erro", error_message)

    def save_layout(self):
        path = filedialog.asksaveasfilename(defaultextension=".visuproj",
                                           filetypes=[("Projetos VisuMaker", "*.visuproj")])
        if not path: return
        
        # Mostra janela de progresso
        progress = tk.Toplevel(self)
        progress.title("Salvando Projeto")
        progress.geometry("300x80")
        progress.transient(self)
        
        tk.Label(progress, text="A guardar projeto...").pack(pady=10)
        progress_bar = ttk.Progressbar(progress, mode="indeterminate")
        progress_bar.pack(fill=tk.X, padx=20, pady=10)
        progress_bar.start()
        progress.update()
        progress.update_idletasks()  # Força a atualização da interface
        
        try:
            # Prepara os dados para salvar
            items_copy = {}
            
            # Copia os dados essenciais de cada item
            for item_id, item in self.items.items():
                item_data = item.copy()
                
                # Remove referências que não são serializáveis
                if 'tkimg' in item_data:
                    item_data.pop('tkimg')
                if 'img_cache' in item_data:
                    item_data.pop('img_cache')
                if 'img_cache_zoom' in item_data:
                    item_data.pop('img_cache_zoom')
                
                # Para itens de imagem, armazena os dados da imagem em base64
                if item_data['tipo'] == 'imagem':
                    try:
                        # Carrega a imagem e converte para base64
                        with open(item_data['path'], 'rb') as img_file:
                            img_data = base64.b64encode(img_file.read()).decode('utf-8')
                            
                        # Substitui o caminho pelo nome do arquivo e adiciona dados
                        item_data['filename'] = os.path.basename(item_data['path'])
                        item_data['img_data'] = img_data
                    except Exception as img_error:
                        logging.error(f"Erro ao processar imagem {item_data['path']}: {img_error}")
                
                items_copy[item_id] = item_data
            
            # Inclui o modelo base se existir
            if self.model_img:
                # Salva o modelo em um buffer de memória
                buffer = io.BytesIO()
                self.model_img.save(buffer, format="PNG")
                model_data = base64.b64encode(buffer.getvalue()).decode('utf-8')
            else:
                model_data = None
            
            # Cria o dicionário de dados do projeto
            project_data = {
                'items': items_copy,
                'item_order': self.item_order,
                'layer_names': self.layer_names,
                'model': model_data
            }
            
            # Salva no arquivo JSON
            with open(path, 'w') as f:
                json.dump(project_data, f)
            
            progress.destroy()
            messagebox.showinfo("Projeto", "Projeto guardado com sucesso.")
        except Exception as e:
            progress.destroy()
            messagebox.showerror("Erro", f"Erro ao guardar projeto: {str(e)}")

    def load_layout(self, path=None):
        """Carrega um layout existente
        
        Args:
            path: Caminho opcional para o arquivo .visuproj. Se None, abre diálogo de arquivo.
        """
        # Se não foi fornecido um caminho, abre diálogo
        if path is None:
            path = filedialog.askopenfilename(defaultextension=".visuproj",
                                             filetypes=[("Projetos VisuMaker", "*.visuproj")])
            if not path: return
        
        # Mostra janela de progresso
        progress = tk.Toplevel(self)
        progress.title("Carregando Projeto")
        progress.geometry("300x80")
        progress.transient(self)
        
        tk.Label(progress, text="A carregar projeto...").pack(pady=10)
        progress_bar = ttk.Progressbar(progress, mode="indeterminate")
        progress_bar.pack(fill=tk.X, padx=20, pady=10)
        progress_bar.start()
        progress.update()
        progress.update_idletasks()  # Força a atualização da interface
        
        try:
            with open(path) as f:
                data = json.load(f)
                temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp", "temp_images")
                os.makedirs(temp_dir, exist_ok=True)
                self.clear_all_layers()
                # Carrega o modelo base se existir
                if 'model' in data and data['model']:
                    try:
                        model_bytes = base64.b64decode(data['model'])
                        model_stream = io.BytesIO(model_bytes)
                        self.model_img = Image.open(model_stream).convert("RGBA")
                        w, h = self.model_img.size
                        self.canvas.config(width=w, height=h)
                        self.canvas.config(scrollregion=(0, 0, w, h))
                    except Exception as model_error:
                        logging.error(f"Erro ao carregar modelo: {model_error}")
                # Carrega os itens
                for item_id, item_data in data['items'].items():
                    if item_data['tipo'] == 'imagem' and 'img_data' in item_data:
                        try:
                            temp_path = os.path.join(temp_dir, item_data['filename'])
                            img_bytes = base64.b64decode(item_data['img_data'])
                            with open(temp_path, 'wb') as img_file:
                                img_file.write(img_bytes)
                            item_data['path'] = temp_path
                        except Exception as img_error:
                            logging.error(f"Erro ao restaurar imagem: {img_error}")
                            continue
                    if item_data['tipo'] == 'texto':
                        item_data['bg_color'] = ""
                    self.items[item_id] = item_data
                if 'item_order' in data:
                    self.item_order = data['item_order']
                if 'layer_names' in data:
                    self.layer_names = data['layer_names']
                for item_id in self.items:
                    self.visible_items[item_id] = True
                self.refresh_layers_list()
                self._redraw_canvas_now()

                # --- UI/Editor/Template update logic ---
                # If the layout contains template info, update the UI/editor
                template_path = data.get('template_path')
                use_template = data.get('use_template')
                if template_path is not None and hasattr(self, 'template_path_var'):
                    self.template_path_var.set(template_path)
                if use_template is not None and hasattr(self, 'use_template_var'):
                    self.use_template_var.set(use_template)
                if hasattr(self, 'toggle_template'):
                    self.toggle_template()
                if hasattr(self, 'email_config'):
                    if template_path is not None:
                        self.email_config['template_path'] = template_path
                    if use_template is not None:
                        self.email_config['use_template'] = use_template
                if hasattr(self, 'update_preview'):
                    self.update_preview()

            progress.destroy()
            messagebox.showinfo("Projeto", "Projeto carregado com sucesso.")
        except Exception as e:
            progress.destroy()
            messagebox.showerror("Erro", f"Erro ao carregar projeto: {str(e)}\n\nVerifique se o formato do arquivo está correto.")

    def get_all_placeholders(self, data_dict):
        """
        Processa um dicionário de dados e retorna todos os placeholders disponíveis.
        Esta função prepara os valores para substituição nos textos dos documentos visuais.
        
        Args:
            data_dict: Dicionário com os dados da linha do CSV
            
        Returns:
            Dicionário com placeholders formatados
        """
        result = {}
        
        # Copia os valores originais
        for key, value in data_dict.items():
            # Converte para string se não for
            if not isinstance(value, str):
                value = str(value)
            # Limpa espaços extras
            value = value.strip()
            # Adiciona ao resultado
            result[key] = value
            
        # Adiciona versões formatadas (primeira letra maiúscula, etc.)
        for key, value in list(result.items()):
            if isinstance(value, str) and value:
                # Primeira letra maiúscula
                result[f"{key}_cap"] = value[0].upper() + value[1:] if len(value) > 1 else value.upper()
                # Tudo maiúsculo
                result[f"{key}_upper"] = value.upper()
                # Tudo minúsculo
                result[f"{key}_lower"] = value.lower()
                
        # Adiciona placeholders especiais
        import datetime
        now = datetime.datetime.now()
        result["data_atual"] = now.strftime("%d/%m/%Y")
        result["hora_atual"] = now.strftime("%H:%M:%S")
        result["ano_atual"] = str(now.year)
        
        # Adiciona placeholders globais do config.py
        import config
        global_placeholders = getattr(config, 'GLOBAL_PLACEHOLDERS', {})
        for key, value_or_func in global_placeholders.items():
            if callable(value_or_func):
                # Se for uma função, chama para obter o valor
                try:
                    value = value_or_func()
                except:
                    value = f"Erro ao calcular GLOBAL_{key}"
            else:
                value = value_or_func
                
            # Adiciona o placeholder com prefixo GLOBAL_
            result[f"GLOBAL_{key}"] = value
        
        return result

    def _process_certificate(self, row_data, out_dir, result_queue):
        """Função auxiliar que processa documentos visuais para ambos os métodos de geração"""
        try:
            index, row = row_data
            # Cria o documento visual com transparência
            cert = self.model_img.copy().convert("RGBA")
            draw = ImageDraw.Draw(cert)
            
            # Processa os itens na ordem correta
            for item_id in self.item_order:
                # Pula se não estiver visível
                if item_id not in self.visible_items:
                    continue
                
                info = self.items[item_id]
                if info['tipo'] == 'texto':
                    # Substitui placeholders
                    txt = info['texto']
                    try:
                        txt = txt.format(**self.get_all_placeholders(row.to_dict()))
                    except KeyError as e:
                        # Se falhar, mantém o texto original
                        logging.error(f"Aviso: Placeholder não encontrado: {e}")
                        
                    # Configurações de fonte e estilo
                    try:
                        # Em Windows, usa o diretório de fontes padrão
                        font_family = info['font_family']
                        font_size = info['size']
                        
                        # Obtém os estilos configurados
                        is_bold = info.get('bold', False)
                        is_italic = info.get('italic', False)
                        
                        # Debug das propriedades de estilo
                        #print(f"Item: {item_id}, Texto: {txt[:20]}...")
                        #print(f"Propriedades: Font={font_family}, Size={font_size}, Bold={is_bold}, Italic={is_italic}")
                        #print(f"Posição: {info['xy']}, Largura: {info.get('width', 'N/A')}")
                        #print(f"Alinhamento: {info.get('text_align', 'left')}, Cor: {info['color']}")
                        
                        # Usa nossa função especializada para encontrar o arquivo de fonte
                        font_file = self.get_pil_font_file(font_family, is_bold, is_italic)
                        
                        if font_file:
                            try:
                                # Verificar se é uma fonte OpenType (.otf)
                                is_otf = font_file.lower().endswith('.otf')
                                
                                # Log detalhado sobre o arquivo da fonte escolhido
                                #print(f"Tentando carregar fonte: {font_file}")
                        
                                # Tenta carregar a fonte de acordo com o tipo
                                if is_otf:
                                    #print(f"Carregando fonte OpenType: {font_file}")
                                    try:
                                        # Tenta carregar como OpenType
                                        font = ImageFont.truetype(font_file, font_size)
                                        #print(f"Fonte OpenType carregada com sucesso: {font_file}")
                                    except Exception as otf_error:
                                        #print(f"Erro ao carregar fonte OpenType: {otf_error}")
                                        # Tenta abordagem alternativa
                                        from PIL import FreeTypeFont
                                        if hasattr(ImageFont, 'FreeTypeFont'):
                                            font = ImageFont.FreeTypeFont(font_file, font_size)
                                            #print(f"Fonte carregada com FreeTypeFont: {font_file}")
                                        else:
                                            # Fallback para o método padrão
                                            font = ImageFont.truetype(font_file, font_size)
                                else:
                                    # Para fontes TrueType e outras
                                    font = ImageFont.truetype(font_file, font_size)
                                    #print(f"Fonte carregada com sucesso: {font_file}")
                            except Exception as e:
                                #print(f"Erro ao carregar fonte {font_file}: {e}")
                                font = None
                        else:
                            #print(f"Arquivo de fonte não encontrado para {font_family}")
                            font = None
                        
                                                # Se não conseguiu carregar a fonte, tenta fallbacks específicos 
                        # baseados no estilo requisitado
                        if font is None:
                            #print("Tentando fallbacks para fonte...")
                            # Lista de fontes fallback para tentar
                            if is_bold and is_italic:
                                fallbacks = ['arialbi', 'timesbi', 'calibriz', 'verdanaz', 'tahomaZ']
                            elif is_bold:
                                fallbacks = ['arialbd', 'timesbd', 'calibrib', 'verdanab', 'tahomabd']
                            elif is_italic:
                                fallbacks = ['ariali', 'timesi', 'calibrii', 'verdanai', 'tahomai']
                            else:
                                fallbacks = ['arial', 'calibri', 'segoeui', 'tahoma', 'verdana']
    
                            for fallback in fallbacks:
                                try:
                                    # Tenta obter o arquivo de fonte apropriado sem usar estilos
                                    # (o nome já inclui o estilo)
                                    fallback_file = self.get_pil_font_file(fallback, False, False)
                                    if fallback_file:
                                        font = ImageFont.truetype(fallback_file, font_size)
                                        if font:
                                            #print(f"Usando fonte fallback: {fallback}")
                                            break
                                except Exception as e:
                                    #print(f"Erro com fonte fallback {fallback}: {e}")
                                    continue

                        # Se tudo falhar, usa a fonte padrão do sistema
                        if font is None:
                            #print("Usando fonte padrão do sistema.")
                            font = ImageFont.load_default()

                        # Obter posição e outras propriedades necessárias para o texto
                        x, y = info['xy']
                        color = info.get('color', '#000000')
                        text_align = info.get('text_align', 'left')
                        width = info.get('width', 400)
                        
                        # Quebra o texto para caber na largura definida
                        wrapped_text = ""
                        lines = []
                        for line in txt.split('\n'):
                            current_line = ""
                            for word in line.split():
                                test_line = current_line + " " + word if current_line else word
                                try:
                                    # Verifica se cabe na largura
                                    # Use getsize() ou getbbox() em vez de getlength() para compatibilidade
                                    # Alguns tipos de fonte retornam medidas diferentes com getlength()
                                    try:
                                        line_width, _ = font.getsize(test_line)
                                    except AttributeError:
                                        # Para versões mais recentes do PIL que não têm getsize
                                        bbox = font.getbbox(test_line)
                                        line_width = bbox[2] - bbox[0]
                                
                                    if line_width <= width:
                                        current_line = test_line
                                    else:
                                        lines.append(current_line)
                                        current_line = word
                                except:
                                    # Fallback se não conseguir obter a largura
                                    if len(test_line) <= 50:  # Estimativa básica
                                        current_line = test_line
                                    else:
                                        lines.append(current_line)
                                        current_line = word
                            if current_line:
                                lines.append(current_line)
                    
                        wrapped_text = '\n'.join(lines)
                    
                        # Ajusta o alinhamento baseado na configuração
                        if text_align == 'center':
                            # Centraliza cada linha
                            y_pos = y
                            for line in wrapped_text.split('\n'):
                                try:
                                    # Tenta obter a largura da linha usando getsize() ou getbbox()
                                    try:
                                        line_width, _ = font.getsize(line)
                                    except AttributeError:
                                        # Para versões mais recentes do PIL
                                        bbox = font.getbbox(line)
                                        line_width = bbox[2] - bbox[0]
                                        
                                    line_x = x + (width - line_width) // 2
                                except:
                                    # Fallback se não conseguir obter a largura
                                    line_width = len(line) * (font_size * 0.6)  # Estimativa
                                    line_x = x + (width - line_width) // 2
                                
                                # Desenha o texto centralizado
                                draw.text((line_x, y_pos), line, font=font, fill=color)
                                y_pos += int(font_size * 1.2)  # Espaçamento entre linhas
                        elif text_align == 'right':
                            # Alinha à direita
                            y_pos = y
                            for line in wrapped_text.split('\n'):
                                try:
                                    # Tenta obter a largura da linha
                                    try:
                                        line_width, _ = font.getsize(line)
                                    except AttributeError:
                                        # Para versões mais recentes do PIL
                                        bbox = font.getbbox(line)
                                        line_width = bbox[2] - bbox[0]
                                        
                                    line_x = x + width - line_width
                                except:
                                    line_width = len(line) * (font_size * 0.6)  # Estimativa
                                    line_x = x + width - line_width
                                
                                # Desenha o texto alinhado à direita
                                draw.text((line_x, y_pos), line, font=font, fill=color)
                                y_pos += int(font_size * 1.2)
                        else:
                            # Alinhamento padrão à esquerda
                            # Desenha linha por linha para manter o espaçamento consistente
                            y_pos = y
                            for line in wrapped_text.split('\n'):
                                draw.text((x, y_pos), line, font=font, fill=color)
                                y_pos += int(font_size * 1.2)  # Espaçamento entre linhas consistente
                        
                        #print(f"Texto desenhado em ({x}, {y}) com alinhamento {text_align}")
                            
                    except Exception as e:
                        logging.error(f"Erro ao renderizar texto: {e}")
                            
                elif info['tipo'] == 'imagem':
                    try:
                        # Debug da imagem
                        logging.debug(f"Processando imagem: {os.path.basename(info['path'])}")
                        logging.debug(f"Posição: {info['xy']}, Tamanho: {info['size']}, Opacidade: {info['opacity']}")
                        
                        # Abre e redimensiona a imagem
                        im = Image.open(info['path']).convert("RGBA")
                        im = im.resize(info['size'], Image.LANCZOS)
                        
                        # Aplica opacidade corretamente
                        if info['opacity'] < 1.0:
                            # Criar uma cópia da imagem com o canal alpha modificado
                            alpha = int(255 * info['opacity'])
                            
                            # Criar um novo array de pixels para alpha
                            data = im.getdata()
                            new_data = []
                            for item in data:
                                # item é (r, g, b, a)
                                new_a = int(item[3] * (alpha / 255))
                                new_data.append((item[0], item[1], item[2], new_a))
                            
                            # Atualiza a imagem com os novos valores alpha
                            im.putdata(new_data)
                        
                        # Compõe a imagem sobre o certificado com o canal alpha
                        # Assegura que as coordenadas são inteiros
                        x, y = int(info['xy'][0]), int(info['xy'][1])
                        cert.paste(im, (x, y), im)
                        logging.debug(f"Imagem colocada em ({x}, {y}) com tamanho {info['size']}")
                    except Exception as e:
                        # Mostra o erro para ajudar a depurar
                        logging.error(f"Erro ao processar imagem {info.get('path', 'desconhecido')}: {e}")
            
            # Guarda o certificado
            nome = row['nome']
            fname = os.path.join(out_dir, f"{nome}.png")
            
            # Verifica se já existe arquivo com este nome e adiciona número se necessário
            if os.path.exists(fname):
                base_name = nome
                counter = 1
                while os.path.exists(os.path.join(out_dir, f"{base_name}_{counter}.png")):
                    counter += 1
                fname = os.path.join(out_dir, f"{base_name}_{counter}.png")
                logging.info(f"Arquivo {nome}.png já existe. Salvando como {base_name}_{counter}.png")
            
            cert.save(fname)
            logging.info(f"Certificado gerado: {fname}")
            
            # Adiciona resultado na fila (success, caminho)
            result = ("success", fname)
            if result_queue:
                result_queue.put(result)
                
            # Retorna o resultado para poder ser usado na função de processamento
            return result
        except Exception as e:
            # Em caso de erro, também notifica
            logging.error(f"Erro ao gerar certificado: {e}")
            error_result = ("error", str(e))
            if result_queue:
                result_queue.put(error_result)
            return error_result

    def configure_email(self):
        """Abre a janela de configuração de email"""
        # Verifica se um CSV foi carregado
        if self.df is None:
            messagebox.showwarning("Aviso", "Carregue primeiro um CSV antes de configurar o email.")
            return
            
        # Verifica se temos uma coluna "email" para envio
        if 'email' not in self.df.columns:
            messagebox.showwarning("Aviso", "O CSV precisa ter uma coluna 'email'.")
            return
            
        # Abre a janela de configuração
        if not hasattr(self, 'email_config_window') or not self.email_config_window.winfo_exists():
            self.email_config_window = EmailConfigWindow(self)
        else:
            self.email_config_window.lift()
            self.email_config_window.focus_force()
            
    def generate_all(self, from_config=False):
        """Gera certificados e envia por email"""
        if self.df is None:
            messagebox.showwarning("Aviso", "Carregue primeiro um CSV.")
            return
            
        # Verifica se temos uma coluna "email" para envio
        if 'email' not in self.df.columns:
            messagebox.showwarning("Aviso", "O CSV precisa ter uma coluna 'email'.")
            return
        
        # Se não estamos vindo da janela de configuração
        if not from_config:
            if not hasattr(self, 'email_config_window') or not self.email_config_window.winfo_exists():
                self.email_config_window = EmailConfigWindow(self)
                self.email_config_window.send_mode = True
                # O botão já está visível na nova estrutura de layout
                return
            else:
                self.email_config_window.lift()
                self.email_config_window.focus_force()
                return
        
        # Verifica se credenciais SMTP estão configuradas
        if not hasattr(config, 'SMTP_USER') or not hasattr(config, 'SMTP_PASSWORD'):
            messagebox.showerror("Erro", "Configure as credenciais SMTP no ficheiro config.py")
            return
            
        out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "gerados")
        os.makedirs(out_dir, exist_ok=True)
        
        # Barra de progresso
        progress = tk.Toplevel(self)
        progress.title("Gerando e Enviando")
        progress.geometry("300x130")
        progress.transient(self)
        
        tk.Label(progress, text="Processando certificados...").pack(pady=5)
        
        progress_bar = ttk.Progressbar(progress, orient="horizontal", length=250, mode="determinate")
        progress_bar.pack(pady=5)
        
        status_label = tk.Label(progress, text="Iniciando...", wraplength=280)
        status_label.pack(pady=5, fill=tk.X, expand=True)
        
        progress.update()
        
        total = len(self.df)
        
        # Fila para comunicação entre threads
        result_queue = queue.Queue()
        email_queue = queue.Queue()
        
        # Função para processar certificados em threads
        def process_job(row_data):
            index, row = row_data
            result = self._process_certificate(row_data, out_dir, result_queue)
            
            # Sempre coloca o resultado na fila de emails 
            try:
                status, cert_path = result
                if status == "success" and os.path.exists(cert_path):
                    # Verifica se as colunas necessárias existem
                    try:
                        if 'email' not in row.index or 'nome' not in row.index:
                            logging.error("Colunas 'email' ou 'nome' não encontradas no CSV")
                            # coloca erro no resultado para manter contagem
                            result_queue.put(("error", "colunas_missing"))
                            return ("error", "colunas_missing")

                        recipient = str(row['email']).strip()
                        name = str(row['nome']).strip()
                    except Exception as e:
                        logging.error(f"Erro ao extrair colunas do CSV: {e}")
                        result_queue.put(("error", str(e)))
                        return ("error", str(e))

                    email_data = {
                        "recipient": recipient,
                        "name": name,
                        "cert_path": cert_path,
                        "row_data": row.to_dict()
                    }
                    email_queue.put(email_data)
            except KeyError as e:
                logging.error(f"Erro ao acessar coluna no CSV: {e}")
                return "error", None
            except Exception as e:
                logging.error(f"Erro ao processar email data: {e}")
                return "error", None
            except Exception as e:
                logging.error(f"Erro ao preparar email para {row.get('nome', 'desconhecido')}: {e}")
        
        # Cria pool de threads para geração dos certificados
        with ThreadPoolExecutor(max_workers=4) as executor:
            # Envia os trabalhos
            futures = [executor.submit(process_job, (i, row)) for i, row in self.df.iterrows()]
            
            # Atualiza progresso em tempo real
            cert_done = 0
            email_done = 0
            total_tasks = len(futures)
            # Agora sempre assume que haverá envio de emails
            progress_bar["maximum"] = total_tasks * 2
            
            def update_progress():
                nonlocal cert_done, email_done
                try:
                    # Processa resultados da geração de certificados
                    while not result_queue.empty():
                        try:
                            result = result_queue.get_nowait()
                            # Incrementa contador apenas se for bem sucedido
                            cert_done += 1
                            progress_bar["value"] = cert_done + email_done
                            progress.update_idletasks()  # Força a atualização da interface
                            
                            # Sempre mostra contador de emails também
                            status_label.config(text=f"Gerando: {cert_done}/{total_tasks} | Emails: {email_done}/{total_tasks}")
                            progress.update_idletasks()  # Força a atualização do texto
                        except queue.Empty:
                            break
                    
                    # Processa a fila de emails
                    while not email_queue.empty() and email_done < cert_done:
                        email_data = None
                        try:
                            email_data = email_queue.get_nowait()
                            if not email_data:
                                email_done += 1
                                continue
                                
                            # Atualiza o status
                            status_label.config(text=f"Enviando email para {email_data.get('name', '?')}...")
                            progress.update_idletasks()
                            
                            # Envia o email (uso seguro com retries)
                            result = self.send_email_safe(email_data)
                            
                            # Incrementa contador de emails
                            email_done += 1
                            progress_bar["value"] = cert_done + email_done
                            progress.update_idletasks()
                            progress.title(f"Gerando/Enviando - {email_done}/{total_tasks}")
                            status_label.config(text=f"Gerando: {cert_done}/{total_tasks} | Emails: {email_done}/{total_tasks}")
                            progress.update_idletasks()
                        except queue.Empty:
                            break
                        except Exception as e:
                            logging.error(f"Erro ao processar email na fila: {e}")
                            email_done += 1
                            name = email_data.get('name', 'desconhecido') if email_data else 'desconhecido'
                            status_label.config(text=f"Erro ao enviar para {name}: {str(e)[:50]}...")
                            progress.update_idletasks()
                    
                    # Verifica se concluímos tudo
                    all_done = cert_done >= total_tasks and email_done >= total_tasks
                    
                    if all_done:
                        # Mostra a mensagem de sucesso
                        messagebox.showinfo("Concluído", 
                                        f"Certificados gerados e enviados com sucesso.\n"
                                        f"Total: {cert_done} certificados gerados, {email_done} emails enviados.")
                        try:
                            progress.destroy()
                        except:
                            pass
                        return
                    
                    # Agenda próxima atualização
                    try:
                        progress.after(100, update_progress)
                    except:
                        return
                
                except Exception as e:
                    logging.error(f"Erro na atualização de progresso: {e}", exc_info=True)
                    try:
                        status_label.config(text=f"Erro: {str(e)[:50]}")
                        progress.update_idletasks()
                    except:
                        pass
            
            # Inicia monitoramento de progresso
            progress.after(100, update_progress)

    def generate_images_only(self):
        """Gera apenas as imagens sem enviar por email"""
        if self.df is None:
            messagebox.showwarning("Aviso", "Carregue primeiro um CSV.")
            return
        
        # Limpa o placeholder global do assunto (não estamos enviando emails)
        config.GLOBAL_PLACEHOLDERS['subject'] = None
        
        out_dir = "gerados"
        os.makedirs(out_dir, exist_ok=True)

        # Barra de progresso
        progress = tk.Toplevel(self)
        progress.title("Gerando certificados")
        progress.geometry("300x100")
        progress.transient(self)
        
        tk.Label(progress, text="Processando certificados...").pack(pady=10)
        
        progress_bar = ttk.Progressbar(progress, orient="horizontal", length=250, mode="determinate")
        progress_bar.pack(pady=10)
        
        progress.update()

        total = len(self.df)
        
        # Fila para comunicação entre threads
        result_queue = queue.Queue()
        
        # Função para processar certificados em threads
        def process_job(row_data):
            self._process_certificate(row_data, out_dir, result_queue)
            
        # Cria pool de threads
        with ThreadPoolExecutor(max_workers=4) as executor:
            # Envia os trabalhos
            futures = [executor.submit(process_job, (i, row)) for i, row in self.df.iterrows()]
            
            # Atualiza progresso em tempo real
            done = 0
            total_tasks = len(futures)
            progress_bar["maximum"] = total_tasks
            
            def update_progress():
                nonlocal done
                if not result_queue.empty():
                    try:
                        result = result_queue.get_nowait()
                        # Incrementa contador apenas se for bem sucedido
                        done += 1
                        progress_bar["value"] = done
                        progress.update_idletasks()  # Força a atualização da barra
                        progress.title(f"Gerando - {done}/{total_tasks}")
                        progress.update_idletasks()  # Força a atualização do título
                        
                        # Se tivermos terminado, habilita o botão
                        if done >= total_tasks:
                            messagebox.showinfo("Concluído", f"{done} certificados gerados com sucesso em {out_dir}")
                            progress.destroy()
                            
                            # Abre a pasta de saída no explorador de arquivos
                            abs_out_dir = os.path.abspath(out_dir)
                            try:
                                if os.path.exists(abs_out_dir):
                                    subprocess.Popen(f'explorer "{abs_out_dir}"')
                            except Exception as e:
                                logging.error(f"Erro ao abrir pasta: {e}")
                                
                            return
                    except queue.Empty:
                        pass
                
                # Agenda próxima atualização
                progress.after(100, update_progress)
            
            # Inicia monitoramento de progresso
            progress.after(100, update_progress)

    # Métodos de zoom e pan
    def zoom_reset(self):
        self.zoom_factor = 1.0
        # Atualiza o valor do slider
        self.zoom_var.set(1.0)
        zoom_percent = int(self.zoom_factor * 100)
        self.zoom_label.config(text=f"{zoom_percent}%")
        self.schedule_redraw(immediate=True)

    def center_view(self):
        if not self.model_img:
            return
            
        # Centraliza a visualização no canvas
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)
        self.redraw_canvas()

    def zoom_in(self):
        """Aumenta o zoom em 20%"""
        self.zoom_factor *= 1.2
        self.apply_zoom()
        
    def zoom_out(self):
        """Diminui o zoom em 20%"""
        self.zoom_factor /= 1.2
        self.apply_zoom()
        
    def mouse_zoom(self, event):
        """Processa o zoom com a roda do rato diretamente sobre o canvas"""
        if not self.model_img:
            return
            
        # Obtém posição do cursor antes do zoom para manter o ponto de zoom
        canvas_x = self.canvas.canvasx(event.x)
        canvas_y = self.canvas.canvasy(event.y)
            
        # Identifica a direção do scroll
        old_zoom = self.zoom_factor
        if event.num == 5 or event.delta < 0:
            # Scroll down - zoom out
            self.zoom_factor /= 1.1
        else:
            # Scroll up - zoom in
            self.zoom_factor *= 1.1
            
        # Limita o zoom entre 10% e 500%
        self.zoom_factor = max(0.1, min(5.0, self.zoom_factor))
        
        # Atualiza o valor do slider de zoom
        self.zoom_var.set(self.zoom_factor)
        
        # Atualiza o label de zoom
        zoom_percent = int(self.zoom_factor * 100)
        self.zoom_label.config(text=f"{zoom_percent}%")
        
        # Calcula ajuste de scroll para manter o ponto sob o cursor
        if old_zoom != self.zoom_factor:
            # Calcula nova posição proporcionalmente ao zoom
            relative_x = canvas_x / (self.model_img.width * old_zoom)
            relative_y = canvas_y / (self.model_img.height * old_zoom)
            
            # Redraw com o novo zoom - força redraw imediato para feedback visual instantâneo
            self._redraw_canvas_now()  # Desenha imediatamente sem debounce
            
            # Após o redesenho, centra no mesmo ponto relativo
            new_width = self.model_img.width * self.zoom_factor
            new_height = self.model_img.height * self.zoom_factor
            
            # Posiciona o scroll para manter o ponto sob o cursor
            self.canvas.xview_moveto(relative_x - (event.x / new_width))
            self.canvas.yview_moveto(relative_y - (event.y / new_height))

    def apply_zoom(self):
        """Aplica o nível de zoom atual à visualização"""
        # Limita o zoom entre 10% e 500%
        self.zoom_factor = max(0.1, min(5.0, self.zoom_factor))
        
        # Atualiza o label de zoom
        zoom_percent = int(self.zoom_factor * 100)
        self.zoom_label.config(text=f"{zoom_percent}%")
        
        # Redraw com o novo zoom - força redraw imediato para feedback instantâneo
        self._redraw_canvas_now()

    def start_pan(self, event):
        """Inicia o processo de pan/movimentação do canvas com botão direito do mouse"""
        # Guarda a posição inicial para pan/movimentação
        self.pan_start_x = event.x
        self.pan_start_y = event.y
        
        # Obtém as coordenadas de scroll atuais
        self.pan_scroll_x = self.canvas.canvasx(0)
        self.pan_scroll_y = self.canvas.canvasy(0)
        
        # Muda o cursor para indicar que estamos movendo
        self.canvas.config(cursor="fleur")
        
        # Previne conflitos com outras operações
        self.pan_active = True

    def pan_canvas(self, event):
        """Movimenta o canvas em resposta ao arrasto com botão direito do mouse"""
        if not self.model_img or not hasattr(self, 'pan_active') or not self.pan_active:
            return
            
        # Calcula quanto mover - usamos uma abordagem mais direta agora
        dx = self.pan_start_x - event.x
        dy = self.pan_start_y - event.y
        
        # Calcula nova posição de scroll absoluta
        new_x = self.pan_scroll_x + dx
        new_y = self.pan_scroll_y + dy
        
        # Define diretamente a posição de scroll (mais eficiente que scroll incremental)
        self.canvas.xview_moveto(new_x / (self.model_img.width * self.zoom_factor))
        self.canvas.yview_moveto(new_y / (self.model_img.height * self.zoom_factor))
        
        # Não atualizamos pan_start_x/y para manter uma referência consistente
        # durante toda a operação de pan, para movimento mais suave

    def end_pan(self, event):
        """Finaliza a operação de pan quando o botão direito é solto"""
        # Restaura o cursor normal quando soltar o botão
        self.canvas.config(cursor="")
        
        # Marca o pan como inativo
        self.pan_active = False

    # Manipulação da seleção múltipla
    def on_shift_press(self, event):
        # Com Shift, adiciona à seleção múltipla
        items = self.canvas.find_withtag("current")
        if items:
            tags = self.canvas.gettags(items[0])
            item_id = None
            
            # Procura por um ID válido nas tags
            for tag in tags:
                if tag in self.items:
                    item_id = tag
                    break
                
            if item_id and item_id in self.items:
                # Adiciona/remove da seleção múltipla
                if item_id in self.selected_items:
                    self.selected_items.remove(item_id)
                else:
                    self.selected_items.add(item_id)
                    self.current_item = item_id  # Define como atual mas mantém os outros selecionados
                
                # Atualiza a interface
                self.schedule_redraw()
                self.show_properties(item_id)

    def refresh_layers_list(self):
        """Atualiza a lista de camadas na interface"""
        # Guarda os dados de arrasto atuais
        preserve_drag_data = self.drag_data.copy() if hasattr(self, 'drag_data') else None
        
        # Limpa a lista de camadas anterior
        for widget in self.layers_list.winfo_children():
            widget.destroy()
        
        # Remove referências antigas
        self.layer_widgets = {}
        
        # Se não há camadas, mostra apenas uma mensagem
        if not self.items:
            tk.Label(self.layers_list, text="Sem camadas", bg='#252525', fg='white').pack(pady=10)
            return
        
        # Adiciona na ordem (do fundo para o topo)
        for item_id in self.item_order:
            # Cria frame para a camada
            layer_frame = tk.Frame(self.layers_list, bg='#333333', bd=1, relief=tk.FLAT)
            layer_frame.pack(fill=tk.X, pady=1)
            
            # Estado de seleção
            is_selected = (item_id == self.current_item)
            if is_selected:
                layer_frame.config(bg='#4a90d9')
            
            # Botão de visibilidade
            is_visible = item_id in self.visible_items
            vis_img = "👁" if is_visible else "⊘"
            vis_btn = tk.Button(layer_frame, text=vis_img, bg='#333333', fg='white',
                             relief=tk.FLAT, borderwidth=0, padx=2, pady=0,
                             command=lambda id=item_id: self.toggle_layer_visibility(id))
            vis_btn.pack(side=tk.LEFT, padx=2)
            
            # Nome da camada
            name_var = tk.StringVar(value=self.layer_names[item_id])
            lbl = tk.Label(layer_frame, textvariable=name_var, bg='#333333' if not is_selected else '#4a90d9',
                         fg='white', anchor=tk.W, padx=2)
            lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            # Bind de eventos
            layer_frame.bind("<Button-1>", lambda e, id=item_id: self.select_layer(id))
            lbl.bind("<Button-1>", lambda e, id=item_id: self.select_layer(id))
            
            # Guarda referências para actualização futura
            self.layer_widgets[item_id] = {
                'frame': layer_frame,
                'vis_btn': vis_btn,
                'label': lbl,
                'name_var': name_var
            }
            
        # Força atualização da scrollbar da sidebar após adicionar/remover camadas
        self.sidebar_interior.update_idletasks()
        self._configure_sidebar_interior(None)
        
        # Restaura os dados de arrasto
        if preserve_drag_data:
            self.drag_data = preserve_drag_data
            
    def toggle_layer_visibility(self, item_id=None):
        """Alterna a visibilidade de uma camada específica"""
        if item_id is None:
            selection = self.current_item
            if not selection:
                return
            item_id = selection
            
        if item_id in self.visible_items:
            self.visible_items.pop(item_id)
            visible = False
        else:
            self.visible_items[item_id] = True
            visible = True
            
        # Atualiza o botão de visibilidade
        if item_id in self.layer_widgets:
            vis_text = "👁" if visible else "⊘"
            self.layer_widgets[item_id]['vis_btn'].config(text=vis_text)
            
        self.schedule_redraw()
        
    def align_selected(self, align_type):
        """Alinha os itens selecionados com base no tipo de alinhamento especificado"""
        if not self.current_item or not self.model_img:
            return
            
        item_id = self.current_item
        if item_id not in self.items:
            return
            
        info = self.items[item_id]
        canvas_width = self.model_img.width
        canvas_height = self.model_img.height
        
        # Pega as coordenadas e dimensões do item
        x, y = info['xy']
        
        # Determina dimensões do item baseado no tipo
        if info['tipo'] == 'texto':
            # Usa a largura definida pelo usuário
            width = info.get('width', 200)
            height = info.get('height', 100)
            
            # Considera o alinhamento do texto ao centralizar
            text_align = info.get('text_align', 'left')
            
            # Alinha horizontalmente
            if align_type == "left":
                # Alinhamento à esquerda é sempre igual
                info['xy'][0] = 0
            elif align_type == "centerx" or align_type == "center":
                # Centraliza no canvas
                info['xy'][0] = (canvas_width - width) / 2
            elif align_type == "right":
                # Alinhamento à direita
                info['xy'][0] = canvas_width - width
        else:  # Imagem
            width, height = info['size']
            
            # Alinha horizontalmente
            if align_type == "left":
                info['xy'][0] = 0
            elif align_type == "centerx" or align_type == "center":
                info['xy'][0] = (canvas_width - width) / 2
            elif align_type == "right":
                info['xy'][0] = canvas_width - width
        
        # Alinha verticalmente (igual para todos os tipos)
        if align_type == "top":
            info['xy'][1] = 0
        elif align_type == "centery" or align_type == "center":
            info['xy'][1] = (canvas_height - height) / 2
        elif align_type == "bottom":
            info['xy'][1] = canvas_height - height
        
        # Redesenha o canvas
        self.redraw_canvas()
        
        # Atualiza os controles de propriedade para refletir as novas coordenadas
        self.update_property_controls(item_id)

    def update_property_controls(self, item_id):
        """Atualiza os controles de propriedade com valores atuais do item"""
        if item_id not in self.items:
            return
            
        # Verifica qual aba está selecionada atualmente
        current_tab_index = None
        notebook = None
        
        # Procura por um notebook no painel de propriedades
        for widget in self.props_frame.winfo_children():
            if isinstance(widget, ttk.Notebook):
                notebook = widget
                current_tab_index = notebook.index(notebook.select())
                break
        
        # Reinicia as propriedades com os valores atuais
        self.show_properties(item_id)
        
        # Se tínhamos um notebook e uma aba selecionada, restaura a seleção
        if notebook and current_tab_index is not None:
            # Procura pelo novo notebook após recriar os controles
            for widget in self.props_frame.winfo_children():
                if isinstance(widget, ttk.Notebook):
                    # Verifica se o índice é válido para o novo notebook
                    if current_tab_index < len(widget.tabs()):
                        widget.select(current_tab_index)
                    break

    def on_zoom_change(self, value):
        """Atualiza o zoom quando o slider é alterado"""
        try:
            self.zoom_factor = float(value)
            zoom_percent = int(self.zoom_factor * 100)
            self.zoom_label.config(text=f"{zoom_percent}%")
            # Atualiza imediatamente para resposta fluida
            self._redraw_canvas_now()
        except ValueError:
            pass

    # Métodos para gerir a barra de deslocamento da barra lateral
    def _configure_sidebar_interior(self, event):
        """Atualiza a região de deslocamento para acomodar todo o conteúdo"""
        # Ajusta a região de deslocamento para garantir que não haja espaço em branco no topo
        width = self.sidebar_interior.winfo_reqwidth()
        height = self.sidebar_interior.winfo_reqheight()
        
        # Configura a região de scrolling estritamente baseada no tamanho do conteúdo
        self.sidebar_canvas.config(scrollregion="0 0 %s %s" % (width, height))
        
        # Força a posição inicial na parte superior
        self.sidebar_canvas.yview_moveto(0)
        
        # Ajusta a largura da janela interior para corresponder à largura do canvas
        canvas_width = self.sidebar_canvas.winfo_width()
        if canvas_width > 0:
            self.sidebar_canvas.itemconfigure(self.sidebar_interior_id, width=canvas_width)

    def _configure_sidebar_canvas(self, event):
        """Ajusta a largura da janela interior quando o canvas é redimensionado"""
        if self.sidebar_interior.winfo_reqwidth() != event.width:
            self.sidebar_canvas.itemconfigure(self.sidebar_interior_id, width=event.width)
            
        # Garante que a janela interior está posicionada no topo
        self.sidebar_canvas.coords(self.sidebar_interior_id, 0, 0)

    def _on_sidebar_mousewheel(self, event):
        """Processa eventos da roda do rato na barra lateral"""
        # Verifica se o rato está sobre a barra lateral
        try:
            x, y = self.winfo_pointerxy()
            target = self.winfo_containing(x, y)
        except (KeyError, AttributeError):
            # Pode falhar se o widget foi destruído ou o widget tree foi alterado
            return
        
        if target and (target == self.sidebar_canvas or target.master == self.sidebar_interior):
            # Identifica a direção do scroll
            if event.num == 4 or event.delta > 0:
                self.sidebar_canvas.yview_scroll(-1, "units")
            elif event.num == 5 or event.delta < 0:
                self.sidebar_canvas.yview_scroll(1, "units")

    def generate_test_certificate(self):
        """Gera um documento visual de teste e abre-o automaticamente"""
        if not self.model_img:
            messagebox.showwarning("Aviso", "Carregue primeiro um modelo.")
            return
        
        # Limpa o placeholder global do assunto (não estamos enviando emails)
        config.GLOBAL_PLACEHOLDERS['subject'] = None
        
        # Criar diretório temporário se não existir
        out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp")
        os.makedirs(out_dir, exist_ok=True)
        
        # Se houver CSV carregado, usar a primeira linha do CSV
        if hasattr(self, 'df') and self.df is not None and len(self.df) > 0:
            test_row = self.df.iloc[0]
            test_data = test_row.to_dict()
            test_nome = test_data.get('nome', 'Teste')
        else:
            # Criar dados de teste
            test_data = {
                'nome': 'João Testes',
                'email': 'teste@exemplo.com',
                'evento': 'Evento Teste',
                'data': '01/01/2025',
                'local': 'Local Teste'
            }
            test_nome = test_data['nome']
        
        # Criar um DataFrame com uma linha de teste
        test_df = pd.DataFrame([test_data])
        
        # Mostrar mensagem de progresso
        progress = tk.Toplevel(self)
        progress.title("Gerando Certificado de Teste")
        progress.geometry("300x80")
        progress.transient(self)
        
        tk.Label(progress, text="A gerar certificado de teste...").pack(pady=10)
        progress_bar = ttk.Progressbar(progress, mode="indeterminate")
        progress_bar.pack(fill=tk.X, padx=20, pady=10)
        progress_bar.start()
        progress.update()
        
        # Fila para resultado
        result_queue = queue.Queue()
        
        # Usar uma thread para não bloquear a interface
        def process_test():
            try:
                # Chamar função de processamento com os dados de teste
                self._process_certificate((0, test_df.iloc[0]), out_dir, result_queue)
            except Exception as e:
                logging.error(f"Erro ao processar documento de teste: {e}")
                result_queue.put(("error", str(e)))
        
        # Inicia thread para processar
        threading.Thread(target=process_test, daemon=True).start()
        
        # Função para verificar o resultado e abrir o certificado
        def check_result():
            if not result_queue.empty():
                try:
                    result = result_queue.get()
                    progress.destroy()
                    
                    # Verifica se o resultado é uma tupla com status e caminho
                    if isinstance(result, tuple) and len(result) == 2:
                        status, data = result
                        
                        # Verifica se o certificado foi gerado, independente do status
                        cert_path = os.path.join(out_dir, f"{test_nome}.png")
                        if os.path.exists(cert_path):
                            # Abre o certificado com o programa padrão do sistema
                            try:
                                if os.name == 'nt':  # Windows
                                    os.startfile(cert_path)
                                elif os.name == 'posix':  # macOS e Linux
                                    import subprocess
                                    subprocess.call(('xdg-open', cert_path))
                                
                                messagebox.showinfo("Teste Concluído", 
                                                  f"Certificado de teste gerado e aberto.\nLocalização: {cert_path}")
                            except Exception as open_error:
                                messagebox.showwarning("Aviso", 
                                                    f"Certificado gerado, mas não foi possível abri-lo automaticamente.\n"
                                                    f"Localização: {cert_path}\n"
                                                    f"Erro: {open_error}")
                        else:
                            messagebox.showerror("Erro", f"Não foi possível gerar o certificado de teste: {data}")
                    else:
                        # Mesmo se o formato for inesperado, verificamos se o arquivo foi gerado
                        cert_path = os.path.join(out_dir, f"{test_nome}.png")
                        if os.path.exists(cert_path):
                            try:
                                if os.name == 'nt':  # Windows
                                    os.startfile(cert_path)
                                elif os.name == 'posix':  # macOS e Linux
                                    import subprocess
                                    subprocess.call(('xdg-open', cert_path))
                                
                                messagebox.showinfo("Teste Concluído", 
                                                 f"Certificado de teste gerado com sucesso.\nLocalização: {cert_path}")
                            except Exception as open_error:
                                messagebox.showwarning("Aviso", 
                                                    f"Certificado gerado, mas não foi possível abri-lo automaticamente.\n"
                                                    f"Localização: {cert_path}")
                        else:
                            messagebox.showerror("Erro", "Resposta inesperada ao gerar certificado de teste.")
                except Exception as e:
                    progress.destroy()
                    
                    # Mesmo com erro, verifica se o arquivo foi gerado
                    cert_path = os.path.join(out_dir, f"{test_nome}.png")
                    if os.path.exists(cert_path):
                        try:
                            if os.name == 'nt':  # Windows
                                os.startfile(cert_path)
                            elif os.name == 'posix':  # macOS e Linux
                                import subprocess
                                subprocess.call(('xdg-open', cert_path))
                            
                            messagebox.showinfo("Teste Concluído", 
                                             f"Certificado de teste gerado apesar de erros.\nLocalização: {cert_path}")
                        except Exception as open_error:
                            messagebox.showwarning("Aviso", 
                                                f"Certificado gerado, mas não foi possível abri-lo automaticamente.\n"
                                                f"Localização: {cert_path}")
                    else:
                        messagebox.showerror("Erro", f"Erro ao abrir certificado: {str(e)}")
            else:
                # Verifica novamente após 100ms
                progress.after(100, check_result)
        
        # Inicia verificação do resultado
        progress.after(100, check_result)

    def load_fonts_map(self):
        """Carrega o mapeamento de fontes do ficheiro ou cria um novo se não existir"""
        try:
            if os.path.exists(self.fonts_map_file):
                with open(self.fonts_map_file, 'r', encoding='utf-8') as f:
                    fonts_map = json.load(f)
                logging.info(f"Fonts mapping loaded successfully: {len(fonts_map)} fonts found.")
                return fonts_map
            else:
                logging.warning("Font mapping file not found. Creating a new one...")
                return self.create_fonts_map()
        except Exception as e:
            logging.error(f"Error loading fonts mapping: {e}")
            return self.create_fonts_map()
    
    def create_fonts_map(self):
        """Cria um mapeamento de fontes e os seus estilos disponíveis"""
        logging.info("Creating fonts mapping...")
        fonts_map = {}
        
        # Obtém a lista de fontes disponíveis usando o tkinter
        font_families = list(tkfont.families())
        
        # Cria um dicionário para armazenar informações sobre cada fonte
        fonts_map = {}
        
        # Adiciona cada fonte ao mapeamento
        for font_family in font_families:
            # Normaliza o nome da fonte para minúsculas para consistência
            font_name_lower = font_family.lower()
            
            # Cria entrada para esta fonte no mapeamento
            fonts_map[font_name_lower] = {
                'styles': {
                    'normal': False,
                    'bold': False,
                    'italic': False,
                    'bold_italic': False
                },
                'files': {
                    'normal': "",
                    'bold': "",
                    'italic': "",
                    'bold_italic': ""
                }
            }
            
            # Verifica quais ficheiros estão disponíveis para cada estilo usando o registo do Windows
            # Normal
            normal_file = self._get_font_file_from_registry(font_family, False, False)
            if normal_file and os.path.exists(normal_file):
                fonts_map[font_name_lower]['styles']['normal'] = True
                fonts_map[font_name_lower]['files']['normal'] = normal_file
            
            # Bold
            bold_file = self._get_font_file_from_registry(font_family, True, False)
            if bold_file and os.path.exists(bold_file):
                fonts_map[font_name_lower]['styles']['bold'] = True
                fonts_map[font_name_lower]['files']['bold'] = bold_file
            
            # Italic
            italic_file = self._get_font_file_from_registry(font_family, False, True)
            if italic_file and os.path.exists(italic_file):
                fonts_map[font_name_lower]['styles']['italic'] = True
                fonts_map[font_name_lower]['files']['italic'] = italic_file
            
            # Bold Italic
            bold_italic_file = self._get_font_file_from_registry(font_family, True, True)
            if bold_italic_file and os.path.exists(bold_italic_file):
                fonts_map[font_name_lower]['styles']['bold_italic'] = True
                fonts_map[font_name_lower]['files']['bold_italic'] = bold_italic_file
        
        # Salva o mapeamento para uso futuro
        self.save_fonts_map(fonts_map)
        
        return fonts_map

    def save_fonts_map(self, fonts_map):
        """Guarda o mapeamento de fontes num ficheiro JSON"""
        try:
            with open(self.fonts_map_file, 'w', encoding='utf-8') as f:
                json.dump(fonts_map, f, ensure_ascii=False, indent=2)
            logging.info(f"Fonts mapping saved successfully: {len(fonts_map)} fonts found.")
        except Exception as e:
            logging.error(f"Error saving fonts mapping: {e}")

    def update_fonts_map(self):
        """Atualiza o mapeamento de fontes com informações detalhadas sobre os arquivos de fonte"""
        try:
            # Mostra uma janela de progresso
            progress = tk.Toplevel(self)
            progress.title("Atualizar Mapeamento de Fontes")
            progress.geometry("300x80")
            progress.transient(self)
            
            tk.Label(progress, text="A mapear fontes do sistema...").pack(pady=10)
            progress_bar = ttk.Progressbar(progress, mode="indeterminate")
            progress_bar.pack(fill=tk.X, padx=20, pady=10)
            progress_bar.start()
            progress.update()
            progress.update_idletasks()  # Força a atualização da interface
            
            # Obtém a lista de fontes disponíveis
            font_families = list(tkfont.families())
            font_families.sort()
            
            # Cria um novo mapeamento
            updated_map = {}
            new_fonts_count = 0
            updated_fonts_count = 0
            otf_fonts_count = 0
            
            # Para cada fonte, atualiza as informações
            for font_family in font_families:
                font_name_lower = font_family.lower()
                
                # Inicializa a entrada para a fonte no mapeamento
                updated_map[font_name_lower] = {
                    'styles': {
                        'normal': False,
                        'bold': False,
                        'italic': False,
                        'bold_italic': False
                    },
                    'files': {}
                }
                
                # Verifica se é uma fonte nova ou atualizada
                if font_name_lower not in self.fonts_map:
                    new_fonts_count += 1
                else:
                    updated_fonts_count += 1
                
                # Obtém os caminhos dos arquivos para cada estilo
                # Normal
                normal_file = self._get_font_file_from_registry(font_family, False, False)
                if normal_file and os.path.exists(normal_file):
                    updated_map[font_name_lower]['styles']['normal'] = True
                    updated_map[font_name_lower]['files']['normal'] = os.path.basename(normal_file)
                    if normal_file.lower().endswith('.otf'):
                        otf_fonts_count += 1
                
                # Bold
                bold_file = self._get_font_file_from_registry(font_family, True, False)
                if bold_file and os.path.exists(bold_file):
                    updated_map[font_name_lower]['styles']['bold'] = True
                    updated_map[font_name_lower]['files']['bold'] = os.path.basename(bold_file)
                    if bold_file.lower().endswith('.otf'):
                        otf_fonts_count += 1
                
                # Italic
                italic_file = self._get_font_file_from_registry(font_family, False, True)
                if italic_file and os.path.exists(italic_file):
                    updated_map[font_name_lower]['styles']['italic'] = True
                    updated_map[font_name_lower]['files']['italic'] = os.path.basename(italic_file)
                    if italic_file.lower().endswith('.otf'):
                        otf_fonts_count += 1
                
                # Bold Italic
                bold_italic_file = self._get_font_file_from_registry(font_family, True, True)
                if bold_italic_file and os.path.exists(bold_italic_file):
                    updated_map[font_name_lower]['styles']['bold_italic'] = True
                    updated_map[font_name_lower]['files']['bold_italic'] = os.path.basename(bold_italic_file)
                    if bold_italic_file.lower().endswith('.otf'):
                        otf_fonts_count += 1
            
            # Adiciona fontes de diretórios específicos que podem não estar no registro
            font_dirs = [
                os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'Fonts'),
                os.path.expandvars('%localappdata%\\Microsoft\\Windows\\Fonts')
            ]
            font_extensions = ['.otf', '.ttf', '.ttc']
            
            # Procura em todos os diretórios
            for font_dir in font_dirs:
                if os.path.exists(font_dir):
                    for file_name in os.listdir(font_dir):
                        file_lower = file_name.lower()
                        
                        # Verifica se é um arquivo de fonte válido
                        if not any(file_lower.endswith(ext) for ext in font_extensions):
                            continue
                        
                        # Se for um arquivo .otf, adiciona ao contador
                        if file_lower.endswith('.otf'):
                            otf_fonts_count += 1
                        
                        # Tenta extrair informações do nome do arquivo
                        # Nomes de arquivos de fonte geralmente seguem padrões como:
                        # FontName-Bold.ttf, FontNameBold.ttf, FontName_Bold.ttf, etc.
                        base_name = os.path.splitext(file_lower)[0]
                        
                        # Determina o estilo com base no nome do arquivo
                        style_key = 'normal'
                        if any(suffix in base_name for suffix in ['bold', 'bd', 'b', 'black']):
                            if any(suffix in base_name for suffix in ['italic', 'it', 'i', 'oblique']):
                                style_key = 'bold_italic'
                            else:
                                style_key = 'bold'
                        elif any(suffix in base_name for suffix in ['italic', 'it', 'i', 'oblique']):
                            style_key = 'italic'
                        
                        # Tenta limpar o nome base para obter o nome da fonte
                        font_base_name = base_name
                        for suffix in ['bold', 'bd', 'b', 'black', 'italic', 'it', 'i', 'oblique', 'regular', 'normal']:
                            # Remove sufixos comuns
                            font_base_name = re.sub(f"[-_]?{suffix}$", "", font_base_name)
                        
                        # Normaliza o nome (remove traços, sublinhados, etc.)
                        font_base_name = font_base_name.replace('-', ' ').replace('_', ' ')
                        
                        # Verifica se o nome é similar a alguma fonte conhecida
                        for font_name in font_families:
                            font_name_lower = font_name.lower()
                            # Usa similaridade de nomes para encontrar correspondências
                            if font_base_name in font_name_lower or font_name_lower in font_base_name:
                                # Se encontrou, adiciona ao mapeamento
                                if font_name_lower not in updated_map:
                                    updated_map[font_name_lower] = {
                                        'styles': {
                                            'normal': False,
                                            'bold': False,
                                            'italic': False,
                                            'bold_italic': False
                                        },
                                        'files': {}
                                    }
                                    new_fonts_count += 1
                                
                                # Adiciona o arquivo ao mapeamento
                                if not updated_map[font_name_lower]['styles'].get(style_key, False):
                                    updated_map[font_name_lower]['styles'][style_key] = True
                                    updated_map[font_name_lower]['files'][style_key] = file_name
                                    
                                    # Se for um .otf, incrementa o contador
                                    if file_lower.endswith('.otf'):
                                        otf_fonts_count += 1
            
            # Atualiza o mapeamento
            self.fonts_map = updated_map
            
            # Conta quantos estilos estão disponíveis
            style_counts = {
                'normal': 0,
                'bold': 0,
                'italic': 0,
                'bold_italic': 0
            }
            
            for font_name, font_data in self.fonts_map.items():
                styles = font_data.get('styles', {})
                for style in style_counts:
                    if styles.get(style, False):
                        style_counts[style] += 1
            
            # Salva o mapeamento
            self.save_fonts_map(self.fonts_map)
            
            # Fecha a janela de progresso
            progress.destroy()
            
            # Mostra mensagem de conclusão
            info_message = f"O mapeamento de fontes foi atualizado com sucesso.\n\n"
            info_message += f"Total de fontes: {len(self.fonts_map)}\n"
            info_message += f"Novas fontes: {new_fonts_count}\n"
            info_message += f"Fontes atualizadas: {updated_fonts_count}\n"
            #info_message += f"Fontes OpenType (.otf): {otf_fonts_count}\n"
            #info_message += f"\nEstilos disponíveis:\n"
            #info_message += f"Normal: {style_counts['normal']}\n"
            #info_message += f"Negrito: {style_counts['bold']}\n"
            #info_message += f"Itálico: {style_counts['italic']}\n"
            #info_message += f"Negrito+Itálico: {style_counts['bold_italic']}\n"
            
            messagebox.showinfo("Fontes Atualizadas", info_message)
            
            return self.fonts_map
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao atualizar mapeamento de fontes: {e}")
            return self.fonts_map
        
    def _get_system_fonts(self):
        """Obtém a lista de fontes instaladas no sistema usando a API do Windows"""
        try:
            installed_fonts = []
            
            # Função para enumerar fontes no Windows usando ctypes
            def enum_font_families():
                # Callback para receber fontes
                def callback(font_family, tm, font_type, param):
                    try:
                        if font_family:
                            # Verifica se font_family é um objeto que pode ser decodificado ou um inteiro
                            if isinstance(font_family, int):
                                # Se for um inteiro, simplesmente ignorar
                                return True
                            
                            try:
                                font_name = font_family.decode('utf-8', errors='ignore')
                                installed_fonts.append(font_name.lower())
                            except (AttributeError, UnicodeDecodeError) as e:
                                # Se não conseguir decodificar, tenta usar como string
                                try:
                                    if isinstance(font_family, str):
                                        installed_fonts.append(font_family.lower())
                                except:
                                    # Ignora se não conseguir processar
                                    pass
                    except Exception as e:
                                # Se não conseguir decodificar, ignora esta fonte
                                logging.error(f"Erro ao processar nome de fonte: {e}")
                    except Exception as e:
                        # Ignora qualquer erro no callback para não interromper o processo
                        print(f"Erro no callback de enumeração de fontes: {e}")
                    return True
                
                # Função mais simples que não tenta enumerar todas as fontes
                try:
                    # Método alternativo: ler do registro do Windows
                    import winreg
                    fonts_reg = winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE)
                    fonts_key = winreg.OpenKey(fonts_reg, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
                    
                    # Lê todas as entradas de fontes
                    i = 0
                    while True:
                        try:
                            name, value, _ = winreg.EnumValue(fonts_key, i)
                            # O nome geralmente contém o estilo (ex: "Arial Bold (TrueType)")
                            # Vamos extrair o nome base da fonte
                            parts = name.split('(')
                            if len(parts) > 0:
                                base_name = parts[0].strip().lower()
                                # Remove informações de estilo do nome
                                for style in ['bold', 'italic', 'oblique', 'light', 'black', 'medium', 'regular']:
                                    base_name = base_name.replace(style, '').strip()
                                # Adiciona à lista se não estiver vazia
                                if base_name:
                                    installed_fonts.append(base_name)
                            i += 1
                        except OSError:
                            # Não há mais fontes para ler
                            break
                    
                    return installed_fonts
                    
                except Exception as reg_error:
                    logging.error(f"Erro ao ler fontes do registro: {reg_error}")
                    # Fallback para a enumeração direta (que pode causar erros)
                    try:
                        # Tipo de função de callback
                        FONTENUMPROC = ctypes.WINFUNCTYPE(
                            wintypes.BOOL,
                            wintypes.LPVOID,
                            wintypes.LPVOID,
                            wintypes.DWORD,
                            wintypes.LPARAM
                        )
                        
                        # Acesso à API do Windows
                        gdi32 = ctypes.WinDLL('gdi32')
                        hdc = ctypes.windll.user32.GetDC(None)
                        
                        # Enumera as famílias de fontes
                        gdi32.EnumFontFamiliesExA(
                            hdc,
                            None,
                            FONTENUMPROC(callback),
                            0,
                            0
                        )
                        
                        ctypes.windll.user32.ReleaseDC(None, hdc)
                    except Exception as enum_error:
                        logging.error(f"Erro na enumeração de fontes: {enum_error}")
                
                # Tipo de função de callback
                FONTENUMPROC = ctypes.WINFUNCTYPE(
                    wintypes.BOOL,
                    wintypes.LPVOID,
                    wintypes.LPVOID,
                    wintypes.DWORD,
                    wintypes.LPARAM
                )
                
                # Acesso à API do Windows
                gdi32 = ctypes.WinDLL('gdi32')
                hdc = ctypes.windll.user32.GetDC(None)
                
                # Enumera as famílias de fontes
                gdi32.EnumFontFamiliesExA(
                    hdc,
                    None,
                    FONTENUMPROC(callback),
                    0,
                    0
                )
                
                ctypes.windll.user32.ReleaseDC(None, hdc)
                
            # Executa a enumeração
            enum_font_families()
            
            # Remove duplicados e ordena
            installed_fonts = list(set(installed_fonts))
            installed_fonts.sort()
            
            return installed_fonts
            
        except Exception as e:
            logging.error(f"Erro ao obter fontes do sistema: {e}")
            return []

    def get_pil_font_file(self, font_name, bold=False, italic=False):
        """
        Encontra o arquivo de fonte correto para o PIL com base no nome da fonte Tkinter.
        Esta função cria uma ponte entre fontes Tkinter e PIL.
        
        Args:
            font_name (str): Nome da fonte como aparece no Tkinter
            bold (bool): Se deve usar estilo negrito
            italic (bool): Se deve usar estilo itálico
            
        Returns:
            str: Caminho do arquivo da fonte ou None se não encontrado
        """
        # Normaliza o nome da fonte para minúsculas
        font_name_lower = font_name.lower()
        file_path = None
        
        # Debug para identificar o que está sendo solicitado
        #print(f"Solicitando fonte: '{font_name}' [bold={bold}, italic={italic}]")
        
        # Determina o estilo desejado
        if bold and italic:
            style_key = 'bold_italic'
        elif bold:
            style_key = 'bold'
        elif italic:
            style_key = 'italic'
        else:
            style_key = 'normal'
        
        # Verifica se a fonte existe no nosso mapeamento
        if font_name_lower in self.fonts_map:
            font_data = self.fonts_map[font_name_lower]
            
            # Verifica se o estilo solicitado está disponível
            if font_data.get('styles', {}).get(style_key, False):
                # Obtém o nome do arquivo
                filename = font_data.get('files', {}).get(style_key, "")
                
                if filename:
                    # Converte o nome do arquivo para caminho completo
                    # Primeiro tenta na pasta de fontes do Windows
                    windows_font_dir = os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'Fonts')
                    full_path = os.path.join(windows_font_dir, filename)
                    
                    if os.path.exists(full_path):
                        logging.debug(f"Arquivo de fonte encontrado: {full_path}")
                        return full_path
                    
                    # Tenta no diretório de fontes do usuário
                    user_font_dir = os.path.expandvars('%localappdata%\\Microsoft\\Windows\\Fonts')
                    user_path = os.path.join(user_font_dir, filename)
                    
                    if os.path.exists(user_path):
                        logging.debug(f"Arquivo de fonte encontrado: {user_path}")
                        return user_path
                    
                    # Se não encontrou, tenta outros diretórios comuns
                    # Para Linux/Mac
                    if os.name == 'posix':
                        for font_dir in ['/usr/share/fonts', '/usr/local/share/fonts', 
                                        '~/.fonts', '~/Library/Fonts', '/Library/Fonts']:
                            expanded_dir = os.path.expanduser(font_dir)
                            # Procura recursivamente
                            for root, dirs, files in os.walk(expanded_dir):
                                if filename in files:
                                    full_path = os.path.join(root, filename)
                                    logging.debug(f"Arquivo de fonte encontrado: {full_path}")
                                    return full_path
            
            # Se o estilo solicitado não estiver disponível, tenta fallbacks
            # Definir prioridades de fallback
            if style_key == 'bold_italic':
                fallbacks = ['bold', 'italic', 'normal']
            elif style_key == 'bold':
                fallbacks = ['bold_italic', 'normal', 'italic']
            elif style_key == 'italic':
                fallbacks = ['bold_italic', 'normal', 'bold']
            else:
                fallbacks = ['bold', 'italic', 'bold_italic']
            
            # Tenta cada fallback
            for fallback_style in fallbacks:
                if font_data.get('styles', {}).get(fallback_style, False):
                    filename = font_data.get('files', {}).get(fallback_style, "")
                    if filename:
                        # Tenta localizar o arquivo
                        windows_font_dir = os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'Fonts')
                        full_path = os.path.join(windows_font_dir, filename)
                        
                        if os.path.exists(full_path):
                            logging.debug(f"Usando fallback {fallback_style} para {style_key}: {full_path}")
                            return full_path
                        
                        # Tenta no diretório de fontes do usuário
                        user_font_dir = os.path.expandvars('%localappdata%\\Microsoft\\Windows\\Fonts')
                        user_path = os.path.join(user_font_dir, filename)
                        
                        if os.path.exists(user_path):
                            logging.debug(f"Usando fallback {fallback_style} para {style_key} (user): {user_path}")
                            return user_path
                        
                        # Tenta outros diretórios comuns
                        if os.name == 'posix':
                            for font_dir in ['/usr/share/fonts', '/usr/local/share/fonts', 
                                            '~/.fonts', '~/Library/Fonts', '/Library/Fonts']:
                                expanded_dir = os.path.expanduser(font_dir)
                                for root, dirs, files in os.walk(expanded_dir):
                                    if filename in files:
                                        full_path = os.path.join(root, filename)
                                        logging.debug(f"Usando fallback {fallback_style} para {style_key}: {full_path}")
                                        return full_path
        
        # Se tudo falhar, tenta obter diretamente do registro
        registry_file = self._get_font_file_from_registry(font_name, bold, italic)
        if registry_file and os.path.exists(registry_file):
            logging.debug(f"Usando fonte do registro: {registry_file}")
            return registry_file
        
        # Tenta procurar arquivos .otf/.ttf diretamente nos diretórios de fontes
        # Baseado no nome, sem necessidade de registro
        font_dirs = [
            os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'Fonts'),
            os.path.expandvars('%localappdata%\\Microsoft\\Windows\\Fonts')
        ]
        
        # Determina o padrão de nome que vamos procurar
        name_patterns = []
        font_base = font_name_lower.replace(' ', '').replace('-', '').replace('_', '')
        name_patterns.append(font_base)
        
        # Adiciona variantes comuns
        if ' ' in font_name_lower:
            name_patterns.append(font_name_lower.replace(' ', ''))
            name_patterns.append(font_name_lower.replace(' ', '_'))
            name_patterns.append(font_name_lower.replace(' ', '-'))
        
        # Adiciona padrões para estilos
        style_suffixes = []
        if bold and italic:
            style_suffixes = ['bolditalic', 'boldoblique', 'blackitalic', 'bi', 'z']
        elif bold:
            style_suffixes = ['bold', 'black', 'bd', 'b']
        elif italic:
            style_suffixes = ['italic', 'oblique', 'it', 'i']
        else:
            style_suffixes = ['regular', 'normal', 'book']
        
        # Extensões de fonte a procurar
        font_extensions = ['.otf', '.ttf', '.ttc']
        
        # Procura todas as combinações possíveis em todos os diretórios
        for font_dir in font_dirs:
            if os.path.exists(font_dir):
                for file_name in os.listdir(font_dir):
                    file_lower = file_name.lower()
                    
                    # Verifica se é um arquivo de fonte válido
                    if not any(file_lower.endswith(ext) for ext in font_extensions):
                        continue
                    
                    # Verifica se o nome contém algum dos padrões
                    for pattern in name_patterns:
                        if pattern in file_lower:
                            # Verifica se o estilo bate
                            if not style_suffixes:  # Se não temos restrição de estilo
                                font_path = os.path.join(font_dir, file_name)
                                logging.debug(f"Encontrado arquivo de fonte por nome: {font_path}")
                                # Atualiza o mapeamento de fontes com esta informação
                                self._update_font_map_entry(font_name_lower, style_key, file_name, font_path)
                                return font_path
                            
                            # Verifica se o arquivo contém algum dos sufixos de estilo
                            for suffix in style_suffixes:
                                if suffix in file_lower:
                                    font_path = os.path.join(font_dir, file_name)
                                    logging.debug(f"Encontrado arquivo de fonte por nome e estilo: {font_path}")
                                    # Atualiza o mapeamento de fontes com esta informação
                                    self._update_font_map_entry(font_name_lower, style_key, file_name, font_path)
                                    return font_path
        
        # Se ainda não encontrou, tenta fontes similares ou genéricas
        similar_font_names = []
        
        # Adiciona variações comuns do nome
        name_parts = font_name_lower.split()
        if len(name_parts) > 1:
            # Tenta apenas a primeira parte do nome (ex: "Arial Black" -> "Arial")
            similar_font_names.append(name_parts[0])
        
        # Adiciona fontes de fallback genéricas
        default_fonts = ['arial', 'times new roman', 'courier new', 'verdana', 'calibri']
        similar_font_names.extend(default_fonts)
        
        # Tenta cada fonte similar
        for similar_name in similar_font_names:
            if similar_name in self.fonts_map:
                similar_font_data = self.fonts_map[similar_name]
                
                # Tenta o mesmo estilo primeiro
                if similar_font_data.get('styles', {}).get(style_key, False):
                    filename = similar_font_data.get('files', {}).get(style_key, "")
                    if filename:
                        windows_font_dir = os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'Fonts')
                        full_path = os.path.join(windows_font_dir, filename)
                        
                        if os.path.exists(full_path):
                            logging.debug(f"Usando fonte similar '{similar_name}' para '{font_name}': {full_path}")
                            return full_path
                        
                        # Tenta no diretório de fontes do usuário
                        user_font_dir = os.path.expandvars('%localappdata%\\Microsoft\\Windows\\Fonts')
                        user_path = os.path.join(user_font_dir, filename)
                        
                        if os.path.exists(user_path):
                            logging.debug(f"Usando fonte similar '{similar_name}' para '{font_name}' (user): {user_path}")
                            return user_path
                
                # Se não encontrou com o mesmo estilo, tenta estilo normal
                if similar_font_data.get('styles', {}).get('normal', False):
                    filename = similar_font_data.get('files', {}).get('normal', "")
                    if filename:
                        windows_font_dir = os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'Fonts')
                        full_path = os.path.join(windows_font_dir, filename)
                        
                        if os.path.exists(full_path):
                            logging.debug(f"Usando fonte similar '{similar_name}' (normal) para '{font_name}': {full_path}")
                            return full_path
                        
                        # Tenta no diretório de fontes do usuário
                        user_font_dir = os.path.expandvars('%localappdata%\\Microsoft\\Windows\\Fonts')
                        user_path = os.path.join(user_font_dir, filename)
                        
                        if os.path.exists(user_path):
                            logging.debug(f"Usando fonte similar '{similar_name}' (normal) para '{font_name}' (user): {user_path}")
                            return user_path
        
        # Última tentativa: fontes padrão do sistema
        if os.name == 'nt':  # Windows
            system_fonts = {
                'normal': os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'Fonts', 'arial.ttf'),
                'bold': os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'Fonts', 'arialbd.ttf'),
                'italic': os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'Fonts', 'ariali.ttf'),
                'bold_italic': os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'Fonts', 'arialbi.ttf')
            }
            
            if os.path.exists(system_fonts[style_key]):
                logging.debug(f"Usando fonte padrão do sistema: {system_fonts[style_key]}")
                return system_fonts[style_key]
        
        # Se chegou aqui, não encontrou nenhuma fonte adequada
        logging.debug(f"Não foi possível encontrar arquivo para a fonte '{font_name}' com estilo {style_key}")
        return None

    def _get_font_file_from_registry(self, font_name, bold=False, italic=False):
        """
        Obtém o ficheiro de fonte diretamente do registo do Windows.
        
        Args:
            font_name (str): Nome da fonte
            bold (bool): Se deve procurar estilo negrito
            italic (bool): Se deve procurar estilo itálico
            
        Returns:
            str: Caminho completo para o ficheiro de fonte ou None se não encontrado
        """
        try:
            import winreg
            import re
            
            # Determina o estilo a procurar no nome da fonte
            style_suffix = ""
            if bold and italic:
                style_suffix = " Bold Italic"
            elif bold:
                style_suffix = " Bold"
            elif italic:
                style_suffix = " Italic"
                
            # Nome completo para procurar no registo
            full_font_name = font_name + style_suffix
            
            # Lista de possíveis candidatos encontrados
            candidates = []
            
            # Lista das chaves de registo a verificar
            registry_keys = [
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
                (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
            ]
            
            # Verifica os registos para cada fonte
            for registry_root, registry_path in registry_keys:
                try:
                    key = winreg.OpenKey(registry_root, registry_path)
                    count = winreg.QueryInfoKey(key)[1]
                    
                    for i in range(count):
                        try:
                            name, value, _ = winreg.EnumValue(key, i)
                            
                            # Verifica se o nome do registro corresponde ao nome da fonte
                            name_matches = False
                            
                            # Verifica correspondência exata
                            if full_font_name.lower() in name.lower():
                                name_matches = True
                            # Verifica correspondência parcial
                            elif font_name.lower() in name.lower():
                                # Se estamos procurando um estilo específico, verifica se o nome contém o estilo
                                if bold and italic and ('bold' in name.lower() and 'italic' in name.lower()):
                                    name_matches = True
                                elif bold and not italic and 'bold' in name.lower() and 'italic' not in name.lower():
                                    name_matches = True
                                elif italic and not bold and 'italic' in name.lower() and 'bold' not in name.lower():
                                    name_matches = True
                                elif not bold and not italic:
                                    # Se estamos procurando estilo normal, verifica se não tem estilo específico
                                    if 'bold' not in name.lower() and 'italic' not in name.lower():
                                        name_matches = True
                            
                            if name_matches:
                                # Processa o valor para obter o caminho do arquivo
                                if os.path.isabs(value):
                                    # Se já é um caminho absoluto
                                    font_path = value
                                else:
                                    # Se é um caminho relativo ao diretório de fontes
                                    if registry_root == winreg.HKEY_LOCAL_MACHINE:
                                        font_dir = os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'Fonts')
                                    else:
                                        # Para fontes instaladas pelo usuário
                                        font_dir = os.path.expandvars('%localappdata%\\Microsoft\\Windows\\Fonts')
                                    
                                    font_path = os.path.join(font_dir, value)
                                
                                # Verifica se o arquivo existe
                                if os.path.exists(font_path):
                                    logging.debug(f"Encontrado pelo registo: {name} -> {font_path}")
                                    
                                    # Dá prioridade a fontes .otf
                                    if font_path.lower().endswith('.otf'):
                                        return font_path
                                    
                                    # Adiciona à lista de candidatos
                                    candidates.append(font_path)
                        except Exception as e:
                            logging.error(f"Erro ao processar entrada de registo {i}: {e}")
                            continue
                    
                    
                    winreg.CloseKey(key)
                except Exception as e:
                    print(f"Erro ao acessar chave de registo {registry_path}: {e}")
            
            # Se não encontrou nenhuma fonte .otf, mas encontrou candidatos, retorna o primeiro
            if candidates:
                return candidates[0]
            
            # Se não encontrou nada no registro, procura diretamente nos diretórios de fontes
            font_dirs = [
                os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'Fonts'),
                os.path.expandvars('%localappdata%\\Microsoft\\Windows\\Fonts')
            ]
            
            # Procura padrões de nome da fonte (removendo espaços, traços, etc.)
            font_patterns = [
                font_name.lower(),
                font_name.lower().replace(' ', ''),
                font_name.lower().replace(' ', '-'),
                font_name.lower().replace(' ', '_')
            ]
            
            # Adiciona padrões com estilo
            if bold and italic:
                style_patterns = ['bolditalic', 'boldoblique', 'blackitalic', 'bi', 'z']
            elif bold:
                style_patterns = ['bold', 'black', 'bd', 'b']
            elif italic:
                style_patterns = ['italic', 'oblique', 'it', 'i']
            else:
                style_patterns = ['regular', 'normal', 'book']
            
            # Extensões de fonte a procurar
            font_extensions = ['.otf', '.ttf', '.ttc']
            
            # Procura em todos os diretórios
            for font_dir in font_dirs:
                if os.path.exists(font_dir):
                    for file_name in os.listdir(font_dir):
                        file_lower = file_name.lower()
                        
                        # Verifica se é um arquivo de fonte válido
                        if not any(file_lower.endswith(ext) for ext in font_extensions):
                            continue
                        
                        # Verifica se o nome do arquivo corresponde a algum padrão
                        for pattern in font_patterns:
                            if pattern in file_lower:
                                # Verifica se tem o estilo correto
                                style_match = not style_patterns  # Se não temos restrição de estilo
                                
                                for style in style_patterns:
                                    if style in file_lower:
                                        style_match = True
                                        break
                                
                                if style_match:
                                    font_path = os.path.join(font_dir, file_name)
                                    print(f"Encontrado diretamente no diretório: {font_path}")
                                    
                                    # Dá prioridade a fontes .otf
                                    if file_lower.endswith('.otf'):
                                        return font_path
                                    
                                    # Adiciona à lista de candidatos
                                    candidates.append(font_path)
            
            # Se encontrou candidatos, retorna o primeiro
            if candidates:
                return candidates[0]
            
            return None
        except Exception as e:
            print(f"Erro ao procurar fonte no registo: {e}")
            return None

    def _update_font_map_entry(self, font_name, style_key, file_name, full_path=None):
        """
        Atualiza o mapeamento de fontes com a informação de ficheiro encontrada
        
        Args:
            font_name (str): Nome da fonte (já normalizado)
            style_key (str): Chave de estilo ('normal', 'bold', 'italic', 'bold_italic')
            file_name (str): Nome do ficheiro
            full_path (str, opcional): Caminho completo do ficheiro
        """
        try:
            # Se encontrou, atualiza o mapeamento para uso futuro
            if font_name in self.fonts_map:
                if isinstance(self.fonts_map[font_name], dict):
                    if 'files' in self.fonts_map[font_name]:
                        self.fonts_map[font_name]['files'][style_key] = full_path if full_path else file_name
                    else:
                        self.fonts_map[font_name]['files'] = {style_key: full_path if full_path else file_name}
                    
                    # Atualiza também o status do estilo
                    if 'styles' in self.fonts_map[font_name]:
                        self.fonts_map[font_name]['styles'][style_key] = True
                else:
                    # Converte para novo formato
                    old_data = self.fonts_map[font_name]
                    self.fonts_map[font_name] = {
                        'styles': old_data,
                        'files': {style_key: full_path if full_path else file_name}
                    }
                    self.fonts_map[font_name]['styles'][style_key] = True
            else:
                # Cria nova entrada
                styles = {'normal': False, 'bold': False, 'italic': False, 'bold_italic': False}
                styles[style_key] = True
                self.fonts_map[font_name] = {
                    'styles': styles,
                    'files': {style_key: full_path if full_path else file_name}
                }
            
            # Salva as atualizações no arquivo para uso futuro
            try:
                with open(self.fonts_map_file, 'w', encoding='utf-8') as f:
                    json.dump(self.fonts_map, f, indent=2)
            except Exception as e:
                logging.error(f"Erro ao salvar mapeamento de fontes: {e}")
        except Exception as e:
            logging.error(f"Erro ao atualizar entrada de mapeamento de fontes: {e}")

    def _send_email(self, email_data):
        """Envia email com certificado anexado utilizando SMTP"""
        try:
            recipient = email_data["recipient"]
            cert_path = email_data["cert_path"]
            row_data = email_data["row_data"]
            
            return self._send_email_with_template(recipient, cert_path, row_data)
                
        except Exception as e:
            logging.error(f"Erro ao processar email: {e}")
            return False

    def send_email_safe(self, email_data, retries=2, delay=2):
        """Tenta enviar email com retries e logging robusto.
        Retorna True se enviado, False caso contrário.
        """
        try:
            recipient = email_data.get('recipient') or email_data.get('to')
            if not recipient or not isinstance(recipient, str):
                logging.error(f"Email inválido: {recipient}")
                return False
            recipient = recipient.strip()
            if '@' not in recipient:
                logging.error(f"Email inválido (sem @): {recipient}")
                return False

            attempt = 0
            while attempt <= retries:
                try:
                    ok = self._send_email(email_data)
                    if ok:
                        return True
                    else:
                        logging.error(f"Falha ao enviar email para {recipient} (tentativa {attempt+1})")
                except smtplib.SMTPAuthenticationError as e:
                    logging.error(f"Erro de autenticação SMTP ao enviar para {recipient}: {e}")
                    return False
                except Exception as e:
                    logging.error(f"Erro ao enviar email para {recipient}: {e}")

                attempt += 1
                if attempt <= retries:
                    backoff = delay * (2 ** (attempt-1))
                    logging.info(f"Re-tentando em {backoff}s...")
                    time.sleep(backoff)

            logging.error(f"Todas tentativas falharam para {recipient}")
            return False

        except Exception as e:
            logging.error(f"Erro inesperado em send_email_safe: {e}")
            return False
    
    def _send_email_with_template(self, recipient, cert_path, row_data):
        """Envia um email com template para o destinatário via SMTP"""
        try:
            # Valida o email do destinatário
            if not recipient or not isinstance(recipient, str):
                logging.error(f"Email inválido: {recipient}")
                return False
            
            recipient = recipient.strip()
            if '@' not in recipient:
                logging.error(f"Email inválido (sem @): {recipient}")
                return False
            
            if not self.email_config:
                logging.error("Configuração de email não encontrada")
                return False
            
            # Obtém dados do template
            template_path = self.email_config.get('template_path')
            subject = self.email_config.get('subject', 'Certificado')
            
            # Usa o conteúdo do widget de texto que guarda o corpo do email
            body_text = ""
            if hasattr(self, 'email_text'):
                body_text = self.email_text.get("1.0", "end-1c")
            
            cc = self.email_config.get('cc', [])
            bcc = self.email_config.get('bcc', [])
            
            # Função para formatar texto com placeholders
            def format_text(text, data):
                if not text:
                    return ""
                try:
                    # Primeiro tenta uma abordagem segura: substitui apenas placeholders conhecidos
                    result = text
                    for key, value in data.items():
                        placeholder = f"{{{key}}}"
                        if placeholder in result:
                            result = result.replace(placeholder, str(value))
                    
                    # Adiciona placeholders globais do config
                    global_placeholders = getattr(config, 'GLOBAL_PLACEHOLDERS', {})
                    for key, value_or_func in global_placeholders.items():
                        if callable(value_or_func):
                            try:
                                value = value_or_func()
                            except:
                                value = f"Erro ao calcular GLOBAL_{key}"
                        else:
                            value = value_or_func
                        
                        placeholder = f"{{GLOBAL_{key}}}"
                        if placeholder in result:
                            result = result.replace(placeholder, str(value))
                    
                    # Processa datas dinâmicas
                    import datetime
                    today = datetime.datetime.now()
                    result = result.replace("{DATA_ATUAL}", today.strftime("%d/%m/%Y"))
                    return result
                except Exception as e:
                    logging.error(f"Erro ao formatar texto: {e}")
                    return text
            
            # Formata o assunto e corpo com os dados do CSV
            subject = format_text(subject, row_data)
            
            # Atualiza o placeholder global do assunto para uso no template
            config.GLOBAL_PLACEHOLDERS['subject'] = subject
            
            # Se há template e estamos usando template, usa ele; senão, usa o corpo simples
            use_template = self.email_config.get('use_template', False)
            if use_template and template_path and os.path.exists(template_path):
                try:
                    with open(template_path, 'r', encoding='utf-8') as f:
                        template_content = f.read()
                    
                    body_html = format_text(template_content, row_data)
                except Exception as e:
                    logging.error(f"Erro ao carregar template: {e}")
                    body_html = format_text(body_text, row_data).replace('\n', '<br>')
            else:
                body_html = format_text(body_text, row_data).replace('\n', '<br>')
            
            # Personaliza o nome do anexo se configurado
            attachment_name = os.path.basename(cert_path)
            if 'attachment_name' in self.email_config and self.email_config['attachment_name']:
                try:
                    # Cria uma cópia temporária com o nome personalizado
                    custom_name = format_text(self.email_config["attachment_name"], row_data)
                    if custom_name:
                        temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp")
                        os.makedirs(temp_dir, exist_ok=True)
                        temp_path = os.path.join(temp_dir, custom_name)
                        
                        # Copia o arquivo para o diretório temporário com o nome personalizado
                        import shutil
                        shutil.copy2(cert_path, temp_path)
                        cert_path = temp_path  # Usa o caminho temporário
                except Exception as e:
                    logging.error(f"Erro ao personalizar nome do anexo: {e}")
            
            # Monta os dados do email
            email_data = {
                'to': recipient,
                'subject': subject,
                'body': body_html,
                'cc': cc,
                'bcc': bcc,
                'attachment': cert_path
            }
            
            # Envia o email via SMTP
            result = self._send_email_smtp(email_data)
            
            if not result:
                logging.error(f"Falha ao enviar email para {recipient}")
                return False
                
            logging.info(f"Email enviado com sucesso para {recipient}")
            return True
            
        except Exception as e:
            logging.error(f"Erro ao enviar email com template: {e}")
            return False

    def validate_smtp_config(self):
        """Valida a configuração SMTP e retorna mensagem de erro se houver problemas"""
        try:
            # Verifica se as configurações existem
            if not hasattr(config, 'SMTP_SERVER') or not config.SMTP_SERVER:
                return "❌ Erro de Configuração: SMTP_SERVER não está definido no config.py"
            
            if not hasattr(config, 'SMTP_PORT') or not config.SMTP_PORT:
                return "❌ Erro de Configuração: SMTP_PORT não está definido no config.py"
            
            if not hasattr(config, 'SMTP_USER') or not config.SMTP_USER:
                return "❌ Erro de Configuração: SMTP_USER não está definido no config.py"
            
            if not hasattr(config, 'SMTP_PASSWORD') or not config.SMTP_PASSWORD:
                return "❌ Erro de Configuração: SMTP_PASSWORD não está definido no config.py"
            
            # Verifica se o servidor é válido (não vazio)
            if config.SMTP_SERVER.strip() == '':
                return "❌ Erro de Configuração: SMTP_SERVER está vazio"
            
            if '@' in config.SMTP_SERVER:
                return "❌ Erro de Configuração: SMTP_SERVER não deve conter '@'. Deve ser algo como 'smtp.gmail.com'"
            
            if len(config.SMTP_USER.strip()) == 0:
                return "❌ Erro de Configuração: SMTP_USER está vazio"
            
            if len(config.SMTP_PASSWORD.strip()) == 0:
                return "❌ Erro de Configuração: SMTP_PASSWORD está vazio"
            
            # Tudo OK
            return None
        except Exception as e:
            return f"❌ Erro ao validar configuração: {e}"
    
    def show_smtp_error(self, error_type, original_error=None):
        """Mostra erro SMTP de forma amigável ao utilizador com soluções sugeridas"""
        error_msg = "❌ Erro de Configuração de Email\n\n"
        
        if isinstance(original_error, str):
            error_str = original_error.lower()
        else:
            error_str = str(original_error).lower() if original_error else ""
        
        # Detecta o tipo de erro e mostra solução apropriada
        if "getaddrinfo failed" in error_str or "11001" in error_str or "nodename nor servname" in error_str:
            error_msg += f"🔍 PROBLEMA: Servidor SMTP '{config.SMTP_SERVER}' não encontrado\n\n"
            error_msg += "SOLUÇÕES (na ordem sugerida):\n\n"
            error_msg += "1️⃣  Verifique se o nome do servidor está correto em config.py\n"
            error_msg += "    Servidores comuns:\n"
            error_msg += "    • Gmail: smtp.gmail.com:587 (TLS=True)\n"
            error_msg += "    • Outlook: smtp-mail.outlook.com:587 (TLS=True)\n"
            error_msg += "    • Seu domínio: mail.seudominio.com\n\n"
            error_msg += "2️⃣  Verifique a sua ligação à Internet\n\n"
            error_msg += "3️⃣  Contacte o administrador de email do seu servidor\n"
        
        elif "timeout" in error_str or "10060" in error_str or "timed out" in error_str:
            error_msg += f"⏱️  PROBLEMA: Servidor SMTP '{config.SMTP_SERVER}:{config.SMTP_PORT}' não responde\n\n"
            error_msg += "SOLUÇÕES (na ordem sugerida):\n\n"
            error_msg += "1️⃣  Verifique se a porta está correta em config.py:\n"
            error_msg += "    • 587 para TLS/STARTTLS\n"
            error_msg += "    • 465 para SSL\n\n"
            error_msg += "2️⃣  Servidor pode estar offline - tente mais tarde\n\n"
            error_msg += "3️⃣  Verifique firewall/antivírus que podem estar bloqueando\n"
        
        elif "authentication" in error_str or "credential" in error_str or "535" in error_str or "unauthorized" in error_str:
            error_msg += "🔐 PROBLEMA: Credenciais SMTP inválidas\n\n"
            error_msg += "SOLUÇÕES (na ordem sugerida):\n\n"
            error_msg += f"1️⃣  Verifique em config.py:\n"
            error_msg += f"    • SMTP_USER: {config.SMTP_USER}\n"
            error_msg += f"    • SMTP_PASSWORD: [verificar se está correto]\n\n"
            error_msg += "2️⃣  Se tem autenticação 2FA (Gmail, Outlook):\n"
            error_msg += "    • Gmail: gere uma 'Senha de Aplicação' em myaccount.google.com\n"
            error_msg += "    • Outlook: ative 'Autenticação de Aplicação' nas definições\n\n"
            error_msg += "3️⃣  Teste a senha manualmente no seu cliente de email\n"
        
        elif "tls" in error_str or "ssl" in error_str or "certificate" in error_str or "ehlo" in error_str:
            error_msg += "🔒 PROBLEMA: Erro de segurança SSL/TLS\n\n"
            error_msg += "SOLUÇÕES (na ordem sugerida):\n\n"
            error_msg += "1️⃣  Verifique em config.py:\n"
            error_msg += f"    • SMTP_USE_TLS = {config.SMTP_USE_TLS}\n"
            error_msg += f"    • SMTP_PORT = {config.SMTP_PORT}\n\n"
            error_msg += "2️⃣  Configure a porta e TLS corretamente:\n"
            error_msg += "    • Porta 587 → SMTP_USE_TLS = True\n"
            error_msg += "    • Porta 465 → SMTP_USE_TLS = False\n\n"
            error_msg += "3️⃣  Se o certificado SSL for inválido, contacte o administrador\n"
        
        elif "connection refused" in error_str or "refused" in error_str:
            error_msg += f"🚫 PROBLEMA: Conexão recusada por '{config.SMTP_SERVER}:{config.SMTP_PORT}'\n\n"
            error_msg += "SOLUÇÕES (na ordem sugerida):\n\n"
            error_msg += "1️⃣  Verifique a porta em config.py (normalmente 587 ou 465)\n\n"
            error_msg += "2️⃣  Servidor pode estar offline ou não suporta essa porta\n\n"
            error_msg += "3️⃣  Firewall/antivírus pode estar bloqueando a ligação\n"
        
        else:
            error_msg += f"⚠️  Erro desconhecido:\n\n{str(original_error)}\n\n"
            error_msg += "Por favor, contacte o administrador com esta mensagem."
        
        # Mostra a mensagem ao utilizador
        try:
            messagebox.showerror("❌ Erro de Email", error_msg)
        except Exception as e:
            print(f"\n{error_msg}\n")
            logging.error(f"Erro ao mostrar dialogo: {e}")

    def _send_email_smtp(self, email_data):
        """Envia email via SMTP simples"""
        try:
            # Valida configuração primeiro
            validation_error = self.validate_smtp_config()
            if validation_error:
                logging.error(validation_error)
                messagebox.showerror("Erro de Configuração SMTP", validation_error)
                return False
            
            # Verifica credenciais no config
            if not all(hasattr(config, attr) for attr in ['SMTP_SERVER', 'SMTP_PORT', 'SMTP_USER', 'SMTP_PASSWORD']):
                error = "Credenciais SMTP não configuradas no config.py"
                logging.error(error)
                messagebox.showerror("Erro de Configuração", error)
                return False
            
            # Prepara dados do email
            recipient = email_data.get('to')
            subject = email_data.get('subject', 'Certificado')
            body = email_data.get('body', '')
            cc = email_data.get('cc', [])
            bcc = email_data.get('bcc', [])
            attachment_path = email_data.get('attachment')
            
            # Cria a mensagem de email
            msg = MIMEMultipart('alternative')
            msg['Subject'] = subject
            msg['From'] = f"{config.SMTP_FROM_NAME} <{config.SMTP_USER}>"
            msg['To'] = recipient
            
            if cc:
                msg['Cc'] = ', '.join(cc) if isinstance(cc, list) else cc
            
            # Adiciona o corpo como HTML
            msg.attach(MIMEText(body, 'html', 'utf-8'))
            
            # Adiciona anexo se existir
            if attachment_path and os.path.exists(attachment_path):
                try:
                    with open(attachment_path, 'rb') as attachment:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(attachment.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
                        msg.attach(part)
                except Exception as e:
                    logging.error(f"Erro ao anexar arquivo: {e}")
                    return False
            
            # Cria lista de destinatários (para, cc, bcc)
            all_recipients = [recipient]
            if cc:
                all_recipients.extend(cc if isinstance(cc, list) else [cc])
            if bcc:
                all_recipients.extend(bcc if isinstance(bcc, list) else [bcc])
            
            # Conecta ao servidor SMTP e envia
            try:
                if config.SMTP_USE_TLS:
                    server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT)
                    server.starttls()
                else:
                    server = smtplib.SMTP_SSL(config.SMTP_SERVER, config.SMTP_PORT)
                
                server.login(config.SMTP_USER, config.SMTP_PASSWORD)
                server.sendmail(config.SMTP_USER, all_recipients, msg.as_string())
                server.quit()
                
                return True
            except smtplib.SMTPAuthenticationError as e:
                error = "Erro de autenticação SMTP - verifique credenciais"
                logging.error(error)
                self.show_smtp_error("auth", e)
                return False
            except smtplib.SMTPException as e:
                logging.error(f"Erro SMTP: {e}")
                self.show_smtp_error("smtp", e)
                return False
            except socket.gaierror as e:
                logging.error(f"Erro ao resolver servidor: {e}")
                self.show_smtp_error("dns", e)
                return False
            except socket.timeout as e:
                logging.error(f"Timeout ao conectar servidor: {e}")
                self.show_smtp_error("timeout", e)
                return False
            except ConnectionRefusedError as e:
                logging.error(f"Conexão recusada: {e}")
                self.show_smtp_error("connection", e)
                return False
                
        except Exception as e:
            logging.error(f"Erro ao enviar email via SMTP: {e}")
            self.show_smtp_error("general", e)
            return False


            
    def on_closing(self):
        """Manipula o evento de fechamento da aplicação.
        
        Esta função:
        1. Apaga a pasta temp e seu conteúdo
        2. Fecha a aplicação
        """
        try:
            # Obtém o caminho da pasta temp
            temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp")
            
            # Verifica se a pasta existe
            if os.path.exists(temp_dir) and os.path.isdir(temp_dir):
                import shutil
                # Apaga a pasta temp e todo seu conteúdo
                shutil.rmtree(temp_dir)
                logging.info(f"Pasta temporária removida: {temp_dir}")
        except Exception as e:
            logging.error(f"Erro ao apagar pasta temporária: {str(e)}")
            
        # Fecha a aplicação
        self.destroy()
    
    def _setup_styles(self):
        """Configura os estilos ttk para a aplicação."""
        style = ttk.Style()
        style.configure("TScale", background="#3a3a3a")
        style.configure("TButton", background="#444444", foreground="white")
        style.configure("TLabel", background="#252525", foreground="white")
        style.configure("TFrame", background="#252525")
        style.configure("TLabelframe", background="#252525", foreground="white")
        style.configure("TLabelframe.Label", background="#252525", foreground="white")
        style.configure("TCombobox", fieldbackground="#333333", background="#444444", foreground="white")
        style.configure("TScrollbar", background="#444444", troughcolor="#333333", bordercolor="#252525", 
                        arrowcolor="white", relief="flat")
        style.configure("TEntry", fieldbackground="#333333", foreground="white")
        
        # Configuração específica para o TNotebook para evitar a borda branca
        style.configure("TNotebook", background="#252525", borderwidth=0)
        style.configure("TNotebook.Tab", background="#333333", foreground="black", borderwidth=0)
        style.map("TNotebook.Tab", background=[("selected", "#444444")], foreground=[("selected", "black")])
        # Elimina a borda das abas do notebook
        style.layout("TNotebook", [("TNotebook.client", {"sticky": "nswe"})])
        style.layout("TNotebook.Tab", [
            ("TNotebook.tab", {
                "sticky": "nswe",
                "children": [
                    ("TNotebook.padding", {
                        "side": "top",
                        "sticky": "nswe",
                        "children": [
                            ("TNotebook.label", {"side": "top", "sticky": ''})
                        ]
                    })
                ]
            })
        ])
    
    def new_project(self):
        """Cria um novo projeto, limpando o canvas e todos os elementos após confirmação"""
        # Verifica se existe conteúdo no canvas para confirmar
        has_content = len(self.items) > 0 or self.model_img is not None
        
        if has_content:
            # Pede confirmação ao utilizador
            if messagebox.askyesno("Novo Projeto", 
                                  "Tem a certeza que deseja criar um novo projeto?\nTodas as alterações não guardadas serão perdidas."):
                # Limpa o canvas e todos os elementos
                self.clear_all_layers()
                self.model_img = None
                self.pil_model = None
                
                # Limpa o canvas visualmente
                self.canvas.delete("all")
                
                # Reseta as dimensões do canvas
                self.canvas.config(width=800, height=600, scrollregion=(0, 0, 800, 600))
                
                # Reseta o zoom
                self.zoom_factor = 1.0
                self.zoom_var.set(1.0)
                self.zoom_label.config(text="100%")
                
                # Centraliza a vista
                self.canvas.xview_moveto(0)
                self.canvas.yview_moveto(0)
                
                # Limpa o título do ficheiro atual
                self.title("VisuMaker - Novo Projeto")
        else:
            # Se não houver conteúdo, apenas limpa o título
            self.title("VisuMaker - Novo Projeto")

    def forget_csv(self):
        """Limpa as referências aos dados do CSV"""
        self.df = None
        if self.forget_csv_btn:
            self.forget_csv_btn.config(state=tk.DISABLED)
        messagebox.showinfo("CSV Esquecido", "O CSV carregado foi esquecido.")

class EmailConfigWindow(tk.Toplevel):
    """Janela para configuração das opções de email"""
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Novo Email")
        self.geometry("900x650")  # Aumentando a largura para acomodar dois painéis lado a lado
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self.configure(bg='#303030')  # Mantido o tema escuro
        
        # Referência às configurações do app principal
        self.email_config = parent.email_config
        
        # Modo de envio (se vai enviar ao fechar ou só configurar)
        self.send_mode = False
        
        # Layout principal
        self.create_widgets()

    def create_widgets(self):
        """Cria os widgets da janela"""
        # Container principal divido em dois painéis
        main_container = tk.Frame(self, bg='#303030')
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Painel esquerdo para configurações
        left_panel = tk.Frame(main_container, bg='#303030', padx=10, pady=10)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Painel direito para preview
        right_panel = tk.Frame(main_container, bg='#303030', padx=10, pady=10)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Cabeçalho no painel esquerdo
        header_frame = tk.Frame(left_panel, bg='#303030', padx=5, pady=5)
        header_frame.pack(fill=tk.X, pady=5)
        
        # Destinatários adicionais: CC
        cc_frame = tk.Frame(header_frame, bg='#303030')
        cc_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(cc_frame, text="Cc:", bg='#303030', fg='white', width=10, anchor='w').pack(side=tk.LEFT)
        
        self.cc_var = tk.StringVar(value=','.join(self.email_config.get("cc", [])))
        cc_entry = tk.Entry(cc_frame, textvariable=self.cc_var, bg='#444444', fg='white')
        cc_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Destinatários adicionais: BCC
        bcc_frame = tk.Frame(header_frame, bg='#303030')
        bcc_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(bcc_frame, text="Bcc:", bg='#303030', fg='white', width=10, anchor='w').pack(side=tk.LEFT)
        
        self.bcc_var = tk.StringVar(value=','.join(self.email_config.get("bcc", [])))
        bcc_entry = tk.Entry(bcc_frame, textvariable=self.bcc_var, bg='#444444', fg='white')
        bcc_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Assunto
        subject_frame = tk.Frame(header_frame, bg='#303030')
        subject_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(subject_frame, text="Assunto:", bg='#303030', fg='white', width=10, anchor='w').pack(side=tk.LEFT)
        
        self.subject_var = tk.StringVar(value=self.email_config.get("subject", "Documento Visual"))
        subject_entry = tk.Entry(subject_frame, textvariable=self.subject_var, bg='#444444', fg='white')
        subject_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Nome do anexo
        attach_frame = tk.Frame(header_frame, bg='#303030')
        attach_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(attach_frame, text="Anexo:", bg='#303030', fg='white', width=10, anchor='w').pack(side=tk.LEFT)
        
        self.attachment_name_var = tk.StringVar(value=self.email_config.get("attachment_name", "Documento.png"))
        attach_entry = tk.Entry(attach_frame, textvariable=self.attachment_name_var, bg='#444444', fg='white')
        attach_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Template OFT (escondido por padrão)
        template_frame = tk.LabelFrame(left_panel, text="Template de Email (opcional)", bg='#303030', fg='white',
                                     padx=10, pady=5)
        template_frame.pack(fill=tk.X, pady=5)
        
        self.use_template_var = tk.BooleanVar(value=self.email_config.get("use_template", False))
        use_template_check = tk.Checkbutton(template_frame, text="Usar template HTML", 
                                         variable=self.use_template_var,
                                         command=self.toggle_template,
                                         bg='#303030', fg='white', selectcolor='#505050',
                                         activebackground='#303030', activeforeground='white')
        use_template_check.pack(fill=tk.X)
        
        # Seleção do template
        template_select_frame = tk.Frame(template_frame, bg='#303030')
        template_select_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(template_select_frame, text="Template:", bg='#303030', fg='white').pack(side=tk.LEFT)
        
        self.template_path_var = tk.StringVar(value=self.email_config.get("template_path", ""))
        template_entry = tk.Entry(template_select_frame, textvariable=self.template_path_var,
                                width=40, bg='#444444', fg='white')
        template_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        browse_button = tk.Button(template_select_frame, text="Procurar", 
                               command=self.browse_template, bg='#444444', fg='white')
        browse_button.pack(side=tk.RIGHT)
        
        # Botão para ver preview HTML
        self.html_preview_button = tk.Button(template_frame, text="Ver Preview HTML", 
                                   command=self.show_html_preview, 
                                   bg='#4CAF50', fg='white', state=tk.DISABLED)
        self.html_preview_button.pack(fill=tk.X, pady=5)
        
        # Corpo do email
        email_body_frame = tk.Frame(left_panel, bg='#303030', padx=5, pady=5)
        email_body_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        email_body_label = tk.Label(email_body_frame, text="Corpo do Email:", bg='#303030', fg='white', anchor='w')
        email_body_label.pack(fill=tk.X, pady=(5, 2))
        
        # Campo de texto do corpo do email (expandido para usar o espaço restante)
        self.email_body_text = tk.Text(email_body_frame, width=40, 
                                    bg='#444444', fg='white', padx=10, pady=10,
                                    insertbackground='white', font=('Arial', 10))
        self.email_body_text.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)
        
        # Adiciona borda ao campo de texto
        self.email_body_text.config(relief=tk.SOLID, borderwidth=1)
        
        # Pegamos o corpo do email do Text widget da aplicação principal
        if hasattr(self.parent, 'email_text'):
            body_content = self.parent.email_text.get("1.0", tk.END)
            self.email_body_text.insert("1.0", body_content)
        else:
            # Ou usamos o padrão do config
            self.email_body_text.insert("1.0", config.EMAIL_BODY)
        
        # Removidos os botões do painel esquerdo
        # (serão adicionados abaixo da preview)
        
        # Preview do email (ocupando todo o painel direito)
        preview_frame = tk.LabelFrame(right_panel, text="Preview do Email", 
                                    bg='#303030', fg='white',
                                    padx=10, pady=5)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Botão para atualizar a preview
        preview_button = tk.Button(preview_frame, text="Atualizar Preview", 
                                command=self.update_preview, 
                                bg='#444444', fg='white')
        preview_button.pack(fill=tk.X, pady=5)
        
        # Campo de texto da preview (somente leitura) - usando todo o espaço disponível
        # Container para a preview (frame que conterá o texto ou HTML)
        preview_container = tk.Frame(preview_frame, bg='#f8f8f8')
        preview_container.pack(fill=tk.BOTH, expand=True)
        
        # Frame para conter o widget de texto (visível por padrão)
        self.text_preview_frame = tk.Frame(preview_container, bg='#f8f8f8')
        self.text_preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Texto simples para a preview (somente leitura)
        self.preview_text = tk.Text(self.text_preview_frame, width=40, 
                                 bg='#f8f8f8', fg='#333333')  # Mantém a preview clara
        self.preview_text.pack(fill=tk.BOTH, expand=True)
        self.preview_text.config(state=tk.DISABLED, relief=tk.SOLID, borderwidth=1)
        
        # Não há mais webview embutido, só preview em texto e botão para abrir no navegador externo
        self.temp_html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp", "preview.html")
        os.makedirs(os.path.dirname(self.temp_html_path), exist_ok=True)
        
        # Botões de ação (na parte inferior do painel direito)
        buttons_frame = tk.Frame(right_panel, bg='#303030')
        buttons_frame.pack(fill=tk.X, pady=10, anchor='e')
        
        # Botão para cancelar
        cancel_button = tk.Button(buttons_frame, text="Cancelar", width=10, command=self.destroy,
                              bg='#444444', fg='white')
        cancel_button.pack(side=tk.LEFT, padx=5)
        
        # Botão para salvar configuração
        save_button = tk.Button(buttons_frame, text="Guardar Config", width=15, command=self.save_config,
                          bg='#2196F3', fg='white', font=('Arial', 9, 'bold'))
        save_button.pack(side=tk.LEFT, padx=5)
        
        # Botão para enviar, sempre visível
        self.send_button = tk.Button(buttons_frame, text="Enviar Emails", width=15, 
                                 command=self.save_and_send, bg='#4CAF50', fg='white',
                                 font=('Arial', 9, 'bold'))
        self.send_button.pack(side=tk.RIGHT, padx=5)
        
        # Criamos uma preview inicial
        self.update_preview()
        
        # Preenche os controles com os valores atuais
        self.toggle_template()
    
    def update_preview(self):
        """Atualiza a preview do email com dados reais do primeiro registro"""
        # Verifica se o CSV foi carregado primeiro
        if hasattr(self.parent, 'df') and self.parent.df is not None and len(self.parent.df) > 0:
            # Usa o primeiro registro do CSV
            sample_data = self.parent.df.iloc[0].to_dict()
        else:
            # Dados de exemplo quando não há CSV carregado
            sample_data = {
                "nome": "João Silva",
                "email": "joao.silva@exemplo.com",
                "evento": "Workshop de Programação",
                "data": "01/01/2023",
                "local": "Online",
                "cargo": "Participante",
                "empresa": "Empresa Exemplo"
            }
        
        # Verifica se estamos usando template HTML
        if self.use_template_var.get() and self.template_path_var.get():
            template_file = self.template_path_var.get()
            if os.path.exists(template_file):
                try:
                    # Lê o arquivo HTML diretamente
                    with open(template_file, 'r', encoding='utf-8') as f:
                        body_html_raw = f.read()
                    
                    # Usa diretamente o assunto informado na interface
                    subject = self.subject_var.get()
                    
                    # Configura o placeholder global do assunto para uso no template
                    import config
                    config.GLOBAL_PLACEHOLDERS['subject'] = subject
                    
                    # Obtém todos os placeholders formatados
                    all_placeholders = self.parent.get_all_placeholders(sample_data)
                    
                    # Formata o assunto e corpo com os dados
                    try:
                        formatted_subject = subject.format(**all_placeholders)
                    except:
                        formatted_subject = subject
                        
                    try:
                        formatted_body = self._format_text_with_data(body_html_raw, all_placeholders)
                    except Exception as e:
                        formatted_body = body_html_raw
                        logging.error(f"Erro ao formatar HTML: {e}")
                    
                    # Verifica se o conteúdo parece ser HTML
                    is_html = '<html' in formatted_body.lower() or '<body' in formatted_body.lower()
                    
                    # Prepara o HTML para abrir no navegador externo
                    if is_html:
                        try:
                            full_html = f"""
                            <!DOCTYPE html>
                            <html>
                            <head>
                                <meta charset='utf-8'>
                                <style>
                                    body {{ font-family: Arial, sans-serif; margin: 0; padding: 20px; background-color: #f8f8f8; }}
                                    .content {{ background-color: white; padding: 15px; border: 1px solid #ddd; border-radius: 4px; }}
                                </style>
                            </head>
                            <body><div class='content'>{formatted_body}</div></body></html>"""
                            with open(self.temp_html_path, "w", encoding="utf-8") as f:
                                f.write(full_html)
                        except Exception as e:
                            logging.error(f"Erro ao preparar HTML para preview: {e}")
                    
                    # Mostra apenas mensagem indicando que template HTML está selecionada
                    self.text_preview_frame.pack(fill=tk.BOTH, expand=True)
                    
                    # Atualiza a preview em modo texto
                    self.preview_text.config(state=tk.NORMAL)
                    self.preview_text.delete("1.0", tk.END)
                    
                    # Mostra apenas a mensagem
                    self.preview_text.insert(tk.END, "Template HTML selecionada\n\nUse o botão 'Ver Preview HTML' para visualizar o template formatado.")
                    self.preview_text.config(state=tk.DISABLED)
                    return
                except Exception as e:
                    # Se houver erro ao processar o template
                    self.preview_text.config(state=tk.NORMAL)
                    self.preview_text.delete("1.0", tk.END)
                    self.preview_text.insert("1.0", f"Erro ao processar template HTML: {str(e)}")
                    self.preview_text.config(state=tk.DISABLED)
                    return
        
        # Se não estiver usando template ou o template não foi encontrado:
        # Pega o texto atual do corpo
        body = self.email_body_text.get("1.0", "end-1c")
        
        # Formata com dados de exemplo
        try:
            formatted_body = self._format_text_with_data(body, sample_data)
            formatted_subject = self._format_text_with_data(self.subject_var.get(), sample_data)
            
            # Verifica se o conteúdo parece ser HTML
            is_html = '<html' in formatted_body.lower() or '<body' in formatted_body.lower()
            
            # Prepara o HTML para abrir no navegador externo
            if is_html:
                try:
                    full_html = f"""
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <meta charset='utf-8'>
                        <style>
                            body {{ font-family: Arial, sans-serif; margin: 0; padding: 20px; background-color: #f8f8f8; }}
                            .content {{ background-color: white; padding: 15px; border: 1px solid #ddd; border-radius: 4px; }}
                        </style>
                    </head>
                    <body><div class='content'>{formatted_body}</div></body></html>"""
                    with open(self.temp_html_path, "w", encoding="utf-8") as f:
                        f.write(full_html)
                except Exception as e:
                    logging.error(f"Erro ao preparar HTML para preview: {e}")
            
            if is_html:
                # Esconde a preview de texto quando for HTML
                self.text_preview_frame.pack_forget()
                return
            # Se não for HTML, mostra a preview de texto normalmente
            self.text_preview_frame.pack(fill=tk.BOTH, expand=True)
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert(tk.END, formatted_body)
            self.preview_text.config(state=tk.DISABLED)
        except Exception as e:
            # Em caso de erro na formatação, mostra o texto original
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert(tk.END, f"Erro ao processar texto: {str(e)}\n\n{body}")
            self.preview_text.config(state=tk.DISABLED)
            
    def _format_text_with_data(self, text, data):
        """Substitui placeholders no texto por valores dos dados e processa escape characters"""
        if not text:
            return ""
            
        # Pré-processamento para lidar com caracteres escapados
        i = 0
        processed_text = ""
        while i < len(text):
            # Verifica se temos um escape character (\)
            if text[i] == '\\' and i + 1 < len(text):
                if text[i+1] == '{':  # \{ torna-se { literal
                    processed_text += '{'
                    i += 2  # Pula o \ e o {
                elif text[i+1] == '\\':  # \\ torna-se \ literal
                    processed_text += '\\'
                    i += 2  # Pula ambos \
                else:  # Qualquer outro caractere após \ é mantido como está
                    processed_text += text[i:i+2]
                    i += 2
            else:
                processed_text += text[i]
                i += 1
        
        # Agora usa o método format padrão do Python com o texto pré-processado
        try:
            formatted = processed_text.format(**self.parent.get_all_placeholders(data))
            return formatted
        except KeyError as e:
            # Se algum placeholder não existir, retorna o texto original
            logging.error(f"Aviso: Placeholder não encontrado: {e}")
            return processed_text
        except Exception as e:
            return processed_text
            
    def toggle_template(self):
        """Ativa/desativa os controles relacionados ao template"""
        use_template = self.use_template_var.get()
        state = "normal" if use_template else "disabled"
        
        # Atualiza visualmente o estado dos controles
        for widget in self.winfo_children():
            if isinstance(widget, tk.Frame):
                for frame in widget.winfo_children():
                    if isinstance(frame, tk.LabelFrame) and frame.cget("text") == "Template de Email (opcional)":
                        for child in frame.winfo_children():
                            if isinstance(child, tk.Frame):
                                for w in child.winfo_children():
                                    if isinstance(w, tk.Entry) or isinstance(w, tk.Button):
                                        w.config(state=state)
        
        # Configura os botões de preview baseado na seleção do template
        if use_template:
            # Se estiver usando template HTML, ativa o botão de preview HTML e desativa o de atualizar preview
            self.html_preview_button.config(state="normal")
            
            # Encontra o botão "Atualizar Preview" e o desativa
            for widget in self.winfo_children():
                if isinstance(widget, tk.Frame):
                    for frame in widget.winfo_children():
                        if isinstance(frame, tk.LabelFrame) and frame.cget("text") == "Preview do Email":
                            for btn in frame.winfo_children():
                                if isinstance(btn, tk.Button) and btn.cget("text") == "Atualizar Preview":
                                    btn.config(state="disabled")
        else:
            # Se não estiver usando template, desativa o botão de preview HTML e ativa o de atualizar preview
            self.html_preview_button.config(state="disabled")
            
            # Encontra o botão "Atualizar Preview" e o ativa
            for widget in self.winfo_children():
                if isinstance(widget, tk.Frame):
                    for frame in widget.winfo_children():
                        if isinstance(frame, tk.LabelFrame) and frame.cget("text") == "Preview do Email":
                            for btn in frame.winfo_children():
                                if isinstance(btn, tk.Button) and btn.cget("text") == "Atualizar Preview":
                                    btn.config(state="normal")
                                        
        # Altera o estado do campo de corpo de email baseado na seleção do template
        email_body_state = "disabled" if use_template else "normal"
        self.email_body_text.config(state=email_body_state)
        
        # NÃO atualiza a preview automaticamente como antes
    
    def browse_template(self):
        """Abre diálogo para selecionar um template HTML"""
        path = filedialog.askopenfilename(
            defaultextension=".html",
            filetypes=[("HTML Template", "*.html;*.htm"), ("All files", "*.*")]
        )
        if path:
            self.template_path_var.set(path)
    
    def save_config(self):
        """Salva as configurações e fecha a janela"""
        # Atualiza as configurações do app principal
        # Sem alterar o use_o365, mantém o valor existente
        self.email_config["use_template"] = self.use_template_var.get()
        self.email_config["template_path"] = self.template_path_var.get()
        self.email_config["subject"] = self.subject_var.get()
        self.email_config["attachment_name"] = self.attachment_name_var.get()
        
        # Processa CC e BCC (separados por vírgula)
        cc_emails = [email.strip() for email in self.cc_var.get().split(',') if email.strip()]
        bcc_emails = [email.strip() for email in self.bcc_var.get().split(',') if email.strip()]
        
        self.email_config["cc"] = cc_emails
        self.email_config["bcc"] = bcc_emails
        
        # Atualiza o corpo do email no app principal
        if hasattr(self.parent, 'email_text'):
            body_content = self.email_body_text.get("1.0", "end-1c")
            self.parent.email_text.delete("1.0", tk.END)
            self.parent.email_text.insert("1.0", body_content)
        
        # Atualiza o arquivo config.py com as novas configurações
        try:
            # Lê o arquivo inteiro para trabalhar com seu conteúdo completo
            with open('config.py', 'r', encoding='utf-8') as f:
                conteudo_completo = f.read()
            
            # Definição padrão das configurações de email
            configuracao_email = f"""
# Configurações adicionais de email
DEFAULT_EMAIL_CONFIG = {{
    "use_template": {str(self.email_config["use_template"])},
    "template_path": "{self.email_config["template_path"]}",
    "subject": "{self.email_config["subject"]}",
    "attachment_name": "{self.email_config["attachment_name"]}",  # Nome do anexo (pode usar placeholders)
    "cc": {repr(self.email_config["cc"])},  # Lista de emails para CC
    "bcc": {repr(self.email_config["bcc"])}  # Lista de emails para BCC
}}
"""
            
            # Usa expressão regular que captura o bloco inteiro da configuração, mesmo com múltiplos }
            import re
            padrao = r'# Configurações adicionais de email\s*DEFAULT_EMAIL_CONFIG\s*=\s*\{[\s\S]*?}\s*(?=\n\s*(?:#|$))'
            
            # Verifica se encontrou o padrão para substituir
            if re.search(padrao, conteudo_completo, re.DOTALL):
                # Substitui o bloco inteiro pelo novo
                conteudo_novo = re.sub(padrao, configuracao_email.strip(), conteudo_completo, flags=re.DOTALL)
            else:
                # Se não encontrou, procura onde inserir (antes do EMAIL_BODY)
                padrao_email_subject = r'# Mensagem de e‑mail com placeholders'
                match_email = re.search(padrao_email_subject, conteudo_completo)
                
                if match_email:
                    # Insere antes do EMAIL_BODY
                    pos = match_email.start()
                    conteudo_novo = conteudo_completo[:pos] + configuracao_email + "\n" + conteudo_completo[pos:]
                else:
                    # Último recurso: adiciona no final do arquivo
                    conteudo_novo = conteudo_completo + "\n" + configuracao_email
            
            # Escreve o arquivo com o conteúdo atualizado
            with open('config.py', 'w', encoding='utf-8') as f:
                f.write(conteudo_novo)
                
        except Exception as e:
            logging.error(f"Erro ao atualizar config.py: {e}")
        
        # Fecha a janela
        self.destroy()
    
    def save_and_send(self):
        """Salva as configurações e inicia o processo de envio"""
        # Verifica se o parent tem um DataFrame com dados
        if not hasattr(self.parent, 'df') or self.parent.df is None:
            messagebox.showwarning("Aviso", "Carregue primeiro um CSV.")
            return
            
        # Verifica se temos uma coluna "email" para envio
        if 'email' not in self.parent.df.columns:
            messagebox.showwarning("Aviso", "O CSV precisa ter uma coluna 'email'.")
            return
            
        # Prepara a lista de destinatários para confirmação
        recipients_df = self.parent.df[['nome', 'email']]
        use_template = self.use_template_var.get()
        
        # Salvar configurações antes de criar a janela de confirmação
        email_config = {
            "use_template": self.use_template_var.get(),
            "template_path": self.template_path_var.get(),
            "subject": self.subject_var.get(),
            "attachment_name": self.attachment_name_var.get(),
            "cc": [email.strip() for email in self.cc_var.get().split(',') if email.strip()],
            "bcc": [email.strip() for email in self.bcc_var.get().split(',') if email.strip()]
        }
        
        # Salvar o conteúdo do corpo do email
        if hasattr(self.parent, 'email_text'):
            body_content = self.email_body_text.get("1.0", "end-1c")
            # Atualiza o app principal antes de fechar esta janela
            self.parent.email_text.delete("1.0", tk.END)
            self.parent.email_text.insert("1.0", body_content)
        
        # Atualiza o email_config no app principal
        self.parent.email_config = email_config
        
        # Fecha esta janela
        self.destroy()
        
        # Agora cria a janela de confirmação como filha da janela principal
        confirm_window = tk.Toplevel(self.parent)
        confirm_window.title("Confirmar Envio de Emails")
        confirm_window.geometry("500x400")
        confirm_window.transient(self.parent)
        confirm_window.grab_set()
        confirm_window.configure(bg='#303030')
        
        # Frame superior para informações
        info_frame = tk.Frame(confirm_window, bg='#303030', padx=10, pady=10)
        info_frame.pack(fill=tk.X)
        
        # Texto de informação
        tk.Label(info_frame, 
                text="Confirme os destinatários para envio dos certificados:", 
                bg='#303030', fg='white', font=('Arial', 11)).pack(anchor='w')
        
        tk.Label(info_frame, 
                text=f"Total: {len(recipients_df)} destinatários", 
                bg='#303030', fg='white').pack(anchor='w')
        
        # Template usado
        template_info = "Será usado template HTML" if use_template else "Será usado corpo de texto personalizado"
        tk.Label(info_frame, text=template_info, 
               bg='#303030', fg='white').pack(anchor='w')
        
        # Frame para a lista de destinatários
        list_frame = tk.Frame(confirm_window, bg='#303030', padx=10, pady=10)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # Lista de destinatários (Treeview)
        columns = ('nome', 'email')
        recipients_tree = ttk.Treeview(list_frame, columns=columns, show='headings')
        recipients_tree.heading('nome', text='Nome')
        recipients_tree.heading('email', text='Email')
        recipients_tree.column('nome', width=200)
        recipients_tree.column('email', width=250)
        
        # Scrollbar para a Treeview
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=recipients_tree.yview)
        recipients_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        recipients_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Preenche a Treeview com os dados
        for i, (_, row) in enumerate(recipients_df.iterrows()):
            recipients_tree.insert('', tk.END, values=(row['nome'], row['email']))
        
        # Frame para os botões
        buttons_frame = tk.Frame(confirm_window, bg='#303030', padx=10, pady=10)
        buttons_frame.pack(fill=tk.X)
        
        # Botões de confirmação e cancelamento
        cancel_button = tk.Button(buttons_frame, text="Cancelar", width=10, 
                               command=confirm_window.destroy,
                               bg='#444444', fg='white')
        cancel_button.pack(side=tk.LEFT, padx=10)
        
        confirm_button = tk.Button(buttons_frame, text="Enviar Emails", width=15, 
                                command=lambda: [confirm_window.destroy(), 
                                               self.parent.generate_all(from_config=True)],
                                bg='#4CAF50', fg='white', font=('Arial', 9, 'bold'))
        confirm_button.pack(side=tk.RIGHT, padx=10)
    
    def show_html_preview(self):
        """Mostra o preview HTML do template diretamente no navegador"""
        if not self.template_path_var.get():
            messagebox.showwarning("Aviso", "Selecione um template primeiro.")
            return
        
        try:
            # Verifica se o CSV foi carregado primeiro
            if hasattr(self.parent, 'df') and self.parent.df is not None and len(self.parent.df) > 0:
                # Usa o primeiro registro do CSV
                sample_data = self.parent.df.iloc[0].to_dict()
            else:
                # Dados de exemplo quando não há CSV carregado
                sample_data = {
                    "nome": "João Silva",
                    "email": "joao.silva@exemplo.com",
                    "evento": "Workshop de Programação",
                    "data": "01/01/2023",
                    "local": "Online",
                    "cargo": "Participante",
                    "empresa": "Empresa Exemplo"
                }

            # Configura o placeholder global do assunto para uso no template
            subject = self.subject_var.get()
            import config
            config.GLOBAL_PLACEHOLDERS['subject'] = subject
            
            # Usa get_all_placeholders para obter todos os placeholders formatados
            all_placeholders = self.parent.get_all_placeholders(sample_data)
            
            # Carrega o template HTML
            with open(self.template_path_var.get(), 'r', encoding='utf-8') as f:
                body_html_raw = f.read()
                
            # Tenta formatar o HTML com os dados de exemplo usando todos os placeholders
            try:
                formatted_body = self._format_text_with_data(body_html_raw, all_placeholders)
            except Exception as e:
                formatted_body = body_html_raw
                logging.error(f"Erro ao formatar HTML: {e}")
                
            # Prepara o HTML com o CSS para melhor visualização
            full_html = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <style>
                    body {{
                        font-family: Arial, sans-serif;
                        margin: 0;
                        padding: 20px;
                        background-color: #f8f8f8;
                    }}
                    .content {{
                        background-color: white;
                        padding: 15px;
                        border: 1px solid #ddd;
                        border-radius: 4px;
                    }}
                </style>
            </head>
            <body>
                <div class="content">
                    {formatted_body}
                </div>
            </body>
            </html>
            """
            
            # Salva o HTML em um arquivo temporário
            self.temp_html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                             "temp", "preview.html")
            os.makedirs(os.path.dirname(self.temp_html_path), exist_ok=True)
            
            with open(self.temp_html_path, "w", encoding="utf-8") as f:
                f.write(full_html)
            
            # Abre o arquivo no navegador padrão
            import webbrowser
            webbrowser.open(self.temp_html_path)
            
        except Exception as e:
            logging.error(f"Erro ao carregar o template HTML: {str(e)}")


class ConfigEditor(tk.Toplevel):
    """Janela de editor de configurações (texto simples)."""
    
    def __init__(self, parent):
        super().__init__(parent)
        
        self.title("Editor de Configurações")
        self.geometry("800x600")
        self.configure(bg='#303030')
        
        # Caminho do arquivo config.py
        self.config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.py")
        
        # Frame superior com informações
        info_frame = tk.Frame(self, bg='#303030')
        info_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(info_frame, text=f"Editando: {self.config_path}", 
                bg='#303030', fg='#aaaaaa', font=('Arial', 9)).pack(anchor='w')
        
        # Cria o editor de texto
        text_frame = tk.Frame(self, bg='#3a3a3a')
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Barra de scroll
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Text widget com tema escuro
        self.text_widget = tk.Text(text_frame, 
                                  bg='#1e1e1e', 
                                  fg='#e0e0e0',
                                  font=('Courier New', 10),
                                  wrap=tk.NONE,
                                  yscrollcommand=scrollbar.set,
                                  insertbackground='#e0e0e0',
                                  selectbackground='#0e639c',
                                  selectforeground='#e0e0e0')
        self.text_widget.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.text_widget.yview)
        
        # Carrega o conteúdo do arquivo config.py
        self._load_config()
        
        # Frame inferior com botões
        button_frame = tk.Frame(self, bg='#303030')
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Button(button_frame, text="Guardar", command=self._save_config,
                 bg='#2196F3', fg='white', font=('Arial', 9, 'bold'),
                 width=15).pack(side=tk.RIGHT, padx=5)
        
        tk.Button(button_frame, text="Cancelar", command=self.destroy,
                 bg='#444444', fg='white', font=('Arial', 9),
                 width=15).pack(side=tk.RIGHT, padx=5)
        
        tk.Button(button_frame, text="Recarregar", command=self._load_config,
                 bg='#444444', fg='white', font=('Arial', 9),
                 width=15).pack(side=tk.LEFT, padx=5)
    
    def _load_config(self):
        """Carrega o conteúdo do arquivo config.py."""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            self.text_widget.delete('1.0', tk.END)
            self.text_widget.insert('1.0', content)
            self.text_widget.edit_reset()
            
            logging.info("Config.py carregado no editor")
        except FileNotFoundError:
            messagebox.showerror("Erro", f"Arquivo não encontrado: {self.config_path}")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar config.py:\n{str(e)}")
            self.destroy()
    
    def _save_config(self):
        """Guarda as alterações ao arquivo config.py."""
        try:
            content = self.text_widget.get('1.0', tk.END)
            
            with open(self.config_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            messagebox.showinfo("Sucesso", "Configurações guardadas com sucesso!\n\nNota: Você pode precisar reiniciar a aplicação para que as alterações façam efeito.")
            logging.info("Config.py guardado com sucesso")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao guardar config.py:\n{str(e)}")
            logging.error(f"Erro ao guardar config.py: {str(e)}")


if __name__ == "__main__":
    # Mensagem de início
    logging.info("Iniciando VisuMaker...")
    # Se estiver em modo verbose, mostra mais informações
    if args.verbose:
        logging.info("Modo verbose ativado - logs detalhados serão mostrados")
    app = App()

    # Lógica simples: se argumentos -csv ou -proj forem passados, preenche os campos na UI
    def fill_cli_loaded_files():
        # Carrega layout/projeto se -proj foi passado
        if hasattr(args, 'proj') and args.proj:
            try:
                # Usa a função load_layout diretamente para garantir consistência
                app.load_layout(path=args.proj)
            except Exception as e:
                logging.error(f"Erro ao carregar projeto via argumento: {e}")
        
        # Preenche CSV se -csv foi passado
        if hasattr(app, 'df') and hasattr(args, 'csv') and args.csv:
            try:
                import pandas as pd
                app.df = pd.read_csv(args.csv)
                if hasattr(app, 'forget_csv_btn'):
                    app.forget_csv_btn.config(state=tk.NORMAL)
                if hasattr(app, 'update_preview'):
                    app.update_preview()
            except Exception as e:
                logging.error(f"Erro ao carregar CSV via argumento: {e}")

    # Executa após a UI estar pronta
    app.after(100, fill_cli_loaded_files)
    app.mainloop()
