import os
import sys
import configparser
from threading import Thread, Lock
from queue import Queue
from tkinter import Menu, NW, WORD, END
from tkinter.filedialog import askopenfilename
from tkinter.scrolledtext import ScrolledText  # Usado para a documentação e mensagem
from PIL import Image, ImageDraw, ImageFont, ImageTk
import pandas as pd
import yagmail
from tkinter import PhotoImage

# Importa ttkbootstrap para a interface moderna
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

class EditCertificate:
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config_path = 'config.ini'
        self.sending = False
        self.stop_requested = False
        self.check_and_create_config()
        self.load_font_families()
        # Define o tema inicial ("darkly" para tema escuro e "flatly" para claro)
        self.current_theme = "darkly"
        self.create_gui()
        self.create_menu()

    def check_and_create_config(self):
        if not os.path.exists(self.config_path):
            self.config['credentials'] = {'email': 'seu_email@gmail.com', 'password': 'sua_senha'}
            with open(self.config_path, 'w') as configfile:
                self.config.write(configfile)
        else:
            self.config.read(self.config_path)

    def load_font_families(self):
        fonts_dir = "C:\\Windows\\Fonts"
        self.font_families = {}
        for root, dirs, files in os.walk(fonts_dir):
            for file in files:
                if file.lower().endswith(('.ttf', '.otf')):
                    self.font_families[os.path.splitext(file)[0]] = file

    def create_menu(self):
        # Usa o Menu nativo do tkinter para a barra de menus
        menu_bar = Menu(self.window, tearoff=0)
        # Menu Arquivo
        file_menu = Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Sair", command=self.window.quit)
        menu_bar.add_cascade(label="📁 Arquivo", menu=file_menu)
        # Menu Ajuda
        help_menu = Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="Documentação", command=self.show_documentation)
        help_menu.add_command(label="Sobre", command=self.show_about)
        menu_bar.add_cascade(label="❓ Ajuda", menu=help_menu)
        # Menu Tema
        theme_menu = Menu(menu_bar, tearoff=0)
        theme_menu.add_command(label="🌙 Tema Escuro", command=lambda: self.toggle_theme("darkly"))
        theme_menu.add_command(label="☀️ Tema Claro", command=lambda: self.toggle_theme("flatly"))
        menu_bar.add_cascade(label="🎨 Tema", menu=theme_menu)
        self.window.config(menu=menu_bar)

    def toggle_theme(self, theme_name):
        self.current_theme = theme_name
        self.window.style.theme_use(theme_name)

    def show_documentation(self):
        docs = ttk.Toplevel(self.window)
        docs.title("Documentação")
        docs.geometry("600x400")
        docs.configure(bg=self.window.style.colors.bg)
        text = ScrolledText(docs, wrap=WORD, font=("Segoe UI", 12))
        text.pack(expand=True, fill="both", padx=10, pady=10)
        text.insert(END, """📚 Guia Rápido:
1. Selecione o modelo do certificado (PNG, JPG ou JPEG);
2. Carregue a planilha com os dados;
3. Configure a posição e o estilo do texto;
4. Envie os certificados por email.

📌 Dicas:
- Utilize coordenadas X/Y para posicionar o texto;
- Teste diferentes fontes e tamanhos;
- Salve suas configurações com frequência.
""")
        text.configure(state="disabled")

    def show_about(self):
        about = ttk.Toplevel(self.window)
        about.title("Sobre")
        about.geometry("300x200")
        about.configure(bg=self.window.style.colors.bg)
        ttk.Label(about,
                  text="Gerador de Certificados\n\nVersão 3.0\nDesenvolvido por Gabriel Batista\n© 2025",
                  font=("Segoe UI", 12),
                  bootstyle=INFO
                  ).pack(expand=True, padx=10, pady=10)

    def create_gui(self):
        # Cria a janela principal usando ttkbootstrap
        self.window = ttk.Window(themename=self.current_theme)
        self.window.title("Gerador de Certificados")
        def resource_path(relative_path):
            """Funciona tanto com .py quanto com .exe"""
            base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            return os.path.join(base_path, relative_path)
        icon_path = resource_path("icon.ico")
        self.window.iconbitmap(icon_path)

        # --- Cálculo do fator de redimensionamento ---
        # Resolução base (padrão): 1920 x 1080
        base_width, base_height = 1920, 1080
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        # Se a tela for menor que a base, calcula o fator; caso contrário, usa 1
        self.resize_factor = min(1, screen_width / base_width, screen_height / base_height)
        default_width, default_height = 1200, 800
        self.window.state('zoomed')
        width = int(default_width * self.resize_factor)
        height = int(default_height * self.resize_factor)
        self.window.geometry(f"{width}x{height}")
        # ------------------------------------------------

        # Cabeçalho
        header = ttk.Frame(self.window)
        header.pack(side=TOP, fill="x", padx=20, pady=10)
        ttk.Label(header, text="Gerador de Certificados", font=("Segoe UI", 20, "bold")).pack(side="left")

        # Container principal: painel lateral e área de pré-visualização
        main_frame = ttk.Frame(self.window)
        main_frame.pack(side=TOP, fill="both", expand=True, padx=20, pady=10)

        # Painel lateral esquerdo – Notebook com as _tabs_
        sidebar = ttk.Frame(main_frame, width=400)
        sidebar.pack(side="left", fill="y", padx=(0, 10))
        self.notebook = ttk.Notebook(sidebar)
        self.tab_config = ttk.Frame(self.notebook)
        self.tab_design = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_config, text="Configurações")
        self.notebook.add(self.tab_design, text="Design")
        self.notebook.pack(fill="both", expand=True)

        # Área de pré-visualização (lado direito)
        preview_container = ttk.Frame(main_frame)
        preview_container.pack(side="left", fill="both", expand=True)
        self.preview_canvas = self.create_preview_canvas(preview_container)

        # Área inferior: barra de progresso, status e botões de ação
        bottom = ttk.Frame(self.window)
        bottom.pack(side="bottom", fill="x", padx=20, pady=10)
        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=5)
        status_frame = ttk.Frame(bottom)
        status_frame.pack(fill="x")
        self.status_bar = ttk.Label(status_frame, text="Pronto")
        self.status_bar.pack(side="left")
        button_frame = ttk.Frame(bottom)
        button_frame.pack(fill="x", pady=5)
        self.stop_button = ttk.Button(button_frame, text="⏹️ Parar", command=self.stop_sending, bootstyle=DANGER, state="disabled")
        self.stop_button.pack(side="right", padx=5)
        self.send_button = ttk.Button(button_frame, text="📧 Enviar Certificados", command=self.start_sending, bootstyle=SUCCESS)
        self.send_button.pack(side="right", padx=5)

        # Cria o conteúdo dos _tabs_
        self.create_config_tab(self.tab_config)
        self.create_design_tab(self.tab_design)

        # Atualiza o canvas e outros elementos sempre que a janela for redimensionada
        self.window.bind("<Configure>", self.on_resize)

    def create_preview_canvas(self, parent):
        # Usa um Canvas do tkinter para a pré-visualização
        from tkinter import Canvas
        canvas = Canvas(parent, bg="white", bd=0, highlightthickness=0)
        canvas.pack(fill="both", expand=True, padx=10, pady=10)
        return canvas

    def on_resize(self, event):
        # Atualiza o tamanho do canvas de pré-visualização
        if hasattr(self, 'preview_canvas'):
            canvas_width = max(self.window.winfo_width() // 2, 400)
            canvas_height = max(self.window.winfo_height() // 2, 300)
            self.preview_canvas.config(width=canvas_width, height=canvas_height)

            # Se houver um modelo selecionado, atualiza a pré-visualização em tempo real
            if hasattr(self, 'template_path') and self.template_path:
                self.preview_certificate()

            # Atualiza o tamanho das fontes do menu (e de outros widgets, se desejar)
            # O fator de escala é baseado na largura atual da janela comparada ao tamanho padrão (1200)
            factor = min(1, self.window.winfo_width() / 1200)
            self.window.option_add("*Menu.font", ("Segoe UI", int(10 * factor)))
            # Se necessário, atualize também outros estilos via self.window.style.configure(...)

    def create_config_tab(self, parent):
        # --- Configurações de Email ---
        email_frame = ttk.Labelframe(parent, text="📧 Configurações de Email", padding=10)
        email_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(email_frame, text="Email:", font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.email_entry = ttk.Entry(email_frame, font=("Segoe UI", 10))
        self.email_entry.insert(0, self.config.get('credentials', 'email'))
        self.email_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ttk.Label(email_frame, text="Senha:", font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.password_entry = ttk.Entry(email_frame, show="*", font=("Segoe UI", 10))
        self.password_entry.insert(0, self.config.get('credentials', 'password'))
        self.password_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        self.save_button = ttk.Button(email_frame, text="Salvar Credenciais", command=self.save_config, bootstyle=PRIMARY)
        self.save_button.grid(row=2, column=0, columnspan=2, pady=10, padx=5)
        email_frame.columnconfigure(1, weight=1)

        # --- Dados e Conteúdo ---
        data_frame = ttk.Labelframe(parent, text="📁 Dados e Conteúdo", padding=10)
        data_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(data_frame, text="Planilha:", font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.data_button = ttk.Button(data_frame, text="Selecionar Excel", command=self.select_data_file, bootstyle=INFO)
        self.data_button.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ttk.Label(data_frame, text="Nome Base:", font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.output_name_entry = ttk.Entry(data_frame, font=("Segoe UI", 10))
        self.output_name_entry.insert(0, 'certificado')
        self.output_name_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        ttk.Label(data_frame, text="Assunto:", font=("Segoe UI", 10)).grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.subject_entry = ttk.Entry(data_frame, font=("Segoe UI", 10))
        self.subject_entry.insert(0, 'Seu Certificado Está Pronto!')
        self.subject_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        ttk.Label(data_frame, text="Mensagem:", font=("Segoe UI", 10)).grid(row=3, column=0, sticky="nw", padx=5, pady=5)
        self.content_text = ScrolledText(data_frame, height=6, wrap=WORD, font=("Segoe UI", 10))
        self.content_text.insert(END, 'Prezado {name},\n\nSegue seu certificado em anexo.\n\nAtenciosamente,\nEquipe ...')
        self.content_text.grid(row=3, column=1, sticky="ew", padx=5, pady=5)
        data_frame.columnconfigure(1, weight=1)

    def create_design_tab(self, parent):
        # --- Posicionamento ---
        pos_frame = ttk.Labelframe(parent, text="📍 Posicionamento", padding=10)
        pos_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(pos_frame, text="Nome (X,Y):", font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.x_entry = ttk.Entry(pos_frame, width=8, font=("Segoe UI", 10))
        self.x_entry.insert(0, '200')
        self.x_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        self.y_entry = ttk.Entry(pos_frame, width=8, font=("Segoe UI", 10))
        self.y_entry.insert(0, '1340')
        self.y_entry.grid(row=0, column=2, sticky="w", padx=5, pady=5)
        ttk.Label(pos_frame, text="Número (X,Y):", font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.cert_x_entry = ttk.Entry(pos_frame, width=8, font=("Segoe UI", 10))
        self.cert_x_entry.insert(0, '570')
        self.cert_x_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        self.cert_y_entry = ttk.Entry(pos_frame, width=8, font=("Segoe UI", 10))
        self.cert_y_entry.insert(0, '1930')
        self.cert_y_entry.grid(row=1, column=2, sticky="w", padx=5, pady=5)

        # --- Configurações de Fonte ---
        font_frame = ttk.Labelframe(parent, text="🔠 Configurações de Fonte", padding=10)
        font_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(font_frame, text="Fonte Principal:", font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.font_combobox = ttk.Combobox(font_frame, values=list(self.font_families.keys()), font=("Segoe UI", 10))
        self.font_combobox.set('times')
        self.font_combobox.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ttk.Label(font_frame, text="Tamanho Texto:", font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.font_size_entry = ttk.Entry(font_frame, width=8, font=("Segoe UI", 10))
        self.font_size_entry.insert(0, '100')
        self.font_size_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(font_frame, text="Tamanho Número:", font=("Segoe UI", 10)).grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.font_cert_size_entry = ttk.Entry(font_frame, width=8, font=("Segoe UI", 10))
        self.font_cert_size_entry.insert(0, '60')
        self.font_cert_size_entry.grid(row=2, column=1, sticky="w", padx=5, pady=5)

        # --- Modelo do Certificado ---
        temp_frame = ttk.Labelframe(parent, text="🎨 Modelo do Certificado", padding=10)
        temp_frame.pack(fill="x", padx=10, pady=5)
        self.template_button = ttk.Button(temp_frame, text="Selecionar Modelo", command=self.select_template, bootstyle=INFO)
        self.template_button.pack(padx=5, pady=5)

        # Atualiza a pré-visualização em tempo real ao alterar os parâmetros
        for widget in (self.x_entry, self.y_entry, self.cert_x_entry, self.cert_y_entry,
                       self.font_combobox, self.font_size_entry, self.font_cert_size_entry):
            widget.bind("<KeyRelease>", lambda e: self.preview_certificate())
        self.font_combobox.bind("<<ComboboxSelected>>", lambda e: self.preview_certificate())

    def select_template(self):
        self.template_path = askopenfilename(
            title="Selecionar Modelo de Certificado",
            filetypes=[("Arquivos Suportados", "*.png *.jpg *.jpeg")]
        )
        if self.template_path:
            self.status_bar.config(text=f"Modelo selecionado: {os.path.basename(self.template_path)}")
            self.preview_certificate()

    def preview_certificate(self):
        try:
            if not hasattr(self, 'template_path') or not self.template_path:
                raise ValueError("Selecione um modelo de certificado!")
            name = "Pré-visualização"
            certificate_number = "0000"
            certificate_number_padded = certificate_number.zfill(4)
            preview_image = Image.open(self.template_path)
            draw = ImageDraw.Draw(preview_image)
            font_name = self.font_combobox.get()
            font_size = int(self.font_size_entry.get())
            font_path = self.get_font_path(font_name)
            font = ImageFont.truetype(font_path, font_size)
            x = int(self.x_entry.get())
            y = int(self.y_entry.get())
            draw.text((x, y), name, font=font, fill=(0, 0, 0))
            font_cert_size = int(self.font_cert_size_entry.get())
            font_cert = ImageFont.truetype(font_path, font_cert_size)
            cert_x = int(self.cert_x_entry.get())
            cert_y = int(self.cert_y_entry.get())
            draw.text((cert_x, cert_y), certificate_number_padded, font=font_cert, fill=(0, 0, 0))
            
            # Obtém as dimensões atuais do canvas para limitar o tamanho da imagem
            canvas_width = self.preview_canvas.winfo_width() or 800
            canvas_height = self.preview_canvas.winfo_height() or 600
            preview_image.thumbnail((canvas_width, canvas_height))
            
            img_tk = ImageTk.PhotoImage(preview_image)
            self.preview_canvas.delete("all")
            self.preview_canvas.config(width=preview_image.width, height=preview_image.height)
            self.preview_canvas.create_image(0, 0, anchor=NW, image=img_tk)
            self.preview_canvas.image = img_tk
        except Exception as e:
            self.shake_window()
            self.status_bar.config(text=f"Erro na pré-visualização: {str(e)}")

    def get_font_path(self, font_name):
        fonts_dir = "C:\\Windows\\Fonts"
        for root, dirs, files in os.walk(fonts_dir):
            for file in files:
                if font_name.lower() in file.lower():
                    return os.path.join(root, file)
        raise FileNotFoundError(f"Fonte '{font_name}' não encontrada no diretório de fontes do Windows.")

    def select_data_file(self):
        self.data_path = askopenfilename(
            title="Selecionar Arquivo de Dados", 
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        if self.data_path:
            self.status_bar.config(text=f"Planilha selecionada: {os.path.basename(self.data_path)}")

    def save_config(self):
        self.config['credentials']['email'] = self.email_entry.get()
        self.config['credentials']['password'] = self.password_entry.get()
        with open(self.config_path, 'w') as configfile:
            self.config.write(configfile)
        self.status_bar.config(text="Credenciais salvas com sucesso!")
        self.animate_success()

    def start_sending(self):
        if self.sending:
            return
        self.sending = True
        self.stop_requested = False
        self.send_button.config(state="disabled")
        self.stop_button.config(state="normal")
        Thread(target=self.send_emails_in_parallel, daemon=True).start()

    def stop_sending(self):
        self.stop_requested = True
        self.sending = False
        self.send_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.status_bar.config(text="Envio interrompido pelo usuário.")

    def send_emails_in_parallel(self):
        try:
            sender_email = self.email_entry.get()
            sender_password = self.password_entry.get()
            subject = self.subject_entry.get()
            content = self.content_text.get(1.0, END)
            df = pd.read_excel(self.data_path)
            total = len(df)
            queue = Queue()
            lock = Lock()
            threads = []
            max_threads = 5  # número máximo de threads simultâneas

            for index, row in df.iterrows():
                queue.put((index, row))

            def worker():
                while not queue.empty() and not self.stop_requested:
                    try:
                        index, row = queue.get_nowait()
                        name = row['Nome']
                        email = row['Email']
                        certificate_number = row['Numero do Certificado']
                        certificate_image = self.create_certificate(name, certificate_number, self.template_path)
                        self.send_email_generic(name, email, certificate_image, sender_email, sender_password, subject, content)
                        progress = int((total - queue.qsize()) / total * 100)
                        self.window.after(0, lambda p=progress: self.update_progress(p))
                        queue.task_done()
                    except Exception as e:
                        with lock:
                            print(f"Erro ao enviar email para {email}: {str(e)}")
                            self.status_bar.config(text=f"Erro ao enviar para {email}: {str(e)}")
                            self.animate_error()

            for _ in range(max_threads):
                thread = Thread(target=worker, daemon=True)
                thread.start()
                threads.append(thread)

            for thread in threads:
                thread.join()

            if not self.stop_requested:
                self.window.after(0, lambda: self.status_bar.config(text="Todos os certificados foram enviados!"))
                self.animate_success()
            self.sending = False
            self.window.after(0, lambda: self.send_button.config(state="normal"))
            self.window.after(0, lambda: self.stop_button.config(state="disabled"))
        except Exception as e:
            self.shake_window()
            self.status_bar.config(text=f"Erro: {str(e)}")
            self.animate_error()

    def create_certificate(self, name, certificate_number, template_path):
        template_extension = os.path.splitext(template_path)[1].lower()
        certificate_number = str(certificate_number)
        certificate_number_padded = certificate_number.zfill(max(4, len(certificate_number)))
        imagem = Image.open(template_path)
        draw = ImageDraw.Draw(imagem)
        font_name = self.font_combobox.get()
        font_size = int(self.font_size_entry.get())
        font_path = self.get_font_path(font_name)
        font = ImageFont.truetype(font_path, font_size)
        x = int(self.x_entry.get())
        y = int(self.y_entry.get())
        draw.text((x, y), name, font=font, fill=(0, 0, 0))
        font_cert_size = int(self.font_cert_size_entry.get())
        font_cert = ImageFont.truetype(font_path, font_cert_size)
        cert_x = int(self.cert_x_entry.get())
        cert_y = int(self.cert_y_entry.get())
        draw.text((cert_x, cert_y), certificate_number_padded, font=font_cert, fill=(0, 0, 0))
        certificate_image = f"{self.output_name_entry.get()}_{name}.pdf"
        imagem.save(certificate_image)
        return certificate_image

    def send_email_generic(self, name, email, certificate_image, sender_email, sender_password, subject, content):
        try:
            usuario = yagmail.SMTP(user=sender_email, password=sender_password)
            personalized_content = content.replace("{name}", name)
            usuario.send(to=email, subject=subject, contents=personalized_content, attachments=certificate_image)
            print(f'Email enviado para {email} com sucesso!')
        except Exception as e:
            raise RuntimeError(f"Erro ao enviar email para {email}: {str(e)}")

    def update_progress(self, value):
        self.progress['value'] = value
        self.status_bar.config(text=f"Enviando... {value}% Completo")

    def shake_window(self):
        # Efeito de "tremor" na janela para sinalizar erro
        x = self.window.winfo_x()
        y = self.window.winfo_y()
        for _ in range(3):
            for dx in [10, -20, 15, -10, 5, -5, 2, -2]:
                self.window.geometry(f"+{x + dx}+{y}")
                self.window.update()
                self.window.after(20)
            self.window.geometry(f"+{x}+{y}")

    def animate_success(self):
        original_color = self.status_bar.cget('foreground')
        self.status_bar.config(foreground="green")
        self.window.after(2000, lambda: self.status_bar.config(foreground=original_color))

    def animate_error(self):
        original_color = self.status_bar.cget('foreground')
        self.status_bar.config(foreground="red")
        self.window.after(2000, lambda: self.status_bar.config(foreground=original_color))


if __name__ == '__main__':
    app = EditCertificate()
    app.window.mainloop()