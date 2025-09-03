#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Processador de Etiquetas Shopee - Versão Corrigida
Recorta a etiqueta da primeira página e coloca horizontalmente abaixo da declaração de conteúdo
"""

import os
import sys
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import fitz  # PyMuPDF

class ShopeeEtiquetaProcessor:
    def __init__(self):
        self.setup_gui()
        
    def setup_gui(self):
        """Configura a interface gráfica"""
        self.root = tk.Tk()
        self.root.title("Processador de Etiquetas Shopee - v2.0")
        self.root.geometry("550x450")
        self.root.resizable(True, True)
        
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        title_label = ttk.Label(main_frame, text="Processador de Etiquetas Shopee", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Instruções
        instructions = ttk.Label(main_frame, 
                               text="Este programa irá:\n" +
                                    "1. Pegar a etiqueta completa da primeira página\n" +
                                    "2. Colocar ela na horizontal na parte inferior\n" +
                                    "3. Manter a declaração de conteúdo na parte superior\n" +
                                    "4. Gerar um PDF otimizado para impressão",
                               justify=tk.LEFT)
        instructions.grid(row=1, column=0, columnspan=2, pady=(0, 20), sticky=tk.W)
        
        # Seleção de arquivo
        ttk.Label(main_frame, text="Arquivo PDF da Shopee:").grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        
        self.file_path_var = tk.StringVar()
        file_frame = ttk.Frame(main_frame)
        file_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=50)
        self.file_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(file_frame, text="Selecionar", command=self.select_file).grid(row=0, column=1)
        
        # Modo de processamento
        mode_frame = ttk.LabelFrame(main_frame, text="Modo de Processamento", padding="10")
        mode_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        
        self.mode_var = tk.StringVar(value="auto")
        ttk.Radiobutton(mode_frame, text="Automático (recomendado)", 
                       variable=self.mode_var, value="auto").grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(mode_frame, text="Manual (ajustar coordenadas)", 
                       variable=self.mode_var, value="manual").grid(row=1, column=0, sticky=tk.W)
        
        # Configurações manuais (inicialmente ocultas)
        self.config_frame = ttk.LabelFrame(main_frame, text="Coordenadas Manuais", padding="10")
        
        ttk.Label(self.config_frame, text="Área da etiqueta (x1, y1, x2, y2):").grid(row=0, column=0, sticky=tk.W)
        
        coords_frame = ttk.Frame(self.config_frame)
        coords_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        self.x1_var = tk.StringVar(value="0")
        self.y1_var = tk.StringVar(value="140")
        self.x2_var = tk.StringVar(value="450")
        self.y2_var = tk.StringVar(value="330")
        
        ttk.Label(coords_frame, text="x1:").grid(row=0, column=0)
        ttk.Entry(coords_frame, textvariable=self.x1_var, width=8).grid(row=0, column=1, padx=2)
        ttk.Label(coords_frame, text="y1:").grid(row=0, column=2)
        ttk.Entry(coords_frame, textvariable=self.y1_var, width=8).grid(row=0, column=3, padx=2)
        ttk.Label(coords_frame, text="x2:").grid(row=0, column=4)
        ttk.Entry(coords_frame, textvariable=self.x2_var, width=8).grid(row=0, column=5, padx=2)
        ttk.Label(coords_frame, text="y2:").grid(row=0, column=6)
        ttk.Entry(coords_frame, textvariable=self.y2_var, width=8).grid(row=0, column=7, padx=2)
        
        # Botão mostrar/ocultar configurações
        self.toggle_button = ttk.Button(main_frame, text="Mostrar Configurações Manuais", 
                                       command=self.toggle_config)
        self.toggle_button.grid(row=5, column=0, columnspan=2, pady=(0, 10))
        
        # Botão processar
        self.process_button = ttk.Button(main_frame, text="Processar PDF", 
                                       command=self.start_processing, style="Accent.TButton")
        self.process_button.grid(row=7, column=0, columnspan=2, pady=20)
        
        # Barra de progresso
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=8, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Status
        self.status_var = tk.StringVar(value="Pronto para processar")
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=9, column=0, columnspan=2)
        
        # Configurar redimensionamento
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        file_frame.columnconfigure(0, weight=1)
        
        # Configurações ocultas inicialmente
        self.config_visible = False
        
    def toggle_config(self):
        """Mostra/oculta as configurações manuais"""
        if self.config_visible:
            self.config_frame.grid_remove()
            self.toggle_button.config(text="Mostrar Configurações Manuais")
            self.config_visible = False
        else:
            self.config_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
            self.toggle_button.config(text="Ocultar Configurações Manuais")
            self.config_visible = True
    
    def select_file(self):
        """Seleciona o arquivo PDF"""
        file_path = filedialog.askopenfilename(
            title="Selecione o PDF da Shopee",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path_var.set(file_path)
    
    def start_processing(self):
        """Inicia o processamento em thread separada"""
        if not self.file_path_var.get():
            messagebox.showerror("Erro", "Por favor, selecione um arquivo PDF")
            return
            
        self.process_button.config(state='disabled')
        self.progress.start()
        self.status_var.set("Processando...")
        
        # Executar em thread separada para não travar a interface
        threading.Thread(target=self.process_pdf, daemon=True).start()
    
    def process_pdf(self):
        """Processa o PDF"""
        try:
            input_path = self.file_path_var.get()
            
            # Gerar nome do arquivo de saída
            input_file = Path(input_path)
            output_path = input_file.parent / f"{input_file.stem}_processado.pdf"
            
            # Abrir o documento PDF
            doc = fitz.open(input_path)
            
            if len(doc) < 2:
                raise Exception("O PDF deve ter pelo menos 2 páginas")
            
            # Obter as páginas
            pagina_etiqueta = doc[0]  # Primeira página (etiqueta)
            pagina_declaracao = doc[1]  # Segunda página (declaração)
            
            # Criar novo documento
            novo_doc = fitz.open()
            
            if self.mode_var.get() == "auto":
                # Modo automático - usa toda a área útil da etiqueta
                self.processar_automatico(novo_doc, doc, pagina_etiqueta, pagina_declaracao)
            else:
                # Modo manual - usa coordenadas especificadas
                x1 = float(self.x1_var.get())
                y1 = float(self.y1_var.get())
                x2 = float(self.x2_var.get())
                y2 = float(self.y2_var.get())
                self.processar_manual(novo_doc, doc, pagina_etiqueta, pagina_declaracao, x1, y1, x2, y2)
            
            # Salvar o documento
            novo_doc.save(str(output_path))
            novo_doc.close()
            doc.close()
            
            # Atualizar interface na thread principal
            self.root.after(0, self.processing_complete, str(output_path))
            
        except Exception as e:
            self.root.after(0, self.processing_error, str(e))
            
    def _bbox_conteudo(self, pagina, pad=2):
        """
        Retorna o retângulo que envolve o conteúdo de texto da página,
        eliminando grandes áreas em branco das bordas. 'pad' adiciona
        uma borda de segurança em pontos.
        """
        try:
            blocks = pagina.get_text("blocks")  # [(x0,y0,x1,y1,text, ...), ...]
            xs, ys, xe, ye = [], [], [], []
            for b in blocks:
                x0, y0, x1, y1 = b[:4]
                # ignora blocos absurdamente pequenos (ruídos)
                if (x1 - x0) > 4 and (y1 - y0) > 4:
                    xs.append(x0); ys.append(y0); xe.append(x1); ye.append(y1)
            if not xs:
                return pagina.rect  # fallback
            r = fitz.Rect(min(xs), min(ys), max(xe), max(ye))
            # aplica "pad" sem passar dos limites da página
            r.x0 = max(0, r.x0 - pad)
            r.y0 = max(0, r.y0 - pad)
            r.x1 = min(pagina.rect.x1, r.x1 + pad)
            r.y1 = min(pagina.rect.y1, r.y1 + pad)
            return r
        except Exception:
            return pagina.rect  # fallback robusto
    
    
    def processar_automatico(self, novo_doc, doc, pagina_etiqueta, pagina_declaracao):
        """
        Declaração no topo (sem deformar, mantendo proporção pelo recorte de conteúdo)
        e etiqueta logo abaixo, colada com gap mínimo, também mantendo proporção.
        """
        # A4 em pontos (72 dpi)
        A4_W, A4_H = 595, 842

        # Margens e espaçamentos (ajuste à vontade)
        MARGEM_LR = 6     # esquerda/direita
        MARGEM_SUP = 6    # topo
        GAP = 6            # espaço entre declaração e etiqueta
        MARGEM_INF = 6    # margem inferior mínima

        nova_pagina = novo_doc.new_page(width=A4_W, height=A4_H)

        # --- DECLARAÇÃO (página 2) ---
        bbox_decl = self._bbox_conteudo(pagina_declaracao, pad=8)
        largura_dest = A4_W - 2 * MARGEM_LR
        # escala para ocupar toda a largura útil mantendo proporção do recorte
        escala_decl = largura_dest / bbox_decl.width
        altura_decl = bbox_decl.height * escala_decl

        # destino da declaração no topo, sem centralizar verticalmente
        decl_x0 = MARGEM_LR
        decl_y0 = MARGEM_SUP
        decl_rect = fitz.Rect(decl_x0, decl_y0, decl_x0 + largura_dest, decl_y0 + altura_decl)

        # insere a declaração recortada (sem rotacionar)
        nova_pagina.show_pdf_page(decl_rect, doc, 1, clip=bbox_decl)

        # --- ETIQUETA (página 1) ---
        bbox_tag = self._bbox_conteudo(pagina_etiqueta, pad=5)

        # Após rotacionar 90°, as dimensões se trocam
        tag_w_rot = bbox_tag.height
        tag_h_rot = bbox_tag.width

        # Espaço disponível logo abaixo da declaração (colado)
        y_inicio_tag = decl_rect.y1 + GAP
        altura_disp = A4_H - y_inicio_tag - MARGEM_INF
        largura_disp = largura_dest

        # Escala para caber no espaço restante (sem sobrar folga desnecessária)
        escala_tag = min(largura_disp / tag_w_rot, altura_disp / tag_h_rot)
        # Evita expandir demais caso o conteúdo já caiba
        escala_tag = min(escala_tag, 1.0)

        tag_w_final = tag_w_rot * escala_tag
        tag_h_final = tag_h_rot * escala_tag

        # Centraliza horizontalmente, mas gruda verticalmente abaixo da declaração
        tag_x0 = MARGEM_LR + (largura_disp - tag_w_final) / 2
        tag_y0 = y_inicio_tag
        tag_rect = fitz.Rect(tag_x0, tag_y0, tag_x0 + tag_w_final, tag_y0 + tag_h_final)

        # Insere etiqueta recortada e rotacionada (90°)
        nova_pagina.show_pdf_page(tag_rect, doc, 0, clip=bbox_tag, rotate=90)   

    
    def processar_manual(self, novo_doc, doc, pagina_etiqueta, pagina_declaracao, x1, y1, x2, y2):
        """Processamento manual com coordenadas especificadas"""
        # Criar nova página A4
        nova_pagina = novo_doc.new_page(width=595, height=842)
        
        # Colocar a declaração na parte superior
        altura_declaracao = 842 * 0.60
        rect_declaracao = fitz.Rect(0, 0, 595, altura_declaracao)
        nova_pagina.show_pdf_page(rect_declaracao, doc, 1)
        
        # Usar coordenadas especificadas
        etiqueta_rect = fitz.Rect(x1, y1, x2, y2)
        
        # Calcular posição para a etiqueta
        etiqueta_largura = x2 - x1
        etiqueta_altura = y2 - y1
        
        # Dimensões após rotação
        largura_final = etiqueta_altura * 0.8
        altura_final = etiqueta_largura * 0.8
        
        # Posicionar com aproveitamento máximo do espaço
        pos_x = 15
        pos_y = altura_declaracao + 15
        
        dest_rect = fitz.Rect(pos_x, pos_y, pos_x + largura_final, pos_y + altura_final)
        
        # Inserir a etiqueta rotacionada
        nova_pagina.show_pdf_page(dest_rect, doc, 0, clip=etiqueta_rect, rotate=90)
    
    def processing_complete(self, output_path):
        """Chamado quando o processamento é concluído"""
        self.progress.stop()
        self.process_button.config(state='normal')
        self.status_var.set("Processamento concluído!")
        
        messagebox.showinfo(
            "Sucesso!", 
            f"PDF processado com sucesso!\n\nArquivo salvo em:\n{output_path}"
        )
        
        if messagebox.askyesno("Abrir pasta", "Deseja abrir a pasta onde o arquivo foi salvo?"):
            try:
                os.startfile(Path(output_path).parent)
            except:
                # Para sistemas não-Windows
                import subprocess
                subprocess.run(['xdg-open', str(Path(output_path).parent)])
    
    def processing_error(self, error_msg):
        """Chamado quando ocorre um erro"""
        self.progress.stop()
        self.process_button.config(state='normal')
        self.status_var.set("Erro no processamento")
        messagebox.showerror("Erro", f"Erro ao processar o PDF:\n\n{error_msg}")
    
    def run(self):
        """Executa a aplicação"""
        self.root.mainloop()

def main():
    """Função principal"""
    try:
        app = ShopeeEtiquetaProcessor()
        app.run()
    except ImportError:
        messagebox.showerror("Erro", 
                           "PyMuPDF não está instalado.\n\n" +
                           "Por favor, execute:\npip install PyMuPDF")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao iniciar o programa:\n\n{str(e)}")

if __name__ == "__main__":
    main()