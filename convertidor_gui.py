"""
Interfaz gráfica para el Convertidor de APU (PDF a Excel)
Permite seleccionar archivos PDF y convertirlos al formato Excel estandarizado.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import threading

# Importar el convertidor
from pdf_to_excel_apu import convert_pdf_to_excel, APUConverter


class APUConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertidor APU - PDF a Excel")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        # Configurar estilo
        style = ttk.Style()
        style.configure('TButton', padding=5)
        style.configure('TLabel', padding=2)
        
        # Frame principal
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title_label = ttk.Label(main_frame, text="Convertidor de Análisis de Precios Unitarios", 
                                font=('Arial', 14, 'bold'))
        title_label.pack(pady=10)
        
        subtitle_label = ttk.Label(main_frame, text="PDF con VAE → Excel Estandarizado")
        subtitle_label.pack()
        
        # Frame de selección de archivo
        file_frame = ttk.LabelFrame(main_frame, text="Archivo de entrada", padding="10")
        file_frame.pack(fill=tk.X, pady=10)
        
        self.file_path = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path, width=50)
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        browse_btn = ttk.Button(file_frame, text="Examinar...", command=self.browse_file)
        browse_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        # Frame de salida
        output_frame = ttk.LabelFrame(main_frame, text="Archivo de salida (opcional)", padding="10")
        output_frame.pack(fill=tk.X, pady=10)
        
        self.output_path = tk.StringVar()
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=50)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        output_btn = ttk.Button(output_frame, text="Examinar...", command=self.browse_output)
        output_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=10)
        
        # Área de log
        log_frame = ttk.LabelFrame(main_frame, text="Estado", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.log_text = tk.Text(log_frame, height=8, state='disabled', wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Botón de convertir
        convert_btn = ttk.Button(main_frame, text="Convertir PDF a Excel", 
                                 command=self.start_conversion, style='TButton')
        convert_btn.pack(pady=10)
        
        self.log("Listo. Selecciona un archivo PDF para convertir.")
    
    def log(self, message):
        """Agrega un mensaje al área de log."""
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state='disabled')
        self.root.update()
    
    def browse_file(self):
        """Abre diálogo para seleccionar archivo PDF."""
        filepath = filedialog.askopenfilename(
            title="Seleccionar archivo PDF",
            filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")]
        )
        if filepath:
            self.file_path.set(filepath)
            # Auto-generar nombre de salida
            base_name = os.path.splitext(os.path.basename(filepath))[0]
            output_name = f"{base_name}_CONVERTIDO.xlsx"
            output_full = os.path.join(os.path.dirname(filepath), output_name)
            self.output_path.set(output_full)
            self.log(f"Archivo seleccionado: {os.path.basename(filepath)}")
    
    def browse_output(self):
        """Abre diálogo para seleccionar archivo de salida."""
        filepath = filedialog.asksaveasfilename(
            title="Guardar archivo Excel como",
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        )
        if filepath:
            self.output_path.set(filepath)
    
    def start_conversion(self):
        """Inicia la conversión en un hilo separado."""
        pdf_path = self.file_path.get()
        output_path = self.output_path.get() or None
        
        if not pdf_path:
            messagebox.showerror("Error", "Por favor selecciona un archivo PDF")
            return
        
        if not os.path.exists(pdf_path):
            messagebox.showerror("Error", f"El archivo no existe: {pdf_path}")
            return
        
        # Iniciar conversión en hilo separado
        self.progress.start()
        thread = threading.Thread(target=self.convert, args=(pdf_path, output_path))
        thread.start()
    
    def convert(self, pdf_path, output_path):
        """Ejecuta la conversión."""
        try:
            self.log("Iniciando conversión...")
            result = convert_pdf_to_excel(pdf_path, output_path)
            self.progress.stop()
            self.log(f"✓ Conversión completada!")
            self.log(f"  Archivo generado: {result}")
            messagebox.showinfo("Éxito", f"Conversión completada!\n\nArchivo: {os.path.basename(result)}")
        except Exception as e:
            self.progress.stop()
            self.log(f"✗ Error: {str(e)}")
            messagebox.showerror("Error", f"Error durante la conversión:\n{str(e)}")


def main():
    root = tk.Tk()
    app = APUConverterGUI(root)
    
    # Si se pasó un archivo como argumento, cargarlo
    if len(sys.argv) > 1:
        app.file_path.set(sys.argv[1])
        base_name = os.path.splitext(os.path.basename(sys.argv[1]))[0]
        output_name = f"{base_name}_CONVERTIDO.xlsx"
        output_full = os.path.join(os.path.dirname(sys.argv[1]), output_name)
        app.output_path.set(output_full)
    
    root.mainloop()


if __name__ == "__main__":
    main()
