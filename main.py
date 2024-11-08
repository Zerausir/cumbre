import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk
import os


class FrequencyManager:
    def __init__(self, excel_path):
        """Inicializa el administrador de frecuencias."""
        self.excel_path = Path(excel_path)
        try:
            self.df = pd.read_excel(self.excel_path)
        except PermissionError:
            messagebox.showerror("Error", "No se puede acceder al archivo Excel. Por favor, ciérrelo si está abierto.")
            raise

    def check_frequency(self, frequency):
        """Verifica si una frecuencia existe en tx o rx."""
        tx_matches = self.df[self.df['Frecuencia_tx'] == frequency]
        rx_matches = self.df[self.df['Frecuencia_rx'] == frequency]

        all_matches = pd.concat([tx_matches, rx_matches])
        return all_matches if not all_matches.empty else None

    def register_frequency(self, frequency, name, area):
        """Registra una nueva frecuencia, nombre y área de operación."""
        try:
            if self.check_frequency(frequency) is not None:
                return False, "La frecuencia ya existe en el registro"

            new_row = pd.Series({
                'Nombres': name,
                'Frecuencia_tx': frequency,
                'Areas_Operacion': area
            })

            self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)

            # Intentar guardar con manejo de errores
            try:
                self.df.to_excel(self.excel_path, index=False)
                return True, "Frecuencia registrada exitosamente"
            except PermissionError:
                return False, "No se puede guardar porque el archivo Excel está abierto. Por favor, ciérrelo e intente nuevamente."
            except Exception as e:
                return False, f"Error al guardar: {str(e)}"

        except Exception as e:
            return False, f"Error al registrar: {str(e)}"


class FrequencyManagerGUI:
    def __init__(self, excel_path):
        try:
            self.manager = FrequencyManager(excel_path)
        except:
            return

        self.root = tk.Tk()
        self.root.title("Administrador de Frecuencias")
        self.root.geometry("800x500")

        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Sección de búsqueda
        search_frame = ttk.LabelFrame(main_frame, text="Búsqueda de Frecuencia", padding="5")
        search_frame.grid(row=0, column=0, columnspan=3, pady=5, sticky=(tk.W, tk.E))

        ttk.Label(search_frame, text="Frecuencia:").grid(row=0, column=0, sticky=tk.W)
        self.search_freq = ttk.Entry(search_frame)
        self.search_freq.grid(row=0, column=1, padx=5)
        ttk.Button(search_frame, text="Buscar", command=self.search_frequency).grid(row=0, column=2)
        ttk.Button(search_frame, text="Usar para Registro", command=self.copy_to_register).grid(row=0, column=3, padx=5)

        # Área de resultados
        self.result_text = tk.Text(main_frame, height=12, width=80)
        self.result_text.grid(row=1, column=0, columnspan=3, pady=10)

        # Sección de registro
        register_frame = ttk.LabelFrame(main_frame, text="Registrar Nueva Frecuencia", padding="5")
        register_frame.grid(row=2, column=0, columnspan=3, pady=10, sticky=(tk.W, tk.E))

        ttk.Label(register_frame, text="Frecuencia:").grid(row=0, column=0, sticky=tk.W)
        self.new_freq = ttk.Entry(register_frame)
        self.new_freq.grid(row=0, column=1, padx=5)

        ttk.Label(register_frame, text="Nombre:").grid(row=1, column=0, sticky=tk.W)
        self.new_name = ttk.Entry(register_frame)
        self.new_name.grid(row=1, column=1, padx=5)

        ttk.Label(register_frame, text="Área de Operación:").grid(row=2, column=0, sticky=tk.W)
        self.new_area = ttk.Entry(register_frame)
        self.new_area.grid(row=2, column=1, padx=5)

        ttk.Button(register_frame, text="Registrar", command=self.register_frequency).grid(row=3, column=0,
                                                                                           columnspan=2, pady=10)

    def copy_to_register(self):
        """Copia la frecuencia buscada al formulario de registro."""
        freq = self.search_freq.get()
        self.new_freq.delete(0, tk.END)
        self.new_freq.insert(0, freq)

    def search_frequency(self):
        """Busca una frecuencia y muestra los resultados."""
        try:
            freq = float(self.search_freq.get())
            matches = self.manager.check_frequency(freq)

            self.result_text.delete(1.0, tk.END)
            if matches is not None:
                self.result_text.insert(tk.END, "Frecuencia encontrada en las siguientes entradas:\n\n")
                for _, row in matches.iterrows():
                    self.result_text.insert(tk.END, f"Nombre: {row['Nombres']}\n")
                    if not pd.isna(row['Frecuencia_tx']):
                        self.result_text.insert(tk.END, f"Frecuencia TX: {row['Frecuencia_tx']}\n")
                    if not pd.isna(row['Frecuencia_rx']):
                        self.result_text.insert(tk.END, f"Frecuencia RX: {row['Frecuencia_rx']}\n")
                    self.result_text.insert(tk.END, f"Área de Operación: {row['Areas_Operacion']}\n")
                    self.result_text.insert(tk.END, "-" * 50 + "\n")
            else:
                self.result_text.insert(tk.END,
                                        "Frecuencia no encontrada. Puede registrarla usando el formulario inferior.\n")
                self.result_text.insert(tk.END,
                                        "Use el botón 'Usar para Registro' para copiar la frecuencia al formulario.")
        except ValueError:
            messagebox.showerror("Error", "Por favor ingrese un valor numérico válido para la frecuencia")

    def register_frequency(self):
        """Registra una nueva frecuencia."""
        try:
            freq = float(self.new_freq.get())
            name = self.new_name.get().strip()
            area = self.new_area.get().strip()

            if not name:
                messagebox.showerror("Error", "Por favor ingrese un nombre")
                return

            if not area:
                messagebox.showerror("Error", "Por favor ingrese un área de operación")
                return

            success, message = self.manager.register_frequency(freq, name, area)
            if success:
                messagebox.showinfo("Éxito", message)
                self.new_freq.delete(0, tk.END)
                self.new_name.delete(0, tk.END)
                self.new_area.delete(0, tk.END)
            else:
                messagebox.showerror("Error", message)
        except ValueError:
            messagebox.showerror("Error", "Por favor ingrese un valor numérico válido para la frecuencia")

    def run(self):
        """Inicia la aplicación."""
        self.root.mainloop()


# Uso de la aplicación
if __name__ == "__main__":
    app = FrequencyManagerGUI("11. SIGER_05nov2024.xlsx")  # Asegúrate de que este sea el nombre correcto de tu archivo
    app.run()
