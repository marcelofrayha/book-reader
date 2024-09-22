import pyttsx3
from PyPDF2 import PdfReader
from ebooklib import epub
import html2text
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import logging
import tempfile
import io
import re

def clean_text(text):
    # Remove referências a notas de rodapé (números sobrescritos)
    text = re.sub(r'[¹²³⁴⁵⁶⁷⁸⁹⁰]+', '', text)
    
    return text.strip()

def extract_text_from_pdf(pdf_path, start_page, progress_callback):
    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)
    text = ""
    for i in range(start_page - 1, total_pages):
        page_text = reader.pages[i].extract_text()
        text += clean_text(page_text) + "\n\n"  # Adiciona quebras de parágrafo entre páginas
        progress_callback(i + 1, total_pages)
    return text, total_pages

def extract_text_from_epub(epub_path, start_page, progress_callback):
    book = epub.read_epub(epub_path)
    h = html2text.HTML2Text()
    h.ignore_links = True
    text = ""
    items = list(book.get_items())
    total_pages = len([item for item in items if item.get_type() == epub.ITEM_DOCUMENT])
    current_page = 0
    for item in items:
        if item.get_type() == epub.ITEM_DOCUMENT:
            current_page += 1
            if current_page >= start_page:
                content = item.get_content()
                if isinstance(content, bytes):
                    content = content.decode('utf-8', errors='ignore')
                page_text = h.handle(content)
                text += clean_text(page_text) + "\n\n"  # Adiciona quebras de parágrafo entre seções
                progress_callback(current_page, total_pages)
    return text, total_pages

def text_to_speech(text, output_file, voice_id, speed, progress_callback):
    engine = None
    temp_file = tempfile.NamedTemporaryFile(suffix=".mp3", delete=False).name
    try:
        logging.info(f"Initializing pyttsx3 engine")
        engine = pyttsx3.init()
        
        # Set properties
        engine.setProperty('rate', speed)  # Speed of speech
        engine.setProperty('volume', 0.9)  # Volume (0.0 to 1.0)
        
        # Set the selected voice if provided
        if voice_id:
            engine.setProperty('voice', voice_id)
        
        def onWord(name, location, length):
            progress = location / len(text)
            progress_callback(progress)
        
        engine.connect('started-word', onWord)
        
        logging.info(f"Saving audio to temporary file: {temp_file}")
        engine.save_to_file(text, temp_file)
        
        logging.info("Running pyttsx3 engine")
        engine.runAndWait()
        
        if os.path.exists(temp_file) and os.path.getsize(temp_file) > 0:
            logging.info(f"Temporary audio file created successfully: {temp_file}")
            os.rename(temp_file, output_file)
            logging.info(f"Audio file moved to final location: {output_file}")
        else:
            raise Exception("Temporary audio file is empty or was not created")
        
    except Exception as e:
        logging.error(f"Error in text_to_speech: {str(e)}", exc_info=True)
        raise
    finally:
        if engine:
            logging.info("Stopping pyttsx3 engine")
            engine.stop()
        if os.path.exists(temp_file):
            os.remove(temp_file)

def generate_output_filename(input_file):
    # Get the directory and filename without extension
    directory, filename = os.path.split(input_file)
    name_without_ext = os.path.splitext(filename)[0]
    
    # Create the new filename with .mp3 extension
    output_filename = f"{name_without_ext}.mp3"
    
    # Join with the original directory
    return os.path.join(directory, output_filename)

class BookToAudioGUI:
    def __init__(self, master):
        self.master = master
        master.title("Book to Audio Converter")
        master.geometry("550x450")

        self.input_label = tk.Label(master, text="Input File:")
        self.input_label.pack()

        self.input_entry = tk.Entry(master, width=50)
        self.input_entry.pack()
        self.input_entry.bind("<KeyRelease>", self.update_output_entry)

        self.input_button = tk.Button(master, text="Browse", command=self.browse_input)
        self.input_button.pack()

        self.output_label = tk.Label(master, text="Output File:")
        self.output_label.pack()

        self.output_entry = tk.Entry(master, width=50, state="readonly")
        self.output_entry.pack()

        self.start_page_label = tk.Label(master, text="Start Page:")
        self.start_page_label.pack()

        self.start_page_entry = tk.Entry(master, width=10)
        self.start_page_entry.insert(0, "1")  # Default to start from page 1
        self.start_page_entry.pack()

        self.voice_label = tk.Label(master, text="Select Voice:")
        self.voice_label.pack()

        self.voice_var = tk.StringVar(master)
        self.voice_dropdown = ttk.Combobox(master, textvariable=self.voice_var, state="readonly", width=50)
        self.voice_dropdown.pack()

        self.voice_info_label = tk.Label(master, text="")
        self.voice_info_label.pack()
        self.speed_label = tk.Label(master, text="Speech Speed:")
        self.speed_label.pack()

        self.speed_scale = tk.Scale(master, from_=50, to=300, orient=tk.HORIZONTAL, length=200)
        self.speed_scale.set(150)  # Valor padrão
        self.speed_scale.pack()

        self.populate_voices()

        self.progress_label = tk.Label(master, text="Progress:")
        self.progress_label.pack()

        self.progress_bar = ttk.Progressbar(master, length=300, mode='determinate')
        self.progress_bar.pack()

        self.page_progress_label = tk.Label(master, text="Pages: 0 / 0")
        self.page_progress_label.pack()


        self.convert_button = tk.Button(master, text="Convert", command=self.convert)
        self.convert_button.pack()

        self.voice_dropdown.bind("<<ComboboxSelected>>", self.update_voice_info)

    def populate_voices(self):
        try:
            engine = pyttsx3.init()
            voices = engine.getProperty('voices')
            self.voices = voices  # Store voices for later use
            voice_options = [f"{voice.name} ({voice.languages[0] if voice.languages else 'Unknown'})" for voice in voices]
            self.voice_dropdown['values'] = voice_options
            if voice_options:
                self.voice_dropdown.set(voice_options[0])
                self.update_voice_info()
            else:
                logging.warning("No voices found on the system.")
                messagebox.showwarning("Warning", "No voices found on your system. The default system voice will be used.")
        except Exception as e:
            logging.error(f"Error populating voices: {str(e)}")
            messagebox.showerror("Error", f"Unable to retrieve system voices: {str(e)}")

    def browse_input(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf"), ("EPUB Files", "*.epub")])
        self.input_entry.delete(0, tk.END)
        self.input_entry.insert(0, filename)
        self.update_output_entry()

    def update_output_entry(self, event=None):
        input_file = self.input_entry.get()
        if input_file:
            output_file = generate_output_filename(input_file)
            self.output_entry.config(state="normal")
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, output_file)
            self.output_entry.config(state="readonly")

    def update_progress(self, progress):
        self.progress_bar['value'] = progress * 100
        self.master.update_idletasks()

    def update_page_progress(self, current_page, total_pages):
        self.page_progress_label.config(text=f"Pages: {current_page} / {total_pages}")
        self.progress_bar['value'] = (current_page / total_pages) * 100
        self.master.update_idletasks()

    def update_voice_info(self, event=None):
        selected_index = self.voice_dropdown.current()
        if selected_index >= 0:
            voice = self.voices[selected_index]
            info = f"Name: {voice.name}\n"
            info += f"ID: {voice.id}\n"
            info += f"Languages: {', '.join(voice.languages) if voice.languages else 'Unknown'}\n"
            info += f"Gender: {voice.gender if hasattr(voice, 'gender') else 'Unknown'}"
            self.voice_info_label.config(text=info)

    def convert(self):
        input_file = self.input_entry.get()
        output_file = self.output_entry.get()
        start_page = int(self.start_page_entry.get())
        selected_voice = self.voice_var.get().split(" (")[0]  # Get only the voice name
        speed = self.speed_scale.get()  # Get the selected speed

        if not input_file or not output_file:
            messagebox.showerror("Error", "Please select both input and output files.")
            return

        def conversion_thread():
            try:
                self.progress_bar['value'] = 0
                self.page_progress_label.config(text="Pages: 0 / 0")
                self.convert_button['state'] = 'disabled'

                if input_file.lower().endswith('.pdf'):
                    text, total_pages = extract_text_from_pdf(input_file, start_page, self.update_page_progress)
                elif input_file.lower().endswith('.epub'):
                    text, total_pages = extract_text_from_epub(input_file, start_page, self.update_page_progress)
                else:
                    messagebox.showerror("Error", "Unsupported file format. Please use PDF or EPUB.")
                    return

                self.progress_bar['value'] = 50
                self.master.update_idletasks()

                # Get the voice ID for the selected voice name
                voice_id = next((voice.id for voice in self.voices if voice.name == selected_voice), None)

                if voice_id is None:
                    logging.warning("Selected voice not found. Using system default.")

                text_to_speech(text, output_file, voice_id, speed, self.update_progress)
                
                self.progress_bar['value'] = 100
                self.master.update_idletasks()
                
                messagebox.showinfo("Success", f"Audio file has been created: {output_file}")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {str(e)}")
            finally:
                self.convert_button['state'] = 'normal'

        thread = threading.Thread(target=conversion_thread)
        thread.start()

def main():
    logging.basicConfig(level=logging.DEBUG)
    root = tk.Tk()
    app = BookToAudioGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
