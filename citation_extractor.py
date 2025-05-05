import re
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Label
from docx import Document

def extract_apa_citations(docx_path):
    doc = Document(docx_path)
    text = "\n".join([para.text for para in doc.paragraphs])

    parenthetical = r'\(([^()]*?\b(?:[\w\-\u00C0-\u024F\u0370-\u03FF]+(?: et al\.)?(?: & [\w\-\u00C0-\u024F\u0370-\u03FF]+)?), \d{4}(?:[a-z])?(?:; [^()]*?, \d{4}[a-z]?)*?)\)'
    # narrative = r'\b([\w\-\u00C0-\u024F\u0370-\u03FF]+(?: et al\.)?(?: and [\w\-\u00C0-\u024F\u0370-\u03FF]+)?) \(\d{4}[a-z]?\)'

    parenthetical_matches = re.findall(parenthetical, text, re.UNICODE)
    narrative_full = re.findall(r'\b[\w\-\u00C0-\u024F\u0370-\u03FF]+(?: et al\.)?(?: and [\w\-\u00C0-\u024F\u0370-\u03FF]+)? \(\d{4}[a-z]?\)', text, re.UNICODE)

    citations = set(f"{match}" for match in parenthetical_matches)
    citations.update(narrative_full)

    return sorted(citations)

def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        try:
            citations = extract_apa_citations(file_path)
            text_box.delete('1.0', tk.END)
            if citations:
                for cite in citations:
                    text_box.insert(tk.END, cite + "\n")
            else:
                text_box.insert(tk.END, "No APA-style citations found.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not extract citations: {e}")

def save_to_txt():
    save_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
    if save_path:
        content = text_box.get("1.0", tk.END).strip()
        with open(save_path, "w", encoding="utf-8") as f:
            f.write(content)
        messagebox.showinfo("Saved", f"Citations saved to {save_path}")

# GUI Setup
root = tk.Tk()
root.title("APA Citation Extractor © 2025 Augustine C. Emmanuel")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(fill="both", expand=True)

open_button = tk.Button(frame, text="Open DOCX File", command=open_file)
open_button.pack(pady=(0, 5))

text_box = scrolledtext.ScrolledText(frame, width=80, height=20, wrap=tk.WORD)
text_box.pack(padx=5, pady=5)

save_button = tk.Button(frame, text="Save Citations as TXT", command=save_to_txt)
save_button.pack(pady=(5, 0))

Label(root, text="© 2025 Augustine C. Emmanuel (ACE)", font=("Arial", 10), fg="gray").pack(side="bottom", pady=5)


root.mainloop()
