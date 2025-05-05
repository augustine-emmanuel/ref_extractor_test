import re
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import win32com.client


def extract_citations_with_pages(docx_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(docx_path)

    results = []
    seen = set()

    # Read full content
    full_text = doc.Content.Text

    # APA-style citation patterns
    patterns = [
        r"\(([^()]*?\b(?:[\w\-\u00C0-\u024F\u0370-\u03FF]+(?: et al\.)?(?: & [\w\-\u00C0-\u024F\u0370-\u03FF]+)?), \d{4}[a-z]?(?:; [^()]*?, \d{4}[a-z]?)*?)\)",
        r"\b([\w\-\u00C0-\u024F\u0370-\u03FF]+(?: et al\.)?(?: and [\w\-\u00C0-\u024F\u0370-\u03FF]+)?) \(\d{4}[a-z]?\)",
    ]

    find = word.Selection.Find
    find.ClearFormatting()

    for pattern in patterns:
        for match in re.finditer(pattern, full_text):
            citation = match.group().strip()
            if citation in seen:
                continue

            find.Text = citation
            if find.Execute():
                page = word.Selection.Information(3)  # 3 = wdActiveEndPageNumber
                results.append(f"{citation} — Page {page}")
                seen.add(citation)

    doc.Close(False)
    word.Quit()
    return sorted(results)


def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        try:
            text_box.delete('1.0', tk.END)
            citations = extract_citations_with_pages(file_path)
            if citations:
                for c in citations:
                    text_box.insert(tk.END, c + "\n")
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
root.title("APA Citation Extractor with Page Numbers")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(fill="both", expand=True)

open_button = tk.Button(frame, text="Open DOCX File", command=open_file)
open_button.pack(pady=(0, 5))

text_box = scrolledtext.ScrolledText(frame, width=80, height=20, wrap=tk.WORD)
text_box.pack(padx=5, pady=5)

save_button = tk.Button(frame, text="Save Citations as TXT", command=save_to_txt)
save_button.pack(pady=(5, 0))

root.mainloop()
