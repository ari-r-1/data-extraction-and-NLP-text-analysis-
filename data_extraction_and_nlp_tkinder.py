import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import threading, os, re, time, shutil, nltk, requests
import pandas as pd
from bs4 import BeautifulSoup
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize

# ───────────────────────────── NLTK one‑time setup ───────────────────────────── #
for pkg in ["punkt", "stopwords", "wordnet", "averaged_perceptron_tagger"]:
    try:
        nltk.data.find(pkg if pkg == "stopwords" else f"tokenizers/{pkg}")
    except LookupError:
        nltk.download(pkg)

# ───────────────────────────── Helper Functions ─────────────────────────────── #

def log(msg):
    """Thread‑safe append to the GUI log window."""
    log_text.configure(state="normal")
    log_text.insert(tk.END, msg + "\n")
    log_text.see(tk.END)
    log_text.configure(state="disabled")

def load_wordlist(path):
    with open(path, encoding="latin-1") as f:
        return {line.strip() for line in f if line.strip() and not line.startswith(";")}

def count_syllables(word):
    word = word.lower().strip(".:;?!")
    if not word: return 0
    if word.endswith(("es", "ed")): word = word[:-2]
    if word.endswith("e") and not word.endswith(("le", "ue")): word = word[:-1]

    vowels, count, prev = "aeiouy", 0, False
    for ch in word:
        is_vowel = ch in vowels
        if is_vowel and not prev: count += 1
        prev = is_vowel
    return count or 1

def count_pronouns(text):
    pronouns = re.findall(r"\b(I|we|my|ours|us)\b", text, flags=re.I)
    return sum(1 for p in pronouns if p.lower() != "us" or
               not re.search(r"\b(united|states|u\.s\.)\b", text, flags=re.I))

# ───────────────────────────── Core Analysis ────────────────────────────────── #

def run_analysis():
    try:
        btn_run["state"] = tk.DISABLED
        progress["value"] = 0
        results = []

        df = pd.read_excel(entry_input.get())
        if not {"URL_ID", "URL"} <= set(df.columns):
            raise ValueError("Input.xlsx must contain URL_ID and URL columns.")

        pos_words = load_wordlist(entry_pos.get())
        neg_words = load_wordlist(entry_neg.get())
        stop_words = set(stopwords.words("english"))

        headers = {"User-Agent":
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/115 Safari/537.36"}

        total = len(df)
        for i, row in df.iterrows():
            url_id, url = row["URL_ID"], row["URL"]
            log(f"Fetching {url_id}: {url}")
            try:
                r = requests.get(url, headers=headers, timeout=15)
                r.raise_for_status()
                if "text/html" not in r.headers.get("Content-Type", ""):
                    raise ValueError("Non‑HTML content")

                soup = BeautifulSoup(r.content, "html.parser")
                text = " ".join(p.get_text(" ", strip=True) for p in soup.find_all("p"))
                text = re.sub(r"\s+", " ", text).strip()
                if not text:
                    raise ValueError("No textual content")

                sents = sent_tokenize(text)
                words = [w for w in word_tokenize(text)
                         if w.isalpha() and w.lower() not in stop_words]

                pos = sum(w.lower() in pos_words for w in words)
                neg = sum(w.lower() in neg_words for w in words)
                tot_words = len(words) or 1
                polarity = (pos - neg) / (pos + neg + 1e-9)
                subj = (pos + neg) / tot_words

                avg_sent = tot_words / (len(sents) or 1)
                complex_w = [w for w in words if count_syllables(w) > 2]
                pct_complex = len(complex_w) / tot_words
                fog = 0.4 * (avg_sent + pct_complex)
                syll_per_word = sum(count_syllables(w) for w in words) / tot_words
                avg_word_len = sum(len(w) for w in words) / tot_words
                pronouns = count_pronouns(text)

                results.append([url_id, url, pos, neg, polarity, subj,
                                avg_sent, pct_complex*100, fog, avg_sent,
                                len(complex_w), tot_words, syll_per_word,
                                pronouns, avg_word_len])

            except Exception as e:
                log(f"  ✖ {url_id} failed: {e}")
                results.append([url_id, url] + [None]*13)

            progress["value"] = ((i+1)/total)*100
            time.sleep(0.2)  # gentle delay

        cols = ["URL_ID","URL","POSITIVE SCORE","NEGATIVE SCORE","POLARITY SCORE",
                "SUBJECTIVITY SCORE","AVG SENTENCE LENGTH","PERCENTAGE OF COMPLEX WORDS",
                "FOG INDEX","AVG NUMBER OF WORDS PER SENTENCE","COMPLEX WORD COUNT",
                "WORD COUNT","SYLLABLE PER WORD","PERSONAL PRONOUNS","AVG WORD LENGTH"]
        pd.DataFrame(results, columns=cols).to_excel("Output.xlsx", index=False)
        log("\n✔ Analysis complete. Output.xlsx created.")
        messagebox.showinfo("Done", "Analysis complete!\nSaved as Output.xlsx.")
    except Exception as e:
        messagebox.showerror("Error", str(e))
        log(f"\n✖ Error: {e}")
    finally:
        btn_run["state"] = tk.NORMAL
        progress["value"] = 0

# ───────────────────────────── Tkinter UI ───────────────────────────────────── #

root = tk.Tk()
root.title("Data Extraction And NLP Analyzer")

rowcfg = dict(padx=6, pady=4, sticky="ew")
root.columnconfigure(1, weight=1)

# Input.xlsx
tk.Label(root, text="Input.xlsx:").grid(row=0, column=0, **rowcfg)
entry_input = tk.Entry(root)
entry_input.grid(row=0, column=1, **rowcfg)
tk.Button(root, text="Browse",
          command=lambda: entry_input.insert(0, filedialog.askopenfilename(
              filetypes=[("Excel files","*.xlsx")]))).grid(row=0, column=2, **rowcfg)

# positive‑words.txt
tk.Label(root, text="positive‑words.txt:").grid(row=1, column=0, **rowcfg)
entry_pos = tk.Entry(root)
entry_pos.grid(row=1, column=1, **rowcfg)
tk.Button(root, text="Browse",
          command=lambda: entry_pos.insert(0, filedialog.askopenfilename(
              filetypes=[("Text files","*.txt")]))).grid(row=1, column=2, **rowcfg)

# negative‑words.txt
tk.Label(root, text="negative‑words.txt:").grid(row=2, column=0, **rowcfg)
entry_neg = tk.Entry(root)
entry_neg.grid(row=2, column=1, **rowcfg)
tk.Button(root, text="Browse",
          command=lambda: entry_neg.insert(0, filedialog.askopenfilename(
              filetypes=[("Text files","*.txt")]))).grid(row=2, column=2, **rowcfg)

# Run button
btn_run = tk.Button(root, text="Run Analysis",
                    command=lambda: threading.Thread(target=run_analysis, daemon=True).start())
btn_run.grid(row=3, column=0, columnspan=3, pady=(10,2))

# Progress bar
progress = ttk.Progressbar(root, length=400, mode="determinate")
progress.grid(row=4, column=0, columnspan=3, pady=(0,8))

# Log window
log_text = scrolledtext.ScrolledText(root, height=15, state="disabled")
log_text.grid(row=5, column=0, columnspan=3, padx=6, pady=6, sticky="nsew")
root.rowconfigure(5, weight=1)

root.mainloop()