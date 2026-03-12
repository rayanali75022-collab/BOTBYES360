"""
Interface de contrôle du Bot BYES 360
Dashboard avec logs en temps réel, bouton stop, alertes
"""

import queue
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import sys
import os

# Ajouter le répertoire courant au path pour importer bot_byes360
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from bot_byes360 import run_bot


# ─── COULEURS & STYLE ─────────────────────────────────────────────────────────
BG_DARK = "#0D1117"
BG_CARD = "#161B22"
BG_INPUT = "#21262D"
ACCENT_BLUE = "#58A6FF"
ACCENT_GREEN = "#3FB950"
ACCENT_RED = "#F85149"
ACCENT_ORANGE = "#D29922"
ACCENT_PURPLE = "#BC8CFF"
TEXT_PRIMARY = "#E6EDF3"
TEXT_MUTED = "#8B949E"
BORDER_COLOR = "#30363D"

LOG_COLORS = {
    "INFO": TEXT_PRIMARY,
    "SUCCESS": ACCENT_GREEN,
    "ERROR": ACCENT_RED,
    "WARNING": ACCENT_ORANGE,
    "DEBUG": TEXT_MUTED,
    "DONE": ACCENT_PURPLE,
}
# ──────────────────────────────────────────────────────────────────────────────


class BotInterface:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("🤖 Bot BYES 360 — Prise de Commande")
        self.root.configure(bg=BG_DARK)
        self.root.geometry("900x680")
        self.root.resizable(True, True)
        self.root.minsize(700, 500)

        self.log_queue = queue.Queue()
        self.stop_event = threading.Event()
        self.bot_thread = None
        self.is_running = False

        self._build_ui()
        self._poll_logs()

    def _build_ui(self):
        # ── HEADER ──
        header = tk.Frame(self.root, bg=BG_DARK, pady=16)
        header.pack(fill="x", padx=24)

        title_frame = tk.Frame(header, bg=BG_DARK)
        title_frame.pack(fill="x")

        tk.Label(
            title_frame, text="⚡ Bot BYES 360",
            font=("Consolas", 22, "bold"),
            fg=ACCENT_BLUE, bg=BG_DARK
        ).pack(side="left")

        tk.Label(
            title_frame, text="  Automatisation Prise de Commande",
            font=("Consolas", 11),
            fg=TEXT_MUTED, bg=BG_DARK
        ).pack(side="left", pady=6)

        # Séparateur
        sep = tk.Frame(self.root, bg=BORDER_COLOR, height=1)
        sep.pack(fill="x", padx=24)

        # ── ZONE URL ──
        url_card = tk.Frame(self.root, bg=BG_CARD, padx=20, pady=16)
        url_card.pack(fill="x", padx=24, pady=(16, 8))

        tk.Label(
            url_card, text="URL du Devis Salesforce",
            font=("Consolas", 10, "bold"),
            fg=TEXT_MUTED, bg=BG_CARD
        ).pack(anchor="w")

        url_input_frame = tk.Frame(url_card, bg=BG_CARD)
        url_input_frame.pack(fill="x", pady=(6, 0))

        self.url_var = tk.StringVar()
        self.url_entry = tk.Entry(
            url_input_frame,
            textvariable=self.url_var,
            font=("Consolas", 10),
            bg=BG_INPUT, fg=TEXT_PRIMARY,
            insertbackground=ACCENT_BLUE,
            relief="flat",
            bd=0,
        )
        self.url_entry.pack(fill="x", ipady=8, padx=(0, 0))
        self.url_entry.config(highlightthickness=1, highlightbackground=BORDER_COLOR,
                              highlightcolor=ACCENT_BLUE)

        # Placeholder
        self.url_entry.insert(0, "https://equans.lightning.force.com/lightning/r/SBQQ__Quote_c/...")
        self.url_entry.config(fg=TEXT_MUTED)
        self.url_entry.bind("<FocusIn>", self._clear_placeholder)
        self.url_entry.bind("<FocusOut>", self._restore_placeholder)

        # ── BOUTONS ──
        btn_frame = tk.Frame(self.root, bg=BG_DARK, pady=8)
        btn_frame.pack(fill="x", padx=24)

        self.start_btn = tk.Button(
            btn_frame,
            text="▶  LANCER LE BOT",
            font=("Consolas", 11, "bold"),
            bg=ACCENT_GREEN, fg=BG_DARK,
            activebackground="#2EA043", activeforeground=BG_DARK,
            relief="flat", bd=0, padx=20, pady=10,
            cursor="hand2",
            command=self._start_bot
        )
        self.start_btn.pack(side="left", padx=(0, 10))

        self.stop_btn = tk.Button(
            btn_frame,
            text="⛔  STOP",
            font=("Consolas", 11, "bold"),
            bg=ACCENT_RED, fg=TEXT_PRIMARY,
            activebackground="#B91C1C", activeforeground=TEXT_PRIMARY,
            relief="flat", bd=0, padx=20, pady=10,
            cursor="hand2",
            state="disabled",
            command=self._stop_bot
        )
        self.stop_btn.pack(side="left", padx=(0, 10))

        self.clear_btn = tk.Button(
            btn_frame,
            text="🗑  Vider les logs",
            font=("Consolas", 10),
            bg=BG_CARD, fg=TEXT_MUTED,
            activebackground=BG_INPUT, activeforeground=TEXT_PRIMARY,
            relief="flat", bd=0, padx=14, pady=10,
            cursor="hand2",
            command=self._clear_logs
        )
        self.clear_btn.pack(side="left")

        # Statut indicator
        self.status_var = tk.StringVar(value="⬤  En attente")
        self.status_label = tk.Label(
            btn_frame,
            textvariable=self.status_var,
            font=("Consolas", 10),
            fg=TEXT_MUTED, bg=BG_DARK
        )
        self.status_label.pack(side="right", padx=10)

        # ── ZONE LOGS ──
        logs_header = tk.Frame(self.root, bg=BG_DARK)
        logs_header.pack(fill="x", padx=24, pady=(8, 4))

        tk.Label(
            logs_header, text="📋 LOGS EN TEMPS RÉEL",
            font=("Consolas", 9, "bold"),
            fg=TEXT_MUTED, bg=BG_DARK
        ).pack(side="left")

        self.log_count_var = tk.StringVar(value="0 entrées")
        tk.Label(
            logs_header,
            textvariable=self.log_count_var,
            font=("Consolas", 9),
            fg=TEXT_MUTED, bg=BG_DARK
        ).pack(side="right")

        # Zone de logs scrollable
        logs_container = tk.Frame(self.root, bg=BG_CARD, padx=1, pady=1)
        logs_container.pack(fill="both", expand=True, padx=24, pady=(0, 16))

        self.log_text = scrolledtext.ScrolledText(
            logs_container,
            font=("Consolas", 10),
            bg=BG_CARD, fg=TEXT_PRIMARY,
            insertbackground=ACCENT_BLUE,
            relief="flat", bd=0,
            wrap="word",
            state="disabled",
        )
        self.log_text.pack(fill="both", expand=True, padx=8, pady=8)

        # Tags de couleur pour les logs
        for level, color in LOG_COLORS.items():
            self.log_text.tag_configure(level, foreground=color)
        self.log_text.tag_configure("TIME", foreground=TEXT_MUTED)
        self.log_text.tag_configure("SEPARATOR", foreground=BORDER_COLOR)

        # ── BARRE DE STATUT ──
        status_bar = tk.Frame(self.root, bg=BG_CARD, height=28)
        status_bar.pack(fill="x", side="bottom")

        self.bottom_status = tk.Label(
            status_bar,
            text="Prêt — Collez l'URL du devis et lancez le bot",
            font=("Consolas", 9),
            fg=TEXT_MUTED, bg=BG_CARD,
            anchor="w"
        )
        self.bottom_status.pack(side="left", padx=12, pady=4)

        self.log_count = 0

    def _clear_placeholder(self, event):
        if self.url_entry.cget("fg") == TEXT_MUTED:
            self.url_entry.delete(0, "end")
            self.url_entry.config(fg=TEXT_PRIMARY)

    def _restore_placeholder(self, event):
        if not self.url_var.get():
            self.url_entry.insert(0, "https://equans.lightning.force.com/lightning/r/SBQQ__Quote_c/...")
            self.url_entry.config(fg=TEXT_MUTED)

    def _start_bot(self):
        url = self.url_var.get().strip()
        if not url or url.startswith("https://equans.lightning.force.com/lightning/r/SBQQ__Quote_c/..."):
            messagebox.showerror(
                "URL manquante",
                "Collez l'URL du devis Salesforce avant de lancer le bot."
            )
            return

        if "equans.lightning.force.com" not in url:
            if not messagebox.askyesno(
                "URL inhabituelle",
                f"L'URL ne ressemble pas à une URL BYES 360:\n{url}\n\nContinuer quand même?"
            ):
                return

        # Reset
        self.stop_event.clear()
        self.is_running = True
        self._set_running_state(True)
        self._add_log("INFO", "Bot démarré")
        self._add_log("INFO", f"URL cible: {url}")

        # Lancer le bot dans un thread
        self.bot_thread = threading.Thread(
            target=run_bot,
            args=(url, self.log_queue, self.stop_event),
            daemon=True
        )
        self.bot_thread.start()

    def _stop_bot(self):
        if self.is_running:
            if messagebox.askyesno("⛔ Arrêter le bot", "Voulez-vous arrêter le bot en cours d'exécution?"):
                self.stop_event.set()
                self._add_log("WARNING", "Arrêt demandé — en attente de la fin de l'étape en cours...")
                self._set_running_state(False)

    def _set_running_state(self, running: bool):
        self.is_running = running
        if running:
            self.start_btn.config(state="disabled", bg="#2D3748")
            self.stop_btn.config(state="normal")
            self.status_var.set("⬤  En cours...")
            self.status_label.config(fg=ACCENT_ORANGE)
            self.bottom_status.config(text="Bot en cours d'exécution...")
        else:
            self.start_btn.config(state="normal", bg=ACCENT_GREEN)
            self.stop_btn.config(state="disabled")
            self.status_var.set("⬤  En attente")
            self.status_label.config(fg=TEXT_MUTED)
            self.bottom_status.config(text="Prêt — Collez l'URL du devis et lancez le bot")

    def _clear_logs(self):
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, "end")
        self.log_text.config(state="disabled")
        self.log_count = 0
        self.log_count_var.set("0 entrées")

    def _add_log(self, level: str, message: str, timestamp: str = None):
        if timestamp is None:
            from datetime import datetime
            timestamp = datetime.now().strftime("%H:%M:%S")

        self.log_text.config(state="normal")
        self.log_text.insert("end", f"[{timestamp}] ", "TIME")
        self.log_text.insert("end", f"{message}\n", level)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

        self.log_count += 1
        self.log_count_var.set(f"{self.log_count} entrées")

        # Alerte visuelle si erreur
        if level == "ERROR":
            self.root.bell()
            self.bottom_status.config(text=f"⚠️  {message}", fg=ACCENT_RED)
        elif level == "SUCCESS":
            self.bottom_status.config(text=f"✅  {message}", fg=ACCENT_GREEN)

    def _poll_logs(self):
        """Vérifie la queue de logs toutes les 100ms."""
        try:
            while True:
                entry = self.log_queue.get_nowait()
                level = entry.get("level", "INFO")
                message = entry.get("message", "")
                timestamp = entry.get("time", "")

                if level == "DONE":
                    self._set_running_state(False)
                    if not self.stop_event.is_set():
                        self._add_log("INFO", "Bot terminé.", timestamp)
                else:
                    self._add_log(level, message, timestamp)

                    # Alerte popup pour erreurs critiques
                    if level == "ERROR" and any(k in message for k in ["introuvable", "échoué", "critique"]):
                        self.root.after(100, lambda m=message: messagebox.showerror(
                            "⚠️ Erreur Bot", m
                        ))

        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._poll_logs)

    def on_close(self):
        if self.is_running:
            if messagebox.askyesno("Quitter", "Le bot est en cours. Forcer la fermeture?"):
                self.stop_event.set()
                self.root.destroy()
        else:
            self.root.destroy()


def main():
    root = tk.Tk()
    app = BotInterface(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)

    # Icône (ignorer si pas dispo)
    try:
        root.iconbitmap("robot.ico")
    except Exception:
        pass

    root.mainloop()


if __name__ == "__main__":
    main()
