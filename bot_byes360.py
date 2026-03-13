"""
Bot BYES 360 - Automatisation Prise de Commande
Extrait le compte depuis le PDF PO et le met à jour dans Salesforce.
"""

import glob
import os
import queue
import re
import sys
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

from selenium import webdriver
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

DOWNLOAD_DIR = str(Path.home() / "Downloads")
WAIT_TIMEOUT = 20
SHORT_WAIT = 5


class BotBYES360:
    def __init__(self, log_queue: queue.Queue, stop_event: threading.Event):
        self.log_queue = log_queue
        self.stop_event = stop_event
        self.driver = None
        self.wait = None

    def log(self, message: str, level: str = "INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put({"time": timestamp, "level": level, "message": message})
        print(f"[{timestamp}] [{level}] {message}")

    def check_stop(self):
        if self.stop_event.is_set():
            raise InterruptedError("⛔ Arrêt demandé par l'utilisateur.")

    def init_driver(self):
        self.log("Initialisation Edge avec ton profil existant...")
        self.log("⚠️ Ferme toutes les fenetres Edge avant de continuer !", "WARNING")
        try:
            opts = EdgeOptions()
            user_data = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Edge", "User Data")
            opts.add_argument(f"--user-data-dir={user_data}")
            opts.add_argument("--profile-directory=Default")
            opts.add_argument("--start-maximized")
            opts.add_argument("--no-first-run")
            opts.add_argument("--no-default-browser-check")
            opts.add_experimental_option("excludeSwitches", ["enable-automation"])
            opts.add_experimental_option("useAutomationExtension", False)
            opts.add_experimental_option(
                "prefs",
                {
                    "download.default_directory": DOWNLOAD_DIR,
                    "download.prompt_for_download": False,
                    "plugins.always_open_pdf_externally": True,
                },
            )
            self.driver = webdriver.Edge(options=opts)
            self.wait = WebDriverWait(self.driver, WAIT_TIMEOUT)
            self.log("Edge ouvert avec ton profil (cookies/session conserves) ✅")
            return True
        except Exception as e:
            self.log(f"❌ Erreur: {e}", "ERROR")
            return False

    def close_driver(self):
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass

    def wait_for_download(self, timeout=45) -> str | None:
        self.log("Attente du téléchargement PDF...")

        def _snapshot_files():
            files = {}
            try:
                for name in os.listdir(DOWNLOAD_DIR):
                    path = os.path.join(DOWNLOAD_DIR, name)
                    if not os.path.isfile(path):
                        continue
                    files[path] = os.path.getmtime(path)
            except Exception:
                pass
            return files

        def _is_completed_pdf(path: str) -> bool:
            base = os.path.basename(path).lower()
            if base.endswith((".crdownload", ".tmp", ".part")):
                return False
            return base.endswith(".pdf")

        before = _snapshot_files()
        start = time.time()

        while time.time() - start < timeout:
            self.check_stop()
            time.sleep(1)
            after = _snapshot_files()

            # 1) nouveaux fichiers PDF
            new_paths = [p for p in after.keys() if p not in before and _is_completed_pdf(p)]
            if new_paths:
                newest = max(new_paths, key=lambda p: after.get(p, 0))
                self.log(f"PDF téléchargé: {os.path.basename(newest)} ✅")
                return newest

            # 2) fallback: fichier PDF modifié après le début (cas renommage/écrasement)
            recent_pdf = [p for p, mtime in after.items() if _is_completed_pdf(p) and mtime >= start]
            if recent_pdf:
                newest = max(recent_pdf, key=lambda p: after.get(p, 0))
                self.log(f"PDF détecté (fallback): {os.path.basename(newest)} ✅")
                return newest

        self.log("Timeout: aucun PDF téléchargé", "WARNING")
        return None

    def extract_account_from_pdf(self, pdf_path: str) -> str | None:
        self.log("Extraction du compte depuis le PDF...")
        if pdfplumber is None:
            self.log("pdfplumber non installé, tentative extraction basique", "WARNING")
            return self._extract_account_basic(pdf_path)

        try:
            with pdfplumber.open(pdf_path) as pdf:
                text = pdf.pages[0].extract_text() if pdf.pages else ""
                if not text:
                    self.log("PDF vide ou non lisible", "ERROR")
                    return None

                lines = text.split("\n")
                found_billing = False
                for line in lines:
                    line_clean = line.strip()
                    if "adresse de facturation" in line_clean.lower():
                        found_billing = True
                        continue
                    if found_billing and line_clean:
                        account_name = self._clean_account_name(line_clean)
                        if account_name:
                            self.log(f"Compte extrait: '{account_name}' ✅")
                            return account_name

                self.log("Compte non trouvé dans le PDF", "WARNING")
                return self._extract_account_fallback(text)
        except Exception as e:
            self.log(f"Erreur extraction PDF: {e}", "ERROR")
            return None

    def _clean_account_name(self, raw: str) -> str | None:
        cleaned = re.sub(r"^\d+\s*[-–]\s*", "", raw).strip()
        cleaned = re.sub(r"\(.*?\)", "", cleaned).strip()
        return cleaned if len(cleaned) >= 3 else None

    def _extract_account_fallback(self, text: str) -> str | None:
        pattern = r"Adresse de facturation\s*\n\s*(\d+\s*[-–]\s*)?([A-Z][A-Z\s&]+(?:SAS|SA|SARL|SNC|SASU|GIE|GROUP|IMMO)?)"
        match = re.search(pattern, text, re.IGNORECASE)
        if match and match.group(2):
            name = match.group(2).strip()
            self.log(f"Compte extrait (fallback): '{name}'")
            return name
        return None

    def _extract_account_basic(self, pdf_path: str) -> str | None:
        self.log("Extraction PDF basique non disponible sans pdfplumber", "ERROR")
        return None

    # -------- Salesforce field targeting (amélioré pour champ exact Compte) --------
    def _find_account_edit_button(self):
        # Fallback 1: édition inline du champ Compte
        selectors = [
            "//button[@title='Modifier Compte']",
            "//button[contains(@title,'Modifier') and contains(@title,'Compte')]",
            "//button[contains(@aria-label,'Modifier') and contains(@aria-label,'Compte')]",
            "//records-record-layout-item//*[self::span or self::label][normalize-space()='Compte']/ancestor::records-record-layout-item[1]//button[contains(@title,'Modifier')]",
            "//records-record-layout-item[.//*[contains(normalize-space(),'Compte')]]//button[contains(@title,'Modifier') or contains(@aria-label,'Modifier')]",
            "//*[contains(@class,'slds-form-element')][.//*[self::span or self::label][normalize-space()='Compte']]//button[contains(@class,'slds-button_icon') or contains(@title,'Modifier')]",
        ]
        for selector in selectors:
            try:
                btn = self.driver.find_element(By.XPATH, selector)
                if btn.is_displayed() and btn.is_enabled():
                    btn.click()
                    time.sleep(0.8)
                    self.log(f"Édition Compte ouverte via: {selector}")
                    return btn
            except Exception:
                continue

        # Fallback 2: ouvrir le mode édition global de l'enregistrement
        global_edit_selectors = [
            "//button[@name='Edit']",
            "//button[contains(@title,'Modifier')]",
            "//button[contains(normalize-space(),'Modifier')]",
            "//*[@role='button'][contains(@aria-label,'Modifier')]",
        ]
        for selector in global_edit_selectors:
            try:
                btn = WebDriverWait(self.driver, 4).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                try:
                    btn.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", btn)
                time.sleep(1)
                self.log(f"Mode édition global ouvert via: {selector}")
                return btn
            except Exception:
                continue

        return None

    def _find_account_lookup_input(self):
        selectors = [
            "//input[@placeholder='Rechercher des comptes...']",
            "//input[contains(@aria-label,'Compte') and (@type='text' or @type='search')]",
            "//input[contains(@name,'Account') and (@type='text' or @type='search')]",
            "//records-record-layout-item[.//*[normalize-space()='Compte']]//input",
            "//records-record-layout-item[.//*[contains(normalize-space(),'Compte')]]//input",
            "//*[contains(@class,'slds-form-element')][.//*[self::span or self::label][normalize-space()='Compte']]//input",
            "//*[contains(@class,'slds-form-element')][.//*[contains(normalize-space(),'Compte')]]//input",
            # Mode édition global (formulaire modal)
            "//div[contains(@class,'modal-container')]//*[contains(@class,'slds-form-element')][.//*[contains(normalize-space(),'Compte')]]//input",
            "//div[contains(@class,'modal-container')]//input[contains(@placeholder,'Rechercher') and (contains(@placeholder,'compte') or contains(@placeholder,'compte'))]",
            "//lightning-grouped-combobox//input[contains(@placeholder,'Rechercher') and contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'compte')]",
            "//input[@role='combobox' and (@type='text' or @type='search')]",
        ]
        for selector in selectors:
            try:
                field = WebDriverWait(self.driver, 3).until(
                    EC.visibility_of_element_located((By.XPATH, selector))
                )
                if field.is_enabled():
                    self.log(f"Champ Compte exact trouvé via: {selector}")
                    return field
            except Exception:
                continue
        return None

    def _clear_account_field(self) -> bool:
        clear_selectors = [
            "//button[@title='Effacer la sélection Compte']",
            "//button[contains(@title,'Effacer') and contains(@title,'Compte')]",
            "//*[contains(@class,'slds-form-element')][.//*[normalize-space()='Compte']]//button[.//*[local-name()='svg'][@data-key='close']]",
        ]
        for selector in clear_selectors:
            try:
                btn = self.driver.find_element(By.XPATH, selector)
                if btn.is_displayed() and btn.is_enabled():
                    btn.click()
                    time.sleep(0.4)
                    return True
            except Exception:
                continue

        field = self._find_account_lookup_input()
        if not field:
            return False
        try:
            field.click()
            field.send_keys(Keys.CONTROL + "a")
            field.send_keys(Keys.DELETE)
            return True
        except Exception:
            return False

    def _type_account(self, account_name: str) -> bool:
        field = self._find_account_lookup_input()
        if not field:
            return False
        try:
            field.click()
            field.send_keys(Keys.CONTROL + "a")
            field.send_keys(Keys.DELETE)
            field.send_keys(account_name)
            return True
        except Exception:
            return False

    def _select_dropdown_result(self, account_name: str) -> bool:
        exact_text_selectors = [
            f"//lightning-base-combobox-item//*[normalize-space()='{account_name}']/ancestor::lightning-base-combobox-item[1]",
            f"//li[@role='presentation' or @role='option'][.//*[normalize-space()='{account_name}']]",
        ]
        for selector in exact_text_selectors:
            try:
                el = WebDriverWait(self.driver, 4).until(EC.element_to_be_clickable((By.XPATH, selector)))
                el.click()
                return True
            except Exception:
                continue

        fallback_selectors = [
            "//lightning-base-combobox-item[1]",
            "//ul[contains(@class,'slds-listbox')]//li[1]",
            "//*[@role='listbox']//*[@role='option'][1]",
        ]
        for selector in fallback_selectors:
            try:
                el = WebDriverWait(self.driver, 4).until(EC.element_to_be_clickable((By.XPATH, selector)))
                el.click()
                return True
            except Exception:
                continue
        return False

    def _save_record(self) -> bool:
        save_selectors = [
            "//button[@name='SaveEdit']",
            "//button[@name='Save']",
            "//button[@title='Enregistrer' or @title='Save']",
            "//button[normalize-space()='Enregistrer' or normalize-space()='Save']",
        ]
        for selector in save_selectors:
            try:
                btn = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, selector)))
                btn.click()
                time.sleep(1.5)
                return True
            except Exception:
                continue
        return False

    def update_account_in_salesforce(self, devis_url: str, account_name: str) -> bool:
        try:
            self.log("Navigation vers le devis...")
            self.driver.get(devis_url)
            self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(2)

            self.log("Recherche du champ Compte...")
            edit_btn = self._find_account_edit_button()
            if not edit_btn:
                self.log("Champ Compte introuvable (mode édition)", "ERROR")
                return False

            if not self._clear_account_field():
                self.log("Impossible d'effacer le compte existant", "ERROR")
                return False

            self.log(f"Saisie du compte: '{account_name}'...")
            if not self._type_account(account_name):
                self.log("Impossible de saisir le compte", "ERROR")
                return False

            self.log("Sélection du résultat...")
            if not self._select_dropdown_result(account_name):
                self.log("Aucun résultat à sélectionner", "ERROR")
                return False

            self.log("Enregistrement...")
            if self._save_record():
                self.log(f"✅ Compte mis à jour avec succès: '{account_name}'", "SUCCESS")
                return True
            self.log("Erreur lors de l'enregistrement", "ERROR")
            return False

        except Exception as e:
            self.log(f"Erreur mise à jour Salesforce: {e}", "ERROR")
            self.log(traceback.format_exc(), "DEBUG")
            return False

    def _click_po_file(self) -> bool:
        def _is_po_name(raw: str) -> bool:
            if not raw:
                return False
            name = raw.strip().upper()
            return name.startswith("PO") and ".PDF" in name

        def _click_element(el) -> bool:
            try:
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
                time.sleep(0.2)
            except Exception:
                pass

            try:
                el.click()
                return True
            except Exception:
                try:
                    self.driver.execute_script("arguments[0].click();", el)
                    return True
                except Exception:
                    return False

        def _collect_candidates():
            selectors = [
                # Carte "Fichiers" (celle visible sur ta capture)
                "//span[normalize-space()='Fichiers']/ancestor::article[1]//a",
                "//span[starts-with(normalize-space(),'Fichiers')]/ancestor::article[1]//a",
                # Liens document Salesforce
                "//a[contains(@href,'ContentDocument') or contains(@href,'sfc/servlet.shepherd') or contains(@href,'/sfc/') ]",
                # Noms de fichiers PO
                "//a[starts-with(normalize-space(@title),'PO') or starts-with(normalize-space(text()),'PO')]",
            ]

            candidates = []
            for selector in selectors:
                try:
                    candidates.extend(self.driver.find_elements(By.XPATH, selector))
                except Exception:
                    continue
            return candidates

        try:
            self.log("Recherche PO dans la page (zone Fichiers + liens documents)...")
            self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

            for attempt in range(1, 6):
                self.check_stop()
                if attempt > 1:
                    time.sleep(1)

                candidates = _collect_candidates()
                if not candidates:
                    self.log(f"Aucun lien fichier détecté (tentative {attempt}/5)", "DEBUG")
                    continue

                for el in candidates:
                    try:
                        text = (el.get_attribute("title") or el.text or "").strip()
                        href = (el.get_attribute("href") or "").strip()
                    except StaleElementReferenceException:
                        # Lightning rerender la zone: on ignore cet élément
                        continue

                    if not _is_po_name(text):
                        if not (text.upper().startswith("PO") and href):
                            continue

                    self.log(f"Fichier PO trouvé: {text or href}")
                    if _click_element(el):
                        time.sleep(2)
                        return True

                # Fallback utile Salesforce: cliquer "Afficher tout" dans Fichiers
                try:
                    show_all = self.driver.find_element(
                        By.XPATH,
                        "//span[starts-with(normalize-space(),'Fichiers')]/ancestor::article[1]//a[contains(.,'Afficher tout')]",
                    )
                    if _click_element(show_all):
                        time.sleep(1)
                except Exception:
                    pass

                self.log(f"Aucun PO détecté (tentative {attempt}/5)", "DEBUG")

            self.log("Fichier PO non trouvé avec les sélecteurs standards", "WARNING")
            return False

        except Exception as e:
            self.log(f"Erreur recherche PO: {e}", "ERROR")
            return False

    def _download_po_pdf(self) -> str | None:
        try:
            # Dans Salesforce, après clic sur le fichier, on peut avoir:
            # - une preview avec bouton Télécharger
            # - un lien direct qui déclenche déjà le download
            # -> on tente d'abord les boutons explicites, puis on attend le fichier.
            download_selectors = [
                "//button[contains(@title,'Télécharger') or contains(@title,'Download')]",
                "//a[contains(@title,'Télécharger') or contains(@title,'Download')]",
                "//button[contains(normalize-space(),'Télécharger') or contains(normalize-space(),'Download')]",
                "//button[@name='download' or @data-key='download']",
                "//*[@role='button' and (contains(@aria-label,'Télécharger') or contains(@aria-label,'Download'))]",
            ]

            clicked = False
            for selector in download_selectors:
                try:
                    btn = WebDriverWait(self.driver, SHORT_WAIT).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    try:
                        btn.click()
                    except Exception:
                        self.driver.execute_script("arguments[0].click();", btn)
                    clicked = True
                    self.log(f"Bouton téléchargement cliqué via: {selector}")
                    break
                except Exception:
                    continue

            if not clicked:
                self.log("Aucun bouton Télécharger trouvé, attente d'un download direct...", "DEBUG")

            return self.wait_for_download()
        except Exception as e:
            self.log(f"Erreur téléchargement: {e}", "ERROR")
            return None

    def process_devis(self, devis_url: str) -> dict:
        result = {"url": devis_url, "success": False, "account_extracted": None, "error": None}
        try:
            self.check_stop()
            self.log(f"{'=' * 50}")
            self.log(f"Traitement du devis: {devis_url}")

            self.driver.get(devis_url)
            time.sleep(3)
            self.check_stop()

            self.log("Recherche du fichier PO...")
            if not self._click_po_file():
                result["error"] = "Fichier PO introuvable"
                return result

            pdf_path = self._download_po_pdf()
            if not pdf_path:
                result["error"] = "Téléchargement PDF échoué"
                return result

            account_name = self.extract_account_from_pdf(pdf_path)
            if not account_name:
                result["error"] = "Extraction compte échouée"
                return result

            result["account_extracted"] = account_name
            if self.update_account_in_salesforce(devis_url, account_name):
                result["success"] = True
            else:
                result["error"] = "Mise à jour Salesforce échouée"

        except InterruptedError:
            result["error"] = "Arrêt demandé"
            raise
        except Exception as e:
            result["error"] = str(e)
            self.log(f"Erreur inattendue: {e}", "ERROR")
        return result


def run_bot(devis_url: str, log_queue: queue.Queue, stop_event: threading.Event):
    bot = BotBYES360(log_queue, stop_event)
    try:
        log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "INFO", "message": "Vérification des dépendances..."})
        try:
            import pdfplumber as _pdfplumber
            globals()["pdfplumber"] = _pdfplumber
        except ImportError:
            log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "INFO", "message": "Installation de pdfplumber..."})
            os.system(f"{sys.executable} -m pip install pdfplumber -q")
            import pdfplumber as _pdfplumber
            globals()["pdfplumber"] = _pdfplumber

        if not bot.init_driver():
            log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "ERROR", "message": "Impossible d'initialiser le navigateur"})
            return

        result = bot.process_devis(devis_url)
        if result["success"]:
            log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "SUCCESS", "message": f"✅ SUCCÈS — Compte: {result['account_extracted']}"})
        else:
            log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "ERROR", "message": f"❌ ÉCHEC — {result['error']}"})

    except InterruptedError:
        log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "WARNING", "message": "⛔ Bot arrêté par l'utilisateur"})
    except Exception as e:
        log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "ERROR", "message": f"Erreur critique: {e}"})
    finally:
        bot.close_driver()
        log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "DONE", "message": "DONE"})
