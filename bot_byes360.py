"""
Bot BYES 360 - Automatisation Prise de Commande
Extrait le compte depuis le PDF PO et le met à jour dans Salesforce
"""

import os
import re
import sys
import time
import glob
import queue
import threading
import traceback
from pathlib import Path
from datetime import datetime

# PDF extraction
try:
    import pdfplumber
except ImportError:
    pdfplumber = None

# Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
    ElementClickInterceptedException, StaleElementReferenceException
)

# ─── CONFIG ───────────────────────────────────────────────────────────────────
DOWNLOAD_DIR = str(Path.home() / "Downloads")
WAIT_TIMEOUT = 20
SHORT_WAIT = 5
BROWSER = "edge"  # "edge" ou "chrome"
# ──────────────────────────────────────────────────────────────────────────────


class BotBYES360:
    def __init__(self, log_queue: queue.Queue, stop_event: threading.Event):
        self.log_queue = log_queue
        self.stop_event = stop_event
        self.driver = None
        self.wait = None

    def log(self, message: str, level: str = "INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        entry = {"time": timestamp, "level": level, "message": message}
        self.log_queue.put(entry)
        print(f"[{timestamp}] [{level}] {message}")

    def check_stop(self):
        if self.stop_event.is_set():
            raise InterruptedError("⛔ Arrêt demandé par l'utilisateur.")

    def init_driver(self):
        self.log("Initialisation Edge avec ton profil existant...")
        self.log("⚠️ Ferme toutes les fenetres Edge avant de continuer !", "WARNING")
        try:
            opts = EdgeOptions()
            user_data = os.path.join(os.environ.get("LOCALAPPDATA", ""),
                                     "Microsoft", "Edge", "User Data")
            opts.add_argument(f"--user-data-dir={user_data}")
            opts.add_argument("--profile-directory=Default")
            opts.add_argument("--start-maximized")
            opts.add_argument("--no-first-run")
            opts.add_argument("--no-default-browser-check")
            opts.add_experimental_option("excludeSwitches", ["enable-automation"])
            opts.add_experimental_option("useAutomationExtension", False)
            prefs = {
                "download.default_directory": DOWNLOAD_DIR,
                "download.prompt_for_download": False,
                "plugins.always_open_pdf_externally": True,
            }
            opts.add_experimental_option("prefs", prefs)
            self.driver = webdriver.Edge(options=opts)
            self.log("Edge ouvert avec ton profil (cookies/session conserves) ✅")
            self.wait = WebDriverWait(self.driver, WAIT_TIMEOUT)
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

    def wait_for_download(self, timeout=30) -> str | None:
        """Attend qu'un nouveau PDF apparaisse dans Downloads et retourne son chemin."""
        self.log("Attente du téléchargement PDF...")
        before = set(glob.glob(os.path.join(DOWNLOAD_DIR, "PO*.pdf")))
        start = time.time()
        while time.time() - start < timeout:
            self.check_stop()
            time.sleep(1)
            after = set(glob.glob(os.path.join(DOWNLOAD_DIR, "PO*.pdf")))
            new_files = after - before
            if new_files:
                # Attendre que le fichier soit complètement téléchargé
                filepath = list(new_files)[0]
                time.sleep(2)
                self.log(f"PDF téléchargé: {os.path.basename(filepath)} ✅")
                return filepath
        self.log("Timeout: aucun PDF téléchargé", "WARNING")
        return None

    def extract_account_from_pdf(self, pdf_path: str) -> str | None:
        """Extrait le nom du compte (adresse de facturation) depuis le PDF PO."""
        self.log(f"Extraction du compte depuis le PDF...")

        if pdfplumber is None:
            self.log("pdfplumber non installé, tentative extraction basique", "WARNING")
            return self._extract_account_basic(pdf_path)

        try:
            with pdfplumber.open(pdf_path) as pdf:
                page = pdf.pages[0]
                text = page.extract_text()

                if not text:
                    self.log("PDF vide ou non lisible", "ERROR")
                    return None

                # Chercher la section "Adresse de facturation"
                lines = text.split('\n')
                found_billing = False
                account_name = None

                for i, line in enumerate(lines):
                    line_clean = line.strip()

                    if "adresse de facturation" in line_clean.lower():
                        found_billing = True
                        continue

                    if found_billing and line_clean:
                        # La première ligne non vide après "Adresse de facturation"
                        # contient souvent un code puis le nom (ex: "7020 - CIPOSTE SAS (...)")
                        # On nettoie pour extraire le nom
                        account_name = self._clean_account_name(line_clean)
                        if account_name:
                            self.log(f"Compte extrait: '{account_name}' ✅")
                            return account_name

                if not account_name:
                    self.log("Compte non trouvé dans le PDF", "WARNING")
                    # Fallback: chercher pattern connu
                    return self._extract_account_fallback(text)

        except Exception as e:
            self.log(f"Erreur extraction PDF: {e}", "ERROR")
            return None

    def _clean_account_name(self, raw: str) -> str | None:
        """Nettoie le nom du compte extrait."""
        # Supprimer les codes numériques en début (ex: "7020 - ")
        cleaned = re.sub(r'^\d+\s*[-–]\s*', '', raw).strip()
        # Supprimer les parenthèses et leur contenu
        cleaned = re.sub(r'\(.*?\)', '', cleaned).strip()
        # Garder seulement si assez long
        if len(cleaned) >= 3:
            return cleaned
        return None

    def _extract_account_fallback(self, text: str) -> str | None:
        """Fallback: cherche des patterns connus dans le texte."""
        # Pattern: ligne après "Adresse de facturation" avec majuscules
        pattern = r'Adresse de facturation\s*\n\s*(\d+\s*[-–]\s*)?([A-Z][A-Z\s&]+(?:SAS|SA|SARL|SNC|SASU|GIE|GROUP|IMMO)?)'
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            name = match.group(2).strip() if match.group(2) else None
            if name:
                self.log(f"Compte extrait (fallback): '{name}'")
                return name
        return None

    def _extract_account_basic(self, pdf_path: str) -> str | None:
        """Extraction basique sans pdfplumber."""
        self.log("Extraction PDF basique non disponible sans pdfplumber", "ERROR")
        return None

    def update_account_in_salesforce(self, devis_url: str, account_name: str) -> bool:
        """Met à jour le champ Compte dans Salesforce."""
        try:
            self.log(f"Navigation vers le devis...")
            self.driver.get(devis_url)
            time.sleep(3)
            self.check_stop()

            # Attendre que la page charge
            self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(2)

            # Chercher le champ Compte - dans Salesforce Lightning il faut d'abord
            # trouver et cliquer l'icône de modification du champ Compte
            self.log("Recherche du champ Compte...")

            # Stratégie 1: Chercher via label "Compte"
            account_field = self._find_account_field()
            if not account_field:
                self.log("Champ Compte introuvable", "ERROR")
                return False

            self.check_stop()

            # Effacer la valeur existante
            self.log("Effacement du compte existant...")
            cleared = self._clear_account_field()
            if not cleared:
                self.log("Impossible d'effacer le compte existant", "ERROR")
                return False

            time.sleep(1)
            self.check_stop()

            # Taper le nouveau compte
            self.log(f"Saisie du compte: '{account_name}'...")
            typed = self._type_account(account_name)
            if not typed:
                self.log("Impossible de saisir le compte", "ERROR")
                return False

            time.sleep(2)
            self.check_stop()

            # Sélectionner le résultat dans la liste déroulante
            self.log("Sélection du résultat...")
            selected = self._select_dropdown_result()
            if not selected:
                self.log("Aucun résultat à sélectionner", "ERROR")
                return False

            time.sleep(1)
            self.check_stop()

            # Enregistrer
            self.log("Enregistrement...")
            saved = self._save_record()
            if saved:
                self.log(f"✅ Compte mis à jour avec succès: '{account_name}'", "SUCCESS")
                return True
            else:
                self.log("Erreur lors de l'enregistrement", "ERROR")
                return False

        except InterruptedError:
            raise
        except Exception as e:
            self.log(f"Erreur mise à jour Salesforce: {e}", "ERROR")
            self.log(traceback.format_exc(), "DEBUG")
            return False

    def _find_account_section(self):
        """Trouve le container Salesforce du champ Compte."""
        section_selectors = [
            "//records-record-layout-item[contains(@field-label, 'Compte')]",
            "//div[contains(@class,'slds-form-element') and .//span[normalize-space()='Compte']]",
            "//*[contains(@data-target-selection-name, 'AccountId') or contains(@data-target-selection-name, 'Account')]",
            "//flexipage-field[contains(@data-field-id, 'Account') or contains(@data-field-id, 'AccountId')]",
        ]
        for sel in section_selectors:
            try:
                el = self.driver.find_element(By.XPATH, sel)
                if el.is_displayed():
                    return el
            except Exception:
                continue
        return None

    def _find_account_field(self):
        """Trouve et clique sur le bouton d'édition du champ Compte."""
        account_section = self._find_account_section()
        scoped_edit_selectors = [
            ".//button[contains(@title,'Modifier') and (contains(@title,'Compte') or contains(@title,'Account'))]",
            ".//button[contains(@class,'inline-edit-trigger') or contains(@class,'slds-button_icon') ]",
            ".//button[contains(@aria-label,'Modifier') or contains(@aria-label,'Edit')]",
        ]

        if account_section:
            for sel in scoped_edit_selectors:
                try:
                    el = account_section.find_element(By.XPATH, sel)
                    if el.is_displayed() and el.is_enabled():
                        self.driver.execute_script("arguments[0].click();", el)
                        WebDriverWait(self.driver, SHORT_WAIT).until(
                            EC.presence_of_element_located((By.XPATH, "//input[contains(@placeholder,'Rechercher') or contains(@role,'combobox') or @type='search']"))
                        )
                        return el
                except Exception:
                    continue

        # fallback global
        selectors = [
            "//button[@title='Modifier Compte']",
            "//button[contains(@title,'Modifier') and contains(@title,'Compte')]",
            "//button[contains(@title,'Edit') and contains(@title,'Account')]",
            "//a[contains(@class,'outputLookupLink') and ancestor::*[contains(.,'Compte') or contains(.,'Account')]]",
        ]
        for sel in selectors:
            try:
                el = WebDriverWait(self.driver, SHORT_WAIT).until(EC.element_to_be_clickable((By.XPATH, sel)))
                self.driver.execute_script("arguments[0].click();", el)
                return el
            except Exception:
                continue

        return None

    def _clear_account_field(self) -> bool:
        """Efface la valeur du champ Compte sans impacter d'autres champs."""
        account_section = self._find_account_section()
        scoped_clear_selectors = [
            ".//button[contains(@title,'Effacer') and (contains(@title,'Compte') or contains(@title,'Account'))]",
            ".//button[contains(@aria-label,'Effacer') or contains(@aria-label,'Clear') or contains(@aria-label,'Remove')]",
            ".//*[local-name()='svg'][@data-key='close']/ancestor::button[1]",
            ".//button[contains(@class,'slds-input__icon')]",
        ]

        if account_section:
            for sel in scoped_clear_selectors:
                try:
                    btn = account_section.find_element(By.XPATH, sel)
                    if btn.is_displayed() and btn.is_enabled():
                        self.driver.execute_script("arguments[0].click();", btn)
                        time.sleep(0.3)
                        self.log("Compte effacé via bouton clear du champ ✅")
                        return True
                except Exception:
                    continue

            try:
                scoped_input = account_section.find_element(By.XPATH, ".//input[@type='text' or @type='search' or contains(@role,'combobox')]")
                scoped_input.click()
                scoped_input.send_keys(Keys.CONTROL + "a")
                scoped_input.send_keys(Keys.DELETE)
                return True
            except Exception:
                pass

        # fallback global
        for sel in [
            "//button[@title='Effacer la sélection Compte']",
            "//button[contains(@title,'Effacer') and contains(@title,'Compte')]",
            "//button[@title='Remove']",
        ]:
            try:
                btn = self.driver.find_element(By.XPATH, sel)
                if btn.is_displayed() and btn.is_enabled():
                    self.driver.execute_script("arguments[0].click();", btn)
                    return True
            except Exception:
                continue

        return False

    def _type_account(self, account_name: str) -> bool:
        """Tape le nom du compte dans le champ de recherche du lookup Compte."""
        account_section = self._find_account_section()
        input_selectors = [
            ".//input[@placeholder='Rechercher des comptes...']",
            ".//input[contains(@placeholder,'Rechercher') and (contains(@placeholder,'compte') or contains(@placeholder,'Account'))]",
            ".//input[contains(@class,'lookup__search-input') or contains(@class,'slds-combobox__input')]",
            ".//input[@type='search' or contains(@role,'combobox')]",
        ]

        containers = [account_section] if account_section else []
        containers.append(self.driver)

        for container in containers:
            for sel in input_selectors:
                try:
                    field = container.find_element(By.XPATH, sel)
                    if field.is_displayed() and field.is_enabled():
                        field.click()
                        field.send_keys(Keys.CONTROL + "a")
                        field.send_keys(Keys.DELETE)
                        field.send_keys(account_name)
                        return True
                except Exception:
                    continue

        return False

    def _select_dropdown_result(self) -> bool:
        """Sélectionne le premier résultat du lookup Compte."""
        account_section = self._find_account_section()
        result_selectors = [
            "//ul[contains(@class,'slds-listbox')]//li[1]",
            "//div[contains(@role,'listbox')]//*[contains(@role,'option')][1]",
            "//lightning-base-combobox-item[1]",
        ]

        for sel in result_selectors:
            try:
                result = WebDriverWait(self.driver, SHORT_WAIT).until(
                    EC.element_to_be_clickable((By.XPATH, sel))
                )
                if account_section and account_section.is_displayed():
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", result)
                self.driver.execute_script("arguments[0].click();", result)
                return True
            except Exception:
                continue

        try:
            active_input = self.driver.find_element(By.XPATH, "//input[@type='search' or contains(@role,'combobox')]")
            active_input.send_keys(Keys.ARROW_DOWN)
            active_input.send_keys(Keys.ENTER)
            return True
        except Exception:
            return False

    def _save_record(self) -> bool:
        """Clique sur le bouton Enregistrer."""
        save_selectors = [
            "//button[@name='SaveEdit']",
            "//button[text()='Enregistrer']",
            "//button[contains(@class,'slds-button') and contains(text(),'Enregistrer')]",
            "//button[@title='Enregistrer']",
            "//button[@name='SaveEdit' or normalize-space()='Save']",
        ]
        for sel in save_selectors:
            try:
                btn = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, sel))
                )
                btn.click()
                time.sleep(2)
                return True
            except Exception:
                continue
        return False

    def process_devis(self, devis_url: str) -> dict:
        """Traite un devis complet."""
        result = {
            "url": devis_url,
            "success": False,
            "account_extracted": None,
            "error": None
        }

        try:
            self.check_stop()
            self.log(f"{'='*50}")
            self.log(f"Traitement du devis: {devis_url}")

            # 1. Ouvrir le devis
            self.driver.get(devis_url)
            time.sleep(3)
            self.check_stop()

            # 2. Trouver et cliquer le fichier PO
            self.log("Recherche du fichier PO...")
            po_found = self._click_po_file()
            if not po_found:
                result["error"] = "Fichier PO introuvable"
                self.log("❌ Fichier PO introuvable", "ERROR")
                return result

            self.check_stop()

            # 3. Télécharger le PDF
            pdf_path = self._download_po_pdf()
            if not pdf_path:
                result["error"] = "Téléchargement PDF échoué"
                self.log("❌ Téléchargement PDF échoué", "ERROR")
                return result

            self.check_stop()

            # 4. Extraire le compte depuis le PDF
            account_name = self.extract_account_from_pdf(pdf_path)
            if not account_name:
                result["error"] = "Extraction compte échouée"
                self.log("❌ Impossible d'extraire le compte du PDF", "ERROR")
                return result

            result["account_extracted"] = account_name
            self.check_stop()

            # 5. Mettre à jour Salesforce
            # Revenir sur le devis
            self.driver.get(devis_url)
            time.sleep(3)

            updated = self.update_account_in_salesforce(devis_url, account_name)
            if updated:
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

    def _click_po_file(self) -> bool:
        """Trouve et ouvre le fichier PO dans la section Fichiers."""
        try:
            # Attendre que les fichiers soient visibles
            time.sleep(2)

            # Chercher le lien qui commence par "PO" dans la section Fichiers
            po_selectors = [
                "//a[starts-with(@title,'PO')]",
                "//span[starts-with(text(),'PO')]",
                "//a[contains(@href,'ContentDocument') and ancestor::*[contains(@class,'files') or contains(.,'Fichiers')]]",
                "//div[contains(@class,'slds-file')]//span[starts-with(text(),'PO')]",
            ]

            for sel in po_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, sel)
                    for el in elements:
                        text = el.get_attribute("title") or el.text
                        if text and text.upper().startswith("PO"):
                            self.log(f"Fichier PO trouvé: {text}")
                            el.click()
                            time.sleep(3)
                            return True
                except Exception:
                    continue

            self.log("Fichier PO non trouvé avec les sélecteurs standards", "WARNING")
            return False

        except Exception as e:
            self.log(f"Erreur recherche PO: {e}", "ERROR")
            return False

    def _download_po_pdf(self) -> str | None:
        """Télécharge le PDF PO depuis la prévisualisation Salesforce."""
        try:
            # Chercher le bouton Télécharger dans la preview
            download_selectors = [
                "//button[contains(@title,'Télécharger')]",
                "//a[contains(@title,'Télécharger')]",
                "//button[contains(text(),'Télécharger')]",
                "//a[contains(@download,'')]",
            ]

            for sel in download_selectors:
                try:
                    btn = WebDriverWait(self.driver, SHORT_WAIT).until(
                        EC.element_to_be_clickable((By.XPATH, sel))
                    )
                    btn.click()
                    break
                except Exception:
                    continue

            # Attendre le téléchargement
            return self.wait_for_download()

        except Exception as e:
            self.log(f"Erreur téléchargement: {e}", "ERROR")
            return None


def run_bot(devis_url: str, log_queue: queue.Queue, stop_event: threading.Event):
    """Fonction principale du bot, exécutée dans un thread séparé."""
    bot = BotBYES360(log_queue, stop_event)

    try:
        # Installer pdfplumber si nécessaire
        log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "INFO",
                       "message": "Vérification des dépendances..."})
        try:
            import pdfplumber
        except ImportError:
            log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "INFO",
                           "message": "Installation de pdfplumber..."})
            os.system(f"{sys.executable} -m pip install pdfplumber -q")
            import pdfplumber as pdfplumber_module
            import bot_byes360
            bot_byes360.pdfplumber = pdfplumber_module

        # Initialiser le driver
        if not bot.init_driver():
            log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "ERROR",
                           "message": "Impossible d'initialiser le navigateur"})
            log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "DONE",
                           "message": "DONE"})
            return

        # Traiter le devis
        result = bot.process_devis(devis_url)

        if result["success"]:
            log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "SUCCESS",
                           "message": f"✅ SUCCÈS — Compte: {result['account_extracted']}"})
        else:
            log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "ERROR",
                           "message": f"❌ ÉCHEC — {result['error']}"})

    except InterruptedError:
        log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "WARNING",
                       "message": "⛔ Bot arrêté par l'utilisateur"})
    except Exception as e:
        log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "ERROR",
                       "message": f"Erreur critique: {e}"})
    finally:
        bot.close_driver()
        log_queue.put({"time": datetime.now().strftime("%H:%M:%S"), "level": "DONE",
                       "message": "DONE"})
