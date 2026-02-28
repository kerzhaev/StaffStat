# -*- coding: utf-8 -*-
"""
WinAppDriver UI test: uf_Search — enter search term, click Export to Excel, verify file.
Requires: WinAppDriver at http://127.0.0.1:4723, StaffState (Access) open with uf_Search visible.

Pure Selenium 3.141.0 with JSON Wire Protocol - no W3C prefixes for WinAppDriver 1.2.1 compatibility.
"""
# Python 3.10+ compatibility: monkey-patch collections aliases removed from collections module
import collections
if not hasattr(collections, 'Mapping'):
    import collections.abc
    collections.Mapping = collections.abc.Mapping
if not hasattr(collections, 'Iterable'):
    import collections.abc
    collections.Iterable = collections.abc.Iterable

import os
import time
import glob
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException


WINAPPDRIVER_URL = "http://127.0.0.1:4723"
# Prefer exact titles; fallback: any window with "Access" in title
WINDOW_TITLES = ("StaffState", "Поиск сотрудников")
SEARCH_TEXT = "Иван"
EXPORT_BUTTON_NAME = "Export to Excel"
SEARCH_BOX_AUTOMATION_ID = "txtFilter"
BUTTON_AUTOMATION_ID = "btnExportExcel"
# Must match VBA: CurrentProject.Path + \\Exports. Set STAFFSTATE_EXPORTS_DIR to that path if DB is elsewhere.
EXPORTS_DIR = os.path.abspath(
    os.environ.get("STAFFSTATE_EXPORTS_DIR")
    or os.path.join(Path(__file__).resolve().parent.parent, "Exports")
)
FILE_WAIT_TIMEOUT_SEC = 15
FILE_MAX_AGE_SEC = 60


def find_app_window(driver, wait_sec=10):
    """Find StaffState or search form window by title; fallback: any window with 'Access' in title."""
    for title in WINDOW_TITLES:
        try:
            el = WebDriverWait(driver, wait_sec).until(
                EC.presence_of_element_located((By.NAME, title))
            )
            return el
        except TimeoutException:
            continue
    # Fallback: any window whose name contains "Access"
    try:
        for el in driver.find_elements(By.XPATH, "//*"):
            name = (el.get_attribute("Name") or el.text or "") if el else ""
            if "Access" in name:
                return el
    except Exception:
        pass
    raise NoSuchElementException(
        f"Window not found by any of: {WINDOW_TITLES} or with 'Access' in title. Open StaffState and uf_Search first."
    )


def test_ui_export():
    driver = None
    try:
        # Selenium 3.141.0: flat capabilities for JSON Wire Protocol (no appium: prefix)
        caps = {
            "app": "Root",
            "platformName": "Windows",
            "deviceName": "WindowsPC",
        }
        # Selenium 3 uses desired_capabilities parameter with JSON Wire Protocol by default
        driver = webdriver.Remote(
            command_executor=WINAPPDRIVER_URL,
            desired_capabilities=caps
        )
        driver.implicitly_wait(5)

        # 1) Find app window (StaffState or "Поиск сотрудников")
        window = find_app_window(driver)
        window.click()

        # 2) Search box (txtFilter): by AutomationId or Name
        try:
            # WinAppDriver uses AccessibilityId strategy for AutomationId attributes
            search_box = driver.find_element("accessibility id", SEARCH_BOX_AUTOMATION_ID)
        except NoSuchElementException:
            search_box = driver.find_element(By.NAME, SEARCH_BOX_AUTOMATION_ID)
        search_box.click()
        time.sleep(0.5)
        search_box.send_keys(Keys.CONTROL + "a")
        search_box.send_keys(Keys.BACKSPACE)
        search_box.send_keys(SEARCH_TEXT)
        time.sleep(1)

        # 3) Export button: by caption "Export to Excel" or AutomationId btnExportExcel
        try:
            export_btn = driver.find_element(By.NAME, EXPORT_BUTTON_NAME)
        except NoSuchElementException:
            # WinAppDriver uses AccessibilityId strategy for AutomationId attributes
            export_btn = driver.find_element("accessibility id", BUTTON_AUTOMATION_ID)
        export_btn.click()

        # 4) Wait for export (message box may appear; Excel may open)
        time.sleep(3)

        # 5) Verify file in Exports folder
        os.makedirs(EXPORTS_DIR, exist_ok=True)
        cutoff = time.time() - FILE_MAX_AGE_SEC
        pattern = os.path.join(EXPORTS_DIR, "*.xlsx")
        deadline = time.time() + FILE_WAIT_TIMEOUT_SEC
        found = []
        while time.time() < deadline:
            for path in glob.glob(pattern):
                if os.path.getmtime(path) >= cutoff:
                    found.append(path)
            if found:
                break
            time.sleep(1)

        assert found, (
            f"No .xlsx file created in {EXPORTS_DIR} in the last {FILE_MAX_AGE_SEC}s. "
            "Note: current app export opens Excel in memory; if your build saves to Exports, ensure path is set."
        )
        return found
    finally:
        if driver:
            driver.quit()


if __name__ == "__main__":
    files = test_ui_export()
    print("OK: Export file(s):", files)
