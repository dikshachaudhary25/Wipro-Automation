"""AutoSys start/end time updater.

This script does the following:
1. Reads the job workbook.
2. Finds each job name in the Excel report.
3. Opens the AutoSys WCC page.
4. Searches for that job in the Name field.
5. Reads the result Start Time and End Time.
6. Writes them back to the Excel file.

Requirements:
    pip install selenium openpyxl

Notes:
- The URL and workbook path are intentionally configurable.
- The browser may require a corporate login from the same machine.
- The page DOM can change slightly depending on AutoSys version; the script tries
  a few common selectors and falls back to a text-based parse.
"""

from __future__ import annotations

import re
import time
import os
from pathlib import Path
from typing import Iterable, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


AUTOSYS_URL = "https://autosysproduction.molina.mhc/wcc/webui/#/"
WORKBOOK_PATH = r"C:\Users\Chaudh24\OneDrive - Molina Healthcare\Desktop\Molina\TransactionAutomation\jobs_updated.xlsx"
SERVER_NAME = "Autosys_PRO"
JOB_NAME_COLUMN = "Job Name"
START_TIME_COLUMN = "Start Time"
END_TIME_COLUMN = "End Time"
FIRST_DATA_ROW = 68
LAST_DATA_ROW = 124


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def get_worksheet(path: str):
    wb = load_workbook(path, data_only=False)
    required_headers = {JOB_NAME_COLUMN.lower(), START_TIME_COLUMN.lower(), END_TIME_COLUMN.lower()}
    for ws in wb.worksheets:
        headers = {normalize_text(cell.value).lower() for cell in ws[1]}
        if required_headers.issubset(headers):
            return wb, ws
    raise ValueError(f"Could not find a worksheet containing {sorted(required_headers)} in workbook: {path}")


def find_column_index(row_values: Iterable[object], target: str) -> int:
    for idx, value in enumerate(row_values):
        if normalize_text(value).lower() == target.lower():
            return idx
    return -1


def read_jobs_from_excel(path: str):
    """Return rows with job name + workbook row index."""
    wb, ws = get_worksheet(path)
    headers = [normalize_text(cell.value) for cell in ws[1]]
    job_idx = find_column_index(headers, JOB_NAME_COLUMN)
    start_idx = find_column_index(headers, START_TIME_COLUMN)
    end_idx = find_column_index(headers, END_TIME_COLUMN)

    if job_idx == -1:
        raise ValueError(f"Could not find '{JOB_NAME_COLUMN}' column in workbook: {path}")

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        if not FIRST_DATA_ROW <= row[0].row <= LAST_DATA_ROW:
            continue
        job_name = normalize_text(row[job_idx].value)
        if not job_name:
            continue
        rows.append({
            "row_number": row[0].row,
            "job_name": job_name,
            "start_cell": row[start_idx] if start_idx >= 0 else None,
            "end_cell": row[end_idx] if end_idx >= 0 else None,
        })

    return wb, ws, rows


def build_driver() -> webdriver.Chrome:
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    # Headless is disabled by default because the AutoSys app is often SSO/cookie based
    # and behaves more reliably in a visible browser session.

    driver = webdriver.Chrome(options=chrome_options)
    driver.get(AUTOSYS_URL)
    return driver


def wait_for_search_ui(driver: webdriver.Chrome, timeout: int = 30):
    # Automatic SSO may take a few seconds; the Quick View link is the authenticated-page signal.
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH, "//a[contains(normalize-space(.), 'Quick View')]") )
    )


def open_quick_view(driver: webdriver.Chrome, timeout: int = 30):
    """Open AutoSys Quick View, where the job Name search is available."""
    driver.switch_to.default_content()
    quick_view = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(normalize-space(.), 'Quick View')]"))
    )
    quick_view.click()
    time.sleep(3)


def find_server_input(driver: webdriver.Chrome):
    candidates = [
        (By.XPATH, "//input[contains(@value, 'Autosys') or @aria-label='Server']"),
        (By.XPATH, "//div[@role='combobox' and .//text()='Server']"),
        (By.XPATH, "//input[@role='combobox']"),
        (By.CSS_SELECTOR, "input.MuiSelect-nativeInput"),
    ]
    for by, value in candidates:
        try:
            element = driver.find_element(by, value)
            if element:
                return element
        except NoSuchElementException:
            continue
    return None


def find_search_input(driver: webdriver.Chrome):
    candidates = [
        (By.ID, "search-autocomplete-input"),
        (By.XPATH, "//input[@placeholder='Search here for job names']"),
        (By.XPATH, "//input[contains(@placeholder, 'job names')]"),
        (By.XPATH, "//input[@name='name' or @id='name']"),
    ]
    for by, value in candidates:
        try:
            element = driver.find_element(by, value)
            if element:
                return element
        except NoSuchElementException:
            continue

    # Fall back to the embedded search fields in the iframe if the top-level page doesn't expose them.
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for iframe in iframes:
        try:
            driver.switch_to.frame(iframe)
            element = find_search_input(driver)
            if element is not None:
                return element
        except Exception:
            pass
        finally:
            try:
                driver.switch_to.default_content()
            except Exception:
                pass

    return None


def switch_to_job_view(driver: webdriver.Chrome) -> bool:
    """Switch into the embedded monitoring view containing the Name search."""
    driver.switch_to.default_content()
    for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
        driver.switch_to.frame(iframe)
        try:
            inputs = [input_element for input_element in driver.find_elements(By.TAG_NAME, "input")
                      if input_element.is_displayed() and input_element.is_enabled()]
            if inputs:
                return True
        except NoSuchElementException:
            pass
        driver.switch_to.default_content()
    return False


def find_selections_and_search_box(driver):
    """Returns a tuple of (server dropdown, name input, go button) when found."""
    server_dropdown = find_server_input(driver)
    name_input = None
    go_button = None

    if switch_to_job_view(driver):
        name_candidates = [
            (By.XPATH, "//input[@aria-label='Name:' or @placeholder='Name:']"),
            (By.XPATH, "//label[contains(normalize-space(.), 'Name:')]/following::input[1]"),
        ]
        for by, value in name_candidates:
            try:
                name_input = driver.find_element(by, value)
                break
            except NoSuchElementException:
                continue

        if name_input is None:
            text_inputs = [input_element for input_element in driver.find_elements(
                By.CSS_SELECTOR, "input[type='text'], input:not([type])"
            ) if input_element.is_displayed() and input_element.is_enabled()]
            if text_inputs:
                name_input = text_inputs[0]

        for by, value in [
            (By.XPATH, "//button[normalize-space(.)='Go']"),
            (By.XPATH, "//button[contains(normalize-space(.), 'Go')]"),
        ]:
            try:
                go_button = driver.find_element(by, value)
                break
            except NoSuchElementException:
                continue

    if name_input is None:
        driver.switch_to.default_content()
        name_input = find_search_input(driver)

    return server_dropdown, name_input, go_button


def set_server(driver: webdriver.Chrome, server_name: str):
    server_input = find_server_input(driver)
    if server_input is None:
        return

    current_value = normalize_text(server_input.get_attribute("value"))
    if current_value and current_value.lower() != server_name.lower():
        server_input.clear()
        server_input.send_keys(server_name)
        time.sleep(1)
        # If a drop-down suggestion appears, click the exact match if available.
        for option in driver.find_elements(By.XPATH, "//*[contains(normalize-space(.), 'Autosys')]"):
            if normalize_text(option.text).lower() == server_name.lower():
                option.click()
                break


def search_job(driver: webdriver.Chrome, job_name: str):
    _, name_input, go_button = find_selections_and_search_box(driver)
    if name_input is None:
        raise RuntimeError("Could not locate the AutoSys name search box.")

    name_input.clear()
    name_input.send_keys(job_name)
    time.sleep(1)

    if go_button is not None:
        go_button.click()
    else:
        name_input.send_keys("\n")

    time.sleep(5)
    open_job_details(driver, job_name)


def open_job_details(driver: webdriver.Chrome, job_name: str):
    """Open the matching Quick View result so its execution fields are available."""
    job_name_xpath = f"//*[normalize-space()={job_name!r}]"
    candidates = driver.find_elements(By.XPATH, job_name_xpath)
    for candidate in candidates:
        if candidate.is_displayed() and candidate.is_enabled():
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", candidate)
            try:
                candidate.click()
            except Exception:
                driver.execute_script("arguments[0].click();", candidate)
            time.sleep(3)
            return

    # Some grid implementations put the name inside a clickable row.
    for row in driver.find_elements(By.XPATH, "//*[@role='row']"):
        if job_name.lower() in normalize_text(row.text).lower():
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", row)
            row.click()
            time.sleep(3)
            return

    # Legacy AutoSys tables may not expose ARIA rows.
    for row in driver.find_elements(By.XPATH, "//tr"):
        if job_name.lower() in normalize_text(row.text).lower():
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", row)
            row.click()
            time.sleep(3)
            return

    raise RuntimeError(f"AutoSys result was not found for job: {job_name}")


def click_job_runs(driver: webdriver.Chrome):
    candidates = driver.find_elements(By.XPATH, "//*[normalize-space()='Job Runs']")
    for candidate in candidates:
        if candidate.is_displayed() and candidate.is_enabled():
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", candidate)
            candidate.click()
            time.sleep(2)
            return
    raise RuntimeError("Could not locate the AutoSys Job Runs tab.")


def extract_time_from_text(text: str) -> Optional[str]:
    match = re.search(r"(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2})", text)
    if match:
        return match.group(1)

    match = re.search(r"([A-Z][a-z]{2}\s\d{1,2},\s\d{4}\s\d{1,2}:\d{2}:\d{2}\s[A|P]M)", text)
    if match:
        return match.group(1)

    return None


def get_job_details_from_table(driver: webdriver.Chrome, job_name: str) -> Tuple[Optional[str], Optional[str]]:
    job_name_lc = job_name.lower()

    def read_rows() -> Tuple[Optional[str], Optional[str], bool]:
        for table in driver.find_elements(By.TAG_NAME, "table"):
            rows = table.find_elements(By.TAG_NAME, "tr")
            if len(rows) < 2:
                continue
            header_cells = rows[0].find_elements(By.TAG_NAME, "th") or rows[0].find_elements(By.TAG_NAME, "td")
            headers = [normalize_text(cell.text).lower() for cell in header_cells]
            start_idx = next((i for i, value in enumerate(headers) if value == "start time"), -1)
            end_idx = next((i for i, value in enumerate(headers) if value == "end time"), -1)
            status_idx = next((i for i, value in enumerate(headers) if "status" in value), -1)
            exit_idx = next((i for i, value in enumerate(headers) if "exit code" in value), -1)
            if start_idx < 0 or end_idx < 0:
                continue

            for row in rows[1:]:
                values = [normalize_text(cell.text) for cell in row.find_elements(By.TAG_NAME, "td")]
                if len(values) <= max(start_idx, end_idx) or job_name_lc not in " | ".join(values).lower():
                    continue
                start_value = extract_time_from_text(values[start_idx])
                end_value = extract_time_from_text(values[end_idx])
                status = values[status_idx].lower() if 0 <= status_idx < len(values) else ""
                exit_code = values[exit_idx].strip() if 0 <= exit_idx < len(values) else ""
                successful = "success" in status or "complete" in status or exit_code == "0"
                return start_value, end_value, successful
        return None, None, False

    start_value, end_value, successful = read_rows()
    if start_value is not None and end_value is not None and successful:
        return start_value, end_value

    # For running, held, terminated, or scheduled jobs, use the newest completed run.
    click_job_runs(driver)
    successful_runs = []
    for table in driver.find_elements(By.TAG_NAME, "table"):
        rows = table.find_elements(By.TAG_NAME, "tr")
        if len(rows) < 2:
            continue
        header_cells = rows[0].find_elements(By.TAG_NAME, "th") or rows[0].find_elements(By.TAG_NAME, "td")
        headers = [normalize_text(cell.text).lower() for cell in header_cells]
        start_idx = next((i for i, value in enumerate(headers) if "start time" in value), -1)
        end_idx = next((i for i, value in enumerate(headers) if "end time" in value), -1)
        status_idx = next((i for i, value in enumerate(headers) if "status" in value), -1)
        if start_idx < 0 or end_idx < 0:
            continue
        for row in rows[1:]:
            values = [normalize_text(cell.text) for cell in row.find_elements(By.TAG_NAME, "td")]
            if len(values) <= max(start_idx, end_idx):
                continue
            status = values[status_idx].lower() if 0 <= status_idx < len(values) else ""
            start_value = extract_time_from_text(values[start_idx])
            end_value = extract_time_from_text(values[end_idx])
            if start_value and end_value and ("success" in status or "complete" in status or not status):
                successful_runs.append((start_value, end_value))

    if not successful_runs:
        return None, None

    def sort_key(run):
        value = run[0]
        for fmt in ("%b %d, %Y %I:%M:%S %p", "%Y-%m-%d %H:%M:%S"):
            try:
                from datetime import datetime
                return datetime.strptime(value, fmt)
            except ValueError:
                continue
        return datetime.min

    return max(successful_runs, key=sort_key)


def write_back_times(ws, row_index: int, start_value: Optional[str], end_value: Optional[str]):
    if start_value is not None:
        ws.cell(row=row_index, column=9, value=start_value)
    if end_value is not None:
        ws.cell(row=row_index, column=10, value=end_value)


def main():
    wb, ws, rows = read_jobs_from_excel(WORKBOOK_PATH)
    target_job = normalize_text(os.environ.get("START_ENDTIME_JOB"))
    if target_job:
        rows = [row for row in rows if row["job_name"].lower() == target_job.lower()]
        if not rows:
            raise ValueError(f"Job was not found in workbook: {target_job}")
    driver = build_driver()

    try:
        wait_for_search_ui(driver)
        open_quick_view(driver)
        set_server(driver, SERVER_NAME)
        if not switch_to_job_view(driver):
            raise RuntimeError("Could not locate the embedded AutoSys job search view.")

        for row in rows:
            job_name = row["job_name"]
            try:
                search_job(driver, job_name)
                start_time, end_time = get_job_details_from_table(driver, job_name)

                # Keep the workbook modification as a simple direct write.
                if start_time is not None:
                    ws.cell(row=row["row_number"], column=9, value=start_time)
                if end_time is not None:
                    ws.cell(row=row["row_number"], column=10, value=end_time)

                print(f"{job_name}: start={start_time}, end={end_time}")
                time.sleep(1)
            except Exception as exc:
                print(f"FAILED FOR {job_name}: {exc}")
                continue

        wb.save(WORKBOOK_PATH)
        print("Workbook updated successfully.")
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
