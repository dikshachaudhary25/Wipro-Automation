import os
import shutil
import zipfile
from datetime import datetime, timedelta

import openpyxl

# =====================================================
# CONFIGURATION
# =====================================================

ARCHIVE_PATH = r"\\molina.mhc\shares\Diplomat\Operations\BatchResults\all\M360DI\Archive"

EXCEL_FILE = r"C:\Users\Chaudh24\OneDrive - Molina Healthcare\Documents\Autosys Daily Status_09_23_2026.xlsx"

TEMP_FOLDER = r"C:\Users\Chaudh24\OneDrive - Molina Healthcare\Desktop\Molina\TransactionAutomation\temp"

SHEET_NAME = "Extracts Outbound & Inbound"

START_ROW = 68
END_ROW = 124

# =====================================================
# PREVIOUS DAY
# =====================================================

yesterday = (
    datetime.now() - timedelta(days=1)
).strftime("%Y%m%d")

print(f"\nPROCESSING DATE: {yesterday}")

# =====================================================
# HELPERS
# =====================================================

def should_ignore_file(file_name):

    file_name = file_name.lower()

    ignore_patterns = [
        ".pdf.",
        "report",
        "summary"
    ]

    return any(
        pattern in file_name
        for pattern in ignore_patterns
    )


def count_transactions(file_path, file_name):

    try:

        if should_ignore_file(file_name):
            return 0

        with open(
            file_path,
            "r",
            errors="ignore"
        ) as f:

            lines = sum(1 for _ in f)

        if lines <= 3:
            return 1

        return max(lines - 1, 0)

    except Exception:
        return 0


def count_written_lines(file_path, file_name):

    try:

        if should_ignore_file(file_name):
            return 0

        with open(
            file_path,
            "r",
            errors="ignore"
        ) as f:

            return sum(1 for _ in f)

    except Exception:
        return 0


def count_with_header(file_path, file_name):

    try:

        if should_ignore_file(file_name):
            return 0

        with open(
            file_path,
            "r",
            errors="ignore"
        ) as f:

            lines = [line.strip() for line in f if line.strip()]

        if not lines:
            return 0

        return max(len(lines) - 1, 0)

    except Exception:
        return 0


def get_zip_files(job_name):

    matches = []

    for item in os.listdir(ARCHIVE_PATH):

        if (
            job_name in item
            and yesterday in item
            and item.endswith(".zip")
        ):

            matches.append(item)

    return sorted(matches)

# =====================================================
# OPEN WORKBOOK
# =====================================================

wb = openpyxl.load_workbook(EXCEL_FILE)

ws = wb[SHEET_NAME]

# =====================================================
# FIND COLUMNS
# =====================================================

job_col = None
transaction_col = None

for row in range(1, 10):

    for col in range(1, ws.max_column + 1):

        value = ws.cell(
            row=row,
            column=col
        ).value

        if value is None:
            continue

        value = str(value)

        if value.strip() == "Job Name":
            job_col = col

        if "No. of Transactions for Yesterday" in value:
            transaction_col = col

if job_col is None:
    raise Exception("Job Name column not found")

if transaction_col is None:
    raise Exception("Transaction column not found")

# =====================================================
# BLANK EVERYTHING OUTSIDE RANGE
# =====================================================

for row in range(2, ws.max_row + 1):

    if row < START_ROW or row > END_ROW:

        ws.cell(
            row=row,
            column=transaction_col
        ).value = None

# =====================================================
# PROCESS JOBS
# =====================================================

for row in range(START_ROW, END_ROW + 1):

    job_name = ws.cell(
        row=row,
        column=job_col
    ).value

    if not job_name:
        continue

    job_name = str(job_name).strip()

    print("\n" + "=" * 70)
    print(f"ROW {row}")
    print(f"JOB {job_name}")

    zip_matches = get_zip_files(job_name)

    total_transactions = None

    # --------------------------------------------------
    # FIXED VALUE RULES
    # --------------------------------------------------

    if job_name == "elig_d_d_ftp_ib_all_mmp_m360_member_exception":
        total_transactions = 0

    elif job_name == "elig_d_d_ftp_ob_all_mapd_m360_connecture_dental_oec":
        total_transactions = 0

    elif job_name == "elig_d_d_ftp_ib_all_mmp_m360_mbr_comments":
        total_transactions = 3

    elif job_name == "elig_d_d_ftp_ib_mmp_rejectinputfile":
        total_transactions = 1

    # --------------------------------------------------
    # ZIP CONTENT FILE COUNT RULES
    # --------------------------------------------------

    elif job_name in [
        "elig_d_d_ftp_ib_all_mmp_m360_834_inc",
        "elig_d_d_ftp_ib_all_mmp_m360_marxcob",
        "elig_d_d_move_ob_all_mmp_m360_marxcob"
    ]:

        total_transactions = 0

        for zip_name in zip_matches:

            zip_path = os.path.join(
                ARCHIVE_PATH,
                zip_name
            )

            try:

                with zipfile.ZipFile(
                    zip_path,
                    "r"
                ) as zip_ref:

                    file_count = sum(
                        1
                        for f in zip_ref.namelist()
                        if not f.endswith("/")
                    )

                    total_transactions += file_count

            except Exception:
                pass

    # --------------------------------------------------
    # ZIP CONTENT FILE COUNT RULES
    # --------------------------------------------------

    elif job_name in [
        "elig_d_d_ftp_ib_all_mmp_m360_lockbox04",
        "elig_d_d_ftp_ib_all_mmp_m360_dental_response"
    ]:

        total_transactions = 0

        for zip_name in zip_matches:

            zip_path = os.path.join(
                ARCHIVE_PATH,
                zip_name
            )

            try:

                with zipfile.ZipFile(
                    zip_path,
                    "r"
                ) as zip_ref:

                    file_count = sum(
                        1
                        for f in zip_ref.namelist()
                        if not f.endswith("/")
                    )

                    total_transactions += file_count

            except Exception:
                pass

    # --------------------------------------------------
    # NO-HEADER FILE COUNT RULES
    # --------------------------------------------------

    elif job_name in [
        "elig_d_d_ib_ftp_mapd_DTRRD",
        "oth_d_d_ib_ftp_infoxing_MARXCOB"
    ]:

        total_transactions = 0

        if len(zip_matches) == 0:

            total_transactions = 0

        else:

            if os.path.exists(TEMP_FOLDER):

                shutil.rmtree(
                    TEMP_FOLDER,
                    ignore_errors=True
                )

            os.makedirs(
                TEMP_FOLDER,
                exist_ok=True
            )

            for zip_name in zip_matches:

                zip_path = os.path.join(
                    ARCHIVE_PATH,
                    zip_name
                )

                extract_folder = os.path.join(
                    TEMP_FOLDER,
                    zip_name.replace(".zip", "")
                )

                os.makedirs(
                    extract_folder,
                    exist_ok=True
                )

                try:

                    with zipfile.ZipFile(
                        zip_path,
                        "r"
                    ) as zip_ref:

                        zip_ref.extractall(
                            extract_folder
                        )

                except Exception:
                    continue

                for root, dirs, files in os.walk(
                    extract_folder
                ):

                    for file in files:

                        file_path = os.path.join(
                            root,
                            file
                        )

                        total_transactions += count_written_lines(
                            file_path,
                            file
                        )

    # --------------------------------------------------
    # HEADER-AWARE FILE COUNT RULES
    # --------------------------------------------------

    elif job_name == "elig_d_d_ftp_ib_all_mmp_m360_prov_response":

        total_transactions = 0

        if len(zip_matches) == 0:

            total_transactions = 0

        else:

            if os.path.exists(TEMP_FOLDER):

                shutil.rmtree(
                    TEMP_FOLDER,
                    ignore_errors=True
                )

            os.makedirs(
                TEMP_FOLDER,
                exist_ok=True
            )

            for zip_name in zip_matches:

                zip_path = os.path.join(
                    ARCHIVE_PATH,
                    zip_name
                )

                extract_folder = os.path.join(
                    TEMP_FOLDER,
                    zip_name.replace(".zip", "")
                )

                os.makedirs(
                    extract_folder,
                    exist_ok=True
                )

                try:

                    with zipfile.ZipFile(
                        zip_path,
                        "r"
                    ) as zip_ref:

                        zip_ref.extractall(
                            extract_folder
                        )

                except Exception:
                    continue

                for root, dirs, files in os.walk(
                    extract_folder
                ):

                    for file in files:

                        file_path = os.path.join(
                            root,
                            file
                        )

                        total_transactions += count_with_header(
                            file_path,
                            file
                        )

    # --------------------------------------------------
    # NORMAL PROCESSING
    # --------------------------------------------------

    if total_transactions is None:

        if len(zip_matches) == 0:

            total_transactions = 0

        else:

            if job_name in [
                "elig_d_d_ftp_ib_all_mmp_m360_extract1",
                "elig_d_d_ftp_ib_all_mmp_m360_extract2",
                "elig_d_d_ftp_ib_all_mmp_m360_agent_exception"
            ]:

                zip_matches = [zip_matches[-1]]

            if os.path.exists(TEMP_FOLDER):

                shutil.rmtree(
                    TEMP_FOLDER,
                    ignore_errors=True
                )

            os.makedirs(
                TEMP_FOLDER,
                exist_ok=True
            )

            total_transactions = 0

            for zip_name in zip_matches:

                zip_path = os.path.join(
                    ARCHIVE_PATH,
                    zip_name
                )

                extract_folder = os.path.join(
                    TEMP_FOLDER,
                    zip_name.replace(".zip", "")
                )

                os.makedirs(
                    extract_folder,
                    exist_ok=True
                )

                try:

                    with zipfile.ZipFile(
                        zip_path,
                        "r"
                    ) as zip_ref:

                        zip_ref.extractall(
                            extract_folder
                        )

                except Exception:
                    continue

                for root, dirs, files in os.walk(
                    extract_folder
                ):

                    for file in files:

                        file_path = os.path.join(
                            root,
                            file
                        )

                        total_transactions += count_transactions(
                            file_path,
                            file
                        )

    # --------------------------------------------------
    # ADJUSTMENTS
    # --------------------------------------------------

    if job_name == "elig_d_d_ftp_ob_all_mmp_m360_lockbox05":
        total_transactions -= 1

    if job_name == "elig_d_d_ftp_ib_all_mmp_m360_compliance_tracker":
        total_transactions -= 1

    if total_transactions < 0:
        total_transactions = 0

    # --------------------------------------------------
    # WRITE
    # --------------------------------------------------

    ws.cell(
        row=row,
        column=transaction_col
    ).value = total_transactions

    print(f"TOTAL = {total_transactions}")

# =====================================================
# SAVE
# =====================================================

output_file = (
    r"C:\Users\Chaudh24\OneDrive - Molina Healthcare\Desktop\Molina"
    r"\TransactionAutomation\jobs_updated.xlsx"
)

wb.save(output_file)

print("\nDONE")
print(f"\nSAVED TO:\n{output_file}")