import os
import re
import threading
import traceback
import unicodedata
from collections import defaultdict

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from openpyxl import load_workbook, Workbook
from rapidfuzz import fuzz, process


# ============================================================
# NAME FUZZY MATCH
# Final Desktop Application
# ============================================================


APP_TITLE = "Name Fuzzy Match"
OUTPUT_SHEET_NAME = "Sheet2"

OUTPUT_HEADERS = [
    "Document ID",
    "Full Name",
    "Unique ID",
    "Full Name Reverse - FN + LN",
    "Full Name Reverse - LN + FN",
    "Similarity Name",
    "Percentage % Matching",
    "Group Match Name",
]


# ============================================================
# NAME NORMALIZATION
# ============================================================

def normalize_name(value):
    """
    Converts a name into a consistent format for matching.

    Handles:
    - Upper/lower case
    - Extra spaces
    - Special characters
    - Unicode normalization
    - Basic punctuation
    """

    if value is None:
        return ""

    value = str(value).strip()

    if not value:
        return ""

    # Unicode normalization
    value = unicodedata.normalize("NFKD", value)

    # Remove accents/diacritics
    value = "".join(
        ch for ch in value
        if not unicodedata.combining(ch)
    )

    # Convert to uppercase
    value = value.upper()

    # Replace punctuation/special characters with spaces
    value = re.sub(r"[^A-Z0-9\s]", " ", value)

    # Collapse multiple spaces
    value = re.sub(r"\s+", " ", value).strip()

    return value


def get_tokens(name):
    """
    Returns individual words/components of a normalized name.
    """
    if not name:
        return []

    return name.split()


def reverse_name(name):
    """
    Reverse complete token order.

    Example:
        RAHUL SHARMA
        ->
        SHARMA RAHUL
    """

    tokens = get_tokens(name)

    if len(tokens) <= 1:
        return name

    return " ".join(reversed(tokens))


def first_last_name(name):
    """
    FN + LN representation.

    Example:
        Rahul Kumar Sharma
        -> Rahul Sharma

    For 2 words:
        Rahul Sharma
        -> Rahul Sharma
    """

    tokens = get_tokens(name)

    if not tokens:
        return ""

    if len(tokens) == 1:
        return tokens[0]

    return f"{tokens[0]} {tokens[-1]}"


def last_first_name(name):
    """
    LN + FN representation.

    Example:
        Rahul Kumar Sharma
        -> Sharma Rahul
    """

    tokens = get_tokens(name)

    if not tokens:
        return ""

    if len(tokens) == 1:
        return tokens[0]

    return f"{tokens[-1]} {tokens[0]}"


# ============================================================
# NAME SIMILARITY ENGINE
# ============================================================

def component_similarity(name1, name2):
    """
    Compares individual name components.

    This helps with:
        RAHUL SHARMA
        RAHULL SHARMA

    and prevents a shared first/last name alone from producing
    an excessively high score.
    """

    tokens1 = get_tokens(name1)
    tokens2 = get_tokens(name2)

    if not tokens1 or not tokens2:
        return 0.0

    # Compare each token from name1 with the best token in name2.
    scores = []

    for token1 in tokens1:
        best = 0

        for token2 in tokens2:
            score = fuzz.ratio(token1, token2)

            if score > best:
                best = score

        scores.append(best)

    if not scores:
        return 0.0

    # Average component quality
    average_score = sum(scores) / len(scores)

    return average_score


def position_aware_similarity(name1, name2):
    """
    Compares names while considering first and last components.

    Example:
        RAJESH VERMA
        RAJESH KUMAR

    Shared first name alone should not make the overall score
    extremely high.
    """

    tokens1 = get_tokens(name1)
    tokens2 = get_tokens(name2)

    if not tokens1 or not tokens2:
        return 0.0

    first_score = fuzz.ratio(tokens1[0], tokens2[0])

    last_score = fuzz.ratio(tokens1[-1], tokens2[-1])

    return (first_score * 0.50) + (last_score * 0.50)


def calculate_similarity(name1, name2):
    """
    Main matching algorithm.

    It evaluates:
        1. Exact normalized match
        2. Full-name similarity
        3. Token-sort similarity
        4. Reversed-name similarity
        5. First/last-name representation
        6. Individual component similarity
        7. Position-aware comparison

    The strongest appropriate score is returned.
    """

    if not name1 or not name2:
        return 0.0

    if name1 == name2:
        return 100.0

    reverse1 = reverse_name(name1)
    reverse2 = reverse_name(name2)

    fnln1 = first_last_name(name1)
    fnln2 = first_last_name(name2)

    lnfn1 = last_first_name(name1)
    lnfn2 = last_first_name(name2)

    # Basic full name
    ratio_score = fuzz.ratio(name1, name2)

    # Partial ratio
    partial_score = fuzz.partial_ratio(name1, name2)

    # Token based comparison
    token_sort_score = fuzz.token_sort_ratio(name1, name2)

    token_set_score = fuzz.token_set_ratio(name1, name2)

    # Reverse order
    reverse_score_1 = fuzz.ratio(name1, reverse2)
    reverse_score_2 = fuzz.ratio(reverse1, name2)

    reverse_token_score_1 = fuzz.token_sort_ratio(name1, reverse2)
    reverse_token_score_2 = fuzz.token_sort_ratio(reverse1, name2)

    # First/last representation
    fnln_score = fuzz.token_sort_ratio(fnln1, fnln2)

    lnfn_score = fuzz.token_sort_ratio(lnfn1, lnfn2)

    # Component-level
    component_score = component_similarity(name1, name2)

    # Position-aware
    positional_score = position_aware_similarity(name1, name2)

    # Strong reverse-order score
    reverse_score = max(
        reverse_score_1,
        reverse_score_2,
        reverse_token_score_1,
        reverse_token_score_2,
    )

    # --------------------------------------------------------
    # Main weighted scores
    # --------------------------------------------------------

    # Normal same-order comparison
    normal_weighted = (
        ratio_score * 0.35
        + token_sort_score * 0.25
        + component_score * 0.20
        + positional_score * 0.20
    )

    # Reverse-order comparison
    reverse_weighted = (
        reverse_score * 0.40
        + component_score * 0.25
        + token_sort_score * 0.15
        + positional_score * 0.20
    )

    # Alternative token score
    token_weighted = (
        token_set_score * 0.35
        + token_sort_score * 0.25
        + component_score * 0.20
        + positional_score * 0.20
    )

    final_score = max(
        normal_weighted,
        reverse_weighted,
        token_weighted,
    )

    # --------------------------------------------------------
    # Safety rules
    # --------------------------------------------------------

    # If names have multiple components and only one component
    # matches strongly, don't allow an excessive score.
    if len(get_tokens(name1)) >= 2 and len(get_tokens(name2)) >= 2:

        if positional_score < 55 and component_score < 70:
            final_score = min(final_score, 69.0)

    # Different last names should matter significantly.
    tokens1 = get_tokens(name1)
    tokens2 = get_tokens(name2)

    if len(tokens1) >= 2 and len(tokens2) >= 2:

        last_score = fuzz.ratio(tokens1[-1], tokens2[-1])

        first_score = fuzz.ratio(tokens1[0], tokens2[0])

        # If first and last names are both substantially different,
        # prevent an accidental high match.
        if first_score < 60 and last_score < 60:
            final_score = min(final_score, 65.0)

    return round(max(0.0, min(100.0, final_score)), 2)


# ============================================================
# DATA STRUCTURE
# ============================================================

class NameRecord:
    def __init__(self, row_number, document_id, full_name):
        self.row_number = row_number
        self.document_id = document_id
        self.full_name = full_name

        self.normalized_name = normalize_name(full_name)

        self.reverse_fn_ln = first_last_name(self.normalized_name)
        self.reverse_ln_fn = last_first_name(self.normalized_name)

        self.similarity_name = self.full_name
        self.similarity_score = 100.0

        self.group_number = None


# ============================================================
# MATCHING ENGINE
# ============================================================

class NameMatcher:

    def __init__(self, records, threshold, progress_callback=None):
        self.records = records
        self.threshold = threshold
        self.progress_callback = progress_callback

        self.name_to_indices = defaultdict(list)

        self.unique_names = []

        self.groups = {}

        self.match_count = 0

    def build_indexes(self):

        for index, record in enumerate(self.records):

            name = record.normalized_name

            if name:
                self.name_to_indices[name].append(index)

        self.unique_names = list(self.name_to_indices.keys())

    def report_progress(
        self,
        processed,
        total,
        matches,
        groups
    ):
        if self.progress_callback:
            self.progress_callback(
                processed,
                total,
                matches,
                groups
            )

    def find_best_match(self, record_index):

        record = self.records[record_index]

        if not record.normalized_name:

            return (
                record.full_name,
                0.0
            )

        current_name = record.normalized_name

        # ----------------------------------------------------
        # Exact match
        # ----------------------------------------------------

        if current_name in self.name_to_indices:

            exact_indices = self.name_to_indices[current_name]

            if len(exact_indices) > 1:

                # Exact duplicate exists.
                # Similarity remains 100%.
                return (
                    self.records[exact_indices[0]].full_name,
                    100.0
                )

        # ----------------------------------------------------
        # RapidFuzz candidate generation
        # ----------------------------------------------------

        candidate_results = process.extract(
            current_name,
            self.unique_names,
            scorer=fuzz.token_sort_ratio,
            limit=30
        )

        best_score = 0.0
        best_name = record.full_name

        for candidate_name, _, _ in candidate_results:

            if not candidate_name:
                continue

            if candidate_name == current_name:
                continue

            score = calculate_similarity(
                current_name,
                candidate_name
            )

            if score > best_score:

                best_score = score

                candidate_indices = self.name_to_indices[
                    candidate_name
                ]

                if candidate_indices:

                    best_name = self.records[
                        candidate_indices[0]
                    ].full_name

        # ----------------------------------------------------
        # If no qualifying match exists
        # ----------------------------------------------------

        if best_score <= self.threshold:

            return (
                record.full_name,
                100.0
            )

        return (
            best_name,
            round(best_score, 2)
        )

    def process_matches(self):

        self.build_indexes()

        total = len(self.records)

        for i in range(total):

            record = self.records[i]

            similarity_name, score = self.find_best_match(i)

            record.similarity_name = similarity_name
            record.similarity_score = score

            if (
                similarity_name != record.full_name
                and score > self.threshold
            ):
                self.match_count += 1

            self.report_progress(
                i + 1,
                total,
                self.match_count,
                0
            )

    # ========================================================
    # GROUP CREATION
    # ========================================================

    def create_groups(self):

        total = len(self.records)

        # Union-Find structure
        parent = list(range(total))

        rank = [0] * total

        def find(x):

            while parent[x] != x:
                parent[x] = parent[parent[x]]
                x = parent[x]

            return x

        def union(a, b):

            root_a = find(a)
            root_b = find(b)

            if root_a == root_b:
                return

            if rank[root_a] < rank[root_b]:
                parent[root_a] = root_b

            elif rank[root_a] > rank[root_b]:
                parent[root_b] = root_a

            else:
                parent[root_b] = root_a
                rank[root_a] += 1

        # ----------------------------------------------------
        # Exact duplicate grouping
        # ----------------------------------------------------

        for indices in self.name_to_indices.values():

            if len(indices) <= 1:
                continue

            first = indices[0]

            for other in indices[1:]:
                union(first, other)

        # ----------------------------------------------------
        # Generate candidates
        # ----------------------------------------------------

        unique_names = self.unique_names

        for index, record in enumerate(self.records):

            if not record.normalized_name:
                continue

            candidates = process.extract(
                record.normalized_name,
                unique_names,
                scorer=fuzz.token_sort_ratio,
                limit=30
            )

            for candidate_name, _, _ in candidates:

                if candidate_name == record.normalized_name:
                    continue

                score = calculate_similarity(
                    record.normalized_name,
                    candidate_name
                )

                if score > self.threshold:

                    candidate_indices = self.name_to_indices[
                        candidate_name
                    ]

                    for candidate_index in candidate_indices:

                        union(
                            index,
                            candidate_index
                        )

        # ----------------------------------------------------
        # Build groups
        # ----------------------------------------------------

        root_to_group = {}

        group_counter = 1

        for index in range(total):

            root = find(index)

            if root not in root_to_group:

                root_to_group[root] = group_counter

                group_counter += 1

            self.records[index].group_number = (
                root_to_group[root]
            )

        groups_created = group_counter - 1

        return groups_created


# ============================================================
# EXCEL PROCESSING
# ============================================================

def find_header_column(ws, required_header):

    for cell in ws[1]:

        if cell.value is None:
            continue

        header = str(cell.value).strip().lower()

        if header == required_header.lower():
            return cell.column

    return None


def read_input_excel(input_file, progress_callback=None):

    wb = load_workbook(
        input_file,
        read_only=True,
        data_only=False
    )

    if "Sheet1" in wb.sheetnames:
        ws = wb["Sheet1"]
    else:
        ws = wb[wb.sheetnames[0]]

    document_id_col = find_header_column(
        ws,
        "Document ID"
    )

    full_name_col = find_header_column(
        ws,
        "Full Name"
    )

    if document_id_col is None:

        wb.close()

        raise ValueError(
            "Column 'Document ID' was not found in Sheet1."
        )

    if full_name_col is None:

        wb.close()

        raise ValueError(
            "Column 'Full Name' was not found in Sheet1."
        )

    records = []

    total_rows = ws.max_row - 1

    if total_rows < 0:
        total_rows = 0

    for row_number, row in enumerate(
        ws.iter_rows(
            min_row=2,
            values_only=True
        ),
        start=2
    ):

        document_id = row[document_id_col - 1]

        full_name = row[full_name_col - 1]

        if document_id is None:
            document_id = ""

        if full_name is None:
            full_name = ""

        document_id = str(document_id).strip()
        full_name = str(full_name).strip()

        # Keep every row.
        record = NameRecord(
            row_number,
            document_id,
            full_name
        )

        records.append(record)

        if progress_callback:
            progress_callback(
                row_number - 1,
                total_rows,
                0,
                0
            )

    wb.close()

    return records


def write_output_excel(
    input_file,
    output_file,
    records,
    progress_callback=None
):

    # Load original workbook normally so Sheet1 remains intact.
    wb = load_workbook(input_file)

    # Delete existing Sheet2 if present.
    if OUTPUT_SHEET_NAME in wb.sheetnames:

        del wb[OUTPUT_SHEET_NAME]

    ws = wb.create_sheet(OUTPUT_SHEET_NAME)

    # --------------------------------------------------------
    # Headers
    # --------------------------------------------------------

    for col, header in enumerate(
        OUTPUT_HEADERS,
        start=1
    ):

        cell = ws.cell(
            row=1,
            column=col,
            value=header
        )

        cell.font = cell.font.copy(
            bold=True
        )

    # --------------------------------------------------------
    # Data
    # --------------------------------------------------------

    total = len(records)

    for row_index, record in enumerate(
        records,
        start=2
    ):

        # Unique ID
        unique_id = (
            f"{record.document_id}_{record.full_name}"
        )

        # If Full Name is empty, still preserve Document ID.
        if not record.full_name:

            unique_id = (
                f"{record.document_id}_"
            )

        values = [
            record.document_id,
            record.full_name,
            unique_id,
            record.reverse_fn_ln,
            record.reverse_ln_fn,
            record.similarity_name,
            record.similarity_score,
            f"G{record.group_number}",
        ]

        for col, value in enumerate(
            values,
            start=1
        ):

            ws.cell(
                row=row_index,
                column=col,
                value=value
            )

        if progress_callback:

            progress_callback(
                row_index - 1,
                total,
                None,
                None
            )

    # --------------------------------------------------------
    # Formatting
    # --------------------------------------------------------

    widths = {
        1: 18,
        2: 30,
        3: 45,
        4: 30,
        5: 30,
        6: 30,
        7: 20,
        8: 20,
    }

    for column_number, width in widths.items():

        ws.column_dimensions[
            ws.cell(
                row=1,
                column=column_number
            ).column_letter
        ].width = width

    ws.freeze_panes = "A2"

    ws.auto_filter.ref = (
        f"A1:H{max(1, len(records) + 1)}"
    )

    # Percentage formatting
    for row in range(2, len(records) + 2):

        ws.cell(
            row=row,
            column=7
        ).number_format = '0.00"%"'

    # Save
    wb.save(output_file)

    wb.close()


# ============================================================
# TKINTER GUI
# ============================================================

class NameFuzzyMatchApp:

    def __init__(self, root):

        self.root = root

        self.root.title(APP_TITLE)

        self.root.geometry("720x570")

        self.root.minsize(680, 520)

        self.root.resizable(True, True)

        self.processing = False

        self.setup_style()

        self.build_gui()

    # ========================================================
    # STYLE
    # ========================================================

    def setup_style(self):

        style = ttk.Style()

        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "Title.TLabel",
            font=("Segoe UI", 22, "bold")
        )

        style.configure(
            "Subtitle.TLabel",
            font=("Segoe UI", 10)
        )

        style.configure(
            "Section.TLabel",
            font=("Segoe UI", 11, "bold")
        )

        style.configure(
            "Start.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=10
        )

    # ========================================================
    # GUI
    # ========================================================

    def build_gui(self):

        main = ttk.Frame(
            self.root,
            padding=25
        )

        main.pack(
            fill="both",
            expand=True
        )

        # ----------------------------------------------------
        # Title
        # ----------------------------------------------------

        title = ttk.Label(
            main,
            text="Name Fuzzy Match",
            style="Title.TLabel"
        )

        title.pack(
            pady=(0, 5)
        )

        subtitle = ttk.Label(
            main,
            text=(
                "Match similar names, calculate similarity "
                "percentage and create groups."
            ),
            style="Subtitle.TLabel"
        )

        subtitle.pack(
            pady=(0, 25)
        )

        # ----------------------------------------------------
        # Input file
        # ----------------------------------------------------

        input_section = ttk.LabelFrame(
            main,
            text="Input Excel File",
            padding=15
        )

        input_section.pack(
            fill="x",
            pady=8
        )

        self.input_var = tk.StringVar()

        input_entry = ttk.Entry(
            input_section,
            textvariable=self.input_var
        )

        input_entry.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 10)
        )

        self.input_button = ttk.Button(
            input_section,
            text="Browse File",
            command=self.browse_input
        )

        self.input_button.pack(
            side="right"
        )

        # ----------------------------------------------------
        # Output file
        # ----------------------------------------------------

        output_section = ttk.LabelFrame(
            main,
            text="Output Excel File",
            padding=15
        )

        output_section.pack(
            fill="x",
            pady=8
        )

        self.output_var = tk.StringVar()

        output_entry = ttk.Entry(
            output_section,
            textvariable=self.output_var
        )

        output_entry.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 10)
        )

        self.output_button = ttk.Button(
            output_section,
            text="Save File",
            command=self.browse_output
        )

        self.output_button.pack(
            side="right"
        )

        # ----------------------------------------------------
        # Matching percentage
        # ----------------------------------------------------

        setting_section = ttk.LabelFrame(
            main,
            text="Matching Setting",
            padding=15
        )

        setting_section.pack(
            fill="x",
            pady=8
        )

        ttk.Label(
            setting_section,
            text="Matching Percentage:"
        ).pack(
            side="left"
        )

        self.threshold_var = tk.StringVar(
            value="80"
        )

        self.threshold_entry = ttk.Entry(
            setting_section,
            textvariable=self.threshold_var,
            width=10
        )

        self.threshold_entry.pack(
            side="left",
            padx=10
        )

        ttk.Label(
            setting_section,
            text="%"
        ).pack(
            side="left"
        )

        ttk.Label(
            setting_section,
            text="  Match only when similarity is greater than this value."
        ).pack(
            side="left"
        )

        # ----------------------------------------------------
        # Progress section
        # ----------------------------------------------------

        progress_section = ttk.LabelFrame(
            main,
            text="Processing Status",
            padding=15
        )

        progress_section.pack(
            fill="x",
            pady=15
        )

        self.status_var = tk.StringVar(
            value="Ready"
        )

        self.status_label = ttk.Label(
            progress_section,
            textvariable=self.status_var
        )

        self.status_label.pack(
            anchor="w"
        )

        self.progress = ttk.Progressbar(
            progress_section,
            orient="horizontal",
            mode="determinate"
        )

        self.progress.pack(
            fill="x",
            pady=10
        )

        self.progress_text_var = tk.StringVar(
            value="0%"
        )

        ttk.Label(
            progress_section,
            textvariable=self.progress_text_var
        ).pack(
            anchor="center"
        )

        self.records_var = tk.StringVar(
            value="Records Processed: 0 / 0"
        )

        self.matches_var = tk.StringVar(
            value="Matches Found: 0"
        )

        self.groups_var = tk.StringVar(
            value="Groups Created: 0"
        )

        ttk.Label(
            progress_section,
            textvariable=self.records_var
        ).pack(
            anchor="w"
        )

        ttk.Label(
            progress_section,
            textvariable=self.matches_var
        ).pack(
            anchor="w"
        )

        ttk.Label(
            progress_section,
            textvariable=self.groups_var
        ).pack(
            anchor="w"
        )

        # ----------------------------------------------------
        # Start button
        # ----------------------------------------------------

        self.start_button = ttk.Button(
            main,
            text="START",
            style="Start.TButton",
            command=self.start_processing
        )

        self.start_button.pack(
            pady=15
        )

        # ----------------------------------------------------
        # Footer
        # ----------------------------------------------------

        ttk.Label(
            main,
            text="Name Fuzzy Match • Excel Processing Tool",
            style="Subtitle.TLabel"
        ).pack(
            side="bottom",
            pady=(10, 0)
        )

    # ========================================================
    # BROWSE
    # ========================================================

    def browse_input(self):

        if self.processing:
            return

        filename = filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=[
                (
                    "Excel Files",
                    "*.xlsx *.xlsm"
                ),
                (
                    "All Files",
                    "*.*"
                ),
            ]
        )

        if filename:

            self.input_var.set(filename)

            # Automatically suggest output filename.
            directory = os.path.dirname(filename)

            base = os.path.splitext(
                os.path.basename(filename)
            )[0]

            output = os.path.join(
                directory,
                f"{base}_Name_Fuzzy_Match.xlsx"
            )

            self.output_var.set(output)

    def browse_output(self):

        if self.processing:
            return

        filename = filedialog.asksaveasfilename(
            title="Save Output Excel File",
            defaultextension=".xlsx",
            filetypes=[
                (
                    "Excel Files",
                    "*.xlsx"
                ),
                (
                    "All Files",
                    "*.*"
                ),
            ]
        )

        if filename:

            self.output_var.set(filename)

    # ========================================================
    # VALIDATION
    # ========================================================

    def validate_inputs(self):

        input_file = self.input_var.get().strip()

        output_file = self.output_var.get().strip()

        threshold_text = self.threshold_var.get().strip()

        if not input_file:

            messagebox.showerror(
                "Input Required",
                "Please select an input Excel file."
            )

            return None

        if not os.path.isfile(input_file):

            messagebox.showerror(
                "Invalid Input",
                "The selected input file does not exist."
            )

            return None

        if not output_file:

            messagebox.showerror(
                "Output Required",
                "Please select an output Excel file."
            )

            return None

        try:

            threshold = float(
                threshold_text
            )

        except ValueError:

            messagebox.showerror(
                "Invalid Percentage",
                "Please enter a valid number between 0 and 100."
            )

            return None

        if threshold < 0 or threshold > 100:

            messagebox.showerror(
                "Invalid Percentage",
                "Matching Percentage must be between 0 and 100."
            )

            return None

        if os.path.abspath(input_file) == os.path.abspath(
            output_file
        ):

            messagebox.showerror(
                "Invalid Output",
                "Input and output files must be different."
            )

            return None

        return (
            input_file,
            output_file,
            threshold
        )

    # ========================================================
    # START
    # ========================================================

    def start_processing(self):

        if self.processing:
            return

        values = self.validate_inputs()

        if values is None:
            return

        input_file, output_file, threshold = values

        self.processing = True

        self.start_button.config(
            state="disabled"
        )

        self.input_button.config(
            state="disabled"
        )

        self.output_button.config(
            state="disabled"
        )

        self.threshold_entry.config(
            state="disabled"
        )

        self.progress["value"] = 0

        self.status_var.set(
            "Starting processing..."
        )

        self.progress_text_var.set(
            "0%"
        )

        self.records_var.set(
            "Records Processed: 0 / 0"
        )

        self.matches_var.set(
            "Matches Found: 0"
        )

        self.groups_var.set(
            "Groups Created: 0"
        )

        thread = threading.Thread(
            target=self.run_processing,
            args=(
                input_file,
                output_file,
                threshold
            ),
            daemon=True
        )

        thread.start()

    # ========================================================
    # PROCESSING
    # ========================================================

    def run_processing(
        self,
        input_file,
        output_file,
        threshold
    ):

        try:

            # ------------------------------------------------
            # Step 1
            # ------------------------------------------------

            self.root.after(
                0,
                lambda: self.status_var.set(
                    "Reading Sheet1..."
                )
            )

            records = read_input_excel(
                input_file,
                progress_callback=self.update_read_progress
            )

            total = len(records)

            if total == 0:

                raise ValueError(
                    "No data rows were found in Sheet1."
                )

            # ------------------------------------------------
            # Step 2
            # ------------------------------------------------

            self.root.after(
                0,
                lambda: self.status_var.set(
                    "Processing Names..."
                )
            )

            matcher = NameMatcher(
                records,
                threshold,
                progress_callback=self.update_match_progress
            )

            matcher.process_matches()

            # ------------------------------------------------
            # Step 3
            # ------------------------------------------------

            self.root.after(
                0,
                lambda: self.status_var.set(
                    "Creating Groups..."
                )
            )

            groups_created = matcher.create_groups()

            self.root.after(
                0,
                lambda: self.groups_var.set(
                    f"Groups Created: {groups_created}"
                )
            )

            # ------------------------------------------------
            # Step 4
            # ------------------------------------------------

            self.root.after(
                0,
                lambda: self.status_var.set(
                    "Writing Sheet2..."
                )
            )

            write_output_excel(
                input_file,
                output_file,
                records,
                progress_callback=self.update_write_progress
            )

            # ------------------------------------------------
            # Finished
            # ------------------------------------------------

            self.root.after(
                0,
                lambda: self.processing_completed(
                    output_file,
                    total,
                    matcher.match_count,
                    groups_created
                )
            )

        except Exception as error:

            error_text = (
                f"{str(error)}\n\n"
                f"Technical Details:\n"
                f"{traceback.format_exc()}"
            )

            self.root.after(
                0,
                lambda: self.processing_failed(
                    error_text
                )
            )

    # ========================================================
    # PROGRESS UPDATES
    # ========================================================

    def update_read_progress(
        self,
        processed,
        total,
        matches,
        groups
    ):

        if total <= 0:
            return

        percentage = (
            processed / total
        ) * 20

        self.root.after(
            0,
            lambda: self.update_progress_ui(
                percentage,
                processed,
                total,
                matches if matches is not None else 0,
                groups if groups is not None else 0
            )
        )

    def update_match_progress(
        self,
        processed,
        total,
        matches,
        groups
    ):

        if total <= 0:
            return

        # Matching = 60% of total progress
        percentage = 20 + (
            processed / total
        ) * 60

        self.root.after(
            0,
            lambda: self.update_progress_ui(
                percentage,
                processed,
                total,
                matches if matches is not None else 0,
                groups if groups is not None else 0
            )
        )

    def update_write_progress(
        self,
        processed,
        total,
        matches,
        groups
    ):

        if total <= 0:
            return

        percentage = 80 + (
            processed / total
        ) * 20

        self.root.after(
            0,
            lambda: self.update_progress_ui(
                percentage,
                processed,
                total,
                0,
                0
            )
        )

    def update_progress_ui(
        self,
        percentage,
        processed,
        total,
        matches,
        groups
    ):

        percentage = max(
            0,
            min(
                100,
                percentage
            )
        )

        self.progress["value"] = percentage

        self.progress_text_var.set(
            f"{percentage:.0f}%"
        )

        self.records_var.set(
            f"Records Processed: {processed:,} / {total:,}"
        )

        if matches is not None:

            self.matches_var.set(
                f"Matches Found: {matches:,}"
            )

        if groups is not None and groups > 0:

            self.groups_var.set(
                f"Groups Created: {groups:,}"
            )

    # ========================================================
    # COMPLETED
    # ========================================================

    def processing_completed(
        self,
        output_file,
        total,
        matches,
        groups
    ):

        self.processing = False

        self.progress["value"] = 100

        self.progress_text_var.set(
            "100%"
        )

        self.status_var.set(
            "Processing Completed Successfully!"
        )

        self.records_var.set(
            f"Records Processed: {total:,} / {total:,}"
        )

        self.matches_var.set(
            f"Matches Found: {matches:,}"
        )

        self.groups_var.set(
            f"Groups Created: {groups:,}"
        )

        self.enable_controls()

        messagebox.showinfo(
            "Processing Completed",
            (
                "Processing Completed Successfully!\n\n"
                f"Records Processed: {total:,}\n"
                f"Matches Found: {matches:,}\n"
                f"Groups Created: {groups:,}\n\n"
                f"Output File:\n{output_file}"
            )
        )

    # ========================================================
    # FAILED
    # ========================================================

    def processing_failed(
        self,
        error_text
    ):

        self.processing = False

        self.status_var.set(
            "Processing Failed"
        )

        self.enable_controls()

        # Show only the main error to the user.
        main_error = error_text.split(
            "Technical Details:"
        )[0].strip()

        messagebox.showerror(
            "Processing Error",
            (
                "The program could not complete the "
                "processing.\n\n"
                f"{main_error}"
            )
        )

        print(
            "\n========== ERROR ==========\n"
            f"{error_text}\n"
            "===========================\n"
        )

    # ========================================================
    # ENABLE CONTROLS
    # ========================================================

    def enable_controls(self):

        self.start_button.config(
            state="normal"
        )

        self.input_button.config(
            state="normal"
        )

        self.output_button.config(
            state="normal"
        )

        self.threshold_entry.config(
            state="normal"
        )


# ============================================================
# MAIN
# ============================================================

def main():

    root = tk.Tk()

    app = NameFuzzyMatchApp(
        root
    )

    root.mainloop()


if __name__ == "__main__":
    main()