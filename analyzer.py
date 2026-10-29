"""
Analyzer for college mark lists — supports PDF / Excel / CSV
Handles consolidated format like: S.No | RegNo | Subj1...SubjN | Total
"""
import os
import re
from datetime import datetime
from xml.sax.saxutils import escape
import pdfplumber
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                Paragraph, Spacer, PageBreak, Flowable)

CLASS_SUBJECT_FACULTY = {
    "II/3/A": {
        "23BS305": ("Discrete Mathematics", "MYTHILI D"),
        "23MC003": ("Indian Constitution", "BABITHA K"),
        "23TPS03": ("Quantitative Aptitude and Logical Reasoning - I",
                    "TAMILSELVAN P"),
        "23CS301": ("Data Structures", "SARAVANAN G"),
        "23AD301": ("Artificial Intelligence", "SATHISH KUMAR P"),
        "23AD303": ("Digital Computer Organization and Architecture",
                    "LINETT SOPHIA D"),
    },
    "II/3/C": {
        "23BS305": ("Discrete Mathematics", "PADMA PRIYA J"),
        "23MC003": ("Indian Constitution", "RAMESH BABU"),
        "23TPS03": ("Quantitative Aptitude and Logical Reasoning - I",
                    "TAMILSELVAN P"),
        "23CS301": ("Data Structures", "SARAVANAN G"),
        "23AD301": ("Artificial Intelligence", "SATHISH KUMAR P"),
        "23AD303": ("Digital Computer Organization and Architecture",
                    "LINETT SOPHIA D"),
    },
    "III/5/A": {
        "23AD501": ("Machine Learning", "MAHALINGAM M"),
        "23AD502": ("Data Exploration and Visualization", "RAMYASRI A"),
        "23ADX08": ("Java Programming", "PUSHPARAJ P"),
        "23TPS05": ("Quantitative Aptitude and Logical Reasoning - III",
                    "MEKALA C"),
        "23ADX11": ("R Programming", "SENTHIL KUMAR M"),
    },
    "III/5/B": {
        "23AD501": ("Machine Learning", "MOHANAPRIYA B"),
        "23AD502": ("Data Exploration and Visualization", "VANITHA T"),
        "23ADX08": ("Java Programming", "PUSHPARAJ P"),
        "23TPS05": ("Quantitative Aptitude and Logical Reasoning - III",
                    "MEKALA C"),
        "23ADX11": ("R Programming", "SENTHIL KUMAR M"),
    },
    "III/5/C": {
        "23AD501": ("Machine Learning", "MOHANAPRIYA B"),
        "23AD502": ("Data Exploration and Visualization", "VANITHA T"),
        "23ADX08": ("Java Programming", "PUSHPARAJ P"),
        "23TPS05": ("Quantitative Aptitude and Logical Reasoning - III",
                    "MEKALA C"),
        "23ADX11": ("R Programming", "SENTHIL KUMAR M"),
    },
    "IV/7/A": {
        "23AD701": ("Cloud Computing", "VINODHINI A N"),
        "23AD702": ("Natural Language Processing", "VINODHINI S"),
        "23AD703": ("Predictive Analytics", "ARCHANADEVI S"),
        "23ADX37": ("Edge Computing", "SENTHILKUMAR C"),
        "23ADX48": ("Management Information System", "NANDHINI V"),
        "23ECY01": ("Basics of VLSI Design", "VIJI D"),
    },
    "IV/7/B": {
        "23AD701": ("Cloud Computing", "VINODHINI A N"),
        "23AD702": ("Natural Language Processing", "VINODHINI S"),
        "23AD703": ("Predictive Analytics", "ARCHANADEVI S"),
        "23ADX37": ("Edge Computing", "SENTHILKUMAR C"),
        "23ADX48": ("Management Information System", "NANDHINI V"),
        "23ECY01": ("Basics of VLSI Design", "VIJI D"),
    },
}


def _normalize_class_key(value):
    return re.sub(r"\s+", "", str(value or "").upper())


def _extract_section_letter(value):
    text = str(value or "").upper()
    match = re.search(r"(?:/|SECTION\s*|YEAR\s+)([A-Z])\b", text)
    return match.group(1) if match else ""


def _class_subject_mapping(class_key, class_info, raw_text, subjects):
    normalized_lookup = {
        _normalize_class_key(key): value
        for key, value in CLASS_SUBJECT_FACULTY.items()
    }
    mapping = normalized_lookup.get(_normalize_class_key(class_key), {})
    if mapping:
        return mapping

    normalized_metadata = _normalize_class_key(raw_text)
    for key, candidate in normalized_lookup.items():
        if key in normalized_metadata:
            return candidate

    subject_codes = {str(subject).strip().upper() for subject in subjects}
    ranked = []
    for key, candidate in CLASS_SUBJECT_FACULTY.items():
        mapped_codes = {code.strip().upper() for code in candidate}
        score = len(subject_codes & mapped_codes)
        if score:
            ranked.append((score, key, candidate))
    if not ranked:
        return {}

    ranked.sort(key=lambda item: (-item[0], item[1]))
    desired_section = (
        _extract_section_letter(class_key)
        or _extract_section_letter(class_info)
        or _extract_section_letter(raw_text)
    )
    if desired_section:
        for score, key, candidate in ranked:
            if score == ranked[0][0] and key.endswith(f"/{desired_section}"):
                return candidate
    return ranked[0][2]


class CircledMark(Flowable):
    """Display a failed mark with a circle around it in the PDF table."""

    def __init__(self, value, width=30 * mm, height=8 * mm):
        Flowable.__init__(self)
        self.value = value
        self.width = width
        self.height = height

    def wrap(self, available_width, available_height):
        return self.width, self.height

    def draw(self):
        self.canv.setFont("Helvetica", 7)
        self.canv.setLineWidth(0.8)
        self.canv.ellipse(
            self.width * 0.18, self.height * 0.08,
            self.width * 0.82, self.height * 0.92,
        )
        text = str(self.value)
        text_width = self.canv.stringWidth(text, "Helvetica", 7)
        self.canv.drawString((self.width - text_width) / 2, self.height * 0.36, text)


class ResultAnalyzer:
    def __init__(self, filepath, pass_mark=40, subject_names=None):
        self.filepath = filepath
        self.pass_mark = pass_mark
        self.df = None
        self.subjects = []
        self.failures_df = None
        self.roll_col = None
        self.name_col = None
        self.subject_names = subject_names or []
        self.pdf_optional_columns = []
        self.raw_text = ""
        self._numeric_df = None

    # ============================================================
    #  LOADER
    # ============================================================
    def load_data(self):
        ext = os.path.splitext(self.filepath)[1].lower()
        if ext == ".pdf":
            self._load_pdf()
        elif ext in (".xlsx", ".xls"):
            self.df = self._read_excel_with_header()
            self._post_load_excel_csv()
        elif ext == ".csv":
            self.df = pd.read_csv(self.filepath)
            self._post_load_excel_csv()
        else:
            raise ValueError("Unsupported format. Use .pdf, .xlsx, .xls, or .csv")

        if self.df is not None:
            self.df.columns = [str(c).strip() for c in self.df.columns]
        return self.df

    def _read_excel_with_header(self):
        """Read files with optional college/title rows above the mark headers."""
        preview = pd.read_excel(self.filepath, header=None, nrows=25)
        header_row = None
        for index, row in preview.iterrows():
            values = {
                str(value).strip().lower().replace("_", " ")
                for value in row.tolist()
                if pd.notna(value)
            }
            has_serial = bool(values & {"s.no", "s no", "sno"})
            has_register = bool(values & {"reg no", "regno", "roll no",
                                           "roll number"})
            if has_serial and has_register:
                header_row = index
                break
        return pd.read_excel(self.filepath, header=header_row or 0)

    def _post_load_excel_csv(self):
        """Setup columns for Excel/CSV input."""
        if len(self.df.columns) < 3:
            raise ValueError("File needs at least 3 columns: RollNo | Name | Subjects...")
        columns = list(self.df.columns)
        normalised = {str(c).strip().lower().replace("_", " "): c
                      for c in columns}

        self.roll_col = next(
            (normalised[key] for key in ("reg no", "regno", "roll no", "roll number")
             if key in normalised),
            columns[1] if str(columns[0]).strip().lower() in ("s.no", "s no", "sno")
            else columns[0],
        )
        serial_col = next(
            (normalised[key] for key in ("s.no", "s no", "sno") if key in normalised),
            None,
        )
        self.name_col = next(
            (normalised[key] for key in ("name", "student name") if key in normalised),
            None,
        )
        if self.name_col is None and self.roll_col == columns[0]:
            self.name_col = columns[1]
        excluded = {self.roll_col, serial_col, self.name_col}
        excluded.update(
            c for c in columns
            if str(c).strip().lower().replace("_", " ") in
            {"total", "total mark", "total marks", "mark", "marks in words"}
        )
        self.subjects = [c for c in columns if c not in excluded]

    # ============================================================
    #  PDF PARSER
    # ============================================================
    def _load_pdf(self):
        """Extract student data from a consolidated mark list PDF."""
        all_text = []
        with pdfplumber.open(self.filepath) as pdf:
            for page in pdf.pages:
                # Try structured tables first (won't hurt)
                try:
                    for tbl in (page.extract_tables() or []):
                        for row in tbl:
                            if row:
                                cleaned = " ".join(str(c) for c in row if c)
                                all_text.append(cleaned)
                except Exception:
                    pass
                # Also raw text
                txt = page.extract_text() or ""
                if txt:
                    all_text.append(txt)

        full_text = "\n".join(all_text)
        self.raw_text = full_text

        print("\n=== PDF TEXT SAMPLE (first 1200 chars) ===")
        print(full_text[:1200])
        print("=== END SAMPLE ===\n")

        self._extract_subject_headers(full_text)
        self._extract_subjects_from_summary(full_text)
        if self.pdf_optional_columns:
            optional_code = re.search(
                r"\((\d{2}[A-Z]{2,4}\d{2,3})\)", full_text, re.IGNORECASE
            )
            if optional_code:
                code = optional_code.group(1).upper()
                if code not in self.subject_names:
                    self.subject_names.append(code)
                self.pdf_optional_columns = []

        # Extract student rows — multiple strategies
        students = self._extract_student_rows(full_text)
        print(f"=== DETECTED {len(students)} STUDENT ROWS ===\n")

        if not students:
            raise ValueError(
                "No student rows detected in PDF. "
                "It may be a scanned image (not text-based), "
                "or format is very different. Check terminal debug output."
            )

        # Pad all students to same subject count
        max_subj = max(len(s["marks"]) for s in students)
        cols = ["S.No", "Reg No"] + [f"Subject{i+1}" for i in range(max_subj)] + ["Total"]

        rows = []
        for s in students:
            marks = list(s["marks"])
            while len(marks) < max_subj:
                marks.append(None)
            rows.append([s["sno"], s["regno"]] + marks + [s["total"]])

        self.df = pd.DataFrame(rows, columns=cols)

        print("=== DATAFRAME COLUMNS ===")
        print(list(self.df.columns))
        print()

        # Rename Subject1..N → actual names if detected
        if self.subject_names:
            rename_map = {}
            for i, name in enumerate(self.subject_names):
                old = f"Subject{i+1}"
                if old in self.df.columns:
                    rename_map[old] = name
            self.df.rename(columns=rename_map, inplace=True)

        self.roll_col = "Reg No"
        self.name_col = None
        self.subjects = [c for c in self.df.columns
                         if c not in ("S.No", "Reg No", "Total")]

        print(f"OK Roll column: {self.roll_col}")
        print(f"OK Subjects: {self.subjects}\n")

    def _extract_subject_headers(self, text):
        """Read subject codes from the consolidated table header when available."""
        header_re = re.compile(
            r"^\s*S\.?No\s+Reg\s+No\s+(.+?)\s+Total\s+Mark(?:\s+MARK)?\s*$",
            re.IGNORECASE | re.MULTILINE,
        )
        match = header_re.search(text)
        if not match:
            return
        headers = match.group(1).split()
        self.subject_names = [
            header for header in headers
            if header.lower() not in {"total", "mark", "marks", "oe"}
        ]
        self.pdf_optional_columns = [
            header for header in headers if header.lower() in {"oe"}
        ]

    def _extract_subjects_from_summary(self, text):
        """Extract subject names from a professor-wise summary section."""
        summary_idx = text.find("Professor wise Summary")
        if summary_idx == -1:
            summary_idx = text.rfind("Subject Code")
        if summary_idx == -1:
            return

        section = text[summary_idx: summary_idx + 2000]
        pattern = re.compile(
            r"(\d)(\d{2}[A-Z]{2,4}\d{3})"
            r"([A-Za-z][A-Za-z0-9\s\-&.,/]+?)"
            r"([A-Z][A-Z\s.]{2,})(\d+)"
        )
        subjects = []
        for match in pattern.finditer(section):
            name = match.group(3).strip()
            if name and len(name) < 80 and name not in subjects:
                subjects.append(name)

        if subjects:
            self.subject_names = subjects[:10]
        print(f"=== DETECTED SUBJECTS: {self.subject_names} ===\n")

    def _extract_student_rows(self, text):
        """Robust multi-strategy student row extraction."""
        students = []
        seen_regnos = set()

        # ---------- Strategy 1: Line-based ----------
        # Format: S.No RegNo mark1 mark2 ... markN Total [English words]
        line_re = re.compile(
            r"^\s*(\d{1,3})\s+(\d{8,20})\s+(.+?)\s*$", re.MULTILINE
        )
        for m in line_re.finditer(text):
            sno, regno = m.group(1), m.group(2)
            if regno in seen_regnos:
                continue
            tokens = [
                token for token in m.group(3).split()
                if not re.fullmatch(
                    r"\(\d{2}[A-Z]{2,4}\d{2,3}\)", token, re.IGNORECASE
                )
            ]
            expected_columns = len(self.subject_names) + len(
                self.pdf_optional_columns
            )
            if expected_columns and len(tokens) < expected_columns + 1:
                continue
            if expected_columns:
                column_values = tokens[:expected_columns]
                total = tokens[expected_columns]
                marks = [
                    value for column, value in zip(
                        self.subject_names + self.pdf_optional_columns,
                        column_values,
                    )
                    if column not in self.pdf_optional_columns
                ]
            else:
                total = next(
                    (value for value in tokens if value.isdigit()), ""
                )
                marks = tokens[:tokens.index(total)] if total else []
            cleaned = self._clean_marks(marks)
            if len(cleaned) >= 3 and total.isdigit():
                seen_regnos.add(regno)
                students.append({
                    "sno": sno,
                    "regno": regno,
                    "marks": cleaned,
                    "total": int(total),
                })

        # ---------- Strategy 2: RegNo-anchored fallback ----------
        if not students:
            print("⚠️  Strategy 1 failed — trying RegNo-anchor strategy...")
            for match in re.finditer(r"\b(\d{8,20})\b", text):
                regno = match.group(1)
                if regno in seen_regnos:
                    continue
                window = text[match.end(): match.end() + 250]
                tokens = re.findall(r"\b(AB|ab|\d{1,3})\b", window)
                if len(tokens) < 4:
                    continue
                # Last numeric = total
                numeric_tokens = [t for t in tokens if t.isdigit()]
                if len(numeric_tokens) < 4:
                    continue
                # Guess: total is at position ~7 (6 subjects + total)
                total = tokens[-1] if tokens[-1].isdigit() else tokens[-2]
                # Marks = first 6 non-total
                marks_tokens = tokens[:-1][:8]
                cleaned = self._clean_marks(marks_tokens)
                if len(cleaned) >= 3:
                    seen_regnos.add(regno)
                    students.append({
                        "sno": str(len(students) + 1),
                        "regno": regno,
                        "marks": cleaned,
                        "total": int(total) if str(total).isdigit() else 0,
                    })

        return students

    def _clean_marks(self, marks):
        cleaned = []
        for v in marks:
            if str(v).upper() == "AB":
                cleaned.append("AB")
            else:
                try:
                    cleaned.append(int(v))
                except (ValueError, TypeError):
                    cleaned.append(v)
        return cleaned

    # ============================================================
    #  ANALYSIS
    # ============================================================
    def _to_numeric(self):
        df = self.df.copy()
        for sub in self.subjects:
            df[sub] = pd.to_numeric(df[sub], errors="coerce")
        return df

    def analyze(self):
        if self.df is None:
            self.load_data()

        numeric = self._to_numeric()
        fail_matrix = pd.DataFrame(index=self.df.index)
        for subject in self.subjects:
            marks = numeric[subject]
            raw_marks = self.df[subject].astype(str).str.strip().str.upper()
            absent = raw_marks.isin({"AB", "ABSENT", "A"})
            fail_matrix[subject] = marks.isna() | (marks < self.pass_mark) | absent
        fail_matrix.insert(0, self.roll_col, self.df[self.roll_col])
        if "S.No" in self.df.columns:
            fail_matrix.insert(0, "S.No", self.df["S.No"])

        fail_matrix["Total_Failures"] = fail_matrix[self.subjects].sum(axis=1)
        fail_matrix["Failed_Subjects"] = fail_matrix[self.subjects].apply(
            lambda row: ", ".join([s for s in self.subjects if row[s]]), axis=1
        )
        fail_matrix["Status"] = fail_matrix["Total_Failures"].apply(
            lambda x: "FAIL" if x > 0 else "PASS"
        )

        self.failures_df = fail_matrix
        self._numeric_df = numeric
        return fail_matrix

    def get_statistics(self):
        if self.failures_df is None:
            self.analyze()

        total = len(self.df)
        failed = int((self.failures_df["Status"] == "FAIL").sum())
        passed = total - failed

        subject_fail_counts = {
            sub: int(self.failures_df[sub].sum()) for sub in self.subjects
        }
        subject_pass_counts = {
            sub: total - subject_fail_counts[sub] for sub in self.subjects
        }
        subject_avg = {
            sub: round(float(self._numeric_df[sub].mean(skipna=True)), 2)
            if not self._numeric_df[sub].dropna().empty else 0
            for sub in self.subjects
        }
        subject_max = {
            sub: int(self._numeric_df[sub].max(skipna=True))
            if not self._numeric_df[sub].dropna().empty else 0
            for sub in self.subjects
        }

        return {
            "total_students": total,
            "passed_students": passed,
            "failed_students": failed,
            "pass_percentage": round((passed / total) * 100, 2) if total else 0,
            "subject_fail_counts": subject_fail_counts,
            "subject_pass_counts": subject_pass_counts,
            "subject_avg": subject_avg,
            "subject_max": subject_max,
        }

    def get_student_wise_failures(self):
        if self.failures_df is None:
            self.analyze()
        cols = [c for c in ["S.No", self.roll_col, "Failed_Subjects",
                            "Total_Failures", "Status"]
                if c in self.failures_df.columns]
        df = self.failures_df[cols].copy()
        df.rename(columns={self.roll_col: "Reg No"}, inplace=True)
        total_subjects = len(self.subjects)
        df["Pass_Percentage"] = (
            (total_subjects - df["Total_Failures"]) / total_subjects * 100
            if total_subjects else 0
        ).round(2)
        return df

    def get_subject_wise_failures(self):
        if self.failures_df is None:
            self.analyze()

        subject_reports = {}
        for sub in self.subjects:
            failed_mask = self.failures_df[sub] == True
            failed_students = self.df[failed_mask]
            cols = [c for c in ["S.No", self.roll_col] if c in failed_students.columns]
            df = failed_students[cols].rename(columns={self.roll_col: "Reg No"})
            df["Marks_Obtained"] = failed_students[sub].values
            subject_reports[sub] = df.to_dict(orient="records")
        return subject_reports

    def generate_round_lists(self):
        if self.failures_df is None:
            self.analyze()

        round_lists = {}
        for sub in self.subjects:
            failed_mask = self.failures_df[sub] == True
            failed_students = self.df[failed_mask]
            if not failed_students.empty:
                rows = []
                for _, row in failed_students.iterrows():
                    rows.append({
                        "S.No": row.get("S.No", ""),
                        "Reg No": row[self.roll_col],
                        "Marks_Obtained": row[sub],
                    })
                round_lists[sub] = rows
        return round_lists

    def _number_in_words(self, value):
        """Return the integer total in the same words-only style as the PDF."""
        try:
            number = int(value)
        except (TypeError, ValueError):
            return ""
        ones = ["ZERO", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX",
                "SEVEN", "EIGHT", "NINE", "TEN", "ELEVEN", "TWELVE",
                "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN",
                "SEVENTEEN", "EIGHTEEN", "NINETEEN"]
        tens = ["", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY",
                "SEVENTY", "EIGHTY", "NINETY"]
        if number < 20:
            return ones[number]
        if number < 100:
            return tens[number // 10] + (f" {ones[number % 10]}"
                                         if number % 10 else "")
        if number < 1000:
            remainder = number % 100
            return f"{ones[number // 100]} HUNDRED" + (
                f" {self._number_in_words(remainder)}" if remainder else "")
        return str(number)

    def consolidated_mark_list(self):
        """Build the dynamic, line-wise mark list used by Excel and PDF exports."""
        if self.df is None:
            self.load_data()
        serial_col = "S.No" if "S.No" in self.df.columns else None
        rows = []
        numeric = self._to_numeric()
        for index, row in self.df.iterrows():
            marks = [row.get(subject, "") for subject in self.subjects]
            total = sum(
                float(numeric.iloc[index][subject])
                for subject in self.subjects
                if pd.notna(numeric.iloc[index][subject])
            )
            total = int(total) if isinstance(total, float) and total.is_integer() else total
            rows.append({
                "S.No": row.get(serial_col, index + 1) if serial_col else index + 1,
                "Reg No": row.get(self.roll_col, ""),
                **{subject: mark for subject, mark in zip(self.subjects, marks)},
                "Total Mark": total,
                "MARK": self._number_in_words(total),
                "Pass Percentage": round(
                    sum(
                        pd.notna(numeric.iloc[index][subject])
                        and str(row.get(subject, "")).strip().upper()
                        not in {"AB", "ABSENT", "A"}
                        and numeric.iloc[index][subject] >= self.pass_mark
                        for subject in self.subjects
                    ) / len(self.subjects) * 100, 2
                ) if self.subjects else 0,
            })
        return pd.DataFrame(rows)

    def export_student_correction_pdf(self, output_dir, college_name="",
                                      class_info=""):
        """Export a printable subject matrix for teacher mark corrections."""
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(
            output_dir, f"student_correction_report_{timestamp}.pdf"
        )

        doc = SimpleDocTemplate(
            out_path, pagesize=landscape(A4),
            leftMargin=8 * mm, rightMargin=8 * mm,
            topMargin=10 * mm, bottomMargin=10 * mm,
        )
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "CorrectionTitle", parent=styles["Title"],
            fontSize=14, leading=17, alignment=1, spaceAfter=2 * mm,
        )
        info_style = ParagraphStyle(
            "CorrectionInfo", parent=styles["Normal"],
            fontSize=9, leading=11, alignment=1,
        )

        if self.df is None:
            self.load_data()
        if self.failures_df is None:
            self.analyze()

        def correction_cell(value):
            normalized = str(value).strip().upper()
            if normalized in {"OD", "ON DUTY", "ON-DUTY", "AB", "ABSENT", "A"}:
                return ""
            numeric_value = pd.to_numeric(value, errors="coerce")
            if pd.notna(numeric_value) and numeric_value >= self.pass_mark:
                return "\N{EN DASH}"
            return ""

        story = [Paragraph(
            college_name or "Student Detailed Correction Report", title_style
        )]
        if college_name:
            story.append(Paragraph("Student Detailed Correction Report",
                                   info_style))
        if class_info:
            story.append(Paragraph(f"Class: {class_info}", info_style))
        story.append(Spacer(1, 4 * mm))

        header_style = ParagraphStyle(
            "CorrectionHeader", parent=styles["Normal"],
            fontSize=8, leading=10, alignment=1, textColor=colors.white,
        )
        cell_style = ParagraphStyle(
            "CorrectionCell", parent=styles["Normal"],
            fontSize=8, leading=10, alignment=1,
        )
        name_style = ParagraphStyle(
            "CorrectionName", parent=cell_style, alignment=0,
        )
        headers = [
            Paragraph("Roll No", header_style),
            Paragraph("Name of<br/>Student", header_style),
        ] + [
            Paragraph(str(subject).replace(" & ", " &<br/>"), header_style)
            for subject in self.subjects
        ]
        rows = [headers]
        for _, row in self.df.iterrows():
            student_name = row.get(self.name_col, "") if self.name_col else ""
            student_name = "" if pd.isna(student_name) else str(student_name)
            rows.append([
                Paragraph(str(row.get(self.roll_col, "")), cell_style),
                Paragraph(student_name.replace(" ", " ", 1), name_style),
            ] + [
                Paragraph(correction_cell(row.get(subject, "")), cell_style)
                for subject in self.subjects
            ])

        usable_width = landscape(A4)[0] - doc.leftMargin - doc.rightMargin
        subject_width = (
            (usable_width - 30 * mm - 50 * mm) / len(self.subjects)
            if self.subjects else usable_width
        )
        column_widths = [30 * mm, 50 * mm] + [
            subject_width
        ] * len(self.subjects)
        table = Table(rows, colWidths=column_widths, repeatRows=1,
                      splitByRow=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#242424")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.white, colors.HexColor("#f8fafc")]),
            ("TOPPADDING", (0, 0), (-1, 0), 8),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
            ("TOPPADDING", (0, 1), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 1), (-1, -1), 5),
        ]))
        story.append(table)
        story.append(Spacer(1, 3 * mm))
        story.append(Paragraph(
            "Passed subjects are shown as –. Blank cells are provided for "
            "failed, OD, and absent subjects.", info_style,
        ))

        doc.build(story)
        return out_path

    def export_counselling_pdf(self, output_dir, college_name="",
                               department="", class_info="",
                               academic_year="", branch="",
                               year_semester="", exam_details="",
                               class_key=""):
        """Export a counselling report for failed and absent subjects."""
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(
            output_dir, f"student_counselling_report_{timestamp}.pdf"
        )

        if self.df is None:
            self.load_data()
        if self.failures_df is None:
            self.analyze()

        doc = SimpleDocTemplate(
            out_path, pagesize=A4,
            leftMargin=12 * mm, rightMargin=12 * mm,
            topMargin=12 * mm, bottomMargin=12 * mm,
        )
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "CounsellingTitle", parent=styles["Title"],
            fontName="Helvetica-Bold", fontSize=14, leading=17,
            alignment=1, spaceAfter=2 * mm,
        )
        header_style = ParagraphStyle(
            "CounsellingHeader", parent=styles["Normal"],
            fontSize=8.5, leading=11, alignment=1,
        )
        cell_style = ParagraphStyle(
            "CounsellingCell", parent=styles["Normal"],
            fontSize=7.5, leading=9, alignment=1,
        )
        left_cell_style = ParagraphStyle(
            "CounsellingLeftCell", parent=cell_style, alignment=0,
        )
        report_detail_style = ParagraphStyle(
            "CounsellingDetail", parent=cell_style,
            fontSize=8, leading=10, spaceAfter=1,
        )
        report_subject_style = ParagraphStyle(
            "CounsellingSubject", parent=cell_style,
            fontSize=8, leading=10, alignment=0, spaceAfter=1,
        )

        story = [Paragraph(
            f"<b>{escape(class_info or 'CAT-II - II Year A Section')}</b>",
            header_style,
        )]
        metadata = [
            ["Branch", branch or department or "-"],
            ["Year / Semester", year_semester or "-"],
            ["Details of Exam", exam_details or academic_year or "-"],
        ]
        metadata_table = Table(
            [[Paragraph(f"<b>{escape(label)}</b>", left_cell_style),
              Paragraph(escape(value), left_cell_style)]
             for label, value in metadata],
            colWidths=[35 * mm, 151 * mm],
        )
        metadata_table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.extend([metadata_table, Spacer(1, 5 * mm)])

        headers = [
            "Roll No", "Name of the<br/>student",
            "Name of the subject &<br/>Faculty",
            "Area of weakness<br/>(Absent / On duty)", "Action taken",
        ]
        rows = [[Paragraph(header, header_style) for header in headers]]
        span_ranges = []
        class_mapping = _class_subject_mapping(
            class_key, class_info, self.raw_text, self.subjects
        )

        for _, row in self.df.iterrows():
            failed_subjects = []
            for subject in self.subjects:
                value = row.get(subject, "")
                normalized = str(value).strip().upper()
                number = pd.to_numeric(value, errors="coerce")
                if normalized in {"AB", "ABSENT", "A"}:
                    weakness = "Absent for Exam"
                    action = "Retest Conducted"
                elif normalized in {"OD", "ON DUTY", "ON-DUTY"}:
                    weakness = "On Duty"
                    action = "OD Recorded"
                elif pd.notna(number) and number < self.pass_mark:
                    weakness = "Subject Understanding"
                    action = "Special Class Conducted"
                else:
                    continue
                failed_subjects.append((subject, weakness, action))

            if not failed_subjects:
                continue
            roll = str(row.get(self.roll_col, ""))
            name = str(row.get(self.name_col, "")) if self.name_col else ""
            block_start = len(rows)
            for idx, (subject, weakness, action) in enumerate(failed_subjects):
                subject_name, faculty_name = class_mapping.get(
                    str(subject), ("Unknown subject", "-")
                )
                rows.append([
                    Paragraph(escape(roll), cell_style) if idx == 0 else "",
                    Paragraph(escape(name), left_cell_style) if idx == 0 else "",
                    Paragraph(
                        f"{escape(subject_name)}<br/><font size='7'>"
                        f"Code: {escape(str(subject))}<br/>"
                        f"Faculty: {escape(faculty_name)}</font>",
                        report_subject_style,
                    ),
                    Paragraph(escape(weakness), report_detail_style),
                    Paragraph(escape(action), report_detail_style),
                ])
            block_end = len(rows) - 1
            if block_end > block_start:
                span_ranges.append((block_start, block_end))

        if len(rows) == 1:
            rows.append([
                Paragraph("-", cell_style),
                Paragraph("No failed or absent subjects", cell_style),
                Paragraph("-", cell_style),
                Paragraph("-", cell_style),
                Paragraph("-", cell_style),
            ])

        table = Table(
            rows,
            colWidths=[20 * mm, 27 * mm, 56 * mm, 43 * mm, 40 * mm],
            repeatRows=1,
        )
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.white, colors.HexColor("#fafafa")]),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (2, 1), (-1, -1), 8),
            ("BOTTOMPADDING", (2, 1), (-1, -1), 8),
        ]))
        for start_row, end_row in span_ranges:
            table.setStyle(TableStyle([
                ("SPAN", (0, start_row), (0, end_row)),
                ("SPAN", (1, start_row), (1, end_row)),
            ]))
        story.append(table)
        doc.build(story)
        return out_path

    def export_counselling_docx(self, output_dir, class_key=""):
        """Export the counselling report as an editable Word document."""
        from docx import Document
        from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Inches, Pt

        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(
            output_dir, f"student_counselling_report_{timestamp}.docx"
        )
        if self.df is None:
            self.load_data()
        if self.failures_df is None:
            self.analyze()

        class_mapping = _class_subject_mapping(
            class_key, "", self.raw_text, self.subjects
        )
        document = Document()
        section = document.sections[0]
        section.top_margin = Inches(0.55)
        section.bottom_margin = Inches(0.55)
        section.left_margin = Inches(0.55)
        section.right_margin = Inches(0.55)

        class_match = re.search(
            r"Year/Sem/Sec\s*:\s*([IVX]+)\s*/\s*(\d+)\s*/\s*([A-Z])",
            self.raw_text,
            re.IGNORECASE,
        )
        exam_match = re.search(r"\bCAT\s*(I{1,3}|IV)\b", self.raw_text, re.I)
        if class_match:
            year, semester, section = class_match.groups()
            class_title = (
                f"{exam_match.group(0).upper() if exam_match else 'Student'}"
                f" - {year.upper()} Year {section.upper()} Section"
            )
            class_value = f"{year.upper()} / {semester} / {section.upper()}"
        else:
            class_title = "Student Counselling Report"
            class_value = "-"

        title = document.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run(class_title)
        run.bold = True
        run.font.size = Pt(12)

        metadata = document.add_table(rows=3, cols=2)
        metadata.style = "Table Grid"
        metadata_values = [
            ("Branch", self._branch_from_source()),
            ("Year / Semester", class_value),
            ("Details of Exam", exam_match.group(0).upper() if exam_match else "-"),
        ]
        for row, (label, value) in zip(metadata.rows, metadata_values):
            row.cells[0].text = label
            row.cells[1].text = value
            row.cells[0].paragraphs[0].runs[0].bold = True

        document.add_paragraph()
        table = document.add_table(rows=1, cols=5)
        table.style = "Table Grid"
        table.autofit = False
        headers = [
            "Roll No", "Name of the student", "Name of the subject & Faculty",
            "Area of weakness (Absent / On duty)", "Action taken",
        ]
        widths = [Inches(0.75), Inches(0.9), Inches(2.0), Inches(1.35), Inches(1.1)]
        for cell, label, width in zip(table.rows[0].cells, headers, widths):
            cell.text = label
            cell.width = width
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = Pt(8)

        student_blocks = []
        for _, row in self.df.iterrows():
            failed_subjects = []
            for subject in self.subjects:
                value = row.get(subject, "")
                normalized = str(value).strip().upper()
                number = pd.to_numeric(value, errors="coerce")
                if normalized in {"AB", "ABSENT", "A"}:
                    weakness, action = "Absent for Exam", "Retest Conducted"
                elif normalized in {"OD", "ON DUTY", "ON-DUTY"}:
                    weakness, action = "On Duty", "OD Recorded"
                elif pd.notna(number) and number < self.pass_mark:
                    weakness, action = (
                        "Subject Understanding", "Special Class Conducted"
                    )
                else:
                    continue
                failed_subjects.append((subject, weakness, action))

            if failed_subjects:
                student_blocks.append((row, failed_subjects))

        for student, failed_subjects in student_blocks:
            first_row = len(table.rows)
            for subject, weakness, action in failed_subjects:
                cells = table.add_row().cells
                subject_name, faculty_name = class_mapping.get(
                    str(subject), ("Unknown subject", "-")
                )
                cells[0].text = str(student.get(self.roll_col, ""))
                cells[1].text = (
                    str(student.get(self.name_col, "")) if self.name_col else ""
                )
                cells[2].text = (
                    f"{subject_name}\nCode: {subject}\nFaculty: {faculty_name}"
                )
                cells[3].text = weakness
                cells[4].text = action
                for cell, width in zip(cells, widths):
                    cell.width = width
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = (
                            WD_ALIGN_PARAGRAPH.LEFT
                            if cell is cells[2]
                            else WD_ALIGN_PARAGRAPH.CENTER
                        )
                        for run in paragraph.runs:
                            run.font.size = Pt(8)

            if len(failed_subjects) > 1:
                last_row = len(table.rows) - 1
                table.cell(first_row, 0).merge(table.cell(last_row, 0))
                table.cell(first_row, 1).merge(table.cell(last_row, 1))

        if not student_blocks:
            cells = table.add_row().cells
            cells[0].text = "-"
            cells[1].text = "No failed or absent subjects"
            for cell in cells[2:]:
                cell.text = "-"

        document.save(out_path)
        return out_path

    def export_correction_docx(self, output_dir):
        """Export the all-subject mark correction matrix as Word."""
        from docx import Document
        from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
        from docx.enum.section import WD_ORIENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Inches, Pt

        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(
            output_dir, f"student_correction_report_{timestamp}.docx"
        )
        if self.df is None:
            self.load_data()
        if self.failures_df is None:
            self.analyze()

        document = Document()
        section = document.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Inches(11)
        section.page_height = Inches(8.5)
        section.top_margin = Inches(0.45)
        section.bottom_margin = Inches(0.45)
        section.left_margin = Inches(0.45)
        section.right_margin = Inches(0.45)
        document.add_heading("Student Detailed Correction Report", level=1)

        headers = ["Roll No", "Name of Student"] + [str(s) for s in self.subjects]
        table = document.add_table(rows=1, cols=len(headers))
        table.style = "Table Grid"
        table.autofit = True
        for cell, label in zip(table.rows[0].cells, headers):
            cell.text = label
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = Pt(8)

        for _, row in self.df.iterrows():
            cells = table.add_row().cells
            name = row.get(self.name_col, "") if self.name_col else ""
            cells[0].text = str(row.get(self.roll_col, ""))
            cells[1].text = "" if pd.isna(name) else str(name)
            for index, subject in enumerate(self.subjects, start=2):
                value = row.get(subject, "")
                normalized = str(value).strip().upper()
                number = pd.to_numeric(value, errors="coerce")
                if normalized in {"OD", "ON DUTY", "ON-DUTY", "AB", "ABSENT", "A"}:
                    rendered = ""
                elif pd.notna(number) and number >= self.pass_mark:
                    rendered = "\N{EN DASH}"
                else:
                    rendered = ""
                cells[index].text = rendered
            for cell in cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in paragraph.runs:
                        run.font.size = Pt(8)

        document.save(out_path)
        return out_path

    def export_full_report_docx(self, output_dir):
        """Export the complete result analysis, including every uploaded subject."""
        from docx import Document
        from docx.enum.section import WD_ORIENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Inches, Pt

        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(output_dir, f"result_report_{timestamp}.docx")
        if self.df is None:
            self.load_data()
        if self.failures_df is None:
            self.analyze()

        document = Document()
        section = document.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Inches(11)
        section.page_height = Inches(8.5)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

        title = document.add_heading("Student Result Analysis Report", level=1)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        branch = self._branch_from_source()
        if branch != "-":
            document.add_paragraph(f"Branch: {branch}")

        stats = self.get_statistics()
        document.add_heading("Class Summary", level=2)
        summary = document.add_table(rows=2, cols=4)
        summary.style = "Table Grid"
        summary_headers = ["Total Students", "Passed", "Failed", "Pass Percentage"]
        summary_values = [
            stats["total_students"], stats["passed_students"],
            stats["failed_students"], f"{stats['pass_percentage']}%",
        ]
        for index, value in enumerate(summary_headers):
            summary.cell(0, index).text = value
            summary.cell(0, index).paragraphs[0].runs[0].bold = True
            summary.cell(1, index).text = str(summary_values[index])

        mapping = _class_subject_mapping("", "", self.raw_text, self.subjects)
        document.add_heading("Subject-wise Statistics", level=2)
        subject_table = document.add_table(rows=1, cols=6)
        subject_table.style = "Table Grid"
        stat_headers = ["Subject Code", "Subject", "Passed", "Failed", "Average", "Maximum"]
        for cell, label in zip(subject_table.rows[0].cells, stat_headers):
            cell.text = label
            cell.paragraphs[0].runs[0].bold = True
        for subject in self.subjects:
            subject_name = mapping.get(subject, (str(subject), "-"))[0]
            row_cells = subject_table.add_row().cells
            values = [
                subject, subject_name,
                stats["subject_pass_counts"][subject],
                stats["subject_fail_counts"][subject],
                stats["subject_avg"][subject],
                stats["subject_max"][subject],
            ]
            for cell, value in zip(row_cells, values):
                cell.text = str(value)

        document.add_heading("Consolidated Mark List", level=2)
        consolidated = self.consolidated_mark_list()
        mark_headers = [str(column) for column in consolidated.columns]
        marks_table = document.add_table(rows=1, cols=len(mark_headers))
        marks_table.style = "Table Grid"
        for cell, label in zip(marks_table.rows[0].cells, mark_headers):
            cell.text = label
            for run in cell.paragraphs[0].runs:
                run.bold = True
                run.font.size = Pt(7)
        for _, data_row in consolidated.iterrows():
            cells = marks_table.add_row().cells
            for cell, column in zip(cells, consolidated.columns):
                value = data_row[column]
                cell.text = "" if pd.isna(value) else str(value)
                for run in cell.paragraphs[0].runs:
                    run.font.size = Pt(7)

        document.add_heading("Subject-wise Round Lists", level=2)
        for subject, students in self.generate_round_lists().items():
            document.add_heading(str(subject), level=3)
            round_table = document.add_table(rows=1, cols=3)
            round_table.style = "Table Grid"
            for cell, label in zip(
                round_table.rows[0].cells, ["S.No", "Roll No", "Marks Obtained"]
            ):
                cell.text = label
                cell.paragraphs[0].runs[0].bold = True
            for student in students:
                cells = round_table.add_row().cells
                cells[0].text = str(student.get("S.No", ""))
                cells[1].text = str(student.get("Reg No", ""))
                cells[2].text = str(student.get("Marks_Obtained", ""))

        document.save(out_path)
        return out_path

    def _branch_from_source(self):
        match = re.search(
            r"B\.?\s*Tech\s*-\s*(.+?)(?:\r?\n|$)",
            self.raw_text,
            re.IGNORECASE,
        )
        return match.group(1).strip() if match else "-"

    # ============================================================
    #  PDF EXPORT
    # ============================================================
    def export_pdf(self, output_dir, college_name="", class_info=""):
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(output_dir, f"round_list_report_{timestamp}.pdf")

        doc = SimpleDocTemplate(
            out_path, pagesize=landscape(A4),
            leftMargin=10 * mm, rightMargin=10 * mm,
            topMargin=12 * mm, bottomMargin=12 * mm
        )
        styles = getSampleStyleSheet()
        story = []

        title_style = ParagraphStyle(
            "TitleStyle", parent=styles["Title"],
            fontSize=14, alignment=1, spaceAfter=4
        )
        sub_style = ParagraphStyle(
            "SubStyle", parent=styles["Normal"],
            fontSize=10, alignment=1, spaceAfter=2
        )
        section_style = ParagraphStyle(
            "SectionStyle", parent=styles["Heading2"],
            fontSize=12, textColor=colors.HexColor("#1e3a8a"),
            spaceBefore=14, spaceAfter=6
        )
        mark_cell_style = ParagraphStyle(
            "MarkCellStyle", parent=styles["Normal"],
            fontSize=7, leading=8, alignment=1,
            wordWrap="LTR",
        )
        mark_header_style = ParagraphStyle(
            "MarkHeaderStyle", parent=mark_cell_style,
            fontName="Helvetica-Bold", textColor=colors.white,
        )

        stats = self.get_statistics()

        # Header
        story.append(Paragraph(college_name or "College Name", title_style))
        story.append(Paragraph("CONSOLIDATED ROUND LIST REPORT", sub_style))
        if class_info:
            story.append(Paragraph(class_info, sub_style))
        story.append(Paragraph(
            f"Generated on: {datetime.now().strftime('%d-%m-%Y %H:%M')}",
            sub_style
        ))
        story.append(Spacer(1, 8 * mm))

        # Class summary
        story.append(Paragraph("Class Summary", section_style))
        summary_data = [
            ["Total Students", "Passed", "Failed", "Pass %"],
            [stats["total_students"], stats["passed_students"],
             stats["failed_students"], f"{stats['pass_percentage']}%"],
        ]
        summary_table = Table(summary_data, colWidths=[45 * mm] * 4)
        summary_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e3a8a")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#f0f4f8")),
        ]))
        story.append(summary_table)
        story.append(Spacer(1, 6 * mm))

        # Subject stats
        story.append(Paragraph("Subject-wise Statistics", section_style))
        headers = ["Subject", "Pass", "Fail", "Avg", "Max", "Pass %"]
        rows = [headers]
        for sub in self.subjects:
            f = stats["subject_fail_counts"][sub]
            p = stats["subject_pass_counts"][sub]
            total_sub = p + f
            pct = round((p / total_sub) * 100, 1) if total_sub else 0
            rows.append([
                str(sub)[:30], p, f,
                stats["subject_avg"][sub], stats["subject_max"][sub],
                f"{pct}%",
            ])

        col_widths = [70 * mm, 18 * mm, 18 * mm, 20 * mm, 20 * mm, 24 * mm]
        sub_table = Table(rows, colWidths=col_widths, repeatRows=1)
        sub_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e3a8a")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("ALIGN", (1, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.white, colors.HexColor("#f8fafc")]),
        ]))
        story.append(sub_table)
        story.append(PageBreak())

        # Consolidated mark list in the same line-wise format as the uploaded sheet.
        story.append(Paragraph("Consolidated Mark List", section_style))
        consolidated = self.consolidated_mark_list()
        mark_headers = list(consolidated.columns)
        mark_rows = [[Paragraph(str(header), mark_header_style)
                      for header in mark_headers]]
        for _, row in consolidated.iterrows():
            rendered = []
            for column in mark_headers:
                value = row[column]
                if column in self.subjects:
                    numeric_value = pd.to_numeric(value, errors="coerce")
                    is_absent = str(value).strip().upper() in {"AB", "ABSENT", "A"}
                    if is_absent or pd.isna(numeric_value) or numeric_value < self.pass_mark:
                        value = CircledMark(value, width=20 * mm)
                rendered.append(value if isinstance(value, Flowable)
                                else Paragraph(str(value), mark_cell_style))
            mark_rows.append(rendered)
        usable_width = landscape(A4)[0] - doc.leftMargin - doc.rightMargin
        mark_widths = (
            [14 * mm, 38 * mm]
            + [20 * mm] * len(self.subjects)
            + [22 * mm, 34 * mm, 24 * mm]
        )
        scale = usable_width / sum(mark_widths)
        mark_widths = [width * scale for width in mark_widths]
        mark_table = Table(mark_rows, colWidths=mark_widths, repeatRows=1)
        mark_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e3a8a")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.white, colors.HexColor("#f8fafc")]),
        ]))
        story.append(mark_table)
        story.append(PageBreak())

        # Round lists
        story.append(Paragraph("Subject-wise Round Lists", section_style))
        story.append(Spacer(1, 3 * mm))

        for sub, students in self.generate_round_lists().items():
            story.append(Paragraph(f"Subject: {sub}", styles["Heading3"]))
            story.append(Paragraph(
                f"Total failed: {len(students)} | Pass mark: {self.pass_mark}",
                styles["Normal"]
            ))
            story.append(Spacer(1, 2 * mm))

            data = [["#", "S.No", "Reg No", "Marks Obtained"]]
            for i, s in enumerate(students, 1):
                data.append([i, s.get("S.No", ""), s["Reg No"], s["Marks_Obtained"]])

            t = Table(data, colWidths=[15 * mm, 20 * mm, 65 * mm, 40 * mm],
                      repeatRows=1)
            t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e3a8a")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("ALIGN", (0, 0), (0, -1), "CENTER"),
                ("ALIGN", (3, 0), (3, -1), "CENTER"),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1),
                 [colors.white, colors.HexColor("#fef2f2")]),
            ]))
            story.append(t)
            story.append(Spacer(1, 8 * mm))

        # Student-wise failure report
        story.append(PageBreak())
        story.append(Paragraph("Student-wise Failure Report", section_style))
        student_rows = [["S.No", "Reg No", "Failed Subjects", "Count",
                         "Pass %"]]
        for _, r in self.get_student_wise_failures().iterrows():
            if r["Status"] == "FAIL":
                student_rows.append([
                    r.get("S.No", ""), r["Reg No"],
                    str(r["Failed_Subjects"])[:60], r["Total_Failures"],
                    f"{r['Pass_Percentage']}%",
                ])
        t = Table(student_rows,
                  colWidths=[15 * mm, 40 * mm, 85 * mm, 18 * mm, 22 * mm],
                  repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e3a8a")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.white, colors.HexColor("#f8fafc")]),
        ]))
        story.append(t)

        doc.build(story)
        return out_path

    # ============================================================
    #  EXCEL EXPORT
    # ============================================================
    def export_report(self, output_dir):
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(output_dir, f"result_report_{timestamp}.xlsx")

        stats_df = pd.DataFrame(
            list(self.get_statistics()["subject_fail_counts"].items()),
            columns=["Subject", "Fail_Count"]
        )
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            self.consolidated_mark_list().to_excel(
                writer, sheet_name="Consolidated_Marks", index=False
            )
            self.get_student_wise_failures().to_excel(
                writer, sheet_name="Student_Wise_Failures", index=False
            )
            stats_df.to_excel(writer, sheet_name="Subject_Stats", index=False)
            for sub, rows in self.generate_round_lists().items():
                sheet_name = f"Round_{sub}"[:31]
                pd.DataFrame(rows).to_excel(
                    writer, sheet_name=sheet_name, index=False
                )
        return out_path