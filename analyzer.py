"""
Analyzer for college mark lists — supports PDF / Excel / CSV
Handles consolidated format like: S.No | RegNo | Subj1...SubjN | Total
"""
import os
import re
import shutil
from io import BytesIO
from datetime import datetime
from xml.sax.saxutils import escape
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                Paragraph, Spacer, PageBreak, Flowable, Image)

CLASS_SUBJECT_FACULTY = {
    "II/3/A": {
        "23BS305": ("Discrete Mathematics", "Mythili D"),
        "23MC003": ("Indian Constitution", "Ramesh Babu"),
        "23TPS03": ("Quantitative Aptitude and Logical Reasoning - I",
                    "Tamilselvan P"),
        "23HS301": ("Professional Communication", ""),
        "23CS301": ("Data Structures", "Saravanan G"),
        "23AD301": ("Artificial Intelligence", "Sathish Kumar P"),
        "23AD302": ("Object Oriented Programming Using Python (Embedded)",
                    "Ragavi K"),
        "23AD303": ("Digital Computer Organization and Architecture",
                    "Gowtham C"),
    },
    "II/3/B": {
        "23BS305": ("Discrete Mathematics", "Mythili D"),
        "23MC003": ("Indian Constitution", "Ramesh Babu"),
        "23TPS03": ("Quantitative Aptitude and Logical Reasoning - I",
                    "Tamilselvan P"),
        "23HS301": ("Professional Communication", ""),
        "23CS301": ("Data Structures", "Saravanan G"),
        "23AD301": ("Artificial Intelligence", "Sathish Kumar P"),
        "23AD302": ("Object Oriented Programming Using Python (Embedded)",
                    "Ragavi K"),
        "23AD303": ("Digital Computer Organization and Architecture",
                    "Gowtham C"),
    },
    "II/3/C": {
        "23BS305": ("Discrete Mathematics", "Padma Priya J"),
        "23MC003": ("Indian Constitution", "Ramesh Babu"),
        "23TPS03": ("Quantitative Aptitude and Logical Reasoning - I",
                    "Tamilselvan P"),
        "23HS301": ("Professional Communication", ""),
        "23CS301": ("Data Structures", "Saravanan G"),
        "23AD301": ("Artificial Intelligence", "Sathish Kumar P"),
        "23AD302": ("Object Oriented Programming Using Python (Embedded)",
                    "Ragavi K"),
        "23AD303": ("Digital Computer Organization and Architecture",
                    "Linett Sophia D"),
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

SUBJECT_WEAKNESS_TOPIC = {
    "23BS305": "Partial Ordering",
    "23MC003": "State Election commission",
    "23AD303": "Memory Management",
    "23AD301": "Statistical Learning",
    "23CS301": "Binary Search Trees",
    "23TPS03": "Probability & Permutations",
    "23AD501": "Linear Regression & Gradient Descent",
    "23AD502": "Dimensionality Reduction",
    "23ADX08": "Inheritance & Exception Handling",
    "23TPS05": "Logical Reasoning & Syllogisms",
    "23ADX11": "Data Manipulation in R",
    "23AD503": "MapReduce Architecture",
    "23AD701": "Virtualization & Cloud Storage",
    "23AD702": "Syntactic Parsing & Word Embeddings",
    "23AD703": "Time Series Forecasting",
    "23ADX37": "Edge Gateway & IoT Architecture",
    "23ADX48": "Enterprise Resource Planning",
    "23ECY01": "CMOS Inverter Characteristics",
}

STUDENT_ROSTER = {
    # II Year AIDS A (from physical counseling report & records)
    "ES24AD01": "Aadhavan K",
    "ES24AD02": "Aakash H",
    "ES24AD03": "Aarthi S",
    "ES24AD04": "Abdul Rahman M",
    "ES24AD05": "Abhinav K",
    "ES24AD06": "Abishek R",
    "ES24AD07": "Ajay S",
    "ES24AD08": "Akash M",
    "ES24AD09": "Akshaya S",
    "ES24AD10": "Anand P",
    "ES24AD11": "Ananya R",
    "ES24AD12": "Archana M",
    "ES24AD13": "Arjun V",
    "ES24AD14": "Arun Kumar G",
    "ES24AD15": "Balamurugan S",
    "ES24AD16": "Barath K",
    "ES24AD17": "Boopathi S",
    "ES24AD18": "Chandru M",
    "ES24AD19": "Deepak R",
    "ES24AD20": "Dharani P",
    # III Year AIDS B (from official docx)
    "ES23AD62": "KEERTHIK SARAVANA D",
    "ES23AD63": "KIRAN G",
    "ES23AD69": "KOWSHICKRAJA R D",
    "ES23AD70": "KUMAR V",
    "ES23AD77": "MANIKANDA RAJA  S",
    "ES23AD82": "MOHAMMED RAFIK J",
    "ES23AD83": "MOHAN KUMAR B",
    "ES23AD87": "MONISH KUMAR K",
    "ES23AD88": "MONISHA L T",
    "ES23AD89": "MOULI L",
    "ES23AD90": "MUHILAN M R",
    "ES23AD101": "NAVEENKRISHNA M",
    "ES23AD103": "NAVEEN T",
    "ES23AD106": "NIKKESH RAJ S",
    "ES23AD107": "NIRANJAN V",
    "ES23AD108": "NISHANTH S",
    "ES23AD109": "NISHANTH S",
    "ES23AD110": "NITHIN KUMAR K",
    "ES23AD111": "PONSRIDHAR N",
    "ES23AD114": "PRASANNA M",
    "ES23AD115": "PRASANNA S",
    "ES23AD116": "PRATHAP M",
    "ES23AD118": "PRIVINKUMAR S",
    "ES23AD119": "PRIYA DHARSHINI A",
    "ES23AD121": "RAGHAVENDRAN G",
    "ES23AD122": "RAHUL K",
    "ES23AD123": "RAJASHREE M",
    "ES23ADL04": "SATHIYAN K SSM",
    # II Year AIDS C (from attendance register)
    "ES25AD126": "SANJAY P",
    "ES25AD127": "SANJITH RAKUL",
    "ES25AD128": "SANTHOSH R",
    "ES25AD130": "SARANRAJ M",
    "ES25AD132": "SARATHY",
    "ES25AD133": "SARAVANAN R",
    "ES25AD135": "SATHYAN K",
    "ES25AD136": "SHAARVESH R",
    "ES25AD138": "SHASHTIKA R",
    "ES25AD140": "SHRENKARTHIC",
    "ES25AD146": "SOUNDARYA K",
    "ES25AD150": "SRI DHAYA N",
    "ES25AD151": "SRIMATHI S",
    "ES25AD152": "SRIRAMU S",
    "ES25AD156": "SUJITHA S",
    "ES25AD157": "TAMILARASAN S",
    "ES25AD158": "TAMILSELVAN S",
}


def to_roman_sem(sem_str):
    mapping = {
        "1": "I", "2": "II", "3": "III", "4": "IV",
        "5": "V", "6": "VI", "7": "VII", "8": "VIII",
    }
    s = str(sem_str).strip()
    return mapping.get(s, mapping.get(s.upper(), s.upper()))


def format_college_roll(reg_no, dept_code="AD"):
    s = str(reg_no).strip()
    if re.match(r"^ES\d+[A-Z]+\d+$", s, re.I):
        return s.upper()
    if len(s) == 16 and s.isdigit():
        batch = s[8:10]
        roll_num = s[-3:].lstrip("0")
        if len(roll_num) == 1:
            roll_num = "0" + roll_num
        elif not roll_num:
            roll_num = "00"
        return f"ES{batch}{dept_code}{roll_num}"
    return s


def resolve_student_name(student_row, roll, name_col=None):
    if name_col and student_row.get(name_col) and str(student_row.get(name_col)).strip():
        val = str(student_row.get(name_col)).strip()
        if val.lower() != "none" and val != "-":
            return val
    short_roll = format_college_roll(roll)
    if short_roll in STUDENT_ROSTER:
        return STUDENT_ROSTER[short_roll]
    if str(roll) in STUDENT_ROSTER:
        return STUDENT_ROSTER[str(roll)]
    m = re.search(r"AD(\d+)$", short_roll, re.I)
    if m:
        num = m.group(1)
        for k, v in STUDENT_ROSTER.items():
            if k.endswith(num) or k.endswith(f"AD{num}"):
                return v
    return ""


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
        self.display_subjects = []
        self.ignored_subjects = []
        self.failures_df = None
        self.roll_col = None
        self.name_col = None
        self.subject_names = subject_names or []
        self.pdf_optional_columns = []
        self.raw_text = ""
        self._numeric_df = None
        self._source_rotation = 0

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
        self.display_subjects = [c for c in columns if c not in excluded]
        if self._is_third_year():
            self.ignored_subjects = [
                c for c in self.display_subjects
                if re.search(r"\bOE\b|23MEY06", str(c), re.IGNORECASE)
            ]
        self.display_subjects = [
            c for c in self.display_subjects if c not in self.ignored_subjects
        ]
        self.subjects = list(self.display_subjects)

    def _is_third_year(self):
        source = f"{os.path.basename(self.filepath)} {self.raw_text}".replace(
            "_", " "
        )
        return bool(re.search(
            r"\bIII\s+YEAR\b|\b3(?:RD)?\s+YEAR\b|"
            r"\bYEAR\s*/\s*SEM\s*/\s*SEC\s*:\s*III\b",
            source,
            re.IGNORECASE,
        ))

    # ============================================================
    #  PDF PARSER
    # ============================================================
    def _load_pdf(self):
        """Extract student data from a consolidated mark list PDF."""
        try:
            import pypdfium2 as pdfium
        except ImportError as exc:
            raise ValueError(
                "PDF support requires pypdfium2 to be installed."
            ) from exc

        all_text = []
        with pdfium.PdfDocument(self.filepath) as document:
            for page_index in range(len(document)):
                page = document[page_index]
                text_page = page.get_textpage()
                txt = text_page.get_text_range() or ""
                if txt:
                    all_text.append(txt)
                page.close()

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
                if self._is_third_year():
                    self.ignored_subjects.append(code)
                self.pdf_optional_columns = []

        # Extract student rows — multiple strategies
        students = self._extract_student_rows(full_text)
        if not students:
            ocr_text = self._extract_pdf_text_with_ocr()
            if ocr_text:
                self.raw_text = f"{full_text}\n{ocr_text}".strip()
                self.subject_names = []
                self.pdf_optional_columns = []
                self._extract_subject_headers(ocr_text)
                self._extract_subjects_from_summary(ocr_text)
                students = self._extract_student_rows(ocr_text)
        print(f"=== DETECTED {len(students)} STUDENT ROWS ===\n")

        if not students:
            raise ValueError(
                "No student rows detected in PDF, including OCR. "
                "Check that the PDF contains a clear mark-list table."
            )

        # Pad all students to same subject count
        max_subj = max(len(s["marks"]) for s in students)
        has_names = any(s.get("name") for s in students)
        cols = ["S.No", "Reg No"]
        if has_names:
            cols.append("Name")
        cols += [f"Subject{i+1}" for i in range(max_subj)] + ["Total"]

        rows = []
        for s in students:
            marks = list(s["marks"])
            while len(marks) < max_subj:
                marks.append(None)
            row = [s["sno"], s["regno"]]
            if has_names:
                row.append(s.get("name", ""))
            rows.append(row + marks + [s["total"]])

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
        self.name_col = "Name" if "Name" in self.df.columns else None
        self.display_subjects = [
            c for c in self.df.columns
            if c not in ("S.No", "Reg No", "Name", "Total")
            and c not in self.ignored_subjects
        ]
        self.subjects = list(self.display_subjects)

        print(f"OK Roll column: {self.roll_col}")
        print(f"OK Subjects: {self.subjects}\n")

    def _extract_pdf_text_with_ocr(self):
        """OCR image-only PDF pages so scanned mark lists can be analyzed."""
        try:
            import pypdfium2 as pdfium
            import pytesseract
        except ImportError as exc:
            raise ValueError(
                "This PDF is image-only. OCR support requires pypdfium2 and "
                "pytesseract to be installed."
            ) from exc

        if not shutil.which("tesseract"):
            for candidate in (
                r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            ):
                if os.path.isfile(candidate):
                    pytesseract.pytesseract.tesseract_cmd = candidate
                    break

        ocr_pages = []
        with pdfium.PdfDocument(self.filepath) as document:
            for page_index in range(len(document)):
                page = document[page_index]
                bitmap = page.render(scale=3.0)
                image = bitmap.to_pil()
                try:
                    rotated = image.rotate(270, expand=True)
                    ocr_pages.append(
                        pytesseract.image_to_string(
                            rotated,
                            lang="eng",
                            timeout=120,
                        )
                    )
                    rotated.close()
                except pytesseract.TesseractNotFoundError as exc:
                    raise ValueError(
                        "This PDF is image-only. Install the Tesseract OCR "
                        "engine before uploading scanned PDFs."
                    ) from exc
                bitmap.close()
                page.close()
        return "\n".join(ocr_pages)

    def _extract_subject_headers(self, text):
        """Read subject codes from the consolidated table header when available."""
        header_re = re.compile(
            r"^\s*S\.?No\s+Reg\s+No\s+(.+?)\s+Total\s+Mark(?:\s+MARK)?\s*$",
            re.IGNORECASE | re.MULTILINE,
        )
        match = header_re.search(text)
        if not match:
            code_matches = re.findall(
                r"\b\d{2}[A-Z]{2,4}\d{2,3}\b", text, re.IGNORECASE
            )
            self.subject_names = list(dict.fromkeys(code_matches))
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

        if subjects and not self.subject_names:
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

        # OCR format: S.No, alphanumeric register number, student name,
        # followed by one mark token per subject and no total column.
        if not students and self.subject_names:
            ocr_row_re = re.compile(
                r"^\s*(\d{1,3})\s+([A-Z]{2}\d{2}[A-Z]{2}\d{2,})\s+(.+?)\s*$",
                re.IGNORECASE | re.MULTILINE,
            )
            mark_token_re = re.compile(
                r"^(?:AB|ABSENT|A|-|\d{1,3})$", re.IGNORECASE
            )
            for match in ocr_row_re.finditer(text):
                sno, regno, remainder = match.groups()
                tokens = remainder.split()
                mark_start = len(tokens)
                while mark_start and mark_token_re.fullmatch(
                    tokens[mark_start - 1]
                ):
                    mark_start -= 1
                mark_tokens = tokens[mark_start:]
                if len(mark_tokens) < len(self.subject_names):
                    continue
                mark_tokens = mark_tokens[-len(self.subject_names):]
                name = " ".join(tokens[:mark_start]).strip()
                cleaned = self._clean_marks(mark_tokens)
                numeric_marks = pd.to_numeric(cleaned, errors="coerce")
                total = int(pd.Series(numeric_marks).dropna().sum())
                students.append({
                    "sno": sno,
                    "regno": regno,
                    "name": name,
                    "marks": cleaned,
                    "total": total,
                })

        # ---------- Strategy 2: RegNo-anchored fallback ----------
        if not students:
            print("Strategy 1 failed - trying RegNo-anchor strategy...")
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
        for sub in self.display_subjects:
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
                    student_name = row.get(self.name_col, "") if self.name_col else ""
                    rows.append({
                        "S.No": row.get("S.No", ""),
                        "Reg No": row[self.roll_col],
                        "Name": "" if pd.isna(student_name) else str(student_name),
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
            marks = [row.get(subject, "") for subject in self.display_subjects]
            total = sum(
                float(numeric.iloc[index][subject])
                for subject in self.subjects
                if pd.notna(numeric.iloc[index][subject])
            )
            total = int(total) if isinstance(total, float) and total.is_integer() else total
            rows.append({
                "S.No": row.get(serial_col, index + 1) if serial_col else index + 1,
                "Reg No": row.get(self.roll_col, ""),
                **{
                    subject: mark for subject, mark in
                    zip(self.display_subjects, marks)
                },
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
            if pd.isna(value) or not normalized or normalized == "-":
                return "-"
            numeric_value = pd.to_numeric(value, errors="coerce")
            if pd.notna(numeric_value) and numeric_value >= self.pass_mark:
                return "\N{EN DASH}"
            return ""

        story = []
        self._add_pdf_source_header(story, info_style, target_width=220 * mm)

        dept_line = next(
            (
                line.strip()
                for line in self.raw_text.splitlines()
                if "DEPARTMENT OF" in line.upper()
            ),
            "",
        )
        if not dept_line:
            branch = self._branch_from_source()
            if branch and branch != "-":
                dept_line = f"DEPARTMENT OF {branch.upper()}"
            else:
                dept_line = "DEPARTMENT OF ARTIFICIAL INTELLIGENCE AND DATA SCIENCE"

        story.append(Paragraph(f"<b>{dept_line}</b>", title_style))

        exam_match = re.search(
            r"\bCAT\s*[-–—]?\s*(I{1,3}|IV|\d+)\b", self.raw_text, re.I
        )
        exam_str = exam_match.group(0).upper().replace("–", "-") if exam_match else "CAT I"

        class_match = re.search(
            r"Year/Sem/Sec\s*:\s*([IVX]+|\d+)\s*/\s*(\d+)\s*/\s*([A-Z])",
            self.raw_text,
            re.I,
        )
        if class_match:
            year, sem, sec = class_match.groups()
            class_title = f"{exam_str} – {year.upper()} Year {sec.upper()} Sec"
        elif class_info:
            class_title = f"{exam_str} – {class_info}"
        else:
            class_title = f"{exam_str} – III Year B Sec"

        story.append(Paragraph(f"<b>{class_title}</b>", info_style))
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
        correction_mapping = _class_subject_mapping(
            "", class_info, self.raw_text, self.subjects
        )
        headers = [
            Paragraph("<b>Roll No</b>", header_style),
            Paragraph("<b>Name of the<br/>Student</b>", header_style),
        ] + [
            Paragraph(
                (
                    f"<b>{escape(str(subject))} &amp; "
                    f"{escape(correction_mapping.get(subject, ('', ''))[0])}</b>"
                )
                if correction_mapping.get(subject, ("", ""))[0]
                else f"<b>{escape(str(subject))}</b>",
                header_style,
            )
            for subject in self.display_subjects
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
                for subject in self.display_subjects
            ])

        usable_width = landscape(A4)[0] - doc.leftMargin - doc.rightMargin
        subject_width = (
            (usable_width - 30 * mm - 50 * mm) / len(self.display_subjects)
            if self.display_subjects else usable_width
        )
        column_widths = [30 * mm, 50 * mm] + [
            subject_width
        ] * len(self.display_subjects)
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
        story.append(Spacer(1, 4 * mm))

        dept_ref = dept_line.upper()
        if "ARTIFICIAL" in dept_ref or "AI" in dept_ref:
            hod_label = "HOD/AI&DS"
        elif "COMPUTER" in dept_ref or "CSE" in dept_ref:
            hod_label = "HOD/CSE"
        elif "COMMUNICATION" in dept_ref or "ECE" in dept_ref:
            hod_label = "HOD/ECE"
        elif "ELECTRICAL" in dept_ref or "EEE" in dept_ref:
            hod_label = "HOD/EEE"
        elif "MECHANICAL" in dept_ref or "MECH" in dept_ref:
            hod_label = "HOD/MECH"
        elif "CIVIL" in dept_ref:
            hod_label = "HOD/CIVIL"
        elif "INFORMATION" in dept_ref or "IT" in dept_ref:
            hod_label = "HOD/IT"
        else:
            hod_label = "HOD/AI&DS"

        hod_style = ParagraphStyle(
            "CorrectionHOD", parent=styles["Normal"],
            fontSize=10, leading=12, alignment=2, fontName="Helvetica-Bold",
        )
        story.append(Paragraph(f"<b>{hod_label}</b>", hod_style))

        doc.build(story)
        return out_path

    def export_counselling_pdf(self, output_dir, college_name="",
                               department="", class_info="",
                               academic_year="", branch="",
                               year_semester="", exam_details="",
                               class_key=""):
        """Export the counselling report matching the official format."""
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.lib.units import mm
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Table, TableStyle, Image as RLImage, Spacer
        )
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from xml.sax.saxutils import escape

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
            leftMargin=14 * mm, rightMargin=14 * mm,
            topMargin=12 * mm, bottomMargin=12 * mm,
        )
        styles = getSampleStyleSheet()

        header_style = ParagraphStyle(
            "CounsellingHeader", parent=styles["Normal"],
            fontSize=8.5, leading=11, alignment=1, fontName="Helvetica-Bold",
        )
        cell_style = ParagraphStyle(
            "CounsellingCell", parent=styles["Normal"],
            fontSize=8.5, leading=11, alignment=1,
        )
        left_cell_style = ParagraphStyle(
            "CounsellingLeftCell", parent=cell_style, alignment=0,
        )

        story = []

        # 1. White Institutional Header with 3 Logos
        base_dir = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(base_dir, "static", "college_logo.png")
        naac_path = os.path.join(base_dir, "static", "naac_logo.png")
        inn_path = os.path.join(base_dir, "static", "innovation_logo.png")

        left_logo = RLImage(logo_path, width=22 * mm, height=22 * mm) if os.path.exists(logo_path) else ""
        right_top = RLImage(naac_path, width=16 * mm, height=16 * mm) if os.path.exists(naac_path) else ""
        right_bot = RLImage(inn_path, width=20 * mm, height=10 * mm) if os.path.exists(inn_path) else ""
        right_cell = Table([[right_top], [Spacer(1, 2 * mm)], [right_bot]], colWidths=[24 * mm])
        right_cell.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))

        center_html = """
        <para align="center">
        <b><font size="12">ERODE SENGUNTHAR</font></b><br/>
        <b><font size="10">ENGINEERING COLLEGE</font></b><br/>
        <font size="7.5">(An Autonomous Institution)</font><br/>
        <font size="6.5">Approved by AICTE, New Delhi, Permanently Affiliated to Anna University- Chennai,<br/>
        Accredited by National Board of Accreditation (NBA), New Delhi &amp;<br/>
        National Assessment and Accreditation Council (NAAC), Bangalore with 'A' Grade</font><br/>
        <b><font size="7.5">PERUNDURAI – 638 057, TAMILNADU, INDIA.</font></b>
        </para>
        """
        center_p = Paragraph(center_html, styles["Normal"])

        head_table = Table([[left_logo, center_p, right_cell]], colWidths=[25 * mm, 132 * mm, 25 * mm])
        head_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))
        story.append(head_table)
        story.append(Spacer(1, 4 * mm))

        # 2. Department & Title
        branch_resolved = branch or department or self._branch_from_source()
        if not branch_resolved or branch_resolved == "-":
            branch_resolved = "ARTIFICIAL INTELLIGENCE AND DATA SCIENCE"
        branch_clean = branch_resolved.replace("DEPARTMENT OF", "").replace("DEPARTMENT", "").strip().upper()

        dept_p = Paragraph(f"<b><font size='10'>DEPARTMENT OF {escape(branch_clean)}</font></b>", header_style)
        title_p = Paragraph("<b><font size='10'>STUDENT COUNSELLING REPORT</font></b>", header_style)
        story.extend([dept_p, Spacer(1, 1.5 * mm), title_p, Spacer(1, 3 * mm)])

        # 3. Exam & Class Line
        exam_match = re.search(r"\bCAT\s*[-–—]?\s*(I{1,3}|IV|\d+)\b", self.raw_text, re.I)
        exam_str = exam_details or (exam_match.group(0).upper().replace("–", "-") if exam_match else "CAT II")

        class_match = re.search(
            r"Year/Sem/Sec\s*:\s*([IVX]+|\d+)\s*/\s*(\d+)\s*/\s*([A-Z])",
            self.raw_text,
            re.I,
        )
        if class_match:
            year_val, sem_val, sec_val = class_match.groups()
            class_line = f"{exam_str}– {year_val.upper()} Year {sec_val.upper()} Section"
            sem_display = f"{year_val.upper()} / {to_roman_sem(sem_val)}"
        else:
            class_line = class_info or f"{exam_str}– II Year A Section"
            sem_display = year_semester or "II / III"

        class_p = Paragraph(f"<b><font size='9'>{escape(class_line)}</font></b>", left_cell_style)
        story.extend([class_p, Spacer(1, 2 * mm)])

        # 4. Borderless Metadata Table
        acad_match = re.search(r"(20\d\d\s*[-–/]\s*\d{2,4})", self.raw_text)
        acad_year = academic_year or (acad_match.group(1).replace(" ", "") if acad_match else "2024-25")
        meta_rows = [
            [Paragraph("<b>Branch</b>", left_cell_style), Paragraph(f": {escape(branch_clean)}", left_cell_style)],
            [Paragraph("<b>Year / Semester</b>", left_cell_style), Paragraph(f": {escape(sem_display)}", left_cell_style)],
            [Paragraph("<b>Details of Exam</b>", left_cell_style), Paragraph(f": {escape(exam_str)} ({escape(acad_year)})", left_cell_style)],
        ]
        meta_table = Table(meta_rows, colWidths=[35 * mm, 147 * mm])
        meta_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ]))
        story.extend([meta_table, Spacer(1, 4 * mm)])

        # 5. Main Table
        headers = [
            "<b>Roll No.</b>",
            "<b>Name of the<br/>student</b>",
            "<b>Name of the subject &amp;<br/>Faculty</b>",
            "<b>Area of weakness/<br/>Absent / On duty</b>",
            "<b>Action taken</b>",
        ]
        t_rows = [[Paragraph(h, header_style) for h in headers]]
        col_widths = [24 * mm, 34 * mm, 61 * mm, 34 * mm, 29 * mm]
        span_ranges = []

        class_mapping = _class_subject_mapping(
            class_key, class_info, self.raw_text, self.subjects
        )

        for _, row in self.df.iterrows():
            failed_subs = []
            for sub in self.subjects:
                val = row.get(sub, "")
                norm = str(val).strip().upper()
                num = pd.to_numeric(val, errors="coerce")
                if norm in {"AB", "ABSENT", "A"}:
                    w_text = "Absent for Exam"
                    act_text = "Retest Conducted"
                elif norm in {"OD", "ON DUTY", "ON-DUTY"}:
                    w_text = "On Duty"
                    act_text = "OD Recorded"
                elif pd.notna(num) and num < self.pass_mark:
                    w_text = SUBJECT_WEAKNESS_TOPIC.get(str(sub), "Partial Ordering")
                    act_text = "Special Class Conducted"
                else:
                    continue
                failed_subs.append((sub, w_text, act_text))

            if not failed_subs:
                continue

            raw_roll = row.get(self.roll_col, "")
            college_roll = format_college_roll(raw_roll)
            student_name = resolve_student_name(row, raw_roll, self.name_col)
            block_start = len(t_rows)

            for idx, (sub, w_text, act_text) in enumerate(failed_subs):
                sub_name, fac_name = class_mapping.get(str(sub), (sub, "-"))
                sub_fac_cell = f"<b>{escape(str(sub))}- {escape(sub_name)}</b><br/><br/>{escape(fac_name)}"
                t_rows.append([
                    Paragraph(escape(college_roll), cell_style) if idx == 0 else "",
                    Paragraph(escape(student_name), cell_style) if idx == 0 else "",
                    Paragraph(sub_fac_cell, left_cell_style),
                    Paragraph(escape(w_text), cell_style),
                    Paragraph(escape(act_text), cell_style),
                ])
            block_end = len(t_rows) - 1
            if block_end > block_start:
                span_ranges.append((block_start, block_end))

        if len(t_rows) == 1:
            t_rows.append([
                Paragraph("-", cell_style),
                Paragraph("No failed or absent subjects", cell_style),
                Paragraph("-", cell_style),
                Paragraph("-", cell_style),
                Paragraph("-", cell_style),
            ])

        table = Table(t_rows, colWidths=col_widths, repeatRows=1)
        t_style = [
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("ALIGN", (2, 1), (2, -1), "LEFT"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]
        for start_r, end_r in span_ranges:
            t_style.append(("SPAN", (0, start_r), (0, end_r)))
            t_style.append(("SPAN", (1, start_r), (1, end_r)))
        table.setStyle(TableStyle(t_style))
        story.append(table)
        story.append(Spacer(1, 10 * mm))

        # 6. Footer Signatures
        hod_label = "HOD/AI&DS" if "AI" in branch_clean or "ARTIFICIAL" in branch_clean else f"HOD/{branch_clean[:4]}"
        sig_table = Table([
            [Paragraph("<b>Class Advisor / Counsellor</b>", left_cell_style),
             Paragraph(f"<b>{escape(hod_label)}</b>", ParagraphStyle("HODStyle", parent=cell_style, alignment=2))]
        ], colWidths=[91 * mm, 91 * mm])
        sig_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ]))
        story.append(sig_table)

        doc.build(story)
        return out_path

    def export_counselling_docx(self, output_dir, department="", class_key=""):
        """Export the counselling report matching the official format as Word."""
        from docx import Document
        from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Inches, Pt, RGBColor

        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(
            output_dir, f"student_counselling_report_{timestamp}.docx"
        )
        if self.df is None:
            self.load_data()
        if self.failures_df is None:
            self.analyze()

        document = Document()
        sec = document.sections[0]
        sec.top_margin = Inches(0.5)
        sec.bottom_margin = Inches(0.5)
        sec.left_margin = Inches(0.55)
        sec.right_margin = Inches(0.55)

        # 1. White Institutional Header with 3 Logos
        base_dir = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(base_dir, "static", "college_logo.png")
        naac_path = os.path.join(base_dir, "static", "naac_logo.png")
        inn_path = os.path.join(base_dir, "static", "innovation_logo.png")

        logo_table = document.add_table(rows=1, cols=3)
        logo_table.autofit = False
        cells = logo_table.rows[0].cells
        cells[0].width = Inches(1.1)
        cells[1].width = Inches(5.1)
        cells[2].width = Inches(1.1)

        # Left logo
        p_left = cells[0].paragraphs[0]
        p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists(logo_path):
            p_left.add_run().add_picture(logo_path, width=Inches(0.9))

        # Center institutional text
        p_center = cells[1].paragraphs[0]
        header_lines = [
            ("ERODE SENGUNTHAR", 13, True),
            ("ENGINEERING COLLEGE", 11, True),
            ("(An Autonomous Institution)", 8, False),
            ("Approved by AICTE, New Delhi, Permanently Affiliated to Anna University- Chennai,", 7, False),
            ("Accredited by National Board of Accreditation (NBA), New Delhi &", 7, False),
            ("National Assessment and Accreditation Council (NAAC), Bangalore with 'A' Grade", 7, False),
            ("PERUNDURAI – 638 057, TAMILNADU, INDIA.", 8, True),
        ]
        for idx, (text, size, bold) in enumerate(header_lines):
            p = p_center if idx == 0 else cells[1].add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
            run = p.add_run(text)
            run.bold = bold
            run.font.size = Pt(size)
            run.font.color.rgb = RGBColor(0, 0, 0)

        # Right logos
        p_right = cells[2].paragraphs[0]
        p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_right.paragraph_format.space_after = Pt(2)
        if os.path.exists(naac_path):
            p_right.add_run().add_picture(naac_path, width=Inches(0.65))
        if os.path.exists(inn_path):
            p_right2 = cells[2].add_paragraph()
            p_right2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_right2.add_run().add_picture(inn_path, width=Inches(0.85))

        # 2. Department & Report Title
        branch = department or self._branch_from_source()
        if not branch or branch == "-":
            branch = "ARTIFICIAL INTELLIGENCE AND DATA SCIENCE"
        branch_clean = branch.replace("DEPARTMENT OF", "").replace("DEPARTMENT", "").strip().upper()

        p_dept = document.add_paragraph()
        p_dept.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_dept.paragraph_format.space_before = Pt(8)
        p_dept.paragraph_format.space_after = Pt(2)
        r_dept = p_dept.add_run(f"DEPARTMENT OF {branch_clean}")
        r_dept.bold = True
        r_dept.font.size = Pt(11)

        p_title = document.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_before = Pt(0)
        p_title.paragraph_format.space_after = Pt(6)
        r_title = p_title.add_run("STUDENT COUNSELLING REPORT")
        r_title.bold = True
        r_title.font.size = Pt(11)

        # 3. Exam & Class Line
        exam_match = re.search(r"\bCAT\s*[-–—]?\s*(I{1,3}|IV|\d+)\b", self.raw_text, re.I)
        exam_str = exam_match.group(0).upper().replace("–", "-") if exam_match else "CAT II"

        class_match = re.search(
            r"Year/Sem/Sec\s*:\s*([IVX]+|\d+)\s*/\s*(\d+)\s*/\s*([A-Z])",
            self.raw_text,
            re.I,
        )
        if class_match:
            year_val, sem_val, sec_val = class_match.groups()
            class_line = f"{exam_str}– {year_val.upper()} Year {sec_val.upper()} Section"
            sem_display = f"{year_val.upper()} / {to_roman_sem(sem_val)}"
        else:
            class_line = f"{exam_str}– II Year A Section"
            sem_display = "II / III"

        p_class = document.add_paragraph()
        p_class.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_class.paragraph_format.space_before = Pt(0)
        p_class.paragraph_format.space_after = Pt(4)
        r_class = p_class.add_run(class_line)
        r_class.bold = True
        r_class.font.size = Pt(10)

        # 4. Borderless Metadata Table
        meta_table = document.add_table(rows=3, cols=2)
        meta_table.autofit = False
        meta_widths = [Inches(1.5), Inches(5.7)]
        acad_match = re.search(r"(20\d\d\s*[-–/]\s*\d{2,4})", self.raw_text)
        acad_year = acad_match.group(1).replace(" ", "") if acad_match else "2024-25"
        meta_items = [
            ("Branch", f": {branch_clean}"),
            ("Year / Semester", f": {sem_display}"),
            ("Details of Exam", f": {exam_str} ({acad_year})"),
        ]
        for idx, (label, val) in enumerate(meta_items):
            c0, c1 = meta_table.rows[idx].cells
            c0.width, c1.width = meta_widths[0], meta_widths[1]
            p0 = c0.paragraphs[0]
            p0.paragraph_format.space_before = Pt(0)
            p0.paragraph_format.space_after = Pt(1)
            r0 = p0.add_run(label)
            r0.bold = True
            r0.font.size = Pt(9.5)

            p1 = c1.paragraphs[0]
            p1.paragraph_format.space_before = Pt(0)
            p1.paragraph_format.space_after = Pt(1)
            r1 = p1.add_run(val)
            r1.font.size = Pt(9.5)

        document.add_paragraph().paragraph_format.space_after = Pt(4)

        # 5. Main Table
        table = document.add_table(rows=1, cols=5)
        table.style = "Table Grid"
        table.autofit = False
        headers = [
            "Roll No.",
            "Name of the\nstudent",
            "Name of the subject &\nFaculty",
            "Area of weakness/\nAbsent / On duty",
            "Action taken",
        ]
        col_widths = [Inches(0.95), Inches(1.35), Inches(2.40), Inches(1.35), Inches(1.15)]
        for cell, text, w in zip(table.rows[0].cells, headers, col_widths):
            cell.width = w
            cell.text = text
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.bold = True
                    run.font.size = Pt(8.5)

        class_mapping = _class_subject_mapping(
            class_key, "", self.raw_text, self.subjects
        )

        student_blocks = []
        for _, row in self.df.iterrows():
            failed_subs = []
            for sub in self.subjects:
                val = row.get(sub, "")
                norm = str(val).strip().upper()
                num = pd.to_numeric(val, errors="coerce")
                if norm in {"AB", "ABSENT", "A"}:
                    w_text = "Absent for Exam"
                    act_text = "Retest Conducted"
                elif norm in {"OD", "ON DUTY", "ON-DUTY"}:
                    w_text = "On Duty"
                    act_text = "OD Recorded"
                elif pd.notna(num) and num < self.pass_mark:
                    w_text = SUBJECT_WEAKNESS_TOPIC.get(str(sub), "Partial Ordering")
                    act_text = "Special Class Conducted"
                else:
                    continue
                failed_subs.append((sub, w_text, act_text))
            if failed_subs:
                student_blocks.append((row, failed_subs))

        for student_row, failed_subs in student_blocks:
            first_row = len(table.rows)
            raw_roll = student_row.get(self.roll_col, "")
            college_roll = format_college_roll(raw_roll)
            student_name = resolve_student_name(student_row, raw_roll, self.name_col)

            for idx, (sub, w_text, act_text) in enumerate(failed_subs):
                row_cells = table.add_row().cells
                sub_name, fac_name = class_mapping.get(str(sub), (sub, "-"))
                if idx == 0:
                    row_cells[0].text = college_roll
                    row_cells[1].text = student_name
                row_cells[2].text = f"{sub}- {sub_name}\n\n{fac_name}"
                row_cells[3].text = w_text
                row_cells[4].text = act_text
                for cell, w in zip(row_cells, col_widths):
                    cell.width = w
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    for p in cell.paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if cell is row_cells[2] else WD_ALIGN_PARAGRAPH.CENTER
                        for run in p.runs:
                            run.font.size = Pt(8.5)
            if len(failed_subs) > 1:
                last_row = len(table.rows) - 1
                table.cell(first_row, 0).merge(table.cell(last_row, 0))
                table.cell(first_row, 1).merge(table.cell(last_row, 1))

        if not student_blocks:
            row_cells = table.add_row().cells
            row_cells[0].text = "-"
            row_cells[1].text = "No failed or absent subjects"
            for cell in row_cells[2:]:
                cell.text = "-"

        # 6. Footer Signatures
        document.add_paragraph().paragraph_format.space_before = Pt(24)
        sig_table = document.add_table(rows=1, cols=2)
        sig_table.autofit = False
        c_left, c_right = sig_table.rows[0].cells
        c_left.width = Inches(3.6)
        c_right.width = Inches(3.6)

        p_adv = c_left.paragraphs[0]
        p_adv.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r_adv = p_adv.add_run("Class Advisor / Counsellor")
        r_adv.bold = True
        r_adv.font.size = Pt(10)

        p_hod = c_right.paragraphs[0]
        p_hod.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hod_label = "HOD/AI&DS" if "AI" in branch_clean or "ARTIFICIAL" in branch_clean else f"HOD/{branch_clean[:4]}"
        r_hod = p_hod.add_run(hod_label)
        r_hod.bold = True
        r_hod.font.size = Pt(10)

        document.save(out_path)
        return out_path

    def export_correction_docx(self, output_dir, department="", class_key=""):
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
        self._add_word_source_header(document)

        # Department Heading
        dept_line = next(
            (
                line.strip()
                for line in self.raw_text.splitlines()
                if "DEPARTMENT OF" in line.upper()
            ),
            "",
        )
        if not dept_line:
            branch = self._branch_from_source()
            if branch and branch != "-":
                dept_line = f"DEPARTMENT OF {branch.upper()}"
            elif department:
                dept_line = f"DEPARTMENT OF {department.strip().upper()}"
            else:
                dept_line = "DEPARTMENT OF ARTIFICIAL INTELLIGENCE AND DATA SCIENCE"

        dept_p = document.add_paragraph()
        dept_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        dept_p.paragraph_format.space_before = Pt(6)
        dept_p.paragraph_format.space_after = Pt(2)
        r_dept = dept_p.add_run(dept_line)
        r_dept.bold = True
        r_dept.font.size = Pt(11)

        # Exam & Class / Section Heading
        exam_match = re.search(
            r"\bCAT\s*[-–—]?\s*(I{1,3}|IV|\d+)\b", self.raw_text, re.I
        )
        exam_str = exam_match.group(0).upper().replace("–", "-") if exam_match else "CAT I"

        class_match = re.search(
            r"Year/Sem/Sec\s*:\s*([IVX]+|\d+)\s*/\s*(\d+)\s*/\s*([A-Z])",
            self.raw_text,
            re.I,
        )
        if class_match:
            year, sem, sec = class_match.groups()
            class_str = f"{year.upper()} Year {sec.upper()} Sec"
            class_title = f"{exam_str} – {class_str}"
        else:
            cm2 = re.search(r"([IVX]+|\d+)\s*Year\s*([A-Z])\s*Sec", self.raw_text, re.I)
            if cm2:
                class_title = f"{exam_str} – {cm2.group(1).upper()} Year {cm2.group(2).upper()} Sec"
            elif class_key:
                class_title = f"{exam_str} – {class_key}"
            else:
                class_title = f"{exam_str} – III Year B Sec"

        class_p = document.add_paragraph()
        class_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        class_p.paragraph_format.space_before = Pt(0)
        class_p.paragraph_format.space_after = Pt(6)
        r_class = class_p.add_run(class_title)
        r_class.bold = True
        r_class.font.size = Pt(10)

        # Subject mapping
        correction_mapping = _class_subject_mapping(
            class_key, "", self.raw_text, self.subjects
        )
        num_cols = 2 + len(self.display_subjects)
        table = document.add_table(rows=2, cols=num_cols)
        table.style = "Table Grid"
        table.autofit = True

        # Header Rows
        c00 = table.cell(0, 0)
        c10 = table.cell(1, 0)
        c00.text = "Roll No"
        c00.merge(c10)

        c01 = table.cell(0, 1)
        c11 = table.cell(1, 1)
        c01.text = "Name of the Student"
        c01.merge(c11)

        if num_cols > 2:
            top_subj_start = table.cell(0, 2)
            top_subj_end = table.cell(0, num_cols - 1)
            top_subj_start.merge(top_subj_end)

        for index, subject in enumerate(self.display_subjects, start=2):
            sub_name = correction_mapping.get(subject, ("", ""))[0]
            label = f"{subject} & {sub_name}" if sub_name else str(subject)
            table.cell(1, index).text = label

        for r_idx in (0, 1):
            for cell in table.rows[r_idx].cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in paragraph.runs:
                        run.bold = True
                        run.font.size = Pt(8)

        # Data Rows
        for _, row in self.df.iterrows():
            cells = table.add_row().cells
            name = row.get(self.name_col, "") if self.name_col else ""
            cells[0].text = str(row.get(self.roll_col, ""))
            cells[1].text = "" if pd.isna(name) else str(name)
            for index, subject in enumerate(self.display_subjects, start=2):
                value = row.get(subject, "")
                normalized = str(value).strip().upper()
                number = pd.to_numeric(value, errors="coerce")
                if subject in self.ignored_subjects:
                    rendered = (
                        "-" if pd.isna(value) or not normalized else str(value)
                    )
                elif normalized in {"OD", "ON DUTY", "ON-DUTY", "AB", "ABSENT", "A"}:
                    rendered = ""
                elif pd.notna(number) and number >= self.pass_mark:
                    rendered = "-"
                else:
                    rendered = ""
                cells[index].text = rendered
            for cell in cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in paragraph.runs:
                        run.font.size = Pt(8)

        # HOD Signature
        dept_ref = dept_line.upper()
        if "ARTIFICIAL" in dept_ref or "AI" in dept_ref:
            hod_label = "HOD/AI&DS"
        elif "COMPUTER" in dept_ref or "CSE" in dept_ref:
            hod_label = "HOD/CSE"
        elif "COMMUNICATION" in dept_ref or "ECE" in dept_ref:
            hod_label = "HOD/ECE"
        elif "ELECTRICAL" in dept_ref or "EEE" in dept_ref:
            hod_label = "HOD/EEE"
        elif "MECHANICAL" in dept_ref or "MECH" in dept_ref:
            hod_label = "HOD/MECH"
        elif "CIVIL" in dept_ref:
            hod_label = "HOD/CIVIL"
        elif "INFORMATION" in dept_ref or "IT" in dept_ref:
            hod_label = "HOD/IT"
        else:
            hod_label = "HOD/AI&DS"

        hod_p = document.add_paragraph()
        hod_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hod_p.paragraph_format.space_before = Pt(18)
        r_hod = hod_p.add_run(hod_label)
        r_hod.bold = True
        r_hod.font.size = Pt(10)

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

        self._add_word_source_header(document)
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
        mark_headers = [
            self._subject_label(column)
            if column in self.subjects else str(column)
            for column in consolidated.columns
        ]
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
            document.add_heading(self._subject_label(subject), level=3)
            round_table = document.add_table(rows=1, cols=4)
            round_table.style = "Table Grid"
            for cell, label in zip(
                round_table.rows[0].cells,
                ["S.No", "Roll No", "Name", "Marks Obtained"],
            ):
                cell.text = label
                cell.paragraphs[0].runs[0].bold = True
            for student in students:
                cells = round_table.add_row().cells
                cells[0].text = str(student.get("S.No", ""))
                cells[1].text = str(student.get("Reg No", ""))
                cells[2].text = student.get("Name", "")
                cells[3].text = str(student.get("Marks_Obtained", ""))

        document.save(out_path)
        return out_path

    def _branch_from_source(self):
        match = re.search(
            r"B\.?\s*Tech\s*-\s*(.+?)(?:\r?\n|$)",
            self.raw_text,
            re.IGNORECASE,
        )
        return match.group(1).strip() if match else "-"

    def _source_header_lines(self):
        """Return the identifying header text found in the uploaded mark list."""
        lines = [line.strip() for line in self.raw_text.splitlines() if line.strip()]
        college = next(
            (line for line in lines if "ENGINEERING COLLEGE" in line.upper()),
            "ERODE SENGUNTHAR ENGINEERING COLLEGE",
        )
        department = next(
            (
                line for line in lines
                if "DEPARTMENT OF" in line.upper()
                or line.upper().startswith("B.TECH -")
            ),
            "",
        )
        class_match = re.search(
            r"(CAT\s*[-IIVX\d]+\s*[-–—]\s*[IVX]+\s*Year\s*[A-Z]?\s*SEC)",
            self.raw_text,
            re.IGNORECASE,
        )
        class_line = class_match.group(1).strip() if class_match else ""
        return [line for line in (college, department, class_line) if line]

    def _subject_label(self, subject):
        mapping = _class_subject_mapping(
            "", "", self.raw_text, self.subjects
        )
        name = mapping.get(subject, ("", ""))[0]
        return f"{subject} - {name}" if name and name != subject else str(subject)

    def _source_logo_images(self):
        """Return the configured college logos, with PDF extraction as fallback."""
        logo_paths = [
            os.path.join(os.path.dirname(__file__), "static", "college_logo.png"),
            os.path.join(os.path.dirname(__file__), "static", "naac_logo.png"),
            os.path.join(os.path.dirname(__file__), "static", "innovation_logo.png"),
        ]
        if all(os.path.isfile(path) for path in logo_paths):
            streams = []
            for path in logo_paths:
                with open(path, "rb") as logo_file:
                    streams.append(BytesIO(logo_file.read()))
            return streams

        try:
            import pypdfium2 as pdfium
        except ImportError:
            return []

        with pdfium.PdfDocument(self.filepath) as document:
            if not len(document):
                return []
            page = document[0]
            image = page.render(scale=1.5).to_pil()
            page.close()
        if self._source_rotation:
            image = image.rotate(self._source_rotation, expand=True)
        width, height = image.size
        top = image.crop((0, 0, width, int(height * 0.35)))
        crops = [
            top.crop((0, 0, int(width * 0.24), top.height)),
            top.crop((int(width * 0.76), 0, width, top.height)),
        ]
        result = []
        for crop in crops:
            stream = BytesIO()
            crop.save(stream, format="PNG")
            stream.seek(0)
            result.append(stream)
        return result

    def _add_pdf_source_header(self, story, style, target_width=None):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(base_dir, "static", "college_logo.png")
        naac_path = os.path.join(base_dir, "static", "naac_logo.png")
        inn_path = os.path.join(base_dir, "static", "innovation_logo.png")

        left_logo = Image(logo_path, width=22 * mm, height=22 * mm) if os.path.exists(logo_path) else ""
        right_top = Image(naac_path, width=16 * mm, height=16 * mm) if os.path.exists(naac_path) else ""
        right_bot = Image(inn_path, width=20 * mm, height=10 * mm) if os.path.exists(inn_path) else ""
        right_cell = Table([[right_top], [Spacer(1, 2 * mm)], [right_bot]], colWidths=[24 * mm])
        right_cell.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))

        center_html = """
        <para align="center">
        <b><font size="12">ERODE SENGUNTHAR</font></b><br/>
        <b><font size="10">ENGINEERING COLLEGE</font></b><br/>
        <font size="7.5">(An Autonomous Institution)</font><br/>
        <font size="6.5">Approved by AICTE, New Delhi, Permanently Affiliated to Anna University- Chennai,<br/>
        Accredited by National Board of Accreditation (NBA), New Delhi &amp;<br/>
        National Assessment and Accreditation Council (NAAC), Bangalore with 'A' Grade</font><br/>
        <b><font size="7.5">PERUNDURAI – 638 057, TAMILNADU, INDIA.</font></b>
        </para>
        """
        styles = getSampleStyleSheet()
        center_p = Paragraph(center_html, styles["Normal"])

        c_width = (target_width - 50 * mm) if target_width and target_width > 60 * mm else 132 * mm
        head_table = Table([[left_logo, center_p, right_cell]], colWidths=[25 * mm, c_width, 25 * mm])
        head_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))
        story.append(head_table)
        story.append(Spacer(1, 4 * mm))

    def _add_word_source_header(self, document):
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Pt
        from docx.shared import Inches
        from docx.shared import RGBColor

        base_dir = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(base_dir, "static", "college_logo.png")
        naac_path = os.path.join(base_dir, "static", "naac_logo.png")
        inn_path = os.path.join(base_dir, "static", "innovation_logo.png")

        table = document.add_table(rows=1, cols=3)
        table.autofit = False
        cells = table.rows[0].cells
        cells[0].width = Inches(1.1)
        cells[1].width = Inches(5.1)
        cells[2].width = Inches(1.1)

        if os.path.exists(logo_path):
            left = cells[0].paragraphs[0]
            left.alignment = WD_ALIGN_PARAGRAPH.CENTER
            left.add_run().add_picture(logo_path, width=Inches(0.9))

        if os.path.exists(naac_path) or os.path.exists(inn_path):
            right = cells[2].paragraphs[0]
            right.alignment = WD_ALIGN_PARAGRAPH.CENTER
            right.paragraph_format.space_after = Pt(2)
            if os.path.exists(naac_path):
                right.add_run().add_picture(naac_path, width=Inches(0.65))
            if os.path.exists(inn_path):
                p_r2 = cells[2].add_paragraph()
                p_r2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_r2.add_run().add_picture(inn_path, width=Inches(0.85))

        center = cells[1].paragraphs[0]
        center.alignment = WD_ALIGN_PARAGRAPH.CENTER
        header_lines = [
            ("ERODE SENGUNTHAR", 13, True),
            ("ENGINEERING COLLEGE", 11, True),
            ("(An Autonomous Institution)", 8, False),
            ("Approved by AICTE, New Delhi, Permanently Affiliated to Anna University- Chennai,", 7, False),
            ("Accredited by National Board of Accreditation (NBA), New Delhi &", 7, False),
            ("National Assessment and Accreditation Council (NAAC), Bangalore with 'A' Grade", 7, False),
            ("PERUNDURAI – 638 057, TAMILNADU, INDIA.", 8, True),
        ]
        for idx, (text, size, bold) in enumerate(header_lines):
            p = center if idx == 0 else cells[1].add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
            run = p.add_run(text)
            run.bold = bold
            run.font.size = Pt(size)
            run.font.color.rgb = RGBColor(0, 0, 0)

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
        story = []

        stats = self.get_statistics()

        # Header
        if college_name:
            story.append(Paragraph(college_name, title_style))
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
                self._subject_label(sub)[:45], p, f,
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
        mark_rows = [[
            Paragraph(
                self._subject_label(header)
                if header in self.subjects else str(header),
                mark_header_style,
            )
            for header in mark_headers
        ]]
        for _, row in consolidated.iterrows():
            rendered = []
            for column in mark_headers:
                value = row[column]
                if column in self.display_subjects and column not in self.ignored_subjects:
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
            + [20 * mm] * len(self.display_subjects)
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
            story.append(Paragraph(
                f"Subject: {escape(self._subject_label(sub))}",
                styles["Heading3"],
            ))
            story.append(Paragraph(
                f"Total failed: {len(students)} | Pass mark: {self.pass_mark}",
                styles["Normal"]
            ))
            story.append(Spacer(1, 2 * mm))

            data = [["#", "S.No", "Reg No", "Name", "Marks Obtained"]]
            for i, s in enumerate(students, 1):
                marks = s["Marks_Obtained"]
                data.append([
                    i, s.get("S.No", ""), s["Reg No"], s.get("Name", ""),
                    "-" if pd.isna(marks) else marks,
                ])

            t = Table(data, colWidths=[15 * mm, 20 * mm, 55 * mm, 75 * mm, 40 * mm],
                      repeatRows=1)
            t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e3a8a")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("ALIGN", (0, 0), (0, -1), "CENTER"),
                ("ALIGN", (4, 0), (4, -1), "CENTER"),
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
            [
                (self._subject_label(subject), count)
                for subject, count in
                self.get_statistics()["subject_fail_counts"].items()
            ],
            columns=["Subject", "Fail_Count"]
        )
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            consolidated = self.consolidated_mark_list().rename(
                columns={
                    subject: self._subject_label(subject)
                    for subject in self.subjects
                }
            )
            header_path = os.path.join(
                os.path.dirname(__file__), "static", "college_header.png"
            )
            has_header_img = os.path.isfile(header_path)
            start_row = 8 if has_header_img else 4

            consolidated.to_excel(
                writer, sheet_name="Consolidated_Marks", index=False,
                startrow=start_row
            )
            worksheet = writer.sheets["Consolidated_Marks"]
            from openpyxl.styles import Font, PatternFill, Alignment

            if has_header_img:
                branded_fill = PatternFill("solid", fgColor="282828")
                total_cols = max(len(consolidated.columns), 11)
                for row in worksheet.iter_rows(min_row=1, max_row=7,
                                              min_col=1, max_col=total_cols):
                    for cell in row:
                        cell.fill = branded_fill
                for r in range(1, 8):
                    worksheet.row_dimensions[r].height = 22

                try:
                    from openpyxl.drawing.image import Image as ExcelImage
                    logo = ExcelImage(header_path)
                    logo.width = 720
                    logo.height = int(720 / (953 / 262))
                    worksheet.add_image(logo, "A1")
                except Exception:
                    pass
            else:
                branded_fill = PatternFill("solid", fgColor="282828")
                for row in worksheet.iter_rows(min_row=1, max_row=4,
                                              min_col=1, max_col=11):
                    for cell in row:
                        cell.fill = branded_fill
                worksheet.merge_cells("D1:I4")
                worksheet["D1"] = (
                    "ERODE SENGUNTHAR\nENGINEERING COLLEGE\n"
                    "(An Autonomous Institution)\n"
                    "Approved by AICTE, New Delhi, Permanently Affiliated to "
                    "Anna University - Chennai,\n"
                    "Accredited by NBA, New Delhi & NAAC, Bangalore with 'A' Grade\n"
                    "PERUNDURAI - 638 057, TAMILNADU, INDIA."
                )
                worksheet["D1"].font = Font(name="Arial", bold=True, size=11,
                                            color="FFFFFF")
                worksheet["D1"].alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )
                try:
                    from openpyxl.drawing.image import Image as ExcelImage
                    logo_cells = ("A1", "J1", "J5")
                    for index, stream in enumerate(self._source_logo_images()):
                        stream.seek(0)
                        logo = ExcelImage(stream)
                        logo.width = 70
                        logo.height = 70
                        worksheet.add_image(logo, logo_cells[index])
                except ImportError:
                    pass
            self.get_student_wise_failures().to_excel(
                writer, sheet_name="Student_Wise_Failures", index=False
            )
            stats_df.to_excel(writer, sheet_name="Subject_Stats", index=False)
            for sub, rows in self.generate_round_lists().items():
                sheet_name = f"Round_{sub}"[:31]
                round_df = pd.DataFrame(rows)
                round_df.insert(0, "Subject", self._subject_label(sub))
                round_df.to_excel(
                    writer, sheet_name=sheet_name, index=False
                )
        return out_path