# atp_docx_insert.py
from docx import Document
from docx.shared import Inches, Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
from PIL import Image
import io
import os


TEXT_PLACEHOLDERS = [
    "[SK_1]",
    "[SITE_ID1]",
    "[SITE_NAME1]",
    "[SITE_ID]",
    "[SITE_NAME]",
    "[HOSTNAME]",
    "[SCOPE_OF_WORK]",
    "[DEVICE_TYPE]",
    "[PROJECT_CODE]",
]


class ATPDocxInserter:
    """
    Handles photo insertion and text replacement in DOCX ATP templates.
    """

    def __init__(self, docx_path):
        """
        Initialize with a DOCX template.

        Args:
            docx_path: Path to the DOCX template file
        """
        self.docx_path = docx_path
        self.doc = Document(docx_path)
        self.photo_mappings = self.detect_photo_placeholders()
        self.text_mappings = self.detect_text_placeholders()

    def detect_photo_placeholders(self):
        """
        Detect photo placeholder text in the DOCX document.
        Looks for patterns like [PHOTO_FRONT_VIEW] in paragraphs and tables.
        """
        mappings = []

        # Check all paragraphs
        # for para_idx, paragraph in enumerate(self.doc.paragraphs):
        #     text = paragraph.text.strip()
        #     if self.is_photo_placeholder(text):
        #         mappings.append(
        #             {
        #                 "type": "paragraph",
        #                 "paragraph_index": para_idx,
        #                 "placeholder": text,
        #                 "photo_type": self.get_photo_type(text),
        #                 "location": f"Paragraph {para_idx + 1}",
        #             }
        #         )
        for para_idx, paragraph in enumerate(self.doc.paragraphs):
            text = paragraph.text.strip()
            if self.is_photo_placeholder(text):
                mappings.append({
                "type": "paragraph",
                "paragraph_index": para_idx,
                "placeholder": text,
                "photo_type": self.get_photo_type(text),
                "location": f"Paragraph {para_idx + 1}",
            })

        # Check all tables
        # for table_idx, table in enumerate(self.doc.tables):
        #     for row_idx, row in enumerate(table.rows):
        #         for cell_idx, cell in enumerate(row.cells):
        #             for para_idx, paragraph in enumerate(cell.paragraphs):
        #                 text = paragraph.text.strip()
        #                 if self.is_photo_placeholder(text):
        #                     mappings.append(
        #                         {
        #                             "type": "table_cell",
        #                             "table_index": table_idx,
        #                             "row_index": row_idx,
        #                             "cell_index": cell_idx,
        #                             "paragraph_index": para_idx,
        #                             "placeholder": text,
        #                             "photo_type": self.get_photo_type(text),
        #                             "location": f"Table {table_idx + 1}, Cell ({row_idx + 1},{cell_idx + 1})",
        #                         }
        #                     )
        for table_idx, table in enumerate(self.doc.tables):
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    for para_idx, paragraph in enumerate(cell.paragraphs):
                        text = paragraph.text.strip()
                        if self.is_photo_placeholder(text):
                            mappings.append({
                            "type": "table_cell",
                            "table_index": table_idx,
                            "row_index": row_idx,
                            "cell_index": cell_idx,
                            "paragraph_index": para_idx,
                            "placeholder": text,
                            "photo_type": self.get_photo_type(text),
                            "location": f"Table {table_idx + 1}, Cell ({row_idx + 1},{cell_idx + 1})",
                        })


        return mappings

    # def detect_text_placeholders(self):
    #     """
    #     Detect text placeholder patterns in the DOCX document.
    #     """
    #     mappings = []
    #     text_patterns = [
    #         r"\[SITE_?ID\]",
    #         r"\[SITE_?NAME\]",
    #         r"\[HOSTNAME\]",
    #         r"\[SCOPE_?OF_?WORK\]",
    #         r"\[DEVICE_?TYPE\]",
    #         r"\[PROJECT_?CODE\]",
    #         r"\[DATE\]",
    #         r"\[ENGINEER\]",
    #         r"\[LOCATION\]",
    #         r"\[ADDRESS\]",
    #         r"\[SK_?1\]",   # 👈 THIS LINE
    #         # Add more patterns as needed
    #     ]

    #     # Check paragraphs
    #     for para_idx, paragraph in enumerate(self.doc.paragraphs):
    #         text = paragraph.text
    #         for pattern in text_patterns:
    #             if re.search(pattern, text, re.IGNORECASE):
    #                 mappings.append(
    #                     {
    #                         "type": "paragraph",
    #                         "paragraph_index": para_idx,
    #                         "placeholder": re.search(
    #                             pattern, text, re.IGNORECASE
    #                         ).group(0),
    #                         "location": f"Paragraph {para_idx + 1}",
    #                     }
    #                 )
    #                 break

    #     # Check tables
    #     for table_idx, table in enumerate(self.doc.tables):
    #         for row_idx, row in enumerate(table.rows):
    #             for cell_idx, cell in enumerate(row.cells):
    #                 for para_idx, paragraph in enumerate(cell.paragraphs):
    #                     text = paragraph.text
    #                     for pattern in text_patterns:
    #                         if re.search(pattern, text, re.IGNORECASE):
    #                             mappings.append(
    #                                 {
    #                                     "type": "table_cell",
    #                                     "table_index": table_idx,
    #                                     "row_index": row_idx,
    #                                     "cell_index": cell_idx,
    #                                     "paragraph_index": para_idx,
    #                                     "placeholder": re.search(
    #                                         pattern, text, re.IGNORECASE
    #                                     ).group(0),
    #                                     "location": f"Table {table_idx + 1}, Cell ({row_idx + 1},{cell_idx + 1})",
    #                                 }
    #                             )
    #                             break

    #     return mappings

    def detect_text_placeholders(self):
        mappings = []
        seen = set()

        def scan_text(text, location):
            for placeholder in TEXT_PLACEHOLDERS:
                if placeholder in text and placeholder not in seen:
                    seen.add(placeholder)

                    key = placeholder.strip("[]").lower()

                    mappings.append({
                        "placeholder": placeholder,
                        "placeholder_key": key,
                        "display_name": placeholder.strip("[]").replace("_", " "),
                        "location": location,
                        "required": key in ["site_id", "site_name", "project_code"],
                    })

        # Scan paragraphs
        for i, paragraph in enumerate(self.doc.paragraphs):
            scan_text(paragraph.text, f"Paragraph {i + 1}")

        # Scan tables
        for t_i, table in enumerate(self.doc.tables):
            for r_i, row in enumerate(table.rows):
                for c_i, cell in enumerate(row.cells):
                    for p_i, paragraph in enumerate(cell.paragraphs):
                        scan_text(
                            paragraph.text,
                            f"Table {t_i + 1}, Cell ({r_i + 1},{c_i + 1})"
                        )

        return mappings

    def is_photo_placeholder(self, text):
        """Check if text is a photo placeholder."""
        photo_placeholders = [
            "[BEFORE_TOWER1]",
            "[AFTER_TOWER1]",
            "[PHOTO_FRONT_SPACE]",
            "[PHOTO_REAR_SPACE]",
            "[PHOTO_POWER_CABLE]",
            "[PHOTO_GROUNDING]",
            "[PHOTO_CONNECTION]",
            "[PHOTO]",
            "[IMAGE]",
            "[PICTURE]",
        ]
        return text.upper() in [p.upper() for p in photo_placeholders]

    def get_photo_type(self, placeholder):
        """Get human-readable photo type from placeholder."""
        photo_types = {
            "[BEFORE_TOWER1]": "Before Tower 1",
            "[AFTER_TOWER1]": "After Tower 1",
            "[PHOTO_FRONT_SPACE]": "Front Space",
            "[PHOTO_REAR_SPACE]": "Rear Space",
            "[PHOTO_POWER_CABLE]": "Power Cable",
            "[PHOTO_GROUNDING]": "Grounding Cable",
            "[PHOTO_CONNECTION]": "Connection Router",
            "[PHOTO]": "Photo",
            "[IMAGE]": "Image",
            "[PICTURE]": "Picture",
        }
        return photo_types.get(placeholder.upper(), placeholder)

    def insert_photo(
        self, mapping_index, photo_path, width_inches=3.0, height_inches=2.0
    ):
        """
        Insert a photo at the specified placeholder location.

        Args:
            mapping_index: Index in photo_mappings list
            photo_path: Path to the photo file
            width_inches: Width of image in inches
            height_inches: Height of image in inches
        """
        if mapping_index >= len(self.photo_mappings):
            return False

        mapping = self.photo_mappings[mapping_index]

        try:
            if mapping["type"] == "paragraph":
                # Insert photo in paragraph
                paragraph = self.doc.paragraphs[mapping["paragraph_index"]]
                # Clear the placeholder text
                paragraph.clear()
                # Add the image
                run = paragraph.add_run()
                run.add_picture(
                    photo_path, width=Inches(width_inches), height=Inches(height_inches)
                )
                # Center the image
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            elif mapping["type"] == "table_cell":
                # Insert photo in table cell
                table = self.doc.tables[mapping["table_index"]]
                cell = table.cell(mapping["row_index"], mapping["cell_index"])
                paragraph = cell.paragraphs[mapping["paragraph_index"]]
                # Clear the placeholder text
                paragraph.clear()
                # Add the image
                run = paragraph.add_run()
                run.add_picture(
                    photo_path, width=Inches(width_inches), height=Inches(height_inches)
                )
                # Center the image
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            return True

        except Exception as e:
            print(f"Error inserting photo: {e}")
            return False

    # def replace_text(self, placeholder, replacement_text):
    #     """
    #     Replace text placeholder with actual text.

    #     Args:
    #         placeholder: The placeholder text to replace (e.g., [SITE_ID])
    #         replacement_text: The text to insert
    #     """
    #     # Search in paragraphs
    #     for paragraph in self.doc.paragraphs:
    #         if placeholder.upper() in paragraph.text.upper():
    #             paragraph.text = paragraph.text.replace(placeholder, replacement_text)

    #     # Search in tables
    #     for table in self.doc.tables:
    #         for row in table.rows:
    #             for cell in row.cells:
    #                 for paragraph in cell.paragraphs:
    #                     if placeholder.upper() in paragraph.text.upper():
    #                         paragraph.text = paragraph.text.replace(
    #                             placeholder, replacement_text
    #                         )

    def replace_text(self, placeholder, replacement_text):
        placeholder_upper = placeholder.upper()

    # Paragraphs
        for paragraph in self.doc.paragraphs:
            for run in paragraph.runs:
                if placeholder_upper in run.text.upper():
                    run.text = run.text.replace(placeholder, replacement_text)

    # Tables
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            if placeholder_upper in run.text.upper():
                                run.text = run.text.replace(placeholder, replacement_text)

    # def replace_all_text(self, text_values):
    #     """
    #     Replace all text placeholders with values from dictionary.

    #     Args:
    #         text_values: Dictionary mapping placeholder keys to values
    #     """
    #     placeholder_mapping = {
    #         "before_tower1": ["[BEFORE_TOWER1]", "[BEFORETOWER1]"],
    #         "after_tower1": ["[AFTER_TOWER1]", "[AFTERTOWER1]"],
    #         "hostname": ["[HOSTNAME]"],
    #         "scope_of_work": ["[SCOPE_OF_WORK]", "[SCOPEOFWORK]"],
    #         "device_type": ["[DEVICE_TYPE]", "[DEVICETYPE]"],
    #         "project_code": ["[PROJECT_CODE]", "[PROJECTCODE]"],
    #         "date": ["[DATE]"],
    #         "engineer": ["[ENGINEER]"],
    #         "location": ["[LOCATION]"],
    #         "address": ["[ADDRESS]"],
    #         "systemkey1" : "[SK_1]",
    #         "site_id1" : "[SITE_ID1]",
    #         "site_name1":"[SITE_NAME1]",
    #     }

    #     for key, value in text_values.items():
    #         if value:
    #             if key in placeholder_mapping:
    #                 for placeholder in placeholder_mapping[key]:
    #                     self.replace_text(placeholder, value)
    #             else:
    #                 # Try generic placeholder format
    #                 generic_placeholder = f"[{key.upper()}]"
    #                 self.replace_text(generic_placeholder, value)


    def replace_all_text(self, text_values):
        """
        text_values example:
        {
            "sk_1": "ABC123",
            "site_id1": "SITE-001",
            "site_name1": "Jakarta Tower"
        }
        """

        for key, value in text_values.items():
            if not value:
                continue

            # Convert key -> placeholder
            # sk_1 -> [SK_1]
            placeholder = f"[{key.upper()}]"

            self.replace_text(placeholder, value)

    def get_available_photo_slots(self):
        """Get photo slots information for frontend display."""
        slots = []

        for idx, mapping in enumerate(self.photo_mappings):
            slots.append(
                {
                    "slot_index": idx,
                    "type": mapping["photo_type"],
                    "photo_type": mapping["photo_type"],
                    "location": mapping["location"],
                    "placeholder": mapping["placeholder"],
                    "status": "empty",
                }
            )

        return slots

    def get_available_text_fields(self):
        """Get text fields information for frontend display."""
        text_fields = []
        seen_placeholders = set()

        for mapping in self.text_mappings:
            placeholder = mapping["placeholder"].upper()
            if placeholder not in seen_placeholders:
                seen_placeholders.add(placeholder)

                # Extract key from placeholder (e.g., [SITE_ID] -> site_id)
                key = placeholder.strip("[]").lower().replace("_", "")

                text_fields.append(
                    {
                        "placeholder": placeholder,
                        "placeholder_key": key,
                        "display_name": placeholder.strip("[]")
                        .replace("_", " ")
                        .title(),
                        "location": mapping["location"],
                        "required": key in ["siteid", "sitename", "projectcode"],
                    }
                )

        return text_fields

    def save(self, output_path):
        """Save the modified document."""
        self.doc.save(output_path)
