# Patch process_references in main.py
import re

with open('main.py', 'r', encoding='utf-8') as f:
    code = f.read()

new_process_references = """    def process_references(self):
        self.log("Processing references...")
        self.top_references = []
        self.related_references = []

        def process_reference(rank_cell, is_related=False):
            ref = self.Reference()
            ref.Rank = rank_cell
            current_row, current_col = rank_cell_coords
            try:
                if not is_related:
                    ref.RawPublicationNumber = str(self.df.iloc[current_row-9, current_col])
                    ref.PublicationNumber = self.clean_publication_number(self.df.iloc[current_row-9, current_col])
                    ref.PriorityDate = self.format_date(current_row-8, current_col)
                    ref.FilingDate = self.format_date(current_row-7, current_col)
                    ref.PublicationDate = self.format_date(current_row-6, current_col)
                    orig_assignee = self.df.iloc[current_row-4, current_col]
                    curr_assignee = self.df.iloc[current_row-5, current_col]
                    ref.OriginalAssignee = str(orig_assignee) if pd.notna(orig_assignee) else ""
                    ref.CurrentAssignee = str(curr_assignee) if pd.notna(curr_assignee) else ""
                    ref.Title = str(self.df.iloc[current_row-3, current_col]).strip()
                    ref.URL = str(self.df.iloc[current_row-2, current_col])
                    ref.isNPL = False if "patents.google" in ref.URL else True
                    self.top_references.append(ref)
                else:
                    ref.URL = str(self.df.iloc[current_row-2, current_col])
                    ref.isNPL = False if "patents.google" in ref.URL else True
                    ref.Title = str(self.df.iloc[current_row-3, current_col]).strip()
                    orig_assignee = self.df.iloc[current_row-4, current_col]
                    curr_assignee = self.df.iloc[current_row-5, current_col]
                    ref.OriginalAssignee = str(orig_assignee) if pd.notna(orig_assignee) else ""
                    ref.CurrentAssignee = str(curr_assignee) if pd.notna(curr_assignee) else ""
                    ref.PublicationNumber = self.clean_publication_number(self.df.iloc[current_row-9, current_col])
                    self.related_references.append(ref)
            except IndexError as e:
                self.log(f"Error processing reference at row {current_row}, col {current_col}: {str(e)}")

        try:
            rank_header_row = self.df[self.df[0] == 'Rank'].index[0]
            current_col = 1
            while current_col < self.df.shape[1]:
                current_row = rank_header_row
                while current_row < self.df.shape[0]:
                    cell_value = self.df.iloc[current_row, current_col]
                    if pd.isna(cell_value):
                        break
                    rank_text = normalize_rank(cell_value)
                    if (
                        is_normal_letter_rank(rank_text)
                        or is_system_parent_rank(rank_text)
                        or is_system_child_rank(rank_text)
                    ):
                        rank_cell_coords = (current_row, current_col)
                        process_reference(rank_text, is_related=False)
                    elif rank_text in ['RR', 'RR NPL']:
                        rank_cell_coords = (current_row, current_col)
                        process_reference(rank_text, is_related=True)
                    current_row += 3
                current_col += 1
            self.include_other_related_references = len(self.related_references) > 0
            self.log("References processed.")
        except Exception as e:
            self.log(f"Error processing references: {str(e)}")
            raise"""

# Find process_references method in main.py and replace it
# Match from "def process_references(self):" to the next "def replace_in_paragraphs_and_tables(self,"
start_p = code.find("    def process_references(self):")
end_p = code.find("    def replace_in_paragraphs_and_tables(self,")

if start_p != -1 and end_p != -1:
    code = code[:start_p] + new_process_references + "\\n\\n" + code[end_p:]
    print("✓ Patched process_references.")
else:
    print("✗ Failed to find process_references start/end in main.py.")

with open('main.py', 'w', encoding='utf-8') as f:
    f.write(code)
