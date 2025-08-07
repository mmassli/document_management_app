import os
from pathlib import Path
from spire.doc import *
from spire.doc.common import *


class WordOperations:
    """Handles Word document operations using Spire.Doc for watermarking"""

    def __init__(self, app):
        self.app = app

    def add_watermark_to_word(self, file_path, watermark_text="UNGÜLTIG", font_size=65, font_color="Red", layout=WatermarkLayout.Diagonal):
        """
        Add watermark using Spire.Doc TextWatermark
        Args:
            file_path (str): Path to the Word document
            watermark_text (str): Watermark text (default: "UNGÜLTIG")
            font_size (int): Font size in points (default: 65)
            font_color (str): Color name (default: "Red")
            layout (WatermarkLayout): Layout type (default: Diagonal)
        """
        try:
            # Create a Document object
            document = Document()

            # Load the Word document
            document.LoadFromFile(str(Path(file_path).resolve()))

            # Create a TextWatermark object
            txtWatermark = TextWatermark()

            # Set the format of the text watermark
            txtWatermark.Text = watermark_text
            txtWatermark.FontSize = font_size
            txtWatermark.Color = getattr(Color, f"get_{font_color}")()
            txtWatermark.Layout = layout

            # Add the text watermark to document
            document.Watermark = txtWatermark

            # Save the result document
            document.SaveToFile(str(Path(file_path).resolve()), FileFormat.Docx)
            document.Close()

            self.app.log_message(f"✅ Spire.Doc watermark added to {Path(file_path).name}")
            return True

        except Exception as e:
            self.app.log_message(f"❌ Error: {str(e)}")
            return False

    def is_word_document(self, file_path):
        return Path(file_path).suffix.lower() in [".docx", ".doc"]

    def add_watermark_to_archived_files(self, files_to_archive, archive_dir):
        processed = []
        for file_path in files_to_archive:
            file_path = Path(file_path)
            if self.is_word_document(file_path):
                success = self.add_watermark_to_word(str(file_path))
                status = "✅" if success else "❌"
                self.app.log_message(f"{status} {file_path.name}")
                processed.append(file_path)
        return processed
