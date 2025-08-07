import os
import tempfile
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.colors import red, blue, green, black, gray, orange, purple
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io


class PDFOperations:
    """Handles PDF document operations including watermarking"""

    def __init__(self, app):
        self.app = app
        self.color_map = {
            'red': red,
            'blue': blue,
            'green': green,
            'black': black,
            'gray': gray,
            'orange': orange,
            'purple': purple
        }

    def add_watermark_to_pdf(self, file_path, watermark_text="UNGÜLTIG", font_name="Helvetica-Bold", 
                            font_size=80, font_color="red", transparency=0.7):
        """
        Add watermark to PDF using ReportLab
        Args:
            file_path (str): Path to the PDF document
            watermark_text (str): Watermark text (default: "UNGÜLTIG")
            font_name (str): Font name (default: "Helvetica-Bold")
            font_size (int): Font size in points (default: 80)
            font_color (str): Color name or RGB tuple (default: "red")
            transparency (float): Transparency level 0.0-1.0 (default: 0.7)
        """
        try:
            # Read the original PDF
            reader = PdfReader(file_path)
            writer = PdfWriter()

            # Create watermark
            watermark_pdf = self._create_watermark_pdf(
                watermark_text, font_name, font_size, font_color, transparency
            )

            # Apply watermark to each page
            for page in reader.pages:
                page.merge_page(watermark_pdf)
                writer.add_page(page)

            # Save the watermarked PDF
            with open(file_path, 'wb') as output_file:
                writer.write(output_file)

            self.app.log_message(f"✅ PDF watermark added to {Path(file_path).name}")
            return True

        except Exception as e:
            self.app.log_message(f"❌ Error adding watermark to PDF: {str(e)}")
            return False

    def _create_watermark_pdf(self, text, font_name, font_size, font_color, transparency):
        """Create a PDF with the watermark text"""
        # Create a temporary PDF with the watermark
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        
        # Set font and size
        can.setFont(font_name, font_size)
        
        # Set color with transparency effect
        if isinstance(font_color, str) and font_color.lower() in self.color_map:
            color = self.color_map[font_color.lower()]
            # Apply transparency by using a lighter color
            if transparency < 1.0:
                # Create a lighter version of the color for transparency effect
                if font_color.lower() == 'red':
                    can.setFillColorRGB(1.0, 0.3, 0.3)  # Light red
                else:
                    can.setFillColor(color)
            else:
                can.setFillColor(color)
        elif isinstance(font_color, (tuple, list)) and len(font_color) == 3:
            # Apply transparency to RGB values
            r, g, b = font_color[0]/255, font_color[1]/255, font_color[2]/255
            if transparency < 1.0:
                # Lighten the color for transparency effect
                r = min(1.0, r + (1 - transparency) * 0.3)
                g = min(1.0, g + (1 - transparency) * 0.3)
                b = min(1.0, b + (1 - transparency) * 0.3)
            can.setFillColorRGB(r, g, b)
        else:
            can.setFillColor(red)  # Default to red
        
        # Get page dimensions
        page_width, page_height = A4
        
        # Calculate text dimensions (approximate)
        text_width = len(text) * font_size * 0.6  # Approximate width
        text_height = font_size
        
        # Position text in center with diagonal rotation
        x = (page_width - text_width) / 2
        y = (page_height + text_height) / 2
        
        # Save canvas state
        can.saveState()
        
        # Move to center and rotate
        can.translate(x + text_width/2, y - text_height/2)
        can.rotate(45)  # Diagonal rotation
        
        # Draw text centered
        can.drawString(-text_width/2, 0, text)
        
        # Restore canvas state
        can.restoreState()
        
        can.save()
        packet.seek(0)
        
        return PdfReader(packet).pages[0]

    def is_pdf_document(self, file_path):
        """Check if file is a PDF document"""
        return Path(file_path).suffix.lower() == ".pdf"

    def add_watermark_to_archived_pdfs(self, files_to_archive, archive_dir):
        """Add watermarks to PDF files before archiving"""
        processed = []
        for file_path in files_to_archive:
            file_path = Path(file_path)
            if self.is_pdf_document(file_path):
                success = self.add_watermark_to_pdf(str(file_path))
                status = "✅" if success else "❌"
                self.app.log_message(f"{status} {file_path.name}")
                processed.append(file_path)
        return processed 