import os
import tempfile
import math
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import red, blue, green, black, gray, orange, purple
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
        """Add watermark to PDF using ReportLab"""
        try:
            reader = PdfReader(file_path)
            writer = PdfWriter()

            # Apply watermark to each page individually
            for i, page in enumerate(reader.pages):
                page_width = float(page.mediabox.width)
                page_height = float(page.mediabox.height)

                # Create watermark for the specific page size
                watermark_pdf = self._create_watermark_pdf(
                    watermark_text, font_name, font_size, font_color, transparency,
                    page_width, page_height
                )

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

    def add_watermark_to_pdf_all_pages(self, file_path, watermark_text="UNGÜLTIG", font_name="Helvetica-Bold",
                                   font_size=80, font_color="red", transparency=0.7):
        """Add watermark to all pages of PDF"""
        return self.add_watermark_to_pdf(file_path, watermark_text, font_name, font_size, 
                                        font_color, transparency)

    def add_watermark_to_pdf_odd_pages_only(self, file_path, watermark_text="UNGÜLTIG", font_name="Helvetica-Bold",
                                        font_size=80, font_color="red", transparency=0.7):
        """Add watermark to odd pages only"""
        return self.add_watermark_to_pdf(file_path, watermark_text, font_name, font_size, 
                                        font_color, transparency)

    def add_watermark_to_pdf_auto_detect(self, file_path, watermark_text="UNGÜLTIG", font_name="Helvetica-Bold",
                                     font_size=80, font_color="red", transparency=0.7):
        """Add watermark to PDF"""
        return self.add_watermark_to_pdf(file_path, watermark_text, font_name, font_size, 
                                        font_color, transparency)
    
    def _create_watermark_pdf(self, text, font_name, font_size, font_color, transparency, width, height):
        """Create a watermark PDF for a specific page size"""
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(width, height))

        can.setFont(font_name, font_size)

        # Set fill color with transparency approximation
        if isinstance(font_color, str) and font_color.lower() in self.color_map:
            color = self.color_map[font_color.lower()]
            if transparency < 1.0:
                if font_color.lower() == 'red':
                    # Use a lighter red for transparency effect
                    can.setFillColorRGB(1.0, 0.3, 0.3)
                else:
                    can.setFillColor(color)
            else:
                can.setFillColor(color)
        elif isinstance(font_color, (tuple, list)) and len(font_color) == 3:
            r, g, b = [c / 255 for c in font_color]
            if transparency < 1.0:
                # Adjust color for transparency effect
                r = min(1.0, r + (1 - transparency) * 0.3)
                g = min(1.0, g + (1 - transparency) * 0.3)
                b = min(1.0, b + (1 - transparency) * 0.3)
            can.setFillColorRGB(r, g, b)
        else:
            can.setFillColor(red)

        # Calculate the diagonal angle from bottom-left to top-right
        # This ensures the watermark goes from bottom-left to top-right
        diagonal_angle = math.degrees(math.atan2(height, width))
        
        # Move to center and rotate across diagonal (bottom-left to top-right)
        can.saveState()
        can.translate(width / 2, height / 2)
        can.rotate(diagonal_angle)
        can.drawCentredString(0, 0, text)
        can.restoreState()

        can.save()
        packet.seek(0)

        return PdfReader(packet).pages[0]
    
    def _create_fallback_watermark(self, text, width, height):
        """Create a simple fallback watermark if the main method fails"""
        try:
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(width, height))
            
            # Use basic settings
            can.setFont("Helvetica-Bold", 60)
            can.setFillColor(red)
            
            # Simple centered text
            center_x = width / 2
            center_y = height / 2
            can.drawCentredString(center_x, center_y, text)
            
            can.save()
            packet.seek(0)
            
            return PdfReader(packet).pages[0]
            
        except Exception as e:
            self.app.log_message(f"❌ Error creating fallback watermark: {str(e)}")
            raise e


    def is_pdf_document(self, file_path):
        """Check if file is a PDF document"""
        return Path(file_path).suffix.lower() == ".pdf"

    def add_watermark_to_archived_pdfs(self, files_to_archive, archive_dir):
        """Add watermarks to PDF files before archiving"""
        processed = []
        for file_path in files_to_archive:
            file_path = Path(file_path)
            if self.is_pdf_document(file_path):
                try:
                    # Use the main watermark method for all PDFs
                    success = self.add_watermark_to_pdf(str(file_path))
                    status = "✅" if success else "❌"
                    self.app.log_message(f"{status} Watermarked PDF: {file_path.name}")
                    
                    processed.append(file_path)
                    
                except Exception as e:
                    self.app.log_message(f"❌ Error processing PDF {file_path.name}: {str(e)}")
                    # Still add to processed list to avoid reprocessing
                    processed.append(file_path)
                    
        return processed

    def test_watermark_creation(self, file_path):
        """Test function to debug watermark creation"""
        try:
            reader = PdfReader(file_path)
            if len(reader.pages) > 0:
                page = reader.pages[0]
                page_width = float(page.mediabox.width)
                page_height = float(page.mediabox.height)
                
                self.app.log_message(f"Page dimensions: {page_width} x {page_height}")
                
                # Test watermark creation with highly visible settings
                watermark_pdf = self._create_watermark_pdf(
                    "TEST", "Helvetica-Bold", 200, "black", 1.0,  # Large, black, opaque
                    page_width, page_height
                )
                
                self.app.log_message("✅ Watermark creation test successful")
                return True
            else:
                self.app.log_message("❌ No pages found in PDF")
                return False
        except Exception as e:
            self.app.log_message(f"❌ Watermark test failed: {str(e)}")
            return False

    def create_test_watermark(self, file_path, watermark_text="TEST_WATERMARK"):
        """Create a test watermark with maximum visibility for debugging"""
        try:
            reader = PdfReader(file_path)
            if len(reader.pages) > 0:
                page = reader.pages[0]
                page_width = float(page.mediabox.width)
                page_height = float(page.mediabox.height)
                rotation = (page.get("/Rotate") or 0) % 360
                
                # Determine watermark dimensions for rotated pages
                w, h = page_width, page_height
                wm_w, wm_h = (h, w) if rotation in (90, 270) else (w, h)
                
                self.app.log_message(f"Creating test watermark: {wm_w:.1f}x{wm_h:.1f} (rotation: {rotation}°)")
                
                # Create highly visible test watermark
                watermark_pdf = self._create_watermark_pdf(
                    watermark_text, "Helvetica-Bold", 200, "black", 1.0,  # Large, black, opaque
                    wm_w, wm_h
                )
                
                self.app.log_message("✅ Test watermark creation successful")
                return watermark_pdf
            else:
                self.app.log_message("❌ No pages found in PDF")
                return None
        except Exception as e:
            self.app.log_message(f"❌ Test watermark creation failed: {str(e)}")
            return None 