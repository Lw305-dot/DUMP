from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from pathlib import Path

# Settings
output_pdf = "ID_Cards.pdf"
# cards_dir = Path("./Employee_Reports26/TrainingIDs")
cards_dir = Path("./Generated_IDs4")
card_files = sorted(cards_dir.glob("ID_CARD_*.png"))

# A4 page setup
PAGE_WIDTH, PAGE_HEIGHT = A4
MARGIN = 5  # Left, Right, Top, Bottom margins
GAP_X = 10  # Horizontal spacing between cards
GAP_Y = 15  # Vertical spacing between rows

# Auto-fit card width so that two cards + gap fit perfectly on one row
CARDS_PER_ROW = 2
CARD_WIDTH = (PAGE_WIDTH - (2 * MARGIN) - GAP_X) / CARDS_PER_ROW
CARD_HEIGHT = CARD_WIDTH * 0.65  # Maintain approximate aspect ratio

# Create PDF canvas
c = canvas.Canvas(output_pdf, pagesize=A4)

# Initial positions
x = MARGIN
y = PAGE_HEIGHT - MARGIN - CARD_HEIGHT
cards_in_row = 0

for i, img_path in enumerate(card_files):
    c.drawImage(ImageReader(img_path), x, y, width=CARD_WIDTH, height=CARD_HEIGHT)

    cards_in_row += 1
    if cards_in_row < CARDS_PER_ROW:
        # Move right for next card
        x += CARD_WIDTH + GAP_X
    else:
        # Start a new row
        cards_in_row = 0
        x = MARGIN
        y -= CARD_HEIGHT + GAP_Y

        # If there's no space for the next row, add a new page
        if y < MARGIN:
            c.showPage()
            x = MARGIN
            y = PAGE_HEIGHT - MARGIN - CARD_HEIGHT

# Finalize PDF
c.save()
print(f"✅ Saved to: {output_pdf}")
