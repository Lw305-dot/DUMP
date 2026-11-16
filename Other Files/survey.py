import qrcode
from pathlib import Path
output_dir = Path("/workspaces/DUMP")
# link="https://docs.google.com/forms/d/e/1FAIpQLSc388co6eV7ZoCMTeVieP1WsfcPzUlXgZTSw3sJUExN4mQvTw/viewform?usp=dialog"
link="https://docs.google.com/forms/d/e/1FAIpQLScYy-PZlUZEqDP1A0iGqPA2UEB7ZvIDZISH0GK2a4i8JUY26w/viewform"
Qr=qrcode.QRCode(version=1,box_size=10,border=4)
Qr.add_data(link)
Qr.make(fit=True)
img = Qr.make_image(fill_color="black", back_color="white")
output_path = output_dir / "SatisfactionSurvey.png"
img.save(output_path)
print(f"QR code saved to: {output_path}")