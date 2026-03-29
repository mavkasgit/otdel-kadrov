"""
HRMS - Утилиты для UI
"""
import os
from PIL import Image, ImageTk


def set_app_icon(window):
    """Устанавливает иконку для любого окна Tkinter"""
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(base_dir, "..", "icon.png")
        icon_path = os.path.abspath(icon_path)
        
        img = Image.open(icon_path)
        photo = ImageTk.PhotoImage(img)
        
        # Сначала скрываем окно, потом показываем - трюк для иконки
        window.withdraw()
        window.iconphoto(True, photo)
        window._icon = photo
        window.deiconify()
        
    except Exception as e:
        print(f"Icon error: {e}")
