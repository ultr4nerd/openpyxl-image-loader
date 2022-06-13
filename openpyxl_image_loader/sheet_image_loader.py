"""
Contains a SheetImageLoader class that allow you to loadimages from a sheet
"""

import io
import string

from PIL import Image

excel_col_name = lambda n: '' if n <= 0 else excel_col_name((n - 1) // 26) + chr((n - 1) % 26 + ord('A'))

class SheetImageLoader:
    """Loads all images in a sheet"""

    def __init__(self, sheet):
        """Loads all sheet images"""
        
        self._images = {}
        
        sheet_images = sheet._images
        for image in sheet_images:
            row = image.anchor._from.row + 1
            col = excel_col_name( image.anchor._from.col )
            self._images[f'{col}{row}'] = image._data

    def image_in(self, cell):
        """Checks if there's an image in specified cell"""
        return cell in self._images

    def get(self, cell):
        """Retrieves image data from a cell"""
        if cell not in self._images:
            raise ValueError("Cell {} doesn't contain an image".format(cell))
        else:
            image = io.BytesIO(self._images[cell]())
            return Image.open(image)
