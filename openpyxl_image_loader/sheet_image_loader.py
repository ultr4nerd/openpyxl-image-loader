"""
Contains a SheetImageLoader class that allow you to loadimages from a sheet
"""

import io
import string
import itertools
from PIL import Image


class SheetImageLoader:
    """Loads all images in a sheet"""

    def __init__(self, sheet):
        """Loads all sheet images"""
        self._images = {}
        # Holds an array of A-ZZ
        col_holder = list(itertools.chain(string.ascii_uppercase, (''.join(pair) for pair in itertools.product(string.ascii_uppercase, repeat=2))))
        sheet_images = sheet._images
        for image in sheet_images:
            row = image.anchor._from.row + 1
            col = col_holder[image.anchor._from.col] # Convert the col to correct Excel Col.
            if self._images.get(f'{col}{row}'): # Append multiple images of one cell
                self._images[f'{col}{row}'].append(image._data)
            else:
                self._images[f'{col}{row}'] = [image._data]

    def image_in(self, cell):
        """Checks if there's an image in specified cell"""
        return cell in self._images

    def get(self, cell):
        """Retrieves image data from a cell"""
        if cell not in self._images:
            raise ValueError("Cell {} doesn't contain an image".format(cell))
        else:
            for ele in self._images[cell]: # Show all images of one cell 
                image = io.BytesIO(ele())
                display(Image.open(image))
                image.close()

    def close(self):
        """Delete image buffer"""
        self._images.clear()
