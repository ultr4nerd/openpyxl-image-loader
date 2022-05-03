"""
Contains a SheetImageLoader class that allow you to loadimages from a sheet
"""

import io
import string

from PIL import Image


class SheetImageLoader:
    """Loads all images in a sheet."""

    def __init__(self, sheet):
        """Loads all sheet images."""
        
        self._images = {}
        
        sheet_images = sheet._images
        for image in sheet_images:
            row = image.anchor._from.row + 1
            col = string.ascii_uppercase[image.anchor._from.col]
            self._images[f'{col}{row}'] = image._data

    def image_in(self, cell):
        """Checks if there's an image in specified cell."""
        return cell in self._images

    def get(self, cell):
        """Retrieves image data from a cell."""
        if cell not in self._images:
            raise ValueError("Cell {} doesn't contain an image".format(cell))
        else:
            image = io.BytesIO(self._images[cell]())
            return Image.open(image)

    def get_all_images(self):
        """Return all images from worksheet."""
        images = [io.BytesIO(image()) for image in self._images.values()]
        return [Image.open(image) for image in images]

    def get_all_images_cells(self):
        """Return all cells containing images."""
        cells = [cel for cel in self._images.keys()]
        return cells
