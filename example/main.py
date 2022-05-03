from openpyxl import load_workbook
from openpyxl_image_loader import SheetImageLoader


def main():
    path_to_file: str = 'some_path'
    # Load your workbook and sheet as you want, for example
    wb = load_workbook(path_to_file)
    sheet = wb['required_sheet']

    # Put your sheet in the loader
    image_loader = SheetImageLoader(sheet)

    # And get image from specified cell
    image = image_loader.get('A1')

    # Image now is a Pillow image, so you can do the following
    image.show()

    # Ask if there's an image in a cell
    if image_loader.image_in('A4'):
        print("Got it!")

    # And if you want get all images, write this
    images = image_loader.get_all_images()
    for image in images:
        image.show()  # if you want save, you can write image.save()

    cells = image_loader.get_all_images_cells()
    print(cells)


if __name__ == '__main__':
    main()
