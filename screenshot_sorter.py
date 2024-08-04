import os
from pptx import Presentation


class Screens2PPTX:

    @staticmethod
    def contains_all(string, lst):
        """Returns True if and only if each meaningful (non-empty, longer than one char) element
        from the list-like lst is contained in the string"""
        for item in lst:
            item = item.strip('": ')
            if len(item) > 1 and item not in string:
                return False
        return True

    def __init__(self, pptx_file, image_dir,
                 keystring=''  # all whitespace-separated fragments of this are to be contained in the image file name
                               # as eligibility screen. If left empty,
                               # the value of title from pres properties is used automatically - TOKENIZE?
                 ):
        self.pptx_file = pptx_file.replace('\\', '/')
        self.prs = Presentation(pptx_file)
        self.image_dir = image_dir.replace('\\', '/')
        self.keystring = keystring
        self.title = self.pptx_file.rsplit('/', 1)[-1].rsplit('.', 1)[0]
        self.prs.core_properties.title = self.title

    def populate_title(self):
        title_slide = self.prs.slides[0]
        title_shape = title_slide.shapes.title
        title_shape.text = self.title

    def build_title_slide(self):
        self.populate_title()
        self.prs.save(self.pptx_file)

    def pull_images(self):
        prs = self.prs
        keystring = self.keystring
        print(f'Pulling from {self.image_dir}\ninto {self.pptx_file}')
        if len(prs.slides) > 1:
            proceed = input(
                'WARNING! More than one slide detected. Eligible images will be added to the end. Any symbol to '
                'proceed, mere Enter to cancel.'
            )
            if not proceed:
                print('Operation aborted. Please check destination file contents.')
                return
        blank_slide_layout = prs.slide_layouts[6]
        if not keystring:
            keystring = self.title
        image_files = [file for file in os.scandir(self.image_dir) if self.contains_all(
            file.name, keystring.split(' ')) and file.name.endswith('.png')]  # TODO: tokenize rather than split?
        image_files.sort(key=lambda x: x.stat().st_ctime)
        print(f'Total eligible images: {len(image_files)}')
        for image in image_files:
            slide = prs.slides.add_slide(blank_slide_layout)
            slide.shapes.add_picture(image.path, 0, 0, prs.slide_width)
            print('.', end="")
        print('\nDone')
        prs.save(self.pptx_file)
