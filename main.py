import logging

from screenshot_sorter import *

logging.basicConfig(level=logging.INFO)


def main():
    pptx_file = input('Full path to the Presentation file: ')
    image_dir = input('Full path to the image directory: ')
    keystring = input('Key string to filter images on: ')

    x = Screens2PPTX(
        pptx_file,
        image_dir=image_dir if image_dir else r'C:\Users\User\Videos\Captures',
        keystring=keystring
    )
    x.build_title_slide()
    x.pull_images()


if __name__ == '__main__':
    main()

# F:\User\Learn\ไทยศึกษา\รุ่นเก๋า เล่าเกร็ด\ep.2.1. เกิดอะไรขึ้นในสุโขทัย หลังการสวรรคตของพ่อขุนรามคำแหง - ดร.ตรงใจ หุตางกูร.pptx